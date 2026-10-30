Attribute VB_Name = "editbox_c"
' lib/widget/editbox.cpp, editbox.h: the widget class W_EDITBOX (docs/CONVENTIONS.md,
' "Widgets"; its fields are W_EDITBOX.cls) and W_EDBINIT's constructor (the struct is
' widget_c's). W_EDITBOX(W_EDBINIT const *) is W_EDITBOX_new_init, W_EDITBOX() is
' W_EDITBOX_new. A text position is a UTF-16 unit of the VB String (WzString's
' code point in the Basic Multilingual Plane): substr, insert and remove are Mid$
' and Left$. optional<PIELIGHT> fixedPlaceholderTextColor is
' has_fixedPlaceholderTextColor and fixedPlaceholderTextColor. The handlers are
' WzFuncs called with the edit box (OnReturnHandler; OnTabHandler returns a
' Boolean), OnPasteTransformFunc with the text (it returns the new text), the
' WIDGET_AUDIOCALLBACK with the audio id. StartTextInput / StopTextInput take the
' edit box's address (WzPlatform drops the rectangle).
' Functions for the edit box widget.
' (VB) widget class W_EDITBOX : WIDGET
Option Explicit

' Edit Box states
Public Const WEDBS_FIXED As Long = &H1      ' No editing is going on
Public Const WEDBS_INSERT As Long = &H2     ' Insertion editing
Public Const WEDBS_OVER As Long = &H3       ' Overwrite editing
Public Const WEDBS_MASK As Long = &HF       '
Public Const WEDBS_HILITE As Long = &H10    '
Public Const WEDBS_DISABLE As Long = &H20   ' disable button from selection

' Pixel gap between edge of edit box and text
Private Const WEDB_XGAP As Long = 4

' Size of the overwrite cursor
Private Const WEDB_CURSORSIZE As Long = 8

' The time the cursor blinks for
Private Const WEDB_BLINKRATE As Long = 800

' Number of characters to jump the edit box text when moving the cursor
Private Const WEDB_CHARJUMP As Long = 6

' Max size for a string in a editbox
Private Const EB_MAX_STRINGSIZE As Long = 72

' W_EDBINIT::W_EDBINIT()
Public Sub W_EDBINIT_construct(init As W_EDBINIT)
    Dim z As W_EDBINIT
    init = z
    init.pText = vbNullString
    init.FontID = font_regular
    Set init.pBoxDisplay = Nothing
End Sub

' (VB) the W_INIT part of a W_EDBINIT
Public Function W_EDBINIT_toW_INIT(init As W_EDBINIT) As W_INIT
    Dim r As W_INIT
    r.formID = init.formID
    r.majorID = init.majorID
    r.id = init.id
    r.style = init.style
    r.x = init.x
    r.y = init.y
    r.width = init.width
    r.height = init.height
    Set r.pDisplay = init.pDisplay
    Set r.pCallback = init.pCallback
    WzFunc_assign r.pUserData, init.pUserData
    r.UserData = init.UserData
    Set r.calcLayout = init.calcLayout
    Set r.onDelete = init.onDelete
    Set r.customHitTest = init.customHitTest
    Set r.initPUserDataFunc = init.initPUserDataFunc
    W_EDBINIT_toW_INIT = r
End Function

' (VB) a WzUniCodepoint as VB text: one UTF-16 unit, or a surrogate pair above U+FFFF
Private Function codepointText(ByVal unicode As Long) As String
    If unicode < &H10000 Then
        codepointText = ChrW$(unicode And &HFFFF&)
    Else
        unicode = unicode - &H10000
        codepointText = ChrW$(&HD800& + (unicode \ &H400&)) & ChrW$(&HDC00& + (unicode And &H3FF&))
    End If
End Function

' std::make_shared<W_EDITBOX>(init)
Public Function W_EDITBOX_new_init(init As W_EDBINIT) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_EDITBOX)
    W_EDITBOX_construct_init w, init
    Set W_EDITBOX_new_init = w
End Function

' std::make_shared<W_EDITBOX>()
Public Function W_EDITBOX_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_W_EDITBOX)
    W_EDITBOX_construct w
    Set W_EDITBOX_new = w
End Function

' (VB) the member initialisers both constructors share
Private Sub W_EDITBOX_initMembers(ByVal self As WIDGET, ByVal FontID As iV_fonts, ByVal pBoxDisplay As WzFunc)
    Dim d As W_EDITBOX
    Set d = W_EDITBOX_(self)
    d.state = WEDBS_FIXED
    d.FontID = FontID
    d.blinkOffset = wzGetTicks()
    d.maxStringSize = EB_MAX_STRINGSIZE
    d.insPos = 0
    d.printStart = 0
    d.printChars = 0
    d.printWidth = 0
    Set d.pBoxDisplay = pBoxDisplay
    d.HilightAudioID = WidgGetHilightAudioID()
    d.ClickedAudioID = WidgGetClickedAudioID()
    d.ErrorAudioID = WidgGetErrorAudioID()
    Set d.AudioCallback = WidgGetAudioCallback()
    d.boxColourFirst = psPalette(WZCOL_FORM_DARK)
    d.boxColourSecond = psPalette(WZCOL_FORM_LIGHT)
    d.boxColourBackground = psPalette(WZCOL_FORM_BACKGROUND)
    d.suppressAudioCallback = False
    Set d.wzDisplayedText = New WzText  ' EditBoxDisplayCache displayCache
    Set d.modeText = New WzText
End Sub

' W_EDITBOX::W_EDITBOX(W_EDBINIT const *init)
Public Sub W_EDITBOX_construct_init(ByVal self As WIDGET, init As W_EDBINIT)
    Dim baseInit As W_INIT, text As String
    baseInit = W_EDBINIT_toW_INIT(init)
    WIDGET_construct_init self, baseInit, WIDG_EDITBOX
    self.kind = K_W_EDITBOX ' (VB) virtual calls from here on reach W_EDITBOX's overrides
    W_EDITBOX_initMembers self, init.FontID, init.pBoxDisplay

    text = init.pText
    If StrPtr(text) = 0 Then
        text = ""
    End If
    W_EDITBOX_(self).aText = text       ' WzString::fromUtf8(text)

    W_EDITBOX_initialise self

    If Not ((init.style And Not (WEDB_PLAIN Or WIDG_HIDDEN)) = 0) Then ASSERT_FAILURE "(init->style & ~(0 | WIDG_HIDDEN)) == 0", "lib/widget/editbox.cpp:85", "W_EDITBOX", "Unknown edit box style"
End Sub

' W_EDITBOX::W_EDITBOX()
Public Sub W_EDITBOX_construct(ByVal self As WIDGET)
    WIDGET_construct self
    self.kind = K_W_EDITBOX ' (VB) virtual calls from here on reach W_EDITBOX's overrides
    W_EDITBOX_initMembers self, font_regular, Nothing
End Sub

Public Sub W_EDITBOX_destruct(ByVal self As WIDGET)
    Dim editState As Long
    ' Note the edit state
    editState = W_EDITBOX_(self).state And WEDBS_MASK

    ' Only have anything to do if the widget is being edited
    If (editState And WEDBS_MASK) = WEDBS_FIXED Then
        Exit Sub
    End If

    ' If the edit box still somehow has focus, and is editable, need to StopTextInput()
    ' (May be able to remove this once more refactoring of the game menus / in-game UI occurs)
    debug_ 121, LOG_INFO, "~W_EDITBOX", "Editbox seems to still have focus, and is editable, as it's being destroyed."
    StopTextInput ObjPtr(self) ' force-stop text input if this EditBox somehow still has the input
End Sub

Private Sub W_EDITBOX_initialise(ByVal self As WIDGET)
    Dim d As W_EDITBOX
    Set d = W_EDITBOX_(self)
    d.state = WEDBS_FIXED
    d.printStart = 0
    d.maxStringSize = EB_MAX_STRINGSIZE
    W_EDITBOX_fitStringStart self
End Sub

' Insert a character into a text buffer
Private Function W_EDITBOX_insertChar(ByVal self As WIDGET, ByVal ch As Long) As Boolean
    Dim d As W_EDITBOX, chText As String
    Set d = W_EDITBOX_(self)
    If ch = 0 Then                      ' ch.isNull()
        W_EDITBOX_insertChar = False
        Exit Function
    End If

    If Not (d.insPos <= Len(d.aText)) Then ASSERT_FAILURE "insPos <= aText.length()", "lib/widget/editbox.cpp:142", "insertChar", "Invalid insertion point"
    If Len(d.aText) >= d.maxStringSize Then
        If Not d.AudioCallback Is Nothing Then
            WzFunc_call d.AudioCallback, d.ErrorAudioID
        End If
        W_EDITBOX_insertChar = False    ' string too big, just return
        Exit Function
    End If
    ' Move the end of the string up by one (including terminating \0)
    ' Insert the character
    chText = codepointText(ch)
    d.aText = Left$(d.aText, d.insPos) & chText & Mid$(d.aText, d.insPos + 1)

    ' Update the insertion point
    d.insPos = d.insPos + Len(chText)

    W_EDITBOX_insertChar = True
End Function

' Put a character into a text buffer overwriting any text under the cursor
Private Function W_EDITBOX_overwriteChar(ByVal self As WIDGET, ByVal ch As Long) As Boolean
    Dim d As W_EDITBOX, chText As String
    Set d = W_EDITBOX_(self)
    If ch = 0 Then                      ' ch.isNull()
        W_EDITBOX_overwriteChar = False
        Exit Function
    End If

    If Not (d.insPos <= Len(d.aText)) Then ASSERT_FAILURE "insPos <= aText.length()", "lib/widget/editbox.cpp:170", "overwriteChar", "overwriteChar: Invalid insertion point"
    self.dirty = True

    If d.insPos = Len(d.aText) Then
        ' At end of string.
        W_EDITBOX_overwriteChar = W_EDITBOX_insertChar(self, ch)
        Exit Function
    End If

    ' Store the character
    chText = codepointText(ch)
    d.aText = Left$(d.aText, d.insPos) & chText & Mid$(d.aText, d.insPos + 2)

    ' Update the insertion point
    d.insPos = d.insPos + Len(chText)

    W_EDITBOX_overwriteChar = True
End Function

' Delete a character to the right of the position
Private Sub W_EDITBOX_delCharRight(ByVal self As WIDGET)
    Dim d As W_EDITBOX
    Set d = W_EDITBOX_(self)
    If Not (d.insPos <= Len(d.aText)) Then ASSERT_FAILURE "insPos <= aText.length()", "lib/widget/editbox.cpp:192", "delCharRight", "Invalid deletion point"

    ' Can't delete if we are at the end of the string
    ' Move the end of the string down by one
    d.aText = Left$(d.aText, d.insPos) & Mid$(d.aText, d.insPos + 2)
End Sub

' Delete a character to the left of the position
Private Sub W_EDITBOX_delCharLeft(ByVal self As WIDGET)
    Dim d As W_EDITBOX
    Set d = W_EDITBOX_(self)
    ' Can't delete if we are at the start of the string
    If d.insPos = 0 Then
        Exit Sub
    End If

    d.insPos = d.insPos - 1
    W_EDITBOX_delCharRight self
End Sub

Public Sub W_EDITBOX_geometryChanged(ByVal self As WIDGET)
    Dim editState As Long
    ' Note the edit state
    editState = W_EDITBOX_(self).state And WEDBS_MASK

    ' For now, only handle fit recalculation if not being edited
    If Not ((editState And WEDBS_MASK) = WEDBS_FIXED) Then
        Exit Sub
    End If
    W_EDITBOX_fitStringStart self
End Sub

' Calculate how much of the start of a string can fit into the edit box
Private Sub W_EDITBOX_fitStringStart(ByVal self As WIDGET)
    Dim d As W_EDITBOX, tmp As String, pixelWidth As Long
    Set d = W_EDITBOX_(self)
    ' We need to calculate the whole string's pixel size.
    ' From QuesoGLC's notes: additional processing like kerning creates strings of text whose dimensions are not directly
    ' related to the simple juxtaposition of individual glyph metrics. For example, the advance width of "VA" isn't the
    ' sum of the advances of "V" and "A" taken separately.
    tmp = Mid$(d.aText, d.printStart + 1)   ' Ignore the first printStart characters.

    Do While Len(tmp) > 0
        pixelWidth = iV_GetTextWidth(tmp, d.FontID)

        If pixelWidth <= self.width() - (WEDB_XGAP * 2 + WEDB_CURSORSIZE) Then
            d.printChars = Len(tmp)
            d.printWidth = pixelWidth
            Exit Sub
        End If

        tmp = Left$(tmp, Len(tmp) - 1)  ' Erase last char.
    Loop

    d.printChars = 0
    d.printWidth = 0
End Sub

' Calculate how much of the end of a string can fit into the edit box
Private Sub W_EDITBOX_fitStringEnd(ByVal self As WIDGET)
    Dim d As W_EDITBOX, tmp As String, pixelWidth As Long
    Set d = W_EDITBOX_(self)
    tmp = d.aText

    d.printStart = 0

    Do While Len(tmp) > 0
        pixelWidth = iV_GetTextWidth(tmp, d.FontID)

        If pixelWidth <= self.width() - (WEDB_XGAP * 2 + WEDB_CURSORSIZE) Then
            d.printChars = Len(tmp)
            d.printWidth = pixelWidth
            Exit Sub
        End If

        tmp = Mid$(tmp, 2)              ' Erase first char.
        d.printStart = d.printStart + 1
    Loop

    d.printChars = 0
    d.printWidth = 0
End Sub

Private Sub W_EDITBOX_setCursorPosPixels(ByVal self As WIDGET, ByVal xPos As Long)
    Dim d As W_EDITBOX, tmp As String, prevDelta As Long, prevPos As Long, pixelWidth As Long, delta As Long, pos As Long
    Set d = W_EDITBOX_(self)
    tmp = Mid$(d.aText, d.printStart + 1)   ' Consider only the visible text.
    tmp = Left$(tmp, d.printChars)

    prevDelta = &H7FFFFFFF              ' INT32_MAX
    prevPos = d.printStart + Len(tmp)
    Do While Len(tmp) > 0
        pixelWidth = iV_GetTextWidth(tmp, d.FontID)
        delta = pixelWidth - (xPos - (WEDB_XGAP + WEDB_CURSORSIZE \ 2))
        pos = d.printStart + Len(tmp)

        If delta <= 0 Then
            If -delta < prevDelta Then
                d.insPos = pos
            Else
                d.insPos = prevPos
            End If
            Exit Sub
        End If

        tmp = Left$(tmp, Len(tmp) - 1)  ' Erase last char.

        prevDelta = delta
        prevPos = pos
    Loop

    d.insPos = d.printStart
End Sub

Public Sub W_EDITBOX_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim d As W_EDITBOX, editState As Long, mx As Long, my As Long, lockedScreen As W_SCREEN
    Dim done As Boolean, unicode As Long, key As Long, length As Long, changedText As Boolean
    Set d = W_EDITBOX_(self)
    ' Note the edit state
    editState = d.state And WEDBS_MASK

    ' Only have anything to do if the widget is being edited
    If (editState And WEDBS_MASK) = WEDBS_FIXED Then
        Exit Sub
    End If
    self.dirty = True
    StartTextInput ObjPtr(self)
    ' If there is a mouse click outside of the edit box - stop editing
    mx = psContext.mx
    my = psContext.my
    If mousePressed(MOUSE_LMB) And Not WzRect_contains(self.geometry(), mx, my) Then
        StopTextInput ObjPtr(self)
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            lockedScreen.setFocus Nothing
        End If
        Exit Sub
    End If

    ' Loop through the characters in the input buffer
    done = False
    key = inputGetKey(unicode)
    Do While key <> 0 And Not done
        ' Don't blink while typing.
        d.blinkOffset = wzGetTicks()

        length = 0

        ' Deal with all the control keys, assume anything else is a printable character
        Select Case key
        Case INPBUF_LEFT
            ' Move the cursor left
            d.insPos = MAX_(d.insPos - 1, 0)

            ' If the cursor has gone off the left of the edit box,
            ' need to update the printable text.
            If d.insPos < d.printStart Then
                d.printStart = MAX_(d.printStart - WEDB_CHARJUMP, 0)
                W_EDITBOX_fitStringStart self
            End If
            debug_ 363, LOG_INPUT, "run", "EditBox cursor left"
        Case INPBUF_RIGHT
            ' Move the cursor right
            length = Len(d.aText)
            d.insPos = MIN_(d.insPos + 1, length)

            ' If the cursor has gone off the right of the edit box,
            ' need to update the printable text.
            If d.insPos > d.printStart + d.printChars Then
                d.printStart = MIN_(d.printStart + WEDB_CHARJUMP, length - 1)
                W_EDITBOX_fitStringStart self
            End If
            debug_ 378, LOG_INPUT, "run", "EditBox cursor right (%d, %d, %d)", d.insPos, d.printStart, d.printChars
        Case INPBUF_UP
            debug_ 381, LOG_INPUT, "run", "EditBox cursor up"
        Case INPBUF_DOWN
            debug_ 384, LOG_INPUT, "run", "EditBox cursor down"
        Case INPBUF_HOME
            ' Move the cursor to the start of the buffer
            d.insPos = 0
            d.printStart = 0
            W_EDITBOX_fitStringStart self
            debug_ 391, LOG_INPUT, "run", "EditBox cursor home"
        Case INPBUF_END
            ' Move the cursor to the end of the buffer
            d.insPos = Len(d.aText)
            If d.insPos <> d.printStart + d.printChars Then
                W_EDITBOX_fitStringEnd self
            End If
            debug_ 400, LOG_INPUT, "run", "EditBox cursor end"
        Case INPBUF_INS
            If editState = WEDBS_INSERT Then
                editState = WEDBS_OVER
            Else
                editState = WEDBS_INSERT
            End If
            debug_ 411, LOG_INPUT, "run", "EditBox cursor insert"
        Case INPBUF_DEL
            W_EDITBOX_delCharRight self

            ' Update the printable text
            W_EDITBOX_fitStringStart self
            debug_ 418, LOG_INPUT, "run", "EditBox cursor delete"
        Case INPBUF_PGUP
            debug_ 421, LOG_INPUT, "run", "EditBox cursor page up"
        Case INPBUF_PGDN
            debug_ 424, LOG_INPUT, "run", "EditBox cursor page down"
        Case INPBUF_BKSPACE
            ' Delete the character to the left of the cursor
            W_EDITBOX_delCharLeft self

            ' Update the printable text
            If d.insPos <= d.printStart Then
                d.printStart = MAX_(d.printStart - WEDB_CHARJUMP, 0)
            End If
            W_EDITBOX_fitStringStart self
            debug_ 436, LOG_INPUT, "run", "EditBox cursor backspace"
        Case INPBUF_TAB
            debug_ 439, LOG_INPUT, "run", "EditBox cursor tab"
            If Not d.onTabHandler Is Nothing Then
                If CBool(WzFunc_call(d.onTabHandler, self)) Then
                    done = True
                End If
            End If
        Case INPBUF_CR, KEY_KPENTER     ' either normal return key || keypad enter
            ' Finish editing
            StopTextInput ObjPtr(self)
            If Not d.onRetHandler Is Nothing Then
                WzFunc_call d.onRetHandler, self
            End If
            Set lockedScreen = self.screenPointer
            If Not lockedScreen Is Nothing Then
                lockedScreen.setFocus Nothing
            End If
            W_EDITBOX_stopEditing self
            debug_ 461, LOG_INPUT, "run", "EditBox cursor return"
            Exit Sub
        Case INPBUF_ESC
            debug_ 465, LOG_INPUT, "run", "EditBox cursor escape"
            If Len(d.aText) > 0 Then
                ' hitting ESC while the editbox contains text clears the text
                d.aText = ""
                d.insPos = 0
                d.printStart = 0
                W_EDITBOX_fitStringStart self
                inputLoseFocus  ' clear the input buffer.
            Else
                ' hitting ESC while the editbox is empty ends editing mode
                Set lockedScreen = self.screenPointer
                If Not lockedScreen Is Nothing Then
                    lockedScreen.setFocus Nothing
                End If
                inputLoseFocus  ' clear the input buffer.
                W_EDITBOX_stopEditing self
                If Not d.onEscHandler Is Nothing Then
                    WzFunc_call d.onEscHandler, self
                End If
                Exit Sub
            End If

        Case Else
            If keyDown(KEY_LCTRL) Or keyDown(KEY_RCTRL) Then
                Select Case key
                Case KEY_V
                    d.aText = wzGetSelection()
                    If Not d.onPasteTransformFunc Is Nothing Then
                        ' call custom onPasteTransformFunc to process the pasted data
                        d.aText = CStr(WzFunc_call(d.onPasteTransformFunc, d.aText))
                    Else
                        ' default behavior:
                        ' remove any \r, \n chars
                        d.aText = Replace$(d.aText, vbCr, "")
                        d.aText = Replace$(d.aText, vbLf, "")
                    End If
                    ' truncate if too long
                    If Len(d.aText) >= d.maxStringSize Then
                        d.aText = Left$(d.aText, d.maxStringSize)
                    End If
                    d.insPos = Len(d.aText)
                    ' Update the printable text
                    W_EDITBOX_fitStringEnd self
                    debug_ 519, LOG_INPUT, "run", "EditBox paste"
                Case Else
                End Select
            Else
                ' Dealt with everything else this must be a printable character
                changedText = False
                If editState = WEDBS_INSERT Then
                    changedText = W_EDITBOX_insertChar(self, unicode)
                Else
                    changedText = W_EDITBOX_overwriteChar(self, unicode)
                End If
                If changedText Then
                    length = Len(d.aText)
                    ' Update the printable chars
                    If d.insPos = length Then
                        W_EDITBOX_fitStringEnd self
                    Else
                        W_EDITBOX_fitStringStart self
                        If d.insPos > d.printStart + d.printChars Then
                            d.printStart = MIN_(d.printStart + WEDB_CHARJUMP, length - 1)
                            If d.printStart >= length Then
                                W_EDITBOX_fitStringStart self
                            End If
                        End If
                    End If
                End If
            End If
        End Select
        key = inputGetKey(unicode)
    Loop

    ' Store the current widget state
    d.state = (d.state And Not WEDBS_MASK) Or editState
End Sub

Public Function W_EDITBOX_getString(ByVal self As WIDGET) As String
    W_EDITBOX_getString = W_EDITBOX_(self).aText
End Function

' Set the current string for the edit box
Public Sub W_EDITBOX_setString(ByVal self As WIDGET, ByVal string_ As String)
    W_EDITBOX_(self).aText = string_
    W_EDITBOX_initialise self
    self.dirty = True
End Sub

Public Sub W_EDITBOX_setPlaceholder(ByVal self As WIDGET, ByVal value As String)
    W_EDITBOX_(self).placeholderText = value
    self.dirty = True
End Sub

' void setPlaceholderTextColor(optional<PIELIGHT> fixedPlaceholderTextColor)
Public Sub W_EDITBOX_setPlaceholderTextColor(ByVal self As WIDGET, ByVal has_fixedPlaceholderTextColor As Boolean, Optional ByVal fixedPlaceholderTextColor As Long = 0)
    W_EDITBOX_(self).has_fixedPlaceholderTextColor = has_fixedPlaceholderTextColor
    W_EDITBOX_(self).fixedPlaceholderTextColor = fixedPlaceholderTextColor
End Sub

Public Sub W_EDITBOX_simulateClick(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal silenceClickAudio As Boolean = False, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    If silenceClickAudio Then
        W_EDITBOX_(self).suppressAudioCallback = True
    End If
    self.clicked psContext, key
    If silenceClickAudio Then
        W_EDITBOX_(self).suppressAudioCallback = False
    End If
End Sub

' Respond to a mouse click
Public Sub W_EDITBOX_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As W_EDITBOX, lockedScreen As W_SCREEN
    Set d = W_EDITBOX_(self)
    If (d.state And WEDBS_DISABLE) <> 0 Then ' disabled button.
        Exit Sub
    End If

    ' Set cursor position to the click location.
    W_EDITBOX_setCursorPosPixels self, psContext.mx - self.x()

    ' Cursor should be visible instantly.
    d.blinkOffset = wzGetTicks()

    If (d.state And WEDBS_MASK) = WEDBS_FIXED Then
        If Not d.AudioCallback Is Nothing And Not d.suppressAudioCallback Then
            WzFunc_call d.AudioCallback, d.ClickedAudioID
        End If

        ' Set up the widget state
        d.state = (d.state And Not WEDBS_MASK) Or WEDBS_INSERT

        ' Calculate how much of the string can appear in the box
        W_EDITBOX_fitStringEnd self
        If d.printStart > 0 Then
            d.insPos = Len(d.aText)
        End If

        ' Clear the input buffer
        inputClearBuffer

        ' Tell the form that the edit box has focus
        Set lockedScreen = self.screenPointer
        If Not lockedScreen Is Nothing Then
            lockedScreen.setFocus self
        Else
            ' If the W_EDITBOX isn't currently attached to a screen when this is triggered, focus issues may occur
            ASSERT_FAILURE "false", "lib/widget/editbox.cpp:644", "clicked", "W_EDITBOX is not attached to any screen?"
        End If
    End If
    self.dirty = True
End Sub

Public Sub W_EDITBOX_stopEditing(ByVal self As WIDGET)
    Dim d As W_EDITBOX, editState As Long
    Set d = W_EDITBOX_(self)
    If (d.state And WEDBS_DISABLE) <> 0 Then ' disabled button.
        Exit Sub
    End If

    ' Note the edit state
    editState = d.state And WEDBS_MASK

    ' Only have anything to do if the widget is being edited
    If (editState And WEDBS_MASK) = WEDBS_FIXED Then
        Exit Sub
    End If

    ' Stop editing the widget
    d.state = WEDBS_FIXED
    d.printStart = 0
    W_EDITBOX_fitStringStart self
    StopTextInput ObjPtr(self)
    self.dirty = True
    If Not d.onEditingStoppedHandler Is Nothing Then
        WzFunc_call d.onEditingStoppedHandler, self
    End If
End Sub

Public Function W_EDITBOX_isEditing(ByVal self As WIDGET) As Boolean
    Dim editState As Long
    editState = W_EDITBOX_(self).state And WEDBS_MASK
    W_EDITBOX_isEditing = (editState And WEDBS_MASK) <> WEDBS_FIXED
End Function

' Respond to loss of focus
Public Sub W_EDITBOX_focusLost(ByVal self As WIDGET)
    Dim lockedScreen As W_SCREEN
    If Not (Not ((W_EDITBOX_(self).state And WEDBS_DISABLE) <> 0)) Then ASSERT_FAILURE "!(state & 0x0020)", "lib/widget/editbox.cpp:688", "focusLost", "editBoxFocusLost: disabled edit box"

    W_EDITBOX_stopEditing self

    Set lockedScreen = self.screenPointer
    If Not lockedScreen Is Nothing Then
        lockedScreen.setReturn self
    End If
End Sub

' Respond to a mouse moving over an edit box
Public Sub W_EDITBOX_highlight(ByVal self As WIDGET, psContext As W_CONTEXT)
    Dim psWidget As W_EDITBOX
    Set psWidget = W_EDITBOX_(self)
    If (psWidget.state And WEDBS_DISABLE) <> 0 Then
        Exit Sub
    End If

    If Not psWidget.AudioCallback Is Nothing Then
        WzFunc_call psWidget.AudioCallback, psWidget.HilightAudioID
    End If

    psWidget.state = psWidget.state Or WEDBS_HILITE
End Sub

' Respond to the mouse moving off an edit box
Public Sub W_EDITBOX_highlightLost(ByVal self As WIDGET)
    Dim psWidget As W_EDITBOX
    Set psWidget = W_EDITBOX_(self)
    If (psWidget.state And WEDBS_DISABLE) <> 0 Then
        Exit Sub
    End If

    psWidget.state = psWidget.state And WEDBS_MASK
End Sub

Public Sub W_EDITBOX_setTip(ByVal self As WIDGET, ByVal string_ As String)
    W_EDITBOX_(self).pTip = string_
End Sub

Public Function W_EDITBOX_getTip(ByVal self As WIDGET) As String
    W_EDITBOX_getTip = W_EDITBOX_(self).pTip
End Function

Public Sub W_EDITBOX_setBoxColours(ByVal self As WIDGET, ByVal first As Long, ByVal second As Long, ByVal background As Long)
    W_EDITBOX_(self).boxColourFirst = first
    W_EDITBOX_(self).boxColourSecond = second
    W_EDITBOX_(self).boxColourBackground = background
End Sub

Public Sub W_EDITBOX_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As W_EDITBOX, x0 As Long, y0 As Long, x1 As Long, y1 As Long, displayedText As String, displayedTextColor As Long
    Dim lineSize As Long, aboveBase As Long, belowBase As Long, fx As Long, fy As Long
    Dim visibleTextBeforeCursor As String, cursorX As Long, cursorY As Long
    Set d = W_EDITBOX_(self)
    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    x1 = x0 + self.width()
    y1 = y0 + self.height()

    If Not d.pBoxDisplay Is Nothing Then
        WzFunc_call d.pBoxDisplay, self, xOffset, yOffset
    Else
        iV_ShadowBox x0, y0, x1, y1, 0, d.boxColourFirst, d.boxColourSecond, d.boxColourBackground
    End If

    ' If there is more text than will fit into the box, display the bit with the cursor in it
    displayedText = Mid$(d.aText, d.printStart + 1)     ' Erase anything there isn't room to display.
    displayedText = Left$(displayedText, d.printChars)

    displayedTextColor = psPalette(WZCOL_FORM_TEXT)
    If Len(d.aText) = 0 And Len(d.placeholderText) > 0 Then
        d.wzDisplayedText.setText d.placeholderText, d.FontID
        If d.has_fixedPlaceholderTextColor Then
            displayedTextColor = d.fixedPlaceholderTextColor
        ElseIf (d.state And WEDBS_MASK) = WEDBS_FIXED Then
            displayedTextColor = psPalette(WZCOL_FORM_TEXT)
        Else
            displayedTextColor = psPalette(WZCOL_GREY)
        End If
    Else
        d.wzDisplayedText.setText displayedText, d.FontID
    End If
    If (d.state And WEDBS_DISABLE) <> 0 Then
        displayedTextColor = psPalette(WZCOL_TEXT_DARK)
    End If

    lineSize = d.wzDisplayedText.lineSize()
    aboveBase = d.wzDisplayedText.aboveBase()
    belowBase = d.wzDisplayedText.belowBase()

    fx = x0 + WEDB_XGAP                 ' + (psEdBox->width - fw) / 2;
    fy = y0 + (self.height() - lineSize) \ 2 - aboveBase

    d.wzDisplayedText.render fx, fy, displayedTextColor

    ' Display the cursor if editing
    If ((wzGetTicks() - d.blinkOffset) \ WEDB_BLINKRATE) Mod 2 = 0 Then
        If d.insPos - d.printStart >= 0 Then
            visibleTextBeforeCursor = Mid$(d.aText, d.printStart + 1, d.insPos - d.printStart)
        Else
            visibleTextBeforeCursor = Mid$(d.aText, d.printStart + 1)  ' (a negative size_t length: the rest)
        End If
        d.modeText.setText visibleTextBeforeCursor, d.FontID
        cursorX = x0 + WEDB_XGAP + d.modeText.width()
        cursorY = fy

        If (d.state And WEDBS_MASK) = WEDBS_INSERT Then
            ' insert mode
            iV_Line cursorX, cursorY + aboveBase, cursorX, cursorY - belowBase, psPalette(WZCOL_FORM_CURSOR)
        ElseIf (d.state And WEDBS_MASK) = WEDBS_OVER Then
            iV_Line cursorX, cursorY, cursorX + WEDB_CURSORSIZE, cursorY, psPalette(WZCOL_FORM_CURSOR)
        End If
    End If

    If d.pBoxDisplay Is Nothing Then
        If (d.state And WEDBS_HILITE) <> 0 Then
            ' Display the button hilite
            iV_Box x0 - 2, y0 - 2, x1 + 2, y1 + 2, psPalette(WZCOL_FORM_HILITE)
        End If
    End If
End Sub

Public Sub W_EDITBOX_setMaxStringSize(ByVal self As WIDGET, ByVal size As Long)
    W_EDITBOX_(self).maxStringSize = size
End Sub

Public Sub W_EDITBOX_setState(ByVal self As WIDGET, ByVal newState As Long)
    Dim mask As Long
    mask = WEDBS_DISABLE
    W_EDITBOX_(self).state = (W_EDITBOX_(self).state And Not mask) Or (newState And mask)
End Sub

Public Sub W_EDITBOX_setOnReturnHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Set W_EDITBOX_(self).onRetHandler = func
End Sub

Public Sub W_EDITBOX_setOnTabHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Set W_EDITBOX_(self).onTabHandler = func
End Sub

Public Sub W_EDITBOX_setOnEscapeHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Set W_EDITBOX_(self).onEscHandler = func
End Sub

Public Sub W_EDITBOX_setOnEditingStoppedHandler(ByVal self As WIDGET, ByVal func As WzFunc)
    Set W_EDITBOX_(self).onEditingStoppedHandler = func
End Sub

Public Sub W_EDITBOX_setOnPasteTransformFunc(ByVal self As WIDGET, ByVal func As WzFunc)
    Set W_EDITBOX_(self).onPasteTransformFunc = func
End Sub
