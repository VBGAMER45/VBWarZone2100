Attribute VB_Name = "paragraph_c"
' lib/widget/paragraph.cpp, paragraph.h: the word wrapping FlowLayout, the
' paragraph elements and the widget classes ParagraphTextWidget (of
' paragraph.cpp) and Paragraph (docs/CONVENTIONS.md, "Widgets"; their fields are
' ParagraphTextWidget.cls and Paragraph.cls). paragraph.h's WzCachedText is
' WzCachedText.cls (paragraph_h; operator-> is deref).
' (VB) The classes that are not widgets: struct ParagraphTextStyle and
' FlowLayoutFragment are UDTs (ParagraphTextStyle_construct is the default
' constructor); ParagraphElement and its two implementations
' (ParagraphTextElement, ParagraphWidgetElement) are ParagraphElement.cls, its
' kind saying which, its virtual methods ParagraphElement_<method> here;
' FlowLayoutElementDescriptor and its implementations (FlowLayoutStringDescriptor,
' ParagraphWidgetElement) are the UDT FlowLayoutElementDescriptor with a kind;
' struct FlowLayout is a UDT (its lines flattened: linesFrag, lines starting at
' lineStart(i)) and FlowLayout_<method>. A text position is a UTF-16 unit of the
' VB String (WzString's code point for text in the Basic Multilingual Plane):
' substr is Mid$, toUtf32 the String's units. W_PARAGRAPH_ONCLICK_FUNC is a
' WzFunc called with the paragraph and the key.
' (VB) widget class ParagraphTextWidget : WIDGET
' (VB) widget class Paragraph : WIDGET
Option Explicit

' Provide information about an element appended to the paragraph. (struct ParagraphElement: its kinds)
Public Const PARAGRAPH_TEXT_ELEMENT As Long = 0     ' struct ParagraphTextElement
Public Const PARAGRAPH_WIDGET_ELEMENT As Long = 1   ' struct ParagraphWidgetElement

Public Type ParagraphTextStyle
    font As iV_fonts                    ' = font_regular
    shadeColour As Long                 ' PIELIGHT = pal_RGBA(0, 0, 0, 0)
    fontColour As Long                  ' PIELIGHT = WZCOL_BLACK
End Type

Public Type FlowLayoutFragment
    elementId As Long                   ' size_t
    begin As Long                       ' size_t
    length As Long                      ' size_t
    width As Long                       ' unsigned int
    offset As Long                      ' unsigned int
End Type

' (VB) FlowLayoutElementDescriptor: kind 0 is a FlowLayoutStringDescriptor, 1 a ParagraphWidgetElement
Private Const DESCRIPTOR_STRING As Long = 0
Private Const DESCRIPTOR_WIDGET As Long = 1

Private Type FlowLayoutElementDescriptor
    kind As Long
    text As String                      ' FlowLayoutStringDescriptor: WzString text
    textUtf32() As Long                 ' std::vector<uint32_t> textUtf32
    textUtf32Size As Long
    rightToLeft As Boolean              ' = false
    font As iV_fonts
    widget As WIDGET                    ' ParagraphWidgetElement: std::shared_ptr<WIDGET> widget
End Type

' FlowLayout implements a word wrapping algorithm.
'
' It doesn't specify how the text should be structured, it only requires that the elements of the text are described
' by FlowLayoutElementDescriptor.
'
' Each element can contain 0 or more words, and a single word can be part of two sequential elements.
'
' Example:
' - Element 1 = "first sec"
' - Element 2 = "ond third"
'
' In the example above, the resulting text is "first second third", so if possible the algorithm will keep
' the second word completely in the same line.
Private Type FlowLayout
    currentLine() As FlowLayoutFragment ' std::vector<FlowLayoutFragment> currentLine
    currentLineCount As Long
    linesFrag() As FlowLayoutFragment   ' std::vector<std::vector<FlowLayoutFragment>> lines: line i is
    linesFragCount As Long              ' linesFrag(lineStart(i)) .. linesFrag(lineStart(i + 1) - 1)
    lineStart() As Long
    linesCount As Long
    maxWidth As Long                    ' unsigned int
    currentElementId As Long            ' unsigned int = 0
    nextOffset As Long                  ' unsigned int = 0
    partialWord() As FlowLayoutFragment ' std::vector<FlowLayoutFragment> partialWord
    partialWordCount As Long
    partialWordWidth As Long            ' unsigned int = 0
End Type

' ParagraphTextStyle's member initialisers
Public Sub ParagraphTextStyle_construct(s As ParagraphTextStyle)
    s.font = font_regular
    s.shadeColour = pal_RGBA(0, 0, 0, 0)
    s.fontColour = psPalette(WZCOL_BLACK)
End Sub

'--------------------------------------------------------------------------
' class FlowLayoutElementDescriptor

Private Function FlowLayoutElementDescriptor_getWidth(desc As FlowLayoutElementDescriptor, ByVal position As Long, ByVal length As Long) As Long
    Select Case desc.kind
    Case DESCRIPTOR_STRING
        FlowLayoutElementDescriptor_getWidth = iV_GetTextWidth(Mid$(desc.text, position + 1, length), desc.font)
    Case DESCRIPTOR_WIDGET
        If position = 0 Then
            FlowLayoutElementDescriptor_getWidth = desc.widget.width()
        Else
            FlowLayoutElementDescriptor_getWidth = 0
        End If
    End Select
End Function

Private Function FlowLayoutElementDescriptor_getMaxElementLengthForWidth(desc As FlowLayoutElementDescriptor, ByVal position As Long, ByVal maxWidth As Long) As Long
    Select Case desc.kind
    Case DESCRIPTOR_STRING
        FlowLayoutElementDescriptor_getMaxElementLengthForWidth = iV_GetMaxTextRunLenForWidth(Mid$(desc.text, position + 1), desc.font, maxWidth, desc.rightToLeft)
    Case DESCRIPTOR_WIDGET
        If position = 0 Then
            If Not ULT(maxWidth, desc.widget.width()) Then
                FlowLayoutElementDescriptor_getMaxElementLengthForWidth = FlowLayoutElementDescriptor_size(desc)
            Else
                FlowLayoutElementDescriptor_getMaxElementLengthForWidth = 0
            End If
            Exit Function
        End If
        If position > FlowLayoutElementDescriptor_size(desc) Then
            FlowLayoutElementDescriptor_getMaxElementLengthForWidth = 0
            Exit Function
        End If
        FlowLayoutElementDescriptor_getMaxElementLengthForWidth = FlowLayoutElementDescriptor_size(desc) - position
    End Select
End Function

Private Function FlowLayoutElementDescriptor_size(desc As FlowLayoutElementDescriptor) As Long
    Select Case desc.kind
    Case DESCRIPTOR_STRING
        FlowLayoutElementDescriptor_size = desc.textUtf32Size
    Case DESCRIPTOR_WIDGET
        FlowLayoutElementDescriptor_size = 2    ' + 1 for "whitespace" at end
    End Select
End Function

Private Function FlowLayoutElementDescriptor_isWhitespace(desc As FlowLayoutElementDescriptor, ByVal position As Long) As Boolean
    Select Case desc.kind
    Case DESCRIPTOR_STRING
        Select Case desc.textUtf32(position)
        Case 32, 9                      ' ' ', '\t'
            FlowLayoutElementDescriptor_isWhitespace = True
        Case Else
            FlowLayoutElementDescriptor_isWhitespace = False
        End Select
    Case DESCRIPTOR_WIDGET
        FlowLayoutElementDescriptor_isWhitespace = (position > 0)
    End Select
End Function

Private Function FlowLayoutElementDescriptor_isLineBreak(desc As FlowLayoutElementDescriptor, ByVal position As Long) As Boolean
    Select Case desc.kind
    Case DESCRIPTOR_STRING
        FlowLayoutElementDescriptor_isLineBreak = desc.textUtf32(position) = 10     ' '\n'
    Case DESCRIPTOR_WIDGET
        FlowLayoutElementDescriptor_isLineBreak = False
    End Select
End Function

Private Function FlowLayoutElementDescriptor_isWord(desc As FlowLayoutElementDescriptor, ByVal position As Long) As Boolean
    If position < FlowLayoutElementDescriptor_size(desc) Then
        If Not FlowLayoutElementDescriptor_isWhitespace(desc, position) Then
            FlowLayoutElementDescriptor_isWord = Not FlowLayoutElementDescriptor_isLineBreak(desc, position)
        End If
    End If
End Function

' FlowLayoutStringDescriptor(WzString const &newText, iV_fonts newFont, bool rtl = false)
Private Sub FlowLayoutStringDescriptor_construct(desc As FlowLayoutElementDescriptor, ByVal newText As String, ByVal newFont As iV_fonts, Optional ByVal rtl As Boolean = False)
    Dim i As Long
    desc.kind = DESCRIPTOR_STRING
    desc.text = newText
    desc.textUtf32Size = Len(newText)   ' newText.toUtf32()
    If desc.textUtf32Size > 0 Then
        ReDim desc.textUtf32(0 To desc.textUtf32Size - 1)
        For i = 0 To desc.textUtf32Size - 1
            desc.textUtf32(i) = AscW(Mid$(newText, i + 1, 1)) And &HFFFF&
        Next
    End If
    desc.rightToLeft = rtl
    desc.font = newFont
End Sub

'--------------------------------------------------------------------------
' struct FlowLayout

Private Sub FlowLayoutFragments_push_back(arr() As FlowLayoutFragment, n As Long, f As FlowLayoutFragment)
    If n = 0 Then
        ReDim arr(0 To 7)
    ElseIf n > UBound(arr) Then
        ReDim Preserve arr(0 To 2 * n - 1)
    End If
    arr(n) = f
    n = n + 1
End Sub

Private Function FlowLayout_shouldMergeFragment(layout As FlowLayout) As Boolean
    If layout.currentLineCount > 0 And layout.partialWordCount > 0 Then
        FlowLayout_shouldMergeFragment = layout.currentLine(layout.currentLineCount - 1).elementId = layout.partialWord(0).elementId
    End If
End Function

Private Sub FlowLayout_endWord(layout As FlowLayout)
    Dim currentLineBack As FlowLayoutFragment, i As Long
    If FlowLayout_shouldMergeFragment(layout) Then
        ' auto partialWordFront = &partialWord.front();
        currentLineBack = layout.currentLine(layout.currentLineCount - 1)
        layout.currentLineCount = layout.currentLineCount - 1
        layout.partialWord(0).length = layout.partialWord(0).length + layout.partialWord(0).begin - currentLineBack.begin
        layout.partialWord(0).width = layout.partialWord(0).offset + layout.partialWord(0).width - currentLineBack.offset
        layout.partialWord(0).begin = currentLineBack.begin
        layout.partialWord(0).offset = currentLineBack.offset
    End If

    For i = 0 To layout.partialWordCount - 1
        FlowLayoutFragments_push_back layout.currentLine, layout.currentLineCount, layout.partialWord(i)
    Next

    layout.nextOffset = layout.nextOffset + layout.partialWordWidth
    layout.partialWordWidth = 0
    layout.partialWordCount = 0
End Sub

Private Sub FlowLayout_pushFragment(layout As FlowLayout, elementDescriptor As FlowLayoutElementDescriptor, ByVal begin As Long, ByVal end_ As Long)
    Dim width As Long, f As FlowLayoutFragment
    width = FlowLayoutElementDescriptor_getWidth(elementDescriptor, begin, end_ - begin)
    f.elementId = layout.currentElementId
    f.begin = begin
    f.length = end_ - begin
    f.width = width
    f.offset = layout.nextOffset + layout.partialWordWidth
    FlowLayoutFragments_push_back layout.partialWord, layout.partialWordCount, f
    layout.partialWordWidth = layout.partialWordWidth + width
End Sub

Private Sub FlowLayout_placeLine(layout As FlowLayout, elementDescriptor As FlowLayoutElementDescriptor, ByVal begin As Long, ByVal end_ As Long)
    Dim current As Long, elementDescriptorSize As Long, spaceUsedForCurrentLine As Long, fitWidth As Long
    Dim fragmentFits As Long, whitespacePosition As Long, fragmentEnd As Long, availableWidth As Long
    current = begin
    elementDescriptorSize = FlowLayoutElementDescriptor_size(elementDescriptor)

    Do While current < end_
        spaceUsedForCurrentLine = layout.nextOffset + layout.partialWordWidth
        If Not ULT(layout.maxWidth, spaceUsedForCurrentLine) Then
            availableWidth = layout.maxWidth - spaceUsedForCurrentLine
        Else
            availableWidth = 0
        End If
        fitWidth = FlowLayoutElementDescriptor_getMaxElementLengthForWidth(elementDescriptor, current, availableWidth)
        fragmentFits = 0
        If fitWidth < end_ - current Then
            fragmentFits = current + fitWidth
        Else
            fragmentFits = end_
        End If

        If fragmentFits >= elementDescriptorSize Then
            whitespacePosition = elementDescriptorSize
        Else
            whitespacePosition = fragmentFits + 1
        End If
        Do While whitespacePosition > current
            If FlowLayoutElementDescriptor_isWhitespace(elementDescriptor, whitespacePosition - 1) Then Exit Do
            whitespacePosition = whitespacePosition - 1
        Loop

        If whitespacePosition > current Then
            ' the fragment ending with a whitespace fits within the line
            FlowLayout_pushFragment layout, elementDescriptor, current, whitespacePosition - 1
            FlowLayout_endWord layout
            layout.nextOffset = layout.nextOffset + FlowLayoutElementDescriptor_getWidth(elementDescriptor, whitespacePosition - 1, 1)
            current = whitespacePosition
        ElseIf fragmentFits = end_ Then
            ' the fragment is a single word, and fits in the line,
            ' but it doesn't end in whitespace and the next element might be part of this word
            FlowLayout_pushFragment layout, elementDescriptor, current, fragmentFits
            current = fragmentFits
        ElseIf layout.nextOffset > 0 Then
            ' word doesn't fit in the current line
            FlowLayout_breakLine layout
        Else
            ' word doesn't fit in an empty line
            fragmentEnd = MAX_(fragmentFits, current + 1)
            FlowLayout_pushFragment layout, elementDescriptor, current, fragmentEnd
            current = fragmentEnd
            FlowLayout_breakLine layout
        End If
    Loop
End Sub

' FlowLayout(unsigned int maxWidth)
Private Sub FlowLayout_construct(layout As FlowLayout, ByVal maxWidth As Long)
    layout.maxWidth = maxWidth
    layout.currentElementId = 0
    layout.nextOffset = 0
    layout.partialWordWidth = 0
End Sub

Private Sub FlowLayout_append(layout As FlowLayout, elementDescriptor As FlowLayoutElementDescriptor)
    Dim position As Long, lineEnd As Long
    position = 0
    Do While position < FlowLayoutElementDescriptor_size(elementDescriptor)
        lineEnd = position

        Do While lineEnd < FlowLayoutElementDescriptor_size(elementDescriptor)
            If FlowLayoutElementDescriptor_isLineBreak(elementDescriptor, lineEnd) Then Exit Do
            lineEnd = lineEnd + 1
        Loop

        FlowLayout_placeLine layout, elementDescriptor, position, lineEnd

        If lineEnd < FlowLayoutElementDescriptor_size(elementDescriptor) Then
            FlowLayout_breakLine layout
        End If

        position = lineEnd + 1
    Loop

    layout.currentElementId = layout.currentElementId + 1
End Sub

Private Sub FlowLayout_end(layout As FlowLayout)
    FlowLayout_endWord layout

    If layout.currentLineCount > 0 Then
        FlowLayout_breakLine layout
    End If
End Sub

Private Sub FlowLayout_breakLine(layout As FlowLayout)
    Dim i As Long
    FlowLayout_endWord layout
    ' lines.push_back(currentLine)
    ReDim Preserve layout.lineStart(0 To layout.linesCount)
    layout.lineStart(layout.linesCount) = layout.linesFragCount
    layout.linesCount = layout.linesCount + 1
    For i = 0 To layout.currentLineCount - 1
        FlowLayoutFragments_push_back layout.linesFrag, layout.linesFragCount, layout.currentLine(i)
    Next
    layout.currentLineCount = 0
    layout.nextOffset = 0
End Sub

' (VB) getLines(): line i's fragments are linesFrag(FlowLayout_lineBegin(i)) .. linesFrag(FlowLayout_lineEnd(i) - 1)
Private Function FlowLayout_lineBegin(layout As FlowLayout, ByVal i As Long) As Long
    FlowLayout_lineBegin = layout.lineStart(i)
End Function

Private Function FlowLayout_lineEnd(layout As FlowLayout, ByVal i As Long) As Long
    If i + 1 < layout.linesCount Then
        FlowLayout_lineEnd = layout.lineStart(i + 1)
    Else
        FlowLayout_lineEnd = layout.linesFragCount
    End If
End Function

'--------------------------------------------------------------------------
' Used to render a fragment of text in the paragraph.
'
' Can be an entire line, or part of a line, but never more than one line.
' class ParagraphTextWidget

' std::make_shared<ParagraphTextWidget>(text, textStyle)
Public Function ParagraphTextWidget_new(ByVal text As String, textStyle As ParagraphTextStyle) As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_ParagraphTextWidget)
    ParagraphTextWidget_construct w, text, textStyle
    Set ParagraphTextWidget_new = w
End Function

' ParagraphTextWidget(WzString text, ParagraphTextStyle const &textStyle)
Public Sub ParagraphTextWidget_construct(ByVal self As WIDGET, ByVal text As String, textStyle As ParagraphTextStyle)
    Dim d As ParagraphTextWidget
    WIDGET_construct self, WIDG_UNSPECIFIED_TYPE
    self.kind = K_ParagraphTextWidget ' (VB) virtual calls from here on reach ParagraphTextWidget's overrides
    Set d = ParagraphTextWidget_(self)
    Set d.cachedText = WzCachedText_new_text(text, textStyle.font)
    d.textStyle = textStyle
    self.setGeometry self.x(), self.y(), d.cachedText.deref().width(), d.cachedText.deref().lineSize()
End Sub

Public Sub ParagraphTextWidget_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim d As ParagraphTextWidget, x0 As Long, y0 As Long, style As ParagraphTextStyle
    Set d = ParagraphTextWidget_(self)
    style = d.textStyle
    x0 = xOffset + self.x()
    y0 = yOffset + self.y()
    pie_UniTransBoxFill x0, y0 + 1, x0 + self.width(), y0 + 1 + self.height(), style.shadeColour
    d.cachedText.deref().render x0, y0 - d.cachedText.deref().aboveBase(), style.fontColour
End Sub

Public Sub ParagraphTextWidget_run(ByVal self As WIDGET, psContext As W_CONTEXT)
    ParagraphTextWidget_(self).cachedText.tick
End Sub

Public Sub ParagraphTextWidget_setFontColour(ByVal self As WIDGET, ByVal colour As Long)
    Dim style As ParagraphTextStyle
    style = ParagraphTextWidget_(self).textStyle
    style.fontColour = colour
    ParagraphTextWidget_(self).textStyle = style
End Sub

'--------------------------------------------------------------------------
' struct ParagraphTextElement and struct ParagraphWidgetElement (ParagraphElement.cls)

' ParagraphTextElement(WzString const &newText, ParagraphTextStyle const &style, bool rtl = false)
Private Function ParagraphTextElement_new(ByVal newText As String, style As ParagraphTextStyle, Optional ByVal rtl As Boolean = False) As ParagraphElement
    Dim e As ParagraphElement
    Set e = New ParagraphElement
    e.kind = PARAGRAPH_TEXT_ELEMENT
    e.style = style
    e.text = newText
    e.rightToLeft = rtl
    Set e.fragments = New Collection
    Set ParagraphTextElement_new = e
End Function

' ParagraphWidgetElement(const std::shared_ptr<WIDGET> &widget, int32_t aboveBase)
Private Function ParagraphWidgetElement_new(ByVal widget As WIDGET, ByVal aboveBase As Long) As ParagraphElement
    Dim e As ParagraphElement
    Set e = New ParagraphElement
    e.kind = PARAGRAPH_WIDGET_ELEMENT
    Set e.widget = widget
    e.aboveBase = aboveBase
    e.layoutWidth = 0
    e.layoutHeight = 0
    Set ParagraphWidgetElement_new = e
End Function

Private Sub ParagraphElement_appendTo(ByVal e As ParagraphElement, layout As FlowLayout)
    Dim desc As FlowLayoutElementDescriptor
    Select Case e.kind
    Case PARAGRAPH_TEXT_ELEMENT
        FlowLayoutStringDescriptor_construct desc, e.text, e.style.font, e.rightToLeft
        FlowLayout_append layout, desc
    Case PARAGRAPH_WIDGET_ELEMENT
        desc.kind = DESCRIPTOR_WIDGET   ' layout.append(*this)
        Set desc.widget = e.widget
        FlowLayout_append layout, desc
    End Select
End Sub

Private Function ParagraphElement_createFragmentWidget(ByVal e As ParagraphElement, ByVal paragraph As WIDGET, fragment As FlowLayoutFragment) As WIDGET
    Dim widget As WIDGET
    Select Case e.kind
    Case PARAGRAPH_TEXT_ELEMENT
        Set widget = ParagraphTextWidget_new(Mid$(e.text, fragment.begin + 1, fragment.length), e.style)
        widget.setTransparentToMouse True
        paragraph.attach widget
        e.fragments.Add widget
        Set ParagraphElement_createFragmentWidget = widget
    Case PARAGRAPH_WIDGET_ELEMENT
        e.layoutWidth = e.widget.width()
        e.layoutHeight = e.widget.height()
        Set ParagraphElement_createFragmentWidget = e.widget
    End Select
End Function

Private Function ParagraphElement_isLayoutDirty(ByVal e As ParagraphElement) As Boolean
    Select Case e.kind
    Case PARAGRAPH_TEXT_ELEMENT
        ParagraphElement_isLayoutDirty = False
    Case PARAGRAPH_WIDGET_ELEMENT
        ParagraphElement_isLayoutDirty = e.widget.width() <> e.layoutWidth Or e.widget.height() <> e.layoutHeight
    End Select
End Function

Private Sub ParagraphElement_destroyFragments(ByVal e As ParagraphElement, ByVal paragraph As WIDGET)
    Dim i As Long
    Select Case e.kind
    Case PARAGRAPH_TEXT_ELEMENT
        For i = 1 To e.fragments.Count
            paragraph.detach e.fragments(i)
        Next

        Set e.fragments = New Collection
    Case PARAGRAPH_WIDGET_ELEMENT
    End Select
End Sub

Private Function ParagraphElement_getAboveBase(ByVal e As ParagraphElement) As Long
    Select Case e.kind
    Case PARAGRAPH_TEXT_ELEMENT
        ParagraphElement_getAboveBase = iV_GetTextAboveBase(e.style.font)
    Case PARAGRAPH_WIDGET_ELEMENT
        ParagraphElement_getAboveBase = e.aboveBase
    End Select
End Function

' ParagraphTextElement::setFontColour
Private Sub ParagraphTextElement_setFontColour(ByVal e As ParagraphElement, ByVal colour As Long)
    Dim style As ParagraphTextStyle, i As Long
    style = e.style
    style.fontColour = colour
    e.style = style
    For i = 1 To e.fragments.Count
        ParagraphTextWidget_setFontColour e.fragments(i), colour
    Next
End Sub

'--------------------------------------------------------------------------
' class Paragraph

' std::make_shared<Paragraph>()
Public Function Paragraph_new() As WIDGET
    Dim w As WIDGET
    Set w = WIDGET_alloc(K_Paragraph)
    Paragraph_construct w
    Set Paragraph_new = w
End Function

' Paragraph(): WIDGET() {}
Public Sub Paragraph_construct(ByVal self As WIDGET)
    Dim d As Paragraph, style As ParagraphTextStyle
    WIDGET_construct self
    self.kind = K_Paragraph ' (VB) virtual calls from here on reach Paragraph's overrides
    Set d = Paragraph_(self)
    d.layoutDirty = True
    d.layoutWidth = 0
    ParagraphTextStyle_construct style
    d.textStyle = style
    d.lineSpacing = 0
    d.isMouseDown = False
End Sub

Private Function Paragraph_hasElementWithLayoutDirty(ByVal self As WIDGET) As Boolean
    Dim d As Paragraph, i As Long
    Set d = Paragraph_(self)
    For i = 1 To d.elements.Count
        If ParagraphElement_isLayoutDirty(d.elements(i)) Then
            Paragraph_hasElementWithLayoutDirty = True
            Exit Function
        End If
    Next

    Paragraph_hasElementWithLayoutDirty = False
End Function

Private Sub Paragraph_updateLayout(ByVal self As WIDGET)
    Dim d As Paragraph, i As Long, nextLineOffset As Long, totalHeight As Long, linesLayout As FlowLayout
    Dim lineIdx As Long, k As Long, lineFragments As Collection, aboveBase As Long, belowBase As Long
    Dim fragmentDescriptor As FlowLayoutFragment, fragment As WIDGET, fragmentAboveBase As Long, element As ParagraphElement
    Set d = Paragraph_(self)
    If Not d.layoutDirty And Not Paragraph_hasElementWithLayoutDirty(self) Then
        Exit Sub
    End If

    d.layoutDirty = False
    d.layoutWidth = self.width()

    For i = 1 To d.elements.Count
        ParagraphElement_destroyFragments d.elements(i), self
    Next
    d.scrollSnapOffsets_clear

    nextLineOffset = 0
    totalHeight = 0
    Paragraph_calculateLinesLayout self, linesLayout
    For lineIdx = 0 To linesLayout.linesCount - 1
        Set lineFragments = New Collection  ' std::vector<WIDGET *> lineFragments
        aboveBase = 0
        belowBase = 0

        d.scrollSnapOffsets_push_back nextLineOffset

        For k = FlowLayout_lineBegin(linesLayout, lineIdx) To FlowLayout_lineEnd(linesLayout, lineIdx) - 1
            fragmentDescriptor = linesLayout.linesFrag(k)
            Set element = d.elements(fragmentDescriptor.elementId + 1)
            Set fragment = ParagraphElement_createFragmentWidget(element, self, fragmentDescriptor)
            fragmentAboveBase = -ParagraphElement_getAboveBase(element)
            aboveBase = MAX_(aboveBase, fragmentAboveBase)
            belowBase = MAX_(belowBase, fragment.height() - fragmentAboveBase)
            fragment.setGeometry fragmentDescriptor.offset, nextLineOffset - fragmentAboveBase, MIN_(self.width(), fragment.width()), fragment.height()
            lineFragments.Add fragment
        Next

        For k = 1 To lineFragments.Count
            Set fragment = lineFragments(k)
            fragment.setGeometry fragment.x(), fragment.y() + aboveBase, fragment.width(), fragment.height()
        Next

        totalHeight = nextLineOffset + aboveBase + belowBase
        nextLineOffset = totalHeight + d.lineSpacing
    Next

    self.setGeometry self.x(), self.y(), self.width(), totalHeight
End Sub

' std::vector<std::vector<FlowLayoutFragment>> calculateLinesLayout(): the FlowLayout's lines
Private Sub Paragraph_calculateLinesLayout(ByVal self As WIDGET, layout As FlowLayout)
    Dim d As Paragraph, i As Long
    Set d = Paragraph_(self)
    FlowLayout_construct layout, self.width()
    For i = 1 To d.elements.Count
        ParagraphElement_appendTo d.elements(i), layout
    Next
    FlowLayout_end layout
End Sub

Public Sub Paragraph_addText(ByVal self As WIDGET, ByVal text As String)
    Dim d As Paragraph, startOffsets() As Long, endOffsets() As Long, rightToLeft() As Boolean, n As Long, i As Long
    Dim runText As String, style As ParagraphTextStyle
    Set d = Paragraph_(self)
    d.layoutDirty = True
    style = d.textStyle
    n = iV_SplitTextParagraphIntoRuns(text, style.font, startOffsets, endOffsets, rightToLeft)
    If n > 1 Then
        For i = 0 To n - 1
            runText = Mid$(text, startOffsets(i) + 1, endOffsets(i) - startOffsets(i))
            d.elements.Add ParagraphTextElement_new(runText, style, rightToLeft(i))
        Next
    Else
        d.elements.Add ParagraphTextElement_new(text, style)
    End If
End Sub

Public Sub Paragraph_addWidget(ByVal self As WIDGET, ByVal widget As WIDGET, ByVal aboveBase As Long)
    Paragraph_(self).layoutDirty = True
    self.attach widget
    Paragraph_(self).elements.Add ParagraphWidgetElement_new(widget, aboveBase)
End Sub

Public Sub Paragraph_setFont(ByVal self As WIDGET, ByVal font As iV_fonts)
    Dim style As ParagraphTextStyle
    style = Paragraph_(self).textStyle
    style.font = font
    Paragraph_(self).textStyle = style
End Sub

Public Sub Paragraph_setFontColour(ByVal self As WIDGET, ByVal colour As Long)
    Dim style As ParagraphTextStyle
    style = Paragraph_(self).textStyle
    style.fontColour = colour
    Paragraph_(self).textStyle = style
End Sub

Public Sub Paragraph_setShadeColour(ByVal self As WIDGET, ByVal colour As Long)
    Dim style As ParagraphTextStyle
    style = Paragraph_(self).textStyle
    style.shadeColour = colour
    Paragraph_(self).textStyle = style
End Sub

Public Sub Paragraph_setLineSpacing(ByVal self As WIDGET, ByVal newLineSpacing As Long)
    Paragraph_(self).lineSpacing = newLineSpacing
    Paragraph_(self).layoutDirty = True
End Sub

Public Function Paragraph_getTextStyle(ByVal self As WIDGET) As ParagraphTextStyle
    Paragraph_getTextStyle = Paragraph_(self).textStyle
End Function

Public Sub Paragraph_geometryChanged(ByVal self As WIDGET)
    If Paragraph_(self).layoutWidth <> self.width() Then
        Paragraph_(self).layoutDirty = True
    End If

    Paragraph_updateLayout self
End Sub

Public Sub Paragraph_displayRecursive(ByVal self As WIDGET, context As WidgetGraphicsContext)
    Dim overrideContext As WidgetGraphicsContext
    Paragraph_updateLayout self

    overrideContext = WidgetGraphicsContext_setAllowChildDisplayRecursiveIfSelfClipped(context, True)

    WIDGET_displayRecursive self, overrideContext
End Sub

Public Sub Paragraph_display(ByVal self As WIDGET, ByVal xOffset As Long, ByVal yOffset As Long)
    Dim x0 As Long, y0 As Long
    If Not Paragraph_(self).isMouseDown Or Paragraph_(self).onClickHandlers.Count = 0 Then Exit Sub

    x0 = self.x() + xOffset
    y0 = self.y() + yOffset
    iV_Box x0, y0, x0 + self.width(), y0 + self.height(), psPalette(WZCOL_RED)
End Sub

Public Sub Paragraph_addOnClickHandler(ByVal self As WIDGET, ByVal onClickFunc As WzFunc)
    If onClickFunc Is Nothing Then
        Exit Sub
    End If
    Paragraph_(self).onClickHandlers.Add onClickFunc
End Sub

Public Sub Paragraph_clicked(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    self.dirty = True
    Paragraph_(self).isMouseDown = True
End Sub

' Respond to a mouse button up
Public Sub Paragraph_released(ByVal self As WIDGET, psContext As W_CONTEXT, Optional ByVal key As WIDGET_KEY = WKEY_PRIMARY)
    Dim d As Paragraph, i As Long, onClickHandler As Variant
    Set d = Paragraph_(self)
    If Not d.isMouseDown Then
        Exit Sub
    End If

    d.isMouseDown = False

    ' Call all onClick event handlers
    For i = 1 To d.onClickHandlers.Count
        Set onClickHandler = d.onClickHandlers(i)
        If Not onClickHandler Is Nothing Then
            WzFunc_call onClickHandler, self, key
        End If
    Next
    self.dirty = True
End Sub

' Respond to the mouse moving off the widget
Public Sub Paragraph_highlightLost(ByVal self As WIDGET)
    Paragraph_(self).isMouseDown = False
    self.dirty = True
End Sub

Public Function Paragraph_getScrollSnapOffsets(ByVal self As WIDGET, offsets() As Long) As Boolean
    Paragraph_(self).scrollSnapOffsets_get offsets
    Paragraph_getScrollSnapOffsets = True
End Function

Public Sub Paragraph_forceSetAllFontColor(ByVal self As WIDGET, ByVal colour As Long)
    Dim d As Paragraph, i As Long, psTextElement As ParagraphElement
    Set d = Paragraph_(self)
    ' set for any newly-added text
    Paragraph_setFontColour self, colour

    ' set for all existing text
    For i = 1 To d.elements.Count
        Set psTextElement = d.elements(i)
        If psTextElement.kind = PARAGRAPH_TEXT_ELEMENT Then
            ParagraphTextElement_setFontColour psTextElement, colour
        End If
    Next
End Sub
