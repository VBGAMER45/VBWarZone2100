VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LobbyBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' (VB) data of widget class LobbyBrowser
' src/titleui/widgets/gamebrowserform.cpp: the fields of class LobbyBrowser (gamebrowserform_c); the vectors are
' Private with Friend accessors: minimumColumnWidths(0 .. minimumColumnWidthsCount - 1), currentResult(0 ..
' currentResultsCount - 1).
Option Explicit

Public backButton As WIDGET             ' std::shared_ptr<W_BUTTON>
Public connectToForm As WIDGET          ' std::shared_ptr<ConnectToForm>
Public filtersButton As WIDGET          ' std::shared_ptr<WzFrontendImageButton>
Public refreshButton As WIDGET          ' std::shared_ptr<WzFrontendImageButton>
Public table As WIDGET                  ' std::shared_ptr<ScrollableTableWidget>
Private minimumColumnWidths_() As Long  ' std::vector<size_t> minimumColumnWidths;
Public minimumColumnWidthsCount As Long
Public minDesiredTableWidth As Long     ' = 0
Public lobbyStatusOverlayWidg As WIDGET ' std::shared_ptr<LobbyStatusOverlayWidget>
Public lobbyStatusMessageContainer As WIDGET ' std::shared_ptr<ScrollableListWidget>

Private currentResults_() As netlobby_GameListing ' std::vector<netlobby::GameListing> currentResults;
Public currentResultsCount As Long
Public filterIncompatible As Boolean    ' = true
Public filterEmpty As Boolean           ' = false
Public filterModded As Boolean          ' = false
Public filterIPv6Only As Boolean        ' = false

Public currentPopoverMenu As WIDGET     ' std::shared_ptr<PopoverMenuWidget>

Public hasPendingFetchRequest As Boolean ' = false

Public weakSerial As Long               ' (VB) its std::weak_ptr<LobbyBrowser>s: the widgweak_vb serial

Friend Sub getMinimumColumnWidths(out() As Long)
    out = minimumColumnWidths_
End Sub

' minimumColumnWidths = { v(0 .. n - 1) }
Friend Sub setMinimumColumnWidths(v() As Long, ByVal n As Long)
    Dim i As Long
    minimumColumnWidthsCount = n
    If n > 0 Then
        ReDim minimumColumnWidths_(0 To n - 1)
        For i = 0 To n - 1
            minimumColumnWidths_(i) = v(i)
        Next
    Else
        Erase minimumColumnWidths_
    End If
End Sub

Friend Property Get currentResult(ByVal i As Long) As netlobby_GameListing
    currentResult = currentResults_(i)
End Property

' currentResults = std::move(games(0 .. n - 1))
Friend Sub setCurrentResults(games() As netlobby_GameListing, ByVal n As Long)
    Dim i As Long
    currentResultsCount = n
    If n > 0 Then
        ReDim currentResults_(0 To n - 1)
        For i = 0 To n - 1
            currentResults_(i) = games(i)
        Next
    Else
        Erase currentResults_
    End If
End Sub

' currentResults.clear()
Friend Sub clearCurrentResults()
    currentResultsCount = 0
    Erase currentResults_
End Sub
