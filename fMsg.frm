VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   6870
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   16050
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------------------------------------------
' UserForm fMsg Provides all means for a message with
'               - up to 3 separated text messages, each either with a
'                 proportional or a fixed font
'               - each of the 3 messages with an optional label
'               - 4 reply buttons either specified with replies known
'                 from the VB MsgSectionBox or any test string.
'
' Design: The implementation is merely design driven. I.e. the names of
'         the elements are not used but the logic of the elements hierarchy.
'         1         Frame MsgArea
'         1.1       Frame Image
'         1.2       Frame MsgSection
'         1.2.1     Frame MsgSection1 to n (currently designed is n=3)
'         1.2.1.1   Label MsgSectionLabel1 to ....3
'         1.2.1.2   Frame MsgSectionFrame1 to ...3
'         1.2.1.2.1 TextBox MsgSectionText1 to ...3
'         2         Frame RepliesArea
'         2.1       Frame RepliesRow1 to ...n
'         2.1.1     CommandButton Reply1 to n
'
'
' lScreenWidth. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const FRAME_CAPTIONS            As Boolean = False
Const MONOSPACED_FONT_NAME      As String = "Courier New"   ' Default monospaced font
Const MONOSPACED_FONT_SIZE      As Single = 9               ' Default monospaced font size
Const FORM_WIDTH_MIN            As Single = 300
Const FORM_WIDTH_MAX_POW        As Long = 80    ' Maximum form width as a percentage of the screen size
Const FORM_HEIGHT_MAX_POW       As Long = 90    ' Maximum form height as a percentage of the screen size
Const F_MARGIN                  As Single = 2
Const L_MARGIN                  As Single = 0   ' Left margin for labels and text boxes
Const R_MARGIN                  As Single = 15  ' Right margin for labels and text boxes
Const H_MARGIN                  As Single = 10  ' Horizontal margin for reply buttons
Const V_MARGIN                  As Single = 10  ' Vertical marging for all displayed elements/controls
Const T_MARGIN                  As Single = 5   ' Top position for the first displayed control
Const B_MARGIN                  As Single = 50  ' Bottom margin after the last displayed control
Const V_MARGIN_LABEL            As Single = 5
Const MIN_WIDTH_REPLY_BUTTON    As Single = 70

' Functions to get the displays DPI
' Used for getting the metrics of the system devices.
'
Const SM_XVIRTUALSCREEN As Long = &H4C&
Const SM_YVIRTUALSCREEN As Long = &H4D&
Const SM_CXVIRTUALSCREEN As Long = &H4E&
Const SM_CYVIRTUALSCREEN As Long = &H4F&
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const TWIPSPERINCH = 1440
Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Dim bFramesWithCaption          As Boolean
Dim sTitle                      As String
Dim sErrSrc                     As String
Dim vReplies                    As Variant
Dim aReplyButtons               As Variant
Dim sTitleFontName              As String
Dim sTitleFontSize              As String   ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim siTitleWidth                As Single
Dim wVirtualScreenLeft          As Single
Dim wVirtualScreenTop           As Single
Dim wVirtualScreenWidth         As Single
Dim wVirtualScreenHeight        As Single
Dim lMaximumFormHeightPoW       As Long       ' % of the screen height
Dim lMaximumFormWidthPoW        As Long       ' % of the screen width
Dim lMinimumFormHeightPoW       As Long       ' % of the screen height
Dim lMinimumFormWidthPoW        As Long       ' % of the screen width
Dim siMaximumFormHeight         As Single     ' above converted to excel userform height
Dim siMaximumFormWidth          As Single     ' above converted to excel userform width
Dim siMinimumFormHeight         As Single
Dim siMinimumFormWidth          As Single
Dim sMonospacedFontName         As String
Dim siMonospacedFontSize        As Single
Dim cllAreas                    As New Collection   ' Collection of the two primary/top frames
Dim cllMsgSections              As New Collection   '
Dim cllMsgSectionsLabel         As New Collection
Dim cllMsgSectionsText          As New Collection   ' Collection of section frames
Dim cllMsgSectionsTextFrame     As New Collection
Dim cllSectionsVisible          As New Collection   ' Collection of visible section frames
Dim cllRepliesRow               As New Collection   ' Collection of the designed reply button row frames
Dim cllReplyRowButtons          As Collection       ' Collection of the designed reply buttons of a certain row
Dim cllReplyRowsButtons         As New Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim cllReplyRowButtonValues     As Collection       ' Collection of the return values of setup reply buttons of a certain row
Dim cllReplyRowsButtonValues    As New Collection   ' Collection of cllReplyRowButtonValues
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim dctSectionsLabel            As New Dictionary   ' User provided through Property SectionsLabel
Dim dctSectionsText             As New Dictionary   ' User provided through Property SectionsText
Dim dctSectionsMonospaced       As New Dictionary   ' User provided through Property Sections
Dim siRepliesWidth              As Single
Dim siRepliesHeight             As Single

Private Sub UserForm_Initialize()
    
    Dim ctl     As MSForms.Control
    Dim fr      As MSForms.Frame
    Dim v       As Variant
    Dim row     As Long
    Dim button  As Long
    Dim cmb     As MSForms.CommandButton
    
    On Error GoTo on_error
    
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_WIDTH_MAX_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_HEIGHT_MAX_POW
    siMinimumFormWidth = FORM_WIDTH_MIN                         ' Default UserForm width
    sMonospacedFontName = MONOSPACED_FONT_NAME                  ' Default monospaced font
    siMonospacedFontSize = MONOSPACED_FONT_SIZE                 ' Default monospaced font
    Me.FramesWithBorder = False
    Me.width = siMinimumFormWidth
    bFramesWithCaption = False
    
    Collect into:=cllAreas, ctltype:="Frame", fromparent:=Me, ctlheight:=10, ctlwidth:=Me.width - F_MARGIN
    RepliesArea.width = 10  ' Will be adjusted to the max replies row width during setup
    
    Collect into:=cllMsgSections, ctltype:="Frame", fromparent:=MsgArea, ctlheight:=50, ctlwidth:=MsgArea.width - F_MARGIN
    Collect into:=cllMsgSectionsLabel, ctltype:="Label", fromparent:=cllMsgSections, ctlheight:=15, ctlwidth:=MsgArea.width - (F_MARGIN * 2)
    Collect into:=cllMsgSectionsTextFrame, ctltype:="Frame", fromparent:=cllMsgSections, ctlheight:=20, ctlwidth:=MsgArea.width - (F_MARGIN * 2)
    Collect into:=cllMsgSectionsText, ctltype:="TextBox", fromparent:=cllMsgSectionsTextFrame, ctlheight:=20, ctlwidth:=MsgArea.width - (F_MARGIN * 3)
    
    Collect into:=cllRepliesRow, ctltype:="Frame", fromparent:=RepliesArea, ctlheight:=10, ctlwidth:=10
        
    For Each v In cllRepliesRow
        Collect into:=cllReplyRowButtons, ctltype:="CommandButton", fromparent:=v, ctlheight:=10, ctlwidth:=MIN_WIDTH_REPLY_BUTTON
        If cllReplyRowButtons.Count > 0 _
        Then cllReplyRowsButtons.Add cllReplyRowButtons
    Next v
    
    Me.Height = V_MARGIN * 4
    bWithFrames = False
    ResizeAndRepositionFrames
exit_sub:
    Exit Sub
    
on_error:
    Stop: Resume Next
End Sub

Private Property Get Monospaced(Optional ByVal section As Long) As Boolean
    Monospaced = MsgSection(section).Font.Name = sMonospacedFontName
End Property
Private Property Let Monospaced(Optional ByVal section As Long, ByVal monospace As Boolean)
    MsgSection(section).Font.Name = sMonospacedFontName
End Property

' Property for testing purpose only, defaulting to False
' When True frames are displayed with a visible border
Public Property Let FramesWithBorder(ByVal b As Boolean)
    
    Dim ctl As MSForms.Control
       
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Or TypeName(ctl) = "TextBox" Then
            ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
            If b = False _
            Then ctl.BorderStyle = fmBorderStyleNone _
            Else ctl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctl
    
End Property

Public Property Let FramesWithCaption(ByVal b As Boolean):                      bFramesWithCaption = b:                                 End Property

Public Property Let MaxFormWidthPrcntgOfScreenSize(ByVal l As Long)
    lMaximumFormWidthPoW = l
    siMaximumFormWidth = wVirtualScreenWidth * (Min(l, 99) / 100)   ' maximum form width based on screen size
End Property
Public Property Get MaxFormWidthPrcntgOfScreenSize() As Long:                   MaxFormWidthPrcntgOfScreenSize = lMaximumFormWidthPoW: End Property
Public Property Get MinFormWidthPrcntgOfScreenSize() As Long:                   MinFormWidthPrcntgOfScreenSize = lMinimumFormWidthPoW: End Property

Public Property Let MaxFormHeightPrcntgOfScreenSize(ByVal l As Long)
    lMaximumFormHeightPoW = l
    siMaximumFormHeight = wVirtualScreenHeight * (Min(l, 99) / 100)   ' maximum form height based on screen size
End Property
Public Property Get MaxFormHeightPrcntgOfScreenSize() As Long:                  MaxFormHeightPrcntgOfScreenSize = lMaximumFormHeightPoW: End Property

Public Property Get MaximumFormWidth() As Single:                               MaximumFormWidth = siMaximumFormWidth:      End Property
Public Property Get MaximumFormHeight() As Single:                              MaximumFormHeight = siMaximumFormHeight:    End Property

Public Property Get MinimumFormWidth() As Single:                               MinimumFormWidth = siMinimumFormWidth:          End Property
Public Property Let MinimumFormWidth(ByVal si As Single)
    siMinimumFormWidth = si
    '~~ The maximum form width must never not become less than the minimum width
    If siMaximumFormWidth < siMinimumFormWidth Then
       siMaximumFormWidth = siMinimumFormWidth
    End If
    lMinimumFormWidthPoW = CInt((siMinimumFormWidth / wVirtualScreenWidth) * 100)
End Property

Private Property Get ReplyButton(Optional ByVal row As Long, Optional ByVal button As Long) As MSForms.CommandButton
    Set ReplyButton = cllReplyRowsButtons(row)(button)
End Property

Private Property Get ReplyButtonValue(Optional ByVal row As Long, Optional ByVal button As Long)
    ReplyButtonValue = cllReplyRowsButtonValues(row)(button)
End Property

Private Property Let FormWidth(ByVal w As Single)
    With Me
        .width = Max(.width, siMinimumFormWidth, w)
    End With
End Property

Private Property Let ReplyRowButtonValues(ByVal v As Variant):          cllReplyRowButtonValues.Add v:      End Property
Private Property Let ReplyRowsButtonValues(ByVal cll As Collection):    cllReplyRowsButtonValues.Add cll:   End Property
Private Property Let ReplyRowButtons(ByVal v As MSForms.CommandButton): cllReplyRowButtons.Add v:           End Property
Private Property Let ReplyRowsButtons(ByVal cll As Collection):         cllReplyRowsButtons.Add cll:        End Property

' UserForm design elements
Private Property Get Areas() As Collection:                                                 Set Areas = cllAreas:                                           End Property
Private Property Get MsgArea() As MSForms.Frame:                                            Set MsgArea = cllAreas(1):                                      End Property
Private Property Get RepliesArea() As MSForms.Frame:                                        Set RepliesArea = cllAreas(2):                                  End Property
Private Property Get ReplyRows() As Collection:                                             Set ReplyRows = cllRepliesRow:                                  End Property
Private Property Get RepliesRow(Optional ByVal row As Long) As MSForms.Frame:               Set RepliesRow = cllRepliesRow(row):                            End Property
Private Property Get MsgFrames() As Collection:                                             Set MsgFrames = cllMsgSectionsTextFrame:                        End Property
Private Property Get MsgFrame(Optional ByVal section As Long) As MSForms.Frame:             Set MsgFrame = cllMsgSectionsTextFrame(section):                End Property
Private Property Get MsgSections() As Collection:                                           Set MsgSections = cllMsgSections:                               End Property
Private Property Get MsgSection(Optional section As Long) As MSForms.Frame:                 Set MsgSection = cllMsgSections(section):                       End Property
Private Property Get MsgSectionLabel(Optional section As Long) As MSForms.Label:           Set MsgSectionLabel = cllMsgSectionsLabel(section):            End Property
Private Property Get MsgSectionTextFrame(Optional ByVal section As Long):                  Set MsgSectionTextFrame = cllMsgSectionsTextFrame(section):    End Property
Private Property Get MsgSectionText(Optional section As Long) As MSForms.TextBox:          Set MsgSectionText = cllMsgSectionsText(section):              End Property
Private Property Get RepliesSetupInRow(Optional ByVal row As Long) As Long:                 RepliesSetupInRow = cllReplyRowsButtonValues(row).Count:        End Property
Private Property Get ReplyRowsSetup() As Long:                                              ReplyRowsSetup = cllReplyRowsButtonValues.Count:                End Property

' Message section properties (label, text, monospaced)
Public Property Let SectionsLabel(Optional ByVal section As Long, ByVal s As String):       dctSectionsLabel.Add section, s:                                End Property
Public Property Get SectionsLabel(Optional ByVal section As Long) As String
    If dctSectionsLabel.Exists(section) _
    Then SectionsLabel = dctSectionsLabel(section) _
    Else SectionsLabel = vbNullString
End Property
Public Property Let SectionsText(Optional ByVal section As Long, ByVal s As String):        dctSectionsText.Add section, s:                             End Property
Public Property Get SectionsText(Optional ByVal section As Long) As String
    If dctSectionsText.Exists(section) _
    Then SectionsText = dctSectionsText(section) _
    Else SectionsText = vbNullString
End Property
Public Property Let SectionsMonospaced(Optional ByVal section As Long, ByVal b As Boolean): dctSectionsMonospaced.Add section, b:                       End Property
Public Property Get SectionsMonospaced(Optional ByVal section As Long) As Boolean
    If dctSectionsMonospaced.Exists(section) _
    Then SectionsMonospaced = dctSectionsMonospaced(section) _
    Else SectionsMonospaced = False
End Property
Public Property Let ErrSrc(ByVal s As String):                                              sErrSrc = s:                                                End Property
Public Property Let Replies(ByVal v As Variant):                                            vReplies = v:                                               End Property
Public Property Let Title(ByVal s As String):                                               sTitle = s:                                                 End Property

' Set the top position for the control (ctl) and return the top posisition for the next one
Private Property Get TopNext(ByVal ctl As MSForms.Control) As Single

    TopNext = siTopNext
    With ctl
        .Top = siTopNext    ' the top position for this one
        '~~ Calculate the top position for any displayed which may come next
        Select Case TypeName(ctl)
            Case "TextBox", "CommandButton":    siTopNext = .Top + .Height + V_MARGIN
            Case "Label"
                Select Case ctl.Name
                    Case "la":                  siTopNext = Me.laMsgTitleSpaceBottom.Top + Me.laMsgTitleSpaceBottom.Height + V_MARGIN
                    Case Else:                  siTopNext = .Top + .Height
                End Select
        End Select
    End With

End Property

Private Sub cmbReply11_Click():  ReplyClicked 1, 1:   End Sub
Private Sub cmbReply12_Click():  ReplyClicked 1, 2:   End Sub
Private Sub cmbReply13_Click():  ReplyClicked 1, 3:   End Sub
Private Sub cmbReply14_Click():  ReplyClicked 1, 4:   End Sub
Private Sub cmbReply15_Click():  ReplyClicked 1, 5:   End Sub

Private Sub cmbReply21_Click():  ReplyClicked 2, 1:   End Sub
Private Sub cmbReply22_Click():  ReplyClicked 2, 2:   End Sub
Private Sub cmbReply23_Click():  ReplyClicked 2, 3:   End Sub
Private Sub cmbReply24_Click():  ReplyClicked 2, 4:   End Sub
Private Sub cmbReply25_Click():  ReplyClicked 2, 5:   End Sub

' Final form height adjustment considering only the maximum height specified
' --------------------------------------------------------------------------
Private Sub FormHeightFinal()
    
    Dim siHeightExceeding   As Single
    Dim s                   As String
    Dim siWidth             As Single
    
    With Me
        '~~ Reduce the height of the largest displayed message paragraph by the amount of exceeding height
        siHeightExceeding = .Height - siMaximumFormHeight
        .Height = siMaximumFormHeight
        With MsgSectionMaxHeight
            siWidth = .width
            s = .value
            .SetFocus
            .AutoSize = False
            .value = vbNullString
            Select Case .ScrollBars
                Case fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsVertical
                    .width = siWidth + 15
                    .Height = .Height - siHeightExceeding - 15
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsVertical
                Case fmScrollBarsBoth
                    .Height = .Height - siHeightExceeding - 15
                    .width = siWidth - 15
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsVertical
                    .width = siWidth + 15
                    .Height = .Height - siHeightExceeding
            End Select
            .value = s
            .SelStart = 0
        End With
    End With
    
End Sub

' Returns the visible textbox with the largest height.
' ----------------------------------------------------------
Private Function MsgSectionMaxHeight() As MSForms.TextBox
Dim v   As Variant
Dim si  As Single
Dim tb  As MSForms.TextBox

    For Each v In MsgSectionsTextVisible
        Set tb = v
        If tb.Height > si Then
            si = tb.Height
            Set MsgSectionMaxHeight = tb
        End If
    Next v
    
End Function

' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' -------------------------------------------------------------
Private Sub MsgSectionSetup( _
            ByVal section As Long)
    
    Dim frArea              As MSForms.Frame
    Dim frSection           As MSForms.Frame
    Dim laSectionLabel      As MSForms.Label
    Dim tbSectionText       As MSForms.TextBox
    Dim frSectionTextFrame  As MSForms.Frame
    Dim sMsgLabel           As String
    Dim sMsgText            As String
    Dim bMsgMonospaced      As Boolean

    Set frArea = MsgArea
    Set frSection = MsgSection(section)
    Set laSectionLabel = MsgSectionLabel(section)
    Set tbSectionText = MsgSectionText(section)
    Set frSectionTextFrame = MsgFrame(section)
    
    sMsgLabel = SectionsLabel(section)
    sMsgText = SectionsText(section)
    bMsgMonospaced = SectionsMonospaced(section)
    
    frSection.width = frArea.width
    laSectionLabel.width = frSection.width
    frSectionTextFrame.width = frSection.width
    tbSectionText.width = frSection.width
        
    If sMsgText <> vbNullString Then
        With frArea
            .Visible = True
            .left = 0
            .width = Me.width - R_MARGIN
            .Height = Max(.Height + (V_MARGIN * 4), V_MARGIN * 4) ' V_MARGIN is the initial height
        End With
        With frSection
            .width = frArea.width
            .Height = V_MARGIN * 4 ' initial height
            .Visible = True
        End With
        frSectionTextFrame.Visible = True
        tbSectionText.Visible = True
        
        ResizeAndRepositionFrames
        
        '~~ Setup above text label/title only when there is a text
        If sMsgLabel <> vbNullString Then
            Set laSectionLabel = MsgSectionLabel(section)
            With laSectionLabel
                .width = frSection.width
                .Caption = sMsgLabel
                .Visible = True
            End With
            frSectionTextFrame.Top = laSectionLabel.Top + laSectionLabel.Height
        Else
            frSectionTextFrame.Top = 0
        End If
        
        If bMsgMonospaced Then
            MsgSectionSetupMonospaced section, sMsgText  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            MsgSectionSetupProportional section, sMsgText
        End If
        DoEvents
        tbSectionText.SelStart = 0
        
    End If
    frSectionTextFrame.Height = tbSectionText.Height
    frSection.Height = frSectionTextFrame.Top + frSectionTextFrame.Height
    frArea.Height = frSection.Top + frSection.Height + V_MARGIN
    Me.Height = Max(Me.Height, siTopNext + (V_MARGIN * 4))
    
    ResizeAndRepositionFrames

End Sub

Private Sub MsgSectionSetupProportional(ByVal section As Long, _
                                        ByVal text As String)
    
    Dim frMsgSection    As MSForms.Frame
    Dim tbMsgText       As MSForms.TextBox
    Dim frMsgText       As MSForms.Frame
    
    Set frMsgSection = MsgSection(section)
    Set frMsgText = MsgSectionTextFrame(section)
    Set tbMsgText = MsgSectionText(section)
    
    '~~ Setup the textbox
    With frMsgText
        .Visible = True
        .Height = V_MARGIN * 4 ' initial height
    End With
    frMsgSection.Height = frMsgText.Top + frMsgText.Height + V_MARGIN
    ResizeAndRepositionFrames
    
    With tbMsgText
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = MsgArea.width
        .value = text
        .SelStart = 0
    End With
    
    ' Adjust surrounding frames accordingly
    With MsgSection(section)
        .width = tbMsgText.width
        .Height = tbMsgText.Height
    End With
    With MsgSection(section)
        .width = tbMsgText.width
        .Height = tbMsgText.Height
    End With
                                       
    ResizeAndRepositionFrames
    
End Sub
                                       
Private Sub AddHorizontalScrollBarToFrame(ByVal section As Long)
    
    Dim frArea          As MSForms.Frame
    Dim frSection       As MSForms.Frame
    Dim frSectionText   As MSForms.Frame
    Dim tbSectionText   As MSForms.TextBox
    
    Set frArea = MsgArea
    Set frSection = MsgSection(section)
    Set frSectionText = MsgSectionTextFrame(section)
    Set tbSectionText = MsgSectionText(section)

    frArea.width = siMaximumFormWidth - L_MARGIN - R_MARGIN
    frSection.width = frArea.width - F_MARGIN
    frSectionText.width = frSection.width - F_MARGIN
    
    frSectionText.Height = tbSectionText.Height + 15 ' space for the scroll bar
    frSection.Height = frSection.Height + 15 ' space for the scroll bar
    frArea.Height = frArea.Height + 15
    
    With frSectionText
        Select Case .ScrollBars
            Case fmScrollBarsBoth
            Case fmScrollBarsHorizontal
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsHorizontal
            Case fmScrollBarsVertical
                .ScrollBars = fmScrollBarsHorizontal
        End Select
        .ScrollWidth = tbSectionText.width
        .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
    End With

End Sub

' Reduce the height of the message section (section) by the amount
' the current form height exceeds the specified maximum form height
' and apply a vertical scrollbar.
' Note: - When the vertical scrollbar is about to be added also the
'         form width must not be changed
'       - A vertical scrollbar may be added to any message section
' -----------------------------------------------------------------
Private Sub MsgSectionScrollBarAddVertical(ByVal section As Long)
    
    Dim frFormSection   As MSForms.Frame
    Dim frMsgSection    As MSForms.Frame
    Dim tbMsgText       As MSForms.TextBox
    
    Set frFormSection = MsgArea
    Set frMsgSection = MsgSection(section)
    Set tbMsgText = MsgSectionText(section)

    frFormSection.Height = frFormSection.Height - (Me.Height - siMaximumFormHeight) ' reduce height by the exceeding amount
    frMsgSection.Height = frFormSection.Height - 2  ' reduce text frame accordinglyy
    tbMsgText.width = tbMsgText.width - 25          ' make room for the vertical scroll bar
    frFormSection.Height = frMsgSection.Height
    With frMsgSection
        Select Case .ScrollBars
            Case fmScrollBarsBoth
            Case fmScrollBarsHorizontal
                .ScrollBars = fmScrollBarsVertical
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsVertical
            Case fmScrollBarsVertical
        End Select
        .ScrollWidth = tbMsgText.Height
        .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
    End With

End Sub

Private Sub AdjustFormWidth(ByVal ctl As MSForms.Control)
    Me.width = mMsg.Max( _
               Me.width, _
               siMinimumFormWidth, _
               ctl.left + ctl.width + R_MARGIN)
End Sub

' Setup the width of the monospaced message section (section) with
' text (text) by means of the monospace width template and
' apply width and height. The section frames are adjusted accordingly.
' --------------------------------------------------------------------
Private Sub MsgSectionSetupMonospaced( _
            ByVal section As Long, _
            ByVal text As String)
            
    Dim tbMsgText   As MSForms.TextBox
    Set tbMsgText = MsgSectionText(section)
    
    '~~ Setup the textbox
    With tbMsgText
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = False
        .Font.Name = sMonospacedFontName
        .Font.Size = siMonospacedFontSize
        .value = text
        .SelStart = 0
        
        If (.width + L_MARGIN + R_MARGIN) > siMaximumFormWidth _
        Then AddHorizontalScrollBarToFrame section ' only applied for monospaced section text
        
    End With
            
End Sub

Private Function MaxFormWidthExceeded(ByVal section As Long) As Boolean
    MaxFormWidthExceeded = MsgSection(section).width + L_MARGIN + R_MARGIN > siMaximumFormWidth
End Function

' Returns a collection of all visible message section frames
' ----------------------------------------------------------
Private Function MsgSectionsVisible() As Collection

    Dim v   As Variant
    Dim cll As New Collection
    
    For Each v In MsgSections
        If v.Visible Then cll.Add v
    Next v
    Set MsgSectionsVisible = cll
    
End Function

' Returns a collection of all visible message section text frames
' ---------------------------------------------------------------
Private Function MsgSectionsTextVisible() As Collection
    
    Dim v   As Variant
    Dim cll As New Collection
    
    For Each v In cllMsgSectionsText
        If v.Visible Then cll.Add v
    Next v
    Set MsgSectionsTextVisible = cll

End Function

' Executed only in case the form width had to be reduced in order to meet its specified maximum height.
' The message section with the largest height will be reduced to fit an will receive a vertical scroll bar.
' ---------------------------------------------------------------------------------------------------------
Private Sub MsgSectionsFinalHeight()
    
    Dim siHeightCurrentRequired As Single
    Dim siHeightExceeding       As Single
    Dim s                       As String

    With Me
        If .frRepliesRow2.Visible Then
            siHeightCurrentRequired = .frRepliesRow2.Top + .frRepliesRow2.Height + B_MARGIN
        Else
            siHeightCurrentRequired = .frRepliesRow1.Top + .frRepliesRow1.Height + B_MARGIN
        End If
    End With
    If siHeightCurrentRequired <= Me.Height Then Exit Sub
    
    siHeightExceeding = siHeightCurrentRequired > Me.Height
    '~~ All displayed controls together take more height than the available form's height
    '~~ The displayed message sections are reduced in their height to fit the available space
    With MsgSectionMaxHeight ' The message paragraph with the maximum height
        .SetFocus
        s = .value
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
            Case fmScrollBarsVertical
                .width = .width + 15
                .ScrollBars = fmScrollBarsBoth
                If Me.width < L_MARGIN + .width + R_MARGIN + 15 Then
                    Me.width = Me.width + 15
                End If
            Case fmScrollBarsNone
                .width = .width + 15
                .ScrollBars = fmScrollBarsVertical
                If Me.width < L_MARGIN + .width + R_MARGIN + 15 Then
                    Me.width = Me.width + 15
                End If
            Case fmScrollBarsBoth       ' nothing required
        End Select
    End With
End Sub

Private Sub MsgSectionsMonospacedSetup()
                             
    If SectionsText(1) <> vbNullString And SectionsMonospaced(1) = True Then MsgSectionSetup section:=1
    If SectionsText(2) <> vbNullString And SectionsMonospaced(2) = True Then MsgSectionSetup section:=2
    If SectionsText(3) <> vbNullString And SectionsMonospaced(3) = True Then MsgSectionSetup section:=3
    
End Sub

Private Sub MsgSectionsProportionalSetup()
                
    If SectionsText(1) <> vbNullString And SectionsMonospaced(1) = False Then MsgSectionSetup section:=1
    If SectionsText(2) <> vbNullString And SectionsMonospaced(2) = False Then MsgSectionSetup section:=2
    If SectionsText(3) <> vbNullString And SectionsMonospaced(3) = False Then MsgSectionSetup section:=3
    
End Sub

' After final adjustment of the form's width the visible the message paragraph's width is re-adjusted.
' Proportional message sections will result in a new height,
' mmonospaced message sections will receive a horizontal sccroll bar.
' ----------------------------------------------------------------------------------------------------
Private Sub MsgSectionsFinalWidth()

    Dim siMax           As Single ' The de-facto available width for any message section
    Dim v               As Variant
    Dim s               As String
    Dim lSection        As Long
    Dim frArea          As MSForms.Frame
    Dim frMsgSection    As MSForms.Frame
    Dim frMsgText       As MSForms.Frame
    Dim tbMsgText       As MSForms.TextBox
    
    siMax = Me.width - L_MARGIN - R_MARGIN
    
    Set frArea = MsgArea
    For lSection = 1 To MsgSections.Count
        Set frMsgSection = MsgSection(lSection)
        If frMsgSection.Visible Then
            Set frMsgText = MsgFrame(lSection)
            Set tbMsgText = MsgSectionText(lSection)
            If Not Monospaced(lSection) Then
                '~~ Adjust the proportional textbox width
                With tbMsgText
                    s = .value
                    .WordWrap = True
                    .AutoSize = True
                    .value = vbNullString
                    .width = Me.width - R_MARGIN
                    frArea.width = tbMsgText.width
                    frMsgSection.width = tbMsgText.width
                    frMsgText.width = tbMsgText.width
                    DoEvents
                    .value = s
                    frMsgText.Height = .Height
                    frMsgSection.Height = frMsgText.Top + frMsgText.Height
                End With
            End If
            ' Monospaced section already done with initial setup
            frArea.Height = frMsgSection.Top + frMsgSection.Height + V_MARGIN
        End If  ' frame visible
    Next lSection
    
End Sub

' - Set the top position for all displayed reply rows,
' - Center the displayed reply rows within the replies area,
' - Set the final height of the replies area.
' ------------------------------------------------------
Private Sub ResizeAndRepositionReplyRows()
    
    Dim frRow       As MSForms.Frame
    Dim v           As Variant
    Dim siCenter    As Single
    Dim siHeight    As Single
    
    On Error GoTo on_error
    
    siTopNext = F_MARGIN
    siCenter = RepliesArea.width / 2
    
    For Each v In ReplyRowsVisible
        Set frRow = v
        With frRow
            siHeight = .Height + V_MARGIN
            .Top = siTopNext
            siTopNext = .Top + .Height + V_MARGIN
            '~~ Center the replies row within the replies area
            .left = siCenter - (.width / 2)
        End With
    Next v
    
    With ReplyRowsVisible
        If .Count > 0 Then RepliesArea.Height = siHeight * .Count
    End With
        
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Return the collection of all visible reply rows
' -----------------------------------------------
Private Function ReplyRowsVisible() As Collection

    Dim v As Variant
    Dim cll As New Collection
    
    For Each v In cllRepliesRow
        If v.Visible Then cll.Add v
    Next v
    
    Set ReplyRowsVisible = cll
    
End Function
' Setup and position the displayed reply buttons.
' Return the max reply button width.
' ------------------------------------------------------
Private Sub ReplyButtonsSetup(ByVal vReplies As Variant)
    
    Dim v                   As Variant
    Dim row                 As Long
    Dim button              As Long
    Dim siLeftNext          As Single
    Dim cmb                 As MSForms.CommandButton
    Dim lRowButtons         As Long
    
    Set cllReplyRowButtonValues = Nothing: Set cllReplyRowButtonValues = New Collection
    
    siLeftNext = 0
    With Me
        '~~ Setup all reply button's caption and return the maximum width and height
        If IsNumeric(vReplies) Then
            '~~ Setup a row of standard VB MsgBox reply command buttons
            Select Case vReplies
                Case vbOKOnly
                    ReplyButtonSetup 1, 1, "Ok", vbOK, siLeftNext
                Case vbOKCancel
                    ReplyButtonSetup 1, 1, "Ok", vbOK, siLeftNext
                    ReplyButtonSetup 1, 2, "Cancel", vbCancel, siLeftNext
                Case vbYesNo
                    ReplyButtonSetup 1, 1, "Yes", vbYes, siLeftNext
                    ReplyButtonSetup 1, 2, "No", vbNo, siLeftNext
                Case vbRetryCancel
                    ReplyButtonSetup 1, 1, "Retry", vbRetry, siLeftNext
                    ReplyButtonSetup 1, 2, "Cancel", vbCancel, siLeftNext
                Case vbYesNoCancel
                    ReplyButtonSetup 1, 1, "Yes", vbYes, siLeftNext
                    ReplyButtonSetup 1, 2, "No", vbNo, siLeftNext
                    ReplyButtonSetup 1, 3, "Cancel", vbCancel, siLeftNext
                Case vbAbortRetryIgnore
                    ReplyButtonSetup 1, 1, "Abort", vbAbort, siLeftNext
                    ReplyButtonSetup 1, 2, "Retry", vbRetry, siLeftNext
                    ReplyButtonSetup 1, 3, "Ignore", vbIgnore, siLeftNext
            End Select
            ReplyRowsButtonValues = cllReplyRowButtonValues
            RepliesRow(row).width = ((siRepliesWidth + H_MARGIN) * RepliesSetupInRow(row)) - H_MARGIN + F_MARGIN

        Else
            '~~ Setup the reply buttons by the area of strings provided with vReplies
            '~~ Consider a new row with each vbLf element found in the array
            aReplyButtons = Split(vReplies, ",")
            row = 1
            button = 0
            For Each v In aReplyButtons
                If v <> vbNullString Then
                    If v = vbLf Then
                        '~~ Finish the setup of a rows reply buttons by adjusting
                        '~~ the surrounding frame's width and height
                        ReplyRowsButtonValues = cllReplyRowButtonValues
                        With RepliesRow(row)
                            .Height = siRepliesHeight + F_MARGIN
                            RepliesArea.width = Max(RepliesArea.width, .width) ' Adjust the area frame to the widest replies row frame
                        End With
                        RepliesArea.width = (siRepliesWidth * button) + (R_MARGIN * (button - 1))
                        '~~ prepare for the next row
                        row = row + 1
                        button = 0
                    Else
                        button = button + 1
                        RepliesRow(row).Visible = True
                        ReplyButtonSetup row:=row, button:=button, sCaption:=v, vReturnValue:=v, leftnext:=siLeftNext
                    End If
                End If
            Next v
        
            ReplyRowsButtonValues = cllReplyRowButtonValues
            With RepliesRow(row)
                .Visible = True
                .Height = siRepliesHeight + 4
                .width = ((siRepliesWidth + H_MARGIN) * RepliesSetupInRow(row)) - H_MARGIN + F_MARGIN
                RepliesArea.width = Max(RepliesArea.width, .width) ' Adjust the area frame to the widest replies row frame
            End With
            
            FormWidth = RepliesArea.width ' will extend the form width if it is a new maximum

        End If
    End With
        
    '~~ Adjust all reply buttons the top (first row) frame reply buttons width, height and left position
    '~~ Adjust the widht and height of the replies frame and the section frame accordingly
    RepliesArea.Visible = True
    ResizeAndRepositionFrames
    
    For row = 1 To cllReplyRowsButtons.Count
        If Not RepliesRow(row).Visible Then Exit For
        Set cllReplyRowButtons = cllReplyRowsButtons(row)
        siLeftNext = 0
        For button = 1 To cllReplyRowButtons.Count
            Set cmb = cllReplyRowButtons(button)
            With cmb
                If Not .Visible Then Exit For
                .width = siRepliesWidth
                .Height = siRepliesHeight
                .left = siLeftNext
                siLeftNext = siLeftNext + .width + H_MARGIN         ' set left pos for the next visible button
                RepliesRow(row).width = .left + .width        ' expand the replies frame accordingly
            End With
        Next button
    
        RepliesRow(row).width = ((siRepliesWidth + H_MARGIN) * RepliesSetupInRow(row)) - H_MARGIN + F_MARGIN
        RepliesRow(row).Height = siRepliesHeight + F_MARGIN
        With RepliesArea
            .Visible = True
            .width = Max(.width, RepliesRow(row).width)
            FormWidth = Max(Me.width, .width)
        End With
        '~~ Center the replies row within the RepliesArea
        RepliesRow(row).left = (RepliesArea.width / 2) - (RepliesRow(row).width / 2)
        '~~ Center the RepliesArea with the message form
    Next row
    With RepliesArea
        .Height = ((siRepliesHeight + V_MARGIN) * ReplyRowsSetup) + F_MARGIN
        .left = (Me.width / 2) - (RepliesArea.width / 2)
    End With
    
    ResizeAndRepositionFrames

End Sub

Private Sub TopPosReplyRows()

    Dim v As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = 0
    For Each v In cllRepliesRow
        Set fr = v
        With fr
            If .Visible = True Then
                .Top = siTopNext
                siTopNext = .Top + .Height + V_MARGIN
                .Height = siRepliesHeight + 2
                RepliesArea.Height = .Top + .Height + V_MARGIN
            End If
        End With
    Next v
    
End Sub

' Return the value of the clicked reply button (button) in row (row).
' -------------------------------------------------------------------
Private Sub ReplyClicked(ByVal row As Long, ByVal button As Long)
    
    Dim s As String
    
    s = ReplyButtonValue(row, button)
    If IsNumeric(s) Then
        mMsg.MsgReply = CLng(s)
    Else
        mMsg.MsgReply = s
    End If
    Unload Me
    
End Sub

' - Setup the reply button's (cmb) visibility and caption
' - Collect the setup command button for the row it is setup
' - Collect the setup command buttons return value when clicked
' - Keep record of the maximum button width (siRepliesWidth)
' - Keep record of the maximum button height (siRepliesHeight)
' - Return the left position for the next button (leftNext).
' -------------------------------------------------------------
Private Sub ReplyButtonSetup(ByVal row As Long, _
                             ByVal button As Long, _
                             ByVal sCaption As String, _
                             ByVal vReturnValue As Variant, _
                             ByRef leftnext As Single)
    
    Dim cmb As MSForms.CommandButton
    Set cmb = ReplyButton(row, button)
    
    With cmb
        .left = leftnext
        .Visible = True
        .AutoSize = True
        .WordWrap = False
        .Caption = sCaption
        siRepliesHeight = mMsg.Max(siRepliesHeight, .Height)
        siRepliesWidth = Max(siRepliesWidth, .width, MIN_WIDTH_REPLY_BUTTON)
        leftnext = .left + siRepliesWidth + R_MARGIN
    End With
    
    ReplyRowButtonValues = vReturnValue
    
End Sub

' An extra title label mimics the title bar in order to determine the required form's width.
' When a specific font name and/or size is specified, the extra title label is actively used
' and the UserForm's title bar is not displayed - which means that there is no X to cancel.
' ------------------------------------------------------------------------------------------
Private Sub TitleSetup(ByRef titlewidth As Single)
    
    siTopNext = 0
    With Me
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            '~~ A title with a specific font is displayed in a dedicated title label
            With .laMsgTitle   ' Hidden by default
                .Visible = True
                .Top = TopNext(Me.laMsgTitle)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .Caption = " " & sTitle    ' some left margin
                titlewidth = .width + H_MARGIN
            End With
            .laMsgTitleSpaceBottom.Visible = True
        Else
            With .laMsgTitle
                '~~ The title label is only used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.65   ' Value which comes to a length close to the length required
                End With
                .Visible = False
                .AutoSize = True
                .Caption = " " & sTitle    ' some left margin
                titlewidth = .width + 30
            End With
            .Caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
                
        .laMsgTitleSpaceBottom.width = titlewidth
        FormWidth = titlewidth
    End With
    
'    ResizeAndRepositionFrames
    
End Sub

' Re-positioning of displayed frames, usually after a width expansion of the UserForm
' -----------------------------------------------------------------------------------
Private Sub ResizeAndRepositionFrames()
    ResizeAndRepositionMsgSections
    ResizeAndRepositionReplyRows
    ResizeAndRepositionAreas
End Sub

' Resize all Message Section relvant (i.e. visible) frames
' and re-position them vertically.
' --------------------------------------------------------
Private Sub ResizeAndRepositionMsgSections()
    
    Dim v               As Variant
    Dim lb              As MSForms.Label
    Dim frSection       As MSForms.Frame
    Dim frText          As MSForms.Frame
    Dim tbText          As MSForms.TextBox
    Dim lSection        As Long
    Dim siMaxWidth      As Single
    
    On Error GoTo on_error
    
    siTopNext = 0
    For lSection = 1 To cllMsgSections.Count
        
        Set frSection = MsgSection(lSection)
        Set lb = MsgSectionLabel(lSection)
        Set frText = MsgSectionTextFrame(lSection)
        Set tbText = MsgSectionText(lSection)
        
        If frSection.Visible Then
            frSection.Top = siTopNext
            If lb.Visible Then
                lb.Top = 0
                frText.Top = lb.Top + lb.Height + F_MARGIN
            Else
                frText.Top = 0
            End If
            tbText.Top = 0
            frText.Height = tbText.Height + F_MARGIN
            frText.width = tbText.width + F_MARGIN
            With frSection
                .Height = frText.Top + frText.Height
                .width = frText.width + F_MARGIN
                FormWidth = .width
                siMaxWidth = Max(siMaxWidth, .width)
                MsgArea.Height = .Top + .Height
                siTopNext = .Top + .Height + V_MARGIN
            End With
        
            MsgArea.width = siMaxWidth
        End If
    Next lSection
       
    Me.Height = Max(Me.Height, MsgArea.Top + MsgArea.Height + (V_MARGIN * 4))
    
exit_proc:
    DoEvents
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub
 
Private Sub ResizeAndRepositionAreas()

    siTopNext = T_MARGIN
    
    With MsgArea
        If .Visible Then
            .Top = siTopNext
            siTopNext = .Top + .Height + V_MARGIN
        End If
    End With
    
    '~~ The replies area is centered
    With RepliesArea
        If .Visible Then
            .Top = siTopNext
            siTopNext = .Top + .Height + V_MARGIN
        End If
        .left = (Me.width / 2) - (.width / 2)
    End With
    
    Me.Height = siTopNext - V_MARGIN + (V_MARGIN * 4)
    
End Sub

' Get coordinates of top-left corner and size of entire screen (stretched over
' all monitors) and convert to Points.
' ----------------------------------------------------------------------------
Private Sub GetScreenMetrics()
    
    wVirtualScreenLeft = GetSystemMetrics32(SM_XVIRTUALSCREEN)
    wVirtualScreenTop = GetSystemMetrics32(SM_YVIRTUALSCREEN)
    wVirtualScreenWidth = GetSystemMetrics32(SM_CXVIRTUALSCREEN)
    wVirtualScreenHeight = GetSystemMetrics32(SM_CYVIRTUALSCREEN)
    '
    ConvertPixelsToPoints wVirtualScreenLeft, wVirtualScreenTop
    ConvertPixelsToPoints wVirtualScreenWidth, wVirtualScreenHeight

End Sub

Public Sub AdjustStartupPosition(ByRef pUserForm As Object, _
                        Optional ByRef pOwner As Object)
    On Error Resume Next
    
'    GetScreenMetrics
    
    Select Case pUserForm.StartupPosition
        Case Manual, WindowsDefault ' Do nothing
        
        Case CenterOwner            ' Position centered on top of the 'Owner'. Usually this is Application.
            If Not pOwner Is Nothing Then Set pOwner = Application
            With pUserForm
                .StartupPosition = 0
                .left = pOwner.left + ((pOwner.width - .width) / 2)
                .Top = pOwner.Top + ((pOwner.Height - .Height) / 2)
            End With
            
        Case CenterScreen           ' Assign the Left and Top properties after switching to Manual positioning.
            With pUserForm
                .StartupPosition = Manual
                .left = (wVirtualScreenWidth - .width) / 2
                .Top = (wVirtualScreenHeight - .Height) / 2
            End With
    End Select
 
    ' Avoid falling off screen. Misplacement can be caused by multiple screens when the primary display
    ' is not the left-most screen (which causes "pOwner.Left" to be negative). First make sure the bottom
    ' right fits, then check if the top-left is still on the screen (which gets priority).
    '
    With pUserForm
        If ((.left + .width) > (wVirtualScreenLeft + wVirtualScreenWidth)) _
        Then .left = ((wVirtualScreenLeft + wVirtualScreenWidth) - .width)
        If ((.Top + .Height) > (wVirtualScreenTop + wVirtualScreenHeight)) _
        Then .Top = ((wVirtualScreenTop + wVirtualScreenHeight) - .Height)
        If (.left < wVirtualScreenLeft) Then .left = wVirtualScreenLeft
        If (.Top < wVirtualScreenTop) Then .Top = wVirtualScreenTop
    End With
End Sub
 
' Returns pixels (device dependent) to points (used by Excel).
' --------------------------------------------------------------------
Private Sub ConvertPixelsToPoints(ByRef X As Single, ByRef Y As Single)
    
    Dim hDC            As Long
    Dim RetVal         As Long
    Dim PixelsPerInchX As Long
    Dim PixelsPerInchY As Long
 
    On Error Resume Next
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    X = X * TWIPSPERINCH / 20 / PixelsPerInchX
    Y = Y * TWIPSPERINCH / 20 / PixelsPerInchY

End Sub

Private Sub UserForm_Activate()
    
    Dim siTitleWidth    As Single
    Dim i               As Long

    DisplayFramesWithCaption bFramesWithCaption
    
    With Me

        '~~ ----------------------------------------------------------------------------------------
        '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
        '~~ returns their individual widths which determines the minimum required message form width
        '~~ This setup ends width the final message form width and all elements adjusted to it.
        '~~ ----------------------------------------------------------------------------------------
        .width = siMinimumFormWidth ' Setup starts with the minimum message form width

        '~~ Setup of the first element which determines the form width
        TitleSetup siTitleWidth
        
        '~~ Setup of monospaced message sections which determine the form width
        MsgSectionsMonospacedSetup           ' Setup monospaced message sections
        
        '~~ Setup of the second element which determines the form width
        ReplyButtonsSetup vReplies     ' Reply buttons text, size and visibility
        
        '~~ Setup of monospaced message sections which determine the form width
        MsgSectionsProportionalSetup         ' Setup proportional spaced message sections
        
        '~~ Determine the minimum required message form width based on the sizes returned from the setup
        '~~ and reduce it if it exceeds the maximum form width specified
        If .width > siMaximumFormWidth Then .width = siMaximumFormWidth ' reduce to maximum when exceeded
        DoEvents
        
        '~~ Adjust all message sections to the final form width. Message sections with a proportional font
        '~~ may be enlarged or shrinked (which will result in a new height). Monospaced message sections
        '~~ when shrinked in their width will receive a horizontal scroll bar.
        MsgSectionsFinalWidth
        
        '~~ ---------------------------------------------------------------------------------------------
        '~~ The  f i n a l  setup considers the height required for the message sections and the reply
        '~~ buttons. This height is reduced when it exceeds the maximum height specified (as a percentage
        '~~ of the available screen size). The largest message section may receive a vertical scroll bar.
        '~~ ---------------------------------------------------------------------------------------------
        If .Height > siMaximumFormHeight Then
        '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
            .Height = siMaximumFormHeight
            FormHeightFinal
        End If
        DoEvents
                
        ResizeAndRepositionFrames
        
    End With
    
    AdjustStartupPosition Me

End Sub

' When False (the default) captions are removed from all frames
' Else they remain visible for testing purpose
' -------------------------------------------------------------
Private Sub DisplayFramesWithCaption(ByVal b As Boolean)
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.Caption = vbNullString
            End If
        Next ctl
    End If

End Sub

' Collect controls of type ctltype with a parent fromparent
' by assigning an initial height and width
' --------------------------------------------------
Private Sub Collect(ByRef into As Collection, _
                    ByVal fromparent As Variant, _
                    ByVal ctltype As String, _
                    ByVal ctlheight As Single, _
                    ByVal ctlwidth As Single)

    Dim ctl As MSForms.Control
    Dim v   As Variant
     
    On Error GoTo on_error
    
    Set into = Nothing: Set into = New Collection
    Select Case TypeName(fromparent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In fromparent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = ctltype And ctl.Parent Is v Then
                        With ctl
                            Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Ctl: " & ctl.Name
                            .Visible = False
                            .Height = ctlheight
                            .width = ctlwidth
                        End With
                        into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = ctltype And ctl.Parent Is fromparent Then
                    With ctl
                        Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Ctl: " & ctl.Name
                        .Visible = False
                        .Height = ctlheight
                        .width = ctlwidth
                    End With
                    into.Add ctl
                End If
            Next ctl
    End Select
exist_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub
