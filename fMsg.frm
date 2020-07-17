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
'         1.2       Frame MsgSections
'         1.2.1     Frame MsgSection1 to ...3
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
Const MONOSPACED_FONT_NAME      As String = "Courier New"   ' Default monospaced font
Const MONOSPACED_FONT_SIZE      As Single = 9               ' Default monospaced font size
Const FORM_WIDTH_MIN            As Single = 300
Const FORM_WIDTH_MAX_POW        As Long = 80    ' Maximum form width as a percentage of the screen size
Const FORM_HEIGHT_MAX_POW       As Long = 90    ' Maximum form height as a percentage of the screen size
Const L_MARGIN                  As Single = 0   ' Left margin for labels and text boxes
Const R_MARGIN                  As Single = 15  ' Right margin for labels and text boxes
Const H_MARGIN                  As Single = 10  ' Horizontal margin for reply buttons
Const V_MARGIN                  As Single = 10  ' Vertical marging for all displayed elements/controls
Const T_MARGIN                  As Single = 5   ' Top position for the first displayed control
Const B_MARGIN                  As Single = 50  ' Bottom margin after the last displayed control
Const MARGIN_VERTIVAL_LABEL     As Single = 5
Const REPLY_BUTTON_MIN_WIDTH    As Single = 70

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

Dim sTitle                      As String
Dim sErrSrc                     As String
Dim vReplies                    As Variant
Dim aReplyButtons               As Variant
Dim sTitleFontName              As String
Dim sTitleFontSize              As String   ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim sMsgSection1Label           As String
Dim sMsgSection1Text            As String
Dim bMsgSection1Monospaced      As Boolean
Dim sMsgSection2Label           As String
Dim sMsgSection2Text            As String
Dim bMsgSection2Monospaced      As Boolean
Dim sMsgSection3Label           As String
Dim sMsgSection3Text            As String
Dim bMsgSection3Monospaced      As Boolean
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
Dim cllMsgSectionsLabel         As New Collection
Dim cllMsgSections              As New Collection   '
Dim cllMsgSectionsText          As New Collection   ' Collection of section frames
Dim cllSectionsVisible          As New Collection   ' Collection of visible section frames
Dim cllReplyButtons1            As New Collection
Dim cllReplyButtons2            As New Collection
Dim cllReplyButtonsVisible1     As New Collection
Dim cllReplyButtonsVisible2     As New Collection
Dim cllReplyButtonsValue1       As New Collection
Dim cllReplyButtonsValue2       As New Collection
Dim cllReplyRows                As New Collection
Dim cllMsgFrames                As New Collection
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim dctSectionsLabel            As New Dictionary
Dim dctSectionsText             As New Dictionary
Dim dctSectionsMonospaced       As New Dictionary
Dim siRepliesWidth              As Single
Dim siRepliesHeight             As Single

Private Sub UserForm_Initialize()
    
    Dim ctl As MSForms.Control
    Dim fr  As MSForms.Frame
    Dim v   As Variant
    
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_WIDTH_MAX_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_HEIGHT_MAX_POW
    siMinimumFormWidth = FORM_WIDTH_MIN                         ' Default UserForm width
    sMonospacedFontName = MONOSPACED_FONT_NAME                  ' Default monospaced font
    siMonospacedFontSize = MONOSPACED_FONT_SIZE                 ' Default monospaced font
    bMsgSection1Monospaced = False
    bMsgSection2Monospaced = False
    bMsgSection3Monospaced = False
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" And ctl.Parent Is Me Then
            cllAreas.Add ctl
        End If
    Next ctl
    
    '~~ Message Section Frames (grouping the label and the message text)
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" And ctl.Parent Is Me.frFormSectionMessage Then
            cllMsgSections.Add ctl
        End If
    Next ctl
    
    '~~ Message sections label
    For Each v In cllMsgSections
        Set fr = v
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Label" And ctl.Parent Is fr Then
                cllMsgSectionsLabel.Add ctl
            End If
        Next ctl
    Next v
            
    '~~ Message section text frame
    For Each v In cllMsgSections
        Set fr = v
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" And ctl.Parent Is fr Then
                cllMsgFrames.Add ctl
            End If
        Next ctl
    Next v
    
    '~~ Message text (in the text frames)
    For Each v In cllMsgFrames
        Set fr = v
        For Each ctl In Me.Controls
            If TypeName(ctl) = "TextBox" And ctl.Parent Is fr Then
                cllMsgSectionsText.Add ctl
            End If
        Next ctl
    Next v
        
    '~~ Reply rows
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" And ctl.Parent Is RepliesArea Then
            cllReplyRows.Add ctl
        End If
    Next ctl
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CommandButton" And ctl.Parent Is Me.frRepliesRow1 Then
            cllReplyButtons1.Add ctl
        End If
    Next ctl
    
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CommandButton" And ctl.Parent Is Me.frRepliesRow2 Then
            cllReplyButtons2.Add ctl
        End If
    Next ctl
    
    bWithFrames = False
    
End Sub

Private Property Get Monospaced(Optional ByVal section As Long) As Boolean
    Monospaced = MsgSection(section).Font.Name = sMonospacedFontName
End Property
Private Property Let Monospaced(Optional ByVal section As Long, ByVal monospace As Boolean)
    MsgSection(section).Font.Name = sMonospacedFontName
End Property

' This property is for testing purpose only. It default to False
' and may be used to see the width and height of the elements.
' --------------------------------------------------------------
Public Property Let DisplayElementsWithFrame(ByVal withframes As Boolean)
    
    Dim ctl As MSForms.Control
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "Frame", "TextBox"
                ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
                If withframes _
                Then ctl.BorderStyle = fmBorderStyleSingle _
                Else ctl.BorderStyle = fmBorderStyleNone
        End Select
    Next ctl

End Property
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

Private Property Get ReplyButton(Optional ByVal row As Long, Optional ByVal Button As Long) As MSForms.CommandButton
    Select Case row
        Case 1: Set ReplyButton = cllReplyButtons1(Button)
        Case 2: Set ReplyButton = cllReplyButtons2(Button)
    End Select
End Property

Private Property Get ReplyButtonValue(Optional ByVal row As Long, Optional ByVal Button As Long)
    Select Case row
        Case 1: ReplyButtonValue = cllReplyButtonsValue1(Button)
        Case 2: ReplyButtonValue = cllReplyButtonsValue2(Button)
    End Select
End Property

Private Property Let ReplyButtonValue(Optional ByVal row As Long, Optional ByVal Button As Long, ByVal v As Variant)
    Select Case row
        Case 1: cllReplyButtonsValue1.Add v
        Case 2: cllReplyButtonsValue2.Add v
    End Select
End Property

Private Property Let FormWidth(ByVal w As Single)
    w = Max(siMinimumFormWidth, w)
    Me.width = w
    MsgArea.width = w - R_MARGIN
    RepliesArea.width = w - R_MARGIN
End Property

' UserForm design elements
Private Property Get Areas() As Collection:                                                 Set Areas = cllAreas:                               End Property
Private Property Get MsgArea() As MSForms.Frame:                                            Set MsgArea = cllAreas(1):                          End Property
Private Property Get RepliesArea() As MSForms.Frame:                                        Set RepliesArea = cllAreas(2):                      End Property
Private Property Get ReplyRows() As Collection:                                             Set ReplyRows = cllReplyRows:                       End Property
Private Property Get RepliesRow(Optional ByVal row As Long) As MSForms.Frame:               Set RepliesRow = cllReplyRows(row):                 End Property
Private Property Get MsgFrames() As Collection:                                             Set MsgFrames = cllMsgFrames:                       End Property
Private Property Get MsgFrame(Optional ByVal section As Long) As MSForms.Frame:             Set MsgFrame = cllMsgFrames(section):               End Property
Private Property Get MsgSections() As Collection:                                           Set MsgSections = cllMsgSections:                   End Property
Private Property Get MsgSectionLabel(Optional section As Long) As MSForms.Label:            Set MsgSectionLabel = cllMsgSectionsLabel(section): End Property
Private Property Get MsgSection(Optional i As Long) As MSForms.Frame:                       Set MsgSection = cllMsgSections(i):                 End Property
Private Property Get MsgSectionText(Optional i As Long) As MSForms.TextBox:                 Set MsgSectionText = cllMsgSectionsText(i):         End Property

' Message section propertis (label, text, monospaced)
Public Property Let SectionsLabel(Optional ByVal section As Long, ByVal s As String):       dctSectionsLabel.Add section, s:                    End Property
Public Property Get SectionsLabel(Optional ByVal section As Long) As String
    If dctSectionsLabel.Exists(section) _
    Then SectionsLabel = dctSectionsLabel(section) _
    Else SectionsLabel = vbNullString
End Property

Public Property Let SectionsText(Optional ByVal section As Long, ByVal s As String):        dctSectionsText.Add section, s:                     End Property
Public Property Get SectionsText(Optional ByVal section As Long) As String
    If dctSectionsText.Exists(section) _
    Then SectionsText = dctSectionsText(section) _
    Else SectionsText = vbNullString
End Property

Public Property Let SectionsMonospaced(Optional ByVal section As Long, ByVal b As Boolean): dctSectionsMonospaced.Add section, b:               End Property
Public Property Get SectionsMonospaced(Optional ByVal section As Long) As Boolean
    If dctSectionsMonospaced.Exists(section) _
    Then SectionsMonospaced = dctSectionsMonospaced(section) _
    Else SectionsMonospaced = False
End Property
Public Property Let ErrSrc(ByVal s As String):                                              sErrSrc = s:                                        End Property
Public Property Let Replies(ByVal v As Variant):                                            vReplies = v:                                       End Property
Public Property Let Title(ByVal s As String):                                               sTitle = s:                                         End Property

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

    For Each v In MsgSectionsVisible
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
            ByVal section As Long, _
            ByRef maxmsgwidth As Single)
    
    Dim frArea                  As MSForms.Frame
    Dim frMsgSection            As MSForms.Frame
    Dim laMsgSectionLabel       As MSForms.Label
    Dim tbMsgSectionText        As MSForms.TextBox
    Dim frMsgSectionTextFrame   As MSForms.Frame
    Dim sMsgLabel               As String
    Dim sMsgText                As String
    Dim bMsgMonospaced          As Boolean

    Set frArea = MsgArea
    Set frMsgSection = MsgSection(section)
    Set laMsgSectionLabel = MsgSectionLabel(section)
    Set tbMsgSectionText = MsgSectionText(section)
    Set frMsgSectionTextFrame = MsgFrame(section)
    
    sMsgLabel = SectionsLabel(section)
    sMsgText = SectionsText(section)
    bMsgMonospaced = SectionsMonospaced(section)
    
    frMsgSection.width = frArea.width
    laMsgSectionLabel.width = frMsgSection.width
    frMsgSectionTextFrame.width = frMsgSection.width
    tbMsgSectionText.width = frMsgSection.width
    
    With frArea
        .Visible = True
        .left = 0
        .width = Me.width - R_MARGIN
    End With
    frMsgSection.width = frArea.width
    
    If sMsgText <> vbNullString Then
        frMsgSection.Visible = True
        '~~ Setup above text label/title only when there is a text
        If sMsgLabel <> vbNullString Then
            Set laMsgSectionLabel = MsgSectionLabel(section)
            With laMsgSectionLabel
                .width = frMsgSection.width
                .Caption = sMsgLabel
                .Visible = True
            End With
            frMsgSectionTextFrame.Top = laMsgSectionLabel.Top + laMsgSectionLabel.Height
        Else
            frMsgSectionTextFrame.Top = 0
        End If
        
        If bMsgMonospaced Then
            MsgSectionSetupMonospaced section, sMsgText, maxmsgwidth  ' returns the maximum width required for monospaced section
            If FormSectionExceedsMaxFormWidth(section) Then
                AddHorizontalScrollBarToFrame section ' only applied for monospaced section text
            End If
        Else ' proportional spaced
            MsgSectionSetupProportional section, sMsgText
        End If
        DoEvents
        tbMsgSectionText.SelStart = 0
        
    End If
    frMsgSectionTextFrame.Height = tbMsgSectionText.Height
    frMsgSection.Height = frMsgSectionTextFrame.Top + frMsgSectionTextFrame.Height
    frArea.Height = frMsgSection.Top + frMsgSection.Height + V_MARGIN
    Me.Height = Max(Me.Height, siTopNext + (V_MARGIN * 4))
    TopPosMsgSections

End Sub

Private Sub MsgSectionSetupProportional(ByVal section As Long, _
                                        ByVal text As String)
    
    Dim tbMsgText   As MSForms.TextBox
    Set tbMsgText = MsgSectionText(section)
    
    '~~ Setup the textbox
    With tbMsgText
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = Me.width - L_MARGIN
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
                                       
    TopPosMsgSections
    
End Sub
                                       
Private Sub AddHorizontalScrollBarToFrame(ByVal section As Long)
    
    Dim frFormSection       As MSForms.Frame
    Dim frMsgSection        As MSForms.Frame
    Dim tbMsgSectionText    As MSForms.TextBox
    
    Set frFormSection = Areas(1)
    Set frMsgSection = MsgSection(section)
    Set tbMsgSectionText = MsgSectionText(section)

    frFormSection.width = siMaximumFormWidth - L_MARGIN - R_MARGIN
    frMsgSection.width = frFormSection.width - 2
    frMsgSection.Height = tbMsgSectionText.Height + 15 ' space for the scroll bar
    frFormSection.Height = frMsgSection.Height + 15
    With frMsgSection
        Select Case .ScrollBars
            Case fmScrollBarsBoth
            Case fmScrollBarsHorizontal
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsHorizontal
            Case fmScrollBarsVertical
                .ScrollBars = fmScrollBarsHorizontal
        End Select
        .ScrollWidth = tbMsgSectionText.width
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
            ByVal text As String, _
            ByRef maxmsgwidth As Single)
            
    Dim tbMsgText   As MSForms.TextBox
    Set tbMsgText = MsgSectionText(section)
    
    '~~ Setup the textbox
    With tbMsgText
        .MultiLine = True
        .AutoSize = True
        .WordWrap = False
        .Font.Name = sMonospacedFontName
        .Font.Size = siMonospacedFontSize
        .value = text
    End With
        
    ' Adjust surrounding frames accordingly
    With MsgSection(section)
        .width = tbMsgText.width
        .Height = tbMsgText.Height
    End With
    With MsgArea
        .width = tbMsgText.width
        .Height = tbMsgText.Height
        Me.width = .width + R_MARGIN
    End With
    
    TopPosMsgSections
    
End Sub

Private Function FormSectionExceedsMaxFormWidth(ByVal section As Long) As Boolean
    FormSectionExceedsMaxFormWidth = MsgSection(section).width + L_MARGIN + R_MARGIN > siMaximumFormWidth
End Function


' Returns a collection of the visible message sections.
' -----------------------------------------------------
Private Function MsgSectionsVisible() As Collection
    
    Dim v   As Variant
    Dim tb  As MSForms.TextBox
    Dim cll As New Collection
    
    For Each v In cllMsgSectionsText
        Set tb = v
        If tb.Visible = True Then
            cll.Add tb
        End If
    Next v
    Set MsgSectionsVisible = cll

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

Private Sub MsgSectionsMonospacedSetup(ByRef msgmaxwidth As Single)
                             
    If SectionsText(1) <> vbNullString And SectionsMonospaced(1) = True Then MsgSectionSetup section:=1, maxmsgwidth:=msgmaxwidth
    If SectionsText(2) <> vbNullString And SectionsMonospaced(2) = True Then MsgSectionSetup section:=2, maxmsgwidth:=msgmaxwidth
    If SectionsText(3) <> vbNullString And SectionsMonospaced(3) = True Then MsgSectionSetup section:=3, maxmsgwidth:=msgmaxwidth
    
End Sub

Private Sub MsgSectionsProportionalSetup(ByRef msgmaxwidth As Single)
                
    If SectionsText(1) <> vbNullString And SectionsMonospaced(1) = False Then MsgSectionSetup section:=1, maxmsgwidth:=msgmaxwidth
    If SectionsText(2) <> vbNullString And SectionsMonospaced(2) = False Then MsgSectionSetup section:=2, maxmsgwidth:=msgmaxwidth
    If SectionsText(3) <> vbNullString And SectionsMonospaced(3) = False Then MsgSectionSetup section:=3, maxmsgwidth:=msgmaxwidth
    
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

' Setup the left position of each setup/visible reply button
' based on the maximum reply button width
' and adjust their frame's width.
' ----------------------------------------------------------
Private Sub RepliesPosAdjust()
    
    Dim fr      As MSForms.Frame
    Dim lRow    As Long
    
    With RepliesArea
        For lRow = 1 To cllReplyRows.Count
            Set fr = RepliesRow(lRow)
            fr.left = (Me.width / 2) - (fr.width / 2)
            .Height = fr.Top + fr.Height + (V_MARGIN * 4) ' replies section height
        Next lRow
    End With
End Sub

' Setup and position the displayed reply buttons.
' Return the max reply button width.
' ------------------------------------------------------
Private Sub ReplyButtonsSetup(ByVal vReplies As Variant)
    
    Dim v                   As Variant
    Dim lRow                As Long
    Dim lButton             As Long
    Dim lButtons            As Long
    Dim siLeft              As Single
    Dim frRepliesSection    As MSForms.Frame
    Dim frRepliesRow        As MSForms.Frame
    
    Set frRepliesSection = RepliesArea
    
    Set cllReplyButtonsValue1 = Nothing: Set cllReplyButtonsValue1 = New Collection
    Set cllReplyButtonsValue2 = Nothing: Set cllReplyButtonsValue2 = New Collection

    siLeft = 0
    With Me
        '~~ Setup all reply button's caption and return the maximum width and height
        If IsNumeric(vReplies) Then
            Select Case vReplies
                Case vbOKOnly
                    ReplyButtonSetup 1, 1, "Ok", vbOK, siLeft
                Case vbOKCancel
                    ReplyButtonSetup 1, 1, "Ok", vbOK, siLeft
                    ReplyButtonSetup 1, 2, "Cancel", vbCancel, siLeft
                Case vbYesNo
                    ReplyButtonSetup 1, 1, "Yes", vbYes, siLeft
                    ReplyButtonSetup 1, 2, "No", vbNo, siLeft
                Case vbRetryCancel
                    ReplyButtonSetup 1, 1, "Retry", vbRetry, siLeft
                    ReplyButtonSetup 1, 2, "Cancel", vbCancel, siLeft
                Case vbYesNoCancel
                    ReplyButtonSetup 1, 1, "Yes", vbYes, siLeft
                    ReplyButtonSetup 1, 2, "No", vbNo, siLeft
                    ReplyButtonSetup 1, 3, "Cancel", vbCancel, siLeft
                Case vbAbortRetryIgnore
                    ReplyButtonSetup 1, 1, "Abort", vbAbort, siLeft
                    ReplyButtonSetup 1, 2, "Retry", vbRetry, siLeft
                    ReplyButtonSetup 1, 3, "Ignore", vbIgnore, siLeft
            End Select
            lButtons = cllReplyButtonsValue1.Count
        Else
            aReplyButtons = Split(vReplies, ",")
            For Each v In aReplyButtons
                If v <> vbNullString Then
                    lButton = lButton + 1
                    If lButton <= 5 Then
                        RepliesRow(1).Visible = True
                        ReplyButtonSetup row:=1, Button:=lButton, s:=v, v:=v, left:=siLeft
                    Else
                        RepliesRow(2).Visible = True
                        ReplyButtonSetup row:=2, Button:=lButton, s:=v, v:=v, left:=siLeft
                    End If
                End If
            Next v
        End If
            
        lButtons = cllReplyButtonsValue1.Count
        Set frRepliesRow = RepliesRow(1)
        frRepliesRow.Height = siRepliesHeight + 4
        frRepliesSection.Height = frRepliesRow.Height
        frRepliesSection.width = (siRepliesWidth * lButtons) + (R_MARGIN * (lButton - 1))
        Me.width = Max(Me.width, RepliesArea.width)
    End With
        
    '~~ Adjust the top (first row) frame reply buttons width, height and left position
    '~~ Adjust the widht and height of the replies frame and the section frame accordingly
    siLeft = 0
    For lButton = 1 To lButtons
        With ReplyButton(1, lButton)
            If .Visible = False Then Exit For
            .width = siRepliesWidth
            .Height = siRepliesHeight
            .left = siLeft
            siLeft = siLeft + .width + H_MARGIN         ' set left pos for the next visible button
            RepliesRow(1).width = .left + .width        ' expand the replies frame accordingly
        End With
    Next lButton
    
    RepliesRow(1).width = RepliesRow(1).width + 2        ' expand the replies frame accordingly
    RepliesRow(1).Height = siRepliesHeight + 4
    FormWidth = RepliesRow(1).width
    RepliesRow(1).left = (RepliesArea.width / 2) - (RepliesRow(1).width / 2)
        
    lButtons = cllReplyButtonsValue2.Count
    If lButtons > 0 Then
        '~~ Adjust the bottom (second row) reply buttons' width, height and left position
        '~~ Adjust the widht and height of the replies frame and the section frame accordingly
        With RepliesRow(2)
            .Visible = True
            .Height = siRepliesHeight
            .Top = Me.frRepliesRow1.Top + siRepliesHeight + V_MARGIN
        End With
        siLeft = 0
        For lButton = 1 To lButtons
            With ReplyButton(2, lButton)
                If .Visible = False Then Exit For
                .width = siRepliesWidth
                .Height = siRepliesHeight
                .left = siLeft
                siLeft = siLeft + .width + H_MARGIN   ' set left pos for the next visible button
                RepliesRow(2).width = .left + .width    ' expand frame accordingly
            End With
        Next lButton
        
        RepliesRow(2).width = RepliesRow(2).width + 2        ' expand the replies frame accordingly
        FormWidth = RepliesRow(2).width
        RepliesRow(2).left = (RepliesArea.width / 2) - (RepliesRow(2).width / 2)
    
    End If
    
    RepliesPosAdjust
    TopPosReplyRows

End Sub

Private Sub TopPosReplyRows()

    Dim v As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = 0
    For Each v In cllReplyRows
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
Private Sub ReplyClicked(ByVal row As Long, ByVal Button As Long)
    
    Dim s As String
    
    s = ReplyButtonValue(row, Button)
    If IsNumeric(s) Then
        mMsg.MsgReply = CLng(s)
    Else
        mMsg.MsgReply = s
    End If
#If test = 0 Then ' allows assertions during testing
    Unload Me
#Else
    Me.Hide
#End If
    
End Sub

' Setup a reply button's visibility and caption
' and return the maximum button width and height and the left pos for the next button.
' -----------------------------------------------
Private Sub ReplyButtonSetup(ByVal row As Long, _
                             ByVal Button As Long, _
                             ByVal s As String, _
                             ByVal v As Variant, _
                             ByRef left As Single)
    
    Dim cmb             As MSForms.CommandButton
    Dim frFormSection   As MSForms.Frame
    Dim frRepliesRow    As MSForms.Frame
    
    Set frFormSection = RepliesArea
    Set frRepliesRow = RepliesRow(row)
    
    If s <> vbNullString Then
        Set cmb = ReplyButton(row, Button)
        With cmb
            .left = left
            .Visible = True
            .AutoSize = True
            .WordWrap = False
            .Caption = s
            siRepliesHeight = mMsg.Max(siRepliesHeight, .Height)
            siRepliesWidth = Max(siRepliesWidth, .width, REPLY_BUTTON_MIN_WIDTH)
            left = (siRepliesWidth + R_MARGIN) * Button
        End With
        frRepliesRow.Height = siRepliesHeight + 2
        frFormSection.Height = frRepliesRow.Top + frRepliesRow.Height + V_MARGIN
        ReplyButtonValue(row, Button) = v
    End If
    
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
            End With
            .laMsgTitleSpaceBottom.Visible = True
        Else
            With .laMsgTitle
                '~~ The title label is only used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.65    ' Value which comes to a length close to the length required
                End With
                .Visible = False
                siTitleWidth = .width + H_MARGIN
            End With
            siTopNext = T_MARGIN
            .Caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
        
        With .laMsgTitle
            '~~ The title label is used to adjust the form width
            .WordWrap = False
            .AutoSize = True
            .Caption = " " & sTitle                 ' some left margin
            .AutoSize = False
            titlewidth = .width + H_MARGIN ' criteria for the final message form width
        End With
        
        .laMsgTitleSpaceBottom.width = titlewidth
        FormWidth = titlewidth
    End With
    
    TopPosAreas
    
End Sub

Public Sub FormFinalPositionOnScreen()
    AdjustStartupPosition Me
End Sub

' Position the control (ctl) at the current next top position (siTopNext)
' and increase the next top position.
' -----------------------------------------------------------------------
Private Sub TopPosMsgSections()
    
    Dim v   As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = 0
    For Each v In cllMsgSections
        Set fr = v
        With fr
            If .Visible Then
                .Top = siTopNext
                siTopNext = .Top + .Height + V_MARGIN
            End If
        End With
    Next v

End Sub
 
Private Sub TopPosAreas()

    Dim fr1 As MSForms.Frame, fr2 As MSForms.Frame
    Set fr1 = MsgArea
    Set fr2 = RepliesArea
    fr1.Top = T_MARGIN
    fr2.Top = fr1.Top + fr1.Height + V_MARGIN
    Me.Height = fr2.Top + fr2.Height + (V_MARGIN * 3)
    
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
    
    GetScreenMetrics
    
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
    Dim siMaxMsgWidth   As Single
    Dim i               As Long

    With Me
        '~~ ----------------------------------------------------------------------------------------
        '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
        '~~ returns their individual widths which determines the minimum required message form width
        '~~ This setup ends width the final message form width and all elements adjusted to it.
        '~~ ----------------------------------------------------------------------------------------
        FormWidth = siMinimumFormWidth

        '~~ Setup of the first element which determines the form width
        TitleSetup siTitleWidth
        
        '~~ Setup of monospaced message sections which determine the form width
        MsgSectionsMonospacedSetup msgmaxwidth:=siMaxMsgWidth           ' Setup monospaced message sections
        
        '~~ Setup of the second element which determines the form width
        ReplyButtonsSetup vReplies     ' Reply buttons text, size and visibility
        
        '~~ Setup of monospaced message sections which determine the form width
        MsgSectionsProportionalSetup msgmaxwidth:=siMaxMsgWidth         ' Setup proportional spaced message sections
        
        '~~ Determine the minimum required message form width based on the sizes returned from the setup
        '~~ and reduce it if it exceeds the maximum form width specified
        If .width > siMaximumFormWidth Then .width = siMaximumFormWidth ' reduce to maximum when exceeded
        DoEvents
        
        '~~ Adjust all message sections to the final form width. Message sections with a proportional font
        '~~ may be enlarged or shrinked (which will result in a new height). Monospaced message sections
        '~~ when shrinked in their width will receive a horizontal scroll bar.
        MsgSectionsFinalWidth
        TopPosMsgSections          ' adjusts all controls' top position
        
        TopPosReplyRows
        
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
                
        TopPosMsgSections   ' adjusts all controls' top position
        TopPosReplyRows
        
        TopPosAreas
        
    End With
    
    AdjustStartupPosition Me

End Sub


