VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
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
' lScreenWidth. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const FONT_NAME_MONOSPACED      As String = "Courier New"   ' Default monospaced font
Const FORM_WIDTH_MIN            As Single = 200
Const FORM_WIDTH_MAX_POW        As Long = 80    ' Maximum form width as a percentage of the screen size
Const FORM_HEIGHT_MAX_POW       As Long = 90    ' Maximum form height as a percentage of the screen size
Const FORM_MARGIN_LEFT          As Single = 5
Const FORM_MARGIN_RIGHT         As Single = 15
Const MARGIN_HORIZONTAL         As Single = 10
Const MARGIN_VERTICAL           As Single = 10
Const MARGIN_FORM_TOP           As Single = 5   ' Top position of the first visible control
Const MARGIN_FORM_BOTTOM        As Single = 50
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
Dim siMaximumFormHeight         As Single     ' above converted to excel userform height
Dim siMaximumFormWidth          As Single     ' above converted to excel userform width
Dim siMinimumFormHeight         As Single
Dim siMinimumFormWidth          As Single
Dim sFontNameMonospaced         As String
Dim cllMsgSections              As New Collection
Dim cllMsgSectionsVisible       As New Collection
Dim cllReplyButtons             As New Collection
Dim cllReplyButtonsVisible      As New Collection
Dim cllReplyButtonsValue        As New Collection
Dim cllMsgLabels                As New Collection
Dim siMaxMonospacedTextWidth    As Single
Dim siMaxMsgWidth               As Single
Dim siMaxReplyWidth             As Single
Dim siMaxReplyHeight            As Single


Private Sub UserForm_Initialize()
    lMaximumFormWidthPoW = FORM_WIDTH_MAX_POW
    lMaximumFormHeightPoW = FORM_HEIGHT_MAX_POW
    siMinimumFormWidth = FORM_WIDTH_MIN         ' Default UserForm width
    GetScreenMetrics            ' provides the screen's width and height
    siMaximumFormWidth = wVirtualScreenWidth * (lMaximumFormWidthPoW / 100)      ' Default maximum form width
    siMaximumFormHeight = wVirtualScreenHeight * (lMaximumFormWidthPoW / 100)   ' Default maximum form height
    sFontNameMonospaced = FONT_NAME_MONOSPACED                          ' Default monospaced font
    bMsgSection1Monospaced = False
    bMsgSection2Monospaced = False
    bMsgSection3Monospaced = False
    
    With Me
        cllMsgSections.Add .tbMsgSection1
        cllMsgSections.Add .tbMsgSection2
        cllMsgSections.Add .tbMsgSection3
        
        cllMsgLabels.Add .laMsgSection1
        cllMsgLabels.Add .laMsgSection2
        cllMsgLabels.Add .laMsgSection3
        
        cllReplyButtons.Add .cmbReply1
        cllReplyButtons.Add .cmbReply2
        cllReplyButtons.Add .cmbReply3
        cllReplyButtons.Add .cmbReply4
        cllReplyButtons.Add .cmbReply5
    End With
End Sub

Public Property Let MaxFormWidthPercentageOfScreenSize(ByVal l As Long):            lMaximumFormWidthPoW = l:                   End Property
Public Property Get MaxFormWidthPercentageOfScreenSize() As Long:                   MaxFormWidthPercentageOfScreenSize = lMaximumFormWidthPoW: End Property

Public Property Let MaxFormHeightPercentageOfScreenSize(ByVal l As Long):            lMaximumFormHeightPoW = l:                   End Property
Public Property Get MaxFormHeightPercentageOfScreenSize() As Long:                   MaxFormHeightPercentageOfScreenSize = lMaximumFormHeightPoW: End Property

Public Property Get MaximumFormWidth() As Single:                                   MaximumFormWidth = siMaximumFormWidth:      End Property
Public Property Let MaximumFormWidth(ByVal si As Single)
    '~~ The maximum specified form size is limited to 99% of the screen size!
    siMaximumFormWidth = wVirtualScreenWidth * (Min(si, 99) / 100)   ' maximum form height based on screen size
    '~~ The maximum form width must never not become less than the minimum width
    If siMaximumFormWidth < siMinimumFormWidth Then
       siMaximumFormWidth = siMinimumFormWidth
    End If
End Property

Public Property Get MaximumFormHeight() As Single:                                  MaximumFormHeight = siMaximumFormHeight:    End Property
Public Property Let MaximumFormHeight(ByVal si As Single)
    '~~ The maximum form height is limited to 99% of the screen size!
    siMaximumFormHeight = wVirtualScreenHeight * (Min(si, 99) / 100)
    '~~ The maximum form width must never not become less than the minimum width
    If siMaximumFormHeight < siMinimumFormHeight Then
       siMaximumFormHeight = siMinimumFormHeight
    End If
End Property

Public Property Get MinimumFormWidth() As Single:                               MinimumFormWidth = siMinimumFormWidth:          End Property
Public Property Let MinimumFormWidth(ByVal si As Single)
    siMinimumFormWidth = si
    '~~ The maximum form width must never not become less than the minimum width
    If siMaximumFormWidth < siMinimumFormWidth Then
       siMaximumFormWidth = siMinimumFormWidth
    End If
End Property

Private Property Get ReplyButton(Optional i As Long) As MsForms.CommandButton:  Set ReplyButton = cllReplyButtons(i):           End Property
Private Property Get ReplyButtonValue(Optional i As Long):                      ReplyButtonValue = cllReplyButtonsValue(i):     End Property
Private Property Let ReplyButtonValue(Optional i As Long, ByVal v As Variant):  cllReplyButtonsValue.Add v:                     End Property

Private Property Get MsgSectionLabel(Optional i As Long) As MsForms.Label:      Set MsgSectionLabel = cllMsgLabels(i):          End Property
Private Property Get MsgSection(Optional i As Long) As MsForms.TextBox:         Set MsgSection = cllMsgSections(i):             End Property

Public Property Let ErrSrc(ByVal s As String):                                  sErrSrc = s:                                    End Property

Public Property Let MsgSection1Label(ByVal s As String):                        sMsgSection1Label = s:                          End Property
Public Property Let MsgSection1Text(ByVal s As String):                         sMsgSection1Text = s:                           End Property
Public Property Let MsgSection1Monospaced(ByVal b As Boolean):                  bMsgSection1Monospaced = b:                     End Property

Public Property Let MsgSection2Label(ByVal s As String):                        sMsgSection2Label = s:                          End Property
Public Property Let MsgSection2Text(ByVal s As String):                         sMsgSection2Text = s:                           End Property
Public Property Let MsgSection2Monospaced(ByVal b As Boolean):                  bMsgSection2Monospaced = b:                     End Property

Public Property Let MsgSection3Label(ByVal s As String):                        sMsgSection3Label = s:                          End Property
Public Property Let MsgSection3Text(ByVal s As String):                         sMsgSection3Text = s:                           End Property
Public Property Let MsgSection3Monospaced(ByVal b As Boolean):                  bMsgSection3Monospaced = b:                     End Property

Public Property Let Replies(ByVal v As Variant):                                vReplies = v:                                   End Property

Public Property Let Title(ByVal s As String):                                   sTitle = s:                                     End Property

' Set the top position for the control (ctl) and return the top posisition for the next one
Private Property Get TopNext(ByVal ctl As MsForms.Control) As Single

    TopNext = siTopNext
    With ctl
        .Top = siTopNext    ' the top position for this one
        '~~ Calculate the top position for any displayed which may come next
        Select Case TypeName(ctl)
            Case "TextBox", "CommandButton":    siTopNext = .Top + .Height + MARGIN_VERTICAL
            Case "Label"
                Select Case ctl.Name
                    Case "la":                  siTopNext = Me.laMsgTitleSpaceBottom.Top + Me.laMsgTitleSpaceBottom.Height + MARGIN_VERTICAL
                    Case Else:                  siTopNext = .Top + .Height
                End Select
        End Select
    End With

End Property

Private Sub cmbReply1_Click():  ReplyClicked 1:    End Sub

Private Sub cmbReply2_Click():  ReplyClicked 2:    End Sub

Private Sub cmbReply3_Click():  ReplyClicked 3:    End Sub

Private Sub cmbReply4_Click():  ReplyClicked 4:    End Sub

Private Sub cmbReply5_Click():  ReplyClicked 5:    End Sub

Private Sub ControlsTopPos()

    siTopNext = MARGIN_FORM_TOP   ' initial top position of first visible element
    
    With Me
        ControlPosTop MsgSectionLabel(1), MARGIN_VERTIVAL_LABEL
        ControlPosTop MsgSection(1), MARGIN_VERTICAL
        ControlPosTop MsgSectionLabel(2), MARGIN_VERTIVAL_LABEL
        ControlPosTop MsgSection(2), MARGIN_VERTICAL
        ControlPosTop MsgSectionLabel(3), MARGIN_VERTIVAL_LABEL
        ControlPosTop MsgSection(3), MARGIN_VERTICAL
        siTopNext = siTopNext + MARGIN_VERTICAL
        
        RepliesPosTop
        .Height = ReplyButton(1).Top + ReplyButton(1).Height + MARGIN_FORM_BOTTOM
    End With

End Sub

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
            siWidth = .Width
            s = .Value
            .SetFocus
            .AutoSize = False
            .Value = vbNullString
            Select Case .ScrollBars
                Case fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsVertical
                    .Width = siWidth + 15
                    .Height = .Height - siHeightExceeding - 15
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsVertical
                Case fmScrollBarsBoth
                    .Height = .Height - siHeightExceeding - 15
                    .Width = siWidth - 15
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsVertical
                    .Width = siWidth + 15
                    .Height = .Height - siHeightExceeding
            End Select
            .Value = s
            .SelStart = 0
        End With
    End With
    
End Sub

' Returns the visible textbox with the largest height.
' ----------------------------------------------------------
Private Function MsgSectionMaxHeight() As MsForms.TextBox
Dim v   As Variant
Dim si  As Single
Dim tb  As MsForms.TextBox

    For Each v In MsgSectionsVisible
        Set tb = v
        If tb.Height > si Then
            si = tb.Height
            Set MsgSectionMaxHeight = tb
        End If
    Next v
    
End Function

' Setup a message section with its label when one is specified.
' -------------------------------------------------------------
Private Sub MsgSectionSetup( _
            ByVal section As Long, _
            ByVal latext As String, _
            ByVal tbtext As String, _
            ByVal monospaced As Boolean)
    
    Dim la  As MsForms.Label
    Dim tb  As MsForms.TextBox
    
    If tbtext <> vbNullString Then
        '~~ Setup above text label/title only when there is a text
        If latext <> vbNullString Then
            Set la = MsgSectionLabel(section)
            With la
                .Caption = latext
                .Visible = True
                .Left = FORM_MARGIN_LEFT
            End With
        End If
        
        Set tb = MsgSection(section)
        With tb
            If monospaced Then
                .Font.Name = sFontNameMonospaced
                MsgSectionMonospacedWidthSet tb, tbtext  ' sets the global siMaxMonospacedTextWidth variable
                .Width = siMaxMonospacedTextWidth
            Else
                .Width = Me.Width - (FORM_MARGIN_LEFT + FORM_MARGIN_RIGHT)
            End If
            .Visible = True
            .MultiLine = True
            .WordWrap = True
            .Left = FORM_MARGIN_LEFT
            If monospaced Then
                .IntegralHeight = True
                DoEvents
                .Width = siMaxMonospacedTextWidth
            End If
            .AutoSize = True
            .Value = Replace(tbtext, vbLf, vbCrLf)
            .SelStart = 0
            siMaxMsgWidth = Max(siMaxMsgWidth, .Width)
        End With
        
        AdjustFormWidth tb
        
    End If
End Sub

Private Sub AdjustFormWidth(ByVal ctl As MsForms.Control)
    With Me
        Debug.Print "Width = " & .Width
        .Width = mMsg.Max(.Width, _
                          siMinimumFormWidth, _
                         ctl.Left + ctl.Width + FORM_MARGIN_RIGHT _
                         )
        Debug.Print "Width = " & .Width
    End With

End Sub
' Setup the width of a the monospaced textbox (tb) with text (sText)
' whereby the fixed font textbox's width is determined by the longest
' text line's length - determined by means of an autosized width-template.
' ------------------------------------------------------------------------
Private Sub MsgSectionMonospacedWidthSet( _
            ByVal tb As MsForms.TextBox, _
            ByVal sText As String)
            
    Dim sSplit          As String
    Dim v               As Variant
    Dim siMaximumFormWidth As Single
    
    '~~ Determine the used line break character
    If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
    If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
    
    '~~ Find the width which fits the largest text line
    With tb
        .MultiLine = False
        .WordWrap = False
        For Each v In Split(sText, sSplit)
            .Value = v
            siMaximumFormWidth = Max(siMaximumFormWidth, .Width)
        Next v
        .Width = Max(siMaximumFormWidth, Me.laMsgTitle.Width) + FORM_MARGIN_RIGHT
        siMaxMonospacedTextWidth = mMsg.Max(siMaxMonospacedTextWidth, .Width)
    End With
End Sub

Private Function MsgSectionIsMonospaced(ByVal tb As MsForms.TextBox) As Boolean
    MsgSectionIsMonospaced = tb.Font.Name = sFontNameMonospaced
End Function

' Returns a collection of the visible message sections.
' -----------------------------------------------------
Private Function MsgSectionsVisible() As Collection
    
    Dim v   As Variant
    Dim tb  As MsForms.TextBox
    Dim cll As New Collection
    
    For Each v In cllMsgSections
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
Private Sub MsgSectionsHeightFinal()
Dim siHeightCurrentRequired As Single
Dim siHeightExceeding       As Single
Dim cllMsgSections          As Collection
Dim s                       As String

    With Me
        siHeightCurrentRequired = .cmbReply1.Top + .cmbReply1.Height + MARGIN_FORM_BOTTOM
    End With
    If siHeightCurrentRequired <= Me.Height Then Exit Sub
    
    Set cllMsgSections = MsgSectionsVisible
    siHeightExceeding = siHeightCurrentRequired > Me.Height
    '~~ All displayed controls together take more height than the available form's height
    '~~ The displayed message sections are reduced in their height to fit the available space
    With MsgSectionMaxHeight ' The message paragraph with the maximum height
        .SetFocus
        s = .Value
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
            Case fmScrollBarsVertical:  .ScrollBars = fmScrollBarsBoth
            Case fmScrollBarsNone:      .ScrollBars = fmScrollBarsVertical
            Case fmScrollBarsBoth       ' nothing required
        End Select
        .Height = .Height - siHeightExceeding - 15 ' 15 is the height required by the scroll bar
    End With

End Sub

Private Sub MsgSectionsSetup()
    With Me
        If sMsgSection1Text <> vbNullString Then MsgSectionSetup section:=1, latext:=sMsgSection1Label, tbtext:=sMsgSection1Text, monospaced:=bMsgSection1Monospaced
        If sMsgSection2Text <> vbNullString Then MsgSectionSetup section:=2, latext:=sMsgSection2Label, tbtext:=sMsgSection2Text, monospaced:=bMsgSection2Monospaced
        If sMsgSection3Text <> vbNullString Then MsgSectionSetup section:=3, latext:=sMsgSection3Label, tbtext:=sMsgSection3Text, monospaced:=bMsgSection3Monospaced
    End With
End Sub

' After final adjustment of the form's width the visible the message paragraph's width is re-adjusted.
' Proportional message sections will result in a new height,
' mmonospaced message sections will receive a horizontal sccroll bar.
' ----------------------------------------------------------------------------------------------------
Private Sub MsgSectionsWidthFinal()

    Dim siMax   As Single
    Dim v       As Variant
    Dim tb      As MsForms.TextBox
    Dim s       As String
 
    siMax = Me.Width - (FORM_MARGIN_LEFT + FORM_MARGIN_RIGHT)
    
    For Each v In MsgSectionsVisible
        Set tb = v
        With tb
            s = .Value
            If .Width > siMax And MsgSectionIsMonospaced(tb) Then
                '~~ Provide a horizontal scroll-bar
                .WordWrap = False
                .AutoSize = False
                .Value = vbNullString
                .SetFocus
                .Value = s
                Select Case .ScrollBars
                    Case fmScrollBarsVertical:                      .ScrollBars = fmScrollBarsBoth
                    Case fmScrollBarsNone:                          .ScrollBars = fmScrollBarsHorizontal
                    Case fmScrollBarsHorizontal, fmScrollBarsBoth   ' do nothing
                End Select
                .Width = Me.Width - FORM_MARGIN_LEFT - FORM_MARGIN_RIGHT / 2 - 15
                DoEvents
                .SelStart = 0
            Else
                '~~ Adjust the textbox width
                .WordWrap = True
                .AutoSize = True
                .Value = vbNullString
                .Width = siMax
                DoEvents
                .Value = s
            End If
        End With
    Next v
    
End Sub

' Setup for each reply button its left position.
' ----------------------------------------------
Private Sub RepliesPosLeft()
    
    Dim v           As Variant
    Dim lVisible    As Long
    Dim siWidth     As Single
    Dim siHeight    As Single
    Dim siByReplies As Single   ' Width required for the visible reply buttons
    
    For Each v In cllReplyButtons
        If v.Visible = False Then Exit For
        lVisible = lVisible + 1
        With ReplyButton(lVisible)
            siWidth = Max(siWidth, .Width, REPLY_BUTTON_MIN_WIDTH)
            siHeight = Max(siHeight, .Height)
        End With
    Next v
    
    siByReplies = FORM_MARGIN_LEFT + ((siWidth + MARGIN_HORIZONTAL) * lVisible)
    Me.Width = Max(Me.Width, Me.MinimumFormWidth, siByReplies)
    
    For Each v In cllReplyButtons
        v.Width = siWidth
        v.Height = siHeight
    Next v
    
    Select Case lVisible
        Case 1
            ReplyButton(1).Left = (Me.Width / 2) - (siWidth / 2) ' center
        Case 2
            ReplyButton(1).Left = (Me.Width / 2) - (MARGIN_HORIZONTAL / 2) - siWidth    ' left from center
            ReplyButton(2).Left = ReplyButton(1).Left + siWidth + MARGIN_HORIZONTAL     ' right from center
        Case 3
            ReplyButton(2).Left = (Me.Width / 2) - (siWidth / 2)                        ' center button 2
            ReplyButton(1).Left = ReplyButton(2).Left - siWidth - MARGIN_HORIZONTAL     ' pos button 1 left from 2
            ReplyButton(3).Left = ReplyButton(2).Left + siWidth + MARGIN_HORIZONTAL     ' pos button 3 right from 2
        Case 4
            ReplyButton(2).Left = (Me.Width / 2) - (MARGIN_HORIZONTAL / 2) - siWidth    ' pos button 2 left from center
            ReplyButton(3).Left = (Me.Width / 2) + (MARGIN_HORIZONTAL / 2)              ' pos button 3 right from center
            ReplyButton(1).Left = ReplyButton(2).Left - MARGIN_HORIZONTAL - siWidth     ' pos button 1 left from 2
            ReplyButton(4).Left = ReplyButton(3).Left + siWidth + MARGIN_HORIZONTAL     ' pos button 4 right from 3
        Case 5
            ReplyButton(3).Left = (Me.Width / 2) - (siWidth / 2)                        ' center button 3
            ReplyButton(2).Left = ReplyButton(3).Left - siWidth - MARGIN_HORIZONTAL     ' pos button 2 left from 3
            ReplyButton(1).Left = ReplyButton(2).Left - siWidth - MARGIN_HORIZONTAL     ' pos button 1 left from 2
            ReplyButton(4).Left = ReplyButton(3).Left + siWidth + MARGIN_HORIZONTAL     ' pos button 4 right from 3
            ReplyButton(5).Left = ReplyButton(4).Left + siWidth + MARGIN_HORIZONTAL     ' pos button 5 right from 4
    End Select

End Sub

Private Sub RepliesPosTop()
Dim siTop   As Single

    With Me
        With .cmbReply1
            .Top = TopNext(Me.cmbReply1)
            siTop = .Top
            .Height = siMaxReplyHeight
        End With
        With .cmbReply2
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply3
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply4
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        With .cmbReply5
            .Top = siTop
            .Height = siMaxReplyHeight
        End With
        .Height = siTop + siMaxReplyHeight + (MARGIN_VERTICAL * 5)
    End With
    
End Sub

' Setup and position the displayed reply buttons.
' Return the max reply button width.
' ------------------------------------------------------
Private Sub ReplyButtonsSetup(ByVal vReplies As Variant)
    
    Dim v   As Variant
    Dim lButton As Long

    Set cllReplyButtonsValue = Nothing
    Set cllReplyButtonsValue = New Collection

    With Me
        '~~ Setup button caption
        If IsNumeric(vReplies) Then
            Select Case vReplies
                Case vbOKOnly
                    ReplyButtonSetup 1, "Ok", vbOK
                Case vbOKCancel
                    ReplyButtonSetup 1, "Ok", vbOK
                    ReplyButtonSetup 2, "Cancel", vbCancel
                Case vbYesNo
                    ReplyButtonSetup 1, "Yes", vbYes
                    ReplyButtonSetup 2, "No", vbNo
                Case vbRetryCancel
                    ReplyButtonSetup 1, "Retry", vbRetry
                    ReplyButtonSetup 2, "Cancel", vbCancel
                Case vbYesNoCancel
                    ReplyButtonSetup 1, "Yes", vbYes
                    ReplyButtonSetup 2, "No", vbNo
                    ReplyButtonSetup 3, "Cancel", vbCancel
                Case vbAbortRetryIgnore
                    ReplyButtonSetup 1, "Abort", vbAbort
                    ReplyButtonSetup 2, "Retry", vbRetry
                    ReplyButtonSetup 3, "Ignore", vbIgnore
            End Select
        Else
            aReplyButtons = Split(vReplies, ",")
            For Each v In aReplyButtons
                If v <> vbNullString Then
                    lButton = lButton + 1
                    ReplyButtonSetup lButton, v, v
                End If
            Next v
        End If
    End With

End Sub

' Return the value of the clicked reply button (button).
' ------------------------------------------------------
Private Sub ReplyClicked(ByVal button As Long)
Dim s As String
    
    s = ReplyButtonValue(button)
    If IsNumeric(s) Then
        mMsg.MsgReply = CLng(s)
    Else
        mMsg.MsgReply = s
    End If
#If Test = 0 Then ' allows assertions during testing
    Unload Me
#Else
    Me.Hide
#End If
    
End Sub

' Setup a reply button's visibility and caption.
' -----------------------------------------------
Private Sub ReplyButtonSetup(ByVal button As Long, _
                             ByVal s As String, _
                             ByVal v As Variant)
    
    Dim cmb As MsForms.CommandButton
    
    If s <> vbNullString Then
        Set cmb = ReplyButton(button)
        With cmb
            .Visible = True
            .Caption = s
            siMaxReplyHeight = mMsg.Max(siMaxReplyHeight, .Height)
            siMaxReplyWidth = Max(siMaxReplyWidth, .Width, REPLY_BUTTON_MIN_WIDTH)
        End With
        ReplyButtonValue(button) = v
    End If
    
End Sub

' An extra title label mimics the title bar in order to determine the required form's width.
' When a specific font name and/or size is specified, the extra title label is actively used
' and the UserForm's title bar is not displayed - which means that there is no X to cancel.
' ------------------------------------------------------------------------------------------
Private Sub TitleSetup()
    
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
                siTitleWidth = .Width + MARGIN_HORIZONTAL
            End With
            siTopNext = MARGIN_FORM_TOP
            .Caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
        
        With .laMsgTitle
            '~~ The title label is used to adjust the form width
            .WordWrap = False
            .AutoSize = True
            .Caption = " " & sTitle    ' some left margin
            .AutoSize = False
            siTitleWidth = .Width + MARGIN_HORIZONTAL
        End With
        
        .laMsgTitleSpaceBottom.Width = .laMsgTitle.Width
    
        AdjustFormWidth .laMsgTitleSpaceBottom
    
    End With

End Sub

Public Sub FormFinalPositionOnScreen()
    AdjustStartupPosition Me
End Sub

Private Sub ControlPosTop(ByVal ctl As MsForms.Control, _
                          ByVal siMargin As Single)
    With ctl
        If .Visible Then
            .Top = siTopNext
            siTopNext = .Top + .Height + siMargin
        End If
    End With
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
                .Left = pOwner.Left + ((pOwner.Width - .Width) / 2)
                .Top = pOwner.Top + ((pOwner.Height - .Height) / 2)
            End With
            
        Case CenterScreen           ' Assign the Left and Top properties after switching to Manual positioning.
            With pUserForm
                .StartupPosition = Manual
                .Left = (wVirtualScreenWidth - .Width) / 2
                .Top = (wVirtualScreenHeight - .Height) / 2
            End With
    End Select
 
    ' Avoid falling off screen. Misplacement can be caused by multiple screens when the primary display
    ' is not the left-most screen (which causes "pOwner.Left" to be negative). First make sure the bottom
    ' right fits, then check if the top-left is still on the screen (which gets priority).
    '
    With pUserForm
        If ((.Left + .Width) > (wVirtualScreenLeft + wVirtualScreenWidth)) _
        Then .Left = ((wVirtualScreenLeft + wVirtualScreenWidth) - .Width)
        If ((.Top + .Height) > (wVirtualScreenTop + wVirtualScreenHeight)) _
        Then .Top = ((wVirtualScreenTop + wVirtualScreenHeight) - .Height)
        If (.Left < wVirtualScreenLeft) Then .Left = wVirtualScreenLeft
        If (.Top < wVirtualScreenTop) Then .Top = wVirtualScreenTop
    End With
End Sub
 
' Returns pixels (device dependent) to points (used by Excel).
' --------------------------------------------------------------------
Private Sub ConvertPixelsToPoints(ByRef x As Single, ByRef y As Single)
On Error Resume Next
    Dim hDC            As Long
    Dim RetVal         As Long
    Dim PixelsPerInchX As Long
    Dim PixelsPerInchY As Long
 
    hDC = GetDC(0)
    PixelsPerInchX = GetDeviceCaps(hDC, LOGPIXELSX)
    PixelsPerInchY = GetDeviceCaps(hDC, LOGPIXELSY)
    RetVal = ReleaseDC(0, hDC)
    x = x * TWIPSPERINCH / 20 / PixelsPerInchX
    y = y * TWIPSPERINCH / 20 / PixelsPerInchY
End Sub

Private Sub UserForm_Activate()
    
    With Me
        .Width = siMinimumFormWidth
        TitleSetup
        MsgSectionsSetup                ' Message sections text and visibility
        ReplyButtonsSetup vReplies      ' Reply buttons text, size and visibility
        
        If .Width > siMaximumFormWidth Then .Width = siMaximumFormWidth
        DoEvents
        
        '~~ An ajusted form width or any message section may have changed the available width
        '~~ The final width adjustment may enlarge the width or shrink it wereas in the latter
        '~~ case for a monospaced section a horizontal scroll bar is provided.
        
        .Width = Max(siMinimumFormWidth, _
                     FORM_MARGIN_LEFT + siMaxMsgWidth + FORM_MARGIN_RIGHT, _
                     FORM_MARGIN_LEFT + ((siMaxReplyWidth + MARGIN_HORIZONTAL) * cllReplyButtonsVisible.Count)) + FORM_MARGIN_RIGHT
        
        RepliesPosLeft                  ' adjust displayed reply buttons left position
        
        MsgSectionsWidthFinal
        
        If .Height > siMaximumFormHeight Then
            '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
            FormHeightFinal
            DoEvents
            MsgSectionsHeightFinal  ' may end up with a horizontal scroll bar for a monospaced message section
        End If
        
        ControlsTopPos          ' adjusts all controls' top position
    
    End With
    
    AdjustStartupPosition Me

End Sub


