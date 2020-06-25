VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   1  'Fenstermitte
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
'                 from the VB MsgBox or any test string.
'
' W. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const FORM_WIDTH_MIN            As Single = 200
Const FORM_WIDTH_MAX            As Long = 80    ' Maximum message form width (percentage of the maximized application window)
Const FORM_HEIGHT_MAX           As Long = 80    ' Maximum message form height (percentage of the maximized application window)
Const FORM_MARGIN_LEFT          As Single = 5
Const FORM_MARGIN_RIGHT         As Single = 10
Const MARGIN_HORIZONTAL         As Single = 10
Const MARGIN_VERTICAL           As Single = 10
Const MARGIN_FORM_TOP           As Single = 5   ' Top position of the first visible control
Const MARGIN_FORM_BOTTOM        As Single = 50
Const MARGIN_VERTIVAL_LABEL       As Single = 5
Const REPLY_BUTTON_MIN_WIDTH    As Single = 70

Dim sTitle                      As String
Dim sErrSrc                     As String
Dim vReplies                    As Variant
Dim aReplyButtons               As Variant
Dim sReplyButtonsReturnValue    As String   ' The provided reply buttons return values a comma delimited string
Dim lNoOfReplyButtons           As Long
Dim siFormWidth                 As Single
Dim sTitleFontName              As String
Dim sTitleFontSize              As String   ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim sMsg1Proportional           As String
Dim sMsg2Proportional           As String
Dim sMsg3Proportional           As String
Dim sMsg1Monospaced             As String
Dim sMsg2Monospaced             As String
Dim sMsg3Monospaced             As String
Dim sLabelMessage1              As String
Dim sLabelMessage2              As String
Dim sLabelMessage3              As String
Dim siTitleWidth                As Single
Dim siMaxMonospacedTextWidth    As Single
Dim siMaxReplyWidth             As Single
Dim siMaxReplyHeight            As Single
Dim vReplyButtons               As Variant
Dim sReplyButtons               As String

Private Sub UserForm_Initialize()
    siFormWidth = FORM_WIDTH_MIN ' Default
End Sub

Public Property Let ErrSrc(ByVal s As String):                  sErrSrc = s:                                    End Property

Public Property Let FormWidth(ByVal si As Single):              siFormWidth = si:                               End Property

Public Property Let LabelMessage1(ByVal s As String):           sLabelMessage1 = s:                             End Property

Public Property Let LabelMessage2(ByVal s As String):           sLabelMessage2 = s:                             End Property

Public Property Let LabelMessage3(ByVal s As String):           sLabelMessage3 = s:                             End Property

Private Property Get LabelMsg1() As MSForms.Label:              Set LabelMsg1 = Me.laMsg1:                      End Property

Private Property Get LabelMsg2() As MSForms.Label:              Set LabelMsg2 = Me.laMsg2:                      End Property

Private Property Get LabelMsg3() As MSForms.Label:              Set LabelMsg3 = Me.laMsg3:                      End Property

Public Property Let Message1Monospaced(ByVal s As String):      sMsg1Monospaced = s:                            End Property

Public Property Let Message1Proportional(ByVal s As String):    sMsg1Proportional = s:                          End Property

Public Property Let Message2Monospaced(ByVal s As String):      sMsg2Monospaced = s:                            End Property

Public Property Let Message2Proportional(ByVal s As String):    sMsg2Proportional = s:                          End Property

Public Property Let Message3Monospaced(ByVal s As String):      sMsg3Monospaced = s:                            End Property

Public Property Let Message3Proportional(ByVal s As String):    sMsg3Proportional = s:                          End Property

Private Property Get msg1monospaced() As MSForms.TextBox:       Set msg1monospaced = Me.tbMsg1Monospaced:       End Property

Private Property Get msg2monospaced() As MSForms.TextBox:       Set msg2monospaced = Me.tbMsg2Monospaced:       End Property

Private Property Get msg3monospaced() As MSForms.TextBox:       Set msg3monospaced = Me.tbMsg3Monospaced:       End Property

Public Property Let replies(ByVal v As Variant):                vReplies = v:                                   End Property

Public Property Let title(ByVal s As String):                   sTitle = s:                                     End Property

Public Property Let titleFontName(ByVal s As String):           sTitleFontName = s:                             End Property

Public Property Let titlefontsize(ByVal l As Long):             sTitleFontSize = l:                             End Property

Private Property Get TopNext(Optional ByVal ctl As Variant = Nothing) As Single
Dim tb  As MSForms.TextBox
Dim la  As MSForms.Label
Dim cb  As MSForms.CommandButton

    TopNext = siTopNext ' Return the current position for control

    If Not ctl Is Nothing Then
        With ctl
                ' Set the top position  for this control and increase it for the next
                .Top = siTopNext
                Select Case TypeName(ctl)
                    Case "TextBox", "CommandButton"
                        siTopNext = .Top + .Height + MARGIN_VERTICAL
                    Case "Label"
                        Select Case ctl.Name
                            Case "la"
                                siTopNext = Me.laTitleSpaceBottom.Top + Me.laTitleSpaceBottom.Height + MARGIN_VERTICAL
                            Case Else ' Message label
                                siTopNext = .Top + .Height
                        End Select
                End Select
        End With
    End If
End Property

Private Sub cmbReply1_Click():  ReplyClicked 0:    End Sub

Private Sub cmbReply2_Click():  ReplyClicked 1:    End Sub

Private Sub cmbReply3_Click():  ReplyClicked 2:    End Sub

Private Sub cmbReply4_Click():  ReplyClicked 3:    End Sub

Private Sub cmbReply5_Click():  ReplyClicked 4:    End Sub

Private Function ControlsNonTextBoxHeight() As Single
' -------------------------------------------------------
' Calculates the required height for all displayed
' non-TextBox controls.
' -------------------------------------------------------
Dim si  As Single
Dim ctl As MSForms.Control
Dim tb  As MSForms.TextBox
    
    With Me
        For Each ctl In Me.Controls
            If TypeName(ctl) <> "Textbox" Then
                With ctl
                    If .Visible Then
                        
                        si = si + .Height + MARGIN_VERTICAL
                    End If
                End With
            End If
        Next ctl
    End With
    ControlsNonTextBoxHeight = si
    
End Function

Private Sub ControlsTopPos()
Dim ctl As MSForms.Control

    siTopNext = MARGIN_FORM_TOP   ' initial top position of first visible element
    
    With Me
        TopPos .laMsg1, MARGIN_VERTIVAL_LABEL
        TopPos .tbMsg1Monospaced, MARGIN_VERTICAL
        TopPos .tbMsg1Proportional, MARGIN_VERTICAL
        TopPos .laMsg2, MARGIN_VERTIVAL_LABEL
        TopPos .tbMsg2Monospaced, MARGIN_VERTICAL
        TopPos .tbMsg2Proportional, MARGIN_VERTICAL
        TopPos .laMsg3, MARGIN_VERTIVAL_LABEL
        TopPos .tbMsg3Monospaced, MARGIN_VERTICAL
        TopPos .tbMsg3Proportional, MARGIN_VERTICAL
        siTopNext = siTopNext + MARGIN_VERTICAL
        
        RepliesPosTop
        .Height = .cmbReply1.Top + .cmbReply1.Height + MARGIN_FORM_BOTTOM
    End With

End Sub

' Final form height adjustment considering only the maximum height specified
' --------------------------------------------------------------------------
Private Sub FormHeightFinal()
Dim siHeightMax         As Single
Dim siHeightUsed        As Single
Dim siHeightExceeding   As Single
Dim siScreenHeight      As Single
Dim s                   As String
Dim siWidth             As Single

'    Application.WindowState = xlMaximized
'    Application.ScreenUpdating = True
    siScreenHeight = Application.Height
    siHeightMax = siScreenHeight * (FORM_HEIGHT_MAX / 100)
'    Application.WindowState = xlNormal
    
    With Me
        siHeightUsed = .cmbReply1.Top + .cmbReply1.Height + MARGIN_FORM_BOTTOM
        If siHeightUsed > siHeightMax Then
            '~~ Reduce the height of the largest displayed message paragraph by the amount of exceeding height
            siHeightExceeding = siHeightUsed - siHeightMax
            With MsgParagraphMaxHeight
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
            .Top = (siScreenHeight - siHeightMax) / 2
        End If
    End With
    
End Sub

' Final form width adjustment considering: title width, maximum fixed message text width,
' width and number of displayed reply buttons, specified minimum message window width
' ---------------------------------------------------------------------------------------
Private Sub FormWidthFinal()
Dim siMaxWidth  As Single

    Application.WindowState = xlMaximized
    siMaxWidth = Application.Width * (FORM_WIDTH_MAX / 100)
    Application.WindowState = xlNormal
    
    With Me
        If .Width > siMaxWidth Then
            .Width = siMaxWidth
        Else
            .Width = Max( _
                         siTitleWidth, _
                         ((siMaxReplyWidth + MARGIN_VERTICAL) * lNoOfReplyButtons) + (MARGIN_VERTICAL * 2), _
                         siMaxMonospacedTextWidth, _
                         FORM_WIDTH_MIN)

        End If
    End With
    
End Sub

' Return the displayed textbox with the largest height
' ----------------------------------------------------------
Private Function MsgParagraphMaxHeight() As MSForms.TextBox
Dim v   As Variant
Dim si  As Single
Dim tb  As MSForms.TextBox

    For Each v In MsgParagraphsDisplayed
        Set tb = v
        If tb.Height > si Then Set MsgParagraphMaxHeight = tb
    Next v
    
End Function

Private Sub MsgParagraphMonospacedSetup( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
' ----------------------------------------
' Setup any fixed font message and its
' above label when one is specified.
' ----------------------------------------

    If sTextBoxText <> vbNullString Then
        '~~ Setup above text label/title only when there is a text
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
                .Left = FORM_MARGIN_LEFT
            End With
        End If
        
        With tb
            .Visible = True
            MsgParagraphMonospacedWidthSet tb, sTextBoxText  ' sets the global siMaxMonospacedTextWidth variable
            .MultiLine = True
            .WordWrap = True
            .AutoSize = True
            .Value = sTextBoxText
            .Left = FORM_MARGIN_LEFT
        End With
        
        With Me
            .Width = mMsg.Max(FORM_WIDTH_MIN, _
                                 siFormWidth, _
                                 .laTitle.Width, _
                                 tb.Left + tb.Width + MARGIN_HORIZONTAL)
            .laTitle.Width = .Width
            .laTitleSpaceBottom.Width = .Width
            .Left = FORM_MARGIN_LEFT
        End With
        
    End If
End Sub

' Setup the width of a the monospaced textbox (tb) with text (sText)
' whereby the fixed font textbox's width is determined by the longest
' text line's length - determined by means of an autosized width-template.
' ------------------------------------------------------------------------
Private Sub MsgParagraphMonospacedWidthSet( _
            ByVal tb As MSForms.TextBox, _
            ByVal sText As String)
    Dim sSplit      As String
    Dim v           As Variant
    Dim siMaxWidth  As Single

    '~~ Determine the used line break character
    If InStr(sText, vbLf) <> 0 Then sSplit = vbLf
    If InStr(sText, vbCrLf) <> 0 Then sSplit = vbCrLf
    
    '~~ Find the width which fits the largest text line
    With Me
        With .tbMsgMonospacedWidthTemplate
            .MultiLine = False
            .WordWrap = False
            For Each v In Split(sText, sSplit)
                .Value = v
                siMaxWidth = Max(siMaxWidth, .Width)
            Next v
        End With
        tb.Width = Max(siMaxWidth, Me.laTitle.Width) + MARGIN_HORIZONTAL
    End With
    siMaxMonospacedTextWidth = mMsg.Max(siMaxMonospacedTextWidth, tb.Width)

End Sub

' Adjust the non-monospaced message paragraph's (tb) width to the form's width
' ----------------------------------------------------------------------------
Private Sub MsgParagraphProportionalSetup( _
            ByVal la As MSForms.Label, _
            ByVal sLabelText As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal sTextBoxText As String)
    
    If sTextBoxText <> vbNullString Then
        '~~ Setup Message Label
        If sLabelText <> vbNullString Then
            With la
                .Caption = sLabelText
                .Visible = True
                .Width = Me.Width - (MARGIN_HORIZONTAL * 2)
                .Left = FORM_MARGIN_LEFT
            End With
        End If
        
        '~~ Setup Message Textbox
        With tb
            .Visible = True
            .MultiLine = True
            .WordWrap = True
            .Width = Me.Width - (MARGIN_HORIZONTAL * 2)
            .AutoSize = True
            .Value = sTextBoxText
            .Left = FORM_MARGIN_LEFT
        End With
    End If

End Sub

' Collection of all displayed message paragraphs.
' -----------------------------------------------------
Private Function MsgParagraphsDisplayed() As Collection
Dim cll As New Collection
Dim ctl As MSForms.Control
    
    With Me
        For Each ctl In Me.Controls
            If TypeName(ctl) = "TextBox" And ctl.Visible = True Then
                cll.Add ctl
            End If
        Next ctl
    End With
    Set MsgParagraphsDisplayed = cll

End Function

' Adjust the largest displayed message paragraph's height
' so that it fits into the final form height.
' -------------------------------------------------------
Private Sub MsgParagraphsHeightFinal()
Dim v                       As Variant
Dim siHeightInitial         As Single
Dim siHeightCurrentRequired As Single
Dim siHeightExceeding       As Single
Dim cllMsgParagraphs        As Collection
Dim tb                      As MSForms.TextBox
Dim s                       As String
Dim siWidth                 As String

    With Me
        siHeightCurrentRequired = .cmbReply1.Top + .cmbReply1.Height + MARGIN_FORM_BOTTOM
    End With
    If siHeightCurrentRequired > Me.Height Then
        Set cllMsgParagraphs = MsgParagraphsDisplayed
        siHeightExceeding = siHeightCurrentRequired > Me.Height
        '~~ All displayed controls together take more height than the available form's height
        '~~ The displayed message paragraphs are reduced in their height to fit the available space
        With MsgParagraphMaxHeight ' The message paragraph with the maximum height
            .SetFocus
            s = .Value
            siWidth = .Width
            Select Case .ScrollBars
                Case fmScrollBarsHorizontal
                    .Height = .Height - siHeightExceeding - 15
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsBoth
                    .Height = .Height - siHeightExceeding - 15
                Case fmScrollBarsBoth
                    .Height = .Height - siHeightExceeding - 15
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsVertical
                    .Height = .Height - siHeightExceeding
            End Select
        End With
    End If

End Sub

Private Function MsgParagraphsHeightInitial() As Single
Dim si  As Single
Dim tb  As MSForms.TextBox
Dim v   As Variant

    For Each v In MsgParagraphsDisplayed
        Set tb = v
        si = si + tb.Height + MARGIN_VERTICAL
    Next v
    MsgParagraphsHeightInitial = si

End Function

Private Sub MsgParagraphsSetup()
    With Me
        If sMsg1Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg1, sLabelMessage1, .tbMsg2Proportional, sMsg1Proportional
        
        If sMsg1Monospaced <> vbNullString _
        Then MsgParagraphMonospacedSetup LabelMsg1, sLabelMessage1, msg1monospaced, sMsg1Monospaced
        
        If sMsg2Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg2, sLabelMessage2, .tbMsg2Proportional, sMsg2Proportional
        
        If sMsg2Monospaced <> vbNullString _
        Then MsgParagraphMonospacedSetup LabelMsg2, sLabelMessage2, msg2monospaced, sMsg2Monospaced
        
        If sMsg3Proportional <> vbNullString _
        Then MsgParagraphProportionalSetup LabelMsg3, sLabelMessage3, .tbMsg3Proportional, sMsg3Proportional
        
        If sMsg3Monospaced <> vbNullString _
        Then MsgParagraphMonospacedSetup LabelMsg3, sLabelMessage3, msg3monospaced, sMsg3Monospaced
    End With
End Sub

' After final adjustment of the form's width all the message paragraph's width
' is re-adjusted. Any message paragraph using a proportinal font will result in
' a new height, any monospaced font paragraph in a vertival scroll bar.
' -----------------------------------------------------------------------------
Private Sub MsgParagraphsWidthFinal()
Dim siMax   As Single
Dim v       As Variant
Dim tb     As MSForms.TextBox
Dim s      As String
 
    siMax = Me.Width - (FORM_MARGIN_LEFT + MARGIN_HORIZONTAL)
    For Each v In MsgParagraphsDisplayed
        Set tb = v
        With tb
            If .Width > siMax Then
                s = .Value
                Select Case .Font.Name
                    Case "Courier New", "Lucida Sans Typewriter"
                        '~~ Manage a vertical scroll-bar instead of a new width
                        .WordWrap = False
                        .AutoSize = False
                        .Value = vbNullString
                        .SetFocus
                        .Value = s
                        Select Case .ScrollBars
                            Case fmScrollBarsVertical
                                .ScrollBars = fmScrollBarsBoth
                                .Width = siMax - 15
                            Case fmScrollBarsHorizontal
                                .Width = siMax
                            Case fmScrollBarsBoth
                                .Width = siMax - 15
                                .Height = .Height - 15
                            Case fmScrollBarsNone
                                .ScrollBars = fmScrollBarsHorizontal
                                .Width = siMax
                        End Select
                        DoEvents
                        .SelStart = 0
                    Case Else
                        '~~ Manage an adjusted width
                        .WordWrap = True
                        .AutoSize = True
                        .Value = vbNullString
                        .Width = siMax
                        DoEvents
                        .Value = s
                End Select
            End If
        End With
    Next v
    
End Sub

' Setup for each reply button its left position.
' ----------------------------------------------
Private Sub RepliesPosLeft()

    With Me
        Select Case lNoOfReplyButtons
            Case 1
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
            Case 2
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (MARGIN_VERTICAL / 2) - siMaxReplyWidth ' left from center
                End With
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply1.Left + siMaxReplyWidth + MARGIN_VERTICAL ' right from center
                End With
            Case 3
                With .cmbReply2
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2) ' center
                End With
                With .cmbReply1
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - MARGIN_VERTICAL ' left from center
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + MARGIN_VERTICAL ' Right from center
                End With
            Case 4
                With .cmbReply2                     ' position 2nd button left from center
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (MARGIN_VERTICAL / 2) - siMaxReplyWidth
                End With
                With .cmbReply1                 ' position left from button 2
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - MARGIN_VERTICAL - siMaxReplyWidth
                End With
                With .cmbReply3
                    .Width = siMaxReplyWidth        ' position 3rd button right from 2nd
                    .Left = Me.cmbReply2.Left + siMaxReplyWidth + MARGIN_VERTICAL
                End With
                With .cmbReply4
                    .Width = siMaxReplyWidth        ' position 4th button right from 3rd
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + MARGIN_VERTICAL
                End With
            Case 5
                With .cmbReply3                                     ' position 3rd reply button in the center
                    .Width = siMaxReplyWidth
                    .Left = (Me.Width / 2) - (siMaxReplyWidth / 2)
                End With
                With .cmbReply2                                     ' position 2nd to left from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left - siMaxReplyWidth - MARGIN_VERTICAL
                End With
                With .cmbReply1                                     ' position 1st to left from 2nd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply2.Left - siMaxReplyWidth - MARGIN_VERTICAL
                End With
                With .cmbReply4                                     ' position 4th right from 3rd
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply3.Left + siMaxReplyWidth + MARGIN_VERTICAL
                End With
                With .cmbReply5                                     ' position 5th right from 4th
                    .Width = siMaxReplyWidth
                    .Left = Me.cmbReply4.Left + siMaxReplyWidth + MARGIN_VERTICAL
                End With
        End Select
    End With

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
Private Sub RepliesSetup(ByVal vReplies As Variant)
Dim v   As Variant

    With Me
        '~~ Setup button caption
        If IsNumeric(vReplies) Then
            Select Case vReplies
                Case vbOKOnly
                    lNoOfReplyButtons = 1
                    ReplySetup .cmbReply1, "Ok"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbOK
                Case vbOKCancel
                    lNoOfReplyButtons = 2
                    ReplySetup .cmbReply1, "Ok"
                    ReplySetup .cmbReply2, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbOK & "," & vbCancel
                Case vbYesNo
                    lNoOfReplyButtons = 2
                    ReplySetup .cmbReply1, "Yes"
                    ReplySetup .cmbReply2, "No"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbYes & "," & vbNo
                Case vbRetryCancel
                    lNoOfReplyButtons = 2
                    ReplySetup .cmbReply1, "Retry"
                    ReplySetup .cmbReply2, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply2.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbRetry & "," & vbCancel
                Case vbYesNoCancel
                    lNoOfReplyButtons = 3
                    ReplySetup .cmbReply1, "Yes"
                    ReplySetup .cmbReply2, "No"
                    ReplySetup .cmbReply3, "Cancel"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply3.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbYes & "," & vbNo & "," & vbCancel
                Case vbAbortRetryIgnore
                    lNoOfReplyButtons = 3
                    ReplySetup .cmbReply1, "Abort"
                    ReplySetup .cmbReply2, "Retry"
                    ReplySetup .cmbReply3, "Ignore"
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply3.Width, REPLY_BUTTON_MIN_WIDTH)
                    sReplyButtonsReturnValue = vbAbort & "," & vbRetry & "," & vbIgnore
            End Select
        Else
            lNoOfReplyButtons = 0
            sReplyButtonsReturnValue = vbNullString
            aReplyButtons = Split(vReplies, ",")
            For Each v In aReplyButtons
                If v <> vbNullString Then
                    lNoOfReplyButtons = lNoOfReplyButtons + 1
                    sReplyButtonsReturnValue = sReplyButtonsReturnValue & v & ","
                End If
            Next v
            Select Case lNoOfReplyButtons
                Case 1
                    ReplySetup .cmbReply1, aReplyButtons(0)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, REPLY_BUTTON_MIN_WIDTH)
                Case 2
                    ReplySetup .cmbReply1, aReplyButtons(0)
                    ReplySetup .cmbReply2, aReplyButtons(1)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, REPLY_BUTTON_MIN_WIDTH)
                Case 3
                    ReplySetup .cmbReply1, aReplyButtons(0)
                    ReplySetup .cmbReply2, aReplyButtons(1)
                    ReplySetup .cmbReply3, aReplyButtons(2)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, REPLY_BUTTON_MIN_WIDTH)
                Case 4
                    ReplySetup .cmbReply1, aReplyButtons(0)
                    ReplySetup .cmbReply2, aReplyButtons(1)
                    ReplySetup .cmbReply3, aReplyButtons(2)
                    ReplySetup .cmbReply4, aReplyButtons(3)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, REPLY_BUTTON_MIN_WIDTH)
                Case 5
                    ReplySetup .cmbReply1, aReplyButtons(0)
                    ReplySetup .cmbReply2, aReplyButtons(1)
                    ReplySetup .cmbReply3, aReplyButtons(2)
                    ReplySetup .cmbReply4, aReplyButtons(3)
                    ReplySetup .cmbReply5, aReplyButtons(4)
                    siMaxReplyWidth = Max(siMaxReplyWidth, .cmbReply1.Width, .cmbReply2.Width, .cmbReply3.Width, .cmbReply4.Width, .cmbReply5.Width, REPLY_BUTTON_MIN_WIDTH)
            End Select
        End If
    End With

End Sub

' Return the value of the clicked reply button (lIndex).
' ------------------------------------------------------
Private Sub ReplyClicked(ByVal lIndex As Long)
Dim s As String
    
    s = Split(sReplyButtonsReturnValue, ",")(lIndex)
    If IsNumeric(s) Then
        mMsg.MsgReply = CLng(s)
    Else
        mMsg.MsgReply = s
    End If
    Unload Me
    
End Sub

' Setup Command Button's visibility and text.
' -----------------------------------------------
Private Sub ReplySetup(ByVal cmb As MSForms.CommandButton, _
                             ByVal s As String)
    If s <> vbNullString Then
        With cmb
            .Visible = True
            .Caption = s
            siMaxReplyHeight = mMsg.Max(siMaxReplyHeight, .Height)
        End With
    End If
End Sub

Private Sub TitleSetup()
' ----------------------------------------------------------------
' When a font name other than the system's font name is provided
' an extra title label mimics the title bar.
' In any case the title label is used to determine the form width
' by autosize of the label.
' ----------------------------------------------------------------
    
    With Me
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            '~~ A title with a specific font is displayed in a dedicated title label
            With .laTitle   ' Hidden by default
                .Top = TopNext(Me.laTitle)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .Visible = True
                siTopNext = .Top + .Height + MARGIN_FORM_TOP
            End With
            
        Else
            .Caption = " " & sTitle
            .laTitleSpaceBottom.Visible = False
            With .laTitle
                '~~ The title label is used to adjust the form width
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.6
                End With
                .Visible = False
                siTitleWidth = .Width + MARGIN_HORIZONTAL
            End With
            siTopNext = MARGIN_FORM_TOP
        End If
        
        With .laTitle
            '~~ The title label is used to adjust the form width
            With .Font
                .Bold = False
                .Size = 8.6
            End With
            .AutoSize = True
            .Caption = " " & sTitle    ' some left margin
            .AutoSize = False
            siTitleWidth = .Width + MARGIN_HORIZONTAL
        End With
        
        .Width = siTitleWidth   ' not the finalwidth though
        .laTitleSpaceBottom.Width = .laTitle.Width
    
    End With

End Sub

Private Sub FormPositionOnScreen()
    With Me
        .Top = 125 '< change 125 to what u want
        .Left = 25 '< change 25 to what u want
    End With
End Sub

Private Sub TopPos(ByVal ctl As MSForms.Control, _
                   ByVal siMargin As Single)
    With ctl
        If .Visible Then
            .Top = siTopNext
            siTopNext = .Top + .Height + siMargin
        End If
    End With
End Sub

Private Sub UserForm_Activate()
    
    With Me
        
        TitleSetup
        
        MsgParagraphsSetup          ' provided message paragraphs text and visibility
        
        RepliesSetup vReplies  ' provided reply buttons text and visibility
        
        FormWidthFinal         ' considers title width, monospaced message paragraphs width and maximum form width
                
        MsgParagraphsWidthFinal
        
        RepliesPosLeft
                
        ControlsTopPos
        
        FormHeightFinal
    
        MsgParagraphsHeightFinal
        
        ControlsTopPos
            
    End With

End Sub

