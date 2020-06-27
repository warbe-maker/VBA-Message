VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   5475
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
'                 from the VB MsgBox or any test string.
'
' lScreenWidth. Rauschenberger Berlin March 2020
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
Dim sReplyButtonsReturnValue    As String   ' The provided reply buttons return values a comma delimited string
Dim lNoOfReplyButtons           As Long
Dim siMinFormWidth              As Single
Dim sTitleFontName              As String
Dim sTitleFontSize              As String   ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim sMsg1TextPrprtional         As String
Dim sMsg2TextPrprtional         As String
Dim sMsg3TextPrprtional         As String
Dim sMsg1TextMonospaced         As String
Dim sMsg2TextMonospaced         As String
Dim sMsg3TextMonospaced         As String
Dim sMsg1Label                  As String
Dim sMsg2Label                  As String
Dim sMsg3Label                  As String
Dim siTitleWidth                As Single
Dim siMaxMonospacedTextWidth    As Single
Dim siMaxReplyWidth             As Single
Dim siMaxReplyHeight            As Single
Dim wVirtualScreenLeft          As Single
Dim wVirtualScreenTop           As Single
Dim wVirtualScreenWidth         As Single
Dim wVirtualScreenHeight        As Single

Private Sub UserForm_Initialize()
    siMinFormWidth = FORM_WIDTH_MIN ' Default
'    lScreenWidth = GetSystemMetrics32(0) ' Screen Resolution width in points
'    lScreenHeight = GetSystemMetrics32(1) ' Screen Resolution height in points
End Sub

'Public Property Get ScreenWidth() As Long:                  ScreenWidth = lScreenWidth:                     End Property
'
'Public Property Get ScreenHeight() As Long:                 ScreenHeight = lScreenHeight:                   End Property

Public Property Let ErrSrc(ByVal s As String):              sErrSrc = s:                                    End Property

Public Property Let FormWidth(ByVal si As Single):          siMinFormWidth = si:                            End Property

Public Property Let Msg1Label(ByVal s As String):           sMsg1Label = s:                                 End Property

Public Property Let Msg2label(ByVal s As String):           sMsg2Label = s:                                 End Property

Public Property Let Msg3label(ByVal s As String):           sMsg3Label = s:                                 End Property

Private Property Get Msg1La() As MSForms.Label:             Set Msg1La = Me.laMsg1:                         End Property

Private Property Get Msg2La() As MSForms.Label:             Set Msg2La = Me.laMsg2:                         End Property

Private Property Get Msg3La() As MSForms.Label:             Set Msg3La = Me.laMsg3:                         End Property

Public Property Let Msg1TextMonospaced(ByVal s As String):  sMsg1TextMonospaced = s:                        End Property

Public Property Let Msg1TextPrprtional(ByVal s As String):  sMsg1TextPrprtional = s:                        End Property

Public Property Let Msg2TextMonospaced(ByVal s As String):  sMsg2TextMonospaced = s:                        End Property

Public Property Let Msg2TextPrprtional(ByVal s As String):  sMsg2TextPrprtional = s:                        End Property

Public Property Let Msg3TextMonospaced(ByVal s As String):  sMsg3TextMonospaced = s:                        End Property

Public Property Let Msg3TextPrprtional(ByVal s As String):  sMsg3TextPrprtional = s:                        End Property

Private Property Get Msg1TbMonospaced() As MSForms.TextBox: Set Msg1TbMonospaced = Me.tbMsg1TextMonospaced: End Property

Private Property Get Msg2TbMonospaced() As MSForms.TextBox: Set Msg2TbMonospaced = Me.tbMsg2TextMonospaced: End Property

Private Property Get Msg3TbMonospaced() As MSForms.TextBox: Set Msg3TbMonospaced = Me.tbMsg3TextMonospaced: End Property

Public Property Let Replies(ByVal v As Variant):            vReplies = v:                                   End Property

Public Property Let Title(ByVal s As String):               sTitle = s:                                     End Property

' Set the top position for the control (ctl) and return the top posisition for the next one
Private Property Get TopNext(ByVal ctl As MSForms.Control) As Single

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

Private Sub cmbReply1_Click():  ReplyClicked 0:    End Sub

Private Sub cmbReply2_Click():  ReplyClicked 1:    End Sub

Private Sub cmbReply3_Click():  ReplyClicked 2:    End Sub

Private Sub cmbReply4_Click():  ReplyClicked 3:    End Sub

Private Sub cmbReply5_Click():  ReplyClicked 4:    End Sub

Private Sub ControlsTopPos()

    siTopNext = MARGIN_FORM_TOP   ' initial top position of first visible element
    
    With Me
        ControlPosTop .laMsg1, MARGIN_VERTIVAL_LABEL
        ControlPosTop .tbMsg1TextMonospaced, MARGIN_VERTICAL
        ControlPosTop .tbMsg1TextPrprtional, MARGIN_VERTICAL
        ControlPosTop .laMsg2, MARGIN_VERTIVAL_LABEL
        ControlPosTop .tbMsg2TextMonospaced, MARGIN_VERTICAL
        ControlPosTop .tbMsg2TextPrprtional, MARGIN_VERTICAL
        ControlPosTop .laMsg3, MARGIN_VERTIVAL_LABEL
        ControlPosTop .tbMsg3TextMonospaced, MARGIN_VERTICAL
        ControlPosTop .tbMsg3TextPrprtional, MARGIN_VERTICAL
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
'Dim siScreenHeight      As Single
Dim s                   As String
Dim siWidth             As Single

'    siScreenHeight = Application.Height
    siHeightMax = wVirtualScreenHeight * (FORM_HEIGHT_MAX / 100)
    
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
        End If
    End With
    
End Sub

' Final form width adjustment considering: title width, maximum fixed message text width,
' width and number of displayed reply buttons, specified minimum message window width
' ---------------------------------------------------------------------------------------
Private Sub FormWidthFinal()
Dim siMaxWidth  As Single

    siMaxWidth = wVirtualScreenWidth * (FORM_WIDTH_MAX / 100)
    If Me.Width > siMaxWidth Then
        Me.Width = siMaxWidth
    End If
    
End Sub

' Return the displayed textbox with the largest height
' ----------------------------------------------------------
Private Function MsgParagraphMaxHeight() As MSForms.TextBox
Dim v   As Variant
Dim si  As Single
Dim tb  As MSForms.TextBox

    For Each v In MsgSectionsDisplayed
        Set tb = v
        If tb.Height > si Then Set MsgParagraphMaxHeight = tb
    Next v
    
End Function

Private Sub MsgParagraphMonospacedSetup( _
            ByVal la As MSForms.Label, _
            ByVal latext As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal tbtext As String)
' ----------------------------------------
' Setup any fixed font message and its
' above label when one is specified.
' ----------------------------------------

    If tbtext <> vbNullString Then
        '~~ Setup above text label/title only when there is a text
        If latext <> vbNullString Then
            With la
                .Caption = latext
                .Visible = True
                .Left = FORM_MARGIN_LEFT
            End With
        End If
        
        With tb
            .Visible = True
            MsgParagraphMonospacedWidthSet tb, tbtext  ' sets the global siMaxMonospacedTextWidth variable
            .MultiLine = True
            .WordWrap = True
            .AutoSize = True
            .Value = tbtext
            .Left = FORM_MARGIN_LEFT
        End With
        
        With Me
            .Width = mMsg.Max(FORM_WIDTH_MIN, _
                                 siMinFormWidth, _
                                 .laMsgTitle.Width, _
                                 tb.Left + tb.Width + MARGIN_HORIZONTAL)
            .laMsgTitle.Width = .Width
            .laMsgTitleSpaceBottom.Width = .Width
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
        tb.Width = Max(siMaxWidth, Me.laMsgTitle.Width) + MARGIN_HORIZONTAL
    End With
    siMaxMonospacedTextWidth = mMsg.Max(siMaxMonospacedTextWidth, tb.Width)

End Sub

' Adjust the non-monospaced message paragraph's (tb) width to the form's width
' ----------------------------------------------------------------------------
Private Sub MsgParagraphPrprtionalSetup( _
            ByVal la As MSForms.Label, _
            ByVal latext As String, _
            ByVal tb As MSForms.TextBox, _
            ByVal tbtext As String)
    
    If tbtext <> vbNullString Then
        '~~ Setup Message Label
        If latext <> vbNullString Then
            With la
                .Caption = latext
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
            .Value = tbtext
            .Left = FORM_MARGIN_LEFT
        End With
    End If

End Sub

' Returns a collection of all displayed message sections.
' -----------------------------------------------------
Private Function MsgSectionsDisplayed() As Collection
Dim cll As New Collection
Dim ctl As MSForms.Control
    
    With Me
        For Each ctl In Me.Controls
            If TypeName(ctl) = "TextBox" And ctl.Visible = True Then
                cll.Add ctl
            End If
        Next ctl
    End With
    Set MsgSectionsDisplayed = cll

End Function

' Adjust the largest displayed message paragraph's height
' so that it fits into the final form height.
' -------------------------------------------------------
Private Sub MsgSectionsHeightFinal()
Dim siHeightCurrentRequired As Single
Dim siHeightExceeding       As Single
Dim cllMsgSections        As Collection
Dim s                       As String

    With Me
        siHeightCurrentRequired = .cmbReply1.Top + .cmbReply1.Height + MARGIN_FORM_BOTTOM
    End With
    If siHeightCurrentRequired <= Me.Height Then Exit Sub
    
    Set cllMsgSections = MsgSectionsDisplayed
    siHeightExceeding = siHeightCurrentRequired > Me.Height
    '~~ All displayed controls together take more height than the available form's height
    '~~ The displayed message sections are reduced in their height to fit the available space
    With MsgParagraphMaxHeight ' The message paragraph with the maximum height
        .SetFocus
        s = .Value
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

End Sub

Private Sub MsgSectionsSetup()
    With Me
        If sMsg1TextPrprtional <> vbNullString Then MsgParagraphPrprtionalSetup la:=.laMsg1, latext:=sMsg1Label, tb:=.tbMsg1TextPrprtional, tbtext:=sMsg1TextPrprtional
        If sMsg1TextMonospaced <> vbNullString Then MsgParagraphMonospacedSetup la:=.laMsg1, latext:=sMsg1Label, tb:=.tbMsg1TextMonospaced, tbtext:=sMsg1TextMonospaced
        If sMsg2TextPrprtional <> vbNullString Then MsgParagraphPrprtionalSetup la:=.laMsg2, latext:=sMsg2Label, tb:=.tbMsg2TextPrprtional, tbtext:=sMsg2TextPrprtional
        If sMsg2TextMonospaced <> vbNullString Then MsgParagraphMonospacedSetup la:=.laMsg2, latext:=sMsg2Label, tb:=.tbMsg2TextMonospaced, tbtext:=sMsg2TextMonospaced
        If sMsg3TextPrprtional <> vbNullString Then MsgParagraphPrprtionalSetup la:=.laMsg3, latext:=sMsg3Label, tb:=.tbMsg3TextPrprtional, tbtext:=sMsg3TextPrprtional
        If sMsg3TextMonospaced <> vbNullString Then MsgParagraphMonospacedSetup la:=.laMsg3, latext:=sMsg3Label, tb:=.tbMsg3TextMonospaced, tbtext:=sMsg3TextMonospaced
    End With
End Sub

' After final adjustment of the form's width all the message paragraph's width
' is re-adjusted. Any message paragraph using a proportinal font will result in
' a new height, any monospaced font paragraph in a vertival scroll bar.
' -----------------------------------------------------------------------------
Private Sub MsgSectionsWidthFinal()
Dim siMax   As Single
Dim v       As Variant
Dim tb     As MSForms.TextBox
Dim s      As String
 
    siMax = Me.Width - (FORM_MARGIN_LEFT + MARGIN_HORIZONTAL)
    For Each v In MsgSectionsDisplayed
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
            With .laMsgTitle   ' Hidden by default
                .Top = TopNext(Me.laMsgTitle)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .Visible = True
                siTopNext = .Top + .Height + MARGIN_FORM_TOP
            End With
            
        Else
            .Caption = " " & sTitle
            .laMsgTitleSpaceBottom.Visible = False
            With .laMsgTitle
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
        
        With .laMsgTitle
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
        .laMsgTitleSpaceBottom.Width = .laMsgTitle.Width
    
    End With

End Sub

Public Sub FormFinalPositionOnScreen()
    AdjustStartupPosition Me
End Sub

Private Sub ControlPosTop(ByVal ctl As MSForms.Control, _
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
    
    GetScreenMetrics            ' provides the screen's width and height
    With Me
        TitleSetup
        MsgSectionsSetup        ' provided message sections text and visibility
        RepliesSetup vReplies   ' provided reply buttons text and visibility
        FormWidthFinal          ' considers title width, monospaced message sections width and maximum form width
        MsgSectionsWidthFinal   ' may end up with a horizontal scroll bar when monospaced
        RepliesPosLeft          ' adjust displayed reply buttons left position
        ControlsTopPos          ' adjust all displayed controls' top position
        FormHeightFinal         ' may end up with a horizontal scroll bar
        MsgSectionsHeightFinal  ' may end up with a horizontal scroll bar for a monospaced message section
        ControlsTopPos          ' adjusts all controls' top position
    End With
    AdjustStartupPosition Me

End Sub

