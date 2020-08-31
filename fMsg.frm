VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   9255.001
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   12390
   OleObjectBlob   =   "fMsg.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "fMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' -------------------------------------------------------------------------------
' UserForm fMsg
' Provides all means for a message with up to 3 separated text sections, either
' proportional- or mono-spaced, with an optional label, and up to 7 reply buttons.
'
' Design:
' Since the implementation is merely design driven its setup is is essential.
' Design changes must comply with the concept. For details please refer to:
'
' Implementation:
'
' Properties:
'
'
' lScreenWidth. Rauschenberger Berlin March 2020
' --------------------------------------------------------------------------
Const BUTTON_MIN_WIDTH          As Single = 70      ' Default minimum reply button width
Const FONT_MONOSPACED_NAME      As String = "Courier New"   ' Default monospaced font
Const FONT_MONOSPACED_SIZE      As Single = 9       ' Default monospaced font size
Const FORM_MAX_HEIGHT_POW       As Long = 90        ' Max form height as a percentage of the screen size
Const FORM_MAX_WIDTH_POW        As Long = 80        ' Max form width as a percentage of the screen size
Const FORM_MIN_WIDTH            As Single = 300     ' Default minimum message form width
Const TEST_WITH_FRAME_BORDERS   As Boolean = False  ' For test purpose only! Display frames with visible border
Const TEST_WITH_FRAME_CAPTIONS  As Boolean = False  ' For test purpose only! Display frames with their test captions (erased by default)
Const HSPACE_BUTTONS            As Single = 10      ' Horizontal margin for reply buttons
Const HSPACE_BUTTON_AREA        As Single = 10      ' Minimum margin between buttons area and form when centered
Const HSPACE_LEFT               As Single = 0       ' Left margin for labels and text boxes
Const HSPACE_RIGHT              As Single = 15      ' Horizontal right space for labels and text boxes
Const HSPACE_SCROLLBAR          As Single = 18      ' Additional horizontal space required for a frame with a vertical scroll bar
Const NEXT_ROW                  As String = vbLf    ' Reply button row break
Const VSPACE_AREAS              As Single = 10      ' Vertical space between message area and replies area
Const VSPACE_BOTTOM             As Single = 50      ' Bottom space after the last displayed reply row
Const VSPACE_BUTTON_ROWS        As Single = 5       ' Vertical space between displayed reply rows
Const VSPACE_LABEL              As Single = 0       ' Vertical space between label and the following text
Const VSPACE_SCROLLBAR          As Single = 12      ' Additional vertical space required for a frame with a horizontal scroll barr
Const VSPACE_SECTIONS           As Single = 5       ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES          As Single = 18      ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                As Single = 2       ' Top position for the first displayed control

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
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

#If Resizable Then
Private resizeEnabled   As Boolean
Private mouseX          As Double
Private mouseY          As Double
Private minWidth        As Double
Private minHeight       As Double
#End If

Dim bDoneTitle                  As Boolean
Dim bDoneButtonsArea            As Boolean
Dim bDoneMonoSpacedSections     As Boolean
Dim bDonePropSpacedSections     As Boolean
Dim bDoneMsgArea                As Boolean
Dim bDoneHdecrement             As Boolean
Dim bDoneHdecrementMsgArea      As Boolean
Dim bDoneHdecrementButtonsArea  As Boolean
Dim bDoneSetup                  As Boolean

Dim bFormEvents                 As Boolean
Dim bTestFrameWithBorders       As Boolean
Dim bDisplayFramesWithCaptions  As Boolean
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim cllDsgnAreas                As New Collection   ' Collection of the two primary/top frames
Dim cllDsgnButtonRows           As New Collection   ' Collection of the designed reply button row frames
Dim cllDsgnButtons              As New Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim cllDsgnRowButtons           As Collection       ' Collection of a designed reply button row's buttons
Dim cllDsgnSections             As New Collection   '
Dim cllDsgnSectionsLabel        As New Collection
Dim cllDsgnSectionsText         As New Collection   ' Collection of section frames
Dim cllDsgnSectionsTextFrame    As New Collection
Dim dctApplied                  As New Dictionary   ' Dictionary of all applied controls (versus just designed)
Dim dctApplButtons              As New Dictionary   ' Dictionary of applied buttons (key=CommandButton, item=row)
Dim dctApplButtonRows           As New Dictionary   ' Dictionary of applied/used/visible button rows (key=frame, item=row)
Dim dctApplButtonsRetVal        As New Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Dim dctSectionsLabel            As New Dictionary   ' User provided Section Labels text through Property ApplLabel
Dim dctSectionsMonoSpaced       As New Dictionary   ' User provided Section Monospaced option through Property SectionMonospaced
Dim dctSectionsText             As New Dictionary   ' User provided Section texts through Property ApplText
Dim cllDsgnButtonsFrame         As New Collection
Dim lMaxFormHeightPoW           As Long             ' % of the screen height
Dim lMaxFormWidthPoW            As Long             ' % of the screen width
Dim lMinimumFormWidthPoW        As Long             ' % of the screen width
Dim sErrSrc                     As String
Dim siMaxButtonHeight           As Single
Dim siMaxButtonWidth            As Single
Dim siMaxFormHeight             As Single           ' above converted to excel userform height
Dim siMaxFormWidth              As Single           ' above converted to excel userform width
Dim siMaxButtonRowWidth         As Single
Dim siMaxSectionWidth           As Single
Dim siMinimumFormWidth          As Single
Dim siMonoSpacedFontSize        As Single
Dim siFramesHmargin             As Single           ' Test property, value defaults to 0
Dim siFramesVmargin             As Single           ' Test property, value defaults to 0
Dim sMonoSpacedFontName         As String
Dim sTitle                      As String
Dim sTitleFontName              As String
Dim sTitleFontSize              As String           ' Ignored when sTitleFontName is not provided
Dim vButtons                    As Variant
Dim vReplyValue                 As Variant
Dim wVirtualScreenHeight        As Single
Dim wVirtualScreenLeft          As Single
Dim wVirtualScreenTop           As Single
Dim wVirtualScreenWidth         As Single

Private Sub UserForm_Initialize()
        
    On Error GoTo on_error
    
    bFormEvents = False
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_MAX_WIDTH_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_MAX_HEIGHT_POW
    siMinimumFormWidth = FORM_MIN_WIDTH                         ' Default UserForm width
    sMonoSpacedFontName = FONT_MONOSPACED_NAME                  ' Default monospaced font
    siMonoSpacedFontSize = FONT_MONOSPACED_SIZE                 ' Default monospaced font
    Me.width = siMinimumFormWidth
    bDisplayFramesWithCaptions = False
    bTestFrameWithBorders = False
    
    Collect into:=cllDsgnAreas, ctltype:="Frame", fromparent:=Me, ctlheight:=10, ctlwidth:=Me.width - siFramesHmargin
    DsgnButtonsArea.width = 10  ' Will be adjusted to the max replies row width during setup
    
    Collect into:=cllDsgnSections, ctltype:="Frame", fromparent:=DsgnMsgArea, ctlheight:=50, ctlwidth:=DsgnMsgArea.width - siFramesHmargin
    Collect into:=cllDsgnSectionsLabel, ctltype:="Label", fromparent:=cllDsgnSections, ctlheight:=15, ctlwidth:=DsgnMsgArea.width - (siFramesHmargin * 2)
    Collect into:=cllDsgnSectionsTextFrame, ctltype:="Frame", fromparent:=cllDsgnSections, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (siFramesHmargin * 2)
    Collect into:=cllDsgnSectionsText, ctltype:="TextBox", fromparent:=cllDsgnSectionsTextFrame, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (siFramesHmargin * 3)
    
    Collect into:=cllDsgnButtonsFrame, ctltype:="Frame", fromparent:=DsgnButtonsArea, ctlheight:=10, ctlwidth:=10
    Collect into:=cllDsgnButtonRows, ctltype:="Frame", fromparent:=cllDsgnButtonsFrame, ctlheight:=10, ctlwidth:=10
        
    Dim v As Variant
    For Each v In cllDsgnButtonRows
        Set cllDsgnRowButtons = New Collection
        Collect into:=cllDsgnRowButtons, ctltype:="CommandButton", fromparent:=v, ctlheight:=10, ctlwidth:=BUTTON_MIN_WIDTH
        If cllDsgnRowButtons.Count > 0 Then
            cllDsgnButtons.Add cllDsgnRowButtons
        End If
    Next v
    
    Me.Height = VSPACE_AREAS * 4
    bWithFrames = False
    siFramesHmargin = 2     ' Ensures proper command buttons framing, may be used for test purpose
    Me.FramesVmargin = 0    ' Ensures proper command buttons framing and vertical positioning of controls
    
    bDoneSetup = False
    bDoneTitle = False
    bDoneButtonsArea = False
    bDoneMonoSpacedSections = False
    bDonePropSpacedSections = False
    bDoneMsgArea = False
    bDoneHdecrement = False
    
'    UnifyBackgroundColor ' not satisfactory
    
exit_sub:
'    Debug.Print cllDsgnButtonsFrame(1).Name
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

Private Sub UserForm_Terminate()

    Set cllDsgnAreas = Nothing
    Set cllDsgnButtonRows = Nothing
    Set cllDsgnButtons = Nothing
    Set cllDsgnRowButtons = Nothing
    Set cllDsgnSections = Nothing
    Set cllDsgnSectionsLabel = Nothing
    Set cllDsgnSectionsText = Nothing
    Set cllDsgnSectionsTextFrame = Nothing
    Set dctApplButtonsRetVal = Nothing
    Set dctSectionsLabel = Nothing
    Set dctSectionsMonoSpaced = Nothing
    Set dctSectionsText = Nothing
    Set cllDsgnButtonsFrame = Nothing
    Set dctApplButtons = Nothing
    
End Sub

Private Property Get ApplButtonRetVal(Optional ByVal Button As MSForms.CommandButton) As Variant
    ApplButtonRetVal = dctApplButtonsRetVal(Button)
End Property

Private Property Let ApplButtonRetVal(Optional ByVal Button As MSForms.CommandButton, ByVal v As Variant)
    dctApplButtonsRetVal.Add Button, v
End Property

Public Property Let ApplButtons(ByVal v As Variant)
    Select Case TypeName(v)
        Case "Long", "String":  vButtons = v
        Case Else:              Set vButtons = v
    End Select
End Property

Private Property Get ApplButtonsRowHeight() As Single:                                  ApplButtonsRowHeight = siMaxButtonHeight + (siFramesVmargin * 2) + 2:       End Property

Private Property Get ApplButtonsRowWidth(Optional ByVal buttons As Long) As Single
    '~~ Extra space rquired for the button's design
    ApplButtonsRowWidth = CInt((siMaxButtonWidth * buttons) + (HSPACE_BUTTONS * (buttons - 1)) + (siFramesHmargin * 2)) + 5
End Property

Private Property Let Applied(ByVal v As Variant)
    If Not IsApplied(v) Then dctApplied.Add v, v.Name
End Property

Public Property Get ApplLabel(Optional ByVal section As Long) As String
    With dctSectionsLabel
        If .Exists(section) _
        Then ApplLabel = dctSectionsLabel(section) _
        Else ApplLabel = vbNullString
    End With
End Property

' Message section properties:
' Message Section Label
Public Property Let ApplLabel(Optional ByVal section As Long, ByVal s As String):        dctSectionsLabel(section) = s:                                             End Property

' Message Section Mono-spaced
Public Property Get ApplMonoSpaced(Optional ByVal section As Long) As Boolean
    With dctSectionsMonoSpaced
        If .Exists(section) _
        Then ApplMonoSpaced = .Item(section) _
        Else ApplMonoSpaced = False
    End With
End Property

Public Property Let ApplMonoSpaced(Optional ByVal section As Long, ByVal b As Boolean): dctSectionsMonoSpaced(section) = b:                                         End Property

Public Property Get ApplText(Optional ByVal section As Long) As String
    With dctSectionsText
        If .Exists(section) _
        Then ApplText = .Item(section) _
        Else ApplText = vbNullString
    End With
End Property

Public Property Let ApplText(Optional ByVal section As Long, ByVal s As String):        dctSectionsText(section) = s:                                               End Property

Public Property Let ApplTitle(ByVal s As String)
    sTitle = s: SetupTitle
End Property

Private Property Get ButtonsFrameWidth() As Single
    ButtonsFrameWidth = siMaxButtonRowWidth + (siFramesHmargin * 2)
End Property

Private Property Get DsgnButton(Optional ByVal row As Long, Optional ByVal Button As Long) As MSForms.CommandButton
    Set DsgnButton = cllDsgnButtons(row)(Button)
End Property

Private Property Get DsgnButtonRow(Optional ByVal row As Long) As MSForms.Frame:        Set DsgnButtonRow = cllDsgnButtonRows(row):                                 End Property

Private Property Get DsgnButtonRows() As Collection:                                    Set DsgnButtonRows = cllDsgnButtonRows:                                     End Property

Private Property Get DsgnButtonsArea() As MSForms.Frame:                                Set DsgnButtonsArea = cllDsgnAreas(2):                                      End Property

Private Property Get DsgnButtonsFrame() As MSForms.Frame:                               Set DsgnButtonsFrame = cllDsgnButtonsFrame(1):                                  End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                    Set DsgnMsgArea = cllDsgnAreas(1):                                          End Property

Private Property Get DsgnSection(Optional section As Long) As MSForms.Frame:            Set DsgnSection = cllDsgnSections(section):                                 End Property

Private Property Get DsgnSectionLabel(Optional section As Long) As MSForms.Label:       Set DsgnSectionLabel = cllDsgnSectionsLabel(section):                       End Property

Private Property Get DsgnSections() As Collection:                                      Set DsgnSections = cllDsgnSections:                                         End Property

Private Property Get DsgnSectionText(Optional section As Long) As MSForms.TextBox:      Set DsgnSectionText = cllDsgnSectionsText(section):                         End Property

Private Property Get DsgnSectionTextFrame(Optional ByVal section As Long):              Set DsgnSectionTextFrame = cllDsgnSectionsTextFrame(section):               End Property

Private Property Get DsgnTextFrame(Optional ByVal section As Long) As MSForms.Frame:    Set DsgnTextFrame = cllDsgnSectionsTextFrame(section):                      End Property

Private Property Get DsgnTextFrames() As Collection:                                    Set DsgnTextFrames = cllDsgnSectionsTextFrame:                              End Property

Public Property Let ErrSrc(ByVal s As String):                                          sErrSrc = s:                                                                End Property

Private Property Let FormWidth(ByVal w As Single)
    Me.width = Max(Me.width, siMinimumFormWidth, w)
End Property

Public Property Let FramesHmargin(ByVal si As Single):                                  siFramesHmargin = si:                                                       End Property

Public Property Get FramesVmargin() As Single:                                          FramesVmargin = siFramesVmargin:                                            End Property

Public Property Let FramesVmargin(ByVal si As Single):                                  siFramesVmargin = VgridPos(si):                                             End Property

Private Property Let HdecrementButtonsArea(ByVal b As Boolean)
    bDoneHdecrementButtonsArea = b
    bDoneHdecrement = b
End Property

Private Property Let HdecrementMsgArea(ByVal b As Boolean)
    bDoneHdecrementMsgArea = b
    bDoneHdecrement = b
End Property

Private Property Get IsApplied(Optional ByVal v As Variant) As Boolean:                 IsApplied = dctApplied.Exists(v):                                           End Property

Private Property Get MaxButtonsAreaWidth() As Single:                                   MaxButtonsAreaWidth = MaxFormWidthUsable - HSPACE_BUTTON_AREA:              End Property

Public Property Get MaxFormHeight() As Single:                                          MaxFormHeight = siMaxFormHeight:                                            End Property

Public Property Get MaxFormHeightPrcntgOfScreenSize() As Long:                          MaxFormHeightPrcntgOfScreenSize = lMaxFormHeightPoW:                        End Property

Public Property Let MaxFormHeightPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormHeightPoW = l
    siMaxFormHeight = wVirtualScreenHeight * (Min(l, 99) / 100)   ' maximum form height based on screen size
End Property

Public Property Get MaxFormWidth() As Single:                                           MaxFormWidth = siMaxFormWidth:                                              End Property

Public Property Get MaxFormWidthPrcntgOfScreenSize() As Long:                           MaxFormWidthPrcntgOfScreenSize = lMaxFormWidthPoW:                          End Property

Public Property Let MaxFormWidthPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormWidthPoW = l
    siMaxFormWidth = wVirtualScreenWidth * (Min(l, 99) / 100)   ' maximum form width based on screen size
End Property

Private Property Get MaxFormWidthUsable() As Single:                                    MaxFormWidthUsable = siMaxFormWidth - (Me.width - Me.InsideWidth):          End Property

Private Property Get MaxMsgAreaWidth() As Single:                                       MaxMsgAreaWidth = MaxFormWidthUsable - siFramesHmargin:                     End Property

Private Property Get MaxRowsHeight() As Single:                                         MaxRowsHeight = siMaxButtonHeight + (siFramesVmargin * 2) + 4:              End Property

Private Property Get MaxSectionWidth() As Single:                                       MaxSectionWidth = MaxMsgAreaWidth - siFramesHmargin - HSPACE_SCROLLBAR:     End Property

Private Property Get MaxTextBoxFrameWidth() As Single:                                  MaxTextBoxFrameWidth = MaxSectionWidth - siFramesHmargin:                   End Property

Private Property Get MaxTextBoxWidth() As Single:                                       MaxTextBoxWidth = MaxTextBoxFrameWidth - siFramesHmargin:                   End Property

Public Property Get MinFormWidth() As Single:                                           MinFormWidth = siMinimumFormWidth:                                          End Property

Public Property Let MinFormWidth(ByVal si As Single)
    siMinimumFormWidth = Max(si, 200) ' cannot be specified less
    '~~ The maximum form width must never not become less than the minimum width
    If siMaxFormWidth < siMinimumFormWidth Then
       siMaxFormWidth = siMinimumFormWidth
    End If
    lMinimumFormWidthPoW = CInt((siMinimumFormWidth / wVirtualScreenWidth) * 100)
End Property

Public Property Get MinFormWidthPrcntg() As Long:                                       MinFormWidthPrcntg = lMinimumFormWidthPoW:                                  End Property

Private Property Get PrcntgHeightButtonsArea() As Single
    PrcntgHeightButtonsArea = Round(DsgnButtonsArea.Height / (DsgnMsgArea.Height + DsgnButtonsArea.Height), 2)
End Property

Private Property Get PrcntgHeightMsgArea() As Single
    PrcntgHeightMsgArea = Round(DsgnMsgArea.Height / (DsgnMsgArea.Height + DsgnButtonsArea.Height), 2)
End Property

Public Property Get ReplyValue() As Variant:                                            ReplyValue = vReplyValue:                                                   End Property

Public Property Let TestFrameWithBorders(ByVal b As Boolean)
    
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

Public Property Let TestFrameWithCaptions(ByVal b As Boolean):                          bDisplayFramesWithCaptions = b:                                                 End Property

Private Sub AdjustButtonsAreaSize()
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim frButtons   As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
    
    If IsApplied(frButtons) Then
        With frButtons
            .Visible = True
            .width = ButtonsFrameWidth
        End With
    End If
    
    If IsApplied(frArea) Then
        With frArea
            .Visible = True
            .Height = frButtons.Top + frButtons.Height + siFramesVmargin
            .width = frButtons.left + frButtons.width + siFramesHmargin
            If bDoneHdecrementButtonsArea Then
                frButtons.left = siFramesHmargin
            Else
                .width = .width + HSPACE_SCROLLBAR
                 CenterHorizontal centerfr:=frButtons, infr:=frArea
            End If
        End With
        CenterHorizontal centerfr:=frArea
    End If
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

Private Sub AdjustButtonsFrameSize()
    
    Dim l As Long:  l = dctApplButtonRows.Count

    With DsgnButtonsFrame
        .Visible = True
        .width = siMaxButtonRowWidth + (siFramesHmargin * 2)
        .Top = siFramesVmargin
        .Height = VgridPos((MaxRowsHeight * l) + (VSPACE_BUTTON_ROWS * (l - 1)) + (siFramesVmargin * 2)) + 5
    End With

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

Public Sub ApplButtonRowsCenter()

    Dim frButtons       As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnButtonRows
    Dim lRow            As Long
    Dim frRow           As MSForms.Frame
    
    For lRow = 1 To cllButtonRows.Count
        Set frRow = cllButtonRows(lRow)
        If IsApplied(frRow) Then
            With frRow
                If .Visible Then
    '                Debug.Print frRow.Name
                    '~~ Center the button rows within the buttons frame horizontaly
                    CenterHorizontal centerfr:=frRow, infr:=frButtons
                End If
            End With
        End If
    Next lRow

End Sub

' Setup an applied reply buttonindex's (buttonindex) visibility and
' caption, calculate the maximum buttonindex width and height,
' keep a record of the setup reply buttonindex's return value.
' --------------------------------------------------------
Private Sub ApplButtonSetup(ByVal buttonrow As Long, _
                            ByVal buttonindex As Long, _
                            ByVal buttoncaption As String, _
                            ByVal buttonreturnvalue As Variant)
    
    Dim cmb As MSForms.CommandButton:   Set cmb = DsgnButton(buttonrow, buttonindex)
    
    With cmb
        .Visible = True
        .AutoSize = True
        .WordWrap = False ' the longest line determines the buttonindex's width
        .caption = buttoncaption
        .AutoSize = False
        .Height = .Height + 1 ' safety margin to ensure proper multilin caption display
        siMaxButtonHeight = Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .width, BUTTON_MIN_WIDTH)
'        Debug.Print "Button " & .Name & " setup"
    End With
    dctApplButtons.Add cmb, buttonrow
    ApplButtonRetVal(cmb) = buttonreturnvalue ' keep record of the setup buttonindex's reply value
    Applied = cmb
    
End Sub

' Setup and position the applied reply buttons and
' calculate the max reply button width.
' Note: When the provided vButtons argument is a string
'       it wil be converted into a collection and the
'       procedure is performed recursively with it.
' -----------------------------------------------------
Private Sub ApplButtonsSetup(ByVal vButtons As Variant)
    
    On Error GoTo on_error
    
    Dim frArea  As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim cll     As Collection
    Dim v       As Variant
    
    Applied = frArea
    Applied = DsgnButtonsFrame

    '~~ Setup all reply button by calculatig their maximum width and height
    Select Case TypeName(vButtons)
        Case "Long":        ApplButtonsSetupFromValue
        Case "String"
            Set cll = New Collection
            For Each v In Split(vButtons, ",")
                cll.Add v
            Next v
            ApplButtonsSetup cll
        Case "Collection"
             ApplButtonsSetupFromCollection vButtons
        Case "Dictionary"
            ApplButtonsSetupFromCollection vButtons
        Case Else
            MsgBox "The format of the provided ""buttons"" argument is not supported!" & vbLf & _
                   "The message will be setup with an Ok only button", vbExclamation
            ApplButtonsSetup vbOKOnly
    End Select
            
    ApplButtonsUnify
                
    AdjustButtonsAreaSize
    
    If frArea.width > MaxButtonsAreaWidth Then
        Debug_Sizes "Buttons area setup done:"
        ApplyScrollBarHorizontal fr:=frArea, widthnew:=MaxButtonsAreaWidth
        CenterHorizontal frArea
    End If

    bDoneButtonsArea = True
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Setup the reply buttons based on the comma delimited string of button captions
' and row breaks indicated by a vbLf, vbCr, or vbCrLf.
' ------------------------------------------------------------------------------
Private Sub ApplButtonsSetupFromCollection(ByVal cllButtons As Collection)

    Dim lRow        As Long
    Dim lButton     As Long
    Dim v           As Variant
    
    lRow = 1
    lButton = 0
    
    For Each v In cllButtons
        If v <> vbNullString Then
            If v = vbLf Or v = vbCr Or v = vbCrLf Then
                '~~ prepare for the next row
                dctApplButtonRows.Add DsgnButtonRow(lRow), lRow
                Applied = DsgnButtonRow(lRow)
                lRow = lRow + 1
                lButton = 0
            Else
                lButton = lButton + 1
                DsgnButtonRow(lRow).Visible = True
                ApplButtonSetup buttonrow:=lRow, buttonindex:=lButton, buttoncaption:=v, buttonreturnvalue:=v
            End If
        End If
    Next v
    dctApplButtonRows.Add DsgnButtonRow(lRow), lRow
    Applied = DsgnButtonRow(lRow)
    DsgnButtonsArea.Visible = True

End Sub

' Setup a row of standard VB MsgBox reply command buttons
' -------------------------------------------------------
Private Sub ApplButtonsSetupFromValue()
    
    Select Case vButtons
        Case vbOKOnly
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Ok", buttonreturnvalue:=vbOK
        Case vbOKCancel
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Ok", buttonreturnvalue:=vbOK
            ApplButtonSetup buttonrow:=1, buttonindex:=2, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNo
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            ApplButtonSetup buttonrow:=1, buttonindex:=2, buttoncaption:="No", buttonreturnvalue:=vbNo
        Case vbRetryCancel
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            ApplButtonSetup buttonrow:=1, buttonindex:=2, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbYesNoCancel
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Yes", buttonreturnvalue:=vbYes
            ApplButtonSetup buttonrow:=1, buttonindex:=2, buttoncaption:="No", buttonreturnvalue:=vbNo
            ApplButtonSetup buttonrow:=1, buttonindex:=3, buttoncaption:="Cancel", buttonreturnvalue:=vbCancel
        Case vbAbortRetryIgnore
            ApplButtonSetup buttonrow:=1, buttonindex:=1, buttoncaption:="Abort", buttonreturnvalue:=vbAbort
            ApplButtonSetup buttonrow:=1, buttonindex:=2, buttoncaption:="Retry", buttonreturnvalue:=vbRetry
            ApplButtonSetup buttonrow:=1, buttonindex:=3, buttoncaption:="Ignore", buttonreturnvalue:=vbIgnore
        Case Else
            Err.Raise AppErr(1), "fMsg.ApplButtonsSetup", "The value provided for the ""buttons"" argument is not a known VB MsgBox value"
    End Select
    DsgnButtonsArea.Visible = True
    DsgnButtonRow(1).Visible = True
    dctApplButtonRows.Add DsgnButtonRow(1), 1
    Applied = DsgnButtonRow(1)
    
End Sub

' - Assign all applied/visible row buttons the same width, height
'   and adjust their left position
' - Assign all applied/visible button rows the same height,
' - Adjust all applied/visible button rows the reqired width
'   calculating the maximum row width.
' ----------------------------------------------------------
Private Sub ApplButtonsUnify()
    
    On Error GoTo on_error
    
    Dim cllButtonRows   As Collection:      Set cllButtonRows = DsgnButtonRows
    Dim frButtons       As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
    Dim frRow           As MSForms.Frame
    Dim siTop           As Single
    Dim vButton         As Variant
    Dim lRow            As Long
    Dim siLeft          As Single
    Dim cll             As New Collection
    Dim v               As Variant
    Dim lButtons        As Long
    
    siTop = siFramesVmargin
    For lRow = 1 To cllButtonRows.Count
        Set frRow = cllButtonRows(lRow)
        With frRow
            lButtons = 0
            If .Visible Then
                cll.Add frRow
'                Debug.Print frRow.Name
                siLeft = siFramesHmargin
                For Each vButton In DsgnRowButtons(lRow)
                    If IsApplied(vButton) Then
                        .Visible = True
                        lButtons = lButtons + 1
                        With vButton
'                            Debug.Print .Name
                            .left = siLeft
                            .width = siMaxButtonWidth
                            .Height = siMaxButtonHeight
                            .Top = siFramesVmargin
                            siLeft = .left + .width + HSPACE_BUTTONS
                        End With
                        .width = vButton.left + vButton.width + siFramesHmargin
                    End If
                Next vButton
                
                siMaxButtonRowWidth = Max(siMaxButtonRowWidth, .width)
                .Height = ApplButtonsRowHeight
                .Top = siTop
                siTop = VgridPos(siTop + .Height + VSPACE_BUTTON_ROWS)
                .width = ApplButtonsRowWidth(lButtons)
            End If
        End With

    Next lRow
    AdjustButtonsFrameSize
    
    '~~ Center all button rows within the buttons frame
    For Each v In cll
        Set frRow = v
'        Debug.Print "Center " & frRow.Name & " (width=" & frRow.width & ") in " & frButtons.Name & " (width=" & frButtons.width & ")"
        CenterHorizontal frRow, frButtons
    Next v
    Set cll = Nothing
        
    ApplButtonRowsCenter
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Reposition all applied/setup control. Executed optionally whenever a control's
' setup had been done (i.e. when a designed control has become an applied/used one),
' obligatory before any height decrement may be due and after any height decrement.
' -----------------------------------------------------------------------------------
Private Sub ApplReposition()
           
    ApplRepositionMsgArea
    ApplRepositionButtonsArea
    ApplRepositionAreas

End Sub

Private Sub ApplRepositionAreas()

    Dim v       As Variant
    Dim siTop   As Single
    
    siTop = siFramesVmargin
    For Each v In cllDsgnAreas
        With v
            If IsApplied(v) Then
                .Visible = True
                .Top = siTop
                siTop = VgridPos(.Top + .Height + VSPACE_AREAS)
            End If
        End With
    Next v
    Me.Height = VgridPos(siTop + (VSPACE_AREAS * 3))
    
End Sub

' Re-position the button frames in the buttons area.
' -------------------------------------------------
Private Sub ApplRepositionButtonsArea()
    
    On Error GoTo on_error
    
    Dim frArea          As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim v               As Variant
    Dim lButtonRows     As Long
    Dim siTop           As Single
    
    siTop = siFramesVmargin
    
    If dctApplButtonRows.Count = 0 Then GoTo exit_proc
    lButtonRows = dctApplButtonRows.Count
    
    '~~ Re-position button rows within the buttons frame vertically
    For Each v In DsgnButtonRows
        If IsApplied(v) Then
            With v
                .Visible = True
                .Top = siTop
                .Height = .Height
                siTop = VgridPos(.Top + .Height + VSPACE_BUTTON_ROWS)
            
            End With
        End If
    Next v
    
    AdjustButtonsFrameSize
    AdjustButtonsAreaSize
       
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Re-position all Message Sections vertically
' and adjust the Message Area height accordingly.
' -----------------------------------------------
Private Sub ApplRepositionMsgArea()
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame: Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame
    Dim lSection    As Long
    Dim siTop       As Single
            
    If Not IsApplied(frArea) Then GoTo exit_proc
    
    siTop = siFramesVmargin
    
    For lSection = 1 To cllDsgnSections.Count
        Set frSection = DsgnSection(lSection)
        With frSection
            If IsApplied(frSection) Then
                .Visible = True
                .left = siFramesVmargin
'                .Height = DsgnSectionTextFrame(lSection).Height + 5 ' extra space to ensure propert display of a scroll bar
                ApplRepositionMsgSection lSection
                .Top = siTop
                If Not bDoneHdecrementMsgArea Then
                    frArea.Height = .Top + .Height + siFramesVmargin
                End If
                siTop = VgridPos(.Top + .Height + VSPACE_SECTIONS)
            Else
                .Visible = False ' not (yet) applied message sections remain invisible
            End If
        End With
    Next lSection
    
    Me.Height = Max(Me.Height, frArea.Top + frArea.Height + (VSPACE_AREAS * 4))
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

Private Sub ApplRepositionMsgSection(ByVal section As Long)
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame:   Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame:   Set frSection = DsgnSection(section)
    Dim la          As MSForms.Label:   Set la = DsgnSectionLabel(section)
    Dim frText      As MSForms.Frame:   Set frText = DsgnSectionTextFrame(section)
    Dim tb          As MSForms.TextBox: Set tb = DsgnSectionText(section)
    Dim siTop       As Single
    
    siTop = siFramesVmargin
    
    If IsApplied(la) Then
        With la
            .Visible = True
            .Top = siTop
            siTop = VgridPos(.Top + .Height + VSPACE_LABEL)
        End With
    End If
    
    If IsApplied(frText) Then
        If IsApplied(tb) Then
            With tb
                .Top = siFramesVmargin
                .Visible = True
            End With
        End If
        With frText
            .Visible = True
            .Top = siTop
            If .ScrollBars = fmScrollBarsBoth Or frText.ScrollBars = fmScrollBarsHorizontal Then
                .Height = tb.Top + tb.Height + VSPACE_SCROLLBAR + siFramesVmargin
            Else
                .Height = tb.Top + tb.Height + siFramesVmargin
            End If
        End With
    End If
    
    If bDoneHdecrementMsgArea _
    Then frSection.left = siFramesHmargin _
    Else CenterHorizontal centerfr:=frSection, infr:=frArea
    
    frSection.Height = frText.Top + frText.Height + siFramesVmargin + 2
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
    
End Sub

' Setup the applied monospaced message section (section) with the text (text),
' and apply width and adjust surrounding frames accordingly.
' Note: All height adjustments except the one for the text box
'       are done by the ApplReposition
' --------------------------------------------------------------------
Private Sub ApplSectionMonoSpacedSetup( _
            ByVal section As Long, _
            ByVal text As String)
            
    Dim frArea          As MSForms.Frame
    Dim frText          As MSForms.Frame
    Dim tbText          As MSForms.TextBox
    Dim frSection       As MSForms.Frame
    
    Set frArea = DsgnMsgArea
    Set frSection = DsgnSection(section)
    Set frText = DsgnSectionTextFrame(section)
    Set tbText = DsgnSectionText(section)
    
    '~~ Setup the textbox
    With tbText
        .Visible = True
        .MultiLine = True
        .WordWrap = False
        .Font.Name = sMonoSpacedFontName
        .Font.Size = siMonoSpacedFontSize
        .AutoSize = True
        .value = text
        .AutoSize = False
        .SelStart = 0
        .left = siFramesHmargin
        .Height = .Height + 2 ' ensure text is not squeeced
        frText.width = .width + (siFramesHmargin * 2)
        frText.left = siFramesHmargin
        
                
        With frSection
            .width = frText.width + (siFramesHmargin * 2)
            .left = siFramesHmargin
        End With
        
        '~~ The area width considers that there might be a need to apply a vertival scroll bar
        '~~ When the space finally isn't required, the sections are centered within the area
        frArea.width = Max(frArea.width, frSection.left + frSection.width + siFramesHmargin + HSPACE_SCROLLBAR)
        FormWidth = frArea.width + siFramesHmargin + 7
        
        Debug.Print .width
        Debug.Print MaxTextBoxWidth
        
        If .width > MaxTextBoxWidth Then
            frSection.width = MaxSectionWidth
            frArea.width = MaxMsgAreaWidth
            Me.width = MaxFormWidth
            ApplyScrollBarHorizontal fr:=frText, widthnew:=MaxTextBoxFrameWidth
        End If
        
    End With
    siMaxSectionWidth = Max(siMaxSectionWidth, frSection.width)
    
    '~~ Keep record of which controls had been applied
    Applied = frArea
    Applied = frSection
    Applied = frText
    Applied = tbText
    
    ApplReposition
    
End Sub

' Setup the proportional spaced Message Section (section) with the text (text)
' Note 1: When proportional spaced Message Sections are setup the width of the
'         Message Form is already final.
' Note 2: All height adjustments except the one for the text box
'         are done by the ApplReposition
' -----------------------------------------------------------------------------
Private Sub ApplSectionPropSpacedSetup(ByVal section As Long, _
                                        ByVal text As String)
    
    Dim frArea      As MSForms.Frame
    Dim frSection   As MSForms.Frame
    Dim frText      As MSForms.Frame
    Dim tbText      As MSForms.TextBox
    
    Set frArea = DsgnMsgArea
    Set frSection = DsgnSection(section)
    Set frText = DsgnSectionTextFrame(section)
    Set tbText = DsgnSectionText(section)
        
    '~~ For proportional spaced message sections the width is determined by the area width
    With frSection
        .width = frArea.width - siFramesHmargin - HSPACE_SCROLLBAR
        .left = HSPACE_LEFT
        siMaxSectionWidth = Max(siMaxSectionWidth, .width)
    End With
    With frText
        .width = frSection.width - siFramesHmargin
        .left = HSPACE_LEFT
    End With
    
    With tbText
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = frText.width - siFramesHmargin
        .value = text
        .SelStart = 0
        .left = HSPACE_LEFT
        frText.width = .left + .width + siFramesHmargin
    End With
    
    '~~ Keep record of which controls had been applied
    Applied = frArea
    Applied = frSection
    Applied = frText
    Applied = tbText

End Sub

' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the ApplReposition
' -------------------------------------------------------------
Private Sub ApplSectionSetup(ByVal section As Long)
    
    Dim frArea      As MSForms.Frame
    Dim frSection   As MSForms.Frame
    Dim la          As MSForms.Label
    Dim tbText      As MSForms.TextBox
    Dim frText      As MSForms.Frame
    Dim sLabel      As String
    Dim sText       As String
    Dim bMonospaced As Boolean

    Set frArea = DsgnMsgArea
    Set frSection = DsgnSection(section)
    Set la = DsgnSectionLabel(section)
    Set tbText = DsgnSectionText(section)
    Set frText = DsgnTextFrame(section)
    
    sLabel = ApplLabel(section)
    sText = ApplText(section)
    bMonospaced = ApplMonoSpaced(section)
    
    frSection.width = frArea.width
    la.width = frSection.width
    frText.width = frSection.width
    tbText.width = frSection.width
        
    If sText <> vbNullString Then
    
        Applied = frArea
        Applied = frSection
        Applied = frText
        Applied = tbText
                
        If sLabel <> vbNullString Then
            Set la = DsgnSectionLabel(section)
            With la
                .width = Me.InsideWidth - (siFramesHmargin * 2)
                .caption = sLabel
            End With
            frText.Top = la.Top + la.Height
            Applied = la
        Else
            frText.Top = 0
        End If
        
        If bMonospaced Then
            ApplSectionMonoSpacedSetup section, sText  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            ApplSectionPropSpacedSetup section, sText
        End If
        tbText.SelStart = 0
        
    End If

End Sub

Private Sub ApplSectionsSetupMonoSpaced()
                             
    If ApplText(1) <> vbNullString And ApplMonoSpaced(1) = True Then ApplSectionSetup section:=1
    If ApplText(2) <> vbNullString And ApplMonoSpaced(2) = True Then ApplSectionSetup section:=2
    If ApplText(3) <> vbNullString And ApplMonoSpaced(3) = True Then ApplSectionSetup section:=3
    bDoneMonoSpacedSections = True

End Sub

Private Sub ApplSectionsSetupPropSpaced()
                
    If ApplText(1) <> vbNullString And ApplMonoSpaced(1) = False Then ApplSectionSetup section:=1
    If ApplText(2) <> vbNullString And ApplMonoSpaced(2) = False Then ApplSectionSetup section:=2
    If ApplText(3) <> vbNullString And ApplMonoSpaced(3) = False Then ApplSectionSetup section:=3
    bDonePropSpacedSections = True
    bDoneMsgArea = True

End Sub

Private Sub ApplyScrollBarHorizontal(ByVal fr As MSForms.Frame, _
                                     ByVal widthnew As Single)
                                     
    Dim siScrollWidth   As Single
    
    With fr
        siScrollWidth = .width
        .width = widthnew
    End With
    Select Case fr.ScrollBars
        Case fmScrollBarsBoth
        Case fmScrollBarsHorizontal
            fr.scrollwidth = siScrollWidth
            fr.Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
        Case fmScrollBarsNone, fmScrollBarsVertical
            fr.ScrollBars = fmScrollBarsHorizontal
            fr.scrollwidth = siScrollWidth
            fr.Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
    End Select
End Sub

' Apply a vertical scroll bar to the frame (scrollframe) and reduce
' the frames height by a percentage (heightreduction). The original
' frame's height becomes the height of the scroll bar.
' ----------------------------------------------------------------------
Private Sub ApplyScrollBarVertical(ByVal scrollframe As MSForms.Frame, _
                                   ByVal newheight As Single)
        
    Dim siScrollHeight As Single: siScrollHeight = scrollframe.Height + VSPACE_SCROLLBAR
        
    With scrollframe
'        Debug.Print "Original frame height = " & .Height
        .Height = newheight
'        Debug.Print "Scroll bar height     = " & siScrollHeight
'        Debug.Print "Reduced frame height  = " & .Height
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
                .ScrollBars = fmScrollBarsBoth
                .ScrollHeight = siScrollHeight
                .KeepScrollBarsVisible = fmScrollBarsBoth
            Case fmScrollBarsNone
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = siScrollHeight
                .KeepScrollBarsVisible = fmScrollBarsVertical
        End Select
    End With
    
End Sub

' Return the value of the clicked reply button (button).
' --------------------------------------------------------------
Private Sub ButtonClicked(ByVal Button As MSForms.CommandButton)
    
    vReplyValue = ApplButtonRetVal(Button)
    Me.Hide
    
End Sub

' Center the frame (fr) horizontally within the frame (frin)
' which defaults to the UserForm when not provided.
' -------------------------------------------------------------
Private Sub CenterHorizontal(ByVal centerfr As MSForms.Frame, _
          Optional ByVal infr As MSForms.Frame = Nothing)
    
    If infr Is Nothing _
    Then centerfr.left = (Me.InsideWidth / 2) - (centerfr.width / 2) _
    Else centerfr.left = (infr.width / 2) - (centerfr.width / 2)
    
End Sub

' Center the frame (fr) vertically within the frame (frin).
' -----------------------------------------------------------
Private Sub CenterVertical(ByVal centerfr As MSForms.Frame, _
                           ByVal infr As MSForms.Frame)
    centerfr.Top = (infr.Height / 2) - (centerfr.heigth / 2)
End Sub

' The reply button click event is the only code using
' control's name - which unfortunately this cannot be avioded.
' ------------------------------------------------------------
Private Sub cmb11_Click():  ButtonClicked Me.cmb11:   End Sub

Private Sub cmb12_Click():  ButtonClicked Me.cmb12:   End Sub

Private Sub cmb13_Click():  ButtonClicked Me.cmb13:   End Sub

Private Sub cmb14_Click():  ButtonClicked Me.cmb14:   End Sub

Private Sub cmb15_Click():  ButtonClicked Me.cmb15:   End Sub

Private Sub cmb16_Click():  ButtonClicked Me.cmb16:   End Sub

Private Sub cmb17_Click():  ButtonClicked Me.cmb17:   End Sub

Private Sub cmb21_Click():  ButtonClicked Me.cmb21:   End Sub

Private Sub cmb22_Click():  ButtonClicked Me.cmb22:   End Sub

Private Sub cmb23_Click():  ButtonClicked Me.cmb23:   End Sub

Private Sub cmb24_Click():  ButtonClicked Me.cmb24:   End Sub

Private Sub cmb25_Click():  ButtonClicked Me.cmb25:   End Sub

Private Sub cmb26_Click():  ButtonClicked Me.cmb26:   End Sub

Private Sub cmb27_Click():  ButtonClicked Me.cmb27:   End Sub

Private Sub cmb31_Click():  ButtonClicked Me.cmb31:   End Sub

Private Sub cmb32_Click():  ButtonClicked Me.cmb32:   End Sub

Private Sub cmb33_Click():  ButtonClicked Me.cmb33:   End Sub

Private Sub cmb34_Click():  ButtonClicked Me.cmb34:   End Sub

Private Sub cmb35_Click():  ButtonClicked Me.cmb35:   End Sub

Private Sub cmb36_Click():  ButtonClicked Me.cmb36:   End Sub

Private Sub cmb37_Click():  ButtonClicked Me.cmb37:   End Sub

Private Sub cmb41_Click():  ButtonClicked Me.cmb41:   End Sub

Private Sub cmb42_Click():  ButtonClicked Me.cmb42:   End Sub

Private Sub cmb43_Click():  ButtonClicked Me.cmb43:   End Sub

Private Sub cmb44_Click():  ButtonClicked Me.cmb44:   End Sub

Private Sub cmb45_Click():  ButtonClicked Me.cmb45:   End Sub

Private Sub cmb46_Click():  ButtonClicked Me.cmb46:   End Sub

Private Sub cmb47_Click():  ButtonClicked Me.cmb47:   End Sub

Private Sub cmb51_Click():  ButtonClicked Me.cmb51:   End Sub

Private Sub cmb52_Click():  ButtonClicked Me.cmb52:   End Sub

Private Sub cmb53_Click():  ButtonClicked Me.cmb53:   End Sub

Private Sub cmb54_Click():  ButtonClicked Me.cmb54:   End Sub

Private Sub cmb55_Click():  ButtonClicked Me.cmb55:   End Sub

Private Sub cmb56_Click():  ButtonClicked Me.cmb56:   End Sub

Private Sub cmb57_Click():  ButtonClicked Me.cmb57:   End Sub

Private Sub cmb61_Click():  ButtonClicked Me.cmb61:   End Sub

Private Sub cmb62_Click():  ButtonClicked Me.cmb62:   End Sub

Private Sub cmb63_Click():  ButtonClicked Me.cmb63:   End Sub

Private Sub cmb64_Click():  ButtonClicked Me.cmb64:   End Sub

Private Sub cmb65_Click():  ButtonClicked Me.cmb65:   End Sub

Private Sub cmb66_Click():  ButtonClicked Me.cmb66:   End Sub

Private Sub cmb67_Click():  ButtonClicked Me.cmb67:   End Sub

Private Sub cmb71_Click():  ButtonClicked Me.cmb71:   End Sub

Private Sub cmb72_Click():  ButtonClicked Me.cmb72:   End Sub

Private Sub cmb73_Click():  ButtonClicked Me.cmb73:   End Sub

Private Sub cmb74_Click():  ButtonClicked Me.cmb74:   End Sub

Private Sub cmb75_Click():  ButtonClicked Me.cmb75:   End Sub

Private Sub cmb76_Click():  ButtonClicked Me.cmb76:   End Sub

Private Sub cmb77_Click():  ButtonClicked Me.cmb77:   End Sub

' Returns all controls of type (ctltype) which do have a parent (fromparent)
' as collection (into) by assigning the an initial height (ctlheight) and width (ctlwidth).
' -----------------------------------------------------------------------------------------
Private Sub Collect(ByRef into As Variant, _
                    ByVal fromparent As Variant, _
                    ByVal ctltype As String, _
                    ByVal ctlheight As Single, _
                    ByVal ctlwidth As Single)

    Dim ctl As MSForms.Control
    Dim v   As Variant
     
    On Error GoTo on_error
        
    Select Case TypeName(fromparent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In fromparent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = ctltype And ctl.Parent Is v Then
                        With ctl
'                            Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Name: " & ctl.Name
                            .Visible = False
                            .Height = ctlheight
                            .width = ctlwidth
                        End With
                        into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
'            Debug.Print "From parent = " & TypeName(fromparent)
            For Each ctl In Me.Controls
                If TypeName(ctl) = ctltype And ctl.Parent Is fromparent Then
                    With ctl
'                        Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Name: " & ctl.Name
                        .Visible = False
                        .Height = ctlheight
                        .width = ctlwidth
                    End With
                    Select Case TypeName(into)
                        Case "Collection":  into.Add ctl
                        Case Else
                            Set into = ctl
                            Debug.Print into.Name
                    End Select
                End If
            Next ctl
    End Select
exist_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
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

Private Sub Debug_Sizes(ByVal stage As String)
#If Test Then
    With Me
        Debug.Print vbLf & stage
        Debug.Print String(Len(stage), "-")
        Debug.Print "Form         width  = " & .width & " (specified max = " & Format(.MaxFormWidth, "##0") & ")"
        Debug.Print "Form         height = " & .Height & " (specified max = " & Format(.MaxFormHeight, "##0") & ")"
        If IsApplied(DsgnMsgArea) Then _
        Debug.Print "Message Area height = " & Format(DsgnMsgArea.Height, "##0") & " (" & PrcntgHeightMsgArea * 100 & "%)"
        If IsApplied(DsgnButtonsArea) Then _
        Debug.Print "Buttons Area height = " & Format(DsgnButtonsArea.Height, "##0") & " (" & PrcntgHeightButtonsArea * 100 & "%)"
    End With
'    Stop
#End If
End Sub

' When False (the default) captions are removed from all frames
' Else they remain visible for testing purpose
' -------------------------------------------------------------
Private Sub DisplayFramesWithCaptions()
            
    Dim ctl As MSForms.Control
       
    If Not bDisplayFramesWithCaptions Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.caption = vbNullString
            End If
        Next ctl
    End If

End Sub

' Return a collection of applied/use/visible buttons in row buttonrow.
' --------------------------------------------------------------------
Private Function DsgnRowButtons(ByVal buttonrow As Long) As Collection
    Set DsgnRowButtons = cllDsgnButtons(buttonrow)
End Function

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

' The user clicked on the lblResizer
' ----------------------------------
Private Sub lblResizer_MouseDown(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)
#If Resizable Then
    resizeEnabled = True    ' Capture the mouse position on click
    mouseX = X
    mouseY = Y
#End If
End Sub

Private Sub lblResizer_MouseMove(ByVal Button As Integer, _
                                 ByVal Shift As Integer, _
                                 ByVal X As Single, _
                                 ByVal Y As Single)
#If Resizable Then
    On Error GoTo on_error
    
    'Check if the UserForm is not resized too small
    Dim allowResize     As Boolean
        
    allowResize = True
    
    If Me.width + X - mouseX < minWidth Then allowResize = False
    If Me.Height + Y - mouseY < minHeight Then allowResize = False
    
    'Check if the mouse clicked on the lblResizer and above minimum size
    If resizeEnabled = True And allowResize = True Then
'        With Me
'            .width = .width + X - mouseX
'            .Height = .Height + Y - mouseY
'        End With
    End If

exit_proc:
    Exit Sub
on_error:
    Debug.Print Err.Description: Stop: Resume Next
#End If
End Sub

Private Sub lblResizer_MouseUp(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal X As Single, _
                               ByVal Y As Single)
#If Resizeable Then
    resizeEnabled = False ' The user un-clicked on the lblResizer
    
    'Resize the UserForm
    With Me
        .width = .width + X - mouseX
        .Height = .Height + Y - mouseY
        .MaxFormWidthPrcntgOfScreenSize = .width / wVirtualScreenWidth
        .MaxFormHeightPrcntgOfScreenSize = .Height / wVirtualScreenHeight
    End With
    Debug.Print "Resized!"
    ApplReposition
    RepositionResizer
#End If
End Sub

Private Function Max(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = 0, _
           Optional ByVal v4 As Variant = 0, _
           Optional ByVal v5 As Variant = 0, _
           Optional ByVal v6 As Variant = 0, _
           Optional ByVal v7 As Variant = 0, _
           Optional ByVal v8 As Variant = 0, _
           Optional ByVal v9 As Variant = 0) As Variant
' -----------------------------------------------------
' Returns the maximum (biggest) of all provided values.
' -----------------------------------------------------
Dim dMax As Double
    dMax = v1
    If v2 > dMax Then dMax = v2
    If v3 > dMax Then dMax = v3
    If v4 > dMax Then dMax = v4
    If v5 > dMax Then dMax = v5
    If v6 > dMax Then dMax = v6
    If v7 > dMax Then dMax = v7
    If v8 > dMax Then dMax = v8
    If v9 > dMax Then dMax = v9
    Max = dMax
End Function

' Reduce the final form height to the maximum height specified by reducing
' one of the two areas by the total exceeding height applying a vertcal
' scroll bar or reducing the height of both areas proportionally and applying
' a vertical scroll bar for both.
' --------------------------------------------------------------------------
Private Sub ReduceAreasHeight(ByVal totalexceedingheight As Single)
    
    Dim frMsgArea               As MSForms.Frame:   Set frMsgArea = DsgnMsgArea
    Dim frButtonsArea           As MSForms.Frame:   Set frButtonsArea = DsgnButtonsArea
    Dim siAreasExceedingHeight  As Single
    
    With Me
        '~~ Reduce the height to the max height specified
        siAreasExceedingHeight = .Height - siMaxFormHeight
        .Height = siMaxFormHeight
        
        If PrcntgHeightMsgArea >= 0.6 Then
            '~~ When the message area requires 60% or more of the total height only this frame
            '~~ will be reduced and applied with a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frMsgArea, _
                                         newheight:=frMsgArea.Height - totalexceedingheight
            HdecrementMsgArea = True
            
        ElseIf PrcntgHeightButtonsArea >= 0.6 Then
            '~~ When the buttons area requires 60% or more it will be reduced and applied with a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frButtonsArea, _
                                     newheight:=frButtonsArea.Height - totalexceedingheight
            HdecrementButtonsArea = True

        Else
            '~~ When one area of the two requires less than 60% of the total areas heigth
            '~~ both will be reduced in the height and get a vertical scroll bar.
            ApplyScrollBarVertical scrollframe:=frMsgArea, _
                                      newheight:=frMsgArea.Height * PrcntgHeightMsgArea
            HdecrementMsgArea = True
            ApplyScrollBarVertical scrollframe:=frButtonsArea, _
                                   newheight:=frButtonsArea.Height * PrcntgHeightButtonsArea
            HdecrementButtonsArea = True
        End If
    End With
    
End Sub

Private Sub RepositionResizer()
#If Resizable Then
    With Me.lblResizer
        .left = Me.InsideWidth - lblResizer.width
        .Top = Me.InsideHeight - lblResizer.Height
        .Visible = True
    End With
#End If
End Sub

Public Sub RowButtons(ByVal row As Long)
    
    Dim i   As Long
    Dim cmb As MSForms.CommandButton
        
    For i = 1 To 7
        Set cmb = DsgnButton(row, i)
        With cmb
            Debug.Print "Name  = " & .Name
            Debug.Print "Top   = " & .Top
            Debug.Print "Left  = " & .left
        End With
    Next i
    
End Sub

Public Sub Setup()
    
    On Error GoTo on_error
    
    If bDoneSetup = True Then GoTo exit_proc
       
    DisplayFramesWithCaptions ' may be True for test purpose
    
    With Me

        '~~ ----------------------------------------------------------------------------------------
        '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
        '~~ returns their individual widths which determines the minimum required message form width
        '~~ This setup ends width the final message form width and all elements adjusted to it.
        '~~ ----------------------------------------------------------------------------------------
        .width = siMinimumFormWidth ' Setup starts with the minimum message form width

        '~~ Setup of those elements which determine the final form width
        If Not bDoneTitle Then SetupTitle
        
        '~~ Setup monospaced message sections
        ApplSectionsSetupMonoSpaced
        ApplReposition
        Debug_Sizes "Monospaced sections setup:"
        
        '~~ Setup the reply buttons
        ApplButtonsSetup vButtons
        ApplReposition
        Debug_Sizes "Monospaced sections and buttons setup:"
            
        '~~ At this point the form width is final - possibly with its specified minimum width.
        '~~ The message area width is adjusted to the form's width
        DsgnMsgArea.width = Me.InsideWidth - siFramesHmargin
        
        '~~ Setup proportional spaced message sections (use the given width)
        ApplSectionsSetupPropSpaced
        ApplReposition
        Debug_Sizes "Message and buttons area setup:"
                
        '~~ At this point the form height is final. It may however exceed the specified maximum form height.
        '~~ In case the message and/or the buttons area (frame) may be reduced to fit and be provided with
        '~~ a vertical scroll bar. When one area of the two requires less than 60% of the total heigth of both
        '~~ areas, both get a vertical scroll bar, else only the one which uses 60% or more of the height.
        If .Height > siMaxFormHeight Then
'            Debug.Print "Form-Height=" & .Height & ", MaxFormHeight=" & siMaxFormHeight
            '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
            ReduceAreasHeight totalexceedingheight:=.Height - siMaxFormHeight
            bDoneHdecrement = True
            Debug_Sizes "Areas had been reduced to fit specified maximum height:"
        End If
                        
    End With
    
    ApplReposition
    Debug_Sizes "All done! Setup and (possibly) height reduced:"

#If Resizable Then
    RepositionResizer
#End If
    AdjustStartupPosition Me

#If Resizable Then
    ' ResizeWindowSettings Me, True
#End If

    bDoneSetup = True
exit_proc:
    Exit Sub

on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' When a specific font name and/or size is specified, the extra title label is actively used
' and the UserForm's title bar is not displayed - which means that there is no X to cancel.
' ------------------------------------------------------------------------------------------
Private Sub SetupTitle()
    
    Dim siTop           As Single
    Dim siTitleWidth    As Single
    
    siTop = 0
    With Me
        '~~ When a font name other then the standard UserForm font name is
        '~~ provided the extra hidden title label which mimics the title bar
        '~~ width is displayed. Otherwise it remains hidden.
        If sTitleFontName <> vbNullString And sTitleFontName <> .Font.Name Then
            With .laMsgTitle   ' Hidden by default
                .Visible = True
                .Top = siTop
                siTop = VgridPos(.Top + .Height)
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                siTitleWidth = .width + HSPACE_RIGHT
            End With
            Applied = .laMsgTitle
            .laMsgTitleSpaceBottom.Visible = True
        Else
            '~~ The extra title label is only used to adjust the form width and remains hidden
            With .laMsgTitle
                With .Font
                    .Bold = False
                    .Name = Me.Font.Name
                    .Size = 8.65   ' Value which comes to a length close to the length required
                End With
                .Visible = False
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                siTitleWidth = .width + 30
            End With
            .caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
                
        .laMsgTitleSpaceBottom.width = siTitleWidth
        FormWidth = siTitleWidth
    End With
    bDoneTitle = True
    
End Sub
'
'Private Sub UnifyBackgroundColor()
'
'    Dim la  As MSForms.Label
'    Dim tb  As MSForms.TextBox
'    Dim fr  As MSForms.Frame
'    Dim v   As Variant
'    Dim l   As Long
'
'    l = Me.BackColor   ' The color for al frames and textboxes
'
'    For Each v In Me.Controls
'        On Error Resume Next
'        Select Case TypeName(v)
'            Case "Label"
'                Set la = v
'                la.BackgroundColor l
'                Debug.Print "BackColor assigned to " & la.Name
'            Case "TextBox"
'                Set tb = v
'                tb.BackgroundColor = l
'                Debug.Print "BackColor assigned to " & tb.Name
'            Case "Frame"
'                Set fr = v
'                fr.BackgroundColor = l
'                fr.BorderColor = l
'                Debug.Print "BackColor assigned to " & fr.Name
'        End Select
'    Next v
'    DoEvents
'
'End Sub

Private Sub UserForm_Activate()
        
    If Not bDoneSetup Then Setup

#If Resizable Then
        RepositionResizer
        minHeight = 125
        minWidth = 125
#End If

    Debug.Print siFramesVmargin
    Debug.Print siFramesHmargin
    
End Sub

Private Sub UserForm_Resize()
#If Resizable Then
    If Not bFormEvents Then GoTo exit_proc
    
exit_proc:
    Exit Sub

on_error:
    Debug.Print Err.Description: Stop: Resume Next
#End If
End Sub

' Returns an integer of which the remainder (Int(si) / 6) is 0.
' Background: A controls content is only properly displayed
' when the top position of it is aligned to such a position
' --------------------------------------------------------------
Public Function VgridPos(ByVal si As Single) As Single
    Dim i As Long
    
    For i = 0 To 6
        If Int(si) = 0 Then
            VgridPos = 0
        Else
            If Int(si) < 6 Then si = 6
            If (Int(si) + i) Mod 6 = 0 Then
                VgridPos = Int(si) + i
                Exit For
            End If
        End If
    Next i

End Function

