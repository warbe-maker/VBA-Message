VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   9255.001
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9480.001
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
Const FRAME_CAPTIONS            As Boolean = False
Const MONOSPACED_FONT_NAME      As String = "Courier New"   ' Default monospaced font
Const MONOSPACED_FONT_SIZE      As Single = 9       ' Default monospaced font size
Const FORM_WIDTH_MIN            As Single = 300     ' Default minimum message form width
Const FORM_WIDTH_MAX_POW        As Long = 80        ' Max form width as a percentage of the screen size
Const FORM_HEIGHT_MAX_POW       As Long = 90        ' Max form height as a percentage of the screen size
Const HSPACE_FRAMES             As Single = 0       ' Horizontal margin of frames
Const HSPACE_LEFT               As Single = 0       ' Left margin for labels and text boxes
Const HSPACE_BUTTONS            As Single = 10      ' Horizontal margin for reply buttons
Const HSPACE_RIGHT              As Single = 15      ' Horizontal right space for labels and text boxes
Const NEXT_ROW                  As String = vbLf    ' Reply button row break
Const VSPACE_AREAS              As Single = 10      ' Vertical space between message area and replies area
Const VSPACE_BOTTOM             As Single = 50      ' Bottom space after the last displayed reply row
Const VSPACE_FRAMES             As Single = 3       ' Vertical space between frames
Const VSPACE_LABEL              As Single = 0       ' Vertical space between label and the following text
Const VSPACE_BUTTON_ROWS        As Single = 5       ' Vertical space between displayed reply rows
Const VSPACE_SECTIONS           As Single = 5       ' Vertical space between displayed message sections
Const VSPACE_TEXTBOXES          As Single = 18      ' Vertical bottom marging for all textboxes
Const VSPACE_TOP                As Single = 2       ' Top position for the first displayed control
Const VSPACE_SCROLLBAR          As Single = 10      ' Additional vertical space required for a frame with a horizontal scroll barr
Const HSPACE_SCROLLBAR          As Single = 18      ' Additional horizontal space required for a frame with a vertical scroll bar
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
Dim vButtons                    As Variant
Dim aButtons                    As Variant
Dim sTitleFontName              As String
Dim sTitleFontSize              As String   ' Ignored when sTitleFontName is not provided
Dim siTopNext                   As Single
Dim siTitleWidth                As Single
Dim wVirtualScreenLeft          As Single
Dim wVirtualScreenTop           As Single
Dim wVirtualScreenWidth         As Single
Dim wVirtualScreenHeight        As Single
Dim lMaxFormHeightPoW           As Long       ' % of the screen height
Dim lMaxFormWidthPoW            As Long       ' % of the screen width
Dim lMinimumFormHeightPoW       As Long       ' % of the screen height
Dim lMinimumFormWidthPoW        As Long       ' % of the screen width
Dim siMaxFormHeight             As Single     ' above converted to excel userform height
Dim siMaxFormWidth              As Single     ' above converted to excel userform width
Dim siMinimumFormHeight         As Single
Dim siMinimumFormWidth          As Single
Dim sMonoSpacedFontName         As String
Dim siMonoSpacedFontSize        As Single
Dim cllDsgnAreas                As New Collection   ' Collection of the two primary/top frames
Dim cllDsgnSections             As New Collection   '
Dim cllDsgnSectionsLabel        As New Collection
Dim cllDsgnSectionsText         As New Collection   ' Collection of section frames
Dim cllDsgnSectionsTextFrame    As New Collection
Dim cllDsgnButtonRows           As New Collection   ' Collection of the designed reply button row frames
Dim cllDsgnRowButtons           As Collection       ' Collection of a designed reply button row's buttons
Dim cllDsgnButtons              As New Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim cllDsgnButtonsFrame         As New Collection   ' The one and only designed reply buttons frame
Dim dctApplButtonsRetVal        As New Dictionary   ' Dictionary of the applied buttons' reply value (key=CommandButton)
Dim cllApplRowButtons           As Collection       ' Collection of the applied buttons of a certain row
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim dctSectionsLabel            As New Dictionary   ' User provided Section Labels text through Property ApplLabel
Dim dctSectionsText             As New Dictionary   ' User provided Section texts through Property ApplText
Dim dctSectionsMonoSpaced       As New Dictionary   ' User provided Section Monospaced option through Property SectionMonospaced
Dim siMaxButtonWidth            As Single
Dim siMaxButtonHeight           As Single
Dim siMaxSectionWidth           As Single
Dim dctSections                 As New Dictionary
Dim vReplyValue                 As Variant

Private Sub UserForm_Initialize()
    
    Dim v       As Variant
    
    On Error GoTo on_error
    
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_WIDTH_MAX_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_HEIGHT_MAX_POW
    siMinimumFormWidth = FORM_WIDTH_MIN                         ' Default UserForm width
    sMonoSpacedFontName = MONOSPACED_FONT_NAME                  ' Default monospaced font
    siMonoSpacedFontSize = MONOSPACED_FONT_SIZE                 ' Default monospaced font
    Me.FramesWithBorder = False
    Me.width = siMinimumFormWidth
    bFramesWithCaption = False
    
    Collect into:=cllDsgnAreas, ctltype:="Frame", fromparent:=Me, ctlheight:=10, ctlwidth:=Me.width - HSPACE_FRAMES
    DsgnButtonsArea.width = 10  ' Will be adjusted to the max replies row width during setup
    
    Collect into:=cllDsgnSections, ctltype:="Frame", fromparent:=DsgnMsgArea, ctlheight:=50, ctlwidth:=DsgnMsgArea.width - HSPACE_FRAMES
    Collect into:=cllDsgnSectionsLabel, ctltype:="Label", fromparent:=cllDsgnSections, ctlheight:=15, ctlwidth:=DsgnMsgArea.width - (HSPACE_FRAMES * 2)
    Collect into:=cllDsgnSectionsTextFrame, ctltype:="Frame", fromparent:=cllDsgnSections, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (HSPACE_FRAMES * 2)
    Collect into:=cllDsgnSectionsText, ctltype:="TextBox", fromparent:=cllDsgnSectionsTextFrame, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (HSPACE_FRAMES * 3)
    
    Collect into:=cllDsgnButtonsFrame, ctltype:="Frame", fromparent:=DsgnButtonsArea, ctlheight:=10, ctlwidth:=10
    Collect into:=cllDsgnButtonRows, ctltype:="Frame", fromparent:=DsgnButtonsFrame, ctlheight:=10, ctlwidth:=10
        
    For Each v In cllDsgnButtonRows
        Collect into:=cllDsgnRowButtons, ctltype:="CommandButton", fromparent:=v, ctlheight:=10, ctlwidth:=MIN_WIDTH_REPLY_BUTTON
        If cllDsgnRowButtons.Count > 0 _
        Then cllDsgnButtons.Add cllDsgnRowButtons
    Next v
    
    Me.Height = VSPACE_AREAS * 4
    bWithFrames = False

exit_sub:
    Exit Sub
    
on_error:
    Stop: Resume Next
End Sub

Private Property Get ApplButtonRetVal(Optional ByVal button As MSForms.CommandButton) As Variant
    ApplButtonRetVal = dctApplButtonsRetVal(button)
End Property

Private Property Let ApplButtonRetVal(Optional ByVal button As MSForms.CommandButton, ByVal v As Variant)
    dctApplButtonsRetVal.Add button, v
End Property

Public Property Let ApplButtons(ByVal v As Variant):                                    vButtons = v:                                                       End Property

Public Property Let ApplTitle(ByVal s As String):                                            sTitle = s:                                                     End Property

Private Property Get DsgnButton(Optional ByVal row As Long, Optional ByVal button As Long) As MSForms.CommandButton
    Set DsgnButton = cllDsgnButtons(row)(button)
End Property

Private Property Get DsgnButtonsFrame() As MSForms.Frame:                               Set DsgnButtonsFrame = cllDsgnButtonsFrame(1):                      End Property

Private Property Get DsgnButtonRow(Optional ByVal row As Long) As MSForms.Frame:        Set DsgnButtonRow = cllDsgnButtonRows(row):                         End Property

Private Property Get DsgnButtonRows() As Collection:                                    Set DsgnButtonRows = cllDsgnButtonRows:                             End Property

Private Property Let DsgnButtons(ByVal cll As Collection):                              cllDsgnButtons.Add cll:                                             End Property

Private Property Get DsgnButtonsArea() As MSForms.Frame:                                Set DsgnButtonsArea = cllDsgnAreas(2):                              End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                    Set DsgnMsgArea = cllDsgnAreas(1):                                  End Property

Private Property Let DsgnRowButtons(ByVal v As MSForms.CommandButton):                  cllDsgnRowButtons.Add v:                                            End Property

Private Property Get DsgnSection(Optional section As Long) As MSForms.Frame:            Set DsgnSection = cllDsgnSections(section):                         End Property

Private Property Get DsgnSectionLabel(Optional section As Long) As MSForms.Label:       Set DsgnSectionLabel = cllDsgnSectionsLabel(section):               End Property

Private Property Get DsgnSections() As Collection:                                      Set DsgnSections = cllDsgnSections:                                 End Property

Private Property Get DsgnSectionText(Optional section As Long) As MSForms.TextBox:      Set DsgnSectionText = cllDsgnSectionsText(section):                 End Property

Private Property Get DsgnSectionTextFrame(Optional ByVal section As Long):              Set DsgnSectionTextFrame = cllDsgnSectionsTextFrame(section):       End Property

Private Property Get DsgnTextFrame(Optional ByVal section As Long) As MSForms.Frame:    Set DsgnTextFrame = cllDsgnSectionsTextFrame(section):              End Property

Private Property Get DsgnTextFrames() As Collection:                                    Set DsgnTextFrames = cllDsgnSectionsTextFrame:                      End Property

Public Property Let ErrSrc(ByVal s As String):                                          sErrSrc = s:                                                        End Property

Private Property Let FormWidth(ByVal w As Single)
    Me.width = Max(Me.width, siMinimumFormWidth, w)
End Property

Public Property Let FramesWithBorder(ByVal B As Boolean)
    
    Dim ctl As MSForms.Control
       
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Frame" Or TypeName(ctl) = "TextBox" Then
            ctl.BorderColor = -2147483638   ' active frame, allows with style none to hide the frame
            If B = False _
            Then ctl.BorderStyle = fmBorderStyleNone _
            Else ctl.BorderStyle = fmBorderStyleSingle
        End If
    Next ctl
    
End Property

Public Property Let FramesWithCaption(ByVal B As Boolean):                              bFramesWithCaption = B:                                             End Property

Private Property Get MaxAreaWidth() As Single:                                          MaxAreaWidth = MaxFormWidthUsable - HSPACE_FRAMES:                 End Property

Public Property Get MaxFormHeight() As Single:                                          MaxFormHeight = siMaxFormHeight:                                    End Property

Public Property Get MaxFormHeightPrcntgOfScreenSize() As Long:                          MaxFormHeightPrcntgOfScreenSize = lMaxFormHeightPoW:                End Property

Public Property Let MaxFormHeightPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormHeightPoW = l
    siMaxFormHeight = wVirtualScreenHeight * (Min(l, 99) / 100)   ' maximum form height based on screen size
End Property

Public Property Get MaxFormWidth() As Single:                                           MaxFormWidth = siMaxFormWidth:                                      End Property

Public Property Get MaxFormWidthPrcntgOfScreenSize() As Long:                           MaxFormWidthPrcntgOfScreenSize = lMaxFormWidthPoW:                  End Property

Public Property Let MaxFormWidthPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormWidthPoW = l
    siMaxFormWidth = wVirtualScreenWidth * (Min(l, 99) / 100)   ' maximum form width based on screen size
End Property

Private Property Get MaxFormWidthUsable() As Single:                                    MaxFormWidthUsable = siMaxFormWidth - (Me.width - Me.InsideWidth):  End Property

Private Property Get MaxSectionWidth() As Single:                                       MaxSectionWidth = MaxAreaWidth - HSPACE_FRAMES - HSPACE_SCROLLBAR:  End Property

Private Property Get MaxTextBoxFrameWidth() As Single:                                  MaxTextBoxFrameWidth = MaxSectionWidth - HSPACE_FRAMES:             End Property

Private Property Get MaxTextBoxWidth() As Single:                                       MaxTextBoxWidth = MaxTextBoxFrameWidth - HSPACE_FRAMES:             End Property

Public Property Get MinFormWidthPrcntgOfScreenSize() As Long:                           MinFormWidthPrcntgOfScreenSize = lMinimumFormWidthPoW:              End Property

Public Property Get MinimumFormWidth() As Single:                                       MinimumFormWidth = siMinimumFormWidth:                              End Property

Public Property Let MinimumFormWidth(ByVal si As Single)
    siMinimumFormWidth = si
    '~~ The maximum form width must never not become less than the minimum width
    If siMaxFormWidth < siMinimumFormWidth Then
       siMaxFormWidth = siMinimumFormWidth
    End If
    lMinimumFormWidthPoW = CInt((siMinimumFormWidth / wVirtualScreenWidth) * 100)
End Property

Public Property Get ReplyValue() As Variant:                                            ReplyValue = vReplyValue:                                           End Property

Public Property Get ApplLabel(Optional ByVal section As Long) As String
    With dctSectionsLabel
        If .Exists(section) _
        Then ApplLabel = dctSectionsLabel(section) _
        Else ApplLabel = vbNullString
    End With
End Property

' Message section properties:
' Message Section Label
Public Property Let ApplLabel(Optional ByVal section As Long, ByVal s As String):        dctSectionsLabel(section) = s:                                  End Property

' Message Section Mono-spaced
Public Property Get ApplMonoSpaced(Optional ByVal section As Long) As Boolean
    With dctSectionsMonoSpaced
        If .Exists(section) _
        Then ApplMonoSpaced = .Item(section) _
        Else ApplMonoSpaced = False
    End With
End Property

Public Property Let ApplMonoSpaced(Optional ByVal section As Long, ByVal B As Boolean):  dctSectionsMonoSpaced(section) = B:                             End Property

Public Property Get ApplText(Optional ByVal section As Long) As String
    With dctSectionsText
        If .Exists(section) _
        Then ApplText = .Item(section) _
        Else ApplText = vbNullString
    End With
End Property

Public Property Let ApplText(Optional ByVal section As Long, ByVal s As String):         dctSectionsText(section) = s:                                   End Property

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

' Return a collection of all visible reply rows
' ---------------------------------------------
Private Function ApplButtonRows() As Collection

    Dim v   As Variant
    Dim cll As New Collection
    
    For Each v In cllDsgnButtonRows
        If v.Visible Then cll.Add v
    Next v
    
    Set ApplButtonRows = cll
    
End Function

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
        Debug.Print .Name
        Debug.Print buttoncaption
        .Visible = True
        .AutoSize = True
        .WordWrap = False ' the longest line determines the buttonindex's width
        .caption = buttoncaption
        .AutoSize = False
        .Height = .Height + 1
        siMaxButtonHeight = mMsg.Max(siMaxButtonHeight, .Height)
        siMaxButtonWidth = Max(siMaxButtonWidth, .width, MIN_WIDTH_REPLY_BUTTON)
    End With
    
    ApplButtonRetVal(cmb) = buttonreturnvalue ' keep record of the setup buttonindex's reply value
    
End Sub

' Setup and position the applied reply buttons and
' calculate the max reply button width.
' -----------------------------------------------------
Private Sub ApplButtonsSetup(ByVal vButtons As Variant)
    
    Dim frRow               As MSForms.Frame
    Dim frArea              As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim frButtons           As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
    Dim v                   As Variant
    Dim lRow                As Long
    Dim lButton             As Long
    Dim siLeftNext          As Single
    Dim cmb                 As MSForms.CommandButton
    Dim lSetupRowButtons    As Long
    Dim cllApplButtonRows   As Collection
    Dim cllApplRowButtons   As Collection
    Dim siMaxRowWidth       As Single
    Dim siMaxRowHeight      As Single
    
    On Error GoTo on_error
    
    '~~ Setup all button's caption and return the maximum button width and height
    If IsNumeric(vButtons) Then
        '~~ Setup a row of standard VB MsgBox reply command ApplButtons
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
        frArea.Visible = True
        DsgnButtonRow(1).Visible = True

    ElseIf TypeName(vButtons) = "String" Then
        '~~ Setup the reply buttons based on the comma delimited string of button captions
        '~~ and row breaks indicated by a vbLf, vbCr, or vbCrLf
        aButtons = Split(vButtons, ",")
        lRow = 1
        lButton = 0
        For Each v In aButtons
            If v <> vbNullString Then
                If v = vbLf Or v = vbCr Or v = vbCrLf Then
                    '~~ prepare for the next row
                    lRow = lRow + 1
                    lButton = 0
                Else
                    frArea.Visible = True
                    lButton = lButton + 1
                    DsgnButtonRow(lRow).Visible = True
                    ApplButtonSetup buttonrow:=lRow, buttonindex:=lButton, buttoncaption:=v, buttonreturnvalue:=v
                End If
            End If
        Next v
    End If
            
    '~~ Assign all applied/visible button the same width and height and adjust their left position
    '~~ Assign all applied/visible button rows the same height
    Set cllApplButtonRows = ApplButtonRows ' Collection of the visible/applied button rows/frames
    For lRow = 1 To cllApplButtonRows.Count
        siLeftNext = HSPACE_FRAMES
        Set frRow = cllApplButtonRows(lRow)
        Debug.Print frRow.Name
        Set cllApplRowButtons = ApplRowButtons(lRow)
        For Each v In cllApplRowButtons
            Set cmb = v
            With cmb
                Debug.Print .Name
                Debug.Print .left
                Debug.Print .Top
                .width = siMaxButtonWidth
                .Height = siMaxButtonHeight
                .left = siLeftNext
                siLeftNext = siLeftNext + .width + HSPACE_BUTTONS         ' set left pos for the next visible button
            End With
        Next v
        
        '~~ Adjust the button's surrounding frame width and height
        '~~ and calculate the maximum botton row's width
        With frRow
            .width = HSPACE_FRAMES + (siMaxButtonWidth * cllApplRowButtons.Count) + (HSPACE_BUTTONS * (cllApplRowButtons.Count - 1)) + HSPACE_FRAMES
            siMaxRowWidth = Max(siMaxRowWidth, .width)
            .Height = siMaxButtonHeight + VSPACE_FRAMES
            siMaxRowHeight = Max(siMaxRowHeight, .Height)
        End With
    
    Next lRow
                                                   
    '~~ Adjust the buttons frame
    With frButtons
        .Visible = True
        .Top = VSPACE_FRAMES
        .width = siMaxRowWidth + (HSPACE_FRAMES * 2)
        .Height = (siMaxRowHeight * cllApplRowButtons.Count) + (VSPACE_BUTTON_ROWS * (cllApplRowButtons.Count - 1)) + VSPACE_FRAMES
        FormWidth = .width + 7
    End With
    '~~ Adjust the button row's surrounding area frame width to the maximum row width
    With frArea
        .Visible = True
        .width = frButtons.width + HSPACE_FRAMES
        FormWidth = .width + 7
    End With
                
    '~~ Center the button rows within the buttons area
    For Each v In cllApplButtonRows
        v.left = (frArea.width / 2) - (v.width / 2)
    Next v

    '~~ Center the buttons area within the UserForm
    With frArea
        .Height = (siMaxButtonHeight * cllApplButtonRows.Count) + (VSPACE_BUTTON_ROWS * (cllApplButtonRows.Count - 1)) + HSPACE_FRAMES + VSPACE_SCROLLBAR
        .left = (Me.InsideWidth / 2) - (frArea.width / 2)
        .width = .width + HSPACE_SCROLLBAR
    End With

exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Return a collection of applied/visible buttons in a button row.
' ---------------------------------------------------------------
Private Function ApplRowButtons(ByVal row As Long) As Collection
    
    Dim cll As New Collection
    Dim cmb As MSForms.CommandButton
    Dim v   As Variant
    
    For Each v In cllDsgnButtons(row)
        Set cmb = v
        If cmb.Visible Then cll.Add v
    Next v
    Set ApplRowButtons = cll
    
End Function

' Apply a vertical scroll bar to the frame (scrollframe) and reduce
' the frames height by a percentage (heightreduction). The original
' frame's height becomes the height of the scroll bar.
' ----------------------------------------------------------------------
Private Sub ApplyVerticalScrollBar(ByVal scrollframe As MSForms.Frame, _
                                   ByVal newheight As Single)
        
    Dim siScrollHeight As Single: siScrollHeight = scrollframe.Height + VSPACE_SCROLLBAR
        
    With scrollframe
        Debug.Print "Original frame height = " & .Height
        .Height = newheight
        Debug.Print "Scroll bar height     = " & siScrollHeight
        Debug.Print "Reduced frame height  = " & .Height
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

' Return a collection of applied/visble message sections
' ------------------------------------------------------
Private Function AppMsgSections() As Collection
        
    Dim v   As Variant
    Dim cll As Collection: Set cll = New Collection
    
    For Each v In cllDsgnSections
        If v.Visible Then cll.Add v
    Next v
    Set AppMsgSections = cll

End Function

' Return the value of the clicked reply button (button).
' --------------------------------------------------------------
Private Sub ButtonClicked(ByVal button As MSForms.CommandButton)
    
    vReplyValue = ApplButtonRetVal(button)
    Me.Hide
    
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
            For Each ctl In Me.Controls
                If TypeName(ctl) = ctltype And ctl.Parent Is fromparent Then
                    With ctl
'                        Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Name: " & ctl.Name
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

' When False (the default) captions are removed from all frames
' Else they remain visible for testing purpose
' -------------------------------------------------------------
Private Sub DisplayFramesWithCaption(ByVal B As Boolean)
            
    Dim ctl As MSForms.Control
       
    If Not B Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.caption = vbNullString
            End If
        Next ctl
    End If

End Sub

' Returns the visible textbox with the largest height.
' ----------------------------------------------------------
Private Function DsgnSectionMaxHeight() As MSForms.TextBox
Dim v   As Variant
Dim si  As Single
Dim tb  As MSForms.TextBox

    For Each v In MsgSectionsTextVisible
        Set tb = v
        If tb.Height > si Then
            si = tb.Height
            Set DsgnSectionMaxHeight = tb
        End If
    Next v
    
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

' Reduces the width of the displayed message sections
' to allow a vertical scroll bar.
' ---------------------------------------------------
Private Sub MsgSectionsDecrementWidth()

    Dim v As Variant
    Dim fr As MSForms.Frame
    
    For Each v In DsgnSections
        Set fr = v
        If fr.Visible Then
            Debug.Print fr.Name & " is visible"
        End If
    Next v

End Sub

' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the RePositioning
' -------------------------------------------------------------
Private Sub MsgSectionSetup(ByVal section As Long)
    
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
    
        frArea.Visible = True
        frSection.Visible = True
        frText.Visible = True
        tbText.Visible = True
                
        If sLabel <> vbNullString Then
            Set la = DsgnSectionLabel(section)
            With la
                .width = Me.InsideWidth - (HSPACE_FRAMES * 2)
                .caption = sLabel
                .Visible = True
            End With
            frText.Top = la.Top + la.Height
        Else
            frText.Top = 0
        End If
        
        If bMonospaced Then
            MsgSectionSetupMonoSpaced section, sText  ' returns the maximum width required for monospaced section
        Else ' proportional spaced
            MsgSectionSetupPropSpaced section, sText
        End If
        tbText.SelStart = 0
        
    End If
        
End Sub

' Setup the monospaced Message Section (section) with the text (text),
' and apply width and adjust surrounding frames accordingly.
' Note: All height adjustments except the one for the text box
'       are done by the RePositioning
' --------------------------------------------------------------------
Private Sub MsgSectionSetupMonoSpaced( _
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
        .left = HSPACE_LEFT
        
        frText.width = .width + HSPACE_FRAMES
        frText.left = HSPACE_LEFT
        
        With frSection
            .width = frText.width + HSPACE_FRAMES
            .left = HSPACE_LEFT
        End With
        
        '~~ The area width considers that there might be a need to apply a vertival scroll bar
        '~~ When the space finally isn't required, the sections are centered within the area
        frArea.width = Max(frArea.width, frSection.left + frSection.width + HSPACE_FRAMES + HSPACE_SCROLLBAR)
        FormWidth = frArea.width + HSPACE_FRAMES + 7
        
        If .width > MaxTextBoxWidth Then
            frText.width = MaxTextBoxFrameWidth
            frSection.width = MaxSectionWidth
            frArea.width = MaxAreaWidth
            Me.width = MaxFormWidth
            
            With frText
                Select Case .ScrollBars
                    Case fmScrollBarsBoth
                    Case fmScrollBarsHorizontal
                    Case fmScrollBarsNone, fmScrollBarsVertical
                        .ScrollBars = fmScrollBarsHorizontal
                        .scrollwidth = tbText.width
                        .Scroll xAction:=fmScrollActionNoChange, yAction:=fmScrollActionEnd
                End Select
            End With
        End If
        
    End With
    siMaxSectionWidth = Max(siMaxSectionWidth, frSection.width)
                    
End Sub

' Setup the proportional spaced Message Section (section) with the text (text)
' Note 1: When proportional spaced Message Sections are setup the width of the
'         Message Form is already final.
' Note 2: All height adjustments except the one for the text box
'         are done by the RePositioning
' -----------------------------------------------------------------------------
Private Sub MsgSectionSetupPropSpaced(ByVal section As Long, _
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
        .width = frArea.width - HSPACE_FRAMES - HSPACE_SCROLLBAR
        .left = HSPACE_LEFT
        siMaxSectionWidth = Max(siMaxSectionWidth, .width)
    End With
    With frText
        .width = frSection.width - HSPACE_FRAMES
        .left = HSPACE_LEFT
    End With
    
    With tbText
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = frText.width - HSPACE_FRAMES
        .value = text
        .SelStart = 0
        .left = HSPACE_LEFT
        frText.width = .left + .width + HSPACE_FRAMES
    End With

End Sub

Private Sub ApplSectionsMonoSpacedSetup()
                             
    If ApplText(1) <> vbNullString And ApplMonoSpaced(1) = True Then MsgSectionSetup section:=1
    If ApplText(2) <> vbNullString And ApplMonoSpaced(2) = True Then MsgSectionSetup section:=2
    If ApplText(3) <> vbNullString And ApplMonoSpaced(3) = True Then MsgSectionSetup section:=3
    
End Sub

Private Sub MsgSectionsPropSpacedSetup()
                
    If ApplText(1) <> vbNullString And ApplMonoSpaced(1) = False Then MsgSectionSetup section:=1
    If ApplText(2) <> vbNullString And ApplMonoSpaced(2) = False Then MsgSectionSetup section:=2
    If ApplText(3) <> vbNullString And ApplMonoSpaced(3) = False Then MsgSectionSetup section:=3
    
End Sub

' Returns a collection of all visible message section text frames
' ---------------------------------------------------------------
Private Function MsgSectionsTextVisible() As Collection
    
    Dim v   As Variant
    Dim cll As New Collection
    
    For Each v In cllDsgnSectionsText
        If v.Visible Then cll.Add v
    Next v
    Set MsgSectionsTextVisible = cll

End Function

' Reduce the final form height to the maximum height specified by reducing
' one of the two areas by the total exceeding height applying a vertcal
' scroll bar or reducing the height of both areas proportionally and applying
' a vertical scroll bar for both.
' --------------------------------------------------------------------------
Private Sub ReduceAreasHeight(ByVal totalexceedingheight As Single)
    
    Dim frMsgArea               As MSForms.Frame:   Set frMsgArea = DsgnMsgArea
    Dim frButtonsArea           As MSForms.Frame:   Set frButtonsArea = DsgnButtonsArea
    Dim frArea                  As MSForms.Frame
    Dim siTotalAreasHeight      As Single
    Dim siAreasExceedingHeight  As Single
    Dim s                       As String
    Dim siWidth                 As Single
    Dim siPrcntgMsgArea         As Single
    Dim siPrcntgButtonsArea     As Single
    With Me
        '~~ Reduce the height to the max height specified
        siAreasExceedingHeight = .Height - siMaxFormHeight
        .Height = siMaxFormHeight
        siTotalAreasHeight = frMsgArea.Height + frButtonsArea.Height
        siPrcntgMsgArea = frMsgArea.Height / siTotalAreasHeight
        siPrcntgButtonsArea = frButtonsArea.Height / siTotalAreasHeight
        
        If siPrcntgMsgArea >= 0.6 Then
            '~~ When the message area requires 60% or more of the total height only this frame
            '~~ will be reduced and applied with a vertical scroll bar.
            ApplyVerticalScrollBar scrollframe:=frMsgArea, _
                                         newheight:=frMsgArea.Height - totalexceedingheight
        ElseIf siPrcntgButtonsArea >= 0.6 Then
            '~~ When the buttons area requires 60% or more it will be reduced and applied with a vertical scroll bar.
            ApplyVerticalScrollBar scrollframe:=frButtonsArea, _
                                     newheight:=frButtonsArea.Height - totalexceedingheight
        Else
            '~~ When one area of the two requires less than 60% of the total areas heigth
            '~~ both will be reduced in the height and get a vertical scroll bar.
            ApplyVerticalScrollBar scrollframe:=frMsgArea, _
                                      newheight:=frMsgArea.Height * siPrcntgMsgArea
            ApplyVerticalScrollBar scrollframe:=frButtonsArea, _
                                   newheight:=frButtonsArea.Height * siPrcntgButtonsArea
        End If
    End With
    
End Sub

' Re-positioning of displayed frames, usually after a width expansion of the UserForm
' -----------------------------------------------------------------------------------
Private Sub RePositioning()
    RePositioningMsgArea
    RePositioningButtonsArea
    RePositioningAreas
End Sub

Private Sub RePositioningAreas()

    Dim v As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = VSPACE_TOP
    For Each v In cllDsgnAreas
        Set fr = v
        With fr
            If .Visible Then
                .Top = siTopNext
                siTopNext = .Top + .Height + VSPACE_AREAS
            End If
        End With
    Next v
    Me.Height = siTopNext - VSPACE_AREAS + (VSPACE_AREAS * 4)
    
End Sub

' Re-position the button frames in the buttons area.
' -------------------------------------------------
Private Sub RePositioningButtonsArea()
    
    On Error GoTo on_error
    
    Dim frArea          As MSForms.Frame:   Set frArea = DsgnButtonsArea
    Dim frButtons       As MSForms.Frame:   Set frButtons = DsgnButtonsFrame
    Dim cllButtonRows   As Collection:      Set cllButtonRows = ApplButtonRows
    Dim frButtonRow     As MSForms.Frame
    Dim v               As Variant
    Dim siHeight        As Single
    Dim lButtonRows     As Long
    
    siTopNext = HSPACE_FRAMES
    
    If cllButtonRows.Count = 0 _
    Then Err.Raise AppErr(1), "fMsg.RePositioningButtonsArea", "None of the designed button rows is visible, has been setup respectively"
    lButtonRows = ApplButtonRows.Count
    
    For Each v In cllButtonRows
        Set frButtonRow = v
        With frButtonRow
            Debug.Print .Name
            siHeight = .Height
            .Top = siTopNext
            siTopNext = .Top + .Height + VSPACE_BUTTON_ROWS
        End With
    Next v

    frButtons.Height = (siHeight * cllButtonRows.Count) + (VSPACE_BUTTON_ROWS * (cllButtonRows.Count - 1)) + VSPACE_FRAMES

    With frArea
        '~~ Additional space for the horizontal scroll bar is added if required
        If .ScrollBars = fmScrollBarsNone Or .ScrollBars = fmScrollBarsVertical Then
            .Height = frButtons.Height + HSPACE_FRAMES + VSPACE_SCROLLBAR
        Else
            .Height = frButtons.Height + HSPACE_FRAMES
        End If
    End With
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

' Re-position all Message Sections vertically
' and adjust the Message Area height accordingly.
' -----------------------------------------------
Private Sub RePositioningMsgArea()
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame: Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame
    Dim lSection    As Long
    Dim siTopNext   As Single
            
    siTopNext = VSPACE_TOP
    
    For lSection = 1 To cllDsgnSections.Count
        Set frSection = DsgnSection(lSection)
        With frSection
            If .Visible Then
                RePositioningMsgSection lSection
                .Top = siTopNext
                frArea.Height = .Top + .Height + VSPACE_FRAMES
                siTopNext = .Top + .Height + VSPACE_SECTIONS
            End If
        End With
    Next lSection
    
    Me.Height = Max(Me.Height, frArea.Top + frArea.Height + (VSPACE_AREAS * 4))
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

Private Sub RePositioningMsgSection(ByVal section As Long)
    
    On Error GoTo on_error
    
    Dim frArea      As MSForms.Frame:   Set frArea = DsgnMsgArea
    Dim frSection   As MSForms.Frame:   Set frSection = DsgnSection(section)
    Dim la          As MSForms.Label:   Set la = DsgnSectionLabel(section)
    Dim frText      As MSForms.Frame:   Set frText = DsgnSectionTextFrame(section)
    Dim tb          As MSForms.TextBox: Set tb = DsgnSectionText(section)
    Dim siTopNext   As Single
    
    siTopNext = 0
    
    If la.Visible Then
        la.Top = 0
        siTopNext = la.Top + la.Height + VSPACE_LABEL
    End If
    
    frText.Top = siTopNext
    tb.Top = 0
    If frText.ScrollBars = fmScrollBarsBoth Or frText.ScrollBars = fmScrollBarsHorizontal Then
        frText.Height = tb.Top + tb.Height + VSPACE_SCROLLBAR + VSPACE_FRAMES
    Else
        frText.Height = tb.Top + tb.Height + VSPACE_FRAMES
    End If
    
    If frArea.ScrollBars = fmScrollBarsBoth Or frArea.ScrollBars = fmScrollBarsVertical Then
        frSection.left = HSPACE_FRAMES
    Else
        frSection.left = (frArea.width / 2) - (frSection.width / 2)
    End If
    
    frSection.Height = frText.Top + frText.Height + VSPACE_FRAMES
    
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
    
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
                .Top = siTopNext
                siTopNext = .Top + .Height + VSPACE_FRAMES
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                titlewidth = .width + HSPACE_RIGHT
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
                .caption = " " & sTitle    ' some left margin
                titlewidth = .width + 30
            End With
            .caption = " " & sTitle    ' some left margin
            .laMsgTitleSpaceBottom.Visible = False
        End If
                
        .laMsgTitleSpaceBottom.width = titlewidth
        FormWidth = titlewidth
    End With
    
End Sub

Private Sub TopPosReplyRows()

    Dim v           As Variant
    Dim frArea      As MSForms.Frame
    Dim frButton    As MSForms.Frame
    
    Set frArea = DsgnButtonsArea
    siTopNext = 0
    
    For Each v In cllDsgnButtonRows
        Set frButton = v
        With frButton
            If .Visible = True Then
                .Top = siTopNext
                siTopNext = .Top + .Height + VSPACE_BUTTON_ROWS
                .Height = siMaxButtonHeight + 2
            End If
        End With
    Next v
    frArea.Height = frButton.Top + frButton.Height + VSPACE_FRAMES + VSPACE_SCROLLBAR
    
End Sub

Private Sub UserForm_Activate()
    
    Dim siTitleWidth    As Single
    
    DisplayFramesWithCaption bFramesWithCaption ' may be True for test purpose
    
    With Me

        '~~ ----------------------------------------------------------------------------------------
        '~~ The  p r i m a r y  setup of the title, the message sections and the reply buttons
        '~~ returns their individual widths which determines the minimum required message form width
        '~~ This setup ends width the final message form width and all elements adjusted to it.
        '~~ ----------------------------------------------------------------------------------------
        .width = siMinimumFormWidth ' Setup starts with the minimum message form width

        '~~ Setup of those elements which determine the final form width
        TitleSetup siTitleWidth
        ApplSectionsMonoSpacedSetup          ' Setup monospaced message sections
        ApplButtonsSetup vButtons           ' Setup the reply buttons
        
        '~~ At this point the form width is final
        '~~ (it may have ended up with its specified minimum width)
        '~~ and the message area's width can be derived from it
        DsgnMsgArea.width = Me.InsideWidth - HSPACE_FRAMES
'        RePositioning
        
        MsgSectionsPropSpacedSetup          ' Setup proportional spaced message sections
        
        RePositioning
        
        '~~ At this point the form height is final. It may however exceed the specified maximum form height.
        '~~ In case the message and/or the buttons area (frame) may be reduced to fit and be provided with
        '~~ a vertical scroll bar. When one area of the two requires less than 60% of the total heigth of both
        '~~ areas, both get a vertical scroll bar, else only the one which uses 60% or more of the height.
        If .Height > siMaxFormHeight Then
            Debug.Print "Form-Height=" & .Height & ", MaxFormHeight=" & siMaxFormHeight
            '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
            ReduceAreasHeight totalexceedingheight:=.Height - siMaxFormHeight
        End If
                        
    End With
    
    RePositioning
    AdjustStartupPosition Me

End Sub

