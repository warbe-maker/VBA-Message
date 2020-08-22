VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsg 
   ClientHeight    =   8550.001
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   9630.001
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
'         1         Frame DsgnMsgArea
'         1.1       Frame Image
'         1.2       Frame DsgnSection
'         1.2.1     Frame MsgSection1 to n (currently designed is n=3)
'         1.2.1.1   Label MsgSectionLabel1 to ....3
'         1.2.1.2   Frame MsgSectionFrame1 to ...3
'         1.2.1.2.1 TextBox MsgSectionText1 to ...3
'         2         Frame DsgnButtonsArea
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
Const FORM_WIDTH_MAX_POW        As Long = 80    ' Max form width as a percentage of the screen size
Const FORM_HEIGHT_MAX_POW       As Long = 90    ' Max form height as a percentage of the screen size
Const H_SPACE_FRAMES            As Single = 2   ' Horizontal margin of frames
Const H_SPACE_LEFT              As Single = 0   ' Left margin for labels and text boxes
Const H_SPACE_REPLIES           As Single = 10  ' Horizontal margin for reply buttons
Const H_SPACE_RIGHT             As Single = 15  ' Horizontal right space for labels and text boxes
Const V_SPACE_AREAS             As Single = 10  ' Vertical space between message area and replies area
Const V_SPACE_BOTTOM            As Single = 50  ' Bottom space after the last displayed reply row
Const V_SPACE_FRAMES            As Single = 5   ' Vertical space between frames
Const V_SPACE_LABEL             As Single = 0   ' Vertical space between label and the following text
Const V_SPACE_REPLY_ROWS        As Single = 10  ' Vertical space between displayed reply rows
Const V_SPACE_SECTIONS          As Single = 5   ' Vertical space between displayed message sections
Const V_SPACE_TEXTBOXES         As Single = 18  ' Vertical bottom marging for all textboxes
Const V_SPACE_TOP               As Single = 2   ' Top position for the first displayed control
Const V_SPACE_SCROLLBAR         As Single = 10  ' Vertical extra space for a frame with a horizontal scroll barr
Const H_SPACE_SCROLLBAR         As Single = 15  ' Horizontal extra space for a frame with a vertical scroll bar
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
Dim cllDsgnRowButtons           As Collection       ' Collection of a designed reply button rows buttons
Dim cllDsgnButtons              As New Collection   ' Collection of the collection of the designed reply buttons of a certain row
Dim dctApplButtons              As New Dictionary   ' Dictionary of the total number applied buttons' reply value (key=CommandButton)
Dim cllApplRowButtons           As Collection       ' Collection of the applied buttons of a certain row
Dim cllApplButtonsValue         As New Collection   ' Collection of the applied/setup buttons reply value
Dim bWithFrames                 As Boolean          ' for test purpose only, defaults to False
Dim dctSectionsLabel            As New Dictionary   ' User provided Section Labels text through Property SectionLabel
Dim dctSectionsText             As New Dictionary   ' User provided Section texts through Property SectionText
Dim dctSectionsMonoSpaced       As New Dictionary   ' User provided Section Monospaced option through Property SectionMonospaced
Dim siButtonsWidth              As Single
Dim siButtonsHeight             As Single
Dim sNextRow                    As String
Dim dctSections                 As New Dictionary

Private Sub UserForm_Initialize()
    
    Dim v       As Variant
    
    On Error GoTo on_error
    
    sNextRow = "||"
    GetScreenMetrics                                            ' This environment screen's width and height
    Me.MaxFormWidthPrcntgOfScreenSize = FORM_WIDTH_MAX_POW
    Me.MaxFormHeightPrcntgOfScreenSize = FORM_HEIGHT_MAX_POW
    siMinimumFormWidth = FORM_WIDTH_MIN                         ' Default UserForm width
    sMonoSpacedFontName = MONOSPACED_FONT_NAME                  ' Default monospaced font
    siMonoSpacedFontSize = MONOSPACED_FONT_SIZE                 ' Default monospaced font
    Me.FramesWithBorder = False
    Me.width = siMinimumFormWidth
    bFramesWithCaption = False
    
    Collect into:=cllDsgnAreas, ctltype:="Frame", fromparent:=Me, ctlheight:=10, ctlwidth:=Me.width - H_SPACE_FRAMES
    DsgnButtonsArea.width = 10  ' Will be adjusted to the max replies row width during setup
    
    Collect into:=cllDsgnSections, ctltype:="Frame", fromparent:=DsgnMsgArea, ctlheight:=50, ctlwidth:=DsgnMsgArea.width - H_SPACE_FRAMES
    Collect into:=cllDsgnSectionsLabel, ctltype:="Label", fromparent:=cllDsgnSections, ctlheight:=15, ctlwidth:=DsgnMsgArea.width - (H_SPACE_FRAMES * 2)
    Collect into:=cllDsgnSectionsTextFrame, ctltype:="Frame", fromparent:=cllDsgnSections, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (H_SPACE_FRAMES * 2)
    Collect into:=cllDsgnSectionsText, ctltype:="TextBox", fromparent:=cllDsgnSectionsTextFrame, ctlheight:=20, ctlwidth:=DsgnMsgArea.width - (H_SPACE_FRAMES * 3)
    
    Collect into:=cllDsgnButtonRows, ctltype:="Frame", fromparent:=DsgnButtonsArea, ctlheight:=10, ctlwidth:=10
        
    For Each v In cllDsgnButtonRows
        Collect into:=cllDsgnRowButtons, ctltype:="CommandButton", fromparent:=v, ctlheight:=10, ctlwidth:=MIN_WIDTH_REPLY_BUTTON
        If cllDsgnRowButtons.Count > 0 _
        Then cllDsgnButtons.Add cllDsgnRowButtons
    Next v
    
    Me.Height = V_SPACE_AREAS * 4
    bWithFrames = False

exit_sub:
    Exit Sub
    
on_error:
    Stop: Resume Next
End Sub

Public Property Get NextRow() As String:                                                NextRow = sNextRow:                                                 End Property
Private Property Let ApplButtonValue(Optional ByVal button As MSForms.CommandButton, ByVal v As Variant)
    dctApplButtons.Add button, v
End Property
Private Property Get ApplButtonValue(Optional ByVal button As MSForms.CommandButton) As Variant
    ApplButtonValue = dctApplButtons(button)
End Property
Public Property Let buttons(ByVal v As Variant):                                        vButtons = v:                                                       End Property
Public Property Let ErrSrc(ByVal s As String):                                          sErrSrc = s:                                                        End Property
Private Property Let FormWidth(ByVal w As Single):                                      Me.width = Max(Me.width, siMinimumFormWidth, w):                    End Property

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

Public Property Let FramesWithCaption(ByVal b As Boolean):                              bFramesWithCaption = b:                                         End Property
Private Property Get MaxAreaWidth() As Single:                                          MaxAreaWidth = MaxFormWidthUsable - H_SPACE_FRAMES:                 End Property
Public Property Get MaxFormHeight() As Single:                                          MaxFormHeight = siMaxFormHeight:                                End Property
Public Property Get MaxFormHeightPrcntgOfScreenSize() As Long:                          MaxFormHeightPrcntgOfScreenSize = lMaxFormHeightPoW:            End Property

Public Property Let MaxFormHeightPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormHeightPoW = l
    siMaxFormHeight = wVirtualScreenHeight * (Min(l, 99) / 100)   ' maximum form height based on screen size
End Property

Public Property Get MaxFormWidth() As Single:                                           MaxFormWidth = siMaxFormWidth:                                  End Property
Public Property Get MaxFormWidthPrcntgOfScreenSize() As Long:                           MaxFormWidthPrcntgOfScreenSize = lMaxFormWidthPoW:              End Property

Public Property Let MaxFormWidthPrcntgOfScreenSize(ByVal l As Long)
    lMaxFormWidthPoW = l
    siMaxFormWidth = wVirtualScreenWidth * (Min(l, 99) / 100)   ' maximum form width based on screen size
End Property

Private Property Get MaxFormWidthUsable() As Single:                                    MaxFormWidthUsable = siMaxFormWidth - (Me.width - Me.InsideWidth):  End Property
Private Property Get MaxSectionWidth() As Single:                                       MaxSectionWidth = MaxAreaWidth - H_SPACE_FRAMES:                    End Property
Private Property Get MaxTextBoxFrameWidth() As Single:                                  MaxTextBoxFrameWidth = MaxSectionWidth - H_SPACE_FRAMES:            End Property
Private Property Get MaxTextBoxWidth() As Single:                                       MaxTextBoxWidth = MaxTextBoxFrameWidth - H_SPACE_FRAMES:            End Property
Public Property Get MinFormWidthPrcntgOfScreenSize() As Long:                           MinFormWidthPrcntgOfScreenSize = lMinimumFormWidthPoW:          End Property
Public Property Get MinimumFormWidth() As Single:                                       MinimumFormWidth = siMinimumFormWidth:                          End Property

Public Property Let MinimumFormWidth(ByVal si As Single)
    siMinimumFormWidth = si
    '~~ The maximum form width must never not become less than the minimum width
    If siMaxFormWidth < siMinimumFormWidth Then
       siMaxFormWidth = siMinimumFormWidth
    End If
    lMinimumFormWidthPoW = CInt((siMinimumFormWidth / wVirtualScreenWidth) * 100)
End Property

Private Property Get DsgnMsgArea() As MSForms.Frame:                                     Set DsgnMsgArea = cllDsgnAreas(1):                                          End Property
Private Property Get DsgnTextFrame(Optional ByVal section As Long) As MSForms.Frame:     Set DsgnTextFrame = cllDsgnSectionsTextFrame(section):                    End Property
Private Property Get DsgnTextFrames() As Collection:                                     Set DsgnTextFrames = cllDsgnSectionsTextFrame:                            End Property
Private Property Get DsgnSection(Optional section As Long) As MSForms.Frame:             Set DsgnSection = cllDsgnSections(section):                           End Property
Private Property Get DsgnSectionLabel(Optional section As Long) As MSForms.Label:        Set DsgnSectionLabel = cllDsgnSectionsLabel(section):                 End Property
Private Property Get DsgnSections() As Collection:                                       Set DsgnSections = cllDsgnSections:                                   End Property
Private Property Get DsgnSectionText(Optional section As Long) As MSForms.TextBox:       Set DsgnSectionText = cllDsgnSectionsText(section):                   End Property
Private Property Get DsgnSectionTextFrame(Optional ByVal section As Long):               Set DsgnSectionTextFrame = cllDsgnSectionsTextFrame(section):         End Property
Private Property Get DsgnButtonsArea() As MSForms.Frame:                                 Set DsgnButtonsArea = cllDsgnAreas(2):                                      End Property
Private Property Get DsgnButtonRow(Optional ByVal row As Long) As MSForms.Frame:         Set DsgnButtonRow = cllDsgnButtonRows(row):                                End Property

Private Property Get DsgnButton(Optional ByVal row As Long, Optional ByVal button As Long) As MSForms.CommandButton
    Set DsgnButton = cllDsgnButtons(row)(button)
End Property

Private Property Let DsgnRowButtons(ByVal v As MSForms.CommandButton):                  cllDsgnRowButtons.Add v:                                           End Property
Private Property Get DsgnButtonRows() As Collection:                                    Set DsgnButtonRows = cllDsgnButtonRows:                                      End Property
Private Property Let DsgnButtons(ByVal cll As Collection):                              cllDsgnButtons.Add cll:                                        End Property
Private Property Get DsgnRowsSetup() As Long:                                           DsgnRowsSetup = cllApplButtonsValue.Count:                    End Property

' Message section properties:
' Message Section Label
Public Property Let SectionLabel(Optional ByVal section As Long, ByVal s As String)
    
    With dctSectionsLabel
        If .Exists(section) Then .Remove section
        .Add section, s
    End With
    
'    With dctSections
'        If .Exists(section) Then
'            MessageSection = .Item(section)
'            MessageSection.sLabel = s
'            .Remove section
'            .Add section, tSection
'        Else
'            tSection.sLabel = s
'            .Add section, tSection
'        End If
'    End With
    
End Property

Public Property Get SectionLabel(Optional ByVal section As Long) As String
    With dctSectionsLabel
        If .Exists(section) _
        Then SectionLabel = dctSectionsLabel(section) _
        Else SectionLabel = vbNullString
    End With
End Property

' Message Section Mono-spaced
Public Property Get SectionMonoSpaced(Optional ByVal section As Long) As Boolean
    With dctSectionsMonoSpaced
        If .Exists(section) _
        Then SectionMonoSpaced = .Item(section) _
        Else SectionMonoSpaced = False
    End With
End Property

Public Property Let SectionMonoSpaced(Optional ByVal section As Long, ByVal b As Boolean)
    With dctSectionsMonoSpaced
        If .Exists(section) Then .Remove section
        .Add section, b
    End With
End Property

Public Property Get SectionText(Optional ByVal section As Long) As String
    With dctSectionsText
        If .Exists(section) _
        Then SectionText = .Item(section) _
        Else SectionText = vbNullString
    End With
End Property

Public Property Let SectionText(Optional ByVal section As Long, ByVal s As String)
    With dctSectionsText
        If .Exists(section) Then .Remove section
        .Add section, s
    End With
End Property

Public Property Let title(ByVal s As String):                                               sTitle = s:                                                 End Property

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
Private Sub DisplayFramesWithCaption(ByVal b As Boolean)
            
    Dim ctl As MSForms.Control
       
    If Not b Then
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Frame" Then
                ctl.caption = vbNullString
            End If
        Next ctl
    End If

End Sub

' Final form height adjustment considering only the maximum height specified
' --------------------------------------------------------------------------
Private Sub FormHeightFinal()
    
    Dim siHeightExceeding   As Single
    Dim s                   As String
    Dim siWidth             As Single
    
    With Me
        '~~ Reduce the height of the largest displayed message paragraph by the amount of exceeding height
        siHeightExceeding = .Height - siMaxFormHeight
        .Height = siMaxFormHeight
        With DsgnSectionMaxHeight
            siWidth = .width
            s = .value
            .SetFocus
            .AutoSize = False
            .value = vbNullString
            Select Case .ScrollBars
                Case fmScrollBarsHorizontal
                    .ScrollBars = fmScrollBarsVertical
                    .width = siWidth + H_SPACE_SCROLLBAR
                    .Height = .Height - siHeightExceeding - V_SPACE_SCROLLBAR
                Case fmScrollBarsVertical
                    .ScrollBars = fmScrollBarsVertical
                Case fmScrollBarsBoth
                    .Height = .Height - siHeightExceeding - H_SPACE_SCROLLBAR
                    .width = siWidth - V_SPACE_SCROLLBAR
                Case fmScrollBarsNone
                    .ScrollBars = fmScrollBarsVertical
                    .width = siWidth + H_SPACE_SCROLLBAR
                    .Height = .Height - siHeightExceeding
            End Select
            .value = s
            .SelStart = 0
        End With
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

' Setup a message section with its label when one is specified
' and return the message's width when greater than any other.
' Note: All height adjustments except the one for the text box
'       are done by the VerticalPositioning
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
    
    sLabel = SectionLabel(section)
    sText = SectionText(section)
    bMonospaced = SectionMonoSpaced(section)
    
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
                .width = Me.width - (H_SPACE_FRAMES * 2)
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
    
    DoEvents
    
End Sub

' Setup the monospaced Message Section (section) with the text (text),
' and apply width and adjust surrounding frames accordingly.
' Note: All height adjustments except the one for the text box
'       are done by the VerticalPositioning
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
    
        frText.width = .width + H_SPACE_FRAMES
        frSection.width = frText.width + H_SPACE_FRAMES
        frArea.width = Max(frArea.width, frSection.width + H_SPACE_FRAMES)
        FormWidth = frArea.width + H_SPACE_FRAMES
        
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
            DoEvents
        End If
        
    End With
                       
End Sub

' Setup the proportional spaced Message Section (section) with the text (text)
' Note 1: When proportional spaced Message Sections are setup the width of the
'         Message Form is already final.
' Note 2: All height adjustments except the one for the text box
'         are done by the VerticalPositioning
' -----------------------------------------------------------------------------
Private Sub MsgSectionSetupPropSpaced(ByVal section As Long, _
                                        ByVal text As String)
    
    Dim frSection   As MSForms.Frame
    Dim frText      As MSForms.Frame
    
    Set frSection = DsgnSection(section)
    Set frText = DsgnSectionTextFrame(section)
    
    With DsgnSectionText(section)
        .Visible = True
        .MultiLine = True
        .AutoSize = True
        .WordWrap = True
        .width = Me.width - (H_SPACE_FRAMES * 4)
        .value = text
        .SelStart = 0
        frText.width = .width + H_SPACE_FRAMES
    End With
    
    frSection.width = frText.width + H_SPACE_FRAMES
                                       
End Sub

' Executed only in case the form width had to be reduced in order to meet the specified maximum height.
' The message section with the largest height will be reduced to fit an will receive a vertical scroll bar.
' ---------------------------------------------------------------------------------------------------------
Private Sub MsgSectionsFinalHeight()
    
    Dim siHeightCurrentRequired As Single
    Dim siHeightExceeding       As Single
    Dim s                       As String

'    With Me
'        If .frRepliesRow2.Visible Then
'            siHeightCurrentRequired = .frRepliesRow2.Top + .frRepliesRow2.Height + V_SPACE_BOTTOM
'        Else
'            siHeightCurrentRequired = .frRepliesRow1.Top + .frRepliesRow1.Height + V_SPACE_BOTTOM
'        End If
'    End With
    If siHeightCurrentRequired <= Me.Height Then Exit Sub
    
    siHeightExceeding = siHeightCurrentRequired > Me.Height
    '~~ All displayed controls together take more height than the available form's height
    '~~ The displayed message sections are reduced in their height to fit the available space
    With DsgnSectionMaxHeight ' The message paragraph with the maximum height
        .SetFocus
        s = .value
        Select Case .ScrollBars
            Case fmScrollBarsHorizontal
            Case fmScrollBarsVertical
                .width = .width + H_SPACE_SCROLLBAR
                .ScrollBars = fmScrollBarsBoth
                If Me.width < H_SPACE_LEFT + .width + H_SPACE_RIGHT + H_SPACE_SCROLLBAR Then
                    Me.width = Me.width + H_SPACE_SCROLLBAR
                End If
            Case fmScrollBarsNone
                .width = .width + H_SPACE_SCROLLBAR
                .ScrollBars = fmScrollBarsVertical
                If Me.width < H_SPACE_LEFT + .width + H_SPACE_RIGHT + H_SPACE_SCROLLBAR Then
                    Me.width = Me.width + H_SPACE_SCROLLBAR
                End If
            Case fmScrollBarsBoth       ' nothing required
        End Select
    End With
End Sub

Private Sub MsgSectionsMonoSpacedSetup()
                             
    If SectionText(1) <> vbNullString And SectionMonoSpaced(1) = True Then MsgSectionSetup section:=1
    If SectionText(2) <> vbNullString And SectionMonoSpaced(2) = True Then MsgSectionSetup section:=2
    If SectionText(3) <> vbNullString And SectionMonoSpaced(3) = True Then MsgSectionSetup section:=3
    
End Sub

Private Sub MsgSectionsPropSpacedSetup()
                
    If SectionText(1) <> vbNullString And SectionMonoSpaced(1) = False Then MsgSectionSetup section:=1
    If SectionText(2) <> vbNullString And SectionMonoSpaced(2) = False Then MsgSectionSetup section:=2
    If SectionText(3) <> vbNullString And SectionMonoSpaced(3) = False Then MsgSectionSetup section:=3
    
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

' - Setup the reply button's (cmb) visibility and caption
' - Collect the setup command button for the row it is setup
' - Collect the setup command buttons return value when clicked
' - Keep record of the maximum button width (siButtonsWidth)
' - Keep record of the maximum button height (siButtonsHeight)
' - Return the left position for the next button (leftNext).
' -------------------------------------------------------------
Private Sub DsgnButtonSetup(ByVal row As Long, _
                             ByVal button As Long, _
                             ByVal caption As String, _
                             ByVal returnvalue As Variant)
    
    Dim cmb As MSForms.CommandButton
    Set cmb = DsgnButton(row, button)
    
    With cmb
        Debug.Print .Name
        .Visible = True
        .AutoSize = True
        .WordWrap = False
        .caption = caption
        siButtonsHeight = mMsg.Max(siButtonsHeight, .Height)
        siButtonsWidth = Max(siButtonsWidth, .width, MIN_WIDTH_REPLY_BUTTON)
    End With
    
    '~~ Keep record of the applied button's return value (when clicked)
    ApplButtonValue(cmb) = returnvalue ' keep record of the setup buttons reply value
    
End Sub

' Setup and position the displayed reply buttons.
' Return the max reply button width.
' ------------------------------------------------------
Private Sub DsgnButtonsSetup(ByVal vButtons As Variant)
    
    Dim fr                  As MSForms.Frame
    Dim frArea              As MSForms.Frame
    Dim v                   As Variant
    Dim row                 As Long
    Dim button              As Long
    Dim siLeftNext          As Single
    Dim cmb                 As MSForms.CommandButton
    Dim lSetupRowButtons    As Long
    Dim cll                 As Collection
    Dim siRowsMaxWidth      As Single
    
    Set frArea = DsgnButtonsArea
    
    With Me
        '~~ Setup all button's caption and return the maximum button width and height
        If IsNumeric(vButtons) Then
            '~~ Setup a row of standard VB MsgBox reply command buttons
            Select Case vButtons
                Case vbOKOnly
                    DsgnButtonSetup row:=1, button:=1, caption:="Ok", returnvalue:=vbOK
                Case vbOKCancel
                    DsgnButtonSetup row:=1, button:=1, caption:="Ok", returnvalue:=vbOK
                    DsgnButtonSetup row:=1, button:=2, caption:="Cancel", returnvalue:=vbCancel
                Case vbYesNo
                    DsgnButtonSetup row:=1, button:=1, caption:="Yes", returnvalue:=vbYes
                    DsgnButtonSetup row:=1, button:=2, caption:="No", returnvalue:=vbNo
                Case vbRetryCancel
                    DsgnButtonSetup row:=1, button:=1, caption:="Retry", returnvalue:=vbRetry
                    DsgnButtonSetup row:=1, button:=2, caption:="Cancel", returnvalue:=vbCancel
                Case vbYesNoCancel
                    DsgnButtonSetup row:=1, button:=1, caption:="Yes", returnvalue:=vbYes
                    DsgnButtonSetup row:=1, button:=2, caption:="No", returnvalue:=vbNo
                    DsgnButtonSetup row:=1, button:=3, caption:="Cancel", returnvalue:=vbCancel
                Case vbAbortRetryIgnore
                    DsgnButtonSetup row:=1, button:=1, caption:="Abort", returnvalue:=vbAbort
                    DsgnButtonSetup row:=1, button:=2, caption:="Retry", returnvalue:=vbRetry
                    DsgnButtonSetup row:=1, button:=3, caption:="Ignore", returnvalue:=vbIgnore
            End Select

        Else
            '~~ Setup the reply buttons in the ButtonsArea
            '~~ An element = vbLf or vbCrLf indicates "next row"
            aButtons = Split(vButtons, ",")
            row = 1
            button = 0
            For Each v In aButtons
                If v <> vbNullString Then
                    If v = NextRow Then
                        '~~ Finish the setup of a rows reply buttons by adjusting
                        '~~ the surrounding frame's width and height
                        With DsgnButtonRow(row)
                            .Height = siButtonsHeight + H_SPACE_FRAMES
                            frArea.width = Max(frArea.width, .width) ' Adjust the area frame to the widest replies row frame
                        End With
                        frArea.width = (siButtonsWidth * button) + (H_SPACE_RIGHT * (lSetupRowButtons - 1))
                        '~~ prepare for the next row
                        row = row + 1
                        button = 0
                    Else
                        button = button + 1
                        DsgnButtonRow(row).Visible = True
                        DsgnButtonSetup row:=row, button:=button, caption:=v, returnvalue:=v
                    End If
                End If
            Next v
        
            With DsgnButtonRow(row)
                .Visible = True
                .Height = siButtonsHeight + V_SPACE_FRAMES
                .width = H_SPACE_FRAMES + (siButtonsWidth * dctApplButtons.Count) + (H_SPACE_REPLIES * (dctApplButtons.Count - 1)) + H_SPACE_FRAMES
                frArea.width = Max(frArea.width, .width) ' Adjust the area frame to the widest replies row frame
            End With
            
            FormWidth = frArea.width ' will extend the form width if it is a new maximum

        End If
    End With
        
    '~~ Assign all visible buttons the same width and height and the left position
    siLeftNext = H_SPACE_FRAMES
    For row = 1 To cllDsgnButtons.Count
        If DsgnButtonRow(row).Visible Then
            frArea.Visible = True
            Set cll = DsgnButtonsRowVisible(row)
            For Each v In cll
                Set cmb = v
                With cmb
                    If Not .Visible Then Exit For
                    .width = siButtonsWidth
                    .Height = siButtonsHeight
                    .left = siLeftNext
                    siLeftNext = siLeftNext + .width + H_SPACE_REPLIES         ' set left pos for the next visible button
                End With
            Next v
            
            '~~ Adjust the button's surrounding frame width and height
            '~~ and calculate the maximum botton row's width
            With DsgnButtonRow(row)
                .width = H_SPACE_FRAMES + (siButtonsWidth * cll.Count) + (H_SPACE_REPLIES * (cll.Count - 1)) + H_SPACE_FRAMES
                siRowsMaxWidth = Max(siRowsMaxWidth, .width)
                .Height = siButtonsHeight + V_SPACE_REPLY_ROWS
            End With
                                    
        End If ' button row visible
    Next row
                
    '~~ Adjust the button row's surrounding area frame to the row with the maximum width
    With frArea
        .Visible = True
        .width = siRowsMaxWidth + H_SPACE_FRAMES * 2
        FormWidth = .width
    End With
                
    '~~ Center the button rows within the buttons area
    For Each v In cllDsgnButtonRows
        Set fr = v
        If fr.Visible Then
            fr.left = (frArea.width / 2) - (fr.width / 2)
        End If
    Next v

    '~~ Center the buttons area within the UserForm
    With frArea
        .Height = ((siButtonsHeight + V_SPACE_REPLY_ROWS) * DsgnRowsSetup) + H_SPACE_FRAMES
        .left = (Me.width / 2) - (frArea.width / 2)
    End With
    
End Sub

Private Function DsgnButtonsRowVisible(ByVal row As Long) As Collection
    
    Dim cll As New Collection
    Dim cmb As MSForms.CommandButton
    Dim v   As Variant
    
    For Each v In cllDsgnButtons(row)
        Set cmb = v
        If cmb.Visible Then cll.Add v
    Next v
    Set DsgnButtonsRowVisible = cll
    
End Function
' Return the value of the clicked reply button (button) in row (row).
' -------------------------------------------------------------------
Private Sub ButtonClicked(ByVal button As MSForms.CommandButton)
    
    Dim s As String
    
    s = ApplButtonValue(button)
    If IsNumeric(s) Then
        mMsg.MsgReply = CLng(s)
    Else
        mMsg.MsgReply = s
    End If
    Unload Me
    
End Sub

' Return the collection of all visible reply rows
' -----------------------------------------------
Private Function ButtonRowsVisible() As Collection

    Dim v As Variant
    Dim cll As New Collection
    
    For Each v In cllDsgnButtonRows
        If v.Visible Then cll.Add v
    Next v
    
    Set ButtonRowsVisible = cll
    
End Function

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
                siTopNext = .Top + .Height + V_SPACE_FRAMES
                .Font.Name = sTitleFontName
                If sTitleFontSize <> 0 Then
                    .Font.Size = sTitleFontSize
                End If
                .AutoSize = True
                .caption = " " & sTitle    ' some left margin
                titlewidth = .width + H_SPACE_RIGHT
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

    Dim v As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = 0
    For Each v In cllDsgnButtonRows
        Set fr = v
        With fr
            If .Visible = True Then
                .Top = siTopNext
                siTopNext = .Top + .Height + V_SPACE_REPLY_ROWS
                .Height = siButtonsHeight + 2
                DsgnButtonsArea.Height = .Top + .Height + V_SPACE_AREAS
            End If
        End With
    Next v
    
End Sub

Private Sub UserForm_Activate()
    
    Dim siTitleWidth    As Single

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
        
        MsgSectionsMonoSpacedSetup          ' Setup monospaced message sections
'        VerticalPositioning                 ' Appropriate here for testing only
        
        DsgnButtonsSetup vButtons          ' Setup Reply Buttons
'        VerticalPositioning                 ' Appropriate here for testing only
        
        MsgSectionsPropSpacedSetup          ' Setup proportional spaced message sections
'        VerticalPositioning                 ' Appropriate here for testing only
        
        '~~ ---------------------------------------------------------------------------------------------
        '~~ The  f i n a l  setup considers the height required for the message sections and the reply
        '~~ buttons. This height is reduced when it exceeds the maximum height specified (as a percentage
        '~~ of the available screen size). The largest message section may receive a vertical scroll bar.
        '~~ ---------------------------------------------------------------------------------------------
        If .Height > siMaxFormHeight Then
        '~~ Reduce height to maximum specified and adjust height of message section(s) accordingly
            .Height = siMaxFormHeight
            FormHeightFinal
        End If
                
'        VerticalPositioning
        
    End With
    
    VerticalPositioning
    AdjustStartupPosition Me

End Sub

' Re-positioning of displayed frames, usually after a width expansion of the UserForm
' -----------------------------------------------------------------------------------
Private Sub VerticalPositioning()
    VerticalPositioningMsgArea
    VerticalPositioningButtonsArea
    VerticalPositioningAreas
    DoEvents
End Sub

Private Sub VerticalPositioningAreas()

    Dim v As Variant
    Dim fr  As MSForms.Frame
    
    siTopNext = V_SPACE_TOP
    For Each v In cllDsgnAreas
        Set fr = v
        With fr
            If .Visible Then
                .Top = siTopNext
                siTopNext = .Top + .Height + V_SPACE_AREAS
            End If
        End With
    Next v
    Me.Height = siTopNext - V_SPACE_AREAS + (V_SPACE_AREAS * 4)
    
End Sub

' Re-position all Message Sections vertically
' and adjust the Message Area height accordingly.
' -----------------------------------------------
Private Sub VerticalPositioningMsgArea()
    
    Dim frArea              As MSForms.Frame
    Dim frSection           As MSForms.Frame
    Dim lSection            As Long
    Dim siTopNext           As Single
    
    On Error GoTo on_error
        
    Set frArea = DsgnMsgArea
    siTopNext = V_SPACE_TOP
    
    For lSection = 1 To cllDsgnSections.Count
        
        Set frSection = DsgnSection(lSection)
        
        With frSection
            If .Visible Then
                VerticalPositioningMsgSection lSection
                
                .Top = siTopNext
                frArea.Height = .Top + .Height + V_SPACE_FRAMES
                siTopNext = .Top + .Height + V_SPACE_SECTIONS
            End If
        End With
        
    Next lSection
    Me.Height = Max(Me.Height, frArea.Top + frArea.Height + (V_SPACE_AREAS * 4))
    
    DoEvents

exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub

Private Sub VerticalPositioningMsgSection(ByVal section As Long)
    
    Dim frSection   As MSForms.Frame
    Dim la          As MSForms.Label
    Dim frText      As MSForms.Frame
    Dim tb          As MSForms.TextBox
    Dim si          As Single
    
    Set frSection = DsgnSection(section)
    Set la = DsgnSectionLabel(section)
    Set frText = DsgnSectionTextFrame(section)
    Set tb = DsgnSectionText(section)
    
    si = 0
    
    If la.Visible Then
        la.Top = 0
        si = la.Top + la.Height + V_SPACE_LABEL
    End If
    
    frText.Top = si
    tb.Top = 0
    If frText.ScrollBars = fmScrollBarsBoth Or frText.ScrollBars = fmScrollBarsHorizontal Then
        frText.Height = tb.Top + tb.Height + V_SPACE_SCROLLBAR + V_SPACE_FRAMES
    Else
        frText.Height = tb.Top + tb.Height + V_SPACE_FRAMES
    End If
    frSection.Height = frText.Top + frText.Height + V_SPACE_FRAMES

End Sub

' - Set the top position for all displayed Reply Rows,
' - Center the displayed Reply Rows within the Buttons Area,
' - Set the final height of the Buttons Area.
' ------------------------------------------------------
Private Sub VerticalPositioningButtonsArea()
    
    Dim frRow       As MSForms.Frame
    Dim v           As Variant
    Dim siCenter    As Single
    Dim siHeight    As Single
    Dim cll         As Collection
    
    On Error GoTo on_error
    
    siTopNext = H_SPACE_FRAMES
    siCenter = DsgnButtonsArea.width / 2
    Set cll = ButtonRowsVisible
    
    For Each v In cll
        Set frRow = v
        With frRow
            siHeight = .Height
            .Top = siTopNext
            siTopNext = .Top + .Height + V_SPACE_REPLY_ROWS
            '~~ Center the replies row within the replies area
            .left = siCenter - (.width / 2)
        End With
    Next v
    
    If cll.Count > 0 Then DsgnButtonsArea.Height = siHeight * cll.Count
        
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
'            Debug.Print "Visble= " & .Visible
        End With
    Next i
    
End Sub
