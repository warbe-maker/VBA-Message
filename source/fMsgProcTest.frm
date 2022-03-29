VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fMsgProcTest 
   Caption         =   "Test-Msg-Form"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
   OleObjectBlob   =   "fMsgProcTest.frx":0000
End
Attribute VB_Name = "fMsgProcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SCROLL_V_WIDTH            As Single = 18              ' Additional horizontal space required for a frame with a vertical scrollbar
Const SCROLL_H_HEIGHT           As Single = 18              ' Additional vertical space required for a frame with a horizontal scroll barr

Private lMonitorSteps           As Long
Private cllSteps                As Collection
Private lStepsDisplayed         As Long
Private tbxHeader               As MSForms.TextBox
Private tbxFooter               As MSForms.TextBox
Private bMonitorInitialized     As Boolean
Private MsgText1                As TypeMsgText  ' common text element
Private TextMonitorHeader       As TypeMsgText
Private TextMonitorFooter       As TypeMsgText
Private TextMonitorStep         As TypeMsgText
Private dctSectText             As New Dictionary
Private dctSectLabel            As New Dictionary

Public Property Get Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1) As TypeMsgText
' ------------------------------------------------------------------------------
' Returns the text (txt_kind_of_text) as section-text or -label, monitor-header,
' -footer, or -step.
' ------------------------------------------------------------------------------
    Dim vArry() As Variant
    Select Case txt_kind_of_text
        Case enMonHeader:    Text = TextMonitorHeader
        Case enMonFooter:    Text = TextMonitorFooter
        Case enMonStep:      Text = TextMonitorStep
        Case enSectText
            If dctSectText Is Nothing Then
                Text.Text = vbNullString
            ElseIf Not dctSectText.Exists(txt_section) Then
                Text.Text = vbNullString
            Else
                vArry = dctSectText(txt_section)
                Text.FontBold = vArry(0)
                Text.FontColor = vArry(1)
                Text.FontItalic = vArry(2)
                Text.FontName = vArry(3)
                Text.FontSize = vArry(4)
                Text.FontUnderline = vArry(5)
                Text.MonoSpaced = vArry(6)
                Text.Text = vArry(7)
            End If
    End Select
End Property

Public Property Let Text(Optional ByVal txt_kind_of_text As KindOfText, _
                         Optional ByVal txt_section As Long = 1, _
                                  ByRef txt_text As TypeMsgText)
' ------------------------------------------------------------------------------
' Provide the text (txt_text) as section (txt_section) text, section label,
' monitor header, footer, or step (txt_kind_of_text).
' ------------------------------------------------------------------------------
    Dim vArry(0 To 7)   As Variant
    
'    Dim t As TypeMsgText
'    t.FontBold = txt_text.FontBold
'    t.FontColor = txt_text.FontColor
'    t.FontItalic = txt_text.FontItalic
'    t.FontName = txt_text.FontName
'    t.FontSize = txt_text.FontSize
'    t.FontUnderline = txt_text.FontUnderline
'    t.MonoSpaced = txt_text.MonoSpaced
'    t.Text = txt_text.Text
    vArry(0) = txt_text.FontBold
    vArry(1) = txt_text.FontColor
    vArry(2) = txt_text.FontItalic
    vArry(3) = txt_text.FontName
    vArry(4) = txt_text.FontSize
    vArry(5) = txt_text.FontUnderline
    vArry(6) = txt_text.MonoSpaced
    vArry(7) = txt_text.Text
    Select Case txt_kind_of_text
        Case enMonHeader:    TextMonitorHeader = txt_text
        Case enMonFooter:    TextMonitorFooter = txt_text
        Case enMonStep:      TextMonitorStep = txt_text
        Case enSectText:     dctSectText.Add txt_section, vArry
    End Select

End Property

Public Property Get ContentHeight( _
                                 ByRef frm As MSForms.Frame, _
                        Optional ByVal applied As Boolean = False) As Single
' ------------------------------------------------------------------------------
' Returns the height of the frame's (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    Dim ctl As MSForms.Control
    
    For Each ctl In frm.Controls
        If ctl.Parent Is frm Then
'            If applied And IsApplied(ctl) Then
                ContentHeight = Max(ContentHeight, ctl.Top + ctl.Height)
'            Else
'                ContentHeight = Max(ContentHeight, ctl.Top + ctl.Height)
'            End If
        End If
    Next ctl

End Property

Public Property Get FrameContentWidth(Optional ByRef v As Variant, _
                                      Optional ByVal applied As Boolean = False) As Single
' ------------------------------------------------------------------------------
' Returns the maximum width of the frames (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    
    Dim ctl     As MSForms.Control
    Dim frm     As MSForms.Frame
    Dim frm_ctl As MSForms.Control
    
    If TypeName(v) = "Frame" Then Set frm_ctl = v Else Stop
    For Each ctl In frm_ctl.Controls
        If ctl.Parent Is frm_ctl Then
'            If applied And IsApplied(ctl) Then
                FrameContentWidth = Max(FrameContentWidth, ctl.Width + ctl.Left)
'            Else
'                FrameContentWidth = Max(FrameContentWidth, ctl.Width - ctl.Left)
'            End If
        End If
    Next ctl
    
End Property

Public Property Let FrameHeight( _
                 Optional ByRef frm As MSForms.Frame, _
                          ByVal frm_height As Single)
' ------------------------------------------------------------------------------
' Mimics a frame's height change event. When the height of the frame (frm) is
' changed (frm_height) to less than the frame's content height and no vertical
' scrollbar is applied one is applied with the frame content's height. If one
' is already applied just the height is adjusted to the frame content's height.
' When the height becomes more than the frame's content height a vertical
' scrollbar becomes obsolete and is removed.
' ------------------------------------------------------------------------------
    Dim yAction     As fmScrollAction
    Dim siHeight    As Single
    
    siHeight = ContentHeight(frm)
    
'    If UsageType = usage_progress_display Then yAction = fmScrollActionEnd Else yAction = fmScrollActionBegin
    frm.Height = frm_height
    If frm.Height < siHeight Then
        '~~ Apply a vertical scrollbar if none is applied yet, adjust its height otherwise
        If Not ScrollV_Applied(frm) _
        Then ScrollV_Apply frm, siHeight, yAction _
        Else frm.ScrollHeight = siHeight
    Else
        '~~ With the frame's height is greater or equal its content height
        '~~ a vertical scrollbar becomes obsolete and is removed
        With frm
            Select Case .ScrollBars
                Case fmScrollBarsBoth:      .ScrollBars = fmScrollBarsHorizontal
                Case fmScrollBarsVertical:  .ScrollBars = fmScrollBarsNone
            End Select
        End With
    End If
End Property

Public Property Let FrameWidth( _
                 Optional ByRef frm As MSForms.Frame, _
                          ByVal frm_width As Single)
' ------------------------------------------------------------------------------
' Mimics a frame's width change event. When the width of the frame (frm) is
' changed (frm_width) a horizontal scrollbar will be applied - or adjusted to
' the frame content's width. I.e. this property must only be used when a
' horizontal scrollbar is applicable/desired in case.
' ------------------------------------------------------------------------------
    Dim ContentWidth As Single: ContentWidth = FrameContentWidth(frm)
    
    frm.Width = frm_width
    If frm_width < ContentWidth Then
        '~~ Apply a horizontal scrollbar if none is applied yet, adjust its width otherwise
        If Not ScrollH_Applied(frm) _
        Then ScrollH_Apply frm, ContentWidth _
        Else frm.ScrollWidth = ContentWidth
    Else
        '~~ With the frame's width greater or equal its content width
        '~~ a horizontal scrollbar becomes obsolete and is removed
        With frm
            Select Case .ScrollBars
                Case fmScrollBarsBoth:          .ScrollBars = fmScrollBarsVertical
                Case fmScrollBarsHorizontal:    .ScrollBars = fmScrollBarsNone
            End Select
        End With
    End If
    
End Property

Private Property Get ScrollBarHeight(Optional ByVal frm As MSForms.Frame) As Single
    If frm.ScrollBars = fmScrollBarsBoth Or frm.ScrollBars = fmScrollBarsHorizontal Then ScrollBarHeight = 14
End Property

Private Property Get ScrollBarWidth(Optional ByVal frm As MSForms.Frame) As Single
    If frm.ScrollBars = fmScrollBarsBoth Or frm.ScrollBars = fmScrollBarsVertical Then ScrollBarWidth = 12
End Property

Private Property Get TitleLengthFactor() As Single
    TitleLengthFactor = CSng(Me.tbxFactor.Value)
End Property

Public Property Get VspaceFrame(Optional frm As MSForms.Frame) As Single
    If frm.Caption = vbNullString Then VspaceFrame = 4 Else VspaceFrame = 8

End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Sub AutoSizeTextBox( _
                     ByRef as_tbx As MSForms.TextBox, _
                     ByVal as_text As String, _
            Optional ByVal as_width_limit As Single = 0, _
            Optional ByVal as_width_min As Single = 0, _
            Optional ByVal as_height_min As Single = 0, _
            Optional ByVal as_width_max As Single = 0, _
            Optional ByVal as_height_max As Single = 0, _
            Optional ByVal as_append As Boolean = False, _
            Optional ByVal as_append_margin As String = vbNullString)
' ------------------------------------------------------------------------------
' Common AutoSize service for an MsForms.TextBox providing a width and height
' for the TextBox (as_tbx) by considering:
' - When a width limit is provided (as_width_limit > 0) the width is regarded a
'   fixed maximum and thus the height is auto-sized by means of WordWrap=True.
' - When no width limit is provided (the default) WordWrap=False and thus the
'   width of the TextBox is determined by the longest line.
' - When a maximum width is provided (as_width_max > 0) and the parent of the
'   TextBox is a frame a horizontal scrollbar is applied for the parent frame.
' - When a maximum height is provided (as_heightmax > 0) and the parent of the
'   TextBox is a frame a vertical scrollbar is applied for the parent frame.
' - When a minimum width (as_width_min > 0) or a minimum height (as_height_min
'   > 0) is provided the size of the textbox is set correspondingly. This
'   option is specifically usefull when text is appended to avoid much flicker.
'
' Uses: FrameWidth, FrameContentWidth, ScrollH_Apply,
'       FrameHeight, ContentHeight, ScrollV_Apply
'
' W. Rauschenberger Berlin June 2021
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            .WordWrap = True
            .AutoSize = False
            .Width = as_width_limit - 7 ' the readability space is added later
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & as_append_margin & vbLf & as_text
                End If
            End If
            .AutoSize = True
        Else
            .MultiLine = True
            .WordWrap = False ' the means to limit the width
            .AutoSize = True
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & vbLf & as_text
                End If
            End If
        End If
        .Width = .Width + 7   ' readability space
        .Height = .Height + 7 ' redability space
        If as_width_min > 0 And .Width < as_width_min Then .Width = as_width_min
        If as_height_min > 0 And .Height < as_height_min Then .Height = as_height_min
        .Parent.Height = .Top + .Height + 2
        .Parent.Width = .Left + .Width + 2
    End With
    
    '~~ When the parent of the TextBox is a frame scrollbars may have become applicable
    '~~ provided a mximimum with and/or height has been provided
    If TypeName(tbx.Parent) = "Frame" Then
        '~~ When a max width is provided and exceeded a horizontal scrollbar is applied
        If as_width_max > 0 Then
            FrameWidth(as_tbx.Parent) = Min(as_width_max, tbx.Width + 2 + ScrollBarWidth(as_tbx.Parent))
        End If
        '~~ When a max height is provided and exceeded a vertical scrollbar is applied
        If as_height_max > 0 Then
            FrameHeight(as_tbx.Parent) = Min(as_height_max, tbx.Height + ScrollBarHeight(as_tbx.Parent))
        End If
    End If
    
xt: Exit Sub

End Sub

Private Sub cbSetupTitle_Click()
    Dim FACTOR As Single
    Dim MinWidth    As Single
    Dim MaxWidth    As Single
    
    MinWidth = 100
    MaxWidth = 1500
    FACTOR = 1.1
    Setup1_Title setup_title:=Me.tbxTestTitle & " " & Format(FACTOR, "0.000") _
               , setup_width_min:=MinWidth _
               , setup_width_max:=MaxWidth
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option
' (Conditional Compile Argument 'Debugging = 1') and an optional additional
' "about the error" information which may be connected to an error message by
' two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case vbPassOn:  Err.Raise Err.Number, ErrSrc(PROC), Err.Description
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsgProcTest." & sProc
End Function

Public Sub MonitorInitialize(ByVal mon_title As String, _
                             ByVal mon_steps_displayed As Long, _
                    Optional ByVal mon_header As String = vbNullString, _
                    Optional ByVal mon_footer As String = vbNullString)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim ctl     As MSForms.Control
    Dim siTop   As Single
    Dim i       As Long
    Dim tbx     As MSForms.TextBox
    
    If Not bMonitorInitialized Then
        lMonitorSteps = mon_steps_displayed
        For Each ctl In Me.Controls
            ctl.Visible = False
        Next ctl
        siTop = 6
        
        With Me
            .Caption = mon_title
            If mon_header <> vbNullString Then
                '~~ Initialize Header
                Set tbxHeader = .Controls.Add("Forms.TextBox.1")
                With tbxHeader
                    .Top = siTop
                    .Left = 0
                    .Value = mon_header
                    .Height = 18
                    .Width = Me.InsideWidth
                    siTop = .Top + .Height
                    .Font.Bold = True
                    .ForeColor = rgbBlue
                    .BackColor = Me.BackColor
                    .BorderColor = Me.BackColor
                    .BorderStyle = fmBorderStyleSingle
                End With
            End If
        
            For i = 1 To lMonitorSteps
                '~~ Initialize Steps
                Set tbx = .Controls.Add("Forms.TextBox.1")
                With tbx
                    .Top = siTop
                    .Left = 0
                    .Visible = False
                    .Height = 18
                    .Width = Me.InsideWidth
                    siTop = .Top + .Height
                    .BackColor = Me.BackColor
                    .BorderColor = Me.BackColor
                    .BorderStyle = fmBorderStyleSingle
               End With
                Qenqueue cllSteps, tbx
            Next i
            
            .Height = siTop + 35
        
            If mon_footer <> vbNullString Then
                '~~ Initialize Footer
                Set tbxFooter = .Controls.Add("Forms.TextBox.1")
                With tbxFooter
                    .Visible = True
                    .Top = siTop + 6
                    .Left = 0
                    .Value = mon_footer
                    .Height = 18
                    .Width = Me.InsideWidth
                    .Font.Bold = True
                    .ForeColor = rgbBlue
                    .BackColor = Me.BackColor
                    .BorderColor = Me.BackColor
                    .BorderStyle = fmBorderStyleSingle
                    Me.Height = .Top + .Height + 35
                End With
            End If
        End With
        bMonitorInitialized = True
    End If

End Sub

Public Sub MonitorStep(Optional ByVal mon_step As String = vbNullString, _
                       Optional ByVal mon_footer As String = vbNullString)
' ------------------------------------------------------------------------------
' Display a step
' ------------------------------------------------------------------------------
    Dim tbx     As MSForms.TextBox
    Dim i       As Long
    Dim lTop    As Long
    
    If mon_step <> vbNullString Then
        If lStepsDisplayed < lMonitorSteps Then
            Set tbx = cllSteps(lStepsDisplayed + 1)
            With tbx
                .Visible = True
                .Value = mon_step
                lStepsDisplayed = lStepsDisplayed + 1
            End With
        Else ' All steps arte displayed
            Set tbx = Qdequeue(cllSteps)
            tbx.Value = vbNullString
            Qenqueue cllSteps, tbx
            If tbxHeader Is Nothing Then
                lTop = 6
            Else
                With tbxHeader
                    lTop = .Top + .Height
                End With
            End If
            
            For i = 1 To lMonitorSteps
                Set tbx = cllSteps(i)
                tbx.Top = lTop + (18 * (i - 1))
            Next i
            tbx.Value = mon_step
        End If
    End If
    
    If mon_footer <> vbNullString Then
        If Not tbxFooter Is Nothing Then tbxFooter.Value = mon_footer
    End If

End Sub

Private Function Qdequeue(ByRef qu As Collection) As Variant
    Const PROC = "DeQueue"
    
    On Error GoTo eh
    If qu Is Nothing Then GoTo xt
    If QisEmpty(qu) Then GoTo xt
    On Error Resume Next
    Set Qdequeue = qu(1)
    If Err.Number <> 0 _
    Then Qdequeue = qu(1)
    qu.Remove 1

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub Qenqueue(ByRef qu As Collection, ByVal qu_item As Variant)
    If qu Is Nothing Then Set qu = New Collection
    qu.Add qu_item
End Sub

Private Function QisEmpty(ByVal qu As Collection) As Boolean
    If Not qu Is Nothing _
    Then QisEmpty = qu.Count = 0 _
    Else QisEmpty = True
End Function

Private Function QLen(ByVal qu As Collection) As Long
    If qu Is Nothing Then Set qu = New Collection
    QLen = qu.Count
End Function

Private Function Qrequeue(ByRef qu As Collection) As Variant
' ------------------------------------------------------------------------------
' Deques the first item and Enqueues it again.
' ------------------------------------------------------------------------------
    Qenqueue qu, Qdequeue(qu)
End Function

Private Function ScrollH_Applied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a horizontal scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollH_Applied = True
    End Select
End Function

Private Sub ScrollH_Apply( _
                            ByRef scroll_frame As MSForms.Frame, _
                            ByVal scrolled_width As Single, _
                   Optional ByVal x_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' Apply a horizontal scrollbar is applied to the frame (scroll_frame) and
' adjusted to the frame content's width (scrolled_width). In case a horizontal
' scrollbar is already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollH_Apply"
    
    On Error GoTo eh
        
    With scroll_frame
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
                .Height = ContentHeight(scroll_frame) + SCROLL_H_HEIGHT
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsHorizontal
                .KeepScrollBarsVisible = fmScrollBarsHorizontal
                .ScrollWidth = scrolled_width
                .Scroll xAction:=x_action
                If .Height < ContentHeight(scroll_frame) + SCROLL_H_HEIGHT Then
                    .Height = ContentHeight(scroll_frame) + SCROLL_H_HEIGHT
                End If
        End Select
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function ScrollV_Applied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a vertical scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollV_Applied = True
    End Select
End Function

Private Sub ScrollV_Apply( _
                          ByRef scroll_frame As MSForms.Frame, _
                          ByVal scrolled_height As Single, _
                 Optional ByVal y_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' A vertical scrollbar is applied to the frame (scroll_frame) and adjusted to
' the frame content's height (scrolled_height). In case a vertical scrollbar is
' already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollV_Apply"
    
    On Error GoTo eh
        
    With scroll_frame
        Select Case .ScrollBars
            Case fmScrollBarsBoth
                '~~ The already displayed horizonzal scrollbar's width is adjusted
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
            Case fmScrollBarsHorizontal
                '~~ Already displayed (no vertical scrollbar yet)
                '~~ No need to adjust the height for the scrollbar
                .ScrollBars = fmScrollBarsBoth
                .KeepScrollBarsVisible = fmScrollBarsBoth
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
'                .Width = FrameContentWidth(scroll_frame) + SCROLL_V_WIDTH
            Case fmScrollBarsVertical
                '~~ Add a horizontal scrollbar to the already displayed vertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
            Case fmScrollBarsNone
                '~~ Add a horizontal scrollbar
                .ScrollBars = fmScrollBarsVertical
                .KeepScrollBarsVisible = fmScrollBarsVertical
                .ScrollHeight = scrolled_height
                .Scroll yAction:=y_action
                .Width = FrameContentWidth(scroll_frame) + SCROLL_V_WIDTH
        End Select
    End With

xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Sub Setup1_Title( _
                ByVal setup_title As String, _
                ByVal setup_width_min As Single, _
                ByVal setup_width_max As Single)
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (setup_title) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_width_max) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' ------------------------------------------------------------------------------
    Const PROC = "Setup1_Title"
    
    On Error GoTo eh
    Dim Correction    As Single
    
    With Me
        .Width = setup_width_min
        '~~ The extra title label is only used to adjust the form width and remains hidden
        With .laMsgTitle
            With .Font
                .Bold = False
                .Name = Me.Font.Name
                .Size = 8    ' Value which comes to a length close to the length required
            End With
            .Caption = vbNullString
            .AutoSize = True
            .Caption = " " & setup_title    ' some left margin
        End With
        .Caption = setup_title
        Correction = (CInt(.laMsgTitle.Width)) / 1700
'        Debug.Print ".laMsgTitle.Width: " & .laMsgTitle.Width, "Factor: " & FACTOR, "FactorCorrection: " & FactorCorrection
        .Width = Min(setup_width_max, .laMsgTitle.Width * (TitleLengthFactor - Correction))
    End With
   
xt: Exit Sub
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

