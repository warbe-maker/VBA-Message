VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fProcTest 
   Caption         =   "Test-Msg-Form"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
   OleObjectBlob   =   "fProcTest.frx":0000
End
Attribute VB_Name = "fProcTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const VSCROLLBAR_WIDTH              As Single = 18              ' Additional horizontal space required for a frame with a vertical scrollbar
Const HSCROLLBAR_HEIGHT             As Single = 18              ' Additional vertical space required for a frame with a horizontal scroll barr

Public Property Get FrameContentHeight( _
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
                FrameContentHeight = Max(FrameContentHeight, ctl.Top + ctl.Height)
'            Else
'                FrameContentHeight = Max(FrameContentHeight, ctl.Top + ctl.Height)
'            End If
        End If
    Next ctl

End Property

Public Property Get FrameContentWidth( _
                       Optional ByRef v As Variant, _
                       Optional ByVal applied As Boolean = False) As Single
' ------------------------------------------------------------------------------
' Returns the maximum width of the frames (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    
    Dim ctl As MSForms.Control
    Dim frm As MSForms.Frame
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
    Dim yAction         As fmScrollAction
    Dim ContentHeight   As Single:          ContentHeight = FrameContentHeight(frm)
    
'    If UsageType = usage_progress_display Then yAction = fmScrollActionEnd Else yAction = fmScrollActionBegin
    frm.Height = frm_height
    If frm.Height < ContentHeight Then
        '~~ Apply a vertical scrollbar if none is applied yet, adjust its height otherwise
        If Not ScrollVerticalApplied(frm) _
        Then ScrollVerticalApply frm, ContentHeight, yAction _
        Else frm.ScrollHeight = ContentHeight
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
        If Not ScrollHorizontalApplied(frm) _
        Then ScrollHorizontalApply frm, ContentWidth _
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

Public Property Get VspaceFrame(Optional frm As MSForms.Frame) As Single
    If frm.Caption = vbNullString Then VspaceFrame = 4 Else VspaceFrame = 8

End Property

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
' Uses: FrameWidth, FrameContentWidth, ScrollHorizontalApply,
'       FrameHeight, FrameContentHeight, ScrollVerticalApply
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

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString, _
    Optional ByVal err_line As Long = 0)
' ------------------------------------------------------------------------------
' This 'Common VBA Component' uses only a kind of minimum error handling!
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim errline As Long
    Dim AtLine  As String
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    If err_dscrptn = vbNullString Then ErrDesc = Err.Description Else ErrDesc = err_dscrptn
    If err_line = 0 Then errline = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    MsgBox Title:=ErrType & ErrNo & " in " & err_source _
         , Prompt:="Error : " & ErrDesc & vbLf & _
                   "Source: " & err_source & AtLine _
         , Buttons:=vbCritical
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "fMsg." & sProc
End Function

Private Function ScrollHorizontalApplied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a horizontal scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollHorizontalApplied = True
    End Select
End Function

Private Sub ScrollHorizontalApply( _
                            ByRef scroll_frame As MSForms.Frame, _
                            ByVal scrolled_width As Single, _
                   Optional ByVal x_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' Apply a horizontal scrollbar is applied to the frame (scroll_frame) and
' adjusted to the frame content's width (scrolled_width). In case a horizontal
' scrollbar is already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollHorizontalApply"
    
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
                .Height = FrameContentHeight(scroll_frame) + HSCROLLBAR_HEIGHT
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
                If .Height < FrameContentHeight(scroll_frame) + HSCROLLBAR_HEIGHT Then
                    .Height = FrameContentHeight(scroll_frame) + HSCROLLBAR_HEIGHT
                End If
        End Select
    End With

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Function ScrollVerticalApplied(ByRef frm As MSForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a vertical scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollVerticalApplied = True
    End Select
End Function

Private Sub ScrollVerticalApply( _
                          ByRef scroll_frame As MSForms.Frame, _
                          ByVal scrolled_height As Single, _
                 Optional ByVal y_action As fmScrollAction = fmScrollActionBegin)
' ------------------------------------------------------------------------------
' A vertical scrollbar is applied to the frame (scroll_frame) and adjusted to
' the frame content's height (scrolled_height). In case a vertical scrollbar is
' already applied only its width is adjusted.
' ------------------------------------------------------------------------------
    Const PROC = "ScrollVerticalApply"
    
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
'                .Width = FrameContentWidth(scroll_frame) + VSCROLLBAR_WIDTH
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
                .Width = FrameContentWidth(scroll_frame) + VSCROLLBAR_WIDTH
        End Select
    End With

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Public Sub Setup1_Title( _
                ByVal setup_title As String, _
                ByVal setup_min_width As Single, _
                ByVal setup_max_width As Single)
' ------------------------------------------------------------------------------
' Setup the message form for the provided title (setup_title) optimized with the
' provided minimum width (setup_width_min) and the provided maximum width
' (setup_max_width) by using a certain factor (setup_factor) for the calculation
' of the width required to display an untruncated title - as long as the maximum
' widht is not exeeded.
' ------------------------------------------------------------------------------
    Const PROC = "Setup1_Title"
    Const FACTOR = 1.45
    
    On Error GoTo eh
    Dim Correction    As Single
    
    With Me
        .Width = setup_min_width
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
        .Width = Min(setup_max_width, .laMsgTitle.Width * (FACTOR - Correction))
    End With
   
xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Sub cbSetupTitle_Click()
    Dim FACTOR As Single
    Dim MinWidth    As Single
    Dim MaxWidth    As Single
    
    MinWidth = 100
    MaxWidth = 1500
    FACTOR = 1.1
    Setup1_Title setup_title:=Me.tbxTestTitle & " " & Format(FACTOR, "0.000") _
               , setup_min_width:=MinWidth _
               , setup_max_width:=MaxWidth
End Sub
