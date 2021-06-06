VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fProcTest 
   Caption         =   "Test-Msg-Form"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
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
                                 ByRef frm As MsForms.Frame, _
                        Optional ByVal applied As Boolean = False) As Single
' ------------------------------------------------------------------------------
' Returns the height of the Frames (frm) content by considering only
' applied/visible controls.
' ------------------------------------------------------------------------------
    Dim ctl As MsForms.Control
    
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
    
    Dim ctl As MsForms.Control
    Dim frm As MsForms.Frame
    Dim frm_ctl As MsForms.Control
    
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
                 Optional ByRef frm As MsForms.Frame, _
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
                 Optional ByRef frm As MsForms.Frame, _
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

Private Function ScrollHorizontalApplied(ByRef frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a horizontal scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsHorizontal: ScrollHorizontalApplied = True
    End Select
End Function

Private Sub ScrollHorizontalApply( _
                            ByRef scroll_frame As MsForms.Frame, _
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

Private Function ScrollVerticalApplied(ByRef frm As MsForms.Frame) As Boolean
' ------------------------------------------------------------------------------
' Returns True when the frame (frm) has already a vertical scrollbar applied.
' ------------------------------------------------------------------------------
    Select Case frm.ScrollBars
        Case fmScrollBarsBoth, fmScrollBarsVertical: ScrollVerticalApplied = True
    End Select
End Function

Private Sub ScrollVerticalApply( _
                          ByRef scroll_frame As MsForms.Frame, _
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

Public Sub SizeWidthAndHeight( _
                        ByRef as_tbx As MsForms.TextBox, _
                        ByVal as_text As String, _
               Optional ByVal as_width As Single = 0, _
               Optional ByVal as_height As Single = 0, _
               Optional ByVal as_append As Boolean = False)
' ------------------------------------------------------------------------------
' Determines the width and height of the TextBox (tbx).
' - When a width is provided (as_width > 0) the width is regarded a fixed
'   maximum and thus the height is auto-sized by means of WordWrap = True. When
'   the width is exceeded a horizontal scrollbar becomes applicable for the
'   parent frame.
' - When no width is provided the width of the TextBox is determined by the
'   longest line and consequently WordWrap = False
' - When a height is provided (as_min_height > 0) the height is regarded fixed.
'   I.e. when it is exceeded a vertical scrollbar becomes applicable for the
'   parent frame.
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width > 0 Then .WordWrap = True Else .WordWrap = False ' the means to limit the width
        If as_width > 0 Then .AutoSize = False Else .AutoSize = True
        
        If as_width > 0 Then .Width = as_width
        
        If Not as_append Then
            .Value = as_text
        Else
            If .Value = vbNullString Then
                .Value = as_text
            Else
                .Value = .Value & vbLf & as_text
            End If
        End If
        If as_width > 0 Then
            .Width = as_width
            .AutoSize = True
        End If
        If as_height > 0 And .Height < as_height Then .Height = as_height
        
    End With
End Sub

Public Property Get VspaceFrame(Optional frm As MsForms.Frame) As Single
    If frm.Caption = vbNullString Then VspaceFrame = 4 Else VspaceFrame = 8

End Property
