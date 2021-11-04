Attribute VB_Name = "mFuncTest"
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------------
' Standard Module mTest
'          All tests for a complete regression test. Obligatory performed after
'          any modification. Ammended when new features or functions are
'          implemented.
'
' Note:    Errors raised by the tested procedures cannot be asserted since they
'          are not passed on to the calling/entry procedure. This would require
'          the Common Standard Error Handling Module mErH which intentionally is
'          not used by this module.
'
' W. Rauschenberger, Berlin June 2020
' -------------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Public Const BTTN_FINISH        As String = "Test Done"
Public Const BTTN_PASSED        As String = "Passed"
Public Const BTTN_FAILED        As String = "Failed"

Dim TestMsgWidthMin         As Long
Dim TestMsgWidthMax         As Long
Dim TestMsgHeightMin        As Long
Dim TestMsgHeightMax        As Long
Dim bRegressionTest         As Boolean
Dim TestMsgHeightIncrDecr   As Long
Dim TestMsgWidthIncrDecr    As Long
Dim Message                 As TypeMsg
Dim sBttnTerminate          As String
Dim vButton4                As Variant
Dim vButton5                As Variant
Dim vButton6                As Variant
Dim vButton7                As Variant
Dim vbuttons                As Variant

Public Property Get BTTN_TERMINATE() As String ' composed constant
    BTTN_TERMINATE = "Terminate" & vbLf & "Regression"
End Property

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Displays a proper designe error message providing the option to resume the
' error line when the Conditional Compile Argument Debugging = 1.
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim errline As Long
    Dim AtLine  As String
    Dim Buttons As Long
    Dim msg     As TypeMsg
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Application error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    
    If err_line = 0 Then errline = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    With msg.Section(1)
        .Label.Text = "Error:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_dscrptn
    End With
    With msg.Section(2)
        .Label.Text = "Source:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_source & AtLine
    End With

#If Debugging Then
    Buttons = vbYesNo
    With msg.Section(3)
        .Label.Text = "Debugging: (Conditional Compile Argument 'Debugging = 1')"
        .Label.FontColor = rgbBlue
        .Text.Text = "Yes = Resume error line, No = Continue"
    End With
    With msg.Section(4)
        .Label.Text = "About debugging:"
        .Label.FontColor = rgbBlue
        .Text.Text = "To make use of the debugging option have an error handling line" & vbLf & _
                     "eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume"
    End With
#Else
    Buttons = vbCritical
#End If
    
    ErrMsg = Dsply(dsply_title:=ErrType & ErrNo & " in " & err_source & AtLine _
                 , dsply_msg:=msg _
                 , dsply_buttons:=Buttons)
End Function

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mFuncTest." & s:  End Property

Public Property Let RegressionTest(ByVal b As Boolean)
    bRegressionTest = b
    If b Then sBttnTerminate = "Terminate" & vbLf & "Regression" Else sBttnTerminate = vbNullString
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no > 0 Then AppErr = app_err_no + vbObjectError Else AppErr = app_err_no - vbObjectError
End Function

Public Sub Test_Regression()
' --------------------------------------------------------------------------------------
' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
    Const PROC = "Test_Regression"
    
    On Error GoTo eh
    Dim rng     As Range
    Dim sTest   As String
    Dim sMakro  As String
    
    ThisWorkbook.Save
    Unload fMsg
    wsTest.RegressionTest = True
    mFuncTest.RegressionTest = True
    
    For Each rng In wsTest.RegressionTests
        If rng.Value = "R" Then
            sTest = Format(rng.Offset(, -2), "00")
            sMakro = "cmdTest" & sTest & "_Click"
            wsTest.TerminateRegressionTest = False
            Application.Run "Msg.xlsb!" & sMakro
            If wsTest.TerminateRegressionTest Then Exit For
        End If
    Next rng

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Sub ClearVBEImmediateWindow()
    Dim v   As Variant
    For Each v In Application.VBE.Windows
        If v.Caption = "Direktbereich" Then
            v.SetFocus
            Application.SendKeys "^g ^a {DEL}"
            DoEvents
            Exit Sub
        End If
    Next v
End Sub

Public Sub cmdTest00_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 0
    mFuncTest.Test_00_ErrMsg
End Sub

Public Sub cmdTest01_Click()
' ------------------------------------------------------------------------------
' Procedures for test start via Command Buttons on Test Worksheet
' ------------------------------------------------------------------------------
'    wsTest.RegressionTest = False
    wsTest.TestNumber = 1
    mFuncTest.Test_01_WidthDeterminedByMinimumWidth
End Sub

Public Sub cmdTest02_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 2
    mFuncTest.Test_02_WidthDeterminedByTitle
End Sub

Public Sub cmdTest03_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 3
    mFuncTest.Test_03_WidthDeterminedByMonoSpacedMessageSection
End Sub

Public Sub cmdTest04_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 4
    mFuncTest.Test_04_WidthDeterminedByReplyButtons
End Sub

Public Sub cmdTest05_Click()
    wsTest.RegressionTest = False
    mFuncTest.Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth
End Sub

Public Sub cmdTest06_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 6
    mFuncTest.Test_06_MonoSpacedMessageSectionExceedsMaxHeight
End Sub

Public Sub cmdTest07_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 7
    mFuncTest.Test_07_ButtonsOnly
End Sub

Public Sub cmdTest08_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 8
    mFuncTest.Test_08_ButtonsMatrix
End Sub

Public Sub cmdTest09_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 9
    mFuncTest.Test_09_ButtonScrollBarVertical
End Sub

Public Sub cmdTest10_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 10
    mFuncTest.Test_10_ButtonScrollBarHorizontal
End Sub

Public Sub cmdTest11_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 11
    mFuncTest.Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
End Sub

Public Sub cmdTest17_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 17
    mFuncTest.Test_17_MessageAsString
End Sub

Public Sub cmdTest30_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 30
    mFuncTest.Test_30_Monitor
End Sub

Public Sub cmdTest90_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 90
    mFuncTest.Test_90_All_in_one_Demonstration
End Sub

Public Sub Test_00_ErrMsg()
    Const PROC = "Test_00_ErrMsg"
    
    On Error GoTo eh
    Dim i As Long
    
    wsTest.TestNumber = 0
    
    i = i / 0
    
xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume

    Select Case mMsg.Box(box_title:="Test result of " & Readable(PROC) _
                       , box_msg:=vbNullString _
                       , box_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED) _
                        )
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

End Sub

Public Sub Explore(ByVal ctl As Variant, _
          Optional ByVal applied As Boolean = True)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Explore"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle   As String
    Dim dct         As New Dictionary
    Dim v           As Variant
    Dim Appl        As String   ' ControlApplied
    Dim l           As String   ' .Left
    Dim W           As String   ' .Width
    Dim T           As String   ' .Top
    Dim H           As String   ' .Height
    Dim SW          As String   ' .ScrollWidth
    Dim SH          As String   ' .ScrollHeight
    Dim FW          As String   ' MsgForm.InsideWidth
    Dim CW          As String   ' Content width
    Dim CH          As String   ' Content height
    Dim FH          As String   ' MsgForm.InsideHeight
    Dim i           As Long
    Dim Item        As String
    Dim j           As String
    Dim frm         As MSForms.Frame
    
    MsgTitle = "Explore"
    Unload mMsg.Form(MsgTitle) ' Ensure there is no process monitoring with this title still displayed
    Set MsgForm = mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC))
    
    If TypeName(ctl) <> "Frame" And TypeName(ctl) <> "fMsg" Then Exit Sub
    
    '~~ Collect Controls
    mDct.DctAdd dct, ctl, ctl.Name, order_byitem, seq_ascending, sense_casesensitive
      
    i = 0: j = 1
    Do
        If TypeName(dct.Keys()(i)) = "Frame" Or TypeName(dct.Keys()(i)) = "fMsg" Then
            For Each v In dct.Keys()(i).Controls
                If v.Parent Is dct.Keys()(i) Then
                    Item = dct.Items()(i) & ":" & v.Name
                    If applied Then
                        If MsgForm.IsApplied(v) Then mDct.DctAdd dct, v, Item
                    Else
                        mDct.DctAdd dct, v, Item
                    End If
                End If
            Next v
        End If
        If TypeName(dct.Keys()(i)) = "Frame" Or TypeName(dct.Keys()(i)) = "fMsg" Then j = j + 1
        If i + 1 < dct.Count Then i = i + 1 Else Exit Do
    Loop
        
    '~~ Display facts
    Debug.Print "====================+====+=======+=======+=======+=======+=======+=======+=======+=======+=======+======="
    Debug.Print "                    |Ctl | Left  | Width |Content| Top   |Height |Content|VScroll|HScroll| Width | Height"
    Debug.Print "Name                |Appl| Pos   |       | Width | Pos   |       |Height |Height | Width | Form  |  Form "
    Debug.Print "--------------------+----+-------+-------+-------+-------+-------+-------+-------+-------+-------+-------"
    For Each v In dct
        Set ctl = v
        If MsgForm.IsApplied(ctl) Then Appl = "Yes " Else Appl = " No "
        l = Align(Format(ctl.Left, "000.0"), 7, AlignCentered, " ")
        W = Align(Format(ctl.Width, "000.0"), 7, AlignCentered, " ")
        T = Align(Format(ctl.top, "000.0"), 7, AlignCentered, " ")
        H = Align(Format(ctl.Height, "000.0"), 7, AlignCentered, " ")
        FH = Align(Format(MsgForm.InsideHeight, "000.0"), 7, AlignCentered, " ")
        FW = Align(Format(MsgForm.InsideWidth, "000.0"), 7, AlignCentered, " ")
        If TypeName(ctl) = "Frame" Then
            Set frm = ctl
            CW = Align(Format(MsgForm.FrameContentWidth(frm), "000.0"), 7, AlignCentered, " ")
            CH = Align(Format(MsgForm.FrameContentHeight(frm), "000.0"), 7, AlignCentered, " ")
            SW = "   -   "
            SH = "   -   "
            With frm
                Select Case .ScrollBars
                    Case fmScrollBarsHorizontal
                        Select Case .KeepScrollBarsVisible
                            Case fmScrollBarsBoth, fmScrollBarsHorizontal
                                SW = Align(Format(.ScrollWidth, "000.0"), 7, AlignCentered, " ")
                        End Select
                    Case fmScrollBarsVertical
                        Select Case .KeepScrollBarsVisible
                            Case fmScrollBarsBoth, fmScrollBarsVertical
                                SH = Align(Format(.ScrollHeight, "000.0"), 7, AlignCentered, " ")
                        End Select
                    Case fmScrollBarsBoth
                        Select Case .KeepScrollBarsVisible
                            Case fmScrollBarsBoth
                                SW = Align(Format(.ScrollWidth, "000.0"), 7, AlignCentered, " ")
                                SH = Align(Format(.ScrollHeight, "000.0"), 7, AlignCentered, " ")
                            Case fmScrollBarsVertical
                                SH = Align(Format(.ScrollHeight, "000.0"), 7, AlignCentered, " ")
                            Case fmScrollBarsHorizontal
                                SW = Align(Format(.ScrollWidth, "000.0"), 7, AlignCentered, " ")
                        End Select
                End Select
            End With
        End If
        
        Debug.Print Align(ctl.Name, 20, AlignLeft) & "|" & Appl & "|" & l & "|" & W & "|" & CW & "|" & T & "|" & H & "|" & CH & "|" & SH & "|" & SW & "|" & FW & "|" & FH
    Next v

xt: Set dct = Nothing

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Function IsUcase(ByVal s As String) As Boolean

    Dim i   As Integer: i = Asc(s)
    IsUcase = (i >= 65 And i <= 90) Or _
              (i >= 192 And i <= 214) Or _
              (i >= 216 And i <= 223) Or _
              (i = 128) Or _
              (i = 138) Or _
              (i = 140) Or _
              (i = 142) Or _
              (i = 154) Or _
              (i = 156) Or _
              (i >= 158 And i <= 159) Or _
              (i = 163) Or _
              (i = 165)
End Function

Private Sub MessageInit(ByRef msg_form As fMsg, _
                        ByVal msg_title As String, _
               Optional ByVal caller As String = vbNullString)
' ------------------------------------------------------------------------------
' Initializes the all message sections with the defaults throughout this test
' module which uses a module global declared Message for a consistent layout.
' ------------------------------------------------------------------------------
    Dim i           As Long
    
    mMsg.Form frm_caption:=msg_title, frm_unload:=True                    ' Ensures a message starts from scratch
    Set msg_form = mMsg.Form(frm_caption:=msg_title, frm_caller:=caller)
    
    For i = 1 To msg_form.NoOfDesignedMsgSects
        With Message.Section(i)
            .Label.Text = vbNullString
            .Label.FontColor = rgbBlue
            .Text.Text = vbNullString
            .Text.MonoSpaced = False
            .Text.FontItalic = False
            .Text.FontUnderline = False
            .Text.FontColor = rgbBlack
        End With
    Next i
    If bRegressionTest Then mFuncTest.RegressionTest = True Else mFuncTest.RegressionTest = False

End Sub

Private Function Readable(ByVal s As String) As String
' ------------------------------------------------------------------------------
' Convert a string (s) into a readable form by replacing all underscores
' with a whitespace and all characters immediately following an underscore
' to a lowercase letter.
' ------------------------------------------------------------------------------
    Dim i       As Long
    Dim sResult As String
    
    s = Replace(s, "_", " ")
    s = Replace(s, "  ", " ")
    For i = 1 To Len(s)
        If IsUcase(Mid(s, i, 1)) Then
            sResult = sResult & " " & Mid(s, i, 1)
        Else
            sResult = sResult & Mid(s, i, 1)
        End If
    Next i
    Readable = Right(sResult, Len(sResult) - 1)

End Function

Private Function Repeat(repeat_string As String, repeat_n_times As Long)
    Dim s As String
    Dim C As Long
    Dim l As Long
    Dim i As Long

    l = Len(repeat_string)
    C = l * repeat_n_times
    s = Space$(C)

    For i = 1 To C Step l
        Mid(s, i, l) = repeat_string
    Next

    Repeat = s
End Function

Private Function RepeatString( _
           ByVal rep_n_times As Long, _
           ByVal rep_pattern As String, _
  Optional ByVal rep_with_line_numbers As Boolean = False, _
  Optional ByVal rep_with_linen_umbers_as_prefix As Boolean = True, _
  Optional ByVal rep_with_with_line_breaks As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Repeat the string (rep_pattern) n (rep_n_times) times, otionally with a line-
' number, either prefixed (linenumbersprefix=True) or attached. When the pattern
' ends with a vbLf, vbCr, or vbCrLf the attached line number is put at the left.
' The string rep_with_with_line_breaks is attached to the assembled rep_pattern.
' ------------------------------------------------------------------------------
    
    Dim i       As Long
    Dim s       As String
    Dim ln      As String
    Dim sFormat As String
    
    On Error Resume Next
    If rep_with_line_numbers Then sFormat = String$(Len(CStr(rep_n_times)), "0")
    
    For i = 1 To rep_n_times
        If rep_with_line_numbers Then ln = Format(i, sFormat)
        If rep_with_linen_umbers_as_prefix Then
            s = s & ln & " " & rep_pattern & rep_with_with_line_breaks
        Else
            s = s & rep_pattern & " " & ln & rep_with_with_line_breaks
        End If
        If Err.Number <> 0 Then
            Debug.Print "Repeate had to stop after " & i & "which resulted in a string length of " & Len(s)
            RepeatString = s
            Exit Function
        End If
    Next i
    RepeatString = s
End Function

Public Sub RepeatTest()
    Debug.Print RepeatString(10, "a", True, False, vbLf)
End Sub

Public Function Test_00_The_Buttons_Service_1() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    Set cll = mMsg.Buttons("B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", "B25", "B26", "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B34", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42", "B43", "B44", "B45", "B46", "B47", "B48", "B49", "B50")
    Test_00_The_Buttons_Service_1 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=cll)

End Function

Public Function Test_00_The_Buttons_Service_2() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    Set cll = mMsg.Buttons(2843, "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", "B25", "B26", "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B34", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42", "B43", "B44", "B45", "B46", "B47", "B48", "B49", "B50")
    Test_00_The_Buttons_Service_2 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=cll)

End Function

Public Function Test_00_The_Buttons_Service_3() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Test_00_The_Buttons_Service_3 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=mMsg.Buttons("B01,B02,B03,B04,B05"))
End Function

Public Function Test_01_WidthDeterminedByMinimumWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_01_WidthDeterminedByMinimumWidth"
    
    On Error GoTo eh
    Dim MsgForm         As fMsg
    Dim MsgTitle        As String
    
    wsTest.TestNumber = 1
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
        MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    End With
    TestMsgWidthIncrDecr = wsTest.MsgWidthIncrDecr
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    
    vButton4 = "Repeat with minimum width" & vbLf & "+ " & TestMsgWidthIncrDecr
    vButton5 = "Repeat with minimum width" & vbLf & "- " & TestMsgWidthIncrDecr
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4, vButton5)
    
    Do
        With Message.Section(1)
            .Label.Text = "Test description:"
            .Text.Text = wsTest.TestDescription
        End With
        With Message.Section(2)
            .Label.Text = "Expected test result:"
            .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & PrcPnt(TestMsgWidthMin, "w") & ") or " _
                       & "to the width determined by the reply buttons."
        End With
        With Message.Section(3)
            .Label.Text = "Please also note:"
            .Text.Text = "1. The message form height is adjusted to the required height up to the specified " & _
                         "maximum heigth which for this test is " & PrcPnt(TestMsgHeightMax, "h") & " and not exceeded." & vbLf & _
                         "2. The minimum width limit for this test is " & PrcPnt(20, "w") & " and the maximum width limit for this test is " & PrcPnt(99, "w") & "."
            .Text.FontColor = rgbRed
        End With
                                                                                                  
        Test_01_WidthDeterminedByMinimumWidth = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vbuttons _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                  )
        Select Case Test_01_WidthDeterminedByMinimumWidth
            Case vButton5
                TestMsgWidthMin = Max(TestMsgWidthMin - TestMsgWidthIncrDecr, 20)
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4, vButton5)
            Case vButton4
                TestMsgWidthMin = Min(TestMsgWidthMin + TestMsgWidthIncrDecr, 99)
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4, vButton5)
            Case BTTN_PASSED:       wsTest.Passed = True:   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
            Case Else ' Stop and Next are passed on to the caller
        End Select
    
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_02_WidthDeterminedByTitle() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_02_WidthDeterminedByTitle"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 2
    MsgTitle = Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the title's lenght."
    End With
    With Message.Section(3)
        .Label.Text = "Please note:"
        .Text.Text = "The two message sections in this test do use a proportional font " & _
                     "and thus are adjusted to form width determined by other factors." & vbLf & _
                     "The message form height is adjusted to the need up to the specified " & _
                     "maximum heigth based on the screen height which for this test is " & _
                     PrcPnt(TestMsgHeightMax, "h") & "."
    End With
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
    
    Test_02_WidthDeterminedByTitle = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vbuttons _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_02_WidthDeterminedByTitle
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_03_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_03_WidthDeterminedByMonoSpacedMessageSection"
        
    On Error GoTo eh
    Dim MsgForm                         As fMsg
    Dim MsgTitle                        As String
    Dim BttnRepeatMaxWidthIncreased     As String
    Dim BttnRepeatMaxWidthDecreased     As String
    Dim BttnRepeatMaxHeightIncreased    As String
    Dim BttnRepeatMaxHeightDecreased    As String
    
    wsTest.TestNumber = 3
    MsgTitle = Readable(PROC)
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgWidthIncrDecr = .MsgWidthIncrDecr
        TestMsgHeightMin = 25
        TestMsgHeightMax = .MsgHeightMax
        TestMsgHeightIncrDecr = .MsgHeightIncrDecr
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    BttnRepeatMaxWidthIncreased = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & TestMsgWidthIncrDecr
    BttnRepeatMaxWidthDecreased = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & TestMsgWidthIncrDecr
    BttnRepeatMaxHeightIncreased = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & TestMsgHeightIncrDecr
    BttnRepeatMaxHeightDecreased = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & TestMsgHeightIncrDecr
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthDecreased, BttnRepeatMaxHeightIncreased, BttnRepeatMaxHeightDecreased)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    Do
        With Message.Section(1)
            .Label.Text = "Test description:"
            .Text.Text = "The length of the longest monospaced message section line determines the width of the message form - " & _
                         "provided it does not exceed the specified maximum form width which for this test is " & PrcPnt(TestMsgWidthMax, "w") & " " & _
                         "of the screen size. The maximum form width may be incremented/decremented by " & PrcPnt(TestMsgWidthIncrDecr, "w") & " in order to test the result."
        End With
        With Message.Section(2)
            .Label.Text = "Expected test result:"
            .Text.Text = "Initally, the message form width is adjusted to the longest line in the " & _
                         "monospaced message section and all other message sections are adjusted " & _
                         "to this (enlarged) width." & vbLf & _
                         "When the maximum form width is reduced by " & PrcPnt(TestMsgWidthIncrDecr, "w") & " the monospaced message section is displayed with a horizontal scrollbar."
        End With
        With Message.Section(3)
            .Label.Text = "Please note the following:"
            .Text.Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                         "  the message text is not ""word wrapped""." & vbLf & _
                         "- The message form height is adjusted to the need up to the specified maximum heigth" & vbLf & _
                         "  based on the screen height which for this test is " & PrcPnt(TestMsgHeightMax, "h") & "."
            .Text.MonoSpaced = True
            .Text.FontUnderline = False
        End With
            
        '~~ Assign test values from the Test Worksheet
        mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC)).DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
        
        AssertWidthAndHeight TestMsgWidthMin _
                           , TestMsgWidthMax _
                           , TestMsgHeightMin _
                           , TestMsgHeightMax
        
        Test_03_WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vbuttons _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_min:=TestMsgHeightMin _
                 , dsply_height_max:=TestMsgHeightMax _
                  )
        Select Case Test_03_WidthDeterminedByMonoSpacedMessageSection
            Case BttnRepeatMaxWidthDecreased
                TestMsgWidthMax = TestMsgWidthMax - TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthIncreased, BttnRepeatMaxWidthDecreased)
            Case BttnRepeatMaxWidthIncreased
                TestMsgWidthMax = TestMsgWidthMax + TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthIncreased, BttnRepeatMaxWidthDecreased)
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do ' Stop, Previous, and Next are passed on to the caller
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_04_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_04_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 4
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    ' Initializations for this test
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    TestMsgWidthMax = wsTest.MsgWidthMax
    
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    End With
    With Message.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                     "which is a percentage of the screen height (for this test = " & PrcPnt(TestMsgHeightMax, "h") & "."
    End With
    vButton4 = "Repeat with 5 buttons"
    vButton5 = "Repeat with 4 buttons"
    vButton6 = "Dummy button"
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, vButton4, vButton5, vButton6, vbLf, BTTN_PASSED, BTTN_FAILED)
    
    Do
        Test_04_WidthDeterminedByReplyButtons = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vbuttons _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_04_WidthDeterminedByReplyButtons
            Case vButton4
                Set vbuttons = mMsg.Buttons(sBttnTerminate, vButton4, vButton5, vButton6, vbLf, BTTN_PASSED, BTTN_FAILED)
            Case vButton5
                Set vbuttons = mMsg.Buttons(sBttnTerminate, vButton4, vButton5, vbLf, BTTN_PASSED, BTTN_FAILED)
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do

        End Select
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 5
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & PrcPnt(TestMsgWidthMax, "w") & "."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The monospaced message section comes with a horizontal scrollbar."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "This (single line!) monspaced message section exceeds the specified maximum form width which for this test is " & PrcPnt(TestMsgWidthMax, "w") & "."
        .Text.MonoSpaced = True
    End With
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
    
    Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vbuttons _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
    
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_06_MonoSpacedMessageSectionExceedsMaxHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_06_MonoSpacedMessageSectionExceedsMaxHeight"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 6
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
       
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The height of the monospaced message section exxceeds the maximum form height (for this test " & _
                      PrcPnt(TestMsgHeightMax, "h") & " of the screen height."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "The monospaced message's height is reduced to fit the maximum form height and a vertical scrollbar is added."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = RepeatString(25, "This monospaced message comes with a vertical scrollbar." & vbLf, True)
        .Text.MonoSpaced = True
    End With
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
    
    Test_06_MonoSpacedMessageSectionExceedsMaxHeight = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vbuttons _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_06_MonoSpacedMessageSectionExceedsMaxHeight
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_07_ButtonsOnly() As Variant
    Const PROC = "Test_07_ButtonsOnly"
    
    On Error GoTo eh
    Dim MsgForm             As fMsg
    Dim MsgTitle            As String
    Dim i                   As Long
    Dim cllStory            As New Collection
    Dim vReply              As Variant
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    
    wsTest.TestNumber = 7
    MsgTitle = Readable(PROC) & ": No message, just buttons (finish with " & BTTN_PASSED & " or " & BTTN_FAILED & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMax = .MsgWidthMax:     lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMin = .MsgWidthMin:     lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax:   lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 1 To 4 ' rows
        cllStory.Add "Click this button in case ...." & vbLf & "(no lengthy message text above but everything is said in the button)"
        cllStory.Add vbLf
    Next i
    cllStory.Add BTTN_PASSED
    cllStory.Add vbLf
    cllStory.Add BTTN_FAILED
    If sBttnTerminate <> vbNullString Then
        cllStory.Add vbLf
        cllStory.Add sBttnTerminate
    End If
    
    Do
        mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC)).DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
        '~~ Obtain initial test values from the Test Worksheet
                         
        Test_07_ButtonsOnly = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cllStory _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_button_width_min:=40 _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                  )
        Select Case Test_07_ButtonsOnly
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case "Ok":                                                      Exit Do ' The very last item in the collection is the "Finished" button
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do

        End Select
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_08_ButtonsMatrix() As Variant
    Const PROC = "Test_08_ButtonsMatrix"
    
    On Error GoTo eh
    Dim MsgForm             As fMsg
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim i, j                As Long
    Dim MsgTitle            As String
    Dim cllMatrix           As Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
        
    wsTest.TestNumber = 8
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMax = .MsgWidthMax:   lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    MsgTitle = "Just to demonstrate what's theoretically possible: Buttons only! Finish with " & BTTN_PASSED & " (default) or " & BTTN_FAILED
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications

    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    Set cllMatrix = New Collection
    For i = 1 To 7 ' rows
        For j = 1 To 7 ' row buttons
            If i = 7 And j = 6 Then
                cllMatrix.Add BTTN_PASSED
                cllMatrix.Add BTTN_FAILED
                Exit For
            Else
                cllMatrix.Add "Button" & vbLf & i & "-" & j
            End If
        Next j
        If i < 7 Then cllMatrix.Add vbLf
    Next i
    
    Do
        '~~ Obtain initial test values from the Test Worksheet
        mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC)).DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
                             
        Test_08_ButtonsMatrix = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cllMatrix _
                 , dsply_button_reply_with_index:=False _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_button_width_min:=40 _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
            
        Select Case Test_08_ButtonsMatrix
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_09_ButtonScrollBarVertical() As Variant
    Const PROC = "Test_09_ButtonScrollBarVertical"
    
    On Error GoTo eh
    Dim MsgForm             As fMsg
    Dim MsgTitle            As String
    Dim i, j                As Long
    Dim cll                 As New Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    
    wsTest.TestNumber = 9
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    With wsTest
        TestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMax = .MsgWidthMax:     lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Obtain initial test values from the Test Worksheet
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                     "the specified maximum forms height - which for this test has been limited to " & _
                     PrcPnt(TestMsgHeightMax, "h") & " of the screen height."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The height for the vertically ordered buttons is reduced to fit the specified " & _
                     "maximum message form heigth and a vertical scrollbar is applied."
    End With
    With Message.Section(3)
        .Label.Text = "Finish test:"
        .Text.Text = "Click " & BTTN_PASSED & " or " & BTTN_FAILED & " (test is repeated with any other button)"
    End With
    For i = 1 To 5
        For j = 0 To 1
            cll.Add "Reply" & vbLf & "Button" & vbLf & i + j
        Next j
        cll.Add vbLf
    Next i
    cll.Add BTTN_PASSED
    cll.Add BTTN_FAILED
    
    Do
        Test_09_ButtonScrollBarVertical = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cll _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_09_ButtonScrollBarVertical
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
    
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_10_ButtonScrollBarHorizontal() As Variant

    Const PROC = "Test_10_ButtonScrollBarHorizontal"
    Const INIT_WIDTH = 40
    Const CHANGE_WIDTH = 10
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim Bttn10Plus  As String
    Dim Bttn10Minus As String
    
    wsTest.TestNumber = 10
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    TestMsgWidthMax = INIT_WIDTH
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgHeightMax = .MsgHeightMax
    End With

    Do
        mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC)).DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
        
        With Message.Section(1)
            .Label.Text = "Test description:"
            .Text.Text = "The button's width (determined by the longest buttons caption text line), " & _
                         "their number, and the button's order (all in one row) exceeds the form's " & _
                         "maximum width, explicitely specified for this test as " & _
                         PrcPnt(TestMsgWidthMax, "w") & " of the screen width."
        End With
        With Message.Section(2)
            .Label.Text = "Expected result:"
            .Text.Text = "The buttons are dsiplayed with a horizontal scroll bar to meet the specified maximimum form width."
        End With
        With Message.Section(3)
            .Label.Text = "Finish test:"
            .Text.Text = "This test is repeated with any button clicked other than the ""Ok"" button"
        End With
        
        Bttn10Plus = "Repeat with maximum form width" & vbLf & "extended by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(TestMsgWidthMax, "w") + PrcPnt(CHANGE_WIDTH, "w")
        Bttn10Minus = "Repeat with maximum form width" & vbLf & "reduced by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(TestMsgWidthMax, "w") - PrcPnt(CHANGE_WIDTH, "w")
            
        '~~ Obtain initial test values from the Test Worksheet
    
        Test_10_ButtonScrollBarHorizontal = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=mMsg.Buttons(Bttn10Plus, Bttn10Minus, BTTN_PASSED, BTTN_FAILED) _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_button_default:=BTTN_PASSED _
                  )
        Select Case Test_10_ButtonScrollBarHorizontal
            Case Bttn10Minus:       TestMsgWidthMax = TestMsgWidthMax - CHANGE_WIDTH
            Case Bttn10Plus:        TestMsgWidthMax = TestMsgWidthMax + CHANGE_WIDTH
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar() As Variant
    Const PROC = "Test_11_ButtonsMatrix_Horizontal_and_Vertical_Scrollbar"
    
    On Error GoTo eh
    Dim MsgForm                     As fMsg
    Dim i, j                        As Long
    Dim MsgTitle                      As String
    Dim cllMatrix                   As Collection
    Dim bMonospaced                 As Boolean: bMonospaced = True ' initial test value
    Dim TestMsgWidthMin     As Long
    Dim TestMsgWidthMaxSpecInPt     As Long
    Dim TestMsgHeightMax  As Long
    
    wsTest.TestNumber = 11
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    
    MsgTitle = "Buttons only! With a vertical and a horizontal scrollbar! Finish with " & BTTN_PASSED & " or " & BTTN_FAILED
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    Set cllMatrix = New Collection
    For i = 1 To 7 ' rows
        For j = 1 To 7 ' row buttons
            If i = 7 And j = 5 Then
                cllMatrix.Add BTTN_PASSED
                cllMatrix.Add BTTN_FAILED
                Exit For
            Else
                cllMatrix.Add vbLf & " ---- Button ---- " & vbLf & i & "-" & j & vbLf & " "
            End If
        Next j
        If i < 7 Then cllMatrix.Add vbLf
    Next i
    
    Do
        '~~ Obtain initial test values from the Test Worksheet
        mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC)).DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
                             
        Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cllMatrix _
                 , dsply_button_reply_with_index:=False _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_button_width_min:=40 _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_16_ButtonByDictionary()
' -----------------------------------------------
' The buttons argument is provided as Dictionary.
' -----------------------------------------------
    Const PROC  As String = "Test_16_ButtonByDictionary"
    
    Dim dct     As New Collection
    Dim MsgTitle   As String
    Dim MsgForm As fMsg
    
    wsTest.TestNumber = 16
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    dct.Add "Yes"
    dct.Add "No"
    
    Test_16_ButtonByDictionary = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=dct _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_17_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_17_Box_MessageAsString"
        
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 17
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
        
    Test_17_MessageAsString = _
    mMsg.Box( _
             box_title:=MsgTitle _
           , box_msg:="This is a message provided as a simple string argument!" _
           , box_buttons:=vbuttons _
           , box_width_min:=TestMsgWidthMin _
           , box_width_max:=TestMsgWidthMax _
           , box_height_max:=TestMsgHeightMax _
            )
    Select Case Test_17_MessageAsString
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_20_ButtonByValue()

    Const PROC  As String = "Test_20_ButtonByValue"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle   As String
    
    wsTest.TestNumber = 20
    MsgTitle = "Test: Button by value (" & PROC & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as VB MsgBox value vbYesNo."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in one row"
    End With
    Test_20_ButtonByValue = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vbOKOnly _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
            
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_21_ButtonByString()

    Const PROC  As String = "Test_21_ButtonByString"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 21
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_21_ButtonByString = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:="Yes," & vbLf & ",No" _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_22_ButtonByCollection()

    Const PROC  As String = "Test_22_ButtonByCollection"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim cll         As New Collection
    
    wsTest.TestNumber = 22
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    cll.Add "Yes"
    cll.Add "No"
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_22_ButtonByCollection = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=cll _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_30_Monitor() As Variant
    Const PROC = "Test_30_Monitor"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim i           As Long
    Dim PrgrsHeader As String
    Dim PrgrsMsg    As String
    Dim iLoops      As Long
    Dim lWait       As Long
    
    PrgrsHeader = " No. Status   Step"
    iLoops = 12
    
    wsTest.TestNumber = 30
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    PrgrsMsg = vbNullString
    
    For i = 1 To iLoops
        PrgrsMsg = mBasic.Align(i, 4, AlignRight, " ") & mBasic.Align("Passed", 8, AlignCentered, " ") & Repeat(repeat_n_times:=Int(((i - 1) / 10)) + 1, repeat_string:="  " & mBasic.Align(i, 2, AlignRight) & ".  Follow-Up line after " & Format(lWait, "0000") & " Milliseconds.")
        If i < iLoops Then
            mMsg.Monitor mntr_title:=MsgTitle _
                       , mntr_msg:=PrgrsMsg _
                       , mntr_msg_monospaced:=True _
                       , mntr_header:=" No. Status  Step"
            '~~ Simmulation of a process
            lWait = 100 * i
            DoEvents
            Sleep 200
        Else
            mMsg.Monitor mntr_title:=MsgTitle _
                       , mntr_msg:=PrgrsMsg _
                       , mntr_header:=" No. Status  Step" _
                       , mntr_footer:="Process finished! Close this window"
        End If
    Next i
    
    Select Case mMsg.Box(box_title:="Test result of " & Readable(PROC) _
                       , box_msg:=vbNullString _
                       , box_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED) _
                        )
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_90_All_in_one_Demonstration() As Variant
' ------------------------------------------------------------------------------
' Demo as test of as many features as possible at once.
' ------------------------------------------------------------------------------
    Const PROC              As String = "Test_90_All_in_one_Demonstration"

    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim cll         As New Collection
    Dim i, j        As Long
    Dim Message     As TypeMsg
   
    wsTest.TestNumber = 90
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    With Message.Section(1)
        .Label.Text = "Displayed message summary "
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "- Display of (all) 4 message sections, each with an (optional) label" & vbLf _
                   & "- One monospaced section text exceeding the specified maximum width" & vbLf _
                   & "- Display of some of the 49(7x7) possible reply buttons" & vbLf _
                   & "- Font options like color, bold, and italic"
    End With
    With Message.Section(2)
        .Label.Text = "Unlimited message width"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "This section's text is mono-spaced and thus not word-wrapped. I.e. the longest line determines the messag width." & vbLf _
                   & "Because the maximimum width for this demo has been specified " & PrcPnt(TestMsgWidthMax, "w") & " of the screen width (defaults to " & PrcPnt(80, "w") & vbLf _
                   & "the text is displayed with a horizontal scrollbar. The size limit for a section's text is only limited by VBA" & vbLf _
                   & "which as about 1GB! (see also unlimited message height below)"
        .Text.MonoSpaced = True
        .Text.FontItalic = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "All the message sections together ecxeed the maximum height, specified for this demo " & PrcPnt(TestMsgHeightMax, "h") & " " _
                   & "of the screen height (defaults to " & PrcPnt(85, "h") & ". Thus the message area is displayed with a vertical scrollbar. I. e. no matter " _
                   & "how much text is displayed, it is never truncated. The only limit is VBA's limit for a text " _
                   & "string which is abut 1GB! With 4 strings, each in one section the limit is thus about 4GB !!!!"
    End With
    With Message.Section(4)
        .Label.Text = "Reply buttons flexibility"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "This demo displays only some of the 49 possible reply buttons (7 rows by 7 buttons). " _
                   & "It also shows that a reply button can have any caption text and the buttons can be " _
                   & "displayed in any order within the 7 x 7 limit. Of cource the VBA.MsgBox classic " _
                   & "vbOkOnly, vbYesNoCancel, etc. are also possible - even in a mixture." & vbLf & vbLf _
                   & "By the way: End this demo with either " & BTTN_PASSED & " or " & BTTN_FAILED & " clicked (else it loops)."
    End With
    '~~ Prepare the buttons collection
    For j = 1 To 1
        For i = 1 To 5
            cll.Add "Sample multiline" & vbLf & "reply button" & vbLf & "Button-" & j & "-" & i
        Next i
        cll.Add vbLf
    Next j
    cll.Add BTTN_PASSED
    cll.Add BTTN_FAILED
        
    Do
        Test_90_All_in_one_Demonstration = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cll _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_90_All_in_one_Demonstration
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function Test_91_MinimumMessage() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_91_MinimumMessage"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    wsTest.TestNumber = 1
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    TestMsgWidthIncrDecr = wsTest.MsgWidthIncrDecr
    TestMsgHeightIncrDecr = wsTest.MsgHeightIncrDecr
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & PrcPnt(TestMsgWidthMin, "w") & ") or " _
                   & "to the width determined by the reply buttons."
    End With
    With Message.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height up to the specified " & _
                     "maximum heigth which is " & PrcPnt(TestMsgHeightMax, "h") & " and not exceeded."
        .Text.FontColor = rgbRed
    End With
                                                                                              
    mMsg.Dsply dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless
             
xt: Exit Function

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Sub Test_99_Individual()
' ---------------------------------------------------------------------------------
' Test 1 (optional arguments are used in conjunction with the Regression test only)
' ---------------------------------------------------------------------------------
    Const PROC = "Test_99_Individual"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    MsgTitle = "This title is rather short"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    TestMsgWidthMax = 80
    MsgForm.DsplyFrmsWthBrdrsTestOnly = wsTest.TestOptionDisplayFrames
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text"
    End With
    With Message.Section(2)
        .Label.Text = "Test label extra long in order to test the adjustment of the message window width:"
        .Text.Text = "A short message text"
        .Text.MonoSpaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text"
    End With
    mMsg.Dsply dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=mMsg.Buttons("Button-1", "Button-2") _
             , dsply_button_default:="Button-1" _
             , dsply_width_min:=30 _
             , dsply_width_max:=TestMsgWidthMax
    
xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function PrcPnt(ByVal pp_value As Single, _
                        ByVal pp_dimension As String) As String
    PrcPnt = mMsg.Prcnt(pp_value, pp_dimension) & "% (" & mMsg.Pnts(pp_value, "w") & ")"
End Function
