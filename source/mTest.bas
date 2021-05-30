Attribute VB_Name = "mTest"
Option Explicit
Option Compare Text
' -----------------------------------------------------
' Standard Module mTest
'          All tests for a complete regression test.
'          Obligatory performed after any modification.
'          Ammended when new features or functions are
'          implemented.
'
' Please note:
' Errors raised by the tested procedures cannot be
' asserted since they are not passed on to the calling
' /entry procedure. This would require the Common
' Standard Module mErrHndlr which intentionally is not
' used by this module.
'
' W. Rauschenberger, Berlin June 2020
' -----------------------------------------------------
Public Const BTTN_FINISH    As String = "Test Done"
Public Const BTTN_PASSED    As String = "Passed"
Public Const BTTN_FAILED    As String = "Failed"

Dim TestMsgWidthMaxSpecAsPoSS   As Long
Dim TestMsgHeightMaxSpecAsPoSS  As Long
Dim bRegressionTest             As Boolean
Dim TestMsgHeightIncrDecr       As Long
Dim TestMsgWidthIncrDecr        As Long
Dim TestMsgWidthMinSpecInPt     As Long
Dim Message                     As TypeMsg
Dim sBttnTerminate              As String
Dim sMsgTitle                   As String
Dim vButton4                    As Variant
Dim vButton5                    As Variant
Dim vButton6                    As Variant
Dim vButton7                    As Variant
Dim vbuttons                    As Variant

Public Property Let RegressionTest(ByVal b As Boolean)
    bRegressionTest = b
    If b Then sBttnTerminate = "Terminate" & vbLf & "Regression" Else sBttnTerminate = vbNullString
End Property
Public Property Get BTTN_TERMINATE() As String ' composed constant
    BTTN_TERMINATE = "Terminate" & vbLf & "Regression"
End Property

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mTest." & s:  End Property

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

' -------------------------------------------------------------------------------------------------
' Procedures for test start via Command Buttons on Test Worksheet
Public Sub cmdTest01_Click()
'    wsTest.RegressionTest = False
    wsTest.TestNumber = 1
    mTest.Test_01_WidthDeterminedByMinimumWidth
End Sub

Public Sub cmdTest02_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 2
    mTest.Test_02_WidthDeterminedByTitle
End Sub

Public Sub cmdTest03_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 3
    mTest.Test_03_WidthDeterminedByMonoSpacedMessageSection
End Sub

Public Sub cmdTest04_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 4
    mTest.Test_04_WidthDeterminedByReplyButtons
End Sub

Public Sub cmdTest05_Click()
    wsTest.RegressionTest = False
    mTest.Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth
End Sub

Public Sub cmdTest06_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 6
    mTest.Test_06_MonoSpacedMessageSectionExceedsMaxHeight
End Sub

Public Sub cmdTest07_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 7
    mTest.Test_07_OnlyButtons
End Sub

Public Sub cmdTest08_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 8
    mTest.Test_08_ButtonsMatrix
End Sub

Public Sub cmdTest09_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 9
    mTest.Test_09_ButtonScrollBarVertical
End Sub

Public Sub cmdTest10_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 10
    mTest.Test_10_ButtonScrollBarHorizontal
End Sub

Public Sub cmdTest11_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 11
    mTest.Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
End Sub

Public Sub cmdTest30_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 30
    mTest.Test_30_Progress_FollowUp
End Sub

Public Sub cmdTest90_Click()
    wsTest.RegressionTest = False
    wsTest.TestNumber = 90
    mTest.Test_90_All_in_one_Demonstration
End Sub

Public Sub DisplayFramesOption()
    If ActiveSheet.Shapes("optDisplayFrames").OLEFormat.Object.Value = 1 _
    Then wsTest.DsplyFrmsWthBrdrsTestOnly = True _
    Else wsTest.DsplyFrmsWthBrdrsTestOnly = False
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

Public Sub Explore(ByVal ctl As Variant, _
          Optional ByVal applied As Boolean = True)

    Dim dct     As New Dictionary
    Dim v       As Variant
    Dim Appl    As String   ' ControlApplied
    Dim l       As String   ' .Left
    Dim W       As String   ' .Width
    Dim T       As String   ' .Top
    Dim H       As String   ' .Height
    Dim SW      As String   ' .ScrollWidth
    Dim SH      As String   ' .ScrollHeight
    Dim FW      As String   ' fMsg.InsideWidth
    Dim CW      As String   ' Content width
    Dim CH      As String   ' Content height
    Dim FH      As String   ' fMsg.InsideHeight
    Dim i       As Long
    Dim Item    As String
    Dim j       As String
    Dim frm     As MSForms.Frame
    
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
                        If fMsg.IsApplied(v) Then mDct.DctAdd dct, v, Item
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
        If fMsg.IsApplied(ctl) Then Appl = "Yes " Else Appl = " No "
        l = Align(Format(ctl.Left, "000.0"), 7, AlignCentered, " ")
        W = Align(Format(ctl.Width, "000.0"), 7, AlignCentered, " ")
        T = Align(Format(ctl.Top, "000.0"), 7, AlignCentered, " ")
        H = Align(Format(ctl.Height, "000.0"), 7, AlignCentered, " ")
        FH = Align(Format(fMsg.InsideHeight, "000.0"), 7, AlignCentered, " ")
        FW = Align(Format(fMsg.InsideWidth, "000.0"), 7, AlignCentered, " ")
        If TypeName(ctl) = "Frame" Then
            Set frm = ctl
            CW = Align(Format(fMsg.FrameContentWidth(frm), "000.0"), 7, AlignCentered, " ")
            CH = Align(Format(fMsg.FrameContentHeight(frm), "000.0"), 7, AlignCentered, " ")
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

Private Sub MessageInit()
' ------------------------------------------------------
' Initializes the all message sections with the defaults
' throughout this test module which uses a module global
' declared Message for a consistent layout.
' ------------------------------------------------------
    Dim i   As Long
    For i = 1 To fMsg.NoOfDesignedMsgSects
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
    If bRegressionTest Then mTest.RegressionTest = True Else mTest.RegressionTest = False
End Sub

' Convert a string (s) into a readable form by replacing all underscores
' with a whitespace and all characters immediately following an underscore
' to a lowercase letter.
' ---------------------------------------------------------------------
Private Function Readable(ByVal s As String) As String
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

Public Sub Test_00_AutoSizeHeight()
    With fMsg
        .AutoSizeHeight .tbMsgSection1Text, 300, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
        .AutoSizeHeight .tbMsgSection1Text, 250, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
        .AutoSizeHeight .tbMsgSection1Text, 200, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
        .AutoSizeHeight .tbMsgSection1Text, 150, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
        .AutoSizeHeight .tbMsgSection1Text, 100, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
        .AutoSizeHeight .tbMsgSection1Text, 50, "This text is a bit longer than the specified with and thus should result in a corresponding height"
        Debug.Print .tbMsgSection1Text.Height
    End With
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
    
    wsTest.TestNumber = 1
    sMsgTitle = Readable(PROC)
    Unload fMsg
    MessageInit ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    TestMsgWidthIncrDecr = wsTest.MsgWidthIncrDecr
    TestMsgHeightIncrDecr = wsTest.MsgHeightIncrDecr
    
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
            .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & fMsg.MsgWidthMinSpecInPt & " pt) or " _
                       & "to the width determined by the reply buttons."
        End With
        With Message.Section(3)
            .Label.Text = "Please also note:"
            .Text.Text = "The message form height is adjusted to the required height up to the specified " & _
                         "maximum heigth which is " & fMsg.MsgHeightMaxSpecAsPoSS & "% and not exceeded."
            .Text.FontColor = rgbRed
        End With
                                                                                                  
        Test_01_WidthDeterminedByMinimumWidth = _
        mMsg.Dsply( _
                 dsply_title:=sMsgTitle, _
                 dsply_msg:=Message, _
                 dsply_buttons:=vbuttons _
                )
        Select Case Test_01_WidthDeterminedByMinimumWidth
            Case vButton5
                fMsg.MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt - TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4)
            Case vButton4
                fMsg.MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt + TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton5)
            Case BTTN_PASSED:       wsTest.Passed = True:   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
            Case Else ' Stop and Next are passed on to the caller
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_02_WidthDeterminedByTitle() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_02_WidthDeterminedByTitle"
    
    On Error GoTo eh
    wsTest.TestNumber = 2
    sMsgTitle = Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    Unload fMsg
    
    '~~ Obtain initial test values from the Test Worksheet
    TestMsgWidthIncrDecr = wsTest.MsgWidthIncrDecr
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
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
                     fMsg.MsgHeightMaxSpecAsPoSS & "%."
    End With
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
    
    Test_02_WidthDeterminedByTitle = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_msg:=Message, _
             dsply_buttons:=vbuttons _
            )
    Select Case Test_02_WidthDeterminedByTitle
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_03_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_03_WidthDeterminedByMonoSpacedMessageSection"
        
    On Error GoTo eh
    Dim BttnRepeatMaxWidthIncreased     As String
    Dim BttnRepeatMaxWidthDecreased     As String
    Dim BttnRepeatMaxHeightIncreased    As String
    Dim BttnRepeatMaxHeightDecreased    As String
    
    wsTest.TestNumber = 3
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    TestMsgWidthIncrDecr = wsTest.MsgWidthIncrDecr
    TestMsgHeightIncrDecr = wsTest.MsgHeightIncrDecr
    
    ' Initializations for this test
    TestMsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
    TestMsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
    
    BttnRepeatMaxWidthIncreased = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & TestMsgWidthIncrDecr
    BttnRepeatMaxWidthDecreased = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & TestMsgWidthIncrDecr
    BttnRepeatMaxHeightIncreased = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & TestMsgHeightIncrDecr
    BttnRepeatMaxHeightDecreased = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & TestMsgHeightIncrDecr
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthDecreased, BttnRepeatMaxHeightIncreased, BttnRepeatMaxHeightDecreased)
    MessageInit ' set test-global message specifications
    Do
        With Message.Section(1)
            .Label.Text = "Test description:"
            .Text.Text = "The length of the longest monospaced message section line determines the width of the message form - " & _
                         "provided it does not exceed the specified maximum form width which for this test is " & TestMsgWidthMaxSpecAsPoSS & "% " & _
                         "of the screen size. The maximum form width may be incremented/decremented by " & TestMsgWidthIncrDecr & "% in order to test the result."
        End With
        With Message.Section(2)
            .Label.Text = "Expected test result:"
            .Text.Text = "Initally, the message form width is adjusted to the longest line in the " & _
                         "monospaced message section and all other message sections are adjusted " & _
                         "to this (enlarged) width." & vbLf & _
                         "When the maximum form width is reduced by " & TestMsgWidthIncrDecr & " % the monospaced message section is displayed with a horizontal scrollbar."
        End With
        With Message.Section(3)
            .Label.Text = "Please note the following:"
            .Text.Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                         "  the message text is not ""word wrapped""." & vbLf & _
                         "- The message form height is adjusted to the need up to the specified maximum heigth" & vbLf & _
                         "  based on the screen height which for this test is " & TestMsgHeightMaxSpecAsPoSS & "%."
            .Text.MonoSpaced = True
            .Text.FontUnderline = False
        End With
            
        '~~ Obtain initial test values from the Test Worksheet
        With fMsg
            .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
            .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
            .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
            .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
        End With
        
        Test_03_WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Dsply(dsply_title:=sMsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vbuttons _
                 , dsply_max_width:=TestMsgWidthMaxSpecAsPoSS _
                 , dsply_max_height:=TestMsgHeightMaxSpecAsPoSS _
                )
        Select Case Test_03_WidthDeterminedByMonoSpacedMessageSection
            Case BttnRepeatMaxWidthDecreased
                TestMsgWidthMaxSpecAsPoSS = TestMsgWidthMaxSpecAsPoSS - TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthIncreased)
            Case BttnRepeatMaxWidthIncreased
                TestMsgWidthMaxSpecAsPoSS = TestMsgWidthMaxSpecAsPoSS + TestMsgWidthIncrDecr
                Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthDecreased)
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do ' Stop, Previous, and Next are passed on to the caller
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_04_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_04_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    wsTest.TestNumber = 4
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    fMsg.DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    
    MessageInit    ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MsgWidthMaxSpecInPt & " (which is the specified " & fMsg.MsgWidthMaxSpecAsPoSS & "% of the screen width)."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    End With
    With Message.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                     "which is a percentage of the screen height (for this test = " & fMsg.MsgHeightMaxSpecAsPoSS & "%."
    End With
    vButton4 = "Repeat with 5 buttons"
    vButton5 = "Repeat with 4 buttons"
    vButton6 = "Dummy button"
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, vButton4, vButton5, vButton6, vbLf, BTTN_PASSED, BTTN_FAILED)
    
    Do
        fMsg.MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        Test_04_WidthDeterminedByReplyButtons = _
        mMsg.Dsply( _
                   dsply_title:=sMsgTitle, _
                   dsply_msg:=Message, _
                   dsply_buttons:=vbuttons _
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

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh
    wsTest.TestNumber = 5
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & fMsg.MsgWidthMaxSpecInPt & " pt (the equivalent of " & fMsg.MsgWidthMaxSpecAsPoSS & "% of the screen width)."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The monospaced message section comes with a horizontal scrollbar."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "This (single line!) monspaced message section exceeds the specified maximum form width which for this test is " & fMsg.MsgWidthMaxSpecInPt & " pt, " & _
                     " which is the equivalent of " & fMsg.MsgWidthMaxSpecAsPoSS & "% of the screen width."
        .Text.MonoSpaced = True
    End With
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
    
    Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_msg:=Message, _
             dsply_buttons:=vbuttons _
            )
    Select Case Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
    
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_06_MonoSpacedMessageSectionExceedsMaxHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_06_MonoSpacedMessageSectionExceedsMaxHeight"
    
    On Error GoTo eh
    wsTest.TestNumber = 6
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
       
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription ' "The height of the monospaced message section exxceeds the maximum form height for this test (" _
                   & fMsg.MaxMsgHeightPts & ") which is the specified " & fMsg.MsgHeightMaxSpecAsPoSS & "% of the screen height."
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
    mMsg.Dsply( _
               dsply_title:=sMsgTitle, _
               dsply_msg:=Message, _
               dsply_buttons:=vbuttons _
              )
    Select Case Test_06_MonoSpacedMessageSectionExceedsMaxHeight
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_07_OnlyButtons() As Variant
    Const PROC              As String = "Test_07_OnlyButtons"
    
    On Error GoTo eh
    Dim i                       As Long
    Dim cllStory                As New Collection
    Dim vReply                  As Variant
    Dim lChangeHeightPcntg      As Long
    Dim lChangeWidthPcntg       As Long
    Dim lChangeMinWidthPt       As Long
    Dim bMonospaced             As Boolean: bMonospaced = True ' initial test value
    
    wsTest.TestNumber = 7
    sMsgTitle = Readable(PROC) & ": No message, just buttons (finish with " & BTTN_PASSED & " or " & BTTN_FAILED & ")"
    Unload fMsg                     ' Ensures a message starts from scratch
       
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMaxSpecAsPoSS = .MsgWidthMaxSpecAsPoSS: lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMinSpecInPt = .MsgWidthMinSpecInPt:     lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMaxSpecAsPoSS = .MsgHeightMaxSpecAsPoSS:  lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 1 To 4 ' rows
        cllStory.Add "Click this button in case ...." & vbLf & "(no lengthy message text above but everything is said in the button)"
        cllStory.Add vbLf
    Next i
    cllStory.Add BTTN_PASSED
    cllStory.Add vbLf
    cllStory.Add BTTN_FAILED
    cllStory.Add vbLf
    cllStory.Add sBttnTerminate
    Do
        '~~ Obtain initial test values from the Test Worksheet
        With fMsg
            .MinButtonWidth = 40
            .MsgWidthMinSpecInPt = TestMsgWidthMinSpecInPt
            .MsgWidthMaxSpecAsPoSS = TestMsgWidthMaxSpecAsPoSS    ' for this demo to enforce a vertical scrollbar
            .MsgHeightMaxSpecAsPoSS = TestMsgHeightMaxSpecAsPoSS  ' for this demo to enbforce a vertical scrollbar for the message section
            .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
        End With
                         
        Test_07_OnlyButtons = _
        mMsg.Dsply(dsply_title:=sMsgTitle, _
                   dsply_msg:=Message, _
                   dsply_buttons:=cllStory _
                  )
        Select Case Test_07_OnlyButtons
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case "Ok":                                                      Exit Do ' The very last item in the collection is the "Finished" button
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do

        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_08_ButtonsMatrix() As Variant
    Const PROC              As String = "Test_08_ButtonsMatrix"
    
    On Error GoTo eh
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim i, j                As Long
    Dim sTitle              As String
    Dim cllMatrix           As Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    
    wsTest.TestNumber = 8
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMinSpecInPt = .MsgWidthMinSpecInPt:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMaxSpecAsPoSS = .MsgWidthMaxSpecAsPoSS:   lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMaxSpecAsPoSS = .MsgHeightMaxSpecAsPoSS: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    
    sTitle = "Just to demonstrate what's theoretically possible: Buttons only! Finish with ""Ok"" (the default button)"
    MessageInit ' set test-global message specifications
'    Message.Section(1).Text.Text = "Some can play around with button matrix of 7 by 7 buttons"
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
        With fMsg
            .MinButtonWidth = 40
            .MsgWidthMaxSpecInPt = TestMsgWidthMinSpecInPt
            .MsgWidthMaxSpecAsPoSS = TestMsgWidthMaxSpecAsPoSS    ' for this demo to enforce a vertical scrollbar
            .MsgHeightMaxSpecAsPoSS = TestMsgHeightMaxSpecAsPoSS  ' for this demo to enbforce a vertical scrollbar for the message section
            .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
        End With
                             
        Test_08_ButtonsMatrix = _
        mMsg.Dsply(dsply_title:=sTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cllMatrix _
                 , dsply_reply_with_index:=False _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_modeless:=True)
            
        Select Case Test_08_ButtonsMatrix
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_09_ButtonScrollBarVertical() As Variant
    Const PROC = "Test_09_ButtonScrollBarVertical"
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim cll                 As New Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    
    wsTest.TestNumber = 9
    sMsgTitle = Readable(PROC)
    Unload fMsg
    
    With wsTest
        TestMsgWidthMinSpecInPt = .MsgWidthMinSpecInPt:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMaxSpecAsPoSS = .MsgWidthMaxSpecAsPoSS:     lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMaxSpecAsPoSS = .MsgHeightMaxSpecAsPoSS: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMaxSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                     "the specified maximum forms height - which for this test has been limited to " & _
                     fMsg.MsgHeightMaxSpecAsPoSS & "% of the screen height."
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
        mMsg.Dsply(dsply_title:=sMsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cll _
                  )
        Select Case Test_09_ButtonScrollBarVertical
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
    
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_10_ButtonScrollBarHorizontal() As Variant

    Const PROC = "Test_10_ButtonScrollBarHorizontal"
    Const INIT_WIDTH = 40
    Const CHANGE_WIDTH = 10
    
    On Error GoTo eh
    Dim Bttn10Plus  As String
    Dim Bttn10Minus As String
    Dim WidthMax    As Long
    
    wsTest.TestNumber = 10
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    WidthMax = INIT_WIDTH

    Do
        Unload fMsg                                         ' Ensures a message starts from scratch
        fMsg.MsgWidthMaxSpecAsPoSS = WidthMax ' enforce horizontal scrollbar
        MessageInit ' set test-global message specifications
        With Message.Section(1)
            .Label.Text = "Test description:"
            .Text.Text = "The width, number, amd order of the displayed reply buttons exceeds " & _
                         "the specified maximum forms width, which for this test is limited to " & _
                         fMsg.MsgWidthMaxSpecAsPoSS & "% of the screen width."
        End With
        With Message.Section(2)
            .Label.Text = "Expected result:"
            .Text.Text = "When the maximum form width is enlarged the message will be displayed without a scrollbar."
        End With
        With Message.Section(3)
            .Label.Text = "Finish test:"
            .Text.Text = "This test is repeated with any button clicked other than the ""Ok"" button"
        End With
        
        Bttn10Plus = "Repeat with maximum form width" & vbLf & "extended by " & CHANGE_WIDTH & "% to " & WidthMax + CHANGE_WIDTH & "%"
        Bttn10Minus = "Repeat with maximum form width" & vbLf & "reduced by " & CHANGE_WIDTH & "% to " & WidthMax - CHANGE_WIDTH & "%"
            
        '~~ Obtain initial test values from the Test Worksheet
        With fMsg
            .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
            .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
            .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
            .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
        End With
    
        Test_10_ButtonScrollBarHorizontal = _
        mMsg.Dsply(dsply_title:=sMsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=mMsg.Buttons(Bttn10Plus, Bttn10Minus, BTTN_PASSED, BTTN_FAILED))
        Select Case Test_10_ButtonScrollBarHorizontal
            Case Bttn10Minus:   WidthMax = WidthMax - CHANGE_WIDTH
            Case Bttn10Plus:    WidthMax = WidthMax + CHANGE_WIDTH
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar() As Variant
    Const PROC              As String = "Test_11_ButtonsMatrix_Horizontal_and_Vertical_Scrollbar"
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim sTitle              As String
    Dim cllMatrix           As Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim TestMsgWidthMinSpecInPt       As Long
    Dim TestMsgWidthMaxSpecInPt        As Long
    Dim TestMsgHeightMaxSpecAsPoSS       As Long
    
    wsTest.TestNumber = 11
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMinSpecInPt = .MsgWidthMinSpecInPt:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMaxSpecAsPoSS = .MsgWidthMaxSpecAsPoSS:   lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMaxSpecAsPoSS = .MsgHeightMaxSpecAsPoSS: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
    
    sTitle = "Buttons only! With a vertical and a horizontal scrollbar! Finish with ""Ok"" (the default button)"
    MessageInit ' set test-global message specifications
'    Message.Section(1).Text.Text = "Some can play around with button matrix of 7 by 7 buttons"
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    Set cllMatrix = New Collection
    For i = 1 To 7 ' rows
        For j = 1 To 7 ' row buttons
            If (j * i) < 48 Then
                cllMatrix.Add vbLf & " ---- Button ---- " & vbLf & i & "-" & j & vbLf & " "
            Else
                cllMatrix.Add BTTN_PASSED
                cllMatrix.Add BTTN_FAILED
            End If
        Next j
        If i < 7 Then cllMatrix.Add vbLf
    Next i
    
    Do
        '~~ Obtain initial test values from the Test Worksheet
        With fMsg
            .MinButtonWidth = 40
            .MsgWidthMaxSpecInPt = TestMsgWidthMaxSpecInPt
            .MsgWidthMaxSpecAsPoSS = TestMsgWidthMaxSpecInPt    ' for this demo to enforce a vertical scrollbar
            .MsgHeightMaxSpecAsPoSS = TestMsgHeightMaxSpecAsPoSS  ' for this demo to enbforce a vertical scrollbar for the message section
            .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
        End With
                             
        Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar = _
        mMsg.Dsply(dsply_title:=sTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cllMatrix _
                 , dsply_reply_with_index:=True _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_modeless:=True)
        Select Case Test_11_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_16_ButtonByDictionary()
' -----------------------------------------------
' The buttons argument is provided as Dictionary.
' -----------------------------------------------
    Const PROC  As String = "Test_16_ButtonByDictionary"
    Dim dct     As New Collection
    
    wsTest.TestNumber = 16
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
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
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=dct _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_17_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_17_Box_MessageAsString"
        
    On Error GoTo eh
    wsTest.TestNumber = 17
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    Set vbuttons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
        
    Test_17_MessageAsString = _
    mMsg.Box( _
             box_title:=sMsgTitle _
           , box_msg:="This is a message provided as a simple string argument!" _
           , box_buttons:=vbuttons _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_20_ButtonByValue()

    Const PROC  As String = "Test_20_ButtonByValue"
    
    On Error GoTo eh
    wsTest.TestNumber = 20
    Unload fMsg                     ' Ensures a message starts from scratch
        
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as VB MsgBox value vbYesNo."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in one row"
    End With
    Test_20_ButtonByValue = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & PROC & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=vbOKOnly _
            )
            
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_21_ButtonByString()

    Const PROC  As String = "Test_21_ButtonByString"
    
    Unload fMsg                     ' Ensures a message starts from scratch
    wsTest.TestNumber = 21
        
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_21_ButtonByString = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:="Yes," & vbLf & ",No" _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_22_ButtonByCollection()

    Const PROC  As String = "Test_22_ButtonByCollection"
    Dim cll     As New Collection
    
    wsTest.TestNumber = 22
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMinSpecInPt = wsTest.MsgWidthMinSpecInPt
        .MsgWidthMaxSpecAsPoSS = wsTest.MsgWidthMaxSpecAsPoSS
        .MsgHeightMaxSpecAsPoSS = wsTest.MsgHeightMaxSpecAsPoSS
        .DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    End With
    
    cll.Add "Yes"
    cll.Add "No"
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_22_ButtonByCollection = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=cll _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_30_Progress_FollowUp() As Variant
    Const PROC = "Test_30_Progress_FollowUp"
    
    On Error GoTo eh
    Dim i As Long
    Dim PrgrsHeader As String: PrgrsHeader = " No. Status   Step"
    Dim PrgrsMsg    As String
    Dim iLoops      As Long: iLoops = 25
        
    wsTest.TestNumber = 30
    sMsgTitle = Readable(PROC)
    Unload mMsg.Form(sMsgTitle)                     ' Ensures a message starts from scratch
    
    mMsg.Form(sMsgTitle).DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    
    For i = 1 To iLoops
        PrgrsMsg = mBasic.Align(i, 4, AlignRight, " ") & mBasic.Align("Passed", 8, AlignCentered, " ") & Repeat(repeat_n_times:=(Int((i / 10)) + 1) + 1, repeat_string:="  " & mBasic.Align(i, 2, AlignRight) & ".  Follow-Up line")
        If i < iLoops Then
'            Debug.Print i & ". Line"
            mMsg.Progress prgrs_title:=sMsgTitle _
                        , prgrs_msg:=PrgrsMsg _
                        , prgrs_msg_monospaced:=True _
                        , prgrs_header:=" No. Status  Step" _
                        , prgrs_max_height:=50 _
                        , prgrs_max_width:=50 _
                        , prgrs_buttons:=vbOKOnly
        Else
            mMsg.Progress prgrs_title:=sMsgTitle _
                        , prgrs_msg:=PrgrsMsg _
                        , prgrs_header:=" No. Status  Step" _
                        , prgrs_footer:="Process finished! Press ""Ok"" to terminate the display." _
                        , prgrs_buttons:=vbOKOnly
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

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Function Test_90_All_in_one_Demonstration() As Variant
' ------------------------------------------------------------------------------
' Demo as test of as many features as possible at once.
' ------------------------------------------------------------------------------
    Const PROC              As String = "Test_90_All_in_one_Demonstration"
    Const TEST_MAX_WIDTH    As Long = 55
    Const TEST_MAX_HEIGHT   As Long = 65

    On Error GoTo eh
    Dim cll                 As New Collection
    Dim i, j                As Long
    Dim Message             As TypeMsg
   
    wsTest.TestNumber = 90
    sMsgTitle = Readable(PROC)
    Unload fMsg                     ' Ensures a message starts from scratch
    
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
                   & "Because the maximimum width for this demo has been specified " & TEST_MAX_WIDTH & "% of the screen width (defaults to 80%)" & vbLf _
                   & "the text is displayed with a horizontal scrollbar. The size limit for a section's text is only limited by VBA" & vbLf _
                   & "which as about 1GB! (see also unlimited message height below)"
        .Text.MonoSpaced = True
        .Text.FontItalic = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "All the message sections together ecxeed the maximum height, specified for this demo " & TEST_MAX_HEIGHT & "% " _
                   & "of the screen height (defaults to 70%). Thus the message area is displayed with a vertical scrollbar. I. e. no matter " _
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
    fMsg.DsplyFrmsWthBrdrsTestOnly = wsTest.DsplyFrmsWthBrdrsTestOnly
    
    Do
        Test_90_All_in_one_Demonstration = _
        mMsg.Dsply(dsply_title:=sMsgTitle _
                   , dsply_msg:=Message _
                   , dsply_buttons:=cll _
                   , dsply_button_default:=BTTN_PASSED _
                   , dsply_max_height:=TEST_MAX_HEIGHT _
                   , dsply_max_width:=TEST_MAX_WIDTH _
                    )
        Select Case Test_90_All_in_one_Demonstration
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Function

Public Sub Test_99_Individual()
' ---------------------------------------------------------------------------------
' Test 1 (optional arguments are used in conjunction with the Regression test only)
' ---------------------------------------------------------------------------------
    Const PROC = "Test_99_Individual"
    
    On Error GoTo eh
    
    '~~ Obtain initial test values from the Test Worksheet
    With fMsg
        .MsgWidthMaxSpecInPt = 300
        .DsplyFrmsWthBrdrsTestOnly = True
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text"
    End With
    With Message.Section(2)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text"
        .Text.MonoSpaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text"
    End With
    mMsg.Dsply dsply_title:="This title is rather short" _
             , dsply_msg:=Message _
             , dsply_buttons:=mMsg.Buttons("Button-1", "Button-2") _
             , dsply_button_default:="Button-1"
    
xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Debugging Then
    Stop: Resume
#End If
End Sub

