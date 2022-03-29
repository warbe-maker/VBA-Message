Attribute VB_Name = "mTestServices"
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------------
' Standard Module mTestServices
' All tests obligatory for a complete regression test performed after any code
' modification. Tests are to be extended when new features or functions are
' implemented.
'
' Note:    Test which explicitely raise an errors are only correctly asserted
'          when the error is passed on to the calling/entry procedure - which
'          requires the Conditional Compile Argument 'Debugging = 1'.
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
Dim vButtons                As Collection
Dim cllButtonsTest          As Collection

Private Property Get BTTN_TERMINATE() As String ' composed constant
    BTTN_TERMINATE = "Terminate" & vbLf & "Regression" & vbLf & "Test"
End Property

Public Property Let RegressionTest(ByVal b As Boolean)
    bRegressionTest = b
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

Private Sub BoP(ByVal b_proc As String, _
           ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Begin of Procedure stub. The service is handed over to the corresponding
' procedures in the Common mTrc Component (Execution Trace) or the Common mErH
' Component (Error Handler) provided the components are installed which is
' indicated by the corresponding Conditional Compile Arguments ErHComp = 1 and
' TrcComp = 1.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Public Sub cmdTest01_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_01_Buttons
End Sub

Public Sub cmdTest02_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_02_ErrMsg
End Sub

Public Sub cmdTest03_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_03_WidthDeterminedByMinimumWidth
End Sub

Public Sub cmdTest04_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_04_WidthDeterminedByTitle
End Sub

Public Sub cmdTest05_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_05_WidthDeterminedByMonoSpacedMessageSection
End Sub

Public Sub cmdTest06_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_06_WidthDeterminedByReplyButtons
End Sub

Public Sub cmdTest07_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_07_MonoSpacedSectionWidthExceedsMaxMsgWidth
End Sub

Public Sub cmdTest08_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_08_MonoSpacedMessageSectionExceedsMaxHeight
End Sub

Public Sub cmdTest09_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_09_ButtonsOnly
End Sub

Public Sub cmdTest10_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_10_ButtonsMatrix
End Sub

Public Sub cmdTest11_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_11_ButtonScrollBarVertical
End Sub

Public Sub cmdTest12_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_12_ButtonScrollBarHorizontal
End Sub

Public Sub cmdTest13_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_13_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
End Sub

Public Sub cmdTest16_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_16_ButtonByDictionary
End Sub

Public Sub cmdTest17_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_17_MessageAsString
End Sub

Public Sub cmdTest20_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_20_ButtonByValue
End Sub

Public Sub cmdTest21_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_21_ButtonByString
End Sub

Public Sub cmdTest22_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_22_ButtonByCollection
End Sub

Public Sub cmdTest23_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_23_MonoSpacedSectionOnly
End Sub

Public Sub cmdTest30_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_30_Monitor
End Sub

Public Sub cmdTest90_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_90_AllInOne
End Sub

Public Sub cmdTest91_Click()
    wsTest.RegressionTest = False
    mTestServices.Test_91_MinimumMessage
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' End of Procedure stub. Handed over to the corresponding procedures in the
' Common Component mTrc (Execution Trace) or mErH (Error Handler) provided the
' components are installed which is indicated by the corresponding Conditional
' Compile Arguments.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
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
    ErrSrc = "mTestServices." & sProc
End Function

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
    Dim t           As String   ' .Top
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
    Unload mMsg.MsgInstance(MsgTitle) ' Ensure there is no process monitoring with this title still displayed
    Set MsgForm = mMsg.MsgInstance(MsgTitle)
    
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
                        If v.Visible Then mDct.DctAdd dct, v, Item
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
        If ctl.Visible Then Appl = "Yes " Else Appl = " No "
        l = Align(Format(ctl.Left, "000.0"), 7, AlignCentered, " ")
        W = Align(Format(ctl.Width, "000.0"), 7, AlignCentered, " ")
        t = Align(Format(ctl.Top, "000.0"), 7, AlignCentered, " ")
        H = Align(Format(ctl.Height, "000.0"), 7, AlignCentered, " ")
        FH = Align(Format(MsgForm.InsideHeight, "000.0"), 7, AlignCentered, " ")
        FW = Align(Format(MsgForm.InsideWidth, "000.0"), 7, AlignCentered, " ")
        If TypeName(ctl) = "Frame" Then
            Set frm = ctl
            CW = Align(Format(MsgForm.ContentWidth(frm), "000.0"), 7, AlignCentered, " ")
            CH = Align(Format(MsgForm.ContentHeight(frm), "000.0"), 7, AlignCentered, " ")
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
        
        Debug.Print Align(ctl.Name, 20, AlignLeft) & "|" & Appl & "|" & l & "|" & W & "|" & CW & "|" & t & "|" & H & "|" & CH & "|" & SH & "|" & SW & "|" & FW & "|" & FH
    Next v

xt: Set dct = Nothing

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    Dim i As Long
    
    mMsg.MsgInstance fi_key:=msg_title, fi_unload:=True                    ' Ensures a message starts from scratch
    Set msg_form = mMsg.MsgInstance(msg_title)
    
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
    If bRegressionTest Then mTestServices.RegressionTest = True Else mTestServices.RegressionTest = False

End Sub

Private Function PrcPnt(ByVal pp_value As Single, _
                        ByVal pp_dimension As String) As String
    PrcPnt = mMsg.Prcnt(pp_value, pp_dimension) & "% (" & mMsg.Pnts(pp_value, "w") & "pt)"
End Function

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
    Dim c As Long
    Dim l As Long
    Dim i As Long

    l = Len(repeat_string)
    c = l * repeat_n_times
    s = Space$(c)

    For i = 1 To c Step l
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

Private Sub SetupTest(ByVal test_no As Long)
    
    wsTest.TestNumber = test_no
    
    If bRegressionTest _
    Then Set cllButtonsTest = mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TERMINATE) _
    Else Set cllButtonsTest = mMsg.Buttons(BTTN_PASSED, BTTN_FAILED)
    
End Sub

Public Sub Test_00_Regression()
' --------------------------------------------------------------------------------------
' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
    Const PROC = "Test_00_Regression"
    
    On Error GoTo eh
    Dim Rng     As Range
    Dim sTest   As String
    Dim sMakro  As String
        
    ' Test initializations
    ThisWorkbook.Save
    Unload fMsg
    wsTest.RegressionTest = True
    mTestServices.RegressionTest = True
    mErH.Regression = True
    mTrc.LogFile = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "RegressionTest.log")
    mTrc.LogTitle = "Regression test module mMsg"
    
    BoP ErrSrc(PROC)
    For Each Rng In wsTest.RegressionTests
        If Rng.Value = "R" Then
            sTest = Format(Rng.OFFSET(, -2), "00")
            sMakro = "cmdTest" & sTest & "_Click"
            wsTest.TerminateRegressionTest = False
            Application.Run "Msg.xlsb!" & sMakro
            If wsTest.TerminateRegressionTest Then Exit For
        End If
    Next Rng

xt: EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_Buttons()
    Const PROC = "Test_01_Buttons"
    BoP ErrSrc(PROC)
    mTestServices.Test_01_Buttons_01_Empty
    mTestServices.Test_01_Buttons_02_Single_String
    mTestServices.Test_01_Buttons_03_Single_Numeric_Item
    mTestServices.Test_01_Buttons_04_String_String
    mTestServices.Test_01_Buttons_05_Collection_String_String
    mTestServices.Test_01_Buttons_06_String_Collection_String
    mTestServices.Test_01_Buttons_07_String_String_Collection
    mTestServices.Test_01_Buttons_08_Semicolon_Delimited_String_Collection
    mTestServices.Test_01_Buttons_09_Comma_Delimited_String_Dictionary
    mTestServices.Test_01_Buttons_10_Box_7_By_7_Matrix
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_01_Empty()
    Const PROC = "Test_01_Buttons_01_Empty"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = Buttons()
    Debug.Assert cll.Count = 0
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_02_Single_String()
    Const PROC = "Test_01_Buttons_02_Single_String"
    Dim cll As New Collection
    
    BoP ErrSrc(PROC)
    Set cll = Buttons("aaa")
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1) = "aaa"
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_01_Buttons_03_Single_Numeric_Item() As Variant
    Const PROC = "Test_01_Buttons_03_Single_Numeric_Item"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = mMsg.Buttons(vbResumeOk)
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1) = vbResumeOk
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Function

Public Sub Test_01_Buttons_04_String_String()
    Const PROC = "Test_01_Buttons_04_String_String"
    Dim cll As New Collection
    
    BoP ErrSrc(PROC)
    Set cll = Buttons("aaa", "bbb")
    Debug.Assert cll.Count = 2
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_05_Collection_String_String()
    Const PROC = "Test_01_Buttons_05_Collection_String_String"
    Dim cll_1 As New Collection
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    cll_1.Add "aaa"
    cll_1.Add "bbb"
    
    Set cll = Buttons(cll_1, "aaa", "bbb")
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "aaa"
    Debug.Assert cll(4) = "bbb"
    
    Set cll = Nothing
    Set cll_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_06_String_Collection_String()
    Const PROC = "Test_01_Buttons_06_String_Collection_String"
    Dim cll     As Collection
    Dim cll_1   As New Collection
    
    BoP ErrSrc(PROC)
    cll_1.Add "aaa"
    cll_1.Add "bbb"
    
    Set cll = Buttons("aaa", cll_1, "bbb")
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "aaa"
    Debug.Assert cll(3) = "bbb"
    Debug.Assert cll(4) = "bbb"
    
    Set cll = Nothing
    Set cll_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_07_String_String_Collection()
    Const PROC = "Test_01_Buttons_07_String_String_Collection"
    Dim cll     As Collection
    Dim cll_1   As New Collection
    
    BoP ErrSrc(PROC)
    cll_1.Add "ccc"
    cll_1.Add "ddd"
    
    Set cll = Buttons("aaa", "bbb", cll_1)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    Set cll_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_08_Semicolon_Delimited_String_Collection()
    Const PROC = "Test_01_Buttons_08_Semicolon_Delimited_String_Collection"
    Dim cll     As Collection
    Dim cll_1   As New Collection
    
    BoP ErrSrc(PROC)
    cll_1.Add "ccc"
    cll_1.Add "ddd"
    
    Set cll = Buttons("aaa;bbb", cll_1)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    Set cll_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Sub Test_01_Buttons_09_Comma_Delimited_String_Dictionary()
    Const PROC = "Test_01_Buttons_09_Comma_Delimited_String_Dictionary"
    Dim cll     As Collection
    Dim dct   As New Dictionary
    
    BoP ErrSrc(PROC)
    dct.Add "ccc", "ccc"
    dct.Add "ddd", "ddd"
    
    Set cll = Buttons("aaa,bbb", dct)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    Set dct = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_01_Buttons_10_Box_7_By_7_Matrix() As Variant
' ------------------------------------------------------------------------------
' The Buttons service "in action": Display a matrix of 7 x 7 buttons
' ------------------------------------------------------------------------------
    Const PROC = "Function Test_01_Buttons_10_Box_7_By_7_Matrix"
    
    Dim cll As New Collection
    Dim i As Long
    
    BoP ErrSrc(PROC)
    SetupTest 1
    
    For i = 1 To 49
        cll.Add "B" & Format(i, "00")
    Next i
    Set cll = mMsg.Buttons(cllButtonsTest, cll) ' excessive buttons are ignored !
    Debug.Assert cll.Count = 55
    Debug.Assert cll(8) = vbLf
    Debug.Assert cll(16) = vbLf
    Debug.Assert cll(24) = vbLf
    Debug.Assert cll(32) = vbLf
    Debug.Assert cll(40) = vbLf
    Debug.Assert cll(48) = vbLf
    
    Test_01_Buttons_10_Box_7_By_7_Matrix = _
    mMsg.Box(Prompt:=vbNullString _
           , Buttons:=cll _
           , Title:="49 buttons ordered in 7 rows, row breaks inserted and excessive buttons ignored by the 'mMsg.Buttons' service")
    EoP ErrSrc(PROC)
    
End Function

Public Sub Test_02_ErrMsg()
' ------------------------------------------------------------------------------
' Test of the "universal error message display which includes
' - the 'Debugging Option' activated by the Conditional Compile Argument
'   'Debugging = 1')
' - an optional additional "about the error" information which may be
'   concatenated with an error message by two vertical bars (||)".
' All tests primarily use the 'Private Function ErrMsg' which passes on the
' display of the error message to the ErrMsg function of the mMsg module when
' the Conditional Compile Argument 'CompMsg = 1' or passes on the function to
' the ErrMsg function of the mErH module when the Conditional Compile Argument
' 'CompErH = 1'.
' Summarized all this means that testing has to be performed with the following
' three Conditional Compile Argument variants:
' ErHComp = 0 : MsgComp = 0 > display of the error message by VBA.MsgBox
' ErHComp = 0 : MsgComp = 1 > display of the error message by mMsg.ErrMsg
' ErHComp = 1               > display of the error message by mErH.ErrMsg
' For the last testing variant the mErH component is installed!
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_ErrMsg"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    SetupTest 1
    
    mErH.Asserted AppErr(5) ' skip error message display when mErH.Regression = True
    
    Err.Raise Number:=AppErr(5), source:=ErrSrc(PROC), _
              Description:="This is a test error description!||This is part of the error description, " & _
                           "concatenated by a double vertical bar and therefore displayed as an additional 'About the error' section " & _
                           "- one of the specific features of the mMsg.ErrMsg service."
        
xt: EoP ErrSrc(PROC)
    Select Case mMsg.Box(Title:="Test result of " & Readable(PROC) _
                       , Prompt:=vbNullString _
                       , Buttons:=mMsg.Buttons(cllButtonsTest) _
                        )
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function Test_03_WidthDeterminedByMinimumWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_03_WidthDeterminedByMinimumWidth"
    
    On Error GoTo eh
    Dim MsgForm         As fMsg
    Dim MsgTitle        As String
    Dim cll             As Collection
    
    BoP ErrSrc(PROC)
    SetupTest 3
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
        MsgForm.VisualizeForTest = wsTest.VisualizeForTest
        TestMsgWidthIncrDecr = .MsgWidthIncrDecr
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    
    vButton4 = "Repeat with minimum width" & vbLf & "+ " & PrcPnt(TestMsgWidthIncrDecr, "w")
    vButton5 = "Repeat with minimum width" & vbLf & "- " & PrcPnt(TestMsgWidthIncrDecr, "w")
    
    Set cll = mMsg.Buttons(cllButtonsTest, vbLf, vButton4, vButton5)
    
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
                                                                                                  
        Test_03_WidthDeterminedByMinimumWidth = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cll _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                  )
        Select Case Test_03_WidthDeterminedByMinimumWidth
            Case vButton5
                TestMsgWidthMin = TestMsgWidthMin - TestMsgWidthIncrDecr
                Set cll = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4, vButton5)
            Case vButton4
                TestMsgWidthMin = TestMsgWidthMin + mMsg.Pnts(TestMsgWidthIncrDecr, "W")
                Set cll = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, vButton4, vButton5)
            Case BTTN_PASSED:       wsTest.Passed = True:   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
            Case Else ' Stop and Next are passed on to the caller
        End Select
    
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_04_WidthDeterminedByTitle() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_04_WidthDeterminedByTitle"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 4
    MsgTitle = Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications

    '~~ Obtain initial test values from the Test Worksheet
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
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
    Set vButtons = mMsg.Buttons(cllButtonsTest)
    
    Test_04_WidthDeterminedByTitle = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vButtons _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_04_WidthDeterminedByTitle
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_05_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_05_WidthDeterminedByMonoSpacedMessageSection"
        
    On Error GoTo eh
    Dim MsgForm                         As fMsg
    Dim MsgTitle                        As String
    Dim BttnRepeatMaxWidthIncreased     As String
    Dim BttnRepeatMaxWidthDecreased     As String
    Dim BttnRepeatMaxHeightIncreased    As String
    Dim BttnRepeatMaxHeightDecreased    As String
    
    BoP ErrSrc(PROC)
    SetupTest 5
    MsgTitle = Readable(PROC)
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = Pnts(.MsgWidthMin, "w")
        TestMsgWidthMax = Pnts(.MsgWidthMax, "w")
        TestMsgWidthIncrDecr = Pnts(.MsgWidthIncrDecr, "w")
        TestMsgHeightMin = Pnts(25, "h")
        TestMsgHeightMax = Pnts(.MsgHeightMax, "h")
        TestMsgHeightIncrDecr = Pnts(.MsgHeightIncrDecr, "h")
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    BttnRepeatMaxWidthIncreased = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & PrcPnt(TestMsgWidthIncrDecr, "w")
    BttnRepeatMaxWidthDecreased = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & PrcPnt(TestMsgWidthIncrDecr, "w")
    BttnRepeatMaxHeightIncreased = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & PrcPnt(TestMsgHeightIncrDecr, "h")
    BttnRepeatMaxHeightDecreased = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & PrcPnt(TestMsgHeightIncrDecr, "h")
    
    Set vButtons = mMsg.Buttons(cllButtonsTest, vbLf, BttnRepeatMaxWidthIncreased, BttnRepeatMaxWidthDecreased)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    Do
        AssertWidthAndHeight TestMsgWidthMin _
                           , TestMsgWidthMax _
                           , TestMsgHeightMin _
                           , TestMsgHeightMax
        
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
        mMsg.MsgInstance(MsgTitle).VisualizeForTest = wsTest.VisualizeForTest
                
        Test_05_WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vButtons _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_min:=TestMsgHeightMin _
                 , dsply_height_max:=TestMsgHeightMax _
                  )
        Select Case Test_05_WidthDeterminedByMonoSpacedMessageSection
            Case BttnRepeatMaxWidthDecreased
                TestMsgWidthMax = TestMsgWidthMax - TestMsgWidthIncrDecr
                Set vButtons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthIncreased, BttnRepeatMaxWidthDecreased)
            Case BttnRepeatMaxWidthIncreased
                TestMsgWidthMax = TestMsgWidthMax + TestMsgWidthIncrDecr
                Set vButtons = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED, vbLf, BttnRepeatMaxWidthIncreased, BttnRepeatMaxWidthDecreased)
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do ' Stop, Previous, and Next are passed on to the caller
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_06_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_06_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim OneBttnMore As String
    Dim OneBttnLess As String
    
    BoP ErrSrc(PROC)
    SetupTest 6
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    ' Initializations for this test
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
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
    OneBttnMore = "Repeat with one button more"
    OneBttnLess = "Repeat with one button less"
    vButton6 = "The one more buttonn"
    
    Set vButtons = mMsg.Buttons(cllButtonsTest, vbLf, OneBttnLess, vButton6)
    
    Do
        Test_06_WidthDeterminedByReplyButtons = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vButtons _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_06_WidthDeterminedByReplyButtons
            Case OneBttnMore
                Set vButtons = mMsg.Buttons(cllButtonsTest, vbLf, OneBttnLess, vButton6)
            Case OneBttnLess
                Set vButtons = mMsg.Buttons(cllButtonsTest, vbLf, OneBttnMore)
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do

        End Select
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_07_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_07_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 7
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
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
    Set vButtons = mMsg.Buttons(cllButtonsTest)
    
    Test_07_MonoSpacedSectionWidthExceedsMaxMsgWidth = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vButtons _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_07_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
    
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_08_MonoSpacedMessageSectionExceedsMaxHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_08_MonoSpacedMessageSectionExceedsMaxHeight"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 8
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
       
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The overall message window height exceeds the for this test specified maximum of " & _
                     PrcPnt(TestMsgHeightMax, "h") & " of the screen height. Because the monospaced section " & _
                     "is the dominating one regarding its height it is displayed with a horizontal scroll-bar."
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
    Set vButtons = mMsg.Buttons(cllButtonsTest)
    
    Test_08_MonoSpacedMessageSectionExceedsMaxHeight = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=vButtons _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    Select Case Test_08_MonoSpacedMessageSectionExceedsMaxHeight
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_09_ButtonsOnly() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_09_ButtonsOnly"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim i           As Long
    Dim j           As Long
    Dim cllStory    As New Collection
    Dim vReply      As Variant
    Dim bMonospaced As Boolean: bMonospaced = True ' initial test value
    
    BoP ErrSrc(PROC)
    SetupTest 9
    MsgTitle = Readable(PROC) & ": No message, just buttons (finish with " & BTTN_PASSED & " or " & BTTN_FAILED & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMax = .MsgWidthMax:     TestMsgWidthIncrDecr = .MsgWidthIncrDecr
        TestMsgWidthMin = .MsgWidthMin:     TestMsgHeightIncrDecr = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax
    End With
    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 1 To 4 ' rows
        If i > 1 Then cllStory.Add vbLf
        For j = 1 To 3
            cllStory.Add "Click " & i & "-" & j & " in case ...." & vbLf & "(instead of a lengthy" & vbLf & "message text above)"
        Next j
    Next i
    Set cllStory = mMsg.Buttons(cllButtonsTest, vbLf, cllStory)
    Do
        mMsg.MsgInstance(MsgTitle).VisualizeForTest = wsTest.VisualizeForTest
        '~~ Obtain initial test values from the Test Worksheet
                         
        Test_09_ButtonsOnly = _
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
        Select Case Test_09_ButtonsOnly
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case "Ok":                                                      Exit Do ' The very last item in the collection is the "Finished" button
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do

        End Select
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_10_ButtonsMatrix() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_10_ButtonsMatrix"
    
    On Error GoTo eh
    Dim MsgForm             As fMsg
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim i, j                As Long
    Dim MsgTitle            As String
    Dim cllMatrix           As Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
        
    BoP ErrSrc(PROC)
    SetupTest 10
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMax = .MsgWidthMax:   lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
'    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
'    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    MsgTitle = "Just to demonstrate what's theoretically possible: Buttons only! Finish with " & BTTN_PASSED & " (default) or " & BTTN_FAILED
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications

    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    Set cllMatrix = New Collection
    For i = 2 To 7 ' rows
        For j = 1 To 7 ' row buttons
            cllMatrix.Add "Button" & vbLf & i & "-" & j
        Next j
    Next i
    Set cllMatrix = mMsg.Buttons(cllButtonsTest, vbLf, cllMatrix)
    
    Do
        '~~ Obtain initial test values from the Test Worksheet
        mMsg.MsgInstance(MsgTitle).VisualizeForTest = wsTest.VisualizeForTest
                             
        Test_10_ButtonsMatrix = _
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
            
        Select Case Test_10_ButtonsMatrix
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_11_ButtonScrollBarVertical() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_11_ButtonScrollBarVertical"
    
    On Error GoTo eh
    Dim MsgForm             As fMsg
    Dim MsgTitle            As String
    Dim i, j                As Long
    Dim cll                 As New Collection
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    
    BoP ErrSrc(PROC)
    SetupTest 11
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    With wsTest
        TestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        TestMsgWidthMax = .MsgWidthMax:     lChangeWidthPcntg = .MsgWidthIncrDecr
        TestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
'    If TestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
'    If TestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Obtain initial test values from the Test Worksheet
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
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
        If i > 1 Then Set cll = mMsg.Buttons(cll, vbLf)
        For j = 1 To 2
            Set cll = mMsg.Buttons(cll, "Reply" & vbLf & "Button" & vbLf & i & "-" & j)
        Next j
    Next i
    Set cll = mMsg.Buttons(cllButtonsTest, vbLf, cll)
    
    Do
        Test_11_ButtonScrollBarVertical = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=cll _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_height_max:=TestMsgHeightMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                  )
        Select Case Test_11_ButtonScrollBarVertical
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_12_ButtonScrollBarHorizontal() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_12_ButtonScrollBarHorizontal"
    Const INIT_WIDTH    As String = 40
    Const CHANGE_WIDTH  As String = 10
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim Bttn10Plus  As String
    Dim Bttn10Minus As String
    
    BoP ErrSrc(PROC)
    SetupTest 12
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    TestMsgWidthMax = INIT_WIDTH
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With

    Do
        mMsg.MsgInstance(MsgTitle).VisualizeForTest = wsTest.VisualizeForTest
        
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
        
        Bttn10Plus = "Repeat with maximum form width" & vbLf & "extended by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(TestMsgWidthMax, "w")
        Bttn10Minus = "Repeat with maximum form width" & vbLf & "reduced by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(TestMsgWidthMax, "w")
            
        '~~ Obtain initial test values from the Test Worksheet
    
        Set vButtons = mMsg.Buttons(cllButtonsTest, vbLf, Bttn10Plus, Bttn10Minus)
        Test_12_ButtonScrollBarHorizontal = _
        mMsg.Dsply(dsply_title:=MsgTitle _
                 , dsply_msg:=Message _
                 , dsply_buttons:=vButtons _
                 , dsply_width_min:=TestMsgWidthMin _
                 , dsply_width_max:=TestMsgWidthMax _
                 , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                 , dsply_button_default:=BTTN_PASSED _
                  )
        Select Case Test_12_ButtonScrollBarHorizontal
            Case Bttn10Minus:       TestMsgWidthMax = TestMsgWidthMax - CHANGE_WIDTH
            Case Bttn10Plus:        TestMsgWidthMax = TestMsgWidthMax + CHANGE_WIDTH
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop
    
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_13_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_13_ButtonsMatrix_Horizontal_and_Vertical_Scrollbar"
    
    On Error GoTo eh
    Dim MsgForm                 As fMsg
    Dim i, j                    As Long
    Dim MsgTitle                As String
    Dim cllMatrix               As Collection
    Dim bMonospaced             As Boolean: bMonospaced = True ' initial test value
    Dim TestMsgWidthMin         As Long
    Dim TestMsgWidthMaxSpecInPt As Long
    Dim TestMsgHeightMax        As Long
    
    BoP ErrSrc(PROC)
    SetupTest 13
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
    For i = 2 To 7 ' rows
        For j = 1 To 7 ' row buttons
            cllMatrix.Add vbLf & " ---- Button ---- " & vbLf & i & "-" & j & vbLf & " "
        Next j
    Next i
    Set cllMatrix = mMsg.Buttons(cllButtonsTest, vbLf, cllMatrix)
    
    Do
        '~~ Obtain initial test values from the Test Worksheet
        mMsg.MsgInstance(MsgTitle).VisualizeForTest = wsTest.VisualizeForTest
                             
        Test_13_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar = _
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
        Select Case Test_13_ButtonsMatrix_with_horizomtal_and_vertical_scrollbar
            Case BTTN_PASSED:       wsTest.Passed = True:                   Exit Do
            Case BTTN_FAILED:       wsTest.Failed = True:                   Exit Do
            Case sBttnTerminate:    wsTest.TerminateRegressionTest = True:  Exit Do
        End Select
    Loop

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_16_ButtonByDictionary()
' ------------------------------------------------------------------------------
' The buttons argument is provided as Dictionary.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_16_ButtonByDictionary"
    
    Dim dct         As New Collection
    Dim MsgTitle    As String
    Dim MsgForm     As fMsg
    
    BoP ErrSrc(PROC)
    SetupTest 16
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is a collection of the test specific buttons " & _
                     "(Passed, Failed) and the two extra Yes, No buttons provided as Dictionary!" & vbLf & vbLf & _
                     "The test proves that the mMsg.Buttons service is able to combine any kind of arguments " & _
                     "provided via the ParamArray."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    dct.Add "Yes"
    dct.Add vbLf
    dct.Add "No"
    
    Test_16_ButtonByDictionary = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=mMsg.Buttons(cllButtonsTest, vbLf, dct) _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_17_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_17_Box_MessageAsString"
        
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 17
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    Set vButtons = mMsg.Buttons(cllButtonsTest)
        
    Test_17_MessageAsString = _
    mMsg.Box( _
             Title:=MsgTitle _
           , Prompt:="This is a message provided as a simple string argument!" _
           , Buttons:=vButtons _
           , box_width_min:=TestMsgWidthMin _
           , box_width_max:=TestMsgWidthMax _
           , box_height_max:=TestMsgHeightMax _
            )
    Select Case Test_17_MessageAsString
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_20_ButtonByValue()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_20_ButtonByValue"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle   As String
    
    BoP ErrSrc(PROC)
    SetupTest 20
    MsgTitle = "Test: Button by value (" & PROC & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is a collection of the test buttons (Passed, Failed) and an additional button provided as value"
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The ""Ok"" button is displayed in the second row."
    End With
    Test_20_ButtonByValue = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=mMsg.Buttons(cllButtonsTest, vbLf, vbOKOnly) _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
            
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_21_ButtonByString()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_21_ButtonByString"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 21
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
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

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_22_ButtonByCollection()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_22_ButtonByCollection"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim cll         As New Collection
    
    BoP ErrSrc(PROC)
    SetupTest 22
    MsgTitle = "Test: Button by value (" & ErrSrc(PROC) & ")"
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
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

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_23_MonoSpacedSectionOnly()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_23_MonoSpacedSectionOnly"
    Const LINES = 50
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim cll         As New Collection
    Dim Msg         As String
    Dim i           As Long

    BoP ErrSrc(PROC)
    SetupTest 23
    MsgTitle = "Test: Monospaced section with " & LINES & " lines exceeding the specified max height (" & ErrSrc(PROC) & ")"
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    Set cll = mMsg.Buttons(sBttnTerminate, BTTN_PASSED, BTTN_FAILED)
            
    i = 1
    Msg = Format(i, "00: ") & Format(Now(), "YY-MM-DD hh:mm:ss") & " Test mono-spaced message section text exceeding the specified maximum width and height"
    For i = 2 To LINES
        Msg = Msg & vbLf & Format(i, "00: ") & Format(Now(), "YY-MM-DD hh:mm:ss") & " Test mono-spaced message section text exceeding the specified maximum width and height"
    Next i
    
    With Message.Section(1).Text
        .Text = Msg
        .MonoSpaced = True
    End With
    
    Test_23_MonoSpacedSectionOnly = _
    mMsg.Dsply(dsply_title:=MsgTitle _
             , dsply_msg:=Message _
             , dsply_buttons:=cll _
             , dsply_width_min:=TestMsgWidthMin _
             , dsply_width_max:=TestMsgWidthMax _
             , dsply_height_max:=TestMsgHeightMax _
             , dsply_modeless:=wsTest.TestOptionDisplayModeless _
              )
    
    Select Case Test_23_MonoSpacedSectionOnly
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_30_Monitor() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_30_Monitor"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    Dim i           As Long
    Dim Header      As TypeMsgText
    Dim Step        As TypeMsgText
    Dim Footer      As TypeMsgText
    Dim iLoops      As Long
    Dim lWait       As Long
    
    BoP ErrSrc(PROC)
    SetupTest 30
    TestMsgWidthMin = wsTest.MsgWidthMin
    TestMsgWidthMax = wsTest.MsgWidthMax
    TestMsgHeightMax = wsTest.MsgHeightMax
    
    With Header
        .Text = "Step Status"
        .MonoSpaced = True
        .FontColor = rgbRed
    End With
    iLoops = 15
    lWait = 300
    
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    For i = 1 To iLoops
        With Step
            .Text = Format(i, "00") & ". Follow-Up line after " & Format(lWait, "0000") & " Milliseconds."
            .Text = Repeat(.Text & " ", Int(i / 5) + 1) & vbLf & "    Second line just for test " & Repeat(".", i)
            .MonoSpaced = True
        End With
        mMsg.Monitor mon_title:=MsgTitle _
                   , mon_header:=Header _
                   , mon_step:=Step _
                   , mon_steps_visible:=10 _
                   , mon_footer:=Footer _
                   , mon_width_max:=TestMsgWidthMax _
                   , mon_width_min:=TestMsgWidthMin _
                   , mon_height_max:=TestMsgHeightMax
                   
        '~~ Simmulation of a process
        DoEvents
        Sleep lWait
    Next i
    Step.Text = vbNullString
    With Footer
        .Text = "Process finished! Close this window"
        .FontBold = True
        .FontColor = rgbBlue
    End With
    mMsg.Monitor mon_title:=MsgTitle _
               , mon_header:=Header _
               , mon_step:=Step _
               , mon_footer:=Footer _
               , mon_width_max:=wsTest.MsgWidthMax _
               , mon_width_min:=wsTest.MsgWidthMin _
               , mon_height_max:=wsTest.MsgHeightMax
    
    MsgTitle = "Test result of: " & Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    Set vButtons = mMsg.Buttons(BTTN_PASSED, BTTN_FAILED)
    Select Case mMsg.Box(Title:=MsgTitle _
                       , Prompt:=vbNullString _
                       , Buttons:=vButtons _
                        )
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select

xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_90_AllInOne() As Variant
    Const PROC      As String = "Test_90_AllInOne"

    Dim MsgTitle    As String
    Dim cll         As New Collection
    Dim i, j        As Long
    Dim Msg         As TypeMsg
    Dim MsgForm     As fMsg
    
    SetupTest 90
    MsgTitle = Readable(PROC)
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC) ' set test-global message specifications
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
    
    With Msg.Section(1)
        .Label.Text = "Service features used by this displayed message:"
        .Label.FontColor = rgbBlue
        .Text.Text = "All 4 message sections, and all with a label, monospaced option for the second section, " _
                   & "some of the 7 x 7 reply buttons in a 4-4-1 order, font color option for all labels."
    End With
    With Msg.Section(2)
        .Label.Text = "Demonstration of the unlimited message width:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section's text is mono-spaced (which by definition is not word-wrapped)" & vbLf _
                   & "the message width is determined by:" & vbLf _
                   & "a) the for this demo specified maximum width of " & TestMsgWidthMax & "% of the screen size" & vbLf _
                   & "   (defaults to 80% when not specified)" & vbLf _
                   & "b) the longest line of this section" & vbLf _
                   & "Because the text exeeds the specified maximum message width, a horizontal scroll-bar is displayed." & vbLf _
                   & "Due to this feature there is no message size limit other than the sytem's limit which for a string is about 1GB !!!!"
        .Text.MonoSpaced = True
    End With
    With Msg.Section(3)
        .Label.Text = "Unlimited message height (not the fact with this message):"
        .Label.FontColor = rgbBlue
        .Text.Text = "As with the message width, the message height is unlimited. When the maximum height (explicitely specified or the default) " _
                   & "is exceeded a vertical scroll-bar is displayed. Due to this feature there is no message size limit other than the sytem's " _
                   & "limit which for a string is about 1GB !!!!"
    End With
    With Msg.Section(4)
        .Label.Text = "Flexibility regarding the displayed reply buttons:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 7 x 7 = 49 possible reply buttons which may have any caption text " _
                   & "including the classic VBA.MsgBox values (vbOkOnly, vbYesNoCancel, etc.) - even in a mixture." & vbLf & vbLf _
                   & "!! This demo ends only with the Ok button and loops with any other."
    End With
    '~~ Prepare the buttons collection
    
    For j = 1 To 2
        If j > 1 Then cll.Add vbLf
        For i = 1 To 4
            cll.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
    Next j
    Set cll = mMsg.Buttons(cllButtonsTest, vbLf, cll)
    
    Select Case mMsg.Dsply(dsply_title:=MsgTitle _
                         , dsply_msg:=Msg _
                         , dsply_buttons:=cll _
                         , dsply_width_min:=TestMsgWidthMin _
                         , dsply_width_max:=TestMsgWidthMax _
                         , dsply_height_max:=TestMsgHeightMax _
                         , dsply_modeless:=wsTest.TestOptionDisplayModeless _
                          )
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
    
End Function

Public Function Test_91_MinimumMessage() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_91_MinimumMessage"
    
    On Error GoTo eh
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    BoP ErrSrc(PROC)
    SetupTest 91
    MsgTitle = Readable(PROC)
    MessageInit msg_form:=MsgForm, msg_title:=MsgTitle, caller:=ErrSrc(PROC)  ' set test-global message specifications
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        TestMsgWidthMin = .MsgWidthMin
        TestMsgWidthMax = .MsgWidthMax
        TestMsgHeightMax = .MsgHeightMax
    End With
    MsgForm.VisualizeForTest = wsTest.VisualizeForTest
        
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
                                                                                              
    Select Case mMsg.Dsply(dsply_title:=MsgTitle _
                         , dsply_msg:=Message _
                         , dsply_buttons:=mMsg.Buttons() _
                         , dsply_width_min:=TestMsgWidthMin _
                         , dsply_width_max:=TestMsgWidthMax _
                         , dsply_height_max:=TestMsgHeightMax _
                         , dsply_modeless:=wsTest.TestOptionDisplayModeless)
        Case BTTN_PASSED:       wsTest.Passed = True
        Case BTTN_FAILED:       wsTest.Failed = True
        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
    End Select
             
xt: EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

