Attribute VB_Name = "mMsgTestServices"
Option Explicit
Option Compare Text
' ------------------------------------------------------------------------------
' Standard Module mMsgTestServices: All tests obligatory for a complete test of
' ================================= regression test of all kind of message
' all services and features, performed after any code modification. It goes
' without saying that test procedures are to be extended, ammended, or modified
' in case of when new implemented features, methods, or functions or in case
' an error has been dedected which was not covered by a test.
'
' Note: - All test procedures (except "Test_12_mMsg_ErrMsg_AppErr_5") display the
' -----   message modeless - regardless the option set - with a "Passed",
'         "Failed", and a "Terminate" button waiting for either of the three is
'         pressed.
'       - For the Regression test (Test_10_Regression) explicitly raised errors
'         are asserted beforehand in order not to interrupt the regression test
'         procedure. This is achived by `mErH.Regression = True` and
'         `mErH.Asserted AppErr(n)` for 'awaited' respectively tested
'         application errors.
'       - Any loops with modified arguments like min and max width and height
'         or the LabelPosSpec are to be implemented by means of button with
'         AppRun arguments, modifying "global" argument values and re-executing
'         the current test-procedure.
'
' W. Rauschenberger, Berlin Aug 2023
' -------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

#If XcTrc_clsTrc = 1 Then
    Private Trc                 As New clsTrc
#End If
Private Const DFLT_SECT_TEXT_PROP   As String = ">Lorem ipsum dolor sit amet, consectetur adipiscing elit, " & _
                                                "sed do eiusmod tempor incididunt ut labore et dolore magna " & _
                                                "aliqua. Ut enim ad minim veniam, quis nostrud exercitation " & _
                                                "ullamco laboris nisi ut aliquip ex ea commodo consequat. " & _
                                                "Duis aute irure dolor in reprehenderit in voluptate velit " & _
                                                "esse cillum dolore eu fugiat nulla pariatur. Excepteur sint " & _
                                                "occaecat cupidatat non proident, sunt in culpa qui officia " & _
                                                "deserunt mollit anim id est laborum.<"
Private Const DFLT_SECT_TEXT_MONO   As String = ">Lorem ipsum dolor sit amet, consectetur adipiscing elit, " & vbLf & _
                                                "sed do eiusmod tempor incididunt ut labore et dolore magna " & vbLf & _
                                                "aliqua." & vbLf & _
                                                "Ut enim ad minim veniam, quis nostrud exercitation " & _
                                                "ullamco laboris nisi ut aliquip ex ea commodo consequat." & vbLf & _
                                                "Duis aute irure dolor in reprehenderit in voluptate velit " & _
                                                "esse cillum dolore eu fugiat nulla pariatur." & vbLf & _
                                                "Excepteur sint occaecat cupidatat non proident, sunt in culpa " & _
                                                "qui officia deserunt mollit anim id est laborum.<"
Private vButton4                As Variant
Private vButton5                As Variant
Private vButton6                As Variant
Private sReadableTestProc       As String

Private Property Get DefaultSectionTextMono() As String
    DefaultSectionTextMono = Replace(DFLT_SECT_TEXT_PROP, ". ", "." & vbLf)
End Property

Private Property Get DefaultSectionTextProp() As String: DefaultSectionTextProp = DFLT_SECT_TEXT_PROP:  End Property

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

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.        | Installed component |
' |-------------------------|---------------------|
' | XcTrc_mTrc = 1          | mTrc                |
' | XcTrc_clsTrc = 1        | clsTrc              |
' | ErHComp = 1             | mErH                |
' I.e. both components are independant from each other!
' Note: This procedure is obligatory for any VB-Component using either the
'       the 'Common VBA Error Services' and/or the 'Common VBA Execution Trace
'       Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    Debug.Print "BoP: " & b_proc
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")

#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf XcTrc_clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf XcTrc_mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Public Sub cmdTest11_Click():   mMsgTestServices.Test_11_mMsg_Box_Buttons_Only:                                 End Sub

Public Sub cmdTest12_Click():   mMsgTestServices.Test_12_mMsg_ErrMsg_AppErr_5:                                  End Sub

Public Sub cmdTest13_Click():   mMsgTestServices.Test_13_mMsg_Dsply_WidthDeterminedByMinimumWidth:              End Sub

Public Sub cmdTest14_Click():   mMsgTestServices.Test_14_mMsg_Dsply_Width_Determined_By_This_eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeextra_long_Title:                    End Sub

Public Sub cmdTest15_Click():   mMsgTestServices.Test_15_mMsg_Dsply_WidthDeterminedByMonoSpacedMessageSection:  End Sub

Public Sub cmdTest16_Click():   mMsgTestServices.Test_16_mMsg_Dsply_WidthDeterminedByReplyButtons:              End Sub

Public Sub cmdTest17_Click():   mMsgTestServices.Test_17_mMsg_Dsply_MonoSpacedSectionWidthExceedsMaxMsgWidth:   End Sub

Public Sub cmdTest18_Click():   mMsgTestServices.Test_18_mMsg_Dsply_MonoSpacedMessageSectionExceedsMaxHeight:   End Sub

Public Sub cmdTest19_Click():   mMsgTestServices.Test_19_mMsg_Dsply_ButtonsOnly:                                End Sub

Public Sub cmdTest20_Click():   mMsgTestServices.Test_20_mMsg_Dsply_ButtonsMatrix:                              End Sub

Public Sub cmdTest21_Click():   mMsgTestServices.Test_21_mMsg_Dsply_ButtonScrollBarVertical:                    End Sub

Public Sub cmdTest23_Click():   mMsgTestServices.Test_23_mMsg_Dsply_Buttons_Only:                               End Sub

Public Sub cmdTest26_Click():   mMsgTestServices.Test_26_mMsg_Dsply_ButtonByDictionary:                         End Sub

Public Sub cmdTest27_Click():   mMsgTestServices.Test_27_mMsg_Box_MessageAsString:                              End Sub

Public Sub cmdTest30_Click():   mMsgTestServices.Test_30_mMsg_Dsply_ButtonByValue:                              End Sub

Public Sub cmdTest33_Click():   mMsgTestServices.Test_33_mMsg_Dsply_Single_MonoSpaced_Section_Without_Label:    End Sub

Public Sub cmdTest34_Click():   mMsgTestServices.Test_34_mMsg_Dsply_Sections_Without_Label_Or_Label_Only:       End Sub

Public Sub cmdTest40_Click():   mMsgTestServices.Test_40_mMsg_Monitor_Services:                                 End Sub

Public Sub cmdTest50_Click():   mMsgTestServices.Test_50_mMsg_Dsply_LabelPos_Left_R30:                          End Sub

Public Sub cmdTest90_Click():   mMsgTestServices.Test_90_mMsg_Dsply_AllInOne:                                   End Sub

Public Sub cmdTest91_Click():   mMsgTestServices.Test_91_mMsg_Dsply_MinimumMessage:                             End Sub

Public Sub cmdTest92_Click():   mMsgTestServices.Test_92_mMsg_Dsply_LabelWithUnderlayedURL:                     End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.        | Installed component |
'         |-------------------------|---------------------|
'         | XcTrc_mTrc = 1          | mTrc                |
'         | XcTrc_clsTrc = 1        | clsTrc              |
'         | ErHComp = 1             | mErH                |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
    Debug.Print "EoP: " & e_proc
#If ErHComp = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf XcTrc_clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf XcTrc_mTrc = 1 Then
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
' Common VBA udtMessage Display Component (mMsg) installed (Conditional Compile
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
    '~~ When only the Common udtMessage Services Component (mMsg) is installed but
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
    ErrSrc = "mMsgTestServices." & sProc
End Function

Public Function Test_10_Regression() As Variant
' --------------------------------------------------------------------------------------
' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
    Const PROC = "Test_10_Regression"
    
    On Error GoTo eh
    Dim Rng     As Range
    Dim sTest   As String
    Dim sMakro  As String
        
    ' Test initializations
    ThisWorkbook.Save
    Unload fMsg
    wsTest.RegressionTest = True
    mErH.Regression = True
    mTrc.FileName = "RegressionTest.ExecTrace.log"
    mTrc.Title = "Regression test module mMsg"
    mTrc.NewFile
    
    BoP ErrSrc(PROC)
    mTest.Current = vbNullString
    Test_10_Regression = mTest.TestProc(wsTest.NextTestNumber)

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_11_mMsg_Box_Buttons_Only() As Variant
' ------------------------------------------------------------------------------
' The Buttons service "in action": Display a matrix of 7 x 7 buttons
' ------------------------------------------------------------------------------
    Const PROC = "Test_11_mMsg_Box_Buttons_Only"
        
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 11, PROC
        
    
    mMsg.Box Prompt:=vbNullString _
           , Buttons:=mTest.BttnsOnly _
           , Title:=sMsgTitle _
           , box_modeless:=mTest.MODE_LESS _
           , box_width_min:=wsTest.MsgWidthMin _
           , box_width_max:=wsTest.MsgWidthMax _
           , box_height_max:=wsTest.MsgHeightMax

    
xt: mTest.Evaluate
    EoP ErrSrc(PROC)
    
End Function

Private Sub Test_11_mMsg_Buttons_Service()
    Const PROC = "Test_11_mMsg_Buttons_Service"
    BoP ErrSrc(PROC)
    mMsgTestServices.Test_11_mMsg_Buttons_Service_01_Empty
    mMsgTestServices.Test_11_mMsg_Buttons_Service_02_Single_String
    mMsgTestServices.Test_11_mMsg_Buttons_Service_03_Single_Numeric_Item
    mMsgTestServices.Test_11_mMsg_Buttons_Service_04_String_String
    mMsgTestServices.Test_11_mMsg_Buttons_Service_05_Collection_String_String
    mMsgTestServices.Test_11_mMsg_Buttons_Service_06_String_Collection_String
    mMsgTestServices.Test_11_mMsg_Buttons_Service_07_String_String_Collection
    mMsgTestServices.Test_11_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection
    mMsgTestServices.Test_11_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary
    mMsgTestServices.Test_11_mMsg_Box_Buttons_Only
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_01_Empty()
    Const PROC = "Test_11_mMsg_Buttons_Service_01_Empty"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = mMsg.Buttons()
    Debug.Assert cll.Count = 0
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_02_Single_String()
    Const PROC = "Test_11_mMsg_Buttons_Service_02_Single_String"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = mMsg.Buttons("aaa")
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1) = "aaa"
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_11_mMsg_Buttons_Service_03_Single_Numeric_Item() As Variant
    Const PROC = "Test_11_mMsg_Buttons_Service_03_Single_Numeric_Item"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = mMsg.Buttons(vbResumeOk)
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1) = vbResumeOk
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Function

Private Sub Test_11_mMsg_Buttons_Service_04_String_String()
    Const PROC = "Test_11_mMsg_Buttons_Service_04_String_String"
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Set cll = mMsg.Buttons("aaa", "bbb")
    Debug.Assert cll.Count = 2
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_05_Collection_String_String()
    Const PROC = "Test_11_mMsg_Buttons_Service_05_Collection_String_String"
    Dim cll As New Collection
    
    BoP ErrSrc(PROC)
    cll.Add "aaa"
    cll.Add "bbb"
    
    Set cll = mMsg.Buttons(cll, "aaa", "bbb")
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "aaa"
    Debug.Assert cll(4) = "bbb"
    
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_06_String_Collection_String()
    Const PROC = "Test_11_mMsg_Buttons_Service_06_String_Collection_String"
    Dim cll  As New Collection
    
    BoP ErrSrc(PROC)
    cll.Add "aaa"
    cll.Add "bbb"
    
    Set cll = mMsg.Buttons("aaa", cll, "bbb")
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "aaa"
    Debug.Assert cll(3) = "bbb"
    Debug.Assert cll(4) = "bbb"
    
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_07_String_String_Collection()
    Const PROC = "Test_11_mMsg_Buttons_Service_07_String_String_Collection"
    Dim cll  As New Collection
    
    BoP ErrSrc(PROC)
    cll.Add "ccc"
    cll.Add "ddd"
    
    Set cll = mMsg.Buttons("aaa", "bbb", cll)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection()
    Const PROC = "Test_11_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection"
    Dim cll   As New Collection
    
    BoP ErrSrc(PROC)
    cll.Add "ccc"
    cll.Add "ddd"
    
    Set cll = mMsg.Buttons("aaa;bbb", cll)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_11_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary()
    Const PROC = "Test_11_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary"
    Dim dct As New Dictionary
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    dct.Add "ccc", "ccc"
    dct.Add "ddd", "ddd"
    
    Set cll = mMsg.Buttons("aaa,bbb", dct)
    Debug.Assert cll.Count = 4
    Debug.Assert cll(1) = "aaa"
    Debug.Assert cll(2) = "bbb"
    Debug.Assert cll(3) = "ccc"
    Debug.Assert cll(4) = "ddd"
    
    Set cll = Nothing
    Set dct = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_12_mMsg_ErrMsg_AppErr_5() As Variant
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
    Const PROC = "Test_12_mMsg_ErrMsg_AppErr_5"
    Const EXPECTED_TITLE = "Application Error  5 in: 'mMsgTestServices.Test_12_mMsg_ErrMsg_AppErr_5'"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 12, PROC
    
    mErH.Asserted AppErr(5) ' skips the display of the error message when mErH.Regression = True
    
    Err.Raise Number:=AppErr(5) _
            , source:=ErrSrc(PROC) _
            , Description:="This is a test error description!||This is part of the error description, " & _
                           "concatenated by a double vertical bar and therefore displayed as an additional 'About the error' section " & _
                           "- one of the specific features of the mMsg.ErrMsg service."
        
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_13_mMsg_Dsply_WidthDeterminedByMinimumWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_13_mMsg_Dsply_WidthDeterminedByMinimumWidth"
    
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 13, PROC
            
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Attention:"
            .Label.FontColor = rgbRed
            .Text.Text = "The Ok button ultimately teminates this test without having been evaluated! " & _
                         "The evaluation should include changing arguments like min/max width/height and " & _
                         "- when appropriate - also the label positioning and width."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Description:"
            .Text.Text = wsTest.TestDescription
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Expected test result:"
            .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & mTest.PrcPnt(wsTest.MsgWidthMin, mMsg.enDsplyDimensionWidth) & ") or " _
                       & "to the width determined by the reply buttons."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Please also note:"
            .Text.Text = "1. The message form height is adjusted to the required height up to the specified " & _
                         "maximum heigth which for this test is " & mTest.PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & "." & vbLf & _
                         "2. The minimum width limit for this test is " & mTest.PrcPnt(wsTest.MsgWidthMin, mMsg.enDsplyDimensionWidth) & " and the maximum width limit for this test is " & mTest.PrcPnt(wsTest.MsgWidthMax, mMsg.enDsplyDimensionWidth) & "."
            .Text.FontColor = rgbRed
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_14_mMsg_Dsply_Width_Determined_By_This_eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeextra_long_Title() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_14_mMsg_Dsply_Width_Determined_By_This_eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeextra_long_Title"
    
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 14, PROC
   
   With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Attention:"
            .Label.FontColor = rgbRed
            .Text.Text = "The Ok button ultimately teminates this test without having been evaluated! " & _
                         "The evaluation should include changing arguments like min/max width/height and " & _
                         "- when appropriate - also the label positioning and width."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Description:"
            .Text.Text = wsTest.TestDescription
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Expected test result:"
            .Text.Text = "Because all sections use a proportional font message's width is adjusted exclusively to the title's lenght."
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_15_mMsg_Dsply_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_15_mMsg_Dsply_WidthDeterminedByMonoSpacedMessageSection"
        
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 15, PROC
            
    AssertWidthAndHeight wsTest.MsgWidthMin _
                       , wsTest.MsgWidthMax _
                       , wsTest.MsgHeightMax
    
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Description:"
            .Text.Text = "The length of the longest monospaced message section line determines the width of the message form - " & _
                         "provided it does not exceed the specified maximum form width which for this test is " & mTest.PrcPnt(wsTest.MsgWidthMax, mMsg.enDsplyDimensionWidth) & "% " & _
                         "of the display's width. The maximum form width may be incremented/decremented by " & mTest.PrcPnt(10, mMsg.enDsplyDimensionWidth) & " in order to test the result."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Expected test result:"
            .Text.Text = "Initally, the message form width is adjusted to the longest line in the " & _
                         "monospaced message section and all other message sections are adjusted " & _
                         "to this (enlarged) width." & vbLf & _
                         "When the maximum form width is reduced by " & mTest.PrcPnt(10, mMsg.enDsplyDimensionWidth) & " the monospaced message section is displayed with a horizontal scrollbar."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Please note the following:"
            .Text.Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                         "  the message text is not ""word wrapped""." & vbLf & _
                         "- The message form height is adjusted to the need up to the specified maximum heigth" & vbLf & _
                         "  based on the screen height which for this test is " & mTest.PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & "."
            .Text.MonoSpaced = True
            .Text.FontUnderline = False
        End With
    End With
                        
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_16_mMsg_Dsply_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_16_mMsg_Dsply_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    Dim OneBttnMore         As String
    Dim OneBttnLess         As String
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 16, PROC
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = wsTest.TestDescription
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                     "which is a percentage of the screen height (for this test = " & mTest.PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & "."
    End With
    OneBttnMore = "Repeat with one button more"
    OneBttnLess = "Repeat with one button less"
    vButton6 = "The one more buttonn"
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mMsg.Buttons("Yes", "No", "Cancel", "Ok") _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_17_mMsg_Dsply_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_17_mMsg_Dsply_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 17, PROC
    
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum message width which for this test is " & mTest.PrcPnt(wsTest.MsgWidthMax, mMsg.enDsplyDimensionWidth) & "."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The monospaced message section comes with a horizontal scrollbar."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "This (single line!) monspaced message section exceeds the specified maximum message width which for this test is " & mTest.PrcPnt(wsTest.MsgWidthMax, mMsg.enDsplyDimensionWidth) & "."
        .Text.MonoSpaced = True
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_18_mMsg_Dsply_MonoSpacedMessageSectionExceedsMaxHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_18_mMsg_Dsply_MonoSpacedMessageSectionExceedsMaxHeight"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 18, PROC
       
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The overall message window height exceeds the for this test specified maximum of " & _
                     PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & " of the screen height. Because the monospaced section " & _
                     "is the dominating one regarding its height it is displayed with a horizontal scroll-bar."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "The monospaced message's height is reduced to fit the maximum form height and a vertical scrollbar is added."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = RepeatString(25, "This monospaced message comes with a vertical scrollbar." & vbLf, True)
        .Text.MonoSpaced = True
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_19_mMsg_Dsply_ButtonsOnly() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_19_mMsg_Dsply_ButtonsOnly"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 19, PROC
        
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mTest.BttnsOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_button_default:=BTTN_PASSED _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax
             
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_20_mMsg_Dsply_ButtonsMatrix() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_20_mMsg_Dsply_ButtonsMatrix"
    
    On Error GoTo eh
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim i, j                As Long
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    Dim cll                 As New Collection
    
    BoP ErrSrc(PROC)
    sMsgTitle = "Just to demonstrate what's theoretically possible: Buttons only! Finish with " & BTTN_PASSED & " (default) or " & BTTN_FAILED
    mTest.SetupMsgTitleInstanceAndNo 20, PROC
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 2 To 7 ' rows
        For j = 1 To 7 ' row buttons
            cll.Add "Button" & vbLf & i & "-" & j
        Next j
    Next i
                         
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mMsg.Buttons(mTest.Bttns, cll) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_button_reply_with_index:=False _
             , dsply_button_default:=BTTN_PASSED _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_21_mMsg_Dsply_ButtonScrollBarVertical() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_21_mMsg_Dsply_ButtonScrollBarVertical"
    
    On Error GoTo eh
    Dim i   As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 21, PROC
    
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Description:"
            .Text.Text = "The vertical order of the displayed buttons make the buttons section to the dominating section. " & _
                         "This means that, in case the max message height is exceeded, this sections height is reduced to " & _
                         "fit the max message height and the section is provided with a vertical scroll-bar."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Expected result:"
            .Text.Text = "When the max height is reduced (may already be the case) the buttons section is displayed with " & _
                         "a vertical scrollbar. When the max height is increased, the vertical scroll-bar vanishes."
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mMsg.Buttons("Yes", vbLf, "No", vbLf, "Cancel", vbLf, "Ok") _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_23_mMsg_Dsply_Buttons_Only() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_23_mMsg_Dsply_Buttons_Only"
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim bMonospaced         As Boolean:         bMonospaced = True ' initial test value
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 23, PROC
                         
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_button_reply_with_index:=False _
             , dsply_button_default:=BTTN_PASSED _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
                 
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_26_mMsg_Dsply_ButtonByDictionary()
' ------------------------------------------------------------------------------
' The buttons argument is provided as Dictionary.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_26_mMsg_Dsply_ButtonByDictionary"
    
    On Error GoTo xt
    Dim dct                 As New Collection
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 26, PROC
    
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The ""buttons"" argument is a collection of the test specific buttons " & _
                     "(Passed, Failed) and the two extra Yes, No buttons provided as Dictionary!" & vbLf & vbLf & _
                     "The test proves that the mMsg.Buttons service is able to combine any kind of arguments " & _
                     "provided via the ParamArray."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    dct.Add "Yes"
    dct.Add vbLf
    dct.Add "No"
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_27_mMsg_Box_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_27_mMsg_Box_MessageAsString"
        
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 27, PROC
        
    mMsg.Box Title:=sMsgTitle _
           , Prompt:="This is a message provided as a simple string argument!" _
           , Buttons:=mTest.Bttns _
           , box_buttons_app_run:=mTest.BttnsAppRunArgs _
           , box_width_min:=wsTest.MsgWidthMin _
           , box_width_max:=wsTest.MsgWidthMax _
           , box_height_max:=wsTest.MsgHeightMax _
           , box_modeless:=mTest.MODE_LESS
           
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_30_mMsg_Dsply_ButtonByValue()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_30_mMsg_Dsply_ButtonByValue"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 30, PROC
        
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The ""buttons"" argument is a collection of the test buttons (Passed, Failed) and an additional button provided as value"
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The ""Ok"" button is displayed in the second row."
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_31_mMsg_Dsply_ButtonByString()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_31_mMsg_Dsply_ButtonByString"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 31, PROC
    
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_32_mMsg_Dsply_ButtonByCollection()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_32_mMsg_Dsply_ButtonByCollection"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 32, PROC
    
    With udtMessage.Section(1)
        .Label.Text = "Description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons are centered in n rows"
    End With
      
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_33_mMsg_Dsply_Single_MonoSpaced_Section_Without_Label()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_33_mMsg_Dsply_Single_MonoSpaced_Section_Without_Label"
    Const LINES = 50
    
    On Error GoTo eh
    Dim Msg                 As String
    Dim i                   As Long
    Dim sLbreak             As String
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 33, PROC
            
    i = 1
    For i = 1 To LINES
        Msg = Msg & sLbreak & Format(i, "00: ") & Format(Now(), "YY-MM-DD hh:mm:ss") & " Line " & Format(i, "00") & " of " & Format(LINES, "00") & " the single mono-spaced message section without label."
        sLbreak = vbLf
    Next i
    
    With udtMessage.Section(1).Text
        .Text = Msg
        .MonoSpaced = True
    End With
      
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_34_mMsg_Dsply_Sections_Without_Label_Or_Label_Only()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_34_mMsg_Dsply_Sections_Without_Label_Or_Label_Only"
    
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 34, PROC
    
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Attention:"
            .Label.FontColor = rgbRed
            .Text.Text = "The Ok button ultimately teminates this test without having been evaluated! " & _
                         "The evaluation should include changing arguments like min/max width/height and " & _
                         "- when appropriate - also the label positioning and width."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Multiline label without text (spans full message width), " & _
                          "Multiline label without text (spans full message width), " & _
                          "Multiline label without text (spans full message width)."
            .Label.FontColor = rgbGreen
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Label with text:"
            .Label.FontColor = rgbGreen
            .Text.Text = "Section text"
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Multiline label without text (spans full message width), " & _
                          "Multiline label without text (spans full message width)."
            .Label.FontColor = rgbGreen
        End With
        i = i + 1
        With .Section(i)
            .Text.Text = "Section text without label (spans full message width), " & _
                         "Section text without label (spans full message width), " & _
                         "Section text without label (spans full message width)."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Multiline label without text (spans full message width), " & _
                          "Multiline label without text (spans full message width), " & _
                          "Multiline label without text (spans full message width)."
            .Label.FontColor = rgbGreen
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_40_mMsg_Monitor_Services() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_40_mMsg_Monitor_Services"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim Header              As TypeMsgText
    Dim Step                As TypeMsgText
    Dim Footer              As TypeMsgText
    Dim iLoops              As Long
    Dim lWait               As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 40, PROC
    
    With Header
        .Text = "Step Status (steps 1 to 10)"
        .MonoSpaced = True
        .FontColor = rgbBlue
    End With
    With Footer
        .Text = "Please wait! Process in progress"
        .FontBold = True
        .FontColor = rgbGreen
    End With
    
    iLoops = 15
    lWait = 300
       
    '~~ Because this is the very first service call the size of the monitoring window is initialized
    mMsg.MonitorHeader mon_title:=sMsgTitle, mon_text:=Header, mon_width_max:=50
    mMsg.MonitorFooter sMsgTitle, Footer
    
    For i = 1 To iLoops
        '~~ The Header may be changed at any point in time
        If i = 10 Then
            With Header
                .Text = "Step Status (steps 11 to " & iLoops & ")"
                .MonoSpaced = True
                .FontColor = rgbDarkBlue
            End With
            mMsg.MonitorHeader sMsgTitle, Header
        End If
        
        With Step
            .Text = Format(i, "00") & ". Follow-Up line after " & Format(lWait, "0000") & " Milliseconds."
            .Text = mTest.Repeat(.Text & " ", Int(i / 5) + 1) & vbLf & "    Second line just for test " & mTest.Repeat(".", i)
            .MonoSpaced = True
        End With
        mMsg.Monitor mon_title:=sMsgTitle _
                   , mon_text:=Step
                   
        '~~ Simmulation of a process
        DoEvents
        Sleep lWait
    Next i
    
    With Footer
        .Text = "Process finished! Close this window"
        .FontBold = True
        .FontColor = rgbRed
    End With
    mMsg.MonitorFooter sMsgTitle, Footer
    
    mTest.Evaluate
        
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_50_mMsg_Dsply_LabelPos_Left_R30()
' ------------------------------------------------------------------------------
' Test procedure for label pos left, width 30, various sections with and without
' label and/or text.
' ------------------------------------------------------------------------------
    Const PROC = "Test_50_mMsg_Dsply_LabelPos_Left_R30"
    
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 50, PROC
    
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Label-" & i
            .Text.Text = DefaultSectionTextProp
            .Text.MonoSpaced = False
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Label-" & i
            .Text.Text = DefaultSectionTextProp
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = vbNullString
            .Text.Text = "A section/paragraph without a corresponding label uses the full available message width *)"
        End With
        i = i + 1
        With .Section(i)
            With .Label
                .Text = "*) Link to VBA-Message repo README! (label without text, uses full available message width)"
                .OpenWhenClicked = "https://github.com/warbe-maker/VBA-Message#readme"
            End With
            .Text.Text = vbNullString
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=30 _
             , dsply_width_max:=50 _
             , dsply_height_max:=70 _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_90_mMsg_Dsply_AllInOne() As Variant
    Const PROC      As String = "Test_90_mMsg_Dsply_AllInOne"

    On Error GoTo eh
    Dim i       As Long
    Dim sBttn   As String: sBttn = "Any caption\," & vbLf & "any number" & vbLf & "of lines"
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 90, PROC
    
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Summary of the mMsg.Dsply service features:" & vbLf
            .Label.FontColor = rgbBlue
            .Text.MonoSpaced = True
            .Text.FontSize = 9
            .Text.Text = "- Up to 8 message sections/paragraphs" & vbLf & _
                         "- An optional label allows qualifying each section/paragraph" & vbLf & _
                         "- Labels may be positioned above their corresponding section/paragraph text (the default)" & vbLf & _
                         "  or at the left" & vbLf & _
                         "- Monospaced font option for each section/paragraph (as used with this one)" & vbLf & _
                         "- (Almost) unlimited message size" & vbLf & _
                         "- Unlimited message width and height due to scroll-bars used in case" & vbLf & _
                         "- 7 x 7 reply buttons with any caption text" & vbLf & _
                         "  (including the VBA.MsgBox values (vbOkOnly, vbYesNoCancel, etc.)" & vbLf & _
                         "- Font options: name, size (9 with his section), color, bold, italic, underline" & vbLf & _
                         "- Label only and text only sections/paragraphs" & vbLf & _
                         "- Labels with an ""open when clicked"" option (for url links for example)."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Unlimited message width demo:"
            .Label.FontColor = rgbBlue
            .Text.MonoSpaced = True
            .Text.Text = "Because this section's text is mono-spaced (which by definition is not word-wrapped) the message width is determined by:" & vbLf _
                       & "a) the for this demo specified maximum width of " & mTest.PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & vbLf _
                       & "   (defaults to " & MSG_LIMIT_WIDTH_MAX_PERCENTAGE & "% when not specified)" & vbLf _
                       & "b) the longest line of this section" & vbLf _
                       & "Because the text exeeds the specified maximum message width, a horizontal scroll-bar is displayed." & vbLf _
                       & "Due to this feature there is no message size limit other than the sytem's limit which for a string is about 1GB !!!!"
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Unlimited message height:"
            .Label.FontColor = rgbBlue
            .Text.Text = "As with the message width, the message height is unlimited. When the maximum height (explicitly specified or the default) " _
                       & "is exceeded a vertical scroll-bar is displayed. Due to this feature there is no message size limit other than the sytem's " _
                       & "limit which for a string is about 1GB !!!!"
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Button flexibility:"
            .Label.FontColor = rgbBlue
            .Text.Text = "This demo displays only some of the 7 x 7 = 49 possible reply buttons which may have any caption text " _
                       & "including the classic VBA.MsgBox values (vbOkOnly, vbYesNoCancel, etc.) - even in a mixture."
        End With
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mMsg.Buttons("Yes", "No", "Cancel", "Ok", vbLf, sBttn) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_91_mMsg_Dsply_MinimumMessage() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_91_mMsg_Dsply_MinimumMessage"
    
    On Error GoTo eh
    Dim i As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 91, PROC
        
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Description:"
            .Text.Text = wsTest.TestDescription
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Expected test result:"
            .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & mTest.PrcPnt(wsTest.MsgWidthMin, mMsg.enDsplyDimensionWidth) & ") or " _
                       & "to the width determined by the reply buttons."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "Please also note:"
            .Text.Text = "The message form height is adjusted to the required height up to the specified " & _
                         "maximum heigth which is " & mTest.PrcPnt(wsTest.MsgHeightMax, mMsg.enDsplyDimensionHeight) & " and not exceeded."
            .Text.FontColor = rgbRed
        End With
    End With
                                                                                                  
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=vbOKOnly _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
                         
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_92_mMsg_Dsply_LabelWithUnderlayedURL() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_92_mMsg_Dsply_LabelWithUnderlayedURL"
    
    On Error GoTo eh
    Dim i As Long
    Dim cll As Collection
    Dim sButton As String: sButton = "Caption\," & vbLf & "multilined"
    Set cll = mMsg.Buttons("A", "B", vbOKOnly)
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 92, PROC
        
    With udtMessage
        i = i + 1
        With .Section(i)
            .Label.Text = "Public github repo Common-VBA-Message-Service"
            .Label.OpenWhenClicked = "https://github.com/warbe-maker/Common-VBA-Message-Service"
            .Text.Text = "The corresponding label is underlayed with a url *)."
        End With
        i = i + 1
        With .Section(i)
            .Label.Text = "About this feature of the 'Common-VBA-Message-Service':"
            .Text.Text = "The Common-VBA-Message-Service makes use of the 'Click' and the 'MouseMove' event " & _
                         "of the Label control to allow not only to open a URL but also to open a file or " & _
                         "start an application (open a Workbook, Word document, etc). Examples:"
        End With
        i = i + 1
        With .Section(i).Text
            .MonoSpaced = True
            .FontSize = 9
            .Text = "Open a folder:       C:\TEMP\                 " & vbLf & _
                    "Call the eMail app:  mailto:xxxxx@tutanota.com" & vbLf & _
                    "Open a url/link:     http://......            " & vbLf & _
                    "Open a file:         C:\TEMP\TestThis           (opens a dialog for the selection of the app" & vbLf & _
                    "Open an application: x:\my\workbooks\this.xlsb  (opens Excel)"
        End With
        i = i + 1
        With .Section(i).Label
            .Text = "*) Common-VBA-Message-Service'"
            .OpenWhenClicked = "https://github.com/warbe-maker/Common-VBA-Message-Service"
            .FontSize = 8
        End With
    End With
                                                                                                     
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_buttons:=mMsg.Buttons(cll, "True", "False", vbYesNoCancel, sButton) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=40 _
             , dsply_width_max:=80 _
             , dsply_height_max:=70 _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    mTest.Evaluate
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub UnusedPublicItems()
' ----------------------------------------------------------------
' Please note:
' - Providing the Workbook argument saves the Workbook selection
'   dialog
' - Providing the specification of the excluded VBComponents saves
'   the selection dialog. If explicitly none are to be excluded
'   a vbNullString need to be provided
' - Providing excluded lines - those which are a kind of standard
'   and for sure will not contain any call/use of a public item -
'   may improve the overall performance
' - The service displays the result by means of ShellRun. In case
'   no application is linked with the file extention .txt a dialog
'   to determain which application to use for the open will be
'   displayed.
'
' W. Rauschenberger, Berlin Apr 2023
' ----------------------------------------------------------------
    Const COMPS_EXCLUDED As String = vbNullString ' Example: "mBasic,mDct,mErH,mObject,mTrc"
    Const LINES_EXCLUDED As String = "Select Case*ErrMsg(ErrSrc(PROC))" & vbCrLf & _
                                        "Case vbResume:*Stop:*Resume" & vbCrLf & _
                                        "Case Else:*GoTo xt"
    Const UNUSED_SERVICE As String = "VBPunusedPublic.xlsb!mUnused.Unused" ' must not be altered
    
    
    '~~ Check if the servicing Workbook is open and terminate of not.
    Dim wbk As Workbook
    On Error Resume Next
    Set wbk = Application.Workbooks("VBPunusedPublic.xlsb")
    If Err.Number <> 0 Then
        MsgBox Title:="The Workbook VBPunusedPublic.xlsb is not open!", Prompt:="The Workbook needs to be opened before this procedure is re-executed." & vbLf & vbLf & _
                      "The Workbook may be downloaded from the link provided in the 'Immediate Window'. Use the download button on the displayed webpage."
        Debug.Print "https://github.com/warbe-maker/Common-Excel-VBP-Unused-Public-Items-Service/blob/main/VBPunusedPublic.xlsb?raw=true"
        Exit Sub
    End If
    
    Application.Run UNUSED_SERVICE, , COMPS_EXCLUDED, LINES_EXCLUDED

End Sub

