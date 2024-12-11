Attribute VB_Name = "mMsgTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mMsgTest: Test servicing procedures.
' ======================
' ------------------------------------------------------------------------------
Public Const TEST_URL           As String = "https://github.com/warbe-maker/Common-VBA-Message-Service"
Public Const MODE_LESS          As Boolean = True
Public Const MSG_DIM_INCR_DECR  As Long = 10 ' %
Public Const LBL_WDTH_INCR_DECR As Long = 5  ' pt
Public udtMessage               As udtMsg
Public ufmMsg                   As fMsg
Public TestProcName                As String

Private Const EVALUATION_TITLE  As String = "Modification of the test-procedure's arguments and result evaluation"

Private vNumber                 As Variant
Private sCurrentProcId          As String
Private sCurrentTitle           As String
Private sCurrentDescription     As String
Private sPrevious               As String
Private sUnderEvaluation        As String
Private siEvaluateTop           As Single
Private siEvaluateLeft          As Single

Public Property Get BttnOkOnly() As String:                 BttnOkOnly = "Ok":                                                          End Property
Public Property Get BttnFailed() As String:                 BttnFailed = "Test" & vbLf & "Failed":                                      End Property
Public Property Get BttnTerminate() As String:              BttnTerminate = "Terminate" & vbLf & "(this/subsequent)" & vbLf & "Tests":  End Property
Public Property Get BttnRepeat() As String:                 BttnRepeat = "Repeat" & vbLf & "Test":                                      End Property

Public Property Get BttnPassed() As String:                 BttnPassed = "Test" & vbLf & "Passed":                                      End Property

Public Property Get BttnLblPosLeftAlgnCnter() As String:    BttnLblPosLeftAlgnCnter = "Set Label Pos" & vbLf & "Left aligned center":   End Property

Public Property Get BttnLblPosLeftAlgnLeft() As String:     BttnLblPosLeftAlgnLeft = "Set Label Pos" & vbLf & "Left aligned left":      End Property

Public Property Get BttnLblPosLeftAlgnRight() As String:    BttnLblPosLeftAlgnRight = "Set Label Pos" & vbLf & "Left aligned right":    End Property

Public Property Get BttnLblPosTop() As String:              BttnLblPosTop = "Set Label Pos" & vbLf & "Top2":                            End Property

Public Property Get BttnLblWdthDecr() As String:            BttnLblWdthDecr = "Label Width" & vbLf & "- " & LBL_WDTH_INCR_DECR & " pt": End Property

Public Property Get BttnLblWdthIncr() As String:            BttnLblWdthIncr = "Label Width" & vbLf & "+ " & LBL_WDTH_INCR_DECR & " pt": End Property

Public Property Get BttnMsgHghtMaxDecr() As String:         BttnMsgHghtMaxDecr = "Height Max" & vbLf & "- " & MSG_DIM_INCR_DECR & "%":  End Property

Public Property Get BttnMsgHghtMaxIncr() As String:         BttnMsgHghtMaxIncr = "Height Max" & vbLf & "+ " & MSG_DIM_INCR_DECR & "%":  End Property

Public Property Get BttnMsgWdthMaxDecr() As String:         BttnMsgWdthMaxDecr = "Width Max" & vbLf & "- " & MSG_DIM_INCR_DECR & "%":   End Property

Public Property Get BttnMsgWdthMaxIncr() As String:         BttnMsgWdthMaxIncr = "Width  Max" & vbLf & "+ " & MSG_DIM_INCR_DECR & "%":   End Property

Public Property Get BttnMsgWdthMinDecr() As String:         BttnMsgWdthMinDecr = "Width Min" & vbLf & "- " & MSG_DIM_INCR_DECR & "%":   End Property

Public Property Get BttnMsgWdthMinIncr() As String:         BttnMsgWdthMinIncr = "Width Min" & vbLf & "+ " & MSG_DIM_INCR_DECR & "%":   End Property

Public Property Get CurrentProcId() As String:              CurrentProcId = sCurrentProcId:                                             End Property

Public Property Let CurrentProcId(ByVal s As String):       sCurrentProcId = s:                                                         End Property

Public Property Get CurrentTitle() As String:               CurrentTitle = sCurrentTitle:                                               End Property

Public Property Let CurrentTitle(ByVal s As String):        sCurrentTitle = s:                                                          End Property

Public Property Get CurrentDescription() As String:         CurrentDescription = sCurrentDescription:                                   End Property

Public Property Let CurrentDescription(ByVal s As String):  sCurrentDescription = s:                                                    End Property


Private Property Get LabelPosSpec(Optional ByRef l_pos As enLabelPos, _
                                  Optional ByRef l_lbl_width As Single) As String
    Dim s As String
    s = wsTest.MsgLabelPosSpec
    LabelPosSpec = s
    Select Case True
       Case s = vbNullString:    l_pos = l_pos = enLabelAboveSectionText
       Case InStr(s, "L") <> 0:  l_pos = enLposLeftAlignedLeft:     s = Replace(s, "L", vbNullString)
       Case InStr(s, "C") <> 0:  l_pos = enLposLeftAlignedCenter:   s = Replace(s, "C", vbNullString)
       Case InStr(s, "R") <> 0:  l_pos = enLposLeftAlignedRight:    s = Replace(s, "R", vbNullString)
    End Select
    l_lbl_width = CInt(LabelPosSpec)
    
End Property

Private Property Let LabelPosSpec(Optional ByRef l_pos As enLabelPos, _
                                  Optional ByRef l_lbl_width As Single, _
                                           ByVal l_spec As String)
    wsTest.MsgLabelPosSpec = l_spec
End Property

Public Property Get Previous() As String:                   Previous = sPrevious:                                                       End Property

Public Property Let Previous(ByVal s As String):            sPrevious = s:                                                              End Property

Public Property Get ProcId() As Variant:                    ProcId = vNumber:                                                           End Property

Public Property Let ProcId(ByVal v As Variant):             vNumber = v:                                                                End Property

Public Property Get UnderEvaluation() As String:            UnderEvaluation = sUnderEvaluation:                                         End Property

Public Property Let UnderEvaluation(ByVal s As String):     sUnderEvaluation = s:                                                       End Property

Public Function BttnFailedAction() As Variant
' ------------------------------------------------------------------------------
' This "service" may be called from a "Failed" button when the message has been
' displayed modeless.
' Note: Because the next test procedure displays the message modeless the
'       execution of it returns and the previous test message instance is
'       unloaded.
' ------------------------------------------------------------------------------
    wsTest.Failed
    
    Previous = CurrentProcId
    '~~ Unload current test proc
    Unload mMsg.Instance(mMsgTest.CurrentProcId)
    UnloadEvaluate
    
    '~~ Envoke next test proc
    BttnFailedAction = mMsgTest.TestProc(wsTest.NextTestNumber)
    
End Function

Public Function BttnPassedAction() As Variant
' ------------------------------------------------------------------------------
' This "service" may be called from a "Passed" button when the message has been
' displayed modeless.
' ------------------------------------------------------------------------------
    wsTest.Passed
    mMsgTest.Previous = mMsgTest.CurrentProcId
    
    '~~ Unload current test proc
    Unload mMsg.Instance(mMsgTest.CurrentTitle)
    UnloadEvaluate
    
    '~~ Envoke next test proc
    BttnPassedAction = mMsgTest.TestProc(wsTest.NextTestNumber)
    
End Function

Public Sub BttnReExecWithModArgsAction(Optional ByVal r_msg_width_min As Single = 0, _
                              Optional ByVal r_msg_width_max As Single = 0, _
                              Optional ByVal r_msg_height_max As Single = 0, _
                              Optional ByVal r_lbl_pos As enLabelPos = 0, _
                              Optional ByVal r_lbl_width As Single = 0)
' ------------------------------------------------------------------------------
' Modifies a test procerdure's (r_number) messages argument value and
' re-executes the test procedure identified by its test number. The service may
' modify any number of arguments, no matter whether they are used by the tested
' message variant.
' ------------------------------------------------------------------------------
    Const PROC = "BttnReExecWithModArgsAction"
    
    On Error GoTo eh
    Dim siWidthMin      As Single
    Dim siWidthMax      As Single
    Dim siHeightMax     As Single
    Dim lLabelPos       As enLabelPos
    Dim lLabelWidth     As Long
    Dim sLabelPosSpec   As String
    
    '~~ Get current values
    With wsTest
        siWidthMin = .FormWidthMin
        siWidthMax = .FormWidthMax
        siHeightMax = .FormHeightMax
        sLabelPosSpec = .MsgLabelPosSpec
        lLabelWidth = .LabelWidth
        lLabelPos = .LabelPos
    End With
    
    '~~ Modify the current values
    siWidthMin = Max(siWidthMin + r_msg_width_min, mMsg.MSG_LIMIT_WIDTH_MIN_PERCENTAGE)         ' limit to min width
    siWidthMax = Min(siWidthMax + r_msg_width_max, mMsg.MSG_LIMIT_WIDTH_MAX_PERCENTAGE)         ' limit to max width
    siHeightMax = Min(siHeightMax + r_msg_height_max, mMsg.MSG_LIMIT_HEIGHT_MAX_PERCENTAGE)     ' limit to max height
    
    If r_lbl_pos <> 0 Then
        If lLabelWidth = 0 Then lLabelWidth = 30
        '~~ Label pos modified
        Select Case r_lbl_pos
            Case enLabelAboveSectionText:   sLabelPosSpec = vbNullString
            Case enLposLeftAlignedCenter:   sLabelPosSpec = "C" & lLabelWidth
            Case enLposLeftAlignedLeft:     sLabelPosSpec = "L" & lLabelWidth
            Case enLposLeftAlignedRight:    sLabelPosSpec = "R" & lLabelWidth
        End Select
    ElseIf r_lbl_width <> 0 Then
        '~~ The width is increased or decreased
        lLabelWidth = lLabelWidth + r_lbl_width
        If Abs(lLabelWidth) = LBL_WDTH_INCR_DECR Then lLabelWidth = 30
        Select Case lLabelPos
            Case enLabelAboveSectionText:   sLabelPosSpec = vbNullString
            Case enLposLeftAlignedCenter:   sLabelPosSpec = "C" & lLabelWidth
            Case enLposLeftAlignedLeft:     sLabelPosSpec = "L" & lLabelWidth
            Case enLposLeftAlignedRight:    sLabelPosSpec = "R" & lLabelWidth
        End Select
    End If

    '~~ Return modified values
    With wsTest
        .FormWidthMin = siWidthMin
        .FormWidthMax = siWidthMax
        .FormHeightMax = siHeightMax
        .MsgLabelPosSpec = sLabelPosSpec
    End With

    mMsgTest.TestProc mMsgTest.ProcId
    
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Public Function BttnRepeatAction() As Variant
    '~~ Unload current test proc
    Unload mMsg.Instance(mMsgTest.CurrentTitle)
    mMsgTest.TestProc mMsgTest.ProcId
End Function

Public Sub BttnTerminatedAction()
    
    mMsg.Instance(mMsgTest.Title(mMsgTest.CurrentProcId)).Hide
    If mErH.Regression Then
        EoP mMsgTest.CurrentProcId
        wsTest.RegressionTest = False
        mErH.Regression = False
#If XcTrc_clsTrc = 1 Then
        Trc.Dsply
#ElseIf XcTrc_mTrc = 1 Then
        mTrc.Dsply
#End If
    End If
    Unload mMsg.Instance(mMsgTest.Title(mMsgTest.CurrentProcId))
    UnloadEvaluate
    
End Sub

Public Function Bttns() As Collection
' ------------------------------------------------------------------------------
' Collection of test buttons displayed
' ------------------------------------------------------------------------------
    '~~ Min/Max Width/Height increment/decrement buttons
    Set Bttns = mMsg.Buttons(mMsgTest.BttnPassed, mMsgTest.BttnFailed, mMsgTest.BttnTerminate, vbLf, _
                             mMsgTest.BttnMsgWdthMinIncr, mMsgTest.BttnMsgWdthMinDecr, mMsgTest.BttnMsgWdthMaxIncr, mMsgTest.BttnMsgWdthMaxDecr, vbLf, mMsgTest.BttnMsgHghtMaxIncr, mMsgTest.BttnMsgHghtMaxDecr, vbLf)
    
    If HasLabelWithText Then
        '~~ Label position spec modification buttons
        Select Case wsTest.LabelPos
            Case enLabelAboveSectionText: Set Bttns = mMsg.Buttons(Bttns, mMsgTest.BttnLblPosLeftAlgnCnter, mMsgTest.BttnLblPosLeftAlgnLeft, mMsgTest.BttnLblPosLeftAlgnRight, vbLf)
            Case enLposLeftAlignedCenter: Set Bttns = mMsg.Buttons(Bttns, mMsgTest.BttnLblPosTop, mMsgTest.BttnLblPosLeftAlgnLeft, mMsgTest.BttnLblPosLeftAlgnRight, vbLf, mMsgTest.BttnLblWdthIncr, mMsgTest.BttnLblWdthDecr, vbLf)
            Case enLposLeftAlignedLeft:   Set Bttns = mMsg.Buttons(Bttns, mMsgTest.BttnLblPosTop, mMsgTest.BttnLblPosLeftAlgnCnter, mMsgTest.BttnLblPosLeftAlgnRight, vbLf, mMsgTest.BttnLblWdthIncr, mMsgTest.BttnLblWdthDecr, vbLf)
            Case enLposLeftAlignedRight:  Set Bttns = mMsg.Buttons(Bttns, mMsgTest.BttnLblPosTop, mMsgTest.BttnLblPosLeftAlgnCnter, mMsgTest.BttnLblPosLeftAlgnLeft, vbLf, mMsgTest.BttnLblWdthIncr, mMsgTest.BttnLblWdthDecr, vbLf)
        End Select
    End If
    Bttns.Remove Bttns.Count
    
End Function

Public Function BttnsAppRunArgs() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary with reasonable App.Run arguments for any button
' displayed along wuth a test procedure.
' ------------------------------------------------------------------------------
    Const PROC = "BttnsAppRunArgs"
    
    On Error GoTo eh
    Dim dct As New Dictionary
    
    mMsg.BttnAppRun dct, BttnOkOnly, ThisWorkbook, "mMsgTest.BttnTerminatedAction"
    mMsg.BttnAppRun dct, "Yes", ThisWorkbook, "mMsgTest.BttnTerminatedAction"
    mMsg.BttnAppRun dct, "No", ThisWorkbook, "mMsgTest.BttnTerminatedAction"
    mMsg.BttnAppRun dct, "Cancel", ThisWorkbook, "mMsgTest.BttnTerminatedAction"
    mMsg.BttnAppRun dct, BttnPassed, ThisWorkbook, "mMsgTest.BttnPassedAction"
    mMsg.BttnAppRun dct, BttnFailed, ThisWorkbook, "mMsgTest.BttnFailedAction"
    mMsg.BttnAppRun dct, BttnTerminate, ThisWorkbook, "mMsgTest.BttnTerminatedAction"
    mMsg.BttnAppRun dct, BttnRepeat, ThisWorkbook, "mMsgTest.BttnRepeatAction"
    
    mMsg.BttnAppRun dct, BttnLblPosTop, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, enLabelAboveSectionText
    mMsg.BttnAppRun dct, BttnLblPosLeftAlgnCnter, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, enLposLeftAlignedCenter
    mMsg.BttnAppRun dct, BttnLblPosLeftAlgnLeft, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, enLposLeftAlignedLeft
    mMsg.BttnAppRun dct, BttnLblPosLeftAlgnRight, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, enLposLeftAlignedRight
    
    mMsg.BttnAppRun dct, BttnLblWdthDecr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, 0, -LBL_WDTH_INCR_DECR
    mMsg.BttnAppRun dct, BttnLblWdthIncr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, 0, 0, LBL_WDTH_INCR_DECR
    
    mMsg.BttnAppRun dct, BttnMsgWdthMaxDecr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, -MSG_DIM_INCR_DECR
    mMsg.BttnAppRun dct, BttnMsgWdthMaxIncr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, MSG_DIM_INCR_DECR
    mMsg.BttnAppRun dct, BttnMsgWdthMinDecr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", -MSG_DIM_INCR_DECR
    mMsg.BttnAppRun dct, BttnMsgWdthMinIncr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", MSG_DIM_INCR_DECR
    
    mMsg.BttnAppRun dct, BttnMsgHghtMaxDecr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, -MSG_DIM_INCR_DECR
    mMsg.BttnAppRun dct, BttnMsgHghtMaxIncr, ThisWorkbook, "mMsgTest.BttnReExecWithModArgsAction", 0, 0, MSG_DIM_INCR_DECR
    
    Set BttnsAppRunArgs = dct
    Set dct = Nothing
    
xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

                    

Public Function BttnsBasic() As Collection
' ------------------------------------------------------------------------------
' Collection of test buttons displayed for a basic test.
' ------------------------------------------------------------------------------
    Set BttnsBasic = mMsg.Buttons(mMsgTest.BttnPassed, mMsgTest.BttnFailed, vbLf, mMsgTest.BttnRepeat, mMsgTest.BttnTerminate)
End Function

Public Function BttnsOnly() As Collection
' ------------------------------------------------------------------------------
' Returns a 7 x 7 buttons only matrix.
' ------------------------------------------------------------------------------
    Dim cll As New Collection
    Dim i   As Long
    Dim j   As Long
    
    For i = 1 To 7
        For j = 1 To 7
            cll.Add "B " & i & "-" & j
        Next j
        cll.Add vbLf
    Next i
    Set BttnsOnly = cll
    Set cll = Nothing
    
End Function

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
    If err_source = vbNullString Then err_source = Err.Source
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
    ErrSrc = "mMsgTest." & sProc
End Function

Public Sub Evaluate(Optional ByVal e_selftest As Boolean = False)
' --------------------------------------------------------------------------------
' Displays a modeless dialog to modify the current test-procedure's arguments and
' finally evaluate the result as Passed or Failed.
' --------------------------------------------------------------------------------
    Const PROC = "Evaluate"
    
    On Error GoTo eh
    Dim ufm     As fMsg
    Dim Msg     As mMsg.udtMsg
    Dim i       As Long
    Dim ufmTest As fMsg
    Dim siLeft  As Single
    Dim siTop   As Single
    
    UnloadEvaluate
    If e_selftest Then
        mMsgTest.InitializeTest "02-0", PROC
    End If
    Set ufm = mMsg.Instance(PROC)
    Set ufmTest = mMsg.Instance(CurrentTitle)
    With ufmTest
        siTop = .Top
        siLeft = .Left + .Width + 5
    End With
    
    ufm.VisualizeForTest = False
    With Msg
        With .Section(NextSect(i))
            .Label.Text = "Test Title:"
            .Label.FontColor = rgbBlue
            .Text.Text = sCurrentTitle
        End With
        With .Section(NextSect(i))
            .Label.Text = "Test Description:"
            .Label.FontColor = rgbBlue
            .Text.Text = sCurrentDescription
        End With
        
        With .Section(NextSect(i))
            .Label.Text = "Width Min:"
            .Label.FontColor = rgbBlue
            .Text.Text = wsTest.FormWidthMin & "% of the dispay's width"
        End With
        
        With .Section(NextSect(i))
            .Label.Text = "Width Max:"
            .Label.FontColor = rgbBlue
            .Text.Text = wsTest.FormWidthMax & "% of the display's width"
        End With
        
        With .Section(NextSect(i))
            .Label.Text = "Height Max:"
            .Label.FontColor = rgbBlue
            .Text.Text = wsTest.FormHeightMax & "% of the display's height"
        End With
        
        If HasLabelWithText Then
            With .Section(NextSect(i))
                .Label.Text = "Label Pos Spec:"
                .Label.FontColor = rgbBlue
                .Text.Text = wsTest.MsgLabelPosSpec
            End With
        End If
        
        With .Section(NextSect(i)).Text
            .Text = "Modify any (width/height/Label pos) arguments of the current test proc and finally evaluate the result with Passed or Failed."
        End With
    
        With .Section(NextSect(i))
            .Label.Text = "Attention!"
            .Label.FontColor = rgbRed
            .Label.FontBold = True
            .Text.Text = "Buttons displayed with the test procedure must not be pressed! Since the message is displayed modeless, " & _
                         "in order to allow an extra ""Evaluate"" dialog for the tests and the final Passed/Failed evaluation, " & _
                         "any pressed button may result in an error being displayed because the button might have not been provided " & _
                         "with App.Run arguments. The only exception is the test of the mMsg.ErrMsg service which has a corresponding " & _
                         """About"" paragraph for explanation."
            .Text.FontBold = True
        End With
    End With
    
    mMsg.Dsply dsply_title:=EVALUATION_TITLE _
             , dsply_msg:=Msg _
             , dsply_width_min:=50 _
             , dsply_width_max:=70 _
             , dsply_height_max:=85 _
             , dsply_Label_spec:="R80" _
             , dsply_buttons:=mMsgTest.Bttns _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_modeless:=True _
             , dsply_pos:=siTop & ";" & siLeft
        
xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function HasLabelWithText() As Boolean
' ------------------------------------------------------------------------------
' Retuns TRUE when the current udtMessage has at least one section with a Label
' and a text.
' ------------------------------------------------------------------------------
    Dim i As Long
    
    With mMsgTest.udtMessage
        For i = 1 To mMsg.NoOfMsgSects
            With .Section(i)
                If .Label.Text <> vbNullString And .Text.Text <> vbNullString Then
                    HasLabelWithText = True
                    Exit For
                End If
            End With
        Next i
    End With
    
End Function

Public Sub InitializeMessage(ByRef m_form As fMsg, _
                             ByVal m_title As String)
' ------------------------------------------------------------------------------
' Initializes all message sections with defaults used with all tests provided
' they use the global declared udtMessage.
' ------------------------------------------------------------------------------
    Dim i As Long
    
    mMsg.Instance i_title:=m_title, i_unload:=True                    ' Ensures a message starts from scratch
    Set m_form = mMsg.Instance(m_title)
    
    For i = 1 To mMsg.NoOfMsgSects ' obtained when the designed controls are collected
        With mMsgTest.udtMessage.Section(i)
            With .Label
                .Text = vbNullString
                .FontColor = rgbBlue
                .OnClickAction = mMsgTest.TEST_URL
            End With
            With .Text
                .Text = vbNullString
                .FontBold = False
                .FontColor = rgbBlack
                .OnClickAction = mMsgTest.TEST_URL
            End With
        End With
    Next i

End Sub

Public Sub InitializeTest(ByVal s_number As Variant, _
                          ByVal s_proc As String)
' ------------------------------------------------------------------------------
' Sets up a message form instance based on the test proc's title.
' ------------------------------------------------------------------------------
    
    mMsgTest.ProcId = s_number
    TestProcName = mMsgTest.Title(s_proc)
    mMsgTest.CurrentProcId = s_proc
    mMsgTest.CurrentTitle = TestProcName
    mMsgTest.CurrentDescription = wsTest.TestDescription
    mMsgTest.InitializeMessage m_form:=ufmMsg, m_title:=TestProcName ' set test-global message specifications
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    wsTest.SelectDescription
    
End Sub

Private Function IsUcase(ByVal s As String) As Boolean

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

Public Function NextSect(ByRef n_sect As Long) As Long
    n_sect = n_sect + 1
    NextSect = n_sect
End Function

Public Function PrcPnt(ByVal pp_value As Single, _
                       ByVal pp_dimension As String) As String
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    PrcPnt = mMsg.ValueAsPercentage(pp_value, pp_dimension)
    Select Case pp_dimension
        Case mMsg.enDsplyDimensionWidth:   PrcPnt = PrcPnt & "% of the display's width (" & mMsg.ValueAsPt(pp_value, mMsg.enDsplyDimensionWidth) & "pt)"
        Case mMsg.enDsplyDimensionHeight:   PrcPnt = PrcPnt & "% of the display's height (" & mMsg.ValueAsPt(pp_value, mMsg.enDsplyDimensionWidth) & "pt)"
    End Select
    
End Function

Public Function Repeat(repeat_string As String, repeat_n_times As Long)
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

Public Function RepeatString(ByVal rep_n_times As Long, _
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

Public Function TestProc(ByVal n_test_number As Variant) As Variant
        
    Select Case n_test_number
        Case "02-1":    TestProc = mMsgTestProcs.Test_02_1_Single_Section_PropSpaced
        Case "02-2":    TestProc = mMsgTestProcs.Test_02_2_Single_Section_MonoSpaced_With_Label
        Case "02-3":    TestProc = mMsgTestProcs.Test_02_3_Single_Section_MonoSpaced_No_Label
        Case "02-4":    TestProc = mMsgTestProcs.Test_02_4_Single_Section_MonoSpaced_With_VH_Scroll
        Case "02-5":    TestProc = mMsgTestProcs.Test_02_5_Single_Section_Label_Only
        Case "00":      TestProc = mMsgTestServices.Test_00_Evaluate
        Case 11:        TestProc = mMsgTestServices.Test_11_mMsg_Box_Buttons_Only
        Case 12:        TestProc = mMsgTestServices.Test_12_mMsg_ErrMsg_AppErr_5
        Case 13:        TestProc = mMsgTestServices.Test_13_mMsg_Dsply_WidthDeterminedByMinimumWidth
        Case 14:        TestProc = mMsgTestServices.Test_14_mMsg_Dsply_Width_Determined_By_This_eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeextra_long_Title
        Case 15:        TestProc = mMsgTestServices.Test_15_mMsg_Dsply_WidthDeterminedByMonoSpacedMessageSection
        Case 16:        TestProc = mMsgTestServices.Test_16_mMsg_Dsply_WidthDeterminedByReplyButtons
        Case 17:        TestProc = mMsgTestServices.Test_17_mMsg_Dsply_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case 18:        TestProc = mMsgTestServices.Test_18_mMsg_Dsply_MonoSpacedMessageSectionExceedsMaxHeight
        Case 19:        TestProc = mMsgTestServices.Test_19_mMsg_Dsply_ButtonsOnly
        Case 20:        TestProc = mMsgTestServices.Test_20_mMsg_Dsply_ButtonsMatrix
        Case 21:        TestProc = mMsgTestServices.Test_21_mMsg_Dsply_ButtonScrollBarVertical
        Case 23:        TestProc = mMsgTestServices.Test_23_mMsg_Dsply_Buttons_Only
        Case 26:        TestProc = mMsgTestServices.Test_26_mMsg_Dsply_ButtonByDictionary
        Case 27:        TestProc = mMsgTestServices.Test_27_mMsg_Box_MessageAsString
        Case 30:        TestProc = mMsgTestServices.Test_30_mMsg_Dsply_ButtonByValue
        Case 31:        TestProc = mMsgTestServices.Test_31_mMsg_Dsply_ButtonByString
        Case 32:        TestProc = mMsgTestServices.Test_32_mMsg_Dsply_ButtonByCollection
        Case 33:        TestProc = mMsgTestServices.Test_33_mMsg_Dsply_Single_MonoSpaced_Section_Without_Label
        Case 34:        TestProc = mMsgTestServices.Test_34_mMsg_Dsply_Sections_Without_Label_Or_Label_Only
        Case 40:        TestProc = mMsgTestServices.Test_40_mMsg_Monitor_Services
        Case 50:        TestProc = mMsgTestServices.Test_50_mMsg_Dsply_LabelPos_Left_R30
        Case 90:        TestProc = mMsgTestServices.Test_90_mMsg_Dsply_AllInOne
        Case 91:        TestProc = mMsgTestServices.Test_91_mMsg_Dsply_MinimumMessage
        Case 0
            mMsg.Instance(Title(mMsgTest.CurrentProcId)).Hide
            If mErH.Regression Then
                EoP "mMsgTestServices.Test_10_Regression"
                wsTest.RegressionTest = False
                mErH.Regression = False
            End If
#If XcTrc_clsTrc = 1 Then
            Trc.Dsply
#ElseIf XcTrc_mTrc = 1 Then
            mTrc.Dsply
#End If
            Unload mMsg.Instance(Title(CurrentProcId))
    End Select
    
End Function

Public Function Title(ByVal t_s As String) As String
' ------------------------------------------------------------------------------
' Convert a string (t_s) into a readable message title by:
' - replacing all underscores with a whitespace
' - characters immediately following an underscore to a lowercase letter.
' ------------------------------------------------------------------------------
    Dim i As Long
    Dim s As String
    
    t_s = Replace(Replace(t_s, "_", " "), "  ", " ")
    For i = 1 To Len(t_s)
        If IsUcase(Mid(t_s, i, 1)) Then
            s = s & " " & Mid(t_s, i, 1)
        Else
            s = s & Mid(t_s, i, 1)
        End If
    Next i
    s = Replace(s, "m Msg ", "(mMsg.", 1, 1)
    s = Replace(s, " m Msg ", ", mMsg.")
    s = Right(s, Len(s) - 1)
    s = Replace(s, "mMsg. Box", "mMsg.Box Service):")
    s = Replace(s, "mMsg. Dsply", "mMsg.Dsply Service):")
    s = Replace(s, "mMsg. Err Msg", "mMsg.ErrMsg Service):")
    
    Title = s
    
End Function

Private Sub UnloadEvaluate()
' --------------------------------------------------------------------------------
'
' --------------------------------------------------------------------------------
    Const PROC = "UnloadEvaluate"
    
    On Error GoTo eh
    Dim ufm As fMsg
    
    If mMsg.Instances.Exists(EVALUATION_TITLE) Then
        On Error Resume Next
        Set ufm = mMsg.Instances(EVALUATION_TITLE)
        If Err.Number <> 0 Then
            mMsg.Instances.Remove EVALUATION_TITLE
        End If
        With ufm
            siEvaluateTop = .Top
            siEvaluateLeft = .Left
        End With
        Unload ufm
    Else
        siEvaluateTop = 20
        siEvaluateLeft = 200
    End If

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

