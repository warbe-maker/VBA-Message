Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mTest: Test servicing procedures.
' ======================
' ------------------------------------------------------------------------------
Public Const BTTN_PASSED    As String = "Test" & vbLf & "Passed"
Public Const BTTN_FAILED    As String = "Test" & vbLf & "Failed"
Public Const BTTN_TRMNTE    As String = "Terminate" & vbLf & "(this and following)" & vbLf & "Tests"
Public Const MODE_LESS      As Boolean = True
Public cllBttnsMsg          As Collection
Public cllBttnsTest         As Collection
Public udtMessage           As TypeMsg
Public ufmMsg               As fMsg
Public sMsgTitle            As String

Private lNumber             As Long
Private sCurrent            As String
Private sPrevious           As String

Public Property Get Current() As String:            Current = sCurrent:     End Property

Public Property Let Current(ByVal s As String):     sCurrent = s:           End Property

Public Property Get Previous() As String:           Previous = sPrevious:   End Property

Public Property Let Previous(ByVal s As String):    sPrevious = s:          End Property

Private Property Get LabelPosSpec(Optional ByRef l_pos As enLabelPos, _
                                  Optional ByRef l_lbl_width As Single, _
                                  Optional ByVal l_test_no As Long) As String
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
                                  Optional ByVal l_test_no As Long, _
                                           ByVal l_spec As String)
    wsTest.MsgLabelPosSpec = l_spec
End Property

Public Sub ModifyTestProcArgument(ByVal m_number As Long, _
                         Optional ByVal m_width_min As Single = 0, _
                         Optional ByVal m_width_max As Single = 0, _
                         Optional ByVal m_height_min As Single = 0, _
                         Optional ByVal m_height_max As Single = 0, _
                         Optional ByVal m_label_pos As String = vbNullString, _
                         Optional ByVal m_label_width As Single = 0)
' ------------------------------------------------------------------------------
' Modifies a test procerdure's (m_number) messages argument value and
' re-executes the test procedure identified by its test number. The service may
' modify any number of arguments, no matter whether they are used by the tested
' message variant.
' ------------------------------------------------------------------------------
    Select Case True
        Case m_width_min <> 0:
        Case m_width_max <> 0:
        Case m_height_min <> 0:
        Case m_height_max <> 0:
        Case m_label_pos <> vbNullString
            If m_label_width = 0 Then m_label_width = 30
            wsTest.MsgLabelPosSpec = m_label_pos & m_label_width
    End Select
End Sub
                        
Public Function BttnsAppRunArgs() As Dictionary
    Dim dct As New Dictionary
    mMsg.BttnAppRun dct, BTTN_PASSED, ThisWorkbook, "mTest.Passed"
    mMsg.BttnAppRun dct, BTTN_FAILED, ThisWorkbook, "mTest.Failed"
    mMsg.BttnAppRun dct, BTTN_TRMNTE, ThisWorkbook, "mTest.Terminated"
    Set BttnsAppRunArgs = dct
    
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

Public Sub Evaluate(ByVal e_proc As String)
' --------------------------------------------------------------------------------
' Obtain the test result for a named (e_proc) test procedure - which is unable to
' do it by its own since the corresponding buttons are (cannot be) displayed.
' The usage of the evaluated test procedure's name displays a title equal to the
' title of the evaluated procedure - and thus this procedure's message window
' needs to be closed beforehand. Since the Passed and the Failed service continue
' with the next subsequently planned test procedure
' --------------------------------------------------------------------------------
    
    Dim cll As Collection
    Dim dct As New Dictionary
    Dim ufm As fMsg
    
    mMsg.MsgInstance mTest.Title(e_proc), True ' unload
    Set ufm = mMsg.MsgInstance(mTest.Title(e_proc))
    ufm.VisualizeForTest = False
    
    Set cll = mMsg.Buttons(BTTN_PASSED, BTTN_FAILED)
    '~~ AppRun buttons! Becomes effective only when the form is displayed modeless!
    mMsg.BttnAppRun dct, BTTN_PASSED, ThisWorkbook, "mTest.Passed"
    mMsg.BttnAppRun dct, BTTN_FAILED, ThisWorkbook, "mTest.Failed"
    
    Select Case mMsg.Box(Title:=mTest.Title(e_proc) _
                       , Prompt:=vbNullString _
                       , Buttons:=cll _
                       , box_buttons_app_run:=dct _
                       , box_modeless:=True _
                        )
        Case BTTN_PASSED:       mTest.Passed
        Case BTTN_FAILED:       mTest.Failed
    End Select
    
    Set cll = Nothing
    Set dct = Nothing
    
End Sub

Public Function Failed() As Variant
' ------------------------------------------------------------------------------
' This "service" may be called from a "Failed" button when the message has been
' displayed modeless.
' Note: Because the next test procedure displays the message modeless the
'       execution of it returns and the previous test message instance is
'       unloaded.
' ------------------------------------------------------------------------------
    wsTest.Failed
    
    Previous = Current
    mMsg.MsgInstance(Title(Previous)).Hide
    Failed = mTest.TestProc(wsTest.NextTestNumber)
    Unload mMsg.MsgInstance(Title(Previous))
    
End Function

Public Property Get Number() As Long:           Number = lNumber:   End Property

Public Property Let Number(ByVal l As Long):    lNumber = l:        End Property

Public Sub SetupMsgTitleInstanceAndNo(ByVal s_number As Long, _
                                      ByVal s_proc As String)
    
    mTest.Number = s_number
    Set cllBttnsTest = mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE)
        
    Set cllBttnsMsg = New Collection
    sMsgTitle = mTest.Title(s_proc)
    mTest.Current = s_proc
    mTest.MessageInit m_form:=ufmMsg, m_title:=sMsgTitle ' set test-global message specifications
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
End Sub

Public Function TestProc(ByVal n_test_number As Long) As Variant
        
    Select Case n_test_number
        Case 1:     TestProc = Test_01_mMsg_Box_Service_Buttons_7_By_7_Matrix
        Case 3:     TestProc = Test_03_mMsg_Dsply_Service_WidthDeterminedByMinimumWidth
        Case 4:     TestProc = Test_04_mMsg_Dsply_Service_WidthDeterminedByTitle
        Case 5:     TestProc = Test_05_mMsg_Dsply_Service_WidthDeterminedByMonoSpacedMessageSection
        Case 6:     TestProc = Test_06_mMsg_Dsply_Service_WidthDeterminedByReplyButtons
        Case 7:     TestProc = Test_07_mMsg_Dsply_Service_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case 8:     TestProc = Test_08_mMsg_Dsply_Service_MonoSpacedMessageSectionExceedsMaxHeight
        Case 9:     TestProc = Test_09_mMsg_Dsply_Service_ButtonsOnly
        Case 10:    TestProc = Test_10_mMsg_Dsply_Service_ButtonsMatrix
        Case 11:    TestProc = Test_11_mMsg_Dsply_Service_ButtonScrollBarVertical
        Case 12:    TestProc = Test_12_mMsg_Dsply_Service_ButtonScrollBarHorizontal
        Case 13:    TestProc = Test_13_mMsg_Dsply_Service_ButtonsMatrix_With_Both_Scroll_Bars
        Case 16:    TestProc = Test_16_mMsg_Dsply_Service_ButtonByDictionary
        Case 17:    TestProc = Test_17_mMsg_Box_Service_MessageAsString
        Case 20:    TestProc = Test_20_mMsg_Dsply_Service_ButtonByValue
        Case 23:    TestProc = Test_23_mMsg_Dsply_Service_MonoSpacedSectionOnly
        Case 24:    TestProc = Test_24_mMsg_All_Sections_Label_Pos_Left
        Case 90:    TestProc = Test_90_mMsg_Dsply_Service_AllInOne
        Case 91:    TestProc = Test_91_mMsg_Dsply_Service_MinimumMessage
        Case 92:    TestProc = Test_92_mMsg_Dsply_Service_LabelWithUnderlayedURL
        Case 0:     mMsg.MsgInstance(Title(mTest.Current)).Hide
                    If mErH.Regression Then
                        EoP "mMsgTestServices.Test_00_Regression"
                        wsTest.RegressionTest = False
                        mErH.Regression = False
                    End If
#If XcTrc_clsTrc = 1 Then
                    Trc.Dsply
#ElseIf XcTrc_mTrc = 1 Then
                    mTrc.Dsply
#End If
                    Unload mMsg.MsgInstance(Title(Current))
    End Select

End Function

Public Function Passed() As Variant
' ------------------------------------------------------------------------------
' This "service" may be called from a "Passed" button when the message has been
' displayed modeless.
' ------------------------------------------------------------------------------
    wsTest.Passed
    mTest.Previous = mTest.Current
    mMsg.MsgInstance(mTest.Title(mTest.Previous)).Hide
    Passed = mTest.TestProc(wsTest.NextTestNumber)
    Unload mMsg.MsgInstance(mTest.Title(mTest.Previous))
    
End Function

Public Sub Terminated()
    
    mMsg.MsgInstance(mTest.Title(mTest.Current)).Hide
    If mErH.Regression Then
        EoP mTest.Current
        wsTest.RegressionTest = False
        mErH.Regression = False
    End If
#If XcTrc_clsTrc = 1 Then
        Trc.Dsply
#ElseIf XcTrc_mTrc = 1 Then
        mTrc.Dsply
#End If
    Unload mMsg.MsgInstance(mTest.Title(mTest.Current))
End Sub

Public Sub MessageInit(ByRef m_form As fMsg, _
                       ByVal m_title As String)
' ------------------------------------------------------------------------------
' Initializes the all message sections with the defaults throughout this test
' module which uses a module global declared udtMessage for a consistent layout.
' ------------------------------------------------------------------------------
    Dim i As Long
    
    mMsg.MsgInstance fi_key:=m_title, fi_unload:=True                    ' Ensures a message starts from scratch
    Set m_form = mMsg.MsgInstance(m_title)
    
    For i = 1 To mMsg.NoOfMsgSects ' obtained when the designed controls are collected
        With udtMessage.Section(i)
            .Label.Text = vbNullString
            .Label.FontColor = rgbBlue
            .Text.Text = vbNullString
            .Text.MonoSpaced = False
            .Text.FontItalic = False
            .Text.FontUnderline = False
            .Text.FontColor = rgbBlack
        End With
    Next i

End Sub

Public Function Title(ByVal s As String) As String
' ------------------------------------------------------------------------------
' Convert a string (s) into a readable message title by:
' - replacing all underscores with a whitespace
' - characters immediately following an underscore to a lowercase letter.
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
    sResult = Replace(sResult, "m Msg ", "(mMsg.", 1, 1)
    sResult = Replace(sResult, " m Msg ", ", mMsg.")
    sResult = Right(sResult, Len(sResult) - 1)
    Title = Replace(sResult, " Service", " Service)")
    
End Function

