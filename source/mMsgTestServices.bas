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
' Note: - All test procedures (except "Test_02_mMsg_ErrMsg_Service") display the
' -----   message modeless - regardless the option set - with a "Passed",
'         "Failed", and a "Terminate" button waiting for either of the three is
'         pressed.
'       - For the Regression test (Test_00_Regression) explicitly raised errors
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
Private lTestMsgHeightIncrDecr  As Long
Private lTestMsgHeightMax       As Long
Private lTestMsgHeightMin       As Long
Private lTestMsgWidthIncrDecr   As Long
Private lTestMsgWidthMax        As Long
Private lTestMsgWidthMin        As Long
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

Public Sub cmdTest01_Click():   mMsgTestServices.Test_01_mMsg_Buttons_Service:                                         End Sub

Public Sub cmdTest02_Click():   mMsgTestServices.Test_02_mMsg_ErrMsg_Service:                                          End Sub

Public Sub cmdTest03_Click():   mMsgTestServices.Test_03_mMsg_Dsply_Service_WidthDeterminedByMinimumWidth:             End Sub

Public Sub cmdTest04_Click():   mMsgTestServices.Test_04_mMsg_Dsply_Service_WidthDeterminedByTitle:                    End Sub

Public Sub cmdTest05_Click():   mMsgTestServices.Test_05_mMsg_Dsply_Service_WidthDeterminedByMonoSpacedMessageSection: End Sub

Public Sub cmdTest06_Click():   mMsgTestServices.Test_06_mMsg_Dsply_Service_WidthDeterminedByReplyButtons:             End Sub

Public Sub cmdTest07_Click():   mMsgTestServices.Test_07_mMsg_Dsply_Service_MonoSpacedSectionWidthExceedsMaxMsgWidth:  End Sub

Public Sub cmdTest08_Click():   mMsgTestServices.Test_08_mMsg_Dsply_Service_MonoSpacedMessageSectionExceedsMaxHeight:  End Sub

Public Sub cmdTest09_Click():   mMsgTestServices.Test_09_mMsg_Dsply_Service_ButtonsOnly:                               End Sub

Public Sub cmdTest10_Click():   mMsgTestServices.Test_10_mMsg_Dsply_Service_ButtonsMatrix:                             End Sub

Public Sub cmdTest11_Click():   mMsgTestServices.Test_11_mMsg_Dsply_Service_ButtonScrollBarVertical:                   End Sub

Public Sub cmdTest12_Click():   mMsgTestServices.Test_12_mMsg_Dsply_Service_ButtonScrollBarHorizontal:                 End Sub

Public Sub cmdTest13_Click():   mMsgTestServices.Test_13_mMsg_Dsply_Service_ButtonsMatrix_With_Both_Scroll_Bars:       End Sub

Public Sub cmdTest16_Click():   mMsgTestServices.Test_16_mMsg_Dsply_Service_ButtonByDictionary:                        End Sub

Public Sub cmdTest17_Click():   mMsgTestServices.Test_17_mMsg_Box_Service_MessageAsString:                             End Sub

Public Sub cmdTest20_Click():   mMsgTestServices.Test_20_mMsg_Dsply_Service_ButtonByValue:                             End Sub

Public Sub cmdTest23_Click():   mMsgTestServices.Test_23_mMsg_Dsply_Service_MonoSpacedSectionOnly:                     End Sub

Public Sub cmdTest24_Click():   mMsgTestServices.Test_24_mMsg_All_Sections_Label_Pos_Left:                           End Sub

Public Sub cmdTest30_Click():   mMsgTestServices.Test_30_mMsg_MonitorHeader_mMsg_Monitor_mMsg_MonitorFooter_Service:   End Sub

Public Sub cmdTest90_Click():   mMsgTestServices.Test_90_mMsg_Dsply_Service_AllInOne:                                  End Sub

Public Sub cmdTest91_Click():   mMsgTestServices.Test_91_mMsg_Dsply_Service_MinimumMessage:                            End Sub

Public Sub cmdTest92_Click():   mMsgTestServices.Test_92_mMsg_Dsply_Service_LabelWithUnderlayedURL:                    End Sub

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

Private Sub Explore(ByVal ctl As Variant, _
          Optional ByVal applied As Boolean = True)
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Explore"
    
    On Error GoTo eh
    Dim dct         As New Dictionary
    Dim v           As Variant
    Dim Appl        As String   ' ControlApplied
    Dim l           As String   ' .Left
    Dim w           As String   ' .Width
    Dim t           As String   ' .Top
    Dim h           As String   ' .Height
    Dim SW          As String   ' .ScrollWidth
    Dim SH          As String   ' .ScrollHeight
    Dim FW          As String   ' ufmMsg.InsideWidth
    Dim CW          As String   ' Content width
    Dim CH          As String   ' Content height
    Dim FH          As String   ' ufmMsg.InsideHeight
    Dim i           As Long
    Dim Item        As String
    Dim j           As String
    Dim frm         As MSForms.Frame
    
    sMsgTitle = "Explore"
    Unload mMsg.MsgInstance(sMsgTitle) ' Ensure there is no process monitoring with this title still displayed
    Set ufmMsg = mMsg.MsgInstance(sMsgTitle)
    
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
        w = Align(Format(ctl.Width, "000.0"), 7, AlignCentered, " ")
        t = Align(Format(ctl.Top, "000.0"), 7, AlignCentered, " ")
        h = Align(Format(ctl.Height, "000.0"), 7, AlignCentered, " ")
        FH = Align(Format(ufmMsg.InsideHeight, "000.0"), 7, AlignCentered, " ")
        FW = Align(Format(ufmMsg.InsideWidth, "000.0"), 7, AlignCentered, " ")
        If TypeName(ctl) = "Frame" Then
            Set frm = ctl
            CW = Align(Format(ufmMsg.ContentWidth(frm), "000.0"), 7, AlignCentered, " ")
            CH = Align(Format(ufmMsg.ContentHeight(frm), "000.0"), 7, AlignCentered, " ")
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
        
        Debug.Print Align(ctl.Name, 20, AlignLeft) & "|" & Appl & "|" & l & "|" & w & "|" & CW & "|" & t & "|" & h & "|" & CH & "|" & SH & "|" & SW & "|" & FW & "|" & FH
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

Private Function PrcPnt(ByVal pp_value As Single, _
                        ByVal pp_dimension As String) As String
    PrcPnt = mMsg.ValueAsPercentage(pp_value, pp_dimension) & "% (" & mMsg.ValueAsPt(pp_value, "w") & "pt)"
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

Public Function Test_00_Regression() As Variant
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
    mErH.Regression = True
    mTrc.FileName = "RegressionTest.ExecTrace.log"
    mTrc.Title = "Regression test module mMsg"
    mTrc.NewFile
    
    BoP ErrSrc(PROC)
    mTest.Current = vbNullString
    Test_00_Regression = mTest.TestProc(wsTest.NextTestNumber)

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_01_mMsg_Box_Service_Buttons_7_By_7_Matrix() As Variant
' ------------------------------------------------------------------------------
' The Buttons service "in action": Display a matrix of 7 x 7 buttons
' ------------------------------------------------------------------------------
    Const PROC = "Test_01_mMsg_Box_Service_Buttons_7_By_7_Matrix"
    
    Dim i                   As Long

    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 1, PROC
    
    For i = 1 To 49
        cllBttnsMsg.Add "B" & Format(i, "00")
    Next i
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, cllBttnsMsg) ' excessive buttons are ignored !
    Debug.Assert cllBttnsMsg.Count = 55
    Debug.Assert cllBttnsMsg(8) = vbLf
    Debug.Assert cllBttnsMsg(16) = vbLf
    Debug.Assert cllBttnsMsg(24) = vbLf
    Debug.Assert cllBttnsMsg(32) = vbLf
    Debug.Assert cllBttnsMsg(40) = vbLf
    Debug.Assert cllBttnsMsg(48) = vbLf
    
    mMsg.Box Prompt:=vbNullString _
           , Buttons:=cllBttnsMsg _
           , box_buttons_app_run:=mTest.BttnsAppRunArgs _
           , Title:=sMsgTitle _
           , box_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    
End Function

Private Sub Test_01_mMsg_Buttons_Service()
    Const PROC = "Test_01_mMsg_Buttons_Service"
    BoP ErrSrc(PROC)
    mMsgTestServices.Test_01_mMsg_Buttons_Service_01_Empty
    mMsgTestServices.Test_01_mMsg_Buttons_Service_02_Single_String
    mMsgTestServices.Test_01_mMsg_Buttons_Service_03_Single_Numeric_Item
    mMsgTestServices.Test_01_mMsg_Buttons_Service_04_String_String
    mMsgTestServices.Test_01_mMsg_Buttons_Service_05_Collection_String_String
    mMsgTestServices.Test_01_mMsg_Buttons_Service_06_String_Collection_String
    mMsgTestServices.Test_01_mMsg_Buttons_Service_07_String_String_Collection
    mMsgTestServices.Test_01_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection
    mMsgTestServices.Test_01_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary
    mMsgTestServices.Test_01_mMsg_Box_Service_Buttons_7_By_7_Matrix
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_01_Empty()
    Const PROC = "Test_01_mMsg_Buttons_Service_01_Empty"
    
    BoP ErrSrc(PROC)
    Set cllBttnsMsg = mMsg.Buttons()
    Debug.Assert cllBttnsMsg.Count = 0
    Set cllBttnsMsg = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_02_Single_String()
    Const PROC = "Test_01_mMsg_Buttons_Service_02_Single_String"
    
    BoP ErrSrc(PROC)
    Set cllBttnsMsg = mMsg.Buttons("aaa")
    Debug.Assert cllBttnsMsg.Count = 1
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Set cllBttnsMsg = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_01_mMsg_Buttons_Service_03_Single_Numeric_Item() As Variant
    Const PROC = "Test_01_mMsg_Buttons_Service_03_Single_Numeric_Item"
    
    BoP ErrSrc(PROC)
    Set cllBttnsMsg = mMsg.Buttons(vbResumeOk)
    Debug.Assert cllBttnsMsg.Count = 1
    Debug.Assert cllBttnsMsg(1) = vbResumeOk
    Set cllBttnsMsg = Nothing
    EoP ErrSrc(PROC)
End Function

Private Sub Test_01_mMsg_Buttons_Service_04_String_String()
    Const PROC = "Test_01_mMsg_Buttons_Service_04_String_String"
    
    BoP ErrSrc(PROC)
    Set cllBttnsMsg = mMsg.Buttons("aaa", "bbb")
    Debug.Assert cllBttnsMsg.Count = 2
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "bbb"
    Set cllBttnsMsg = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_05_Collection_String_String()
    Const PROC = "Test_01_mMsg_Buttons_Service_05_Collection_String_String"
    Dim cllBttnsMsg_1 As New Collection
    
    BoP ErrSrc(PROC)
    cllBttnsMsg_1.Add "aaa"
    cllBttnsMsg_1.Add "bbb"
    
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsMsg_1, "aaa", "bbb")
    Debug.Assert cllBttnsMsg.Count = 4
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "bbb"
    Debug.Assert cllBttnsMsg(3) = "aaa"
    Debug.Assert cllBttnsMsg(4) = "bbb"
    
    Set cllBttnsMsg = Nothing
    Set cllBttnsMsg_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_06_String_Collection_String()
    Const PROC = "Test_01_mMsg_Buttons_Service_06_String_Collection_String"
    Dim cllBttnsMsg_1   As New Collection
    
    BoP ErrSrc(PROC)
    cllBttnsMsg_1.Add "aaa"
    cllBttnsMsg_1.Add "bbb"
    
    Set cllBttnsMsg = mMsg.Buttons("aaa", cllBttnsMsg_1, "bbb")
    Debug.Assert cllBttnsMsg.Count = 4
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "aaa"
    Debug.Assert cllBttnsMsg(3) = "bbb"
    Debug.Assert cllBttnsMsg(4) = "bbb"
    
    Set cllBttnsMsg = Nothing
    Set cllBttnsMsg_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_07_String_String_Collection()
    Const PROC = "Test_01_mMsg_Buttons_Service_07_String_String_Collection"
    Dim cllBttnsMsg_1   As New Collection
    
    BoP ErrSrc(PROC)
    cllBttnsMsg_1.Add "ccc"
    cllBttnsMsg_1.Add "ddd"
    
    Set cllBttnsMsg = mMsg.Buttons("aaa", "bbb", cllBttnsMsg_1)
    Debug.Assert cllBttnsMsg.Count = 4
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "bbb"
    Debug.Assert cllBttnsMsg(3) = "ccc"
    Debug.Assert cllBttnsMsg(4) = "ddd"
    
    Set cllBttnsMsg = Nothing
    Set cllBttnsMsg_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection()
    Const PROC = "Test_01_mMsg_Buttons_Service_08_Semicolon_Delimited_String_Collection"
    Dim cllBttnsMsg_1   As New Collection
    
    BoP ErrSrc(PROC)
    cllBttnsMsg_1.Add "ccc"
    cllBttnsMsg_1.Add "ddd"
    
    Set cllBttnsMsg = mMsg.Buttons("aaa;bbb", cllBttnsMsg_1)
    Debug.Assert cllBttnsMsg.Count = 4
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "bbb"
    Debug.Assert cllBttnsMsg(3) = "ccc"
    Debug.Assert cllBttnsMsg(4) = "ddd"
    
    Set cllBttnsMsg = Nothing
    Set cllBttnsMsg_1 = Nothing
    EoP ErrSrc(PROC)
End Sub

Private Sub Test_01_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary()
    Const PROC = "Test_01_mMsg_Buttons_Service_09_Comma_Delimited_String_Dictionary"
    Dim dct   As New Dictionary
    
    BoP ErrSrc(PROC)
    dct.Add "ccc", "ccc"
    dct.Add "ddd", "ddd"
    
    Set cllBttnsMsg = mMsg.Buttons("aaa,bbb", dct)
    Debug.Assert cllBttnsMsg.Count = 4
    Debug.Assert cllBttnsMsg(1) = "aaa"
    Debug.Assert cllBttnsMsg(2) = "bbb"
    Debug.Assert cllBttnsMsg(3) = "ccc"
    Debug.Assert cllBttnsMsg(4) = "ddd"
    
    Set cllBttnsMsg = Nothing
    Set dct = Nothing
    EoP ErrSrc(PROC)
End Sub

Public Function Test_02_mMsg_ErrMsg_Service() As Variant
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
    Const PROC = "Test_02_mMsg_ErrMsg_Service"
    Const EXPECTED_TITLE = "Application Error  5 in: 'mMsgTestServices.Test_02_mMsg_ErrMsg_Service'"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 2, EXPECTED_TITLE
    
    mErH.Asserted AppErr(5) ' skips the display of the error message when mErH.Regression = True
    
    Err.Raise Number:=AppErr(5) _
            , source:=ErrSrc(PROC) _
            , Description:="This is a test error description!||This is part of the error description, " & _
                           "concatenated by a double vertical bar and therefore displayed as an additional 'About the error' section " & _
                           "- one of the specific features of the mMsg.ErrMsg service."
        
xt: EoP ErrSrc(PROC)
    mTest.Evaluate PROC
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      Test_02_mMsg_ErrMsg_Service = vbOK: GoTo xt
    End Select
End Function

Public Function Test_03_mMsg_Dsply_Service_WidthDeterminedByMinimumWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_03_mMsg_Dsply_Service_WidthDeterminedByMinimumWidth"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 3, PROC
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
        ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
        lTestMsgWidthIncrDecr = .MsgWidthIncrDecr
    End With
    If lTestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    
    vButton4 = "Repeat with minimum width" & vbLf & "+ " & PrcPnt(lTestMsgWidthIncrDecr, "w")
    vButton5 = "Repeat with minimum width" & vbLf & "- " & PrcPnt(lTestMsgWidthIncrDecr, "w")
    
    Set cllBttnsMsg = Nothing
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, vButton4, vButton5)
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & PrcPnt(lTestMsgWidthMin, "w") & ") or " _
                   & "to the width determined by the reply buttons."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "1. The message form height is adjusted to the required height up to the specified " & _
                     "maximum heigth which for this test is " & PrcPnt(lTestMsgHeightMax, "h") & " and not exceeded." & vbLf & _
                     "2. The minimum width limit for this test is " & PrcPnt(20, "w") & " and the maximum width limit for this test is " & PrcPnt(99, "w") & "."
        .Text.FontColor = rgbRed
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_04_mMsg_Dsply_Service_WidthDeterminedByTitle() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_04_mMsg_Dsply_Service_WidthDeterminedByTitle"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)

    mTest.SetupMsgTitleInstanceAndNo 4, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the title's lenght."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please note:"
        .Text.Text = "The two message sections in this test do use a proportional font " & _
                     "and thus are adjusted to form width determined by other factors." & vbLf & _
                     "The message form height is adjusted to the need up to the specified " & _
                     "maximum heigth based on the screen height which for this test is " & _
                     PrcPnt(lTestMsgHeightMax, "h") & "."
    End With
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_max:=wsTest.MsgWidthMax _
             , dsply_width_min:=wsTest.MsgWidthMin _
             , dsply_height_max:=wsTest.MsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_05_mMsg_Dsply_Service_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_05_mMsg_Dsply_Service_WidthDeterminedByMonoSpacedMessageSection"
        
    On Error GoTo eh
    Dim sBttnRepeatMaxWidthIncreased    As String
    Dim sBttnRepeatMaxWidthDecreased    As String
    Dim sBttnRepeatMaxHeightIncreased   As String
    Dim sBttnRepeatMaxHeightDecreased   As String
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 5, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = Pnts(.MsgWidthMin, "w")
        lTestMsgWidthMax = Pnts(.MsgWidthMax, "w")
        lTestMsgWidthIncrDecr = Pnts(.MsgWidthIncrDecr, "w")
        lTestMsgHeightMin = Pnts(25, "h")
        lTestMsgHeightMax = Pnts(.MsgHeightMax, "h")
        lTestMsgHeightIncrDecr = Pnts(.MsgHeightIncrDecr, "h")
    End With
    If lTestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If lTestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    sBttnRepeatMaxWidthIncreased = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & PrcPnt(lTestMsgWidthIncrDecr, "w")
    sBttnRepeatMaxWidthDecreased = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & PrcPnt(lTestMsgWidthIncrDecr, "w")
    sBttnRepeatMaxHeightIncreased = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & PrcPnt(lTestMsgHeightIncrDecr, "h")
    sBttnRepeatMaxHeightDecreased = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & PrcPnt(lTestMsgHeightIncrDecr, "h")
    
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, sBttnRepeatMaxWidthIncreased, sBttnRepeatMaxWidthDecreased)
    
    AssertWidthAndHeight lTestMsgWidthMin _
                       , lTestMsgWidthMax _
                       , lTestMsgHeightMin _
                       , lTestMsgHeightMax
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The length of the longest monospaced message section line determines the width of the message form - " & _
                     "provided it does not exceed the specified maximum form width which for this test is " & PrcPnt(lTestMsgWidthMax, "w") & " " & _
                     "of the screen size. The maximum form width may be incremented/decremented by " & PrcPnt(lTestMsgWidthIncrDecr, "w") & " in order to test the result."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "Initally, the message form width is adjusted to the longest line in the " & _
                     "monospaced message section and all other message sections are adjusted " & _
                     "to this (enlarged) width." & vbLf & _
                     "When the maximum form width is reduced by " & PrcPnt(lTestMsgWidthIncrDecr, "w") & " the monospaced message section is displayed with a horizontal scrollbar."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                     "  the message text is not ""word wrapped""." & vbLf & _
                     "- The message form height is adjusted to the need up to the specified maximum heigth" & vbLf & _
                     "  based on the screen height which for this test is " & PrcPnt(lTestMsgHeightMax, "h") & "."
        .Text.MonoSpaced = True
        .Text.FontUnderline = False
    End With
                        
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_min:=lTestMsgHeightMin _
             , dsply_height_max:=lTestMsgHeightMax

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_06_mMsg_Dsply_Service_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_06_mMsg_Dsply_Service_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    Dim OneBttnMore         As String
    Dim OneBttnLess         As String
    
    BoP ErrSrc(PROC)
    
    mTest.SetupMsgTitleInstanceAndNo 6, PROC
    
    ' Initializations for this test
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    lTestMsgWidthMax = wsTest.MsgWidthMax
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                     "which is a percentage of the screen height (for this test = " & PrcPnt(lTestMsgHeightMax, "h") & "."
    End With
    OneBttnMore = "Repeat with one button more"
    OneBttnLess = "Repeat with one button less"
    vButton6 = "The one more buttonn"
    
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, OneBttnLess, vButton6)
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_07_mMsg_Dsply_Service_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_07_mMsg_Dsply_Service_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh

    BoP ErrSrc(PROC)
    
    mTest.SetupMsgTitleInstanceAndNo 7, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & PrcPnt(lTestMsgWidthMax, "w") & "."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The monospaced message section comes with a horizontal scrollbar."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "This (single line!) monspaced message section exceeds the specified maximum form width which for this test is " & PrcPnt(lTestMsgWidthMax, "w") & "."
        .Text.MonoSpaced = True
    End With
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_08_mMsg_Dsply_Service_MonoSpacedMessageSectionExceedsMaxHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_08_mMsg_Dsply_Service_MonoSpacedMessageSectionExceedsMaxHeight"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    
    mTest.SetupMsgTitleInstanceAndNo 8, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
       
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The overall message window height exceeds the for this test specified maximum of " & _
                     PrcPnt(lTestMsgHeightMax, "h") & " of the screen height. Because the monospaced section " & _
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
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_09_mMsg_Dsply_Service_ButtonsOnly() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_09_mMsg_Dsply_Service_ButtonsOnly"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim j                   As Long
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 9, PROC
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        lTestMsgWidthMax = .MsgWidthMax:     lTestMsgWidthIncrDecr = .MsgWidthIncrDecr
        lTestMsgWidthMin = .MsgWidthMin:     lTestMsgHeightIncrDecr = .MsgWidthIncrDecr
        lTestMsgHeightMax = .MsgHeightMax
    End With
    If lTestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
    If lTestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
'    For i = 1 To 4 ' rows
    For i = 1 To 1 ' rows
        If i > 1 Then cllBttnsMsg.Add vbLf
'        For j = 1 To 3
        For j = 1 To 2
            cllBttnsMsg.Add "Click " & i & "-" & j & " in case ...." & vbLf & "(instead of a lengthy" & vbLf & "message text above)"
        Next j
    Next i
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, cllBttnsMsg)
    
    mMsg.MsgInstance(sMsgTitle).VisualizeForTest = wsTest.VisualizeForTest
    '~~ Obtain initial test values from the Test Worksheet
                     
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_button_default:=BTTN_PASSED _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax
             
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_10_mMsg_Dsply_Service_ButtonsMatrix() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_10_mMsg_Dsply_Service_ButtonsMatrix"
    
    On Error GoTo eh
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim i, j                As Long
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
        
    BoP ErrSrc(PROC)
    sMsgTitle = "Just to demonstrate what's theoretically possible: Buttons only! Finish with " & BTTN_PASSED & " (default) or " & BTTN_FAILED
    mTest.SetupMsgTitleInstanceAndNo 10, PROC
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        lTestMsgWidthMax = .MsgWidthMax:   lChangeWidthPcntg = .MsgWidthIncrDecr
        lTestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
'    If lTestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
'    If lTestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    

    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 2 To 7 ' rows
        For j = 1 To 7 ' row buttons
            cllBttnsMsg.Add "Button" & vbLf & i & "-" & j
        Next j
    Next i
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, cllBttnsMsg)
    '~~ Obtain initial test values from the Test Worksheet
    mMsg.MsgInstance(sMsgTitle).VisualizeForTest = wsTest.VisualizeForTest
                         
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_button_reply_with_index:=False _
             , dsply_button_default:=BTTN_PASSED _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_11_mMsg_Dsply_Service_ButtonScrollBarVertical() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_11_mMsg_Dsply_Service_ButtonScrollBarVertical"
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    
    BoP ErrSrc(PROC)
    
    mTest.SetupMsgTitleInstanceAndNo 11, PROC
    
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin:   lChangeMinWidthPt = .MsgWidthIncrDecr
        lTestMsgWidthMax = .MsgWidthMax:     lChangeWidthPcntg = .MsgWidthIncrDecr
        lTestMsgHeightMax = .MsgHeightMax: lChangeHeightPcntg = .MsgHeightIncrDecr
    End With
'    If lTestMsgWidthIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Width increment/decrement must not be 0 for this test!"
'    If lTestMsgHeightIncrDecr = 0 Then Err.Raise AppErr(1), ErrSrc(PROC), "Height increment/decrement must not be 0 for this test!"
    
    '~~ Obtain initial test values from the Test Worksheet
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                     "the specified maximum forms height - which for this test has been limited to " & _
                     PrcPnt(lTestMsgHeightMax, "h") & " of the screen height."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The height for the vertically ordered buttons is reduced to fit the specified " & _
                     "maximum message form heigth and a vertical scrollbar is applied."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Finish test:"
        .Text.Text = "Click " & BTTN_PASSED & " or " & BTTN_FAILED & " (test is repeated with any other button)"
    End With
    
    For i = 1 To 5
        If i > 1 Then Set cllBttnsMsg = mMsg.Buttons(cllBttnsMsg, vbLf)
        For j = 1 To 2
            Set cllBttnsMsg = mMsg.Buttons(cllBttnsMsg, "Reply" & vbLf & "Button" & vbLf & i & "-" & j)
        Next j
    Next i
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, cllBttnsMsg)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_12_mMsg_Dsply_Service_ButtonScrollBarHorizontal() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_12_mMsg_Dsply_Service_ButtonScrollBarHorizontal"
    Const INIT_WIDTH    As String = 40
    Const CHANGE_WIDTH  As String = 10
    
    On Error GoTo eh
    Dim Bttn10Plus          As String
    Dim Bttn10Minus         As String
    
    BoP ErrSrc(PROC)
    
    mTest.SetupMsgTitleInstanceAndNo 12, PROC
    
    lTestMsgWidthMax = INIT_WIDTH
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With

    mMsg.MsgInstance(sMsgTitle).VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The button's width (determined by the longest buttons caption text line), " & _
                     "their number, and the button's order (all in one row) exceeds the form's " & _
                     "maximum width, explicitly specified for this test as " & _
                     PrcPnt(lTestMsgWidthMax, "w") & " of the screen width."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons are dsiplayed with a horizontal scroll bar to meet the specified maximimum form width."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Finish test:"
        .Text.Text = "This test is repeated with any button clicked other than the ""Ok"" button"
    End With
    
    Bttn10Plus = "Repeat with maximum form width" & vbLf & "extended by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(lTestMsgWidthMax, "w")
    Bttn10Minus = "Repeat with maximum form width" & vbLf & "reduced by " & PrcPnt(CHANGE_WIDTH, "w") & " to " & PrcPnt(lTestMsgWidthMax, "w")
        
    '~~ Obtain initial test values from the Test Worksheet
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, Bttn10Plus, Bttn10Minus)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_modeless:=mTest.MODE_LESS _
             , dsply_button_default:=BTTN_PASSED
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_13_mMsg_Dsply_Service_ButtonsMatrix_With_Both_Scroll_Bars() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_13_mMsg_Dsply_Service_ButtonsMatrix_With_Both_Scroll_Bars"
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim bMonospaced         As Boolean:         bMonospaced = True ' initial test value
    Dim lTestMsgWidthMin    As Long
    Dim lTestMsgHeightMax   As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 13, PROC
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test  from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
        
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 2 To 7 ' rows
        For j = 1 To 7 ' row buttons
            cllBttnsMsg.Add vbLf & " ---- Button ---- " & vbLf & i & "-" & j & vbLf & " "
        Next j
    Next i
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, cllBttnsMsg)
    
        '~~ Obtain initial test values from the Test Worksheet
        mMsg.MsgInstance(sMsgTitle).VisualizeForTest = wsTest.VisualizeForTest
                             
        mMsg.Dsply dsply_title:=sMsgTitle _
                 , dsply_msg:=udtMessage _
                 , dsply_label_spec:=vbNullString _
                 , dsply_buttons:=cllBttnsMsg _
                 , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
                 , dsply_button_reply_with_index:=False _
                 , dsply_button_default:=BTTN_PASSED _
                 , dsply_width_min:=lTestMsgWidthMin _
                 , dsply_width_max:=lTestMsgWidthMax _
                 , dsply_height_max:=lTestMsgHeightMax _
                 , dsply_modeless:=mTest.MODE_LESS
                 
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_16_mMsg_Dsply_Service_ButtonByDictionary()
' ------------------------------------------------------------------------------
' The buttons argument is provided as Dictionary.
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_16_mMsg_Dsply_Service_ButtonByDictionary"
    
    On Error GoTo xt
    Dim dct                 As New Collection
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 16, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
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
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=mMsg.Buttons(cllBttnsTest, vbLf, dct) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS

xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_17_mMsg_Box_Service_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_17_mMsg_Box_Service_MessageAsString"
        
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 17, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
        
    mMsg.Box Title:=sMsgTitle _
           , Prompt:="This is a message provided as a simple string argument!" _
           , Buttons:=cllBttnsMsg _
           , box_buttons_app_run:=mTest.BttnsAppRunArgs _
           , box_width_min:=lTestMsgWidthMin _
           , box_width_max:=lTestMsgWidthMax _
           , box_height_max:=lTestMsgHeightMax _
           , box_modeless:=mTest.MODE_LESS
           
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_20_mMsg_Dsply_Service_ButtonByValue()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_20_mMsg_Dsply_Service_ButtonByValue"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 20, PROC
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is a collection of the test buttons (Passed, Failed) and an additional button provided as value"
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The ""Ok"" button is displayed in the second row."
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=mMsg.Buttons(cllBttnsTest, vbLf, vbOKOnly) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_21_mMsg_Dsply_Service_ButtonByString()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_21_mMsg_Dsply_Service_ButtonByString"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 21, PROC
        
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE, vbLf, "Yes," & vbLf & ",No") _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_22_mMsg_Dsply_Service_ButtonByCollection()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_22_mMsg_Dsply_Service_ButtonByCollection"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 22, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    cllBttnsMsg.Add "Yes"
    cllBttnsMsg.Add "No"
    
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
      
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_23_mMsg_Dsply_Service_MonoSpacedSectionOnly()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_23_mMsg_Dsply_Service_MonoSpacedSectionOnly"
    Const LINES = 50
    
    On Error GoTo eh
    Dim Msg                 As String
    Dim i                   As Long

    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 23, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
            
    i = 1
    Msg = Format(i, "00: ") & Format(Now(), "YY-MM-DD hh:mm:ss") & " Test mono-spaced message section text exceeding the specified maximum width and height"
    For i = 2 To LINES
        Msg = Msg & vbLf & Format(i, "00: ") & Format(Now(), "YY-MM-DD hh:mm:ss") & " Test mono-spaced message section text exceeding the specified maximum width and height"
    Next i
    
    With udtMessage.Section(1).Text
        .Text = Msg
        .MonoSpaced = True
    End With
      
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
             
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_24_mMsg_All_Sections_Label_Pos_Left()
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_24_mMsg_All_Sections_Label_Pos_Left"
    
    On Error GoTo eh
    Dim MsgSection          As String
    Dim i                   As Long

    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 24, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest)
    MsgSection = "udtMessage text section "
    
    For i = 1 To mMsg.NoOfMsgSects
        With udtMessage.Section(i)
            .Label.Text = "Label section " & i
            .Label.FontColor = rgbGreen
            .Text.Text = MsgSection & i
        End With
    Next i
       
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:="L30" _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_30_mMsg_MonitorHeader_mMsg_Monitor_mMsg_MonitorFooter_Service() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_30_mMsg_MonitorHeader_mMsg_Monitor_mMsg_MonitorFooter_Service"
    
    On Error GoTo eh
    Dim i                   As Long
    Dim Header              As TypeMsgText
    Dim Step                As TypeMsgText
    Dim Footer              As TypeMsgText
    Dim iLoops              As Long
    Dim lWait               As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 30, PROC
    
    lTestMsgWidthMin = wsTest.MsgWidthMin
    lTestMsgWidthMax = wsTest.MsgWidthMax
    lTestMsgHeightMax = wsTest.MsgHeightMax
    
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
       
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
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
            .Text = Repeat(.Text & " ", Int(i / 5) + 1) & vbLf & "    Second line just for test " & Repeat(".", i)
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
    
    mTest.Evaluate PROC
        
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_40_mMsg_Dsply_Service_LabelPos_Left_R30()
' ------------------------------------------------------------------------------
' Test procedure for label pos left, width 30, various sections with and without
' label and/or text.
' ------------------------------------------------------------------------------
    Const PROC = "Test_40_mMsg_Dsply_Service_LabelPos_Left_R30"
    
    On Error GoTo eh
    Dim iSect               As Long
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 40, PROC
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    iSect = 1
    With udtMessage.Section(iSect)
        .Label.Text = "Label-" & iSect
        .Text.Text = DefaultSectionTextProp
        .Text.MonoSpaced = False
    End With
    iSect = iSect + 1
    With udtMessage.Section(iSect)
        .Label.Text = "Label-" & iSect
        .Text.Text = DefaultSectionTextProp
    End With
    iSect = iSect + 1
    With udtMessage.Section(iSect)
        .Label.Text = vbNullString
        .Text.Text = "Section without label *)"
    End With
    iSect = iSect + 1
    With udtMessage.Section(iSect)
        With .Label
            .Text = "*) Label-" & iSect & " (without text, uses full available width)"
            .OpenWhenClicked = "https://github.com/warbe-maker/VBA-udtMessage#readme"
        End With
        .Text.Text = vbNullString
    End With
       
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:="R60" _
             , dsply_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE, vbLf, vbOKOnly) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=30 _
             , dsply_width_max:=50 _
             , dsply_height_max:=70 _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_90_mMsg_Dsply_Service_AllInOne() As Variant
    Const PROC      As String = "Test_90_mMsg_Dsply_Service_AllInOne"

    On Error GoTo eh
    Dim i, j                As Long
    Dim Msg                 As TypeMsg
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 9, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
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
                   & "a) the for this demo specified maximum width of " & lTestMsgWidthMax & "% of the screen size" & vbLf _
                   & "   (defaults to 80% when not specified)" & vbLf _
                   & "b) the longest line of this section" & vbLf _
                   & "Because the text exeeds the specified maximum message width, a horizontal scroll-bar is displayed." & vbLf _
                   & "Due to this feature there is no message size limit other than the sytem's limit which for a string is about 1GB !!!!"
        .Text.MonoSpaced = True
    End With
    With Msg.Section(3)
        .Label.Text = "Unlimited message height (not the fact with this message):"
        .Label.FontColor = rgbBlue
        .Text.Text = "As with the message width, the message height is unlimited. When the maximum height (explicitly specified or the default) " _
                   & "is exceeded a vertical scroll-bar is displayed. Due to this feature there is no message size limit other than the sytem's " _
                   & "limit which for a string is about 1GB !!!!"
    End With
    With Msg.Section(4)
        .Label.Text = "Flexibility regarding the displayed reply buttons:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 7 x 7 = 49 possible reply buttons which may have any caption text " _
                   & "including the classic VBA.MsgBox values (vbOkOnly, vbYesNoCancel, etc.) - even in a mixture." & vbLf & vbLf _
                   & "!! This test ends with any button !!"
    End With
    '~~ Prepare the buttons collection
    
    For j = 1 To 2
        If j > 1 Then cllBttnsMsg.Add vbLf
        For i = 1 To 4
            cllBttnsMsg.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
    Next j
    
    Set cllBttnsMsg = mMsg.Buttons(cllBttnsTest, vbLf, cllBttnsMsg)
    
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=Msg _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=cllBttnsMsg _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_91_mMsg_Dsply_Service_MinimumMessage() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_91_mMsg_Dsply_Service_MinimumMessage"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 9, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
        
    With udtMessage.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsTest.TestDescription
    End With
    With udtMessage.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The width of all message sections is adjusted either to the specified minimum form width (" & PrcPnt(lTestMsgWidthMin, "w") & ") or " _
                   & "to the width determined by the reply buttons."
    End With
    With udtMessage.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height up to the specified " & _
                     "maximum heigth which is " & PrcPnt(lTestMsgHeightMax, "h") & " and not exceeded."
        .Text.FontColor = rgbRed
    End With
                                                                                                  
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=lTestMsgWidthMin _
             , dsply_width_max:=lTestMsgWidthMax _
             , dsply_height_max:=lTestMsgHeightMax _
             , dsply_modeless:=mTest.MODE_LESS
                         
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_92_mMsg_Dsply_Service_LabelWithUnderlayedURL() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_92_mMsg_Dsply_Service_LabelWithUnderlayedURL"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 92, PROC
    
    '~~ Obtain initial test values from the Test Worksheet
    With wsTest
        lTestMsgWidthMin = .MsgWidthMin
        lTestMsgWidthMax = .MsgWidthMax
        lTestMsgHeightMax = .MsgHeightMax
    End With
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
        
    With udtMessage.Section(1)
        .Label.Text = "Public github repo Common-VBA-udtMessage-Service"
        .Label.OpenWhenClicked = "https://github.com/warbe-maker/Common-VBA-udtMessage-Service"
        .Text.Text = "The label above is underlayed with a url *)."
    End With
    With udtMessage.Section(2)
        .Label.Text = "About this feature of the 'Common-VBA-udtMessage-Service':"
        .Text.Text = "The Common-VBA-udtMessage-Service makes use of the 'Click' and the 'MouseMove' event " & _
                     "of the Label control to allow not only to open a URL but also to open a file or " & _
                     "start an application (open a Workbook, Word document, etc). Examples:"
    End With
    With udtMessage.Section(3).Text
        .Text = "Open a folder:       C:\TEMP\                 " & vbLf & _
                "Call the eMail app:  mailto:dash10@hotmail.com" & vbLf & _
                "Open a Url:          http://......            " & vbLf & _
                "Open a file:         C:\TEMP\TestThis   (opens a dialog for the selection of the app" & vbLf & _
                "Open an application: x:\my\workbooks\this.xlsb (opens Excel)"
        .MonoSpaced = True
    End With
    
    With udtMessage.Section(4).Text
        .Text = "*) 'https://github.com/warbe-maker/Common-VBA-udtMessage-Service'"
        .MonoSpaced = True
        .FontSize = 8
    End With
                                                                                                 
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=vbNullString _
             , dsply_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE, vbLf, vbOKOnly) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=40 _
             , dsply_width_max:=80 _
             , dsply_height_max:=70 _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_93_mMsg_Dsply_Service_LabelPos() As Variant
    Const PROC = "Test_93_mMsg_Dsply_Service_LabelPos"
    
    BoP ErrSrc(PROC)
'    Test_93_mMsg_Dsply_Service_LabelPos = Test_94_mMsg_Dsply_Service_LabelPos(t_lbl_spec:="R60" _
'                                                                            , t_lbl_1_txt:="Section-1:" & vbLf & "Label" & Chr$(160) _
'                                                                            , t_tbx_1_txt:=DefaultSectionTextProp _
'                                                                            , t_lbl_2_txt:="Section-2:" & vbLf & "Label" & Chr$(160) _
'                                                                            , t_tbx_2_txt:=DefaultSectionTextMono _
'                                                                            , t_tbx_2_mon:=True _
'                                                                              )
'    Test_93_mMsg_Dsply_Service_LabelPos = Test_94_mMsg_Dsply_Service_LabelPos(t_lbl_spec:="L60" _
'                                                                            , t_lbl_1_txt:="Section-1:" _
'                                                                            , t_tbx_1_txt:=DefaultSectionTextMono _
'                                                                            , t_tbx_1_mon:=True _
'                                                                            , t_lbl_2_txt:="Section-2:" _
'                                                                            , t_tbx_2_txt:=DefaultSectionTextProp _
'                                                                              )
'    Test_93_mMsg_Dsply_Service_LabelPos = Test_94_mMsg_Dsply_Service_LabelPos(t_lbl_spec:="C60" _
'                                                                            , t_lbl_1_txt:="Section-1:" _
'                                                                            , t_tbx_1_txt:=DefaultSectionTextProp _
'                                                                            , t_lbl_2_txt:="Section-2:" _
'                                                                            , t_tbx_2_txt:=DefaultSectionTextMono _
'                                                                            , t_tbx_2_mon:=True _
'                                                                              )
    Test_93_mMsg_Dsply_Service_LabelPos = Test_94_mMsg_Dsply_Service_LabelPos(t_lbl_spec:=vbNullString _
                                                                            , t_lbl_1_txt:=vbNullString _
                                                                            , t_tbx_1_txt:=DefaultSectionTextProp _
                                                                            , t_lbl_2_txt:="Section-2:" _
                                                                            , t_tbx_2_txt:=DefaultSectionTextMono _
                                                                            , t_tbx_2_mon:=True _
                                                                              )
    EoP ErrSrc(PROC)

End Function

Public Function Test_94_mMsg_Dsply_Service_LabelPos(Optional ByVal t_lbl_spec As String = vbNullString, _
                                                     Optional ByVal t_lbl_1_txt As String = vbNullString, _
                                                     Optional ByVal t_tbx_1_txt As String = vbNullString, _
                                                     Optional ByVal t_tbx_1_mon As Boolean = False, _
                                                     Optional ByVal t_lbl_2_txt As String = vbNullString, _
                                                     Optional ByVal t_tbx_2_txt As String = vbNullString, _
                                                     Optional ByVal t_tbx_2_mon As Boolean = False)
' ------------------------------------------------------------------------------
' Test procedure for label pos top, left aligned left, centered, or right
' ------------------------------------------------------------------------------
    Const PROC = "Test_94_mMsg_Dsply_Service_LabelPos"
    
    On Error GoTo eh
    
    BoP ErrSrc(PROC)
    mTest.SetupMsgTitleInstanceAndNo 93, PROC
    ufmMsg.VisualizeForTest = wsTest.VisualizeForTest
    
    With udtMessage.Section(1)
        .Label.Text = t_lbl_1_txt
        .Text.Text = "LabelPosSpec = " & t_lbl_spec
        .Text.MonoSpaced = t_tbx_1_mon
    End With
    With udtMessage.Section(2)
        .Label.Text = t_lbl_2_txt
        .Text.Text = t_tbx_2_txt
        .Text.MonoSpaced = t_tbx_2_mon
    End With
        
    mMsg.Dsply dsply_title:=sMsgTitle _
             , dsply_msg:=udtMessage _
             , dsply_label_spec:=t_lbl_spec _
             , dsply_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED, BTTN_TRMNTE, vbLf, vbOKOnly) _
             , dsply_buttons_app_run:=mTest.BttnsAppRunArgs _
             , dsply_width_min:=30 _
             , dsply_width_max:=50 _
             , dsply_height_max:=70 _
             , dsply_modeless:=mTest.MODE_LESS
    
xt: EoP ErrSrc(PROC)
    Set cllBttnsMsg = Nothing
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

