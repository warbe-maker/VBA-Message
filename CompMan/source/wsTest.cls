VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private bRegressionTest As Boolean
Private bTerminated     As Boolean

Public Property Get Value(Optional ByVal v_col As String) As Variant
    Value = Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range(v_col).EntireColumn).Value
End Property
Public Property Let Value(Optional ByVal v_col As String, _
                                   ByVal v_value As Variant)
    Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range(v_col).EntireColumn).Value = v_value
End Property

Public Property Get FailedText() As String
    FailedText = Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range("colTestPassedFailedNote").EntireColumn).Value
End Property

Public Property Get FormHeightMax() As Long
    FormHeightMax = Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range("colMsgMaxHeight").EntireColumn).Value
End Property

Public Property Let FormHeightMax(ByVal l As Long)
    Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range("colMsgMaxHeight").EntireColumn).Value = l
End Property

Public Property Get FormWidthMax() As Long:             FormWidthMax = Value("colMsgMaxWidth"):                     End Property

Public Property Let FormWidthMax(ByVal l As Long):      Value("colMsgMaxWidth") = l:                                End Property

Public Property Get FormWidthMin() As Long:             FormWidthMin = Value("colMsgMinWidth"):                     End Property

Public Property Let FormWidthMin(ByVal l As Long):      Value("colMsgMinWidth") = l:                                End Property

Public Property Get LabelPos() As enLabelPos:           LabelPos = mMsg.LabelPos(MsgLabelPosSpec):                  End Property

Public Property Get LabelWidth() As Long:               LabelWidth = mMsg.LabelWidth(MsgLabelPosSpec):              End Property

Public Property Get MsgLabelPosSpec() As String:        MsgLabelPosSpec = Value("colMsgLabelPosSpec"):              End Property

Public Property Let MsgLabelPosSpec(ByVal s As String): Value("colMsgLabelPosSpec") = s:                            End Property

Public Property Get NextTestNumber() As Long
' -------------------------------------------------------------------------
' Returns the number of the next consequtive test indicated for regression
' ("R"). Returns 0 if there's no next test.
' -------------------------------------------------------------------------
    Dim rng     As Range
    Dim iRow    As Long
    
    NextTestNumber = 0
    For Each rng In TestProcIds
        If rng.Value = vbNullString Then Exit Property  ' End of test reached
        If rng.Value > mMsgTest.ProcId Then                ' This is the current/last test
            If rng.OFFSET(iRow, 2).Value = "R" Then
                NextTestNumber = rng.Value
                Exit Property
            End If
        End If
    Next rng
    
End Property

Public Property Get RegressionTest() As Boolean:                    RegressionTest = bRegressionTest:                               End Property

Public Property Let RegressionTest(ByVal b As Boolean):             bRegressionTest = b:                                            End Property

Public Property Get RegressionTests() As Range
    Set RegressionTests = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colRegression").EntireColumn)
End Property

Public Property Get Terminated() As Boolean:                        Terminated = bTerminated:                                       End Property

Public Property Let Terminated(ByVal b As Boolean)
    bTerminated = b
    If bTerminated Then bRegressionTest = False
End Property

Public Property Get TestDescription() As String:        TestDescription = Value("colTestDescription"):  End Property

Public Property Let TestDescription(ByVal s As String): Value("colTestDescription") = s:                End Property

Public Property Get TestLabelPos() As enLabelPos
    Select Case True
        Case Me.Shapes("optLabelPosTop").OLEFormat.Object.Value = 1:                TestLabelPos = enLabelAboveSectionText
        Case Me.Shapes("optLabelPosLeftAlignedLeft").OLEFormat.Object.Value = 1:    TestLabelPos = enLposLeftAlignedLeft
        Case Me.Shapes("optLabelPosLeftAlignedRight").OLEFormat.Object.Value = 1:   TestLabelPos = enLposLeftAlignedRight
    End Select
End Property

Private Property Get TestProcIds() As Range
    Set TestProcIds = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colTestProcId").EntireColumn)
End Property

Public Property Get TestOptionModelessMessageDisplay() As Boolean
    Const PROC          As String = "TestOptionModelessMessageDisplay_Get"
    Const SHAPE_NAME    As String = "optDisplayModeless"
    
    On Error GoTo eh

'    If wsTest.RegressionTest Then GoTo xt ' Regression and ModeLess is not supported
    If Not mObject.WsShapeExists(wse_ws:=Me, wse_shape_name:=SHAPE_NAME) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A shape named '" & SHAPE_NAME & "' does not exist in Worksheet '" & Me.Name & "'!"
        
    If Me.Shapes("optDisplayModeless").OLEFormat.Object.Value = 1 _
    Then TestOptionModelessMessageDisplay = True _
    Else TestOptionModelessMessageDisplay = False

xt: Exit Property

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Private Property Let TestResultNote(ByVal s As String):             Value(Me.Range("colTestPassedFailedNote")) = s:                             End Property

Public Property Get VisualizeForTest() As Boolean
' ------------------------------------------------------------------------------
' Returns the status True or False of the shape named SHAPE_NAME.
' ------------------------------------------------------------------------------
    Const PROC          As String = "VisualizeForTest"
    Const SHAPE_NAME    As String = "optVisualizeForTest"
    
    On Error GoTo eh
    If Not mObject.WsShapeExists(wse_ws:=Me, wse_shape_name:=SHAPE_NAME) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A shape named '" & SHAPE_NAME & "' does not exist in Worksheet '" & Me.Name & "'!"
    
    If Me.Shapes(SHAPE_NAME).OLEFormat.Object.Value = 1 _
    Then VisualizeForTest = True _
    Else VisualizeForTest = False
    
xt: Exit Property

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
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

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Common, minimum VBA error handling providing the means to resume the error
' line when the Conditional Compile Argument Debugging=1.
' Usage: When this procedure is copied into any desired module the statement
'        If ErrMsg(ErrSrc(PROC) = vbYes Then: Stop: Resume
'        is appropriate
'        The caller provides the source of the error through ErrSrc(PROC) where
'        ErrSrc is a procedure available in the module using this ErrMsg and
'        PROC is the constant identifying the procedure
' Uses: AppErr to translate a negative programmed application error into its
'              original positive number
' ------------------------------------------------------------------------------
    Dim ErrNo       As Long
    Dim ErrDesc     As String
    Dim ErrType     As String
    Dim ErrLine     As Long
    Dim ErrAtLine   As String
    Dim ErrBttns    As Long
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    
    If err_line = 0 Then ErrLine = Erl
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    ErrDesc = "Error: " & vbLf & err_dscrptn & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine

    
#If Debugging = 1 Then
    ErrBttns = vbYesNo
    ErrDesc = ErrDesc & vbLf & vbLf & "Debugging: Yes=Resume error line, No=Continue"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrType & ErrNo & " in " & err_source _
                  , Prompt:=ErrDesc _
                  , Buttons:=ErrBttns)
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "wsTest." & s
End Function

Public Sub Failed()
    Value("colTestPassedFailed") = "Failed"
    With Intersect(RowTest(mMsgTest.ProcId).EntireRow, Range("colTestPassedFailedNote").EntireColumn)
        .Value = _
        VBA.InputBox(Prompt:="Enter the failed text" _
                   , Title:="Test failed" _
                   , Default:=.Value _
                   , XPos:=.Top _
                   , YPos:=.Left _
                )
        Me.Activate
        Application.EnableEvents = False
        .Select
        Application.EnableEvents = True
    End With
End Sub

Public Sub SelectDescription()
    Application.EnableEvents = False
    Intersect(RowTest(mMsgTest.ProcId).EntireRow, Me.Range("colTestDescription").EntireColumn).Select
    Application.EnableEvents = True
End Sub

Public Sub Passed()
    Value("colTestPassedFailed") = "Passed"
    Value("colTestPassedFailedNote") = vbNullString
End Sub

Public Function RowTest(ByVal r_test_proc_id As String) As Range
' ------------------------------------------------------------------------------
' Returns the entire row indicated by the test number. When the row cannot be
' identified by its name it will be identified computed.
' ------------------------------------------------------------------------------
    Const PROC = "RowTest"
    
    On Error GoTo eh
    Dim rng As Range
    
    With Me
        For Each rng In Intersect(.Range("colTestProcId").EntireColumn, .UsedRange.EntireColumn).Cells
            With rng
                If .Value = r_test_proc_id Then
                    Set RowTest = .EntireRow
                    Exit Function
                End If
            End With
        Next rng
    End With
    Err.Raise AppErr(1), ErrSrc(PROC), "There is no test row with the name '" & _
                                       r_test_proc_id & "' in column 'colTestProcId'!"
xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub Worksheet_Activate()
    Dim rng As Range
    Application.ScreenUpdating = False
    Set rng = Selection
    wsTest.Range("rngZoom").Select
    ActiveWindow.Zoom = True
    rng.Select
    Application.EnableEvents = True
End Sub

