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
Private bTerminated As Boolean

Private Property Get ColMsgHeightIncrDecr() As Range:               Set ColMsgHeightIncrDecr = Me.Range("colMsgHeightIncDecr"):                                 End Property

Private Property Get ColMsgHeightMaxSpecAsPoSS() As Range:          Set ColMsgHeightMaxSpecAsPoSS = Me.Range("colMsgHeightMaxSpecAsPoSS"):                      End Property

Private Property Get ColMsgLabelPosSpec() As Range:                 Set ColMsgLabelPosSpec = Me.Range("colMsgLabelPosSpec"):                                    End Property

Private Property Get ColMsgWidthIncrDecr() As Range:                Set ColMsgWidthIncrDecr = Me.Range("colMsgWidthIncDecr"):                                   End Property

Private Property Get ColMsgWidthMax() As Range:                     Set ColMsgWidthMax = Me.Range("colMsgWidthMax"):                                            End Property

Private Property Get ColMsgHeightMax() As Range:                    Set ColMsgHeightMax = Me.Range("colMsgHeightMax"):                                            End Property

Private Property Get ColMsgWidthMin() As Range:                     Set ColMsgWidthMin = Me.Range("colMsgWidthMin"):                                            End Property

Private Property Get FailedText() As String
    Intersect(Range("rowTest" & mTest.Number), Range("colTestPassedFailed")).Value = "Failed"
    FailedText = Intersect(RowTest(mTest.Number).EntireRow, Range("colTestPassedFailedNote").EntireColumn).Value
End Property

Public Property Get MsgHeightIncrDecr() As Long:                    MsgHeightIncrDecr = Intersect(RowTest(mTest.Number), ColMsgHeightIncrDecr).Value:           End Property

Public Property Get MsgHeightMax() As Long:                         MsgHeightMax = TestValue(ColMsgHeightMax).Value:                                            End Property

Public Property Get MsgLabelPosSpec() As String:                    MsgLabelPosSpec = TestValue(ColMsgLabelPosSpec).Value:                                      End Property

Public Property Let MsgLabelPosSpec(ByVal s As String):             TestValue(ColMsgLabelPosSpec).Value = s:                                                    End Property

Public Property Get MsgWidthIncrDecr() As Long:                     MsgWidthIncrDecr = Intersect(RowTest(mTest.Number), ColMsgWidthIncrDecr).Value:             End Property

Public Property Get MsgWidthMax() As Single:                        MsgWidthMax = Intersect(RowTest(mTest.Number), ColMsgWidthMax).Value:                       End Property

Public Property Get MsgWidthMin() As Single
    On Error Resume Next
    MsgWidthMin = Intersect(RowTest(mTest.Number).EntireRow, Me.Range("colMsgWidthMin").EntireColumn).Value
    If Err.Number <> 0 Then Stop
End Property

Public Property Get NextTestNumber() As Long
' -------------------------------------------------------------------------
' Returns the number of the next consequtive test indicated for regression
' ("R"). Returns 0 if there's no next test.
' -------------------------------------------------------------------------
    Dim Rng     As Range
    Dim iRow    As Long
    
    NextTestNumber = 0
    For Each Rng In TestNumbers
        If Rng.Value = vbNullString Then Exit Property  ' End of test reached
        If Rng.Value > mTest.Number Then                ' This is the current/last test
            If Rng.OFFSET(iRow, 2).Value = "R" Then
                NextTestNumber = Rng.Value
                Exit Property
            End If
        End If
    Next Rng
    
End Property

Public Property Get RegressionTest() As Boolean:                    RegressionTest = bRegressionTest:                                                           End Property

Public Property Let RegressionTest(ByVal b As Boolean):             bRegressionTest = b:                                                                        End Property

Public Property Get RegressionTests() As Range
    Set RegressionTests = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colRegression").EntireColumn)
End Property

Private Property Get RowTest(Optional ByVal Test As Long) As Range: Set RowTest = Me.Range("rowTest" & mTest.Number):                                      End Property

Public Property Get Terminated() As Boolean:           Terminated = bTerminated:                                         End Property

Public Property Let Terminated(ByVal b As Boolean)
    bTerminated = b
    If bTerminated Then bRegressionTest = False
End Property

Public Property Get TestDescription()
    With Me
        Intersect(.Range("rowTest" & mTest.Number), .Range("colTestDescription")).Select
        TestDescription = Intersect(.Range("rowTest" & mTest.Number), .Range("colTestDescription")).Value
    End With
End Property

Public Property Get TestLabelPos() As enLabelPos
    Select Case True
        Case Me.Shapes("optLabelPosTop").OLEFormat.Object.Value = 1:                TestLabelPos = enLabelAboveSectionText
        Case Me.Shapes("optLabelPosLeftAlignedLeft").OLEFormat.Object.Value = 1:    TestLabelPos = enLposLeftAlignedLeft
        Case Me.Shapes("optLabelPosLeftAlignedRight").OLEFormat.Object.Value = 1:   TestLabelPos = enLposLeftAlignedRight
    End Select
End Property

Public Property Get LabelPosSpec(Optional ByVal l_test_no As Long)
    LabelPosSpec = vbNullString
End Property

Private Property Get TestNumbers() As Range
    Set TestNumbers = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colTestNo").EntireColumn)
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

Private Property Let TestResultNote(ByVal s As String)
    Intersect(Range("rowTest" & mTest.Number), Range("colTestPassedFailedNote")).Value = s
End Property

Private Property Get TestValue(ByVal r As Range):       Set TestValue = Intersect(RowTest(mTest.Number), r.EntireColumn): End Property

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
    Intersect(Range("rowTest" & mTest.Number), Range("colTestPassedFailed")).Value = "Failed"
    With Intersect(RowTest(mTest.Number).EntireRow, Range("colTestPassedFailedNote").EntireColumn)
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

Public Sub Passed()
    Intersect(Range("rowTest" & mTest.Number), Range("colTestPassedFailed")).Value = "Passed"
    With Intersect(RowTest(mTest.Number).EntireRow, Range("colTestPassedFailedNote").EntireColumn)
        .ClearContents
        Application.EnableEvents = False
        .Select
        Application.EnableEvents = True
    End With
End Sub

Private Sub Worksheet_Activate()
    Dim Rng As Range
    Application.ScreenUpdating = False
    Set Rng = Selection
    wsTest.Range("rngZoom").Select
    ActiveWindow.Zoom = True
    Rng.Select
    Application.EnableEvents = True
End Sub

