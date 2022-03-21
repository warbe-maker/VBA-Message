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

Private lTestNumber                 As Long
Private bRegressionTest             As Boolean
Private bTerminateRegressionTest    As Boolean

Private Property Get ColMsgHeightIncrDecr() As Range:       Set ColMsgHeightIncrDecr = Me.Range("colMsgHeightIncDecr"):                                 End Property

Private Property Get ColMsgHeightMaxSpecAsPoSS() As Range:  Set ColMsgHeightMaxSpecAsPoSS = Me.Range("colMsgHeightMaxSpecAsPoSS"):                      End Property

Private Property Get ColMsgWidthIncrDecr() As Range:        Set ColMsgWidthIncrDecr = Me.Range("colMsgWidthIncDecr"):                                   End Property

Private Property Get ColMsgWidthMax() As Range:             Set ColMsgWidthMax = Me.Range("colMsgWidthMax"):                                            End Property

Private Property Get ColMsgWidthMin() As Range:             Set ColMsgWidthMin = Me.Range("colMsgWidthMin"):                                            End Property

Public Property Let Failed(ByVal b As Boolean)
    If b Then
        TestPassedFailed = "Failed!"
        TestPassedFailedNote.Select
    End If
End Property

Public Property Get MsgHeightIncrDecr() As Long:            MsgHeightIncrDecr = Intersect(RowTest(lTestNumber), ColMsgHeightIncrDecr).Value:            End Property

Public Property Get MsgHeightMax() As Long
    On Error Resume Next
    MsgHeightMax = Intersect(RowTest(lTestNumber), Me.Range("colMsgHeightMax").EntireColumn).Value
    If Err.Number <> 0 Then MsgHeightMax = 85 ' Default in case
End Property

Public Property Get MsgWidthIncrDecr() As Long:             MsgWidthIncrDecr = Intersect(RowTest(lTestNumber), ColMsgWidthIncrDecr).Value:              End Property

Public Property Get MsgWidthMax() As Single:                MsgWidthMax = Intersect(RowTest(lTestNumber), ColMsgWidthMax).Value:                        End Property

Public Property Get MsgWidthMin() As Single
    On Error Resume Next
    MsgWidthMin = Intersect(RowTest(lTestNumber).EntireRow, Me.Range("colMsgWidthMin").EntireColumn).Value
    If Err.Number <> 0 Then Stop
End Property

Public Property Let Passed(ByVal b As Boolean)
    If b Then
        TestPassedFailed = "Passed"
        TestPassedFailedNote.Select
    End If
End Property

Public Property Get RegressionTest() As Boolean:            wsTest.RegressionTest = bRegressionTest:                                                    End Property

Public Property Let RegressionTest(ByVal b As Boolean):     bRegressionTest = b:                                                                        End Property

Public Property Get RegressionTests() As Range
    Set RegressionTests = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colRegression").EntireColumn)
End Property

Private Property Get RowTest(Optional ByVal Test As Long) As Range: Set RowTest = Me.Range("rowTest" & wsTest.TestNumber):  End Property

Public Property Get TerminateRegressionTest() As Boolean:           TerminateRegressionTest = bTerminateRegressionTest:     End Property

Public Property Let TerminateRegressionTest(ByVal b As Boolean):    bTerminateRegressionTest = b:                           End Property

Public Property Get TestDescription()
    With Me
        Intersect(.Range("rowTest" & wsTest.TestNumber), .Range("colTestDescription")).Select
        TestDescription = Intersect(.Range("rowTest" & wsTest.TestNumber), .Range("colTestDescription")).Value
    End With
End Property

Public Property Get TestNumber() As Long:                           TestNumber = lTestNumber:                               End Property

Public Property Let TestNumber(ByVal l As Long):                    lTestNumber = l:                                        End Property

Private Property Get TestNumbers() As Range
    Set TestNumbers = Intersect(Me.Range("rngTests").EntireRow, Me.Range("colTestNo").EntireColumn)
End Property

Public Property Get TestOptionDisplayModeless() As Boolean
    Const PROC          As String = "TestOptionDisplayModeless_Get"
    Const SHAPE_NAME    As String = "optDisplayModeless"
    
    On Error GoTo eh

    If Not mObject.WsShapeExists(wse_ws:=Me, wse_shape_name:=SHAPE_NAME) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A shape named '" & SHAPE_NAME & "' does not exist in Worksheet '" & Me.name & "'!"
        
    If Me.Shapes("optDisplayModeless").OLEFormat.Object.Value = 1 _
    Then TestOptionDisplayModeless = True _
    Else TestOptionDisplayModeless = False

xt: Exit Property

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Property

Public Property Let TestPassed(ByVal b As Boolean)
    With Me
        If b _
        Then Intersect(.Range("rowTest" & wsTest.TestNumber), .Range("colTestPassedFailed")).Value = "Passed" _
        Else Intersect(.Range("rowTest" & wsTest.TestNumber), .Range("colTestPassedFailed")).Value = "Failed"
    End With
End Property

Private Property Let TestPassedFailed(ByVal s As String)
    Intersect(RowTest(lTestNumber).EntireRow, Range("colTestPassedFailed").EntireColumn).Value = s
End Property

Private Property Get TestPassedFailedNote() As Range
    Set TestPassedFailedNote = Intersect(RowTest(lTestNumber).EntireRow, Range("colTestPassedFailedNote").EntireColumn)
    TestPassedFailedNote.Worksheet.Activate ' just in case it is yet not active
End Property

Public Property Let TestResultNote(ByVal s As String)
    With Me
        Intersect(.Range("rowTest" & wsTest.TestNumber), .Range("colTestPassedFailedNote")).Value = s
    End With
End Property

Public Property Get VisualizeControls() As Boolean
' ------------------------------------------------------------------------------
' Returns the status True or False of the shape named SHAPE_NAME.
' ------------------------------------------------------------------------------
    Const PROC          As String = "VisualizeControls"
    Const SHAPE_NAME    As String = "optDisplayFrames"
    
    On Error GoTo eh
    If Not mObject.WsShapeExists(wse_ws:=Me, wse_shape_name:=SHAPE_NAME) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A shape named '" & SHAPE_NAME & "' does not exist in Worksheet '" & Me.name & "'!"
    
    If Me.Shapes(SHAPE_NAME).OLEFormat.Object.Value = 1 _
    Then VisualizeControls = True _
    Else VisualizeControls = False
    
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

Private Function NextRegressionTest() As Long
' -------------------------------------------------------------------------
' Returns the Testnumber of the next regression test, 0 if there's no next.
' -------------------------------------------------------------------------
    Dim Rng     As Range
    Dim iRow    As Long
    
    NextRegressionTest = 0
    For Each Rng In TestNumbers
        If Rng.Value = vbNullString Then Exit Function  ' End of test reached
        If Rng.Value > TestNumber Then                  ' This is the current/last test
            If Rng.OFFSET(iRow, 2).Value = "R" Then
                NextRegressionTest = Rng.Value
                Exit Function
            End If
        End If
    Next Rng
    
End Function

Private Sub Worksheet_Activate()
    Dim Rng As Range
    Application.ScreenUpdating = False
    Set Rng = Selection
    wsTest.Range("rngZoom").Select
    ActiveWindow.Zoom = True
    Rng.Select
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub

