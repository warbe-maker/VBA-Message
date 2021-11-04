Attribute VB_Name = "mProcTest"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mProcTest
'          Test of procedures - rather than fMsg/mMsg services/functions.
'
' ------------------------------------------------------------------------------
Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mFuncTest." & s:  End Property

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

Public Sub Test_DisplayWithWithoutFrames()
    Const PROC = "Test_DisplayWithWithoutFrames"
    
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    MsgTitle = "With frames test"
    Set MsgForm = mMsg.Form(frm_caption:=MsgTitle, frm_caller:=ErrSrc(PROC))
    
    MsgForm.DsplyFrmsWthBrdrsTestOnly = True
    
    mMsg.Box box_title:="With frames test" _
           , box_msg:="Message should be displayed with visible frames"

    mMsg.Box box_title:="With frames test" _
           , box_msg:="Message should be displayed with frames invible"
           
End Sub

Public Sub Test_AutoSizeTextBox_Width_Limited()
    Const PROC = "Test_AutoSizeTextBox_Width_Limited"
    
    Dim i                   As Long
    Dim iFrom               As Long
    Dim iStep               As Long
    Dim iTo                 As Long
    Dim TestAppend          As Boolean
    Dim TestAppendMargin    As String
    Dim TestHeightMax       As Single
    Dim TestHeightMin       As Single
    Dim TestWidthLimit      As Single
    Dim TestWidthMax        As Single
    
    iFrom = 400
    iStep = -100
    iTo = 200
    TestAppend = True
    TestAppendMargin = vbLf
    TestHeightMin = 0
    TestHeightMax = 120
    TestWidthMax = 310
    
again:
    With fProcTest
        .top = 0
        .Left = 0
        .show False
        
        For TestWidthLimit = iFrom To iTo Step iStep
            i = i + 1
            .Caption = PROC
            .frm.Width = TestWidthLimit + 3
            .frm.Left = 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .tbx.Left = 0
            .tbx.top = 0
            .tbx.ForeColor = rgbBlue

            .AutoSizeTextBox as_tbx:=.tbx _
                           , as_width_limit:=TestWidthLimit _
                           , as_height_min:=TestHeightMin _
                           , as_text:="For this " & i & ". test the width is limited to " & TestWidthLimit & ". " & _
                                      "The height is determined at first by the height resulting from the AutoSize " & _
                                      "and second by the provided minimum height which for this test is " & TestHeightMin & "." _
                           , as_width_max:=TestWidthMax _
                           , as_height_max:=TestHeightMax _
                           , as_append:=TestAppend _
                           , as_append_margin:=TestAppendMargin
            
            With .tbxTestAndResult
                .MultiLine = True
                .WordWrap = False
                With .Font
                    .Name = "Courier New"
                    .Size = 8
                End With
                .top = 5
                .AutoSize = True
            End With
            .tbxTestAndResult.Value = "Provided arguments:" & vbLf & _
                                      "-------------------" & vbLf & _
                                      "as_width_limit = " & TestWidthLimit & vbLf & _
                                      "as_height_min  = " & TestHeightMin & vbLf & _
                                      "as_width_max   = " & TestWidthMax & vbLf & _
                                      "as_height_max  = " & TestHeightMax & vbLf & _
                                      "as_append      = " & CStr(TestAppend) & vbLf & vbLf & _
                                      "Results:" & vbLf & _
                                      "--------" & vbLf & _
                                      "tbx.Width      = " & .tbx.Width & vbLf & _
                                      "tbx.Height     = " & .tbx.Height & vbLf & _
                                      "TestHeightMin  = " & TestHeightMin
                       
            .frm.top = .tbxTestAndResult.top + .tbxTestAndResult.Height + 5
            
            '~~ The UserForm's height is adjusted to the resulting frame size
            fProcTest.Height = .frm.top + .frm.Height + (fProcTest.Height - fProcTest.InsideHeight) + 5
            fProcTest.Width = .frm.Left + .frm.Width + (fProcTest.Width - fProcTest.InsideWidth) + 5
            
            If TestWidthLimit <> iTo Then
                Select Case MsgBox(Title:="Continue? > Yes, Finish > No, Terminate? > Cancel", Buttons:=vbYesNoCancel, Prompt:=vbNullString)
                    Case vbYes
                    Case vbNo:                          Exit Sub
                    Case vbCancel: Unload fProcTest: Exit Sub
                End Select
            Else
                Select Case MsgBox(Title:="Done? > Abort, Repeat? > Retry, Finish > Innore", Buttons:=vbAbortRetryIgnore, Prompt:=vbNullString)
                    Case vbAbort:   Unload fProcTest:   Exit Sub
                    Case vbRetry:   Unload fProcTest:   GoTo again
                    Case vbIgnore:  Exit Sub
                End Select
            End If
        Next TestWidthLimit
    End With

End Sub

Public Sub Test_AutoSizeTextBox_Width_Unlimited()
    Const PROC = "Test_AutoSizeTextBox_Width_Unlimited"
    
    Dim i               As Long
    Dim iFrom           As Long
    Dim iStep           As Long
    Dim iTo             As Long
    Dim TestAppend      As Boolean
    Dim TestHeightMax   As Single
    Dim TestHeightMin   As Single
    Dim TestWidthLimit  As Single
    Dim TestWidthtMax   As Single
    
    iFrom = 1
    iTo = 5
    iStep = 1
    TestAppend = True
    TestHeightMin = 200
    TestWidthLimit = 0

again:
    With fProcTest
        .show False
        .top = 0
        .Left = 0
        For i = iFrom To iTo Step iStep
            .Caption = PROC
            .frm.Left = 5
            .tbx.Left = 0
            .tbx.top = 0
            .tbx.ForeColor = rgbBlue
            
            .AutoSizeTextBox as_tbx:=.tbx _
                           , as_width_limit:=TestWidthLimit _
                           , as_height_min:=TestHeightMin _
                           , as_text:="This " & i & ". test is with an unlimited width. " & _
                                      "The width is determined by the longest text line and WordWrap = False. " & _
                                      "the provided height minimum is used for the TextBox even when not used." _
                           , as_append:=TestAppend
            
            With .tbxTestAndResult
                .MultiLine = True
                .WordWrap = False
                With .Font
                    .Name = "Courier New"
                    .Size = 8
                End With
                .top = 5
                .AutoSize = True
            End With
            .tbxTestAndResult.Value = "Provided arguments:" & vbLf & _
                                      "-------------------" & vbLf & _
                                      "as_width_limit = " & TestWidthLimit & vbLf & _
                                      "as_height_min  = " & TestHeightMin & vbLf & _
                                      "as_append      = " & CStr(TestAppend) & vbLf & vbLf & _
                                      "Results:" & vbLf & _
                                      "--------" & vbLf & _
                                      "tbx.Width      = " & .tbx.Width & vbLf & _
                                      "tbx.Height     = " & .tbx.Height & vbLf & _
                                      "TestHeightMin  = " & TestHeightMin
            
            .frm.top = .tbxTestAndResult.top + .tbxTestAndResult.Height + 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .Height = .frm.top + .frm.Height + (.Height - .InsideHeight) + 5
            
            If TestWidthLimit <> iTo Then
                Select Case MsgBox(Title:="Continue? > Yes, Finish > No, Terminate? > Abbrechen", Buttons:=vbYesNoCancel, Prompt:=vbNullString)
                    Case vbYes
                    Case vbNo:                          Exit Sub
                    Case vbCancel: Unload fProcTest: Exit Sub
                End Select
            Else
                Select Case MsgBox(Title:="Done? > Abort, Repeat? > Retry, Finish > Ignore", Buttons:=vbAbortRetryIgnore, Prompt:=vbNullString)
                    Case vbAbort:   Unload fProcTest:   Exit Sub
                    Case vbRetry:   Unload fProcTest:   GoTo again
                    Case vbIgnore:  Exit Sub
                End Select
            End If
            
        
        Next i
    End With

End Sub

Public Sub Test_SetupTitle()
    fProcTest.show False
End Sub

Public Sub Test_AssertWidthAndHeight()
' ------------------------------------------------------------------------------
' - All values are returned as pt
' - All values are within their limit
' - Any min value above its max values is set equal to the max value
' ------------------------------------------------------------------------------

    Dim WidthMin    As Long
    Dim WidthMax    As Long
    Dim HeightMin   As Long
    Dim HeightMax   As Long
    
    '~~ Test 1: All values conform with their min/max limit
    WidthMin = MSG_WIDTH_MIN_LIMIT_PERCENTAGE
    WidthMax = MSG_WIDTH_MAX_LIMIT_PERCENTAGE
    HeightMin = MSG_HEIGHT_MIN_LIMIT_PERCENTAGE
    HeightMax = MSG_HEIGHT_MAX_LIMIT_PERCENTAGE
    
    mMsg.AssertWidthAndHeight WidthMin, WidthMax, HeightMin, HeightMax
    Debug.Assert WidthMin = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Debug.Assert WidthMax = Pnts(MSG_WIDTH_MAX_LIMIT_PERCENTAGE, "w")
    Debug.Assert HeightMin = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    Debug.Assert HeightMax = Pnts(MSG_HEIGHT_MAX_LIMIT_PERCENTAGE, "h")
    
    '~~ Test 2         : Min width > width max and height min > height max
    '~~ Expected result: The min values are set equal with their corresponding max value
    WidthMin = 41
    WidthMax = 40
    HeightMin = 31
    HeightMax = 30
    mMsg.AssertWidthAndHeight WidthMin, WidthMax, HeightMin, HeightMax
    Debug.Assert WidthMin = Pnts(40, "w")
    Debug.Assert WidthMax = Pnts(40, "w")
    Debug.Assert HeightMin = Pnts(30, "h")
    Debug.Assert HeightMax = Pnts(30, "h")
    
    '~~ Test 3          : Min values are less than their limit, max values are greater than their limit
    '~~ Expected results: All values are reset to their corresponding limit
    WidthMin = MSG_WIDTH_MIN_LIMIT_PERCENTAGE - 1
    WidthMax = MSG_WIDTH_MAX_LIMIT_PERCENTAGE + 1
    HeightMin = MSG_HEIGHT_MIN_LIMIT_PERCENTAGE - 1
    HeightMax = MSG_HEIGHT_MAX_LIMIT_PERCENTAGE + 1
    mMsg.AssertWidthAndHeight WidthMin, WidthMax, HeightMin, HeightMax
    Debug.Assert WidthMin = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Debug.Assert WidthMax = Pnts(MSG_WIDTH_MAX_LIMIT_PERCENTAGE, "w")
    Debug.Assert HeightMin = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    Debug.Assert HeightMax = Pnts(MSG_HEIGHT_MAX_LIMIT_PERCENTAGE, "h")
        
    '~~ Test 4         : All values are 0
    '~~ Expected result: All values are set to their corresponding limit
    WidthMin = 0
    WidthMax = 0
    HeightMin = 0
    HeightMax = 0
    mMsg.AssertWidthAndHeight WidthMin, WidthMax, HeightMin, HeightMax
    Debug.Assert WidthMin = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Debug.Assert WidthMax = Pnts(MSG_WIDTH_MAX_LIMIT_PERCENTAGE, "w")
    Debug.Assert HeightMin = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    Debug.Assert HeightMax = Pnts(MSG_HEIGHT_MAX_LIMIT_PERCENTAGE, "h")

End Sub
