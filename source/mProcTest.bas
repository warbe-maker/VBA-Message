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
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Displays a proper designe error message providing the option to resume the
' error line when the Conditional Compile Argument Debugging = 1.
' ------------------------------------------------------------------------------
    Dim ErrNo   As Long
    Dim ErrDesc As String
    Dim ErrType As String
    Dim ErrLine As Long
    Dim AtLine  As String
    Dim Buttons As Long
    Dim msg     As TypeMsg
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Application error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    
    If err_line = 0 Then ErrLine = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    With msg.Section(1)
        .Label.Text = "Error:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_dscrptn
    End With
    With msg.Section(2)
        .Label.Text = "Source:"
        .Label.FontColor = rgbBlue
        .Text.Text = err_source & AtLine
    End With

#If Debugging = 1 Then
    Buttons = vbYesNo
    With msg.Section(3)
        .Label.Text = "Debugging: (Conditional Compile Argument 'Debugging = 1')"
        .Label.FontColor = rgbBlue
        .Text.Text = "Yes = Resume error line, No = Continue"
    End With
    With msg.Section(4)
        .Label.Text = "About debugging:"
        .Label.FontColor = rgbBlue
        .Text.Text = "To make use of the debugging option have an error handling line" & vbLf & _
                     "eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume"
    End With
#Else
    Buttons = vbCritical
#End If
    
    ErrMsg = Dsply(dsply_title:=ErrType & ErrNo & " in " & err_source & AtLine _
                 , dsply_msg:=msg _
                 , dsply_buttons:=Buttons)
End Function

Private Function FormNew(ByVal uf_wb As Workbook, _
                         ByVal uf_name As String, _
                         ByVal uf_buttons As Variant) As UserForm
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "FormNew"
    
    On Error GoTo eh
    Dim MyUserForm          As VBComponent
    Dim NewCommandButton1   As Msforms.CommandButton
    Dim NewCommandButton2   As Msforms.CommandButton
    Dim N                   As Long
    Dim X                   As Long
    Dim MaxWidth            As Long
    Dim cmp                 As VBComponent
    Dim frm                 As UserForm
    Dim LeftPos             As Single
    
    '~~ Check the form doesn't already exist
    For Each cmp In uf_wb.VBProject.VBComponents
        If cmp.Name = uf_name Then
            Set FormNew = uf_wb.VBProject.VBComponents(uf_name)
            Exit Function
        End If
    Next cmp
     
    '~~ Create a new UserForm named uf_name
    Set cmp = uf_wb.VBProject.VBComponents.Add(vbext_ct_MSForm)
    DoEvents
    With cmp
        .Name = uf_name
        .Properties("Height") = 100
        .Properties("Width") = 200
        On Error Resume Next
        .Properties("Caption") = "UserForm named '" & uf_name & "'"
    End With
     
    '~~ Add buttons
    LeftPos = 10
    If uf_buttons = vbOKCancel Or uf_buttons = vbOKOnly Then
        ' Add an OK button to the form
        Set NewCommandButton2 = cmp.Designer.Controls.Add("forms.CommandButton.1")
        With NewCommandButton2
            .Caption = "OK"
            .Height = 18
            .Width = 44
            .Left = LeftPos
            LeftPos = LeftPos + .Width + 10
            .Top = 6
        End With
    End If
    
    If uf_buttons = vbOKCancel Or uf_buttons = vbYesNoCancel Or uf_buttons = vbRetryCancel Then
        ' Add a Cancel button to the form
        Set NewCommandButton1 = cmp.Designer.Controls.Add("forms.CommandButton.1")
        With NewCommandButton1
            .Caption = "Cancel"
            .Height = 18
            .Width = 44
            .Left = LeftPos
            .Top = 6
        End With
    End If
         
    '~~ Add code on the form for the CommandButtons
    With cmp.CodeModule
        X = .CountOfLines
        .InsertLines .CountOfLines + 1, "Option Explict"
        .InsertLines .CountOfLines + 1, vbNullString
        .InsertLines .CountOfLines + 1, "Sub CommandButton1_Click()"
        .InsertLines .CountOfLines + 1, "    Unload Me"
        .InsertLines .CountOfLines + 1, "End Sub"
        .InsertLines .CountOfLines + 1, vbNullString
        .InsertLines .CountOfLines + 1, "Sub CommandButton2_Click()"
        .InsertLines .CountOfLines + 1, "    Unload Me"
        .InsertLines .CountOfLines + 1, "End Sub"
    End With
     
    Set FormNew = cmp

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Sub FormRemove(ByVal wb As Workbook, _
                       ByVal FRM_NAME As String)
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "FormRemove"
    
    On Error GoTo eh
    Dim i As Long
    Dim cmp As VBComponent
    
    With wb.VBProject
        For Each cmp In .VBComponents
            If cmp.Name = FRM_NAME Then
                .VBComponents.Remove cmp
                Exit Sub
            End If
        Next cmp
    End With

xt: Exit Sub

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub

Private Function TestInstance(ByVal fi_key As String, _
                     Optional ByVal fi_unload As Boolean = False) As fProcTest
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fProcTest which is definitely
' identified by anything uniqe for the instance (fi_key). This may be what
' becomes the title (property Caption) or even an object such like a
' Worksheet (if the instance is Worksheet specific). An already existing or
' new created instance is maintained in a static Dictionary with fi_key as
' the key and returned to the caller. When fi_unload is true only a possibly
' already existing Userform identified by fi_key is unloaded.
'
' Requires: Reference to the "Microsoft Scripting Runtime".
' Usage   : The fProcTest has to be replaced by the name of the desired
'           UserForm
' -------------------------------------------------------------------------
    Const PROC = "TestInstance"
    
    On Error GoTo eh
    Static Instances As Dictionary    ' Collection of (possibly still)  active form instances
    
    If Instances Is Nothing Then Set Instances = New Dictionary
    
    If fi_unload Then
        If Instances.Exists(fi_key) Then
            On Error Resume Next
            Unload Instances(fi_key) ' The instance may be already unloaded
            Instances.Remove fi_key
        End If
        Exit Function
    End If
    
    If Not Instances.Exists(fi_key) Then
        '~~ There is no evidence of an already existing instance
        Set TestInstance = New fProcTest
        Instances.Add fi_key, TestInstance
    Else
        '~~ An instance identified by fi_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set TestInstance = Instances(fi_key)
        Select Case Err.Number
            Case 0
            Case 13
                If Instances.Exists(fi_key) Then
                    '~~ The apparently no longer existing instance is removed from the Dictionarys
                    Instances.Remove fi_key
                End If
                Set TestInstance = New fProcTest
                Instances.Add fi_key, TestInstance
            Case Else
                '~~ Unknown error!
                Err.Raise 1 + vbObjectError, ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
    End If

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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
    '~~ Expected result: Min values are set to their corresponding limit, max values are set to the width value
    WidthMin = 0
    WidthMax = 0
    HeightMin = 0
    HeightMax = 0
    mMsg.AssertWidthAndHeight WidthMin, WidthMax, HeightMin, HeightMax
    Debug.Assert WidthMin = Pnts(MSG_WIDTH_MIN_LIMIT_PERCENTAGE, "w")
    Debug.Assert WidthMax = WidthMin
    Debug.Assert HeightMin = Pnts(MSG_HEIGHT_MIN_LIMIT_PERCENTAGE, "h")
    Debug.Assert HeightMax = HeightMin


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
        .Top = 0
        .Left = 0
        .Show False
        
        For TestWidthLimit = iFrom To iTo Step iStep
            i = i + 1
            .Caption = PROC
            .frm.Width = TestWidthLimit + 3
            .frm.Left = 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .tbx.Left = 0
            .tbx.Top = 0
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
                .Top = 5
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
                       
            .frm.Top = .tbxTestAndResult.Top + .tbxTestAndResult.Height + 5
            
            '~~ The UserForm's height is adjusted to the resulting frame size
            fProcTest.Height = .frm.Top + .frm.Height + (fProcTest.Height - fProcTest.InsideHeight) + 5
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
        .Show False
        .Top = 0
        .Left = 0
        For i = iFrom To iTo Step iStep
            .Caption = PROC
            .frm.Left = 5
            .tbx.Left = 0
            .tbx.Top = 0
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
                .Top = 5
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
            
            .frm.Top = .tbxTestAndResult.Top + .tbxTestAndResult.Height + 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .Height = .frm.Top + .frm.Height + (.Height - .InsideHeight) + 5
            
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

Public Sub Test_DisplayWithWithoutFrames()
    Const PROC = "Test_DisplayWithWithoutFrames"
    
    Dim MsgForm     As fMsg
    Dim MsgTitle    As String
    
    MsgTitle = "With frames test"
    Set MsgForm = mMsg.MsgInstance(MsgTitle)
    
    MsgForm.DsplyFrmsWthBrdrsTestOnly = True
    
    mMsg.Box box_title:="With frames test" _
           , box_msg:="Message should be displayed with visible frames"

    mMsg.Box box_title:="With frames test" _
           , box_msg:="Message should be displayed with frames invible"
           
End Sub

Public Sub Test_SetupTitle()
    fProcTest.Show False
End Sub

Public Sub Test_TestInstance()
' ------------------------------------------------------------------------------
' Creates a number of instance of the UserForm named fProcTest and unloads them
' in the revers order. Application.Wait is used to allow the observation of the
' process.
' Note: The test shows that is not required to have a variable for the instance
'       object. It may however make sense in practise.
' ------------------------------------------------------------------------------
    Const INIT_TOP = 50
    Const INIT_LEFT = 50
    
    Dim i   As Long
    Dim key As String
    Dim obj As Object ' not required for the function but only to get the UserForm's name
    
    For i = 1 To 5
        key = "Instance-" & i
        '~~ Set obj ... will create the instance. However, this is not not required.
        '~~ It is just used to obtain the UserForms name
        Set obj = TestInstance(fi_key:=key)
        With TestInstance(fi_key:=key)
            .Height = 80
            .Width = 200
            .Caption = key & " of UserForm '" & obj.Name & "'"
            .Show Modeless
            .Top = INIT_TOP + (30 * i)
            .Left = INIT_LEFT + (30 * i)
        End With
        Application.Wait Now() + 0.000006
    Next i
    
    For i = 5 To 1 Step -1
        key = "Instance-" & i
        '~~ Unloading the instance this way has two advantages:
        '~~ 1. The instance is removed from the Dictionary
        '~~ 2. No error in case the instance no longer exists
        TestInstance fi_key:=key, fi_unload:=True
        Application.Wait Now() + 0.000006
    Next i
    
End Sub

Public Sub Test_SizingAndPositioning()

    Dim Instance1 As String
    Dim Instance2 As String
    Dim Instance3 As String
    Dim Instance4 As String
    Dim Instance5 As String
    
    
    Instance1 = "Test Sizing and Positioning: Top=0, Left=0, Width=" & PrcPnt(50, "w") & ", Height=" & PrcPnt(50, "h")
    With TestInstance(Instance1)
        .Top = 0
        .Left = 0
        .Width = Pnts(50, "w")
        .Height = Pnts(50, "h")
        .Caption = Instance1
        .Show Modeless
    End With
    
    Instance2 = "Test Sizing and Positioning: Top=0, Left=" & PrcPnt(50, "w") & ", Width=" & PrcPnt(50, "w") & ", Height=" & PrcPnt(50, "h")
    With TestInstance(Instance2)
        .Top = 0
        .Left = TestInstance(Instance1).Width
        .Width = Pnts(50, "w")
        .Height = Pnts(50, "h")
        .Caption = Instance2
        .Show Modeless
    End With
    
    Instance3 = "Test Sizing and Positioning: Top=" & PrcPnt(50, "h") & ", Left=0, Width=" & PrcPnt(50, "w") & ", Height=" & PrcPnt(50, "h")
    With TestInstance(Instance3)
        .Top = TestInstance(Instance1).Height
        .Left = 0
        .Width = Pnts(50, "w")
        .Height = Pnts(50, "h")
        .Caption = Instance3
        .Show Modeless
    End With
    
    Instance4 = "Test Sizing and Positioning: Top=" & PrcPnt(50, "h") & ", Left=" & PrcPnt(50, "w") & ", Width=" & PrcPnt(50, "w") & ", Height=" & PrcPnt(50, "h")
    With TestInstance(Instance4)
        .Top = TestInstance(Instance2).Height
        .Left = TestInstance(Instance3).Width
        .Width = Pnts(50, "w")
        .Height = Pnts(50, "h")
        .Caption = Instance4
        .Show Modeless
    End With
    
    Instance5 = "Test Sizing and Positioning: Top=0, Left=0, Width=" & PrcPnt(100, "w") & ", Height=" & PrcPnt(100, "h")
    With TestInstance(Instance5)
        .Top = 0
        .Left = 0
        .Width = Pnts(100, "w")
        .Height = Pnts(95, "h")
        .Caption = Instance5
        .Show Modeless
    End With
       
    TestInstance Instance1, True
    TestInstance Instance2, True
    TestInstance Instance3, True
    TestInstance Instance4, True
    TestInstance Instance5, True

End Sub

Private Function PrcPnt(ByVal pp_value As Single, _
                        ByVal pp_dimension As String) As String
    PrcPnt = mMsg.Prcnt(pp_value, pp_dimension) & "% (" & mMsg.Pnts(pp_value, "w") & "pt)"
End Function

