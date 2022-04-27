Attribute VB_Name = "mMsgTestProcs"
Option Explicit

' ------------------------------------------------------------------------------
' Standard Module mProcTest
'          Test of procedures - rather than fMsg/mMsg services/functions.
'
' ------------------------------------------------------------------------------
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mMsgTestServices." & s:  End Property

Public Function AdjustToVgrid(ByVal atvg_si As Single, _
                     Optional ByVal atvg_threshold As Single = 1.5, _
                     Optional ByVal atvg_grid As Single = 6) As Single
' -------------------------------------------------------------------------------
' Returns an integer which is a multiple of the grid value (stvg_grid) which
' defaults to 6, by considering a certain threshold (atvg_threshold) which
' defaults to 1.5.
' The function is used to vertically align form controls with the grid in order
' result vertically aligns a control in a userform to a grid value which ensures
' to have any text within the control correctly displayed in accordance with its
' font size. A certain threshold prevents an optically irritating large space to
' a control abovel. Examples:
'  7.5 < si >= 0   results to 6
' 13.5 < si >= 7.5 results in 12
' -------------------------------------------------------------------------------
    AdjustToVgrid = (Int((atvg_si - atvg_threshold) / atvg_grid) * atvg_grid) + atvg_grid
End Function

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

Private Function BttnsNo(ByVal v As Variant) As Long
    Select Case v
        Case vbYesNo, vbRetryCancel, vbResumeOk:    BttnsNo = 2
        Case vbAbortRetryIgnore, vbYesNoCancel:     BttnsNo = 3
        Case Else:                                  BttnsNo = 1
    End Select
End Function

Private Function Buttons(ParamArray bttns() As Variant) As Collection
' --------------------------------------------------------------------------
' Returns a collection if the items (bttns) provided. If an item is a
' Collection its items are included. When the number of buttons in a row
' exceeds 7 a vbLf is included to add a new row. When the number of rows is
' exieeded any subsequent items are ignored.
' --------------------------------------------------------------------------
    Const PROC          As String = "Buttons"
    
    On Error GoTo eh
    Static StackItems   As Collection
    Static QueueResult  As Collection
    Static cllResult    As Collection
    Static lBttnsInRow  As Long         ' buttons in a row counter (excludes break items)
    Static lBttns       As Long         ' total buttons in cllAdd
    Static lRows        As Long         ' button rows counter
    Static SubItems     As Long
    Static SubItemsDone As Long
    Dim cll             As Collection
    Dim v2              As Variant
    Dim i               As Long
    
    If cllResult Is Nothing Then
        Set StackItems = New Collection
        Set QueueResult = New Collection
        Set cllResult = New Collection
        lBttnsInRow = 0
        lBttns = 0
        lRows = 0
        SubItemsDone = 0
    End If
    If UBound(bttns) = -1 Then GoTo xt
    If UBound(bttns) = 0 Then
        '~~ Only one item
        Select Case TypeName(bttns(0))
            Case "Collection"
                Set cll = bttns(0)
                For i = cll.Count To 1 Step -1
                    StckPush StackItems, cll(i)
                Next i
            Case Else
                If bttns(0) = vbNullString Then Exit Function
                Select Case bttns(0)
                    Case vbLf, vbCr, vbCrLf
                        cllResult.Add bttns(0)
                        lBttnsInRow = 0
                        lRows = lRows + 1
                    Case Else
                        If lBttnsInRow + BttnsNo(bttns(0)) >= 7 Then
                            If lRows < 7 Then
                                cllResult.Add vbLf
                                lRows = lRows + 1
                                lBttnsInRow = 1
                            End If
                        End If
                        cllResult.Add bttns(0)
                        lBttnsInRow = lBttnsInRow + BttnsNo(bttns(0))
                End Select
        End Select
    Else
        '~~ More than one item in ParamArray
        For i = UBound(bttns) To 0 Step -1
            StckPush StackItems, bttns(i)
        Next i
    End If
    
    If Not StckIsEmpty(StackItems) Then
        Do
            If lRows >= 7 And lBttnsInRow >= 7 Then
                GoTo xt
            End If
            Qenqueue QueueResult, cllResult
            Buttons StckPop(StackItems)
            If StckIsEmpty(StackItems) Then Exit Do
        Loop
    End If

xt: If Not QisEmpty(QueueResult) Then
        Set cllResult = Qdequeue(QueueResult)
        Exit Function
    End If
    
    If StckIsEmpty(StackItems) And QisEmpty(QueueResult) Then
        Debug.Print "1. cllResult.Count: " & cllResult.Count
        Set Buttons = cllResult
        Set cllResult = Nothing
        Debug.Print "2. Buttons.Count: " & Buttons.Count
        Set QueueResult = Nothing
        Set StackItems = Nothing
    End If
    Exit Function
        
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

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

Private Function FormNew(ByVal uf_wb As Workbook, _
                         ByVal uf_name As String, _
                         ByVal uf_buttons As Variant) As UserForm
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "FormNew"
    
    On Error GoTo eh
    Dim MyUserForm          As VBComponent
    Dim NewCommandButton1   As MSForms.CommandButton
    Dim NewCommandButton2   As MSForms.CommandButton
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function IsForm(ByVal v As Object) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = v.Parent
    IsForm = Err.Number <> 0
End Function

Private Function IsFrameOrForm(ByVal v As Object) As Boolean
    IsFrameOrForm = TypeOf v Is MSForms.UserForm Or TypeOf v Is MSForms.Frame
End Function

Private Function IsValidMsgButtonsArg(ByVal v_arg As Variant) As Boolean
' -------------------------------------------------------------------------------------
' Returns TRUE when the buttons argument (v_arg) is valid. When v_arg is an Array,
' a Collection, or a Dictionary, TRUE is returned when all items are valid.
' -------------------------------------------------------------------------------------
    Dim i As Long
    Dim v As Variant
    
    Select Case VarType(v_arg)
        Case vbString, vbEmpty
            IsValidMsgButtonsArg = True
        Case Else
            Select Case True
                Case IsArray(v_arg), TypeName(v_arg) = "Collection", TypeName(v_arg) = "Dictionary"
                     For Each v In v_arg
                        If Not IsValidMsgButtonsArg(v) Then Exit Function
                     Next v
                    IsValidMsgButtonsArg = True
                Case IsNumeric(v_arg)
                    Select Case BttnsArgs(v_arg) ' The numeric buttons argument with all additional option 'unstripped'
                        Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                            IsValidMsgButtonsArg = True
                    End Select
            End Select
    End Select

End Function

Private Sub IsValidMsgButtonsArg_Test()
' -------------------------------------------------------------------------------------
' Test of the "IsValidMsgButtonsArg" function, !! the copy from mMsg !!
' -------------------------------------------------------------------------------------
    Dim ValidCollection         As New Collection
    Dim ValidDictionary         As New Dictionary
    Dim ValidArray(1 To 3)      As Variant
    Dim InValidCollection       As New Collection
    Dim InValidDictionary       As New Dictionary
    Dim InValidArray(1 To 3)    As Variant
    
    ValidArray(1) = vbOKCancel
    ValidArray(2) = "xxx"
    ValidArray(3) = "xxx,yyy"
    
    InValidArray(1) = vbOKCancel
    InValidArray(2) = 2377
    InValidArray(3) = "xxx,yyy"
    
    ValidCollection.Add vbOKCancel
    ValidCollection.Add "xxx"
    ValidCollection.Add "xxx,yyy"
    
    ValidDictionary.Add vbOKCancel, vbOKCancel
    ValidDictionary.Add "xxx", "xxx"
    ValidDictionary.Add "xxx,yyy", "xxx,yyy"
    
    Debug.Assert IsValidMsgButtonsArg(vbYesNo) = True
    Debug.Assert IsValidMsgButtonsArg("xxx") = True
    Debug.Assert IsValidMsgButtonsArg("xxx,yyy") = True
    Debug.Assert IsValidMsgButtonsArg("xxx") = True
    Debug.Assert IsValidMsgButtonsArg(ValidArray) = True
    Debug.Assert IsValidMsgButtonsArg(ValidCollection) = True
    Debug.Assert IsValidMsgButtonsArg(ValidDictionary) = True

    Debug.Assert IsValidMsgButtonsArg(InValidArray) = False

    Set ValidCollection = Nothing
    Set ValidDictionary = Nothing
End Sub

Private Sub Monitor_Test()
    Dim i       As Long
    Dim fMon    As fMsg
    Dim Text    As TypeMsgText
    Dim Title   As String
    Dim Footer  As TypeMsgText
    Dim Step    As TypeMsgText
    Dim Header  As TypeMsgText
    
    Title = "Test of the services:  MonitorHeader, MsgMonitor, and MonitorFooter"
    With Header
        .Text = "Process steps 1 to 10"
        .FontColor = rgbBlue
    End With
    With Footer
        .Text = "Process (steps 1 to 10) in progress, please hang on."
        .FontColor = rgbBlue
    End With
    Set fMon = mMsg.MsgInstance(Title)
    fMon.VisualizeForTest = wsTest.VisualizeForTest
    
    mMsg.MonitorHeader Title, Header
    mMsg.MonitorFooter Title, Footer
    
    For i = 1 To 20
        If i = 11 Then
            With Header
                .Text = "Process steps 11 to 20"
                .FontColor = rgbRed
            End With
            mMsg.MonitorHeader Title, Header
            With Footer
                .Text = "Process (steps 10 to 20) in progress, please hang on."
                .FontColor = rgbBlue
            End With
            mMsg.MonitorFooter Title, Footer
        End If
        DoEvents
        With Step
            .Text = Format(i, "00 ") & "Step"
            .MonoSpaced = True
        End With
        mMsg.Monitor mon_title:=Title _
                   , mon_text:=Step
        DoEvents
        Sleep 150
    Next i
    
    With Footer
        .Text = "Process finished! Close window."
        .FontColor = rgbDarkGreen
    End With
    mMsg.MonitorFooter Title, Footer

End Sub

Private Function PrcPnt(ByVal pp_value As Single, _
                        ByVal pp_dimension As String) As String
    PrcPnt = mMsg.Prcnt(pp_value, pp_dimension) & "% (" & mMsg.Pnts(pp_value, "w") & "pt)"
End Function

Private Function Qdequeue(ByRef qu As Collection) As Variant
    Const PROC = "DeQueue"
    
    On Error GoTo eh
    If qu Is Nothing Then GoTo xt
    If QisEmpty(qu) Then GoTo xt
    On Error Resume Next
    Set Qdequeue = qu(1)
    If Err.Number <> 0 _
    Then Qdequeue = qu(1)
    qu.Remove 1

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Private Sub Qenqueue(ByRef qu As Collection, ByVal qu_item As Variant)
    If qu Is Nothing Then Set qu = New Collection
    qu.Add qu_item
End Sub

Private Function QisEmpty(ByVal qu As Collection) As Boolean
    If Not qu Is Nothing _
    Then QisEmpty = qu.Count = 0 _
    Else QisEmpty = True
End Function

Private Function StckIsEmpty(ByVal stck As Collection) As Boolean
' ----------------------------------------------------------------------------
' Common Stack Empty check service. Returns True when either there is no stack
' (stck Is Nothing) or when the stack is empty (items count is 0).
' ----------------------------------------------------------------------------
    StckIsEmpty = stck Is Nothing
    If Not StckIsEmpty Then StckIsEmpty = stck.Count = 0
End Function

Private Function StckPop(ByVal stck As Collection) As Variant
' ----------------------------------------------------------------------------
' Common Stack Pop service. Returns the last item pushed on the stack (stck)
' and removes the item from the stack. When the stack (stck) is empty a
' vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "StckPop"
    
    On Error GoTo eh
    If StckIsEmpty(stck) Then GoTo xt
    
    On Error Resume Next
    Set StckPop = stck(stck.Count)
    If Err.Number <> 0 _
    Then StckPop = stck(stck.Count)
    stck.Remove stck.Count

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Function

Private Sub StckPush(ByRef stck As Collection, _
                     ByVal stck_item As Variant)
' ----------------------------------------------------------------------------
' Common Stack Push service. Pushes (adds) an item (stck_item) to the stack
' (stck). When the provided stack (stck) is Nothing the stack is created.
' ----------------------------------------------------------------------------
    Const PROC = "StckPush"
    
    On Error GoTo eh
    If stck Is Nothing Then Set stck = New Collection
    stck.Add stck_item

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case Else:  GoTo xt
    End Select
End Sub

Private Function TestInstance(ByVal fi_key As String, _
                     Optional ByVal fi_unload As Boolean = False) As fMsgProcTest
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fMsgProcTest which is definitely
' identified by anything uniqe for the instance (fi_key). This may be what
' becomes the title (property Caption) or even an object such like a
' Worksheet (if the instance is Worksheet specific). An already existing or
' new created instance is maintained in a static Dictionary with fi_key as
' the key and returned to the caller. When fi_unload is true only a possibly
' already existing Userform identified by fi_key is unloaded.
'
' Requires: Reference to the "Microsoft Scripting Runtime".
' Usage   : The fMsgProcTest has to be replaced by the name of the desired
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
        Set TestInstance = New fMsgProcTest
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
                Set TestInstance = New fMsgProcTest
                Instances.Add fi_key, TestInstance
            Case Else
                '~~ Unknown error!
                Err.Raise 1 + vbObjectError, ErrSrc(PROC), "Unknown/unrecognized error!"
        End Select
        On Error GoTo -1
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub Test_AdjustToVgrid()
    Dim i As Single
    
    For i = 5.8 To 8 Step 0.1
        Debug.Print Format(i, "00.0: ") & AdjustToVgrid(i)
    Next i
    Debug.Print Format(fMsgProcTest.tbxFactor.Object, "00.0: ") & AdjustToVgrid(i)
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
    With fMsgProcTest
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
            fMsgProcTest.Height = .frm.Top + .frm.Height + (fMsgProcTest.Height - fMsgProcTest.InsideHeight) + 5
            fMsgProcTest.Width = .frm.Left + .frm.Width + (fMsgProcTest.Width - fMsgProcTest.InsideWidth) + 5
            
            If TestWidthLimit <> iTo Then
                Select Case MsgBox(Title:="Continue? > Yes, Finish > No, Terminate? > Cancel", Buttons:=vbYesNoCancel, Prompt:=vbNullString)
                    Case vbYes
                    Case vbNo:                          Exit Sub
                    Case vbCancel: Unload fMsgProcTest: Exit Sub
                End Select
            Else
                Select Case MsgBox(Title:="Done? > Abort, Repeat? > Retry, Finish > Innore", Buttons:=vbAbortRetryIgnore, Prompt:=vbNullString)
                    Case vbAbort:   Unload fMsgProcTest:   Exit Sub
                    Case vbRetry:   Unload fMsgProcTest:   GoTo again
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
    With fMsgProcTest
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
                    Case vbCancel: Unload fMsgProcTest: Exit Sub
                End Select
            Else
                Select Case MsgBox(Title:="Done? > Abort, Repeat? > Retry, Finish > Ignore", Buttons:=vbAbortRetryIgnore, Prompt:=vbNullString)
                    Case vbAbort:   Unload fMsgProcTest:   Exit Sub
                    Case vbRetry:   Unload fMsgProcTest:   GoTo again
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
    MsgForm.VisualizeForTest = True
    mMsg.Box "Message should be displayed with visible frames", "With frames test"
    mMsg.Box "Message should be displayed with frames invisible", "With frames test"
           
End Sub

Private Sub Test_IsFrameOrForm()
    Debug.Assert IsFrameOrForm(fMsgProcTest.TextBox1) = False
    Debug.Assert IsFrameOrForm(fMsgProcTest.frm) = True
    Debug.Assert IsFrameOrForm(fMsgProcTest.frm.Parent) = True
    Debug.Assert IsForm(fMsgProcTest.frm) = False
    Debug.Assert IsForm(fMsgProcTest.frm.Parent) = True
End Sub

Public Sub Test_MultipleMessageInstances()
' ------------------------------------------------------------------------------
' Creates a number of instance of the UserForm named fMsgProcTest and unloads them
' in the revers order. Application.Wait is used to allow the observation of the
' process.
' Note: The test shows that is not required to have a variable for the instance
'       object. It may however make sense in practise.
' ------------------------------------------------------------------------------
    Const INIT_TOP = 50
    Const INIT_LEFT = 50
    
    Dim i   As Long
    Dim key As String
    Dim Obj As Object ' not required for the function but only to get the UserForm's name
    
    For i = 1 To 5
        key = "Instance-" & i
        '~~ Set obj ... will create the instance. However, this is not not required.
        '~~ It is just used to obtain the UserForms name
        Set Obj = TestInstance(fi_key:=key)
        With TestInstance(fi_key:=key)
            .Height = 80
            .Width = 200
            .Caption = key & " of UserForm '" & Obj.Name & "'"
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

Private Sub Test_OpenFile()
    mMsg.ShellRun "E:\Ablage\Excel VBA\DevAndTest\Common-VBA-Message-Service\ExecTrace.log", WIN_NORMAL
End Sub

Private Sub Test_OpenFile_No_Assoc()
    mMsg.ShellRun "E:\Ablage\Excel VBA\DevAndTest\Common-VBA-Message-Service\.gitattributes", WIN_NORMAL
End Sub

Private Sub Test_OpenHyperlink()
    mMsg.ShellRun "https://github.com/warbe-maker/Common-VBA-Message-Service", WIN_NORMAL
End Sub

Private Sub Test_Pass_TypeMsgText()
' ------------------------------------------------------------------------------
' Test of passing on any 'kind of' TypeMsgText to a UserForm and retrieving it
' again as a TypeMsgText.
' ------------------------------------------------------------------------------
    Const PROC = "Test_Pass_TypeMsgText"
    
    On Error GoTo eh
    Dim t As TypeMsgText
    Dim f As fMsgProcTest
    Dim k As KindOfText
    Dim i As Long
    
    Set f = TestInstance(fi_key:="Test-Title", fi_unload:=True)
    Set f = TestInstance(fi_key:="Test-Title")
    
    t.FontBold = True
    With f
        k = enMonHeader
        .Text(k) = t
        Debug.Assert .Text(k).FontBold = True
        Debug.Assert .Text(enMonFooter).FontBold = False
    
        k = enMonFooter
        .Text(k) = t
        Debug.Assert .Text(k).FontBold = True
    
        k = enMonStep
        .Text(k) = t
        Debug.Assert .Text(k).FontBold = True
    
        For i = 1 To 4
            k = enSectText
            .Text(k, i) = t
            Debug.Assert .Text(k, i).FontBold = True
        Next i
      
    End With
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_SetupTitle()
    fMsgProcTest.Show False
End Sub

Public Sub Test_SizingAndPositioning()

    Dim Instance1 As String
    Dim Instance2 As String
    Dim Instance3 As String
    Dim Instance4 As String
    Dim Instance5 As String
    Dim Title     As String
    Dim i         As Long
    
    For i = 1 To 7
        Title = RepeatStrng("Test Sizing and Positioning", i)
        With TestInstance(Title)
            .Top = 0
            .Left = 0
            .Height = Pnts(50, "h")
            .Caption = Title
            .Top = i * 35
            .Left = i * 10
            .Setup1_Title Title, 200, 800
            .Show False
        End With
    Next i

End Sub

Private Function RepeatStrng( _
                       ByVal rs_s As String, _
                       ByVal rs_n As Long) As String
' ----------------------------------------------------------------------------
' Returns the string (s) concatenated (n) times. VBA.String in not appropriate
' because it does not support leading and trailing spaces.
' ----------------------------------------------------------------------------
    Dim i   As Long
    For i = 1 To rs_n: RepeatStrng = RepeatStrng & rs_s:  Next i
End Function


