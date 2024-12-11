Attribute VB_Name = "mMsgTestProcs"
Option Explicit
' ------------------------------------------------------------------------------
' Standard Module mMsgTestProcs: Test of basic procedures.
' ==============================
'
' ------------------------------------------------------------------------------
Private Const DFLT_SECT_TEXT_PROP   As String = ">Lorem ipsum dolor sit amet, consectetur adipiscing elit, " & _
                                                "sed do eiusmod tempor incididunt ut labore et dolore magna " & _
                                                "aliqua. Ut enim ad minim veniam, quis nostrud exercitation " & _
                                                "ullamco laboris nisi ut aliquip ex ea commodo consequat. " & _
                                                "Duis aute irure dolor in reprehenderit in voluptate velit " & _
                                                "esse cillum dolore eu fugiat nulla pariatur. Excepteur sint " & _
                                                "occaecat cupidatat non proident, sunt in culpa qui officia " & _
                                                "deserunt mollit anim id est laborum.<"
Private Const DFLT_SECT_TEXT_MONO   As String = ">Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed " & _
                                                "do eiusmod tempor incididunt ut labore et dolore magna aliqua." & vbLf & _
                                                "Ut enim ad minim veniam, quis nostrud exercitation " & _
                                                "ullamco laboris nisi ut aliquip ex ea commodo consequat." & vbLf & _
                                                "Duis aute irure dolor in reprehenderit in voluptate velit " & _
                                                "esse cillum dolore eu fugiat nulla pariatur." & vbLf & _
                                                "Excepteur sint occaecat cupidatat non proident, sunt in culpa " & _
                                                "qui officia deserunt mollit anim id est laborum.<"

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mMsgTestServices." & s:  End Property

Private Function AdjustToVgrid(ByVal atvg_si As Single, _
                      Optional ByVal atvg_threshold As Single = 1.5, _
                      Optional ByVal atvg_grid As Single = 6) As Single
' -------------------------------------------------------------------------------
' Returns an integer which is a multiple of the grid value (stvg_grid) which
' defaults to 6, by considering a certain threshold (atvg_threshold) which
' defaults to 1.5.
' The function is used to vertically align form controls with the grid in order
' result vertically aligns a control in a userform to a grid value which ensures
' to have any text within the control correctly displayed in accordance with its
' Font size. A certain threshold prevents an optically irritating large space to
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

Private Function Buttons(ParamArray Bttns() As Variant) As Collection
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
    Static SubItemsDone As Long
    Dim cll             As Collection
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
    If UBound(Bttns) = -1 Then GoTo xt
    If UBound(Bttns) = 0 Then
        '~~ Only one item
        Select Case TypeName(Bttns(0))
            Case "Collection"
                Set cll = Bttns(0)
                For i = cll.Count To 1 Step -1
                    StckPush StackItems, cll(i)
                Next i
            Case Else
                If Bttns(0) = vbNullString Then Exit Function
                Select Case Bttns(0)
                    Case vbLf, vbCr, vbCrLf
                        cllResult.Add Bttns(0)
                        lBttnsInRow = 0
                        lRows = lRows + 1
                    Case Else
                        If lBttnsInRow + BttnsNo(Bttns(0)) >= 7 Then
                            If lRows < 7 Then
                                cllResult.Add vbLf
                                lRows = lRows + 1
                                lBttnsInRow = 1
                            End If
                        End If
                        cllResult.Add Bttns(0)
                        lBttnsInRow = lBttnsInRow + BttnsNo(Bttns(0))
                End Select
        End Select
    Else
        '~~ More than one item in ParamArray
        For i = UBound(Bttns) To 0 Step -1
            StckPush StackItems, Bttns(i)
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

Private Function FormNew(ByVal uf_wb As Workbook, _
                         ByVal uf_name As String, _
                         ByVal uf_buttons As Variant) As UserForm
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "FormNew"
    
    On Error GoTo eh
    Dim NewCommandButton1   As MSForms.CommandButton
    Dim NewCommandButton2   As MSForms.CommandButton
    Dim x                   As Long
    Dim cmp                 As VBComponent
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
        x = .CountOfLines
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

Private Function IsValidMsgButtonsArg(ByVal v_arg As Variant) As Boolean
' -------------------------------------------------------------------------------------
' Returns TRUE when the buttons argument (v_arg) is valid. When v_arg is an Array,
' a Collection, or a Dictionary, TRUE is returned when all items are valid.
' -------------------------------------------------------------------------------------
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
                    Select Case BttnArg(v_arg) ' The numeric buttons argument with all additional option 'unstripped'
                        Case vbOKOnly, vbOKCancel, vbYesNo, vbRetryCancel, vbYesNoCancel, vbAbortRetryIgnore, vbYesNo, vbResumeOk
                            IsValidMsgButtonsArg = True
                    End Select
            End Select
    End Select

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

Private Function RepeatStrng(ByVal rs_s As String, _
                             ByVal rs_n As Long) As String
' ----------------------------------------------------------------------------
' Returns the string (s) concatenated (n) times. VBA.String in not appropriate
' because it does not support leading and trailing spaces.
' ----------------------------------------------------------------------------
    Dim i   As Long
    For i = 1 To rs_n: RepeatStrng = RepeatStrng & rs_s:  Next i
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

Private Function TestInstance(ByVal t_key As String, _
                     Optional ByVal t_unload As Boolean = False) As fMsgProcTest
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fMsgProcTest which is definitely
' identified by anything uniqe for the instance (t_key). This may be what
' becomes the title (property Caption) or even an object such like a
' Worksheet (if the instance is Worksheet specific). An already existing or
' new created instance is maintained in a static Dictionary with t_key as
' the key and returned to the caller. When t_unload is true only a possibly
' already existing Userform identified by t_key is unloaded.
'
' Requires: Reference to the "Microsoft Scripting Runtime".
' Usage   : The fMsgProcTest has to be replaced by the name of the desired
'           UserForm
' -------------------------------------------------------------------------
    Const PROC = "TestInstance"
    
    On Error GoTo eh
    Static Instances As Dictionary    ' Collection of (possibly still)  active form instances
    
    If Instances Is Nothing Then Set Instances = New Dictionary
    
    If t_unload Then
        If Instances.Exists(t_key) Then
            On Error Resume Next
            Unload Instances(t_key) ' The instance may be already unloaded
            Instances.Remove t_key
        End If
        Exit Function
    End If
    
    If Not Instances.Exists(t_key) Then
        '~~ There is no evidence of an already existing instance
        Set TestInstance = New fMsgProcTest
        Instances.Add t_key, TestInstance
    Else
        '~~ An instance identified by t_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set TestInstance = Instances(t_key)
        Select Case Err.Number
            Case 0
            Case 13
                If Instances.Exists(t_key) Then
                    '~~ The apparently no longer existing instance is removed from the Dictionarys
                    Instances.Remove t_key
                End If
                Set TestInstance = New fMsgProcTest
                Instances.Add t_key, TestInstance
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

Public Sub Test_00_1_Guide()

    Dim sTitle As String
    Dim sMsg As String
    
    sTitle = "Guidance title"
    sMsg = "This is a guidance message. This is a guidance message. This is a guidance message. This is a guidance message. This is a guidance message. This is a guidance message. This is a guidance message. This is a guidance message."
    With New clsTestAid
        .TestId = "00-1.1"
        .Guide sMsg
        .GuideDone
        .TestId = "00-1.2"
        .Guide sMsg
        .GuideDone
        .GuideUnload
    End With
    
End Sub

Public Sub Test_01_1_AdjustToVgrid()
    Debug.Assert AdjustToVgrid(7.4) = 6
    Debug.Assert AdjustToVgrid(7.5) = 12
End Sub

Public Sub Test_01_2_AutoSizeTextBox_Width_Limited()
    Const PROC = "Test_01_2_AutoSizeTextBox_Width_Limited"
    
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
                           , as_text:="For this test the width is limited to " & mMsg.ValueAsPercentage(wsTest.FormWidthMax, mMsg.enDsplyDimensionWidth) & ". " & _
                                      "The height is determined at first by the height resulting from the AutoSize." _
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

Public Sub Test_01_3_AutoSizeTextBox_Width_Unlimited()
    Const PROC = "Test_01_3_AutoSizeTextBox_Width_Unlimited"
    
    Dim i               As Long
    Dim iFrom           As Long
    Dim iStep           As Long
    Dim iTo             As Long
    Dim TestAppend      As Boolean
    Dim TestHeightMin   As Single
    Dim TestWidthLimit  As Single
    
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

Public Sub Test_01_4_Pass_udtMsgText()
' ------------------------------------------------------------------------------
' Test of passing on any 'kind of' udtMsgText to a UserForm and retrieving it
' again as a udtMsgText.
' ------------------------------------------------------------------------------
    Const PROC = "Test_01_4_Pass_udtMsgText"
    
    On Error GoTo eh
    Dim t As udtMsgText
    Dim f As fMsgProcTest
    Dim k As KindOfText
    Dim i As Long
    
    Set f = TestInstance(t_key:="Test-Title", t_unload:=True)
    Set f = TestInstance(t_key:="Test-Title")
    
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

Public Function Test_02_1_Single_Section_PropSpaced() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_1_Single_Section_PropSpaced"
        
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-1", PROC
    With mMsgTest.udtMessage.Section(1)
        .Text.Text = DFLT_SECT_TEXT_PROP
    End With
    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=vbNullString _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
           
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_02_2_Single_Section_MonoSpaced_With_Label() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_2_Single_Section_MonoSpaced_With_Label"
        
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-2", PROC
    With mMsgTest.udtMessage.Section(1)
        .Label.Text = "Path:"
        .Label.FontColor = rgbBlue
        .Text.Text = "mMsgTestServices.Test_12_mMsg_ErrMsg_AppErr_5: Application Error 5" ' DFLT_SECT_TEXT_MONO
        .Text.MonoSpaced = True
        .Text.FontSize = 9
    End With
    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=mMsgTest.BttnsBasic _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
           
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_02_3_Single_Section_MonoSpaced_No_Label() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_3_Single_Section_MonoSpaced_No_Label"
        
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-3", PROC

    With mMsgTest.udtMessage.Section(1).Text
        .MonoSpaced = True
        .FontSize = 9
        .Text = "Open a folder      : C:\TEMP\              " & vbLf & _
                "Call the eMail app : mailto:xxxxx@gmail.com" & vbLf & _
                "Open a url/link    : http://......         " & vbLf & _
                "Open a file        : C:\TEMP\TestThis          (opens a dialog for the selection of the app)" & vbLf & _
                "Open an application: x:\my\workbooks\this.xlsb (opens Excel)"
    End With

    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=mMsgTest.BttnsBasic _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
           
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_02_5_Single_Section_Label_Only() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_5_Single_Section_Label_Only"
        
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-5", PROC

    With mMsgTest.udtMessage.Section(1)
        .Label.Text = "Label-only section (no section text specified). The Label spans the full message " & _
                                "width and may even be multi-lined. Any Label position specs are ignored."
    End With
    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=mMsgTest.BttnsBasic _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
           
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_02_4_Single_Section_MonoSpaced_With_VH_Scroll()
' ------------------------------------------------------------------------------
' With the vertical scroll bar applied the horizontal scroll-bar of the section
' is replaced by one for the message area.
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_4_Single_Section_MonoSpaced_With_VH_Scroll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-4", PROC
            
    With mMsgTest.udtMessage.Section(1).Text
        .Text = "Text only section! " & DFLT_SECT_TEXT_MONO
        .MonoSpaced = True
    End With
      
    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=mMsgTest.BttnsBasic _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
             
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Test_02_6_Single_Section_MonoSpaced_With_Label_And_VH_Scroll()
' ------------------------------------------------------------------------------
' With the vertical scroll bar applied the horizontal scroll-bar of the section
' is replaced by one for the message area.
' ------------------------------------------------------------------------------
    Const PROC = "Test_02_6_Single_Section_MonoSpaced_With_Label_And_VH_Scroll"
    
    On Error GoTo eh
    
    mBasic.BoP ErrSrc(PROC)
    mMsgTest.InitializeTest "02-6", PROC
            
    With mMsgTest.udtMessage.Section(1)
        .Label.Text = "Test-Label:"
        .Label.FontColor = rgbGreen

        .Text.Text = DFLT_SECT_TEXT_MONO & vbLf & _
                     DFLT_SECT_TEXT_MONO & vbLf & _
                     DFLT_SECT_TEXT_MONO & vbLf & _
                     DFLT_SECT_TEXT_MONO
        .Text.MonoSpaced = True
    End With
      
    mMsg.Dsply dsply_title:=TestProcName _
             , dsply_msg:=mMsgTest.udtMessage _
             , dsply_buttons:=mMsgTest.BttnsBasic _
             , dsply_buttons_app_run:=mMsgTest.BttnsAppRunArgs _
             , dsply_Label_spec:=wsTest.MsgLabelPosSpec _
             , dsply_width_min:=wsTest.FormWidthMin _
             , dsply_width_max:=wsTest.FormWidthMax _
             , dsply_height_max:=wsTest.FormHeightMax _
             , dsply_modeless:=mMsgTest.MODE_LESS
             
xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub Test_Regression()

'    mMsgTestProcs.Test_01_1_AdjustToVgrid
'    mMsgTestProcs.Test_01_2_AutoSizeTextBox_Width_Limited
'    mMsgTestProcs.Test_01_3_AutoSizeTextBox_Width_Unlimited
'    mMsgTestProcs.Test_01_4_Pass_udtMsgText
    mMsgTestProcs.Test_02_1_Single_Section_PropSpaced
    mMsgTestProcs.Test_02_2_Single_Section_MonoSpaced_With_Label
    mMsgTestProcs.Test_02_3_Single_Section_MonoSpaced_No_Label
    mMsgTestProcs.Test_02_4_Single_Section_MonoSpaced_With_VH_Scroll
    mMsgTestProcs.Test_02_5_Single_Section_Label_Only
End Sub

