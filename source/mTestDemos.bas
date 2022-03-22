Attribute VB_Name = "mTestDemos"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Private vButtons As New Collection
Private Message  As TypeMsg

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mMsgDemo." & s:  End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed 'Application' error number not conflicts with the
' number of a 'VB Runtime Error' or any other system error.
' - Returns a given positive 'Application Error' number (app_err_no) into a
'   negative by adding the system constant vbObjectError
' - Returns the original 'Application Error' number when called with a negative
'   error number.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Sub BoC(ByVal boc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(C)ode with id (boc_id) trace. Procedure to be copied as Private
' into any module potentially using the Common VBA Execution Trace Service. Has
' no effect when Conditional Compile Argument is 0 or not set at all.
' Note: The begin id (boc_id) has to be identical with the paired EoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC boc_id, s
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Public Sub Demos()
    Demo_Box
    Demo_Dsply_1
    Demo_Dsply_2
    Demo_Monitor
    Demo_Monitor_Instances
End Sub

Public Sub Demo_Box()
    Const PROC      As String = "Demo_Box_service"
    Const BTTN_1    As String = "Button-1 caption"
    Const BTTN_2    As String = "Button-2 caption"
    Const BTTN_3    As String = "Button-3 caption"
    Const BTTN_4    As String = "Button-4 caption"
    
    On Error GoTo eh
    Dim DemoMessage As String
    
    DemoMessage = "The        The ""Box"" service displays one string just like the VBA MsgBox." & vbLf & _
                  "message    However, the monospaced option allows a slightly better layout for an indented " & vbLf & _
                  "string     like this one. It should also be noted that there is in fact no message width limit." & vbLf & _
                  vbLf & _
                  "The        7 buttons in 7 rows are possible each with any caption string or a VBA.MsgBox value." & vbLf & _
                  "displayed  The latter may result in more than one button, e.g. vbYesNoCancel." & vbLf & _
                  "buttons" & vbLf & _
                  vbLf & _
                  "The        When the message exceeds the specified maximum width a horizontal scroll-bar," & vbLf & _
                  "message    when it exceeds the specified maximum height a vertical scroll-bar is displayed with" & vbLf & _
                  "window     the width exceeding section which may be a message section of the buttons section." & vbLf & vbLf & _
                  "Note       Press any button to terminate the dispplay!"
    
    
    Set vButtons = mMsg.Buttons(BTTN_1, BTTN_2, BTTN_3, BTTN_4, vbLf, vbYesNoCancel)
    Select Case mMsg.Box(Prompt:=DemoMessage _
                       , Buttons:=vButtons _
                       , Title:="Demonstration of the Box service" _
                       , box_monospaced:=True _
                       , box_width_max:=50 _
                       , box_button_default:=5 _
                        )
        Case BTTN_1:    MsgBox """" & BTTN_1 & """ pressed"
        Case BTTN_2:    MsgBox """" & BTTN_2 & """ pressed"
        Case BTTN_3:    MsgBox """" & BTTN_3 & """ pressed"
        Case BTTN_4:    MsgBox """" & BTTN_4 & """ pressed"
        Case vbYes:     MsgBox """ Yes"" pressed"
        Case vbNo:      MsgBox """No"" pressed"
        Case vbCancel:  MsgBox """Cancel"" pressed"
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Dsply_1()
    Const width_max     As Long = 35
    Const MAX_HEIGHT    As Long = 50

    Dim sTitle          As String
    Dim cll             As New Collection
    Dim i, j            As Long
    Dim Message         As TypeMsg
   
    sTitle = "Usage demo: Full featured multiple choice message"
    With Message.Section(1)
        .Label.Text = "Service features used by this displayed message:"
        .Label.FontColor = rgbBlue
        .Text.Text = "All 4 message sections, and all with a label, monospaced option for the second section, " _
                   & "some of the 7 x 7 reply buttons in a 4-4-1 order, font color option for all labels."
    End With
    With Message.Section(2)
        .Label.Text = "Demonstration of the unlimited message width:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section's text is mono-spaced (which by definition is not word-wrapped)" & vbLf _
                   & "the message width is determined by:" & vbLf _
                   & "a) the for this demo specified maximum width of " & width_max & "% of the screen size" & vbLf _
                   & "   (defaults to 80% when not specified)" & vbLf _
                   & "b) the longest line of this section" & vbLf _
                   & "Because the text exeeds the specified maximum message width, a horizontal scroll-bar is displayed." & vbLf _
                   & "Due to this feature there is no message size limit other than the sytem's limit which for a string is about 1GB !!!!"
        .Text.MonoSpaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height (not the fact with this message):"
        .Label.FontColor = rgbBlue
        .Text.Text = "As with the message width, the message height is unlimited. When the maximum height (explicitely specified or the default) " _
                   & "is exceeded a vertical scroll-bar is displayed. Due to this feature there is no message size limit other than the sytem's " _
                   & "limit which for a string is about 1GB !!!!"
    End With
    With Message.Section(4)
        .Label.Text = "Flexibility regarding the displayed reply buttons:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 7 x 7 = 49 possible reply buttons which may have any caption text " _
                   & "including the classic VBA.MsgBox values (vbOkOnly, vbYesNoCancel, etc.) - even in a mixture." & vbLf & vbLf _
                   & "!! This demo ends only with the Ok button and loops with any other."
    End With
    '~~ Prepare the buttons collection
    mMsg.Buttons cll, vbOKOnly, vbLf ' The reply when clicked will be vbOK though
    For j = 1 To 2
        For i = 1 To 4
            cll.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
        If j < 2 Then cll.Add vbLf
    Next j
    
    While mMsg.Dsply(dsply_title:=sTitle _
                   , dsply_msg:=Message _
                   , dsply_buttons:=cll _
                   , dsply_height_max:=MAX_HEIGHT _
                   , dsply_width_max:=width_max _
                    ) <> vbOK
    Wend
    
End Sub

Public Sub Demo_Dsply_2()
' ---------------------------------------------------------
' Displays a message with 3 sections, each with a label and
' 7 reply buttons ordered in rows 3-3-1
' ---------------------------------------------------------
    Const B1 = "Caption Button 1"
    Const B2 = "Caption Button 2"
    Const B3 = "Caption Button 3"
    Const B4 = "Caption Button 4"
    Const B5 = "Caption Button 5"
    Const B6 = "Caption Button 6"
    Const B7 = "Caption Button 7"
    
    Dim vReturn As Variant
    Dim Message As TypeMsg
    
    ' Preparing the message
    With Message.Section(1)
        .Label.Text = "Any section-1 label (bold, blue)"
        .Label.FontBold = True
        .Label.FontColor = rgbBlue
        .Text.Text = "This is a section-1 text (darkgreen)"
        .Text.FontColor = rgbDarkGreen
    End With
    With Message.Section(2)
        .Label.Text = "Any section-2 label (bold, blue)"
        .Label.FontBold = True
        .Label.FontColor = rgbBlue
        With .Text
            .Text = "This is a section-2 text (bold, italic, red, monospaced, font-size=10)"
            .FontBold = True
            .FontItalic = True
            .FontColor = rgbRed
            .MonoSpaced = True ' Just to demonstrate
            .FontSize = 10
        End With
    End With
    With Message.Section(3)
        .Text.Text = "Any section-3 text (without a label)"
   End With
       
   Set vButtons = mMsg.Buttons(vbAbortRetryIgnore, vbLf, B1, B2, B3, vbLf, B4, B5, B6, vbLf, B7)
   vReturn = mMsg.Dsply(dsply_title:="Any title", _
                        dsply_msg:=Message, _
                        dsply_buttons:=vButtons _
                  )
   MsgBox "Button """ & ReplyString(vReturn) & """ had been clicked"
   
End Sub

Public Sub Demo_ErrMsg()
    Const PROC = "Demo_ErrMsg"
    
    On Error GoTo eh
    Dim i As Long
    i = i / 0
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Monitor()
    Const PROC          As String = "Demo_Monitor"
    Const PROCESS_STEPS As Long = 12

    On Error GoTo eh
    Dim Header          As TypeMsgText
    Dim Footer          As TypeMsgText
    Dim i               As Long
    Dim lWait           As Long
    Dim Title           As String
    Dim Step            As TypeMsgText
    
    With Header
        .Text = " No. Status   Step"
        .MonoSpaced = True
        .FontColor = rgbBlue
    End With
    
    With Footer
        .Text = "Process in progress! Please wait."
        .FontBold = True
    End With
    
    Title = "Demonstration of the monitoring of a process step by step"
    mMsg.MsgInstance Title, fi_unload:=True ' Ensure there is no process monitoring with this title still displayed
        
    For i = 1 To PROCESS_STEPS
        '~~ Preparing a process step message string
        With Step
            .Text = mBasic.Align(i, 4, AlignRight, " ") & _
                            mBasic.Align("Passed", 8, AlignCentered, " ") & _
                            Repeat(repeat_n_times:=Int(((i - 1) / 10)) + 1, repeat_string:="  " & _
                            mBasic.Align(i, 2, AlignRight) & _
                            ".  Follow-Up line after " & _
                            Format(lWait, "0000") & _
                            " Milliseconds.")
            .MonoSpaced = True
        End With
        
        mMsg.Monitor mon_title:=Title _
                   , mon_header:=Header _
                   , mon_step:=Step _
                   , mon_footer:=Footer

        '~~ Simmulation of a process
        lWait = 100 * i
        DoEvents
        Sleep 200
        
    Next i
    
    Step.Text = vbNullString
    Footer.Text = "Process finished! Close this window"
    mMsg.Monitor mon_title:=Title _
               , mon_header:=Header _
               , mon_step:=Step _
               , mon_footer:=Footer
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Private Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
'
' - Displays a debugging option button when the Conditional Compile Argument
'   'Debugging = 1'
' - Displays an optional additional "About the error:" section when a string is
'   concatenated with the error message by two vertical bars (||)
' - Invokes mErH.ErrMsg when the Conditional Compile Argument ErHComp = !
' - Invokes mMsg.ErrMsg when the Conditional Compile Argument MsgComp = ! (and
'   the mErH module is not installed / MsgComp not set)
' - Displays the error message by means of VBA.MsgBox when neither of the two
'   components is installed
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn them into negative and in the error message back into a
'          positive number.
' - ErrSrc To provide an unambiguous procedure name by prefixing is with the
'          module name.
'
' See:
' https://github.com/warbe-maker/Common-VBA-Error-Services
'
' W. Rauschenberger Berlin, Feb 2022
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
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
    
    '~~ Consider extra information is provided with the error description
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
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
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

Private Sub EoC(ByVal eoc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(C)ode id (eoc_id) trace. Procedure to be copied as Private into
' any module potentially using the Common VBA Execution Trace Service. Has no
' effect when the Conditional Compile Argument is 0 or not set at all.
' Note: The end id (eoc_id) has to be identical with the paired BoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC eoc_id, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub


Private Function ReplyString(ByVal vReply As Variant) As String
' ------------------------------------------------------------------------------
' Returns the Dsply or Box return value as string. An invalid value is ignored.
' Only used with these demonstation examples.
' ------------------------------------------------------------------------------

    If VarType(vReply) = vbString Then
        ReplyString = vReply
    Else
        Select Case vReply
            Case vbAbort:       ReplyString = "Abort"
            Case vbCancel:      ReplyString = "Cancel"
            Case vbIgnore:      ReplyString = "Ignore"
            Case vbNo:          ReplyString = "No"
            Case vbOK:          ReplyString = "Ok"
            Case vbRetry:       ReplyString = "Retry"
            Case vbYes:         ReplyString = "Yes"
            Case vbResume:      ReplyString = "Resume Error Line"
        End Select
    End If
    
End Function

Private Sub Demo_Monitor_Instances()
' ------------------------------------------------------------------------------
' - uses the mMsg.Monitor service
' - displays 5 monitor instances
' - updates the text in each with up to five lines
' - removes them in reverse order.
' ------------------------------------------------------------------------------
    Const PROC = "Demo_Monotor_Instances"
    Const INIT_TOP  As Single = 100
    Const INIT_LEFT As Single = 50
    Const OFFSET_H  As Single = 80
    Const OFFSET_V  As Single = 20
    Const T_WAIT    As Single = 0.000003
    
    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim MsgForm     As fMsg
    Dim sTitle      As String
    Dim Header      As TypeMsgText
    Dim Step        As TypeMsgText
    Dim Footer        As TypeMsgText
    
    j = 1
    For i = 1 To 5
        '~~ Establish 5 monitoring instances
        '~~ Note that the instances are identified by their different titles
        sTitle = "Instance-" & i
        Step.Text = "Process step " & j
        mMsg.Monitor mon_title:=sTitle _
                   , mon_header:=Header _
                   , mon_step:=Step _
                   , mon_footer:=Footer _
                   , mon_width_min:=15
        With mMsg.MsgInstance(sTitle)
            .Top = INIT_TOP + OFFSET_V * (i - 1)
            .Left = INIT_LEFT + OFFSET_H * (i - 1)
        End With
        Application.Wait Now() + T_WAIT
    Next i
    
    For j = 2 To 5
        '~~ Display in each of the instances an additional progress message
        For i = 1 To 5
            '~~ Go through all instances and add a message line
            sTitle = "Instance-" & i
            Step.Text = "Process step " & j
            mMsg.Monitor mon_title:=sTitle _
                   , mon_header:=Header _
                   , mon_step:=Step _
                   , mon_footer:=Footer
            Application.Wait Now() + T_WAIT
        Next i
    Next j
    
    For i = 5 To 1 Step -1
        '~~ Unload the instances in reverse order
        mMsg.MsgInstance fi_key:="Instance-" & i, fi_unload:=True
        Application.Wait Now() + (T_WAIT * 2)
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

