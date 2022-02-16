Attribute VB_Name = "mDemo"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Private vButtons As Collection

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mDemo." & s:  End Property

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
    Demo_Dsply_Service_1
    Demo_Dsply_Service_2
End Sub

Public Sub Demo_Box_Service()
    Const PROC          As String = "Demo_Box_service"
    Const BTTN_1        As String = "Button-1 caption"
    Const BTTN_2        As String = "Button-2 caption"
    Const BTTN_3        As String = "Button-3 caption"
    Const BTTN_4        As String = "Button-4 caption"
    Const DEMO_TITLE    As String = "Demonstration of the Box service"
    
    On Error GoTo eh
    Dim DemoMessage     As String
    
    DemoMessage = "The message : The ""Box"" service displays one string just like the VBA MsgBox. However, the monospaced" & vbLf & _
                  "              option allows a better layout for an indented text like this one for example. It should also be noted" & vbLf & _
                  "              that there is in fact no message width limit." & vbLf & _
                  "The buttons : 7 buttons in 7 rows are possible each with any caption string or a VBA MsgBox value. The latter may" & vbLf & _
                  "              result in more than one button, e.g. vbYesNoCancel." & vbLf & _
                  "The window  : When the message exceeds the specified maximum width a horizontal scroll-bar, when it exceeds" & vbLf & _
                  "              the specified maximum height a vertical scroll.bar is displayed  the message is displayed with a horizontal scroll-bar." & vbLf
    
    
    mMsg.Buttons vButtons, BTTN_1, BTTN_2, BTTN_3, BTTN_4, vbLf, vbYesNoCancel
    Select Case mMsg.Box( _
             box_title:=DEMO_TITLE _
           , box_msg:=DemoMessage _
           , box_monospaced:=True _
           , box_width_max:=50 _
           , box_buttons:=vButtons _
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

Public Sub Demo_Dsply_Service_1()
    Const width_max     As Long = 50
    Const MAX_HEIGHT    As Long = 60

    Dim sTitle          As String
    Dim cll             As New Collection
    Dim i, j            As Long
    Dim Message         As TypeMsg
   
    sTitle = "Usage demo: Full featured multiple choice message"
    With Message.Section(1)
        .Label.Text = "Demonstration overview:"
        .Label.FontColor = rgbBlue
        .Text.Text = "- Use of all 4 message sections" & vbLf _
                   & "- All sections with a label" & vbLf _
                   & "- One section monospaced exceeding the specified maximum message form width" & vbLf _
                   & "- Use of some of the 7x7 reply buttons in a 4-4-1 order" & vbLf _
                   & "- An an example for available text font options all labels in blue"
    End With
    With Message.Section(2)
        .Label.Text = "Unlimited message width!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section's text is mono-spaced (which is not word-wrapped) and the maximimum message form width" & vbLf _
                   & "for this demo has been specified " & width_max & "% of the screen width (the default would be 80%)" & vbLf _
                   & "the text is displayed with a horizontal scrollbar. There is no message size limit for the display despite the" & vbLf & vbLf _
                   & "limit of VBA for text strings  which is about 1GB!"
        .Text.MonoSpaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section lext has many lines (line breaks)" & vbLf _
                   & "the default word-wrapping for this proportional-spaced text" & vbLf _
                   & "has not the otherwise usuall effect. The message area thus" & vbLf _
                   & "exeeds the for this demo specified " & MAX_HEIGHT & "% of the screen size" & vbLf _
                   & "(defaults to 80%) it is displayed with a vertical scrollbar." & vbLf _
                   & "So even a proportional spaced text's size - which usually is word-wrapped -" & vbLf _
                   & "is only limited by the system's limit for a String which is abut 1GB !!!"
    End With
    With Message.Section(4)
        .Label.Text = "Great reply buttons flexibility:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 49 possible reply buttons (7 rows by 7 buttons). " _
                   & "It also shows that a reply button can have any caption text and the buttons can be " _
                   & "displayed in any order within the 7 x 7 limit. Of cource the VBA.MsgBox classic " _
                   & "vbOkOnly, vbYesNoCancel, etc. are also possible - even in a mixture." & vbLf & vbLf _
                   & "By the way: This demo ends only with the Ok button clicked and loops with all the ohter."
    End With
    '~~ Prepare the buttons collection
    For j = 1 To 2
        For i = 1 To 4
            cll.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
        cll.Add vbLf
    Next j
    cll.Add vbOKOnly ' The reply when clicked will be vbOK though
    
    While mMsg.Dsply(dsply_title:=sTitle _
                   , dsply_msg:=Message _
                   , dsply_buttons:=cll _
                   , dsply_height_max:=MAX_HEIGHT _
                   , dsply_width_max:=width_max _
                    ) <> vbOK
    Wend
    
End Sub

Public Sub Demo_Dsply_Service_2()
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
       
   mMsg.Buttons vButtons, vbAbortRetryIgnore, vbLf, B1, B2, B3, vbLf, B4, B5, B6, vbLf, B7
   vReturn = mMsg.Dsply(dsply_title:="Any title", _
                        dsply_msg:=Message, _
                        dsply_buttons:=vButtons _
                  )
   MsgBox "Button """ & ReplyString(vReturn) & """ had been clicked"
   
End Sub

Public Sub Demo_ErrMsg_Service()
    Const PROC = "Demo_ErrMsg_Service"
    
    On Error GoTo eh
    Dim i As Long
    i = i / 0
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Monitor_Service()
    Const PROC              As String = "Demo_Monitor_Service"
    Const MONITOR_HEADER    As String = " No. Status   Step"
    Const MONITOR_FOOTER    As String = "Process finished! Close this window"
    Const PROCESS_STEPS     As Long = 12
    
    On Error GoTo eh
    Dim i               As Long
    Dim lWait           As Long
    Dim MonitorTitle    As String
    Dim ProgressStep    As String
    
    MonitorTitle = "Demonstration of the monitoring of a process step by step"
    mMsg.MsgInstance MonitorTitle, fi_unload:=True ' Ensure there is no process monitoring with this title still displayed
        
    For i = 1 To PROCESS_STEPS
        '~~ Preparing a process step message string
        ProgressStep = mBasic.Align(i, 4, AlignRight, " ") & _
                   mBasic.Align("Passed", 8, AlignCentered, " ") & _
                   Repeat(repeat_n_times:=Int(((i - 1) / 10)) + 1, repeat_string:="  " & _
                   mBasic.Align(i, 2, AlignRight) & _
                   ".  Follow-Up line after " & _
                   Format(lWait, "0000") & _
                   " Milliseconds.")
        
        If i < PROCESS_STEPS Then
            '~~ Steps 1 to n - 1
            mMsg.Monitor mntr_title:=MonitorTitle _
                       , mntr_msg:=ProgressStep _
                       , mntr_msg_monospaced:=True _
                       , mntr_header:=MONITOR_HEADER
            
            '~~ Simmulation of a process
            lWait = 100 * i
            DoEvents
            Sleep 200
        
        Else
            '~~ The last step, separated in order to display the footer along with it
            mMsg.Monitor mntr_title:=MonitorTitle _
                       , mntr_msg:=ProgressStep _
                       , mntr_header:=MONITOR_HEADER _
                       , mntr_footer:=MONITOR_FOOTER
        End If
    Next i
    
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
' Universal error message display service. Displays a debugging option button
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section when information is concatenated with
' the error message by two vertical bars (||).
'
' May be copied as Private Function into any module. Considers the Common VBA
' Message Service and the Common VBA Error Services as optional components.
' When neither is installed the error message is displayed by the VBA.MsgBox.
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
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
' Note:  The above may seem to be a lot of code but will be a godsend in case
'        of an error!
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn tem into negative and in the error mesaage back into a positive
'          number.
' - ErrSrc To provide an unambigous procedure name - prefixed by the module name
'
' W. Rauschenberger Berlin, Nov 2021
'
' See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
' ------------------------------------------------------------------------------
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
    Dim C As Long
    Dim l As Long
    Dim i As Long

    l = Len(repeat_string)
    C = l * repeat_n_times
    s = Space$(C)

    For i = 1 To C Step l
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

