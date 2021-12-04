Attribute VB_Name = "mDemo"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Private vButtons As Collection

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
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume    ' go back to the error line
'               Case vbNo:  Resume Next     ' continue with line below the error line
'               Case Else:  Goto xt         ' clean exit (equivalent to Ok )
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
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    '~~ which also includes the installation of the mMsg component for the display of the error message.
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
    GoTo xt
#Else
#If MsgComp = 1 Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
    GoTo xt
#End If
#End If

    '~~ -------------------------------------------------------------------
    '~~ Neither the Common mMsg not the Commen mErH Component is installed.
    '~~ The error message is prepared for the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim errline     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then errline = Erl
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
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mDemo." & s:  End Property

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
    
    
    mMsg.Buttons vButtons, False, BTTN_1, BTTN_2, BTTN_3, BTTN_4, vbLf, vbYesNoCancel
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
        Case vbResumeNext:  Stop: Resume Next
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
       
   mMsg.Buttons vButtons, False, vbAbortRetryIgnore, vbLf, B1, B2, B3, vbLf, B4, B5, B6, vbLf, B7
   vReturn = Dsply(dsply_title:="Any title", _
                   dsply_msg:=Message, _
                   dsply_buttons:=vButtons _
                  )
   MsgBox "Button """ & mMsg.ReplyString(vReturn) & """ had been clicked"
   
End Sub

Public Sub Demo_ErrMsg_Service()
    Const PROC = "Demo_ErrMsg_Service"
    
    On Error GoTo eh
    Dim i As Long
    i = i / 0
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case vbResumeNext:  Stop: Resume Next
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
        Case vbResumeNext:  Stop: Resume Next
        Case Else:          GoTo xt
    End Select
End Sub

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

