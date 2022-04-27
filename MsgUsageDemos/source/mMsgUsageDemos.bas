Attribute VB_Name = "mMsgUsageDemos"
Option Explicit
' ----------------------------------------------------------------------------
' Standard module mMsgUsageDemo Demonstrates the usage of the mMsg services
'                               - Box
'                               - Dsply
'                               - ErrMsg
'                               - Monitor
'                               - MsgInstance
' Just for demonstration!
' -----------------------
' Not only the process monitor window but also the mMsg.Box and the mMsg.Dsply
' window is displayed  m o d e l e s s  at an individual position on screen.
'
' Requires: - Installation/import of the Common Components fMsg, mMsg
'           - Reference to the "Microsoft Scripting Runtime"
'
' W. Rauschenberger, Berlin April 2022
' ----------------------------------------------------------------------------
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Private Message  As TypeMsg

Public Sub Demos()
    Demo_Service_mMsg_Box
    Sleep 1000
    Demo_Service_mMsg_Dsply
    Sleep 1000
    Demo_Service_mMsg_Monitor
    Sleep 1000
    Demo_Service_mMsg_MsgInstance
End Sub

Public Sub Demo_Service_mMsg_Box()
' ----------------------------------------------------------------------------
' Demonstration of the mMsg.Box service. The demonstration uses the mMsg.Bttns
' service to compile the to-be-displayed buttons. Buttons with a user defined
' caption string are declared as constants to allow the selection of the
' returned clicked button.
' ----------------------------------------------------------------------------
    Const PROC = "Demo_Service_mMsg_Box"
    Const BTN1 = "Button-1 caption"
    Const BTN2 = "Button-2 caption"
    Const BTN3 = "Button-3 caption"
    Const BTN4 = "Button-4 caption"
    
    On Error GoTo eh
    Dim Message As String
    
    Message = "The        The ""Box"" service displays one string just like the VBA MsgBox." & vbLf & _
              "displayed  However, the monospaced option allows a slightly better layout for structured message " & vbLf & _
              "message    like this one. It should also be noted that the message string lenght is limited only " & vbLf & _
              "string     by Windows which is about 1GB." & vbLf & _
              vbLf & _
              "The        7 buttons in 7 rows are possible each with any caption string in any combination " & vbLf & _
              "displayed  with a VBA.MsgBox value." & vbLf & _
              "buttons" & vbLf & _
              vbLf & _
              "The        When the message exceeds the specified maximum width a horizontal scroll-bar," & vbLf & _
              "message    when it exceeds the specified maximum height a vertical scroll-bar is displayed with" & vbLf & _
              "window     the width exceeding section which may be a message section of the buttons section." & vbLf & vbLf & vbLf & _
              "Press any button to close this message!"
        
    Select Case mMsg.Box(Prompt:=Message _
                       , Buttons:=mMsg.Buttons(BTN1, BTN2, BTN3, BTN4, vbLf, vbYesNoCancel) _
                       , Title:="Demonstration of the mMsg.Box service" _
                       , box_monospaced:=True _
                       , box_width_max:=50 _
                       , box_modeless:=True _
                       , box_pos:="100;20")
'        Case BTN1:      MsgBox """" & BTN1 & """ pressed"
'        Case BTN2:      MsgBox """" & BTN2 & """ pressed"
'        Case BTN3:      MsgBox """" & BTN3 & """ pressed"
'        Case BTN4:      MsgBox """" & BTN4 & """ pressed"
'        Case vbYes:     MsgBox """ Yes"" pressed"
'        Case vbNo:      MsgBox """No"" pressed"
'        Case vbCancel:  MsgBox """Cancel"" pressed"
    End Select

xt: Exit Sub

eh: Select Case mMsg.ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Service_mMsg_Dsply()
' ----------------------------------------------------------------------------
' Demonstration of the mMsg.Dsply service, displaying pretty much the same
' message as with the mMsg.Box service. The demonstration uses the mMsg.Bttns
' service to compile the to-be-displayed buttons. Buttons with a user defined
' caption string are declared as constants to allow the selection of the
' returned clicked button.
' ----------------------------------------------------------------------------
    Const PROC = "Demo_Service_mMsg_Dsply"
    Const BTN1 = "Button-1 caption"
    Const BTN2 = "Button-2 caption"
    Const BTN3 = "Button-3 caption"
    Const BTN4 = "Button-4 caption"
    
    On Error GoTo eh
    Dim Message As TypeMsg
    
    With Message
        With .Section(1)
            .Label.Text = "The displayed message:"
            .Label.FontColor = rgbBlue
            .Text.Text = "The ""mMsg.Dsply"" service allows a much better designed message than the VBA.MsgBox and the mMsg.Box service. " & _
                         "The monospaced option is available for any of the 4 possible message sections (not used with this message though). " & _
                         "It should also be noted that total message has in fact no limit (4 strings, each with about 1GB (Windows' string length limit)."
        End With
        With .Section(2)
            .Label.Text = "The displayed buttons:"
            .Label.FontColor = rgbBlue
            .Text.Text = "7 buttons in 7 rows are possible each with any caption string in any combination displayed with a VBA.MsgBox value."
        End With
        With .Section(3)
            .Label.Text = "The message window:"
            .Label.FontColor = rgbBlue
            .Text.Text = "The message window's width and height adjusts with the displayed message title, message and buttons. When the " & _
                         "maximum width and/or height is exceeded (defaults to 85% of the screen) scroll-bars are displayed."
        End With
        With Message.Section(4).Text
            .Text = "Press any button to close this message!"
            .MonoSpaced = True
            .FontColor = rgbRed
        End With
    End With
    
    Select Case mMsg.Dsply(dsply_msg:=Message _
                         , dsply_buttons:=mMsg.Buttons(BTN1, BTN2, BTN3, BTN4, vbLf, vbYesNoCancel) _
                         , dsply_title:="Demonstration of the mMsg.Dsply service" _
                         , dsply_width_max:=50 _
                         , dsply_modeless:=True _
                         , dsply_pos:="20;150")
'        Case BTN1:      MsgBox """" & BTN1 & """ pressed"
'        Case BTN2:      MsgBox """" & BTN2 & """ pressed"
'        Case BTN3:      MsgBox """" & BTN3 & """ pressed"
'        Case BTN4:      MsgBox """" & BTN4 & """ pressed"
'        Case vbYes:     MsgBox """ Yes"" pressed"
'        Case vbNo:      MsgBox """No"" pressed"
'        Case vbCancel:  MsgBox """Cancel"" pressed"
    End Select

xt: Exit Sub

eh: Select Case mMsg.ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Service_mMsg_ErrMsg()
' ----------------------------------------------------------------------------
' Demonstrates the display of a well (user friendly) structured error message.
' ----------------------------------------------------------------------------
    Const PROC = "Demo_Service_mMsg_ErrMsg"
    
    On Error GoTo eh
    Dim o As Object
    Debug.Print o.Name
    
xt: Exit Sub

eh: Select Case mMsg.ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Public Sub Demo_Service_mMsg_Monitor()
' ----------------------------------------------------------------------------
' Demonstrates monitoring a process' steps.
' ----------------------------------------------------------------------------
    Const WIDTH_MAX             As Long = 35
    Const DEMO_STEPS_PROCESSED  As Long = 9
    Const DEMO_STEPS_DISPLAYED  As Long = 5
    Dim i                       As Long
    Dim Title                   As String
    Dim Header                  As TypeMsgText
    Dim Step                    As TypeMsgText
    Dim Footer                  As TypeMsgText
       
    Title = "Demo a process' monitoring (by displaying the last " & DEMO_STEPS_DISPLAYED & " of " & DEMO_STEPS_PROCESSED & " steps)"
    With Header
        .Text = "Note: - The steps' line length exceeds the max message window width" & vbLf & _
                "        (specified for this demo to " & WIDTH_MAX & "% of the sreen width)" & vbLf & _
                "        which triggers the display of a horizontal scroll-bar" & vbLf & _
                "      - The mMsg.MonitorHeader service is used to display this information."
        .FontColor = rgbRed
        .MonoSpaced = True
        .FontSize = 8
    End With
    Footer.FontColor = rgbBlue
    Footer.Text = "Process in progress! Please wait."
    
    '~~ With the very first service call the monitoring message window is initialized
    '~~ For this demo the max window with is liited to 30% of the screen width in order to demonstrate a horizontal scroll-bar
    mMsg.MonitorHeader mon_title:=Title _
                     , mon_text:=Header _
                     , mon_width_max:=WIDTH_MAX _
                     , mon_steps_displayed:=DEMO_STEPS_DISPLAYED _
                     , mon_pos:="40;300"
    mMsg.MonitorFooter Title, Footer

    For i = 1 To DEMO_STEPS_PROCESSED
        Step.MonoSpaced = True
        
        '~~ The below 2 lines is all what is required to monitor a process step
        Step.Text = Format(i, "00") & ". Process follow-Up after " & 200 * (i - 1) & " Milliseconds (the line length exceeds the max message window width)."
        mMsg.Monitor Title, Step
                   
        Sleep 300 ' Simmulation of some process time
    Next i
    
    Footer.Text = "Process finished! Close this window (displayed by the mMsg.MonitorFooter service)"
    mMsg.MonitorFooter Title, Footer
     
End Sub

Private Sub Demo_Service_mMsg_MsgInstance()
' ------------------------------------------------------------------------------
' Demonstration of the mMsg.MsgInstance service called by the mMsg.Monitor
' service for each process identified by the Title displayed at the window
' handle bar. Displayed are 5 individual processes, each updated with 5 process
' steps. Finally the 5 monitoring windows are unloaded in reverse order, again
' by means of the mMsg.MsgInstance service.
' ------------------------------------------------------------------------------
    Const PROC              As String = "Demo_Service_mMsg_MsgInstance"
    Const INIT_TOP          As Single = 100
    Const INIT_LEFT         As Single = 50
    Const OFFSET_H          As Single = 80
    Const OFFSET_V          As Single = 20
    Const T_WAIT            As Single = 200
    Const DEMO_PROCESSES    As Long = 5
    Const DEMO_STEPS        As Long = 4
    Const DEMO_PROCESS_NAME As String = "(Sub)Process/Instance-"
    
    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim sTitle      As String
    Dim Step        As TypeMsgText
    
    j = 1
    For i = 1 To DEMO_PROCESSES
        '~~ Establish 5 monitoring instances
        '~~ Note that the instances are identified by their different titles and positioned on screen
        '~~ along with the very first service call
        sTitle = DEMO_PROCESS_NAME & i
        Step.Text = "Process step " & j
        mMsg.Monitor mon_title:=sTitle _
                   , mon_text:=Step _
                   , mon_steps_displayed:=DEMO_STEPS + 1 _
                   , mon_width_min:=20 _
                   , mon_pos:=INIT_TOP + OFFSET_V * (i - 1) & ";" & INIT_LEFT + OFFSET_H * (i - 1)
        Sleep T_WAIT
    Next i
    
    For j = 2 To DEMO_STEPS
        '~~ Display in each of the instances an additional progress message
        For i = 1 To DEMO_PROCESSES
            '~~ Go through all instances and add a message line
            sTitle = DEMO_PROCESS_NAME & i
            Step.Text = "Process step " & j
            mMsg.Monitor sTitle, Step
            Sleep T_WAIT
        Next i
    Next j
    
    Stop
    For i = DEMO_PROCESSES To 1 Step -1
        '~~ Unload the instances in reverse order
        mMsg.MsgInstance fi_key:=DEMO_PROCESS_NAME & i, fi_unload:=True
        Sleep T_WAIT
    Next i
    
xt: Exit Sub

eh: Select Case mMsg.ErrMsg(ErrSrc(PROC))
        Case vbResume:      Stop: Resume
        Case Else:          GoTo xt
    End Select
End Sub

Private Function ErrSrc(Optional ByVal s As String) As String
    ErrSrc = "mMsgUsageDemos." & s
End Function

