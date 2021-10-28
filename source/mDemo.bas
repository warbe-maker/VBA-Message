Attribute VB_Name = "mDemo"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Public Sub Demos()
    DemoMsgDsplyService_1
    DemoMsgDsplyService_2
End Sub

Public Sub DemoMsgDsplyService_1()
    Const max_width     As Long = 50
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
                   & "for this demo has been specified " & max_width & "% of the screen width (the default would be 80%)" & vbLf _
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
                   , dsply_max_height:=MAX_HEIGHT _
                   , dsply_max_width:=max_width _
                    ) <> vbOK
    Wend
    
End Sub

Public Sub DemoMsgDsplyService_2()
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
       
   vReturn = Dsply(dsply_title:="Any title", _
                   dsply_msg:=Message, _
                   dsply_buttons:=mMsg.Buttons(vbAbortRetryIgnore, vbLf, B1, B2, B3, vbLf, B4, B5, B6, vbLf, B7) _
                  )
   MsgBox "Button """ & mMsg.ReplyString(vReturn) & """ had been clicked"
   
End Sub

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mDemo." & s:  End Property

Public Sub Demo_Monitor_Service()
' ------------------------------------------------------------------------------
'
 ' ------------------------------------------------------------------------------
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
    mMsg.Form MonitorTitle, frm_unload:=True ' Ensure there is no process monitoring with this title still displayed
        
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
            mMsg.Monitor prgrs_title:=MonitorTitle _
                       , prgrs_msg:=ProgressStep _
                       , prgrs_msg_monospaced:=True _
                       , prgrs_header:=MONITOR_HEADER _
                       , prgrs_max_height:=wsTest.MsgHeightMaxSpecAsPoSS _
                       , prgrs_max_width:=wsTest.MsgWidthMinSpecInPt
            
            '~~ Simmulation of a process
            lWait = 100 * i
            DoEvents
            Sleep 200
        
        Else
            '~~ The last step, separated in order to display the footer along with it
            mMsg.Monitor prgrs_title:=MonitorTitle _
                       , prgrs_msg:=ProgressStep _
                       , prgrs_header:=MONITOR_HEADER _
                       , prgrs_footer:=MONITOR_FOOTER
        End If
    Next i
    
'    Select Case mMsg.Box(box_title:="Test result of " & Readable(PROC) _
'                       , box_msg:=vbNullString _
'                       , box_buttons:=mMsg.Buttons(BTTN_PASSED, BTTN_FAILED) _
'                        )
'        Case BTTN_PASSED:       wsTest.Passed = True
'        Case BTTN_FAILED:       wsTest.Failed = True
'        Case sBttnTerminate:    wsTest.TerminateRegressionTest = True
'    End Select

xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
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

