Attribute VB_Name = "mTest"
Option Private Module
Option Explicit
Option Compare Text
' -----------------------------------------------------
' Standard Module mMsg
'          Basic declarations, procedures, methods and
'          functions coomon im most VBProjects.
'
' Please note:
' Errors raised by the tested procedures cannot be
' asserted since they are not passed on to the calling
' /entry procedure. This would require the Common
' Standard Module mErrHndlr which is not used with this
' module by intention.
'
' lScreenWidth. Rauschenberger, Berlin Feb 2020
' -----------------------------------------------------
Dim sMsgTitle   As String
Dim sMsgLabel   As String
Dim sMsg1Label  As String
Dim sMsg2Label  As String
Dim sMsg3Label  As String
Dim sMsgText    As String
Dim sMsg1Text   As String
Dim sMsg2Text   As String
Dim sMsg3Text   As String
Dim vReply1     As Variant
Dim vReply2     As Variant
Dim vReply3     As Variant
Dim vReply4     As Variant
Dim vReply5     As Variant
Dim vReplies    As Variant
Dim vReply      As Variant
Dim vReplied    As Variant

' Regression testing makes use of all available design means - by the way testing them.
' Each test procedure is independant from another one and thus may be performed solo.
' -------------------------------------------------------------------------------------
Public Sub Regression()
                             
    If mTest.MinimumWidth1() = vReply1 Then Exit Sub
    If mTest.MinimumWidth2() = vReply1 Then Exit Sub
    If mTest.WidthDeterminedByTitle() = vReply1 Then Exit Sub
    If mTest.WidthDeterminedByMonospacedMessageSection() = vReply1 Then Exit Sub
    If mTest.WidthDeterminedByReplyButtons1() = vReply1 Then Exit Sub
    If mTest.WidthDeterminedByReplyButtons2() = vReply1 Then Exit Sub

End Sub

Private Function MinimumWidth1() As Variant
    
    Const PROC  As String = "MinimumWidth1"
    Unload fMsg                     ' Ensures a message starts from scratch
    fMsg.MinimumFormWidth = 250     ' Overwrite default
    
    sMsgTitle = PROC
    sMsg1Label = "Test description:"
    sMsg1Text = "None of the elements which determine the message form width (title, monospaced message section, reply buttons) " & _
                "exceed the specified minimum form width (explicitely specified this test = " & fMsg.MinimumFormWidth & ")."
    sMsg2Label = "Expected test result"
    sMsg2Text = "The message form width is adjusted to the specified minimum width."
    sMsg3Label = "Please note:"
    sMsg3Text = "The two message sections in this test use a proportional font and thus are " & _
                "adjusted to the form's width determined by other factors." & vbLf & _
                "The message form height is ajusted to the need up to the specified maximum heigth " & _
                "based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReply1 = "Stop"
    vReply2 = "Continue"
    vReplies = "Stop,Continue"

    MinimumWidth1 = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function

Private Function MinimumWidth2() As Variant

    Const PROC  As String = "MinimumWidth2"
    Unload fMsg                     ' Ensures a message starts from scratch
    fMsg.MinimumFormWidth = 350     ' Overwrite default
    
    sMsgTitle = PROC
    sMsg1Label = "Test description:"
    sMsg1Text = "None of the elements which determine the message form width (title, monospaced message section, reply buttons) " & _
                "exceed the specified minimum form width (explicitely specified this test = " & fMsg.MinimumFormWidth & ")."
    sMsg2Label = "Expected test result"
    sMsg2Text = "The message form width is adjusted to the specified minimum width."
    sMsg3Label = "Please note:"
    sMsg3Text = "The two message sections in this test use a proportional font and thus are " & _
                "adjusted to the form's width determined by other factors." & vbLf & _
                "The message form height is ajusted to the need up to the specified maximum heigth " & _
                "based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReplies = "Stop,Continue"
        
    MinimumWidth2 = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function

Private Function WidthDeterminedByTitle() As Variant
    
    Const PROC  As String = "WidthDeterminedByTitle"
    Unload fMsg                     ' Ensures a message starts from scratch
    '~~ Go with defaults
    
    sMsgTitle = ": This title uses more space than the minimum specified message form width and thus the width is determined by the title"
    sMsg1Text = "The length of the title determines the minimum width of the message form - unless the title's lenght would exceed the specified maximum form width which for this test is " & fMsg.MaximumFormWidth & " (which is the specified " & fMsg.MaxFormWidthPercentageOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result"
    sMsg2Text = "The message form width is adjusted to the title's lenght."
    sMsg3Label = "Please note:"
    sMsg3Text = "The two message sections in this test do use a proportional font and thus are adjusted to form width determined by other factors." & vbLf & _
                "The message form height is ajusted to the need up to the specified maximum heigth based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReplies = "Stop,Continue"
    
    WidthDeterminedByTitle = _
    mMsg.Msg( _
             msgtitle:=PROC & sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function

Private Function WidthDeterminedByMonospacedMessageSection() As Variant
    
    Const PROC  As String = "WidthDeterminedByMonospacedMessageSection"
    Unload fMsg                     ' Ensures a message starts from scratch
    ' Go with defaults
    
    sMsgTitle = PROC
    sMsg1Label = "Test description:"
    sMsg1Text = "The length of the longest monospaced message section line determines the minimum width of the message form - unless it does not exceed the specified maximum form width which for this test is " & fMsg.MaximumFormWidth & " (which is the specified " & fMsg.MaxFormWidthPercentageOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result:"
    sMsg2Text = "The message form width is adjusted to the longest line in the monospaced message section and all other message sections are adjusted to this (enlarged) width."
    sMsg3Label = "Please note the following:"
    sMsg3Text = "- In contrast to the message section above, this section makes use of the ""monospaced"" option" & vbLf & _
                "  which ensures the message text is not ""wrapped around""." & vbLf & _
                "- The message form height is ajusted to the need up to the specified maximum heigth" & vbLf & _
                "  based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReplies = "Stop,Continue"
    
    WidthDeterminedByMonospacedMessageSection = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msg3monospaced:=True, _
             msgreplies:=vReplies _
            )
End Function

Private Function WidthDeterminedByReplyButtons1() As Variant
    
    Const PROC  As String = "WidthDeterminedByReplyButtons1"
    Unload fMsg                     ' Ensures a message starts from scratch
    ' Go with defaults
    
    sMsgTitle = PROC
    sMsg1Label = "Test description:"
    sMsg1Text = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaximumFormWidth & " (which is the specified " & fMsg.MaxFormWidthPercentageOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result:"
    sMsg2Text = "The message form width is adjusted to the space required by the reply buttons and all message sections are adjusted to this (enlarged) width."
    sMsg3Label = "Please note the following:"
    sMsg3Text = "The message form height is adjusted to the required height limited by the specified maximum heigth " & _
                "based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReplies = "Stop,Continue,Unused Reply Button 1" & vbLf & "(continues testing),Unused Reply Button 2" & vbLf & "(continues testing)"
    
    WidthDeterminedByReplyButtons1 = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function


Private Function WidthDeterminedByReplyButtons2() As Variant
    
    Const PROC  As String = "WidthDeterminedByReplyButtons2"
    Unload fMsg                     ' Ensures a message starts from scratch
    ' Go with defaults
    
    sMsgTitle = PROC
    sMsg1Label = "Test description:"
    sMsg1Text = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaximumFormWidth & " (which is the specified " & fMsg.MaxFormWidthPercentageOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result:"
    sMsg2Text = "The message form width is adjusted to the space required by the reply buttons and all message sections are adjusted to this (enlarged) width."
    sMsg3Label = "Please note the following:"
    sMsg3Text = "The message form height is adjusted to the required height limited by the specified maximum heigth " & _
                "based on the screen size which for this test is " & fMsg.MaxFormHeightPercentageOfScreenSize & "%."
    vReplies = "Stop,Continue,Unused Reply Button 1" & vbLf & "(continues testing),Unused Reply Button 2" & vbLf & "(continues testing), Unused Reply Button 3" & vbLf & "(continues testing)"
    
    WidthDeterminedByReplyButtons2 = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function

Public Sub All_Tests()
    mTest.Test_1_reply_button_1_proportional_spaced_message_section
    mTest.Test_2_reply_buttons_1_proportional_spaced_message_section
    mTest.Test_3_reply_buttons_1_proportional_spaced_message_section
    mTest.Test_4_reply_buttons_1_monospaced_message_section
    mTest.Test_4_reply_buttons_3_mixed_message_sections
    mTest.Test_5_reply_buttons_3_mixed_message_sections
    mTest.Test_1_reply_button_1_monospaced_message_section_exceeding_max_form_width
    mTest.Test_1_reply_button_1_monospaced_message_section_which_exceeds_the_specified_max_form_height_and_width
    mTest.Test_1_reply_button_1_proportional_space_message_section
End Sub

' Test: The maximum form size is reduced stepwise until it has reached the minimum form with
' Note: In order to be able to assert message form properties
'       the form is not unloaded when the conditional compile argument Test = TRUE
' --------------------------------------------------------------------------------
Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

' Repeat the string (pattern) n (ntimes) times, otionally prefixed
' with a line number
' -------------------------------------------------------------------------
Private Function Repeat(ByVal ntimes As Long, _
                        ByVal pattern As String, _
                        Optional withlinenumbers As Boolean = False) As String
    
    Const MAX_STRING_LENGTH = 12000
    Dim i   As Long
    Dim s   As String
    Dim ln  As String
    Dim sFormat As String
    
    On Error Resume Next
    If withlinenumbers Then sFormat = String$(Len(CStr(ntimes)), "0") & " "
    
    For i = 1 To ntimes
        If withlinenumbers Then ln = Format(i, sFormat)
        s = s & ln & pattern
        If Err.Number <> 0 Then
            Debug.Print "Repeate had to stop after " & i & "which resulted in a string length of " & Len(s)
            Repeat = s
            Exit Function
        End If
    Next i
    Repeat = s
End Function

Public Sub Test_1_reply_button_1_monospaced_message_section_which_exceeds_the_specified_max_form_height_and_width()

    sMsgTitle = "Test: 1 monospaced message section which exceeds the maximum specified forms width and height (specified by the fMsg constant FORM_WIDTH_MAX_POW and FORM_HEIGHT_MAX_POW)"
    sMsgLabel = "Note that the message comes with a vertical and a horizontal scroll bar"
    sMsgText = Repeat(ntimes:=100, pattern:=Repeat(5, " Test Message with 1 reply button. Reply with <Ok>!") & vbLf, withlinenumbers:=True)
    vReplies = vbOKOnly
    vReply = vbOK
    
    vReplied = mMsg.Msg1( _
                        msgtitle:=sMsgTitle, _
                        msgtext:=sMsgText, _
                        msgmonospaced:=True, _
                        msgreplies:=vReplies _
                       )
    Debug.Assert vReplied = vReply

End Sub

' Common error message test
' ------------------------------------
Public Sub Test_Error_Message_Simple()
    
    Const PROC = "Test_Error_Message_Simple"
    Dim sMsg    As String
    Dim sInfo   As String
    
    sMsg = "The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog."
    sInfo = "Note 1: The error line is only displayed when one not 0 is provided" & vbLf & _
            "Note 2: This optional additional info has been provided via the errdescription parameter delimited by two vertical bars (""|"")"
    
    mMsg.ErrMsg errnumber:=1, _
                errdescription:=sMsg & "||" & sInfo, _
                errsource:=ErrSrc(PROC), _
                errpath:="None provided (optional!)", _
                errline:="12"
End Sub

Public Sub Test_1_reply_button_1_monospaced_message_section_exceeding_max_form_width()

    fMsg.MaximumFormWidth = 60
    
    sMsgTitle = "Test: The width of the first message paragraph (monospaced) exceeds the specified maximum forms width (fMsg constant FORM_WIDTH_MAX_POW)"
    sMsg1Text = Repeat(ntimes:=10, pattern:=Repeat(ntimes:=4, pattern:="Test Message with 1 reply button. Reply with <Ok>!") & vbLf, withlinenumbers:=True)
    sMsg2Label = "Please note!"
    sMsg2Text = "The width of the above monospaced message section exceeds the maximum specified form width, which is set " & fMsg.MaximumFormWidth & "% of the screen size. " & _
                "Because it is monospaced the text is  n o t  wrapped around but provided with a horizontal scrollbar instead."
    vReplies = vbOKOnly
    vReply = vbOK
    
    vReplied = mMsg.Msg( _
                        msgtitle:=sMsgTitle, _
                        msg1label:=sMsg1Label, _
                        msg1text:=sMsg1Text, _
                        msg1monospaced:=True, _
                        msg2label:=sMsg2Label, _
                        msg2text:=sMsg2Text, _
                        msgreplies:=vReplies _
                       )
    Debug.Assert vReplied = vReply

End Sub

Public Sub Test_1_reply_button_1_proportional_space_message_section()

    sMsgTitle = "Test: The width of the first message paragraph (proportional spaced) exceeds the specified maximum forms width (fMsg constant FORM_WIDTH_MAX_POW)"
    sMsg1Label = "Note that the text below is wrapped around because it is proportional spaced"
    sMsg1Text = Repeat(ntimes:=5, pattern:="Test Message with 1 reply button. Reply with <Ok>! ")
    vReplies = vbOKOnly
    vReply = vbOK
    
    vReplied = mMsg.Msg( _
                        msgtitle:=sMsgTitle, _
                        msg1label:=sMsg1Label, _
                        msg1text:=sMsg1Text, _
                        msgreplies:=vReplies _
                  )
    Debug.Assert vReplied = vReply

End Sub

Public Sub Test_1_reply_button_1_proportional_spaced_message_section()

    Unload fMsg ' Not unloaded when conditional compile argument Test = 1 !
    fMsg.MaximumFormHeight = 50
    fMsg.MaximumFormWidth = 90
    
    sMsgTitle = "Test: One proportional spaced message section with 1 reply button"
    sMsg1Text = Repeat(20, Repeat(2, "Test Message with 1 reply button. Reply with <Ok>!") & vbLf, True)
    vReplies = vbOKOnly
    vReply = vbOK
    
    vReplied = mMsg.Msg1( _
                         msgtitle:=sMsgTitle, _
                         msgtext:=sMsg1Text, _
                         msgreplies:=vReplies _
                        )
    
    Debug.Assert vReplied = vReply
    Unload fMsg ' Not unloaded when conditional compile argument Test = 1 !
    
End Sub

Public Sub Test_2_reply_buttons_1_proportional_spaced_message_section()

    sMsgTitle = "Test: One proportional spaced message section with 2 reply buttons"
    sMsg1Text = "The 2 reply buttons width, height and position is adjusted in accordance with their caption texts." & vbLf & vbLf & _
                "Reply this test with <Yes> - which is asserted!"
    vReplies = vbYesNo
    vReply = vbYes
    
    vReplied = mMsg.Msg1( _
                         msgtitle:=sMsgTitle, _
                         msgtext:=sMsg1Text, _
                         msgreplies:=vReplies _
                        )
    Debug.Assert vReplied = vReply
    
End Sub

Public Sub Test_3_reply_buttons_1_proportional_spaced_message_section()

    sMsgTitle = "Test: One proportional spaced message section with 3 reply buttons"
    sMsg1Text = "The 2 reply buttons width, height and position is adjusted in accordance with their caption texts." & vbLf & vbLf & _
                "Reply this test with <Yes> - which is asserted!"
    
    vReplies = vbYesNoCancel
    vReply = vbYes
    
    vReplied = mMsg.Msg1( _
                         msgtitle:=sMsgTitle, _
                         msgtext:=sMsg1Text, _
                         msgreplies:=vReplies _
                        )
    Debug.Assert vReplied = vReply
End Sub

Public Sub Test_4_reply_buttons_1_monospaced_message_section()

    sMsgTitle = "Test: One proportional spaced message section with 4 reply buttons"
    
    sMsg1Text = "1. The title is never trunctated (see 4. final width adjustment)" & vbLf & _
                "2. The message text may optionally be displayed monospaced" & vbLf & _
                "   to support a message section with proper indented lines (like these here or the error-path in an error message)" & vbLf & _
                "3. There may be up to 5 reply buttons either up to 3 like the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
                "   or up to 5 with any number of text and lines (all buttons are adjusted accordingly)" & vbLf & _
                "   and the reply value corresponds with the button's content, i.e. either vbOk, vbYes, ..." & vbLf & _
                "   or the displayed text" & vbLf & _
                "4. The final form width is dertermined by:" & vbLf & _
                "   - the width of the title (not truncated unless the form width does not exceed its specified maximum)" & vbLf & _
                "   - the longest fixed font message text line (the above one)" & vbLf & _
                "   - the required space for the displayed reply buttons"
    sMsg2Text = "Reply this test with <Display Execution Trace>!"
    
    vReply1 = "Update Target" & vbLf & "with Source"
    vReply2 = "Update Source" & vbLf & "with Target"
    vReply3 = "Display" & vbLf & "Execution Trace"
    vReply4 = "Ignore"
    vReply5 = vbNullString
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply3
    
    vReplied = mMsg.Msg1( _
                         msgtitle:=sMsgTitle, _
                         msgtext:=sMsg1Text & vbLf & vbLf & sMsg2Text, _
                         msgmonospaced:=True, _
                         msgreplies:=vReplies _
                       )
    Debug.Assert CStr(vReplied) = CStr(vReply)
    
End Sub

Public Sub Test_4_reply_buttons_3_mixed_message_sections()

    sMsgTitle = "Test: 3 mixed (proportional and monospaced) message sections with 4 reply buttons"
    
    sMsg1Label = "General enhancements in contrast to the MsgBox"
    sMsg1Text = "1. The title is never trunctated (see 4. final width adjustment)" & vbLf & _
                "2. The message text may optionally be displayed monospaced" & vbLf & _
                "   to support a message section with proper indented lines (like these here or the error-path in an error message)" & vbLf & _
                "3. There may be up to 5 reply buttons either up to 3 like the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
                "   or up to 5 with any number of text and lines (all buttons are adjusted accordingly)" & vbLf & _
                "   and the reply value corresponds with the button's content, i.e. either vbOk, vbYes, ..." & vbLf & _
                "   or the displayed text" & vbLf & _
                "4. Final form with and height adjustment (see below)"
    
    sMsg2Label = "Form width and height adjustment"
    sMsg2Text = "The width adjustment considers:" & vbLf & _
                "- the title width (never truncated anymore)" & vbLf & _
                "- the maximum length of any fixed font message block" & vbLf & _
                "- total width of the displayed reply buttons" & vbLf & _
                "- the minimum width, optionally specified by the parameter ""minformwidth"", defaults to constant FORM_WIDTH_MIN" & vbLf & _
                "- the maximum width, specified by the constant FORM_WIDTH_MAX_POW which is a percentage of the screen height" & vbLf & _
                "- the largest messae section may end up with a horizontal scroll bar when the max form width would be exeeded otherwise" & vbLf & _
                vbLf & _
                "The height adjustment considers:" & vbLf & _
                "- the height required by the displayed controls" & vbLf & _
                "- the maximum height, specified by the constnad FORM_HEIGHT_MAX_POW which is a percentage of the screen height" & vbLf & _
                "- the largest message section may end up with a vertical scroll bar when the max form height exceeds otherwise"
    
    sMsg3Label = vbNullString
    sMsg3Text = "Reply this test with <Reply 4>"
    
    vReply1 = "Reply 1"
    vReply2 = "Reply 2"
    vReply3 = "Reply lines" & vbLf & "determine the" & vbLf & "button height"
    vReply4 = "Reply 4"
    vReply5 = vbNullString
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply4
    
    vReplied = mMsg.Msg( _
                   msgtitle:=sMsgTitle, _
                   msg1text:=sMsg1Text, msg1label:=sMsg1Label, msg1monospaced:=True, _
                   msg2text:=sMsg2Text, msg2label:=sMsg2Label, _
                   msg3text:=sMsg3Text, msg3label:=sMsg3Label, _
                   msgreplies:=vReplies _
                  )
    Debug.Assert vReplied = vReply
    
End Sub

Public Sub Test_5_reply_buttons_3_mixed_message_sections()

    sMsgTitle = "mMsg.Msg works pretty much like MsgBox but with significant enhancements (see below)"
    
    sMsg1Label = "General"
    sMsg1Text = "- The title will never be truncated" & vbLf & _
             "- There are up to 3 message paragraphs, each with an optional label/header" & vbLf & _
             "- Each message paragraph may be in a proportional or fixed font (like these two)" & vbLf & _
             "  supporting indented text like this one - or the display of the error-path in an error message" & vbLf & _
             "- There are up to 5 reply buttons. 3 work exactly like with the MsgBox (vbOkOnly, vbYesNo, ...)" & vbLf & _
             "  but all may as well contain any string and the reply value corresponds with the clicked reply button"
    
    sMsg2Label = "Window width adjustment"
    sMsg2Text = "The message window width considers:" & vbLf & _
             "- the title width (never truncated anymore)" & vbLf & _
             "- the maximum length of any fixed font message block" & vbLf & _
             "- total width of the displayed reply buttons" & vbLf & _
             "- specified minimum window width"
    
    sMsg3Label = vbNullString
    sMsg3Text = "Reply this test with <Reply lines determine the button height>"
    
    vReply1 = "Reply 1"
    vReply2 = "Reply 2"
    vReply3 = "Reply lines" & vbLf & "determine the" & vbLf & "button height"
    vReply4 = "Reply 4"
    vReply5 = "Reply 5"
    vReplies = Join(Array(vReply1, vReply2, vReply3, vReply4, vReply5), ",")
    vReply = vReply3
    
    vReplied = mMsg.Msg( _
                        msgtitle:=sMsgTitle, _
                        msg1text:=sMsg1Text, msg1label:=sMsg1Label, msg1monospaced:=True, _
                        msg2text:=sMsg2Text, msg2label:=sMsg2Label, _
                        msg3text:=sMsg3Text, msg3label:=sMsg3Label, _
                        msgreplies:=vReplies _
                       )
    Debug.Assert vReplied = vReply
End Sub

' Takes several seconds !
' A veeeery long lasting test produced 7,300.000 characters long string.
' So theres no limit but the exec time which means that this simple
' solution suffice only this testing purpose.
' ----------------------------------------------------------------------
Public Sub Test_Repeat()
    Debug.Print Repeat(10, "a" & vbLf, True)
End Sub

