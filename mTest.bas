Attribute VB_Name = "mTest"
Option Explicit
Option Compare Text
' -----------------------------------------------------
' Standard Module mTest
'          All tests for a complete regression test.
'          Obligatory performed after any modification.
'          Ammended when new features or functions are
'          implemented.
'
' Replies setup:
' - Reply button 1 = "Previous" when test procedure is performed via the Regresson procedure
'                    And the test is not the first one
' - Reply Button 2 = "Stop" throughout all tests
' - Reply Button 3 = "Next" when the test procedure is performed via the Regressin procedure
'                    and the procedure is not the very last one
' - Reply Button 4 = Optionally used for each test specifically
' - Reply Button 5 = Optionally used for each test specifically
'
' Please note:
' Errors raised by the tested procedures cannot be
' asserted since they are not passed on to the calling
' /entry procedure. This would require the Common
' Standard Module mErrHndlr which intentionally is not
' used by this module.
'
' W. Rauschenberger, Berlin June 2020
' -----------------------------------------------------
Const C_PREV        As String = "Previous"
Const C_STOP        As String = "Stop"
Const C_NEXT        As String = "Next"

Dim lMinFormWidth   As Long
Dim lTest           As Long
Dim sMsgTitle       As String
Dim sMsgLabel       As String
Dim sMsg1Label      As String
Dim sMsg2Label      As String
Dim sMsg3Label      As String
Dim sMsgText        As String
Dim sMsg1Text       As String
Dim sMsg2Text       As String
Dim sMsg3Text       As String
Dim vReplies        As Variant
Dim vReply          As Variant
Dim vReply1         As Variant
Dim vReply2         As Variant
Dim vReply3         As Variant
Dim vReply4         As Variant
Dim vReply5         As Variant
Dim vReplied        As Variant
Dim siUsedPoSW      As Long     ' The test specific used % of the screen width (default to 80%)

' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
Public Sub Regression()
                                     
    ThisWorkbook.Save
    Unload fMsg
    
1:  Select Case mTest.WidthDeterminedByMinimumWidth()
        Case C_STOP:    Exit Sub
    End Select

2:  Select Case mTest.WidthDeterminedByTitle()
        Case C_STOP:    Exit Sub
        Case C_PREV:    GoTo 2
    End Select

3:  Select Case mTest.WidthDeterminedByMonoSpacedMessageSection()
        Case C_STOP:    Exit Sub
        Case C_PREV:    GoTo 3
    End Select

4:  Select Case mTest.WidthDeterminedByReplyButtons()
        Case C_STOP:    Exit Sub
        Case C_PREV:    GoTo 3
    End Select

5:  Select Case mTest.MonoSpacedSectionWidthExceedsMaxFormWidth()
        Case C_STOP:    Exit Sub
        Case C_PREV:    GoTo 4
    End Select
    
End Sub

' Test 1
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByMinimumWidth( _
                 Optional ByVal vReply1 As Variant = vbNullString, _
                 Optional ByVal vReply3 As Variant = vbNullString) As Variant
    
    Const PROC      As String = "WidthDeterminedByMinimumWidth"
    Dim lIncrDecr   As Long
    
    vReply2 = C_STOP
    lTest = 1
    
    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecr = wsMsgTest.MinFormWidthIncrDecr(lTest)
    
    
    ' Initializations for this test
    With fMsg
'        .FramesWithBorder = True
'        .FramesWithCaption = True
        .MinimumFormWidth = wsMsgTest.InitMinFormWidth(lTest)
    End With
    
    vReply4 = "Repeat with" & vbLf & "minimum width" & vbLf & "+ " & lIncrDecr
    vReply5 = "Repeat with" & vbLf & "minimum width" & vbLf & "- " & lIncrDecr
    vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply4
    
repeat:
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    sMsg1Label = "Test description:":       sMsg1Text = wsMsgTest.TestDescription(lTest)
    sMsg2Label = "Expected test result:":   sMsg2Text = "The width of all message sections is adjusted to the current specified minimum form width (" & fMsg.MinimumFormWidth & " pt)."
    sMsg3Label = "Please also note:":       sMsg3Text = "The message form height is ajusted to the need " & _
                                                        "up to the specified maximum heigth which is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% and not exceeded."

    WidthDeterminedByMinimumWidth = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, msg2text:=sMsg2Text, _
             msg3label:=sMsg3Label, msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
    Select Case WidthDeterminedByMinimumWidth
        Case vReply5
            fMsg.MinimumFormWidth = wsMsgTest.InitMinFormWidth(lTest) - lIncrDecr
            vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply4
            GoTo repeat
        Case vReply4
            fMsg.MinimumFormWidth = wsMsgTest.InitMinFormWidth(lTest) + lIncrDecr
            vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply5
            GoTo repeat
        Case Else ' Stop and Next are passed on to the caller
    End Select
    
End Function

' Test 2
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByTitle( _
                 Optional ByVal vReply1 As Variant = vbNullString, _
                 Optional ByVal vReply3 As Variant = vbNullString) As Variant
    
    Const PROC  As String = "WidthDeterminedByTitle"
    Dim lIncrDecr           As Long
    
    vReply2 = C_STOP
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 2
    
    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecr = wsMsgTest.MinFormWidthIncrDecr(lTest)
    With fMsg
        .MinimumFormWidth = wsMsgTest.InitMinFormWidth(lTest)
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    sMsg1Label = "Test description:":       sMsg1Text = wsMsgTest.TestDescription(lTest)
    sMsg2Label = "Expected test result:":   sMsg2Text = "The message form width is adjusted to the title's lenght."
    sMsg3Label = "Please note:":            sMsg3Text = "The two message sections in this test do use a proportional font " & _
                                                        "and thus are adjusted to form width determined by other factors." & vbLf & _
                                                        "The message form height is ajusted to the need up to the specified " & _
                                                        "maximum heigth based on the screen size which for this test is " & _
                                                        fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vReplies = vReply1 & "," & vReply2 & "," & vReply3
    
    WidthDeterminedByTitle = _
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

' Test 3
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByMonoSpacedMessageSection( _
                 Optional ByVal vReply1 As Variant = vbNullString, _
                 Optional ByVal vReply3 As Variant = vbNullString) As Variant
    
    Const PROC      As String = "WidthDeterminedByMonoSpacedMessageSection"
    Dim lIncrDecr   As Long
    
    lTest = 3
    vReply2 = C_STOP

    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecr = wsMsgTest.MaxFormWidthIncrDecr(lTest)
    
    ' Initializations for this test
    fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsgTest.InitMaxFormWidth(lTest)
    
    vReply4 = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & lIncrDecr
    vReply5 = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & lIncrDecr
    vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply5
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    sMsg1Label = "Test description:":           sMsg1Text = wsMsgTest.TestDescription(lTest)
    sMsg2Label = "Expected test result:":       sMsg2Text = "Initally, the message form width is adjusted to the longest line in the " & _
                                                            "monospaced message section and all other message sections are adjusted " & _
                                                            "to this (enlarged) width." & vbLf & _
                                                            "When the maximum form width is reduced by " & lIncrDecr & " % the monospaced message section is displayed with a horizontal scroll bar."
    sMsg3Label = "Please note the following:":  sMsg3Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                                                            "  the message text is not ""wrapped around""." & vbLf & _
                                                            "- The message form height is ajusted to the need up to the specified maximum heigth" & vbLf & _
                                                            "  based on the screen size which for this test is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    
repeat:
    With fMsg
        .FramesWithCaption = True  ' defaults to false, set to true for test purpose only
        .FramesWithBorder = True  ' defaults to false, set to true for test purpose only
    End With
    WidthDeterminedByMonoSpacedMessageSection = _
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
    Select Case WidthDeterminedByMonoSpacedMessageSection
        Case vReply5
            fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsgTest.InitMaxFormWidth(lTest) - lIncrDecr
            vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply4
            GoTo repeat
        Case vReply4
            fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsgTest.InitMaxFormWidth(lTest) + lIncrDecr
            vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply5
            GoTo repeat
        Case Else ' Stop and Next are passed on to the caller
    End Select
    
End Function

' Test 4
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByReplyButtons( _
                 Optional ByVal vReply1 As Variant = vbNullString, _
                 Optional ByVal vReply3 As Variant = vbNullString) As Variant
    
    Const PROC  As String = "WidthDeterminedByReplyButtons1"
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 4
    vReply2 = C_STOP
    
    ' Initializations for this test
    With fMsg
        .FramesWithBorder = True
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    sMsg1Label = "Test description:"
    sMsg1Text = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " (which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result:"
    sMsg2Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    sMsg3Label = "Please also note:"
    sMsg3Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                "which is a percentage of the screen size (for this test = " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vReply4 = "Repeat with" & vbLf & "5 buttons"
    vReply5 = "Repeat with" & vbLf & "4 buttons"
    
    If vReply1 = vbNullString And vReply3 = vbNullString Then
        '~~ Test is performed "standalone"
        vReplies = "Dummy," & vReply2 & ",Dummy," & vReply4
    Else
        vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply4
    End If
    
repeat:
    WidthDeterminedByReplyButtons = _
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
    
    Select Case WidthDeterminedByReplyButtons
        Case vReply4
            If vReply1 = vbNullString And vReply3 = vbNullString Then
                '~~ Test is performed "standalone"
                vReplies = "Dummy," & vReply2 & ",Dummy,Dummy," & vReply5
            Else
                '~~ Test is performed within Regression
                vReplies = vReply1 & "," & vReply2 & "," & vReply3 & "," & vReply5
            End If
            GoTo repeat
        Case vReply5
            If vReply1 = vbNullString And vReply3 = vbNullString Then
                '~~ Test is performed "standalone"
                vReplies = "Dummy," & vReply2 & ",Dummy," & vReply4
            Else
                '~~ Test is performed within Regression
                vReplies = vReply1 & "," & vReply2 & "," & vReply3 & ",Dummy," & vReply4
            End If
            GoTo repeat
        Case Else ' passed on to caller
    End Select
    
End Function

' Test 5
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function MonoSpacedSectionWidthExceedsMaxFormWidth( _
                 Optional ByVal vReply1 As Variant = vbNullString, _
                 Optional ByVal vReply3 As Variant = vbNullString) As Variant

    Const PROC  As String = "MonoSpacedSectionWidthExceedsMaxFormWidth"
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 5
    vReply2 = C_STOP
    
    ' Initializations for this test
    With fMsg
        .FramesWithBorder = True
        .MaxFormWidthPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    sMsg1Label = "Test description:"
    sMsg1Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & fMsg.MaxFormWidth & " pt (the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    sMsg2Label = "Expected test result:"
    sMsg2Text = "The monospaced message section comes with a horizontal scroll bar."
    sMsg3Label = "Please note the following:"
    sMsg3Text = "- This monspaced message section exceeds the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " pt," & vbLf & _
                "  the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size." & vbLf & _
                "- The message form height is adjusted to the required height, limited to " & fMsg.MaxFormHeight & " pt," & vbLf & _
                "  the equivalent of " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% of the screen size, for this test and not reached or exceeded."
    vReplies = vReply1 & "," & vReply2 & "," & vReply3
    
    MonoSpacedSectionWidthExceedsMaxFormWidth = _
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

' Test 6
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function MonoSpacedMessageSectionExceedMaxFormHeight() As Variant

    Const PROC  As String = "MonoSpacedMessageSectionExceedMaxFormHeight"
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    With fMsg
        .FramesWithBorder = True
        .MaxFormWidthPrcntgOfScreenSize = 80
        .MaxFormHeightPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    sMsg1Label = "Test description:"
    sMsg1Text = "The width used by the monospaced message section exxceeds the maximum form width for this test (" & fMsg.MaxFormWidth & ") which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size."
    sMsg2Label = "Expected test result:"
    sMsg2Text = repeat(20, "This monospaced message comes with a horizontal scroll bar." & vbLf, True)
    sMsg3Label = "Please note the following:"
    sMsg3Text = "The message form height is adjusted to the required height limited by the specified percentage of the screen height, " & _
                "which for this test is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vReplies = vReply1 & "," & vReply2 & "," & vReply3
    
    MonoSpacedMessageSectionExceedMaxFormHeight = _
    mMsg.Msg( _
             msgtitle:=sMsgTitle, _
             msg1label:=sMsg1Label, _
             msg1text:=sMsg1Text, _
             msg2label:=sMsg2Label, _
             msg2text:=sMsg2Text, _
             msg2monospaced:=True, _
             msg3label:=sMsg3Label, _
             msg3text:=sMsg3Text, _
             msgreplies:=vReplies _
            )
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

' Repeat the string (pattern) n (ntimes) times, otionally prefixed
' with a line number
' -------------------------------------------------------------------------
Private Function repeat(ByVal ntimes As Long, _
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
            repeat = s
            Exit Function
        End If
    Next i
    repeat = s
End Function

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

Public Sub Test_Repeat()
    Debug.Print repeat(10000, "a" & vbLf, True)
End Sub

' Repeat string function good enough for up to 10 thousand repetitions.
' ---------------------------------------------------------------------
Public Function Readable(ByVal s As String) As String

    Dim i       As Long
    Dim sResult As String
    
    For i = 1 To Len(s)
        If IsUcase(Mid(s, i, 1)) Then
            sResult = sResult & " " & Mid(s, i, 1)
        Else
            sResult = sResult & Mid(s, i, 1)
        End If
    Next i
    Readable = Right(sResult, Len(sResult) - 1)

End Function

Function IsUcase(ByVal s As String) As Boolean

    Dim i   As Integer
    i = Asc(s)
    IsUcase = (i >= 65 And i <= 90) Or _
              (i >= 192 And i <= 214) Or _
              (i >= 216 And i <= 223) Or _
              (i = 128) Or _
              (i = 138) Or _
              (i = 140) Or _
              (i = 142) Or _
              (i = 154) Or _
              (i = 156) Or _
              (i >= 158 And i <= 159) Or _
              (i = 163) Or _
              (i = 165)
End Function
