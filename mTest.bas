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
Dim vButtons        As Variant
Dim vButton         As Variant
Dim vButton1        As Variant
Dim vButton2        As Variant
Dim vButton3        As Variant
Dim vButton4        As Variant
Dim vButton5        As Variant
Dim vButton6        As Variant
Dim vButton7        As Variant
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
    Optional ByVal vButton1 As Variant = vbNullString, _
    Optional ByVal vButton3 As Variant = vbNullString) As Variant
    
    Const PROC          As String = "WidthDeterminedByMinimumWidth"
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMessage
    vButton2 = C_STOP
    lTest = 1
    
    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecrWidth = wsMsg.MinFormWidthIncrDecr(lTest)
    
    
    ' Initializations for this test
    fMsg.MinFormWidth = wsMsg.InitMinFormWidth(lTest)
    
    vButton4 = "Repeat with" & vbLf & "minimum width" & vbLf & "+ " & lIncrDecrWidth
    vButton5 = "Repeat with" & vbLf & "minimum width" & vbLf & "- " & lIncrDecrWidth
    vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton4
    
Repeat:
    With fMsg
'        .TestFrameWithBorders = True
'        .FramesWithCaption = True
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    tMsg.section(1).sLabel = "Test description:":       tMsg.section(1).sText = wsMsg.TestDescription(lTest)
    tMsg.section(2).sLabel = "Expected test result:":   tMsg.section(2).sText = "The width of all message sections is adjusted to the current specified minimum form width (" & fMsg.MinFormWidth & " pt)."
    tMsg.section(3).sLabel = "Please also note:":       tMsg.section(3).sText = "The message form height is ajusted to the need " & _
                                                                                "up to the specified maximum heigth which is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% and not exceeded."
    WidthDeterminedByMinimumWidth = _
    mMsg.Msg( _
             title:=sMsgTitle, _
             message:=tMsg, _
             buttons:=vButtons _
            )
    Select Case WidthDeterminedByMinimumWidth
        Case vButton5
            fMsg.MinFormWidth = wsMsg.InitMinFormWidth(lTest) - lIncrDecrWidth
            vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton4
            GoTo Repeat
        Case vButton4
            fMsg.MinFormWidth = wsMsg.InitMinFormWidth(lTest) + lIncrDecrWidth
            vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton5
            GoTo Repeat
        Case Else ' Stop and Next are passed on to the caller
    End Select
    
End Function

Public Function WidthDeterminedByTitle(Optional ByVal vButton1 As Variant = vbNullString, _
                                       Optional ByVal vButton3 As Variant = vbNullString) As Variant
' --------------------------------------------------------------------------------------------------
' Test 2: The optional parameters are used in conjunction with the Regression test only
' --------------------------------------------------------------------------------------------------
    
    Const PROC          As String = "WidthDeterminedByTitle"
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMessage
    
    vButton2 = C_STOP
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 2
    
    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecrWidth = wsMsg.MinFormWidthIncrDecr(lTest)
    With fMsg
        .MinFormWidth = wsMsg.InitMinFormWidth(lTest)
        .TestFrameWithBorders = True
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    tMsg.section(1).sLabel = "Test description:":       tMsg.section(1).sText = wsMsg.TestDescription(lTest)
    tMsg.section(2).sLabel = "Expected test result:":   tMsg.section(2).sText = "The message form width is adjusted to the title's lenght."
    tMsg.section(3).sLabel = "Please note:":            tMsg.section(3).sText = "The two message sections in this test do use a proportional font " & _
                                                                                "and thus are adjusted to form width determined by other factors." & vbLf & _
                                                                                "The message form height is ajusted to the need up to the specified " & _
                                                                                "maximum heigth based on the screen size which for this test is " & _
                                                                                fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vButtons = vButton1 & "," & vButton2 & "," & vButton3
    
    WidthDeterminedByTitle = _
    mMsg.Msg( _
             title:=sMsgTitle, _
             message:=tMsg, _
             buttons:=vButtons _
            )
End Function

' Test 3
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByMonoSpacedMessageSection( _
                 Optional ByVal vButton1 As Variant = vbNullString, _
                 Optional ByVal vButton3 As Variant = vbNullString) As Variant
    
    Const PROC          As String = "WidthDeterminedByMonoSpacedMessageSection"
    Dim lIncrDecrHeight As Long
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMessage
    
    lTest = 3
    vButton2 = C_STOP

    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecrWidth = wsMsg.MaxFormWidthIncrDecr(lTest)
    lIncrDecrHeight = wsMsg.MaxFormHeightIncrDecr(lTest)
    
    ' Initializations for this test
    fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(lTest)
    
    vButton4 = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & lIncrDecrWidth
    vButton5 = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & lIncrDecrWidth
    vButton6 = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & lIncrDecrHeight
    vButton7 = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & lIncrDecrHeight
    vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton5
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = wsMsg.TestDescription(lTest)
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = "Initally, the message form width is adjusted to the longest line in the " & _
                            "monospaced message section and all other message sections are adjusted " & _
                            "to this (enlarged) width." & vbLf & _
                            "When the maximum form width is reduced by " & lIncrDecrWidth & " % the monospaced message section is displayed with a horizontal scroll bar."
    tMsg.section(3).sLabel = "Please note the following:"
    tMsg.section(3).sText = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                            "  the message text is not ""wrapped around""." & vbLf & _
                            "- The message form height is ajusted to the need up to the specified maximum heigth" & vbLf & _
                            "  based on the screen size which for this test is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    tMsg.section(3).bMonspaced = True
    
    Do
        With fMsg
    '        .TestFrameWithCaptions = True  ' defaults to false, set to true for test purpose only
            .TestFrameWithBorders = True  ' defaults to false, set to true for test purpose only
        End With
        WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Msg( _
                 title:=sMsgTitle, _
                 message:=tMsg, _
                 buttons:=vButtons _
                )
        Select Case WidthDeterminedByMonoSpacedMessageSection
            Case vButton5
                fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(lTest) - lIncrDecrWidth
                vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton4
            Case vButton4
                fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(lTest) + lIncrDecrWidth
                vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton5
            Case Else: Exit Do ' Stop and Next are passed on to the caller
        End Select
    Loop
    
End Function

' Test 4
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function WidthDeterminedByReplyButtons( _
                 Optional ByVal vButton1 As Variant = vbNullString, _
                 Optional ByVal vButton3 As Variant = vbNullString) As Variant
    
    Const PROC  As String = "WidthDeterminedByReplyButtons1"
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 4
    vButton2 = C_STOP
    
    ' Initializations for this test
    With fMsg
'        .TestFrameWithBorders = True
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " (which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    tMsg.section(3).sLabel = "Please also note:"
    tMsg.section(3).sText = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                            "which is a percentage of the screen size (for this test = " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vButton4 = "Repeat with" & vbLf & "5 buttons"
    vButton5 = "Repeat with" & vbLf & "4 buttons"
    
    If vButton1 = vbNullString And vButton3 = vbNullString Then
        '~~ Test is performed "standalone"
        vButtons = "Dummy," & vButton2 & ",Dummy," & vButton4
    Else
        vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton4
    End If
    
    Do
        WidthDeterminedByReplyButtons = _
        mMsg.Msg( _
                 title:=sMsgTitle, _
                 message:=tMsg, _
                 buttons:=vButtons _
                )
        
        Select Case WidthDeterminedByReplyButtons
            Case vButton4
                If vButton1 = vbNullString And vButton3 = vbNullString Then
                    '~~ Test is performed "standalone"
                    vButtons = "Dummy," & vButton2 & ",Dummy,Dummy," & vButton5
                Else
                    '~~ Test is performed within Regression
                    vButtons = vButton1 & "," & vButton2 & "," & vButton3 & "," & vButton5
                End If
            Case vButton5
                If vButton1 = vbNullString And vButton3 = vbNullString Then
                    '~~ Test is performed "standalone"
                    vButtons = "Dummy," & vButton2 & ",Dummy," & vButton4
                Else
                    '~~ Test is performed within Regression
                    vButtons = vButton1 & "," & vButton2 & "," & vButton3 & ",Dummy," & vButton4
                End If
            Case Else: Exit Do ' passed on to caller
        End Select
    Loop
    
End Function

' Test 5
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function MonoSpacedSectionWidthExceedsMaxFormWidth( _
                 Optional ByVal vButton1 As Variant = vbNullString, _
                 Optional ByVal vButton3 As Variant = vbNullString) As Variant

    Const PROC  As String = "MonoSpacedSectionWidthExceedsMaxFormWidth"
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    lTest = 5
    vButton2 = C_STOP
    
    ' Initializations for this test
    With fMsg
        .TestFrameWithBorders = True
        .MaxFormWidthPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & fMsg.MaxFormWidth & " pt (the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = "The monospaced message section comes with a horizontal scroll bar."
    tMsg.section(3).sLabel = "Please note the following:"
    tMsg.section(3).sText = "- This monspaced message section exceeds the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " pt," & vbLf & _
                            "  the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size." & vbLf & _
                            "- The message form height is adjusted to the required height, limited to " & fMsg.MaxFormHeight & " pt," & vbLf & _
                            "  the equivalent of " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% of the screen size, for this test and not reached or exceeded."
    vButtons = vButton1 & "," & vButton2 & "," & vButton3
    
    MonoSpacedSectionWidthExceedsMaxFormWidth = _
    mMsg.Msg( _
             title:=sMsgTitle, _
             message:=tMsg, _
             buttons:=vButtons _
            )
End Function

' Test 6
' The optional parameters are used in conjunction with the Regression test only
' -----------------------------------------------------------------------------
Public Function MonoSpacedMessageSectionExceedMaxFormHeight() As Variant

    Const PROC  As String = "MonoSpacedMessageSectionExceedMaxFormHeight"
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    With fMsg
        .TestFrameWithBorders = True
        .MaxFormWidthPrcntgOfScreenSize = 80
        .MaxFormHeightPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = "Test " & lTest & ": " & Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the monospaced message section exxceeds the maximum form width for this test (" & fMsg.MaxFormWidth & ") which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = Repeat(20, "This monospaced message comes with a horizontal scroll bar." & vbLf, True)
    tMsg.section(3).sLabel = "Please note the following:"
    tMsg.section(3).sText = "The message form height is adjusted to the required height limited by the specified percentage of the screen height, " & _
                            "which for this test is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vButtons = vButton1 & "," & vButton2 & "," & vButton3
    
    MonoSpacedMessageSectionExceedMaxFormHeight = _
    mMsg.Msg( _
             title:=sMsgTitle, _
             message:=tMsg, _
             buttons:=vButtons _
            )
End Function

Private Function ErrSrc(ByVal s As String) As String
    ErrSrc = "mTest." & s
End Function

' Repeat the string (pattern) n (ntimes) times, otionally with a linenumber,
' either prefixed (linenumbersprefix=True) or attached. When the pattern
' ends with a vbLf, vbCr, or vbCrLf the attached line number is put at the
' left. The string withlinebreak is attached to the assembled pattern.
' -------------------------------------------------------------------------
Private Function Repeat(ByVal ntimes As Long, _
                        ByVal pattern As String, _
               Optional ByVal withlinenumbers As Boolean = False, _
               Optional ByVal linenumbersprefix As Boolean = True, _
               Optional ByVal withlinebreak As String = vbNullString) As String
    
    Const MAX_STRING_LENGTH = 12000
    Dim i       As Long
    Dim s       As String
    Dim ln      As String
    Dim sFormat As String
    
    On Error Resume Next
    If withlinenumbers Then sFormat = String$(Len(CStr(ntimes)), "0")
    
    For i = 1 To ntimes
        If withlinenumbers Then ln = Format(i, sFormat)
        If linenumbersprefix Then
            s = s & ln & " " & pattern & withlinebreak
        Else
            s = s & pattern & " " & ln & withlinebreak
        End If
        If Err.Number <> 0 Then
            Debug.Print "Repeate had to stop after " & i & "which resulted in a string length of " & Len(s)
            Repeat = s
            Exit Function
        End If
    Next i
    Repeat = s
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

Public Function ButtonByValue()

    Const PROC  As String = "ButtonByValue"
    Dim tMsg     As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    With fMsg
        .TestFrameWithBorders = True
        .TestFrameWithCaptions = True
'        .VmarginFrames = 5
'        .HmarginFrames = 6
    End With
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The ""buttons"" argument is provided as VB MsgBox value vbYesNo."
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The buttons ""Yes"" an ""No"" are displayed centered in one row"

    ButtonByValue = _
    mMsg.Msg( _
             title:="Test: Button by value (" & PROC & ")", _
             message:=tMsg, _
             buttons:=vbYesNo _
            )
End Function

Public Function ButtonByString()

    Const PROC  As String = "ButtonByString"
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The ""buttons"" argument is provided as string expression."
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    
    ButtonByString = _
    mMsg.Msg( _
             title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             message:=tMsg, _
             buttons:="Yes," & vbLf & ",No" _
            )
End Function

Public Function ButtonByCollection()

    Const PROC  As String = "ButtonByCollection"
    Dim cll     As New Collection
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    cll.Add "Yes"
    cll.Add "No"
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The ""buttons"" argument is provided as string expression."
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"

    ButtonByCollection = _
    mMsg.Msg( _
             title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             message:=tMsg, _
             buttons:=cll _
            )
End Function

Public Function ButtonByDictionary()
' -----------------------------------------------
' The buttons argument is provided as Dictionary.
' -----------------------------------------------
    Const PROC  As String = "ButtonByDictionary"
    Dim dct     As New Collection
    Dim tMsg    As tMessage
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The ""buttons"" argument is provided as string expression."
    tMsg.section(1).sLabel = "Expected result:"
    tMsg.section(1).sText = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    dct.Add "Yes"
    dct.Add "No"
    
    ButtonByDictionary = _
    mMsg.Msg( _
             title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             message:=tMsg, _
             buttons:=dct _
            )
End Function


Public Function Test_ButtonScrollBarVertical()

    Const PROC      As String = "Test_ButtonScrollBarVertical"
    Dim sButtons    As String
    Dim i, j        As Long
    Dim tMsg        As tMessage
    Dim cll         As New Collection
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
        .TestFrameWithBorders = True
        .MaxFormHeightPrcntgOfScreenSize = 60 ' enforce vertical scroll bar
    End With
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                           "the specified maximum forms height (for this test limited to " & _
                           fMsg.MaxFormHeightPrcntgOfScreenSize & "% of the screen height"
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The height for the vertically ordered buttons is reduced to fit the specified " & _
                           "maximum message form heigth and a vertical scroll bar is applied."
    tMsg.section(3).sLabel = "Finish test:"
    tMsg.section(3).sText = "This test is repeated with any button clicked othe than the ""Ok"" button"

    For i = 1 To 5
        For j = 0 To 1
            cll.Add "Reply" & vbLf & "Button" & vbLf & i + j
        Next j
        cll.Add vbLf
    Next i
    cll.Add "Ok"
    
    While mMsg.Msg( _
                   title:=Readable(PROC), _
                   message:=tMsg, _
                   buttons:=cll _
                  ) <> "Ok"
    Wend
    
End Function

Public Function Test_ButtonScrollBarHorizontal()

    Const PROC      As String = "ButtonScrollBarHorizontal"
    Dim sButtons    As String
    Dim i           As Long
    Dim tMsg        As tMessage
    Dim cll         As New Collection
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                            "the specified maximum forms width (for this test limited to " & _
                            fMsg.MaxFormHeightPrcntgOfScreenSize & "% of the screen height"
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The width for the horizontally ordered buttons is reduced to fit the specified " & _
                            "maximum message form width and a horizontal scroll bar is applied."
    tMsg.section(3).sLabel = "Finish test:"
    tMsg.section(3).sText = "This test is repeated with any button clicked othe than the ""Ok"" button"

    For i = 1 To 6
        cll.Add "Reply Button " & i
    Next i
    cll.Add "Ok"
    
    Do
        With fMsg
'            .TestFrameWithBorders = True
'            .TestFrameWithCaptions = True
'            .VmarginFrames = 5
'            .HmarginFrames = 6
            .MaxFormWidthPrcntgOfScreenSize = 50 ' enforce horizontal scroll bar
        End With

        If mMsg.Msg( _
             title:=Readable(PROC), _
             message:=tMsg, _
             buttons:=cll _
            ) = "Ok" Then Exit Do
    Loop
    
End Function

' By nature this test has become quite complex because default values, usually unchanged,
' are optionally adjusted by means of this "alternative MsgBox".
' In practice the constraints tested will become rarely effective. However, it is one
' of the major differences compared with the VB MsgBox that there is absolutely no message
' size limit - other than the VB limit for a string lenght.
' ----------------------------------------------------------------------------------------
Public Sub AllInOne()

    Dim lB1, lB2, lB3, lB4, lB5, lB6, lB7 As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMessage
    Dim cll                     As New Collection
    Dim vReply                  As Variant
    Dim lChangeHeightPcntg      As Long
    Dim lChangeWidthPcntg       As Long
    Dim lChangeMinWidthPt       As Long
    Dim bMonospaced             As Boolean: bMonospaced = True ' initial test value
    Dim lMinFormWidth           As Long
    Dim lMaxFormWidth           As Long
    Dim lMaxFormHeight          As Long
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test (lTest) from the Test Worksheet
    lTest = 5
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(lTest):   lChangeMinWidthPt = .MinFormWidthIncrDecr(lTest)
        lMaxFormWidth = .InitMaxFormWidth(lTest):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(lTest)
        lMaxFormHeight = .InitMaxFormHeight(lTest): lChangeHeightPcntg = .MaxFormHeightIncrDecr(lTest)
    End With
    
    sTitle = "All-in-1-Test: Combines as much behaviour a possible"
    tMsg.section(1).sLabel = "Test Description:"
    tMsg.section(1).sText = "This test specifically focuses on constraint issues." & vbLf & _
                            "The test environment allows to increase/decrease the maximum and minimm form width and height " & _
                            "in order to test what happens when the message and/or the buttons area's width and height " & _
                            "exceed the specified limits."
    tMsg.section(2).sLabel = "Test Results:"
    tMsg.section(3).sLabel = "Please note:"
    tMsg.section(3).bMonspaced = True
    
    '~~ Assemble the buttons argument as collection
    cll.Add "Increase ""Minimum Width"" by " & lChangeMinWidthPt & "pt":    lB1 = cll.Count
    cll.Add "Decrease ""Minimum Width"" by " & lChangeMinWidthPt & "pt":    lB2 = cll.Count
    cll.Add vbLf
    cll.Add "Increase ""Maximum Width"" by " & lChangeWidthPcntg & "%":     lB3 = cll.Count
    cll.Add "Decrease ""Maximum Width"" by " & lChangeWidthPcntg & "%":     lB4 = cll.Count
    cll.Add vbLf
    cll.Add "Increase ""Maximum Height"" by " & lChangeHeightPcntg & "%":   lB5 = cll.Count
    cll.Add "Decrease ""Maximum Height"" by " & lChangeHeightPcntg & "%":   lB6 = cll.Count
    cll.Add vbLf
    cll.Add "Finished":                                                     lB7 = cll.Count
    
    Do
        '~~ Assign initial - and as the test repeats the changed - values (contraints)
        '~~ for this test to the UserForm's properties
        With fMsg
            .MinFormWidth = lMinFormWidth
            .MaxFormWidthPrcntgOfScreenSize = lMaxFormWidth    ' for this demo to enforce a vertical scroll bar
            .MaxFormHeightPrcntgOfScreenSize = lMaxFormHeight  ' for this demo to enbforce a vertical scroll bar for the message section
'            .TestFrameWithBorders = True ' Just during test helpfull
        End With
        
        tMsg.section(2).sText = "When the specified minimum form width (currently " & lMinFormWidth & "pt) is increased, the form height will decrease because the proportional spaced " & _
                                "message section will require less height." & vbLf & _
                                "When the specified maximum width is reduced, the monospaced message section below and also the buttons area may get a horizontal scroll-bar." & vbLf & _
                                "When the specified maximum height is reduced, the message area and/or the buttons area may get a vertical scroll bar." & vbLf & _
                                "When the maximum is squeezed enough the scroll-bars may be applied alltogether."
        tMsg.section(3).sText = "- The specified maximum form width (currently " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% = " & Format(fMsg.MaxFormWidth, "##0") & "pt) cannot become less " & vbLf & _
                                "  than the minimum form width (currently " & fMsg.MinFormWidthPrcntg & "% = " & fMsg.MinFormWidth & "pt.) it may thus may have been limited automatically." & vbLf & _
                                "  (it should be noted that the maximum constraints are a percentage value while" & vbLf & _
                                "   the minimum width setting is in pt)" & vbLf & _
                                "- This longest line of this section determines the width of the displayed form" & vbLf & _
                                "  limited by specified maximum width (currently " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% = " & Format(fMsg.MaxFormWidth, "##0") & "pt)" & vbLf & _
                                "- In may take some time to understand the change of the displayed message" & vbLf & _
                                "  depending on the changed contraint values."
                 
        vReply = mMsg.Msg( _
                          title:=sTitle, _
                          message:=tMsg, _
                          buttons:=cll _
                         )
        With fMsg
            Select Case vReply
                Case cll(lB1): lMinFormWidth = lMinFormWidth + lChangeMinWidthPt
                Case cll(lB2): lMinFormWidth = lMinFormWidth - lChangeMinWidthPt
                Case cll(lB3): lMaxFormWidth = lMaxFormWidth + lChangeWidthPcntg
                Case cll(lB4): lMaxFormWidth = lMaxFormWidth - lChangeWidthPcntg
                Case cll(lB5): lMaxFormHeight = lMaxFormHeight + lChangeHeightPcntg
                Case cll(lB6): lMaxFormHeight = lMaxFormHeight - lChangeHeightPcntg
                Case cll(lB7): Exit Do ' The very last item in the collection is the "Finished" button
                Case Else
            End Select
        End With
    Loop
   
End Sub

Public Sub ButtonsMatrix()

    Dim i, j                    As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMessage
    Dim cllMatrix               As New Collection
    Dim cllStory                As New Collection
    Dim vReply                  As Variant
    Dim lChangeHeightPcntg      As Long
    Dim lChangeWidthPcntg       As Long
    Dim lChangeMinWidthPt       As Long
    Dim bMonospaced             As Boolean: bMonospaced = True ' initial test value
    Dim lMinFormWidth           As Long
    Dim lMaxFormWidth           As Long
    Dim lMaxFormHeight          As Long
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test (lTest) from the Test Worksheet
    lTest = 6
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(lTest):   lChangeMinWidthPt = .MinFormWidthIncrDecr(lTest)
        lMaxFormWidth = .InitMaxFormWidth(lTest):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(lTest)
        lMaxFormHeight = .InitMaxFormHeight(lTest): lChangeHeightPcntg = .MaxFormHeightIncrDecr(lTest)
    End With
    
    sTitle = "Buttons only test: No message, just buttons (finish with ""Ok"")"
    tMsg.section(1).sText = "Some can play around with button matrix of 7 by 7 buttons"
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 1 To 5 ' rows
        For j = 1 To 7 ' row buttons
            If j * i < 49 Then
                cllMatrix.Add "Button" & vbLf & i & "-" & j
            Else
                cllMatrix.Add "Next"
            End If
        Next j
        cllMatrix.Add vbLf
    Next i
    For i = 1 To 6
        cllMatrix.Add "Button" & vbLf & "6-" & i
    Next i
    cllMatrix.Add vbLf
    cllMatrix.Add "Ok"
    
    Do
        '~~ Assign initial - and as the test repeats the changed - values (contraints)
        '~~ for this test to the UserForm's properties
        With fMsg
            .MinButtonWidth = 40
            .MinFormWidth = lMinFormWidth
            .MaxFormWidthPrcntgOfScreenSize = lMaxFormWidth    ' for this demo to enforce a vertical scroll bar
            .MaxFormHeightPrcntgOfScreenSize = lMaxFormHeight  ' for this demo to enbforce a vertical scroll bar for the message section
            .TestFrameWithBorders = True
        End With
                         
        vReply = mMsg.Msg( _
                          title:=sTitle, _
                          message:=tMsg, _
                          buttons:=cllMatrix _
                         )
        With fMsg
            Select Case vReply
                Case "Ok": Exit Do ' The very last item in the collection is the "Finished" button
            End Select
        End With
    Loop
   
End Sub


Public Sub MostlyButtons()

    Dim i, j                    As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMessage
    Dim cllStory                As New Collection
    Dim vReply                  As Variant
    Dim lChangeHeightPcntg      As Long
    Dim lChangeWidthPcntg       As Long
    Dim lChangeMinWidthPt       As Long
    Dim bMonospaced             As Boolean: bMonospaced = True ' initial test value
    Dim lMinFormWidth           As Long
    Dim lMaxFormWidth           As Long
    Dim lMaxFormHeight          As Long
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test (lTest) from the Test Worksheet
    lTest = 6
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(lTest):   lChangeMinWidthPt = .MinFormWidthIncrDecr(lTest)
        lMaxFormWidth = .InitMaxFormWidth(lTest):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(lTest)
        lMaxFormHeight = .InitMaxFormHeight(lTest): lChangeHeightPcntg = .MaxFormHeightIncrDecr(lTest)
    End With
    
    sTitle = "Buttons only test: No message, just buttons (finish with ""Ok"")"
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    For i = 1 To 6 ' rows
        cllStory.Add "Click this button in case ...." & vbLf & "(no lengthy message text above but everything is said in the button)"
        cllStory.Add vbLf
    Next i
    cllStory.Add "Ok"
    
    Do
        '~~ Assign initial - and as the test repeats the changed - values (contraints)
        '~~ for this test to the UserForm's properties
        With fMsg
            .MinButtonWidth = 40
            .MinFormWidth = lMinFormWidth
            .MaxFormWidthPrcntgOfScreenSize = lMaxFormWidth    ' for this demo to enforce a vertical scroll bar
            .MaxFormHeightPrcntgOfScreenSize = lMaxFormHeight  ' for this demo to enbforce a vertical scroll bar for the message section
'            .TestFrameWithBorders = True
        End With
                         
        vReply = mMsg.Msg( _
                          title:=sTitle, _
                          message:=tMsg, _
                          buttons:=cllStory _
                         )
        With fMsg
            Select Case vReply
                Case "Ok": Exit Do ' The very last item in the collection is the "Finished" button
            End Select
        End With
    Loop
   
End Sub

Public Sub RepeatTest()
    Debug.Print Repeat(10, "a", True, False, vbLf)
End Sub

' Convert a string (s) into a readable form by replacing all underscores
' with a whitespace and all characters immediately following an underscore
' to a lowercase letter.
' ---------------------------------------------------------------------
Private Function Readable(ByVal s As String) As String

    Dim i       As Long
    Dim sResult As String
    
    s = Replace(s, "_", " ")
    s = Replace(s, "  ", " ")
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

Public Sub Test_WidthDeterminedByReplyButtons()
    WidthDeterminedByReplyButtons
End Sub
