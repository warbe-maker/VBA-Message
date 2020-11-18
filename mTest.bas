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
' Please note:
' Errors raised by the tested procedures cannot be
' asserted since they are not passed on to the calling
' /entry procedure. This would require the Common
' Standard Module mErrHndlr which intentionally is not
' used by this module.
'
' W. Rauschenberger, Berlin June 2020
' -----------------------------------------------------
Const BTTN_FINISH       As String = "Test Done"
Const BTTN_NEXT         As String = "Next Test"
Const BTTN_PREVIOUS     As String = "Previous Test"

Dim lMinFormWidth   As Long
Dim sMsgTitle       As String
Dim sMsgLabel       As String
Dim sMsg1Label      As String
Dim sMsg2Label      As String
Dim sMsg3Label      As String
Dim sMsgText        As String
Dim sMsg1Text       As String
Dim sMsg2Text       As String
Dim sMsg3Text       As String
Dim vbuttons        As Variant
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

Private Property Get BTTN_TERMINATE() As String ' composed constant
    BTTN_TERMINATE = "Terminate this" & vbLf & "Regression-Test"
End Property

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mTest." & s:  End Property

Private Function ButtonsCollection( _
      ParamArray vbuttons() As Variant) As Collection
' ---------------------------------------------------
' Returns a collection of provided strings.
' ---------------------------------------------------
    Dim cll As New Collection
    Dim i As Long
    
    For i = LBound(vbuttons) To UBound(vbuttons)
        If vbuttons(i) <> vbNullString Then cll.Add vbuttons(i)
    Next i
    Set ButtonsCollection = cll
    
End Function

Function IsUcase(ByVal s As String) As Boolean

    Dim i   As Integer: i = Asc(s)
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

Public Sub Regression()
' --------------------------------------------------------------------------------------
' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
                                     
    ThisWorkbook.Save
    Unload fMsg
    
1:  Select Case mTest.Test_01_WidthDeterminedByMinimumWidth(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
    End Select

2:  Select Case mTest.Test_02_WidthDeterminedByTitle(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 2
    End Select

3:  Select Case mTest.Test_03_WidthDeterminedByMonoSpacedMessageSection(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 3
    End Select

4:  Select Case mTest.Test_04_WidthDeterminedByReplyButtons(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 3
    End Select

5:  Select Case mTest.Test_05_MonoSpacedSectionWidthExceedsMaxFormWidth(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 4
    End Select

6:  Select Case mTest.Test_06_MonoSpacedMessageSectionExceedMaxFormHeight(regression_test:=True)
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 5
    End Select
    
End Sub

Private Function Repeat( _
           ByVal rep_n_times As Long, _
           ByVal rep_pattern As String, _
  Optional ByVal rep_with_line_numbers As Boolean = False, _
  Optional ByVal rep_with_linen_umbers_as_prefix As Boolean = True, _
  Optional ByVal rep_with_with_line_breaks As String = vbNullString) As String
' ------------------------------------------------------------------------------
' Repeat the string (rep_pattern) n (rep_n_times) times, otionally with a line-
' number, either prefixed (linenumbersprefix=True) or attached. When the pattern
' ends with a vbLf, vbCr, or vbCrLf the attached line number is put at the left.
' The string rep_with_with_line_breaks is attached to the assembled rep_pattern.
' ------------------------------------------------------------------------------
    
    Const MAX_STRING_LENGTH = 12000
    Dim i       As Long
    Dim s       As String
    Dim ln      As String
    Dim sFormat As String
    
    On Error Resume Next
    If rep_with_line_numbers Then sFormat = String$(Len(CStr(rep_n_times)), "0")
    
    For i = 1 To rep_n_times
        If rep_with_line_numbers Then ln = Format(i, sFormat)
        If rep_with_linen_umbers_as_prefix Then
            s = s & ln & " " & rep_pattern & rep_with_with_line_breaks
        Else
            s = s & rep_pattern & " " & ln & rep_with_with_line_breaks
        End If
        If err.Number <> 0 Then
            Debug.Print "Repeate had to stop after " & i & "which resulted in a string length of " & Len(s)
            Repeat = s
            Exit Function
        End If
    Next i
    Repeat = s
End Function

Public Sub RepeatTest()
    Debug.Print Repeat(10, "a", True, False, vbLf)
End Sub

Public Function Test_01_WidthDeterminedByMinimumWidth( _
       Optional regression_test As Boolean = False) As Variant
' ---------------------------------------------------------------------------------
' Test 1 (optional arguments are used in conjunction with the Regression test only)
' ---------------------------------------------------------------------------------
    Const PROC          As String = "Test_01_WidthDeterminedByMinimumWidth"
    Const TEST_NO       As Long = 1
    
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMsg
       
    ' Initializations for this test
    fMsg.MinFormWidth = wsMsg.InitMinFormWidth(TEST_NO)
    
    vButton4 = "Repeat with" & vbLf & "minimum width" & vbLf & "+ " & lIncrDecrWidth
    vButton5 = "Repeat with" & vbLf & "minimum width" & vbLf & "- " & lIncrDecrWidth
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4, vButton5) _
    Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vButton4, vButton5) _

    
Repeat:
    With fMsg
'        .TestFrameWithBorders = True
'        .FramesWithCaption = True
    End With
    
    sMsgTitle = Readable(PROC)
    tMsg.section(1).sLabel = "Test description:":       tMsg.section(1).sText = wsMsg.TestDescription(TEST_NO)
    tMsg.section(2).sLabel = "Expected test result:":   tMsg.section(2).sText = "The width of all message sections is adjusted to the current specified minimum form width (" & fMsg.MinFormWidth & " pt)."
    tMsg.section(3).sLabel = "Please also note:":       tMsg.section(3).sText = "The message form height is ajusted to the need " & _
                                                                                "up to the specified maximum heigth which is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% and not exceeded."
    Test_01_WidthDeterminedByMinimumWidth = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_message:=tMsg, _
             dsply_buttons:=vbuttons _
            )
    Select Case Test_01_WidthDeterminedByMinimumWidth
        Case vButton5
            fMsg.MinFormWidth = wsMsg.InitMinFormWidth(TEST_NO) - lIncrDecrWidth
            If regression_test _
            Then Set vbuttons = ButtonsCollection(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
            Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton4)
            GoTo Repeat
        Case vButton4
            fMsg.MinFormWidth = wsMsg.InitMinFormWidth(TEST_NO) + lIncrDecrWidth
            If regression_test _
            Then Set vbuttons = ButtonsCollection(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
            Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton5)
            GoTo Repeat
        Case Else ' Stop and Next are passed on to the caller
    End Select
    
End Function

Public Function Test_02_WidthDeterminedByTitle( _
       Optional regression_test As Boolean = False) As Variant
' --------------------------------------------------------------------------------------------------
' Test 2 (optional arguments are used in conjunction with the Regression test only)
' --------------------------------------------------------------------------------------------------
    Const PROC          As String = "Test_02_WidthDeterminedByTitle"
    Const TEST_NO       As Long = 2
    
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecrWidth = wsMsg.MinFormWidthIncrDecr(TEST_NO)
    With fMsg
        .MinFormWidth = wsMsg.InitMinFormWidth(TEST_NO)
'        .TestFrameWithBorders = True
    End With
    
    sMsgTitle = Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    tMsg.section(1).sLabel = "Test description:":       tMsg.section(1).sText = wsMsg.TestDescription(TEST_NO)
    tMsg.section(2).sLabel = "Expected test result:":   tMsg.section(2).sText = "The message form width is adjusted to the title's lenght."
    tMsg.section(3).sLabel = "Please note:":            tMsg.section(3).sText = "The two message sections in this test do use a proportional font " & _
                                                                                "and thus are adjusted to form width determined by other factors." & vbLf & _
                                                                                "The message form height is ajusted to the need up to the specified " & _
                                                                                "maximum heigth based on the screen size which for this test is " & _
                                                                                fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    
    Test_02_WidthDeterminedByTitle = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_message:=tMsg, _
             dsply_buttons:=vbuttons _
            )
End Function

Public Function Test_03_WidthDeterminedByMonoSpacedMessageSection( _
       Optional regression_test As Boolean = False) As Variant
' -------------------------------------------------------------------------------------
' Test 3 (optional arguments are used in conjunction with the Regression test only)
' -------------------------------------------------------------------------------------
    Const PROC          As String = "Test_03_WidthDeterminedByMonoSpacedMessageSection"
    Const TEST_NO       As Long = 3
    
    Dim lIncrDecrHeight As Long
    Dim lIncrDecrWidth  As Long
    Dim tMsg            As tMsg
    

    '~~ Initial test values obtained from the Test Worksheet
    lIncrDecrWidth = wsMsg.MaxFormWidthIncrDecr(TEST_NO)
    lIncrDecrHeight = wsMsg.MaxFormHeightIncrDecr(TEST_NO)
    
    ' Initializations for this test
    fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(TEST_NO)
    
    vButton4 = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & lIncrDecrWidth
    vButton5 = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & lIncrDecrWidth
    vButton6 = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & lIncrDecrHeight
    vButton7 = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & lIncrDecrHeight
    
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5, vButton6, vButton7) _
    Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vButton5, vButton6, vButton7)

    sMsgTitle = Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = wsMsg.TestDescription(TEST_NO)
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
'            .TestFrameWithCaptions = True  ' defaults to false, set to true for test purpose only
'            .TestFrameWithBorders = True  ' defaults to false, set to true for test purpose only
        End With
        Test_03_WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Dsply( _
                   dsply_title:=sMsgTitle, _
                   dsply_message:=tMsg, _
                   dsply_buttons:=vbuttons _
                )
        Select Case Test_03_WidthDeterminedByMonoSpacedMessageSection
            Case vButton5
                fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(TEST_NO) - lIncrDecrWidth
                If regression_test _
                Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
                Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton4)
            Case vButton4
                fMsg.MaxFormWidthPrcntgOfScreenSize = wsMsg.InitMaxFormWidth(TEST_NO) + lIncrDecrWidth
                If regression_test _
                Then Set vbuttons = ButtonsCollection(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
                Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton5)
            Case Else: Exit Do ' Stop, Previous, and Next are passed on to the caller
        End Select
    Loop
    
End Function

Public Function Test_04_WidthDeterminedByReplyButtons( _
       Optional regression_test As Boolean = False) As Variant
' ---------------------------------------------------------------------------------
' Test 4 (optional arguments are used in conjunction with the Regression test only)
' ---------------------------------------------------------------------------------
    Const PROC      As String = "WidthDeterminedByReplyButtons1"
    Const TEST_NO   As Long = 4
    
    Dim tMsg        As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
'    fMsg.TestFrameWithBorders = True
    
    sMsgTitle = Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " (which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    tMsg.section(3).sLabel = "Please also note:"
    tMsg.section(3).sText = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                            "which is a percentage of the screen size (for this test = " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    vButton4 = "Repeat with" & vbLf & "5 buttons"
    vButton5 = "Repeat with" & vbLf & "4 buttons"
    
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4, vButton5, vButton6) _
    Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton4, vButton5, vButton6)
    
    Do
        Test_04_WidthDeterminedByReplyButtons = _
        mMsg.Dsply( _
                   dsply_title:=sMsgTitle, _
                   dsply_message:=tMsg, _
                   dsply_buttons:=vbuttons _
                  )
        
        Select Case Test_04_WidthDeterminedByReplyButtons
            Case vButton4
                If regression_test _
                Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
                Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton5)
            Case vButton5
                If regression_test _
                Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
                Else Set vbuttons = ButtonsCollection(BTTN_FINISH, vbLf, vButton4)
            Case Else: Exit Do ' passed on to caller
        End Select
    Loop
    
End Function

Public Function Test_05_MonoSpacedSectionWidthExceedsMaxFormWidth( _
       Optional regression_test As Boolean = False) As Variant
' -----------------------------------------------------------------------------
' Test 5 (optional arguments are used in conjunction with Regression test only)
' -----------------------------------------------------------------------------
    Const PROC      As String = "Test_05_MonoSpacedSectionWidthExceedsMaxFormWidth"
    Const TEST_NO   As Long = 5
    
    Dim tMsg        As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    With fMsg
'        .TestFrameWithBorders = True
        .MaxFormWidthPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & fMsg.MaxFormWidth & " pt (the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size)."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = "The monospaced message section comes with a horizontal scroll bar."
    tMsg.section(3).sLabel = "Please note the following:"
    tMsg.section(3).sText = "- This monspaced message section exceeds the specified maximum form width which for this test is " & fMsg.MaxFormWidth & " pt," & vbLf & _
                            "  the equivalent of " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size." & vbLf & _
                            "- The message form height is adjusted to the required height, limited to " & fMsg.MaxFormHeight & " pt," & vbLf & _
                            "  the equivalent of " & fMsg.MaxFormHeightPrcntgOfScreenSize & "% of the screen size, for this test and not reached or exceeded."
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    Else Set vbuttons = ButtonsCollection(BTTN_FINISH)
    
    Test_05_MonoSpacedSectionWidthExceedsMaxFormWidth = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_message:=tMsg, _
             dsply_buttons:=vbuttons _
            )
End Function

Public Function Test_06_MonoSpacedMessageSectionExceedMaxFormHeight( _
       Optional regression_test As Boolean = False) As Variant
' -----------------------------------------------------------------------------
' Test 6 (optional arguments are used in conjunction with Regression test only)
' -----------------------------------------------------------------------------

    Const PROC      As String = "Test_06_MonoSpacedMessageSectionExceedMaxFormHeight"
    Const TEST_NO   As Long = 6

    Dim tMsg    As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    
    sMsgTitle = Readable(PROC)
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The width used by the monospaced message section exxceeds the maximum form width for this test (" & fMsg.MaxFormWidth & ") which is the specified " & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size."
    tMsg.section(2).sLabel = "Expected test result:"
    tMsg.section(2).sText = Repeat(20, "This monospaced message comes with a horizontal scroll bar." & vbLf, True)
    tMsg.section(3).sLabel = "Please note the following:"
    tMsg.section(3).sText = "The message form height is adjusted to the required height limited by the specified percentage of the screen height, " & _
                            "which for this test is " & fMsg.MaxFormHeightPrcntgOfScreenSize & "%."
    If regression_test _
    Then Set vbuttons = ButtonsCollection(BTTN_PREVIOUS, BTTN_TERMINATE) _
    Else Set vbuttons = ButtonsCollection(BTTN_FINISH)
    
    Test_06_MonoSpacedMessageSectionExceedMaxFormHeight = _
    mMsg.Dsply( _
               dsply_max_width:=80, _
               dsply_max_height:=50, _
               dsply_title:=sMsgTitle, _
               dsply_message:=tMsg, _
               dsply_buttons:=vbuttons _
              )
              
End Function

Public Sub Test_07_AllInOne()
' ----------------------------------------------------------------------------------------
' By nature this test has become quite complex because default values, usually unchanged,
' are optionally adjusted by means of this "alternative MsgBox".
' In practice the constraints tested will become rarely effective. However, it is one
' of the major differences compared with the VB MsgBox that there is absolutely no message
' size limit - other than the VB limit for a string lenght.
' ----------------------------------------------------------------------------------------
    Const PROC                  As String = "Test_07_AllInOne"
    Const TEST_NO               As Long = 7
    
    Dim lB1, lB2, lB3, lB4, lB5, lB6, lB7 As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMsg
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
    '~~ for this test (TEST_NO) from the Test Worksheet
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(TEST_NO):   lChangeMinWidthPt = .MinFormWidthIncrDecr(TEST_NO)
        lMaxFormWidth = .InitMaxFormWidth(TEST_NO):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(TEST_NO)
        lMaxFormHeight = .InitMaxFormHeight(TEST_NO): lChangeHeightPcntg = .MaxFormHeightIncrDecr(TEST_NO)
    End With
    
    sTitle = Readable(PROC) & ": Combines as much behaviour a possible"
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
            .Setup
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
                 
        vReply = mMsg.Dsply( _
                          dsply_title:=sTitle, _
                          dsply_message:=tMsg, _
                          dsply_buttons:=cll _
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

Public Sub Test_08_MostlyButtons()
    Const PROC                  As String = "Test_08_MostlyButtons"
    Const TEST_NO               As Long = 8
    
    Dim i, j                    As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMsg
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
    '~~ for this test (TEST_NO) from the Test Worksheet
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(TEST_NO):   lChangeMinWidthPt = .MinFormWidthIncrDecr(TEST_NO)
        lMaxFormWidth = .InitMaxFormWidth(TEST_NO):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(TEST_NO)
        lMaxFormHeight = .InitMaxFormHeight(TEST_NO): lChangeHeightPcntg = .MaxFormHeightIncrDecr(TEST_NO)
    End With
    
    sTitle = Readable(PROC) & ": No message, just buttons (finish with ""Ok"")"
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
                         
        vReply = mMsg.Dsply( _
                          dsply_title:=sTitle, _
                          dsply_message:=tMsg, _
                          dsply_buttons:=cllStory _
                         )
        With fMsg
            Select Case vReply
                Case "Ok": Exit Do ' The very last item in the collection is the "Finished" button
            End Select
        End With
    Loop
   
End Sub

Public Sub Test_09_ButtonsMatrix()
    Const PROC                  As String = "Test_09_ButtonsMatrix"
    Const TEST_NO               As Long = 9
    
    Dim i, j                    As Long
    Dim sTitle                  As String
    Dim tMsg                    As tMsg
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
    '~~ for this test (TEST_NO) from the Test Worksheet
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(TEST_NO):   lChangeMinWidthPt = .MinFormWidthIncrDecr(TEST_NO)
        lMaxFormWidth = .InitMaxFormWidth(TEST_NO):   lChangeWidthPcntg = .MaxFormWidthIncrDecr(TEST_NO)
        lMaxFormHeight = .InitMaxFormHeight(TEST_NO): lChangeHeightPcntg = .MaxFormHeightIncrDecr(TEST_NO)
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
'            .TestFrameWithBorders = True
        End With
                         
        vReply = mMsg.Dsply( _
                          dsply_title:=sTitle, _
                          dsply_message:=tMsg, _
                          dsply_buttons:=cllMatrix, _
                          dsply_returnindex:=True _
                         )
        Select Case vReply
            Case "Ok": Exit Do ' The very last item in the collection is the "Finished" button
            Case 42: Exit Do
        End Select
    Loop
   
End Sub

Public Function Test_10_ButtonScrollBarVertical()

    Const PROC      As String = "Test_10_ButtonScrollBarVertical"
    Dim sButtons    As String
    Dim i, j        As Long
    Dim tMsg        As tMsg
    Dim cll         As New Collection
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
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
    
    While mMsg.Dsply( _
                   dsply_title:=Readable(PROC), _
                   dsply_message:=tMsg, _
                   dsply_buttons:=cll _
                  ) <> "Ok"
    Wend
    
End Function

Public Function Test_11_ButtonScrollBarHorizontal()

    Const PROC      As String = "ButtonScrollBarHorizontal"
    Dim sButtons    As String
    Dim i           As Long
    Dim tMsg        As tMsg
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

        If mMsg.Dsply( _
             dsply_title:=Readable(PROC), _
             dsply_message:=tMsg, _
             dsply_buttons:=cll _
            ) = "Ok" Then Exit Do
    Loop
    
End Function

Public Function Test_13_ButtonByValue()

    Const PROC  As String = "Test_13_ButtonByValue"
    Dim tMsg     As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 5
'        .HmarginFrames = 6
    End With
    
    tMsg.section(1).sLabel = "Test description:"
    tMsg.section(1).sText = "The ""buttons"" argument is provided as VB MsgBox value vbYesNo."
    tMsg.section(2).sLabel = "Expected result:"
    tMsg.section(2).sText = "The buttons ""Yes"" an ""No"" are displayed centered in one row"

    Test_13_ButtonByValue = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & PROC & ")", _
             dsply_message:=tMsg, _
             dsply_buttons:=vbOKOnly _
            )
End Function

Public Function Test_14_ButtonByString()

    Const PROC  As String = "Test_14_ButtonByString"
    Dim tMsg    As tMsg
    
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
    
    Test_14_ButtonByString = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_message:=tMsg, _
             dsply_buttons:="Yes," & vbLf & ",No" _
            )
End Function

Public Function Test_15_ButtonByCollection()

    Const PROC  As String = "Test_15_ButtonByCollection"
    Dim cll     As New Collection
    Dim tMsg    As tMsg
    
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

    Test_15_ButtonByCollection = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_message:=tMsg, _
             dsply_buttons:=cll _
            )
End Function

Public Function Test_16_ButtonByDictionary()
' -----------------------------------------------
' The buttons argument is provided as Dictionary.
' -----------------------------------------------
    Const PROC  As String = "Test_16_ButtonByDictionary"
    Dim dct     As New Collection
    Dim tMsg    As tMsg
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
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
    
    Test_16_ButtonByDictionary = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_message:=tMsg, _
             dsply_buttons:=dct _
            )
End Function

