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
Const BTTN_FINISH   As String = "Test Done"
Const BTTN_NEXT     As String = "Next Test"
Const BTTN_PREVIOUS As String = "Previous Test"

Dim lMinFormWidth   As Long
Dim lWidthIncrDecr  As Long
Dim lHeightIncrDecr As Long
Dim sMsgTitle       As String
Dim vbuttons        As Variant
Dim vButton4        As Variant
Dim vButton5        As Variant
Dim vButton6        As Variant
Dim vButton7        As Variant
Dim Message         As TypeMsg
Dim RegressionTest  As Boolean

Private Property Get BTTN_TERMINATE() As String ' composed constant
    BTTN_TERMINATE = "Terminate this" & vbLf & "Regression-Test"
End Property

Private Property Get ErrSrc(Optional ByVal s As String) As String:  ErrSrc = "mTest." & s:  End Property

' -------------------------------------------------------------------------------------------------
' Procedures for test start via Command Buttons on Test Worksheet
Public Sub cmdTest1_Click()
    RegressionTest = False
    mTest.Test_01_WidthDeterminedByMinimumWidth
End Sub

Public Sub cmdTest2_Click()
    RegressionTest = False
    mTest.Test_02_WidthDeterminedByTitle
End Sub

Public Sub cmdTest3_Click():   mTest.Test_03_WidthDeterminedByMonoSpacedMessageSection:    End Sub

Public Sub cmdTest4_Click():   mTest.Test_04_WidthDeterminedByReplyButtons:                End Sub

Public Sub cmdTest5_Click():   mTest.Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth:     End Sub

Public Sub cmdTest6_Click():   mTest.Test_06_MonoSpacedMessageSectionExceedMaxMsgHeight:   End Sub

Public Sub cmdTest7_Click():   mTest.Test_08_MostlyButtons:                                End Sub

Public Sub cmdTest8_Click():   mTest.Test_09_ButtonsMatrix:                                End Sub

Public Sub cmdTest9_Click():   mTest.Demonstration:                                        End Sub
' -------------------------------------------------------------------------------------------------

Public Function Demonstration() As Variant
' ------------------------------------------------------------------------------
' Demo as test of as many features as possible at once.
' ------------------------------------------------------------------------------
    Const MAX_WIDTH     As Long = 60
    Const MAX_HEIGHT    As Long = 55

    Dim sTitle          As String
    Dim cll             As New Collection
    Dim i, j            As Long
    Dim Message         As TypeMsg
   
    sTitle = "Demonstration (all in one)"
    With Message.Section(1)
        .Label.Text = "Displayed message summary "
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "- Display of (all) 4 message sections, each with an (optional) label" & vbLf _
                   & "- One monospaced section text exceeding the specified maximum width" & vbLf _
                   & "- Display of some of the 49(7x7) possible reply buttons" & vbLf _
                   & "- Font options like color, bold, and italic"
    End With
    With Message.Section(2)
        .Label.Text = "Unlimited message width"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "This section's text is mono-spaced and thus not word-wrapped. I.e. the longest line determines the messag width." & vbLf _
                   & "Because the maximimum width for this demo has been specified " & MAX_WIDTH & "% of the sreen width (defaults to 80%)" & vbLf _
                   & "the text is displayed with a horizontal scrollbar. The size limit for a section's text is only limited by VBA" & vbLf _
                   & "which as about 1GB! (see also unlimited message height below)"
        .Text.Monospaced = True
        .Text.FontItalic = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "All the message together would ecxeed the maximum height, spcified for this demo " & MAX_HEIGHT & "%" & vbLf _
                   & "of the sreen height (defaults to 70%) it is displayed with a vertical scrollbar. So no matter " _
                   & "how much text is displayed, it is never truncated. The only limit is VBA's limit for a text " _
                   & "string which is abut 1GB! With 4 strings, each in one section the limit is thus about 4GB !!!!"
    End With
    With Message.Section(4)
        .Label.Text = "Reply buttons flexibility"
        .Label.FontColor = rgbBlue
        .Label.FontBold = True
        .Text.Text = "This demo displays only some of the 49 possible reply buttons (7 rows by 7 buttons). " _
                   & "It also shows that a reply button can have any caption text and the buttons can be " _
                   & "displayed in any order within the 7 x 7 limit. Of cource the VBA.MsgBox classic " _
                   & "vbOkOnly, vbYesNoCancel, etc. are also possible - even in a mixture." & vbLf & vbLf _
                   & "By the way: This demo ends only with the Ok button clicked and loops with all the ohter."
    End With
    '~~ Prepare the buttons collection
    For j = 1 To 1
        For i = 1 To 5
            cll.Add "Sample multiline" & vbLf & "reply button" & vbLf & "Button-" & j & "-" & i
        Next i
        cll.Add vbLf
    Next j
    cll.Add vbOKOnly ' The reply when clicked will be vbOK though
    
    While mMsg.Dsply(dsply_title:=sTitle _
                   , dsply_msg:=Message _
                   , dsply_buttons:=cll _
                   , dsply_max_height:=MAX_HEIGHT _
                   , dsply_max_width:=MAX_WIDTH _
                    ) <> vbOK
    Wend
    
End Function

Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString)
' ------------------------------------------------------
' This Common Component does not have its own error
' handling. Instead it passes on any error to the
' caller's error handling.
' ------------------------------------------------------
    
    If err_no = 0 Then err_no = Err.Number
    If err_dscrptn = vbNullString Then
        On Error Resume Next ' Err.Description may not be available
        err_dscrptn = Err.Description
        If err_dscrptn = vbNullString Then err_dscrptn = "No error description provided by the system!"
    End If
    Debug.Print "Error in: "; err_source & ": Error = " & err_no & " " & err_dscrptn
    Err.Raise Number:=err_no, Source:=err_source, Description:=err_dscrptn

End Sub

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

Private Sub MessageInit()
' ------------------------------------------------------
' Initializes the all message sections with the defaults
' throughout this test module which uses a module global
' declared Message for a consistent layout.
' ------------------------------------------------------
    Dim i   As Long
    For i = 1 To fMsg.NoOfDesignedMsgSects
        With Message.Section(i)
            .Label.Text = vbNullString
            .Label.FontColor = rgbBlue
            .Text.Text = vbNullString
            .Text.Monospaced = False
            .Text.FontItalic = False
            .Text.FontUnderline = False
            .Text.FontColor = rgbBlack
        End With
    Next i
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

Public Sub Regression()
' --------------------------------------------------------------------------------------
' Regression testing makes use of all available design means - by the way testing them.
' Note: Each test procedure is completely independant and thus may be executed directly.
' --------------------------------------------------------------------------------------
    Const PROC = "Regression"
    
    On Error GoTo eh
    ThisWorkbook.Save
    Unload fMsg
    RegressionTest = True
    
1:  Select Case mTest.Test_01_WidthDeterminedByMinimumWidth
        Case BTTN_TERMINATE:    Exit Sub
    End Select

2:  Select Case mTest.Test_02_WidthDeterminedByTitle
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 2
    End Select

3:  Select Case mTest.Test_03_WidthDeterminedByMonoSpacedMessageSection
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 3
    End Select

4:  Select Case mTest.Test_04_WidthDeterminedByReplyButtons
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 3
    End Select

5:  Select Case mTest.Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 4
    End Select

6:  Select Case mTest.Test_06_MonoSpacedMessageSectionExceedMaxMsgHeight
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 5
    End Select
    
7: Select Case mTest.Test_17_MessageAsString
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 6
    End Select
    
8: Select Case mTest.Demonstration
        Case BTTN_TERMINATE:    Exit Sub
        Case BTTN_PREVIOUS:     GoTo 7
    End Select
    

xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Private Function RepeatString( _
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
        If Err.Number <> 0 Then
            Debug.Print "Repeate had to stop after " & i & "which resulted in a string length of " & Len(s)
            RepeatString = s
            Exit Function
        End If
    Next i
    RepeatString = s
End Function

Public Sub RepeatTest()
    Debug.Print RepeatString(10, "a", True, False, vbLf)
End Sub

Public Function Test_00_The_Buttons_Service_1() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    Set cll = mMsg.Buttons("B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", "B25", "B26", "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B34", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42", "B43", "B44", "B45", "B46", "B47", "B48", "B49", "B50")
    Test_00_The_Buttons_Service_1 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=cll)

End Function

Public Function Test_00_The_Buttons_Service_2() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Dim cll As Collection
    
    Set cll = mMsg.Buttons(2843, "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10", "B11", "B12", "B13", "B14", "B15", "B16", "B17", "B18", "B19", "B20", "B21", "B22", "B23", "B24", "B25", "B26", "B27", "B28", "B29", "B30", "B31", "B32", "B33", "B34", "B35", "B36", "B37", "B38", "B39", "B40", "B41", "B42", "B43", "B44", "B45", "B46", "B47", "B48", "B49", "B50")
    Test_00_The_Buttons_Service_2 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=cll)

End Function

Public Function Test_00_The_Buttons_Service_3() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Test_00_The_Buttons_Service_3 = _
    mMsg.Box(box_title:="49 buttons ordered in 7 rows, row breaks are inserted by the Buttons service, excessive 50th button ignored)", _
             box_buttons:=mMsg.Buttons("B01,B02,B03,B04,B05"))
End Function

Public Function Test_01_WidthDeterminedByMinimumWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_01_WidthDeterminedByMinimumWidth"
    Const TEST_NO   As Long = 1
    
    On Error GoTo eh
       
    ' Parameters for this test obtained from the Test Worksheet
    fMsg.MinMsgWidthPts = wsMsg.InitMinFormWidth(TEST_NO)
    lWidthIncrDecr = wsMsg.MsgWidthIncrDecr(TEST_NO)
    lHeightIncrDecr = wsMsg.MsgHeightIncrDecr(TEST_NO)
    
    vButton4 = "Repeat with" & vbLf & "minimum width" & vbLf & "+ " & lWidthIncrDecr
    vButton5 = "Repeat with" & vbLf & "minimum width" & vbLf & "- " & lWidthIncrDecr
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4, vButton5) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vButton4, vButton5) _
    
repeat:
    With fMsg
'        .TestFrameWithBorders = True
'        .FramesWithCaption = True
    End With
        
    sMsgTitle = Readable(PROC)
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsMsg.TestDescription(TEST_NO)
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The width of all message sections is adjusted to the current specified minimum form width (" & fMsg.MinMsgWidthPts & " pt)."
    End With
    With Message.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height up to the specified " & _
                     "maximum heigth which is " & fMsg.MaxMsgHeightPrcntgOfScreenSize & "% and not exceeded."
        .Text.FontColor = rgbRed
    End With
                                                                                              
    Test_01_WidthDeterminedByMinimumWidth = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_msg:=Message, _
             dsply_buttons:=vbuttons _
            )
    Select Case Test_01_WidthDeterminedByMinimumWidth
        Case vButton5
            fMsg.MinMsgWidthPts = wsMsg.InitMinFormWidth(TEST_NO) - lWidthIncrDecr
            If RegressionTest _
            Then Set vbuttons = mMsg.Buttons(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
            Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton4)
            GoTo repeat
        Case vButton4
            fMsg.MinMsgWidthPts = wsMsg.InitMinFormWidth(TEST_NO) + lWidthIncrDecr
            If RegressionTest _
            Then Set vbuttons = mMsg.Buttons(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
            Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton5)
            GoTo repeat
        Case Else ' Stop and Next are passed on to the caller
    End Select

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_02_WidthDeterminedByTitle() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_02_WidthDeterminedByTitle"
    Const TEST_NO       As Long = 2
    
    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    
    '~~ Initial test values obtained from the Test Worksheet
    lWidthIncrDecr = wsMsg.MsgWidthIncrDecr(TEST_NO)
    With fMsg
        .MinMsgWidthPts = wsMsg.InitMinFormWidth(TEST_NO)
'        .TestFrameWithBorders = True
    End With
    
    sMsgTitle = Readable(PROC) & "  (This title uses more space than the minimum specified message form width and thus the width is determined by the title)"
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsMsg.TestDescription(TEST_NO)
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the title's lenght."
    End With
    With Message.Section(3)
        .Label.Text = "Please note:"
        .Text.Text = "The two message sections in this test do use a proportional font " & _
                     "and thus are adjusted to form width determined by other factors." & vbLf & _
                     "The message form height is ajusted to the need up to the specified " & _
                     "maximum heigth based on the sreen height which for this test is " & _
                     fMsg.MaxMsgHeightPrcntgOfScreenSize & "%."
    End With
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH)
    
    Test_02_WidthDeterminedByTitle = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_msg:=Message, _
             dsply_buttons:=vbuttons _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_03_WidthDeterminedByMonoSpacedMessageSection() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC          As String = "Test_03_WidthDeterminedByMonoSpacedMessageSection"
    Const TEST_NO       As Long = 3
    
    On Error GoTo eh
    
    '~~ Initial test values obtained from the Test Worksheet
    lWidthIncrDecr = wsMsg.MsgWidthIncrDecr(TEST_NO)
    lHeightIncrDecr = wsMsg.MsgHeightIncrDecr(TEST_NO)
    
    ' Initializations for this test
    fMsg.MaxMsgWidthPrcntgOfScreenSize = wsMsg.InitMaxMsgWidth(TEST_NO)
    
    vButton4 = "Repeat with" & vbLf & "maximum width" & vbLf & "+ " & lWidthIncrDecr
    vButton5 = "Repeat with" & vbLf & "maximum width" & vbLf & "- " & lWidthIncrDecr
    vButton6 = "Repeat with" & vbLf & "maximum height" & vbLf & "+ " & lHeightIncrDecr
    vButton7 = "Repeat with" & vbLf & "maximum height" & vbLf & "- " & lHeightIncrDecr
    
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5, vButton6, vButton7) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vButton5, vButton6, vButton7)

    sMsgTitle = Readable(PROC)
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = wsMsg.TestDescription(TEST_NO)
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "Initally, the message form width is adjusted to the longest line in the " & _
                     "monospaced message section and all other message sections are adjusted " & _
                     "to this (enlarged) width." & vbLf & _
                     "When the maximum form width is reduced by " & lWidthIncrDecr & " % the monospaced message section is displayed with a horizontal scrollbar."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "- In contrast to the message sections above, this section uses the ""monospaced"" option which ensures" & vbLf & _
                            "  the message text is not ""wrapped around""." & vbLf & _
                            "- The message form height is ajusted to the need up to the specified maximum heigth" & vbLf & _
                            "  based on the sreen height which for this test is " & fMsg.MaxMsgHeightPrcntgOfScreenSize & "%."
        .Text.Monospaced = True
    End With
    Do
        With fMsg
'            .TestFrameWithCaptions = True  ' defaults to false, set to true for test purpose only
'            .TestFrameWithBorders = True  ' defaults to false, set to true for test purpose only
        End With
        Test_03_WidthDeterminedByMonoSpacedMessageSection = _
        mMsg.Dsply( _
                   dsply_title:=sMsgTitle, _
                   dsply_msg:=Message, _
                   dsply_buttons:=vbuttons _
                )
        Select Case Test_03_WidthDeterminedByMonoSpacedMessageSection
            Case vButton5
                fMsg.MaxMsgWidthPrcntgOfScreenSize = wsMsg.InitMaxMsgWidth(TEST_NO) - lWidthIncrDecr
                If RegressionTest _
                Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
                Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton4)
            Case vButton4
                fMsg.MaxMsgWidthPrcntgOfScreenSize = wsMsg.InitMaxMsgWidth(TEST_NO) + lWidthIncrDecr
                If RegressionTest _
                Then Set vbuttons = mMsg.Buttons(BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
                Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton5)
            Case Else: Exit Do ' Stop, Previous, and Next are passed on to the caller
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_04_WidthDeterminedByReplyButtons() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_04_WidthDeterminedByReplyButtons"
    
    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
'    fMsg.TestFrameWithBorders = True
    
    sMsgTitle = Readable(PROC)
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the reply buttons determines the width of the message form - unless they does not exceed the specified maximum form width which for this test is " & fMsg.MaxMsgWidthPts & " (which is the specified " & fMsg.MaxMsgWidthPrcntgOfScreenSize & "% of the sreen width)."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The message form width is adjusted to the space required by the number of reply buttons and all message sections are adjusted to this (enlarged) width."
    End With
    With Message.Section(3)
        .Label.Text = "Please also note:"
        .Text.Text = "The message form height is adjusted to the required height limited only by the specified maximum heigth " & _
                     "which is a percentage of the sreen height (for this test = " & fMsg.MaxMsgHeightPrcntgOfScreenSize & "%."
    End With
    vButton4 = "Repeat with 5 buttons"
    vButton5 = "Repeat with 4 buttons"
    
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4, vButton5, vButton6) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton4, vButton5, vButton6)
    
    Do
        Test_04_WidthDeterminedByReplyButtons = _
        mMsg.Dsply( _
                   dsply_title:=sMsgTitle, _
                   dsply_msg:=Message, _
                   dsply_buttons:=vbuttons _
                  )
        
        Select Case Test_04_WidthDeterminedByReplyButtons
            Case vButton4
                If RegressionTest _
                Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton5) _
                Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton5)
            Case vButton5
                If RegressionTest _
                Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT, vbLf, vButton4) _
                Else Set vbuttons = mMsg.Buttons(BTTN_FINISH, vbLf, vButton4)
            Case Else: Exit Do ' passed on to caller
        End Select
    Loop

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth"
    
    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    
    ' Initializations for this test
    With fMsg
'        .TestFrameWithBorders = True
        .MaxMsgWidthPrcntgOfScreenSize = 50
    End With
    
    sMsgTitle = Readable(PROC)
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the 3rd ""monospaced"" message section exceeds the maximum form width which for this test is " & fMsg.MaxMsgWidthPts & " pt (the equivalent of " & fMsg.MaxMsgWidthPrcntgOfScreenSize & "% of the sreen width)."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = "The monospaced message section comes with a horizontal scrollbar."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "- This monspaced message section exceeds the specified maximum form width which for this test is " & fMsg.MaxMsgWidthPts & " pt, " & _
                     "  which is the equivalent of " & fMsg.MaxMsgWidthPrcntgOfScreenSize & "% of the sreen width."
        .Text.Monospaced = True
    End With
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH)
    
    Test_05_MonoSpacedSectionWidthExceedsMaxMsgWidth = _
    mMsg.Dsply( _
             dsply_title:=sMsgTitle, _
             dsply_msg:=Message, _
             dsply_buttons:=vbuttons _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_06_MonoSpacedMessageSectionExceedMaxMsgHeight() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC      As String = "Test_06_MonoSpacedMessageSectionExceedMaxMsgHeight"

    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    
    ' Initializations for this test
    
    sMsgTitle = Readable(PROC)
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width used by the monospaced message section exxceeds the maximum form width for this test (" & fMsg.MaxMsgWidthPts & ") which is the specified " & fMsg.MaxMsgWidthPrcntgOfScreenSize & "% of the sreen width."
    End With
    With Message.Section(3)
        .Label.Text = "Please note the following:"
        .Text.Text = "The message form height is adjusted to the required height limited by the specified percentage of the screen height, " & _
                     "which for this test is " & fMsg.MaxMsgHeightPrcntgOfScreenSize & "%."
    End With
    With Message.Section(2)
        .Label.Text = "Expected test result:"
        .Text.Text = RepeatString(25, "This monospaced message comes with a horizontal scrollbar." & vbLf, True)
        .Text.Monospaced = True
    End With
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH)
    
    Test_06_MonoSpacedMessageSectionExceedMaxMsgHeight = _
    mMsg.Dsply( _
               dsply_max_width:=80, _
               dsply_max_height:=70, _
               dsply_title:=sMsgTitle, _
               dsply_msg:=Message, _
               dsply_buttons:=vbuttons _
              )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Sub Test_08_MostlyButtons()
    Const PROC              As String = "Test_08_MostlyButtons"
    Const TEST_NO           As Long = 8
    
    On Error GoTo eh
    Dim i                   As Long
    Dim sTitle              As String
    Dim cllStory            As New Collection
    Dim vReply              As Variant
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim lMinFormWidth       As Long
    Dim lMaxMsgWidth        As Long
    Dim lMaxMsgHeight       As Long
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test (TEST_NO) from the Test Worksheet
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(TEST_NO):   lChangeMinWidthPt = .MsgWidthIncrDecr(TEST_NO)
        lMaxMsgWidth = .InitMaxMsgWidth(TEST_NO):   lChangeWidthPcntg = .MsgWidthIncrDecr(TEST_NO)
        lMaxMsgHeight = .InitMaxMsgHeight(TEST_NO): lChangeHeightPcntg = .MsgHeightIncrDecr(TEST_NO)
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
            .MinMsgWidthPts = lMinFormWidth
            .MaxMsgWidthPrcntgOfScreenSize = lMaxMsgWidth    ' for this demo to enforce a vertical scrollbar
            .MaxMsgHeightPrcntgOfScreenSize = lMaxMsgHeight  ' for this demo to enbforce a vertical scrollbar for the message section
'            .TestFrameWithBorders = True
        End With
                         
        vReply = mMsg.Dsply( _
                          dsply_title:=sTitle, _
                          dsply_msg:=Message, _
                          dsply_buttons:=cllStory _
                         )
        With fMsg
            Select Case vReply
                Case "Ok": Exit Do ' The very last item in the collection is the "Finished" button
            End Select
        End With
    Loop

xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Public Sub Test_09_ButtonsMatrix()
    Const PROC              As String = "Test_09_ButtonsMatrix"
    Const TEST_NO           As Long = 9
    
    On Error GoTo eh
    Dim i, j                As Long
    Dim sTitle              As String
    Dim cllMatrix           As Collection
    Dim vReply              As Variant
    Dim lChangeHeightPcntg  As Long
    Dim lChangeWidthPcntg   As Long
    Dim lChangeMinWidthPt   As Long
    Dim bMonospaced         As Boolean: bMonospaced = True ' initial test value
    Dim lMinFormWidth       As Long
    Dim lMaxMsgWidth        As Long
    Dim lMaxMsgHeight       As Long
    
    '~~ Obtain initial test values and their corresponding change (increment/decrement) value
    '~~ for this test (TEST_NO) from the Test Worksheet
    With wsMsg
        lMinFormWidth = .InitMinFormWidth(TEST_NO):   lChangeMinWidthPt = .MsgWidthIncrDecr(TEST_NO)
        lMaxMsgWidth = .InitMaxMsgWidth(TEST_NO):   lChangeWidthPcntg = .MsgWidthIncrDecr(TEST_NO)
        lMaxMsgHeight = .InitMaxMsgHeight(TEST_NO): lChangeHeightPcntg = .MsgHeightIncrDecr(TEST_NO)
    End With
    
    sTitle = "Just to demonstrate what's theoretically possible: Buttons only! (finish with ""Ok"")"
    MessageInit ' set test-global message specifications
'    Message.Section(1).Text.Text = "Some can play around with button matrix of 7 by 7 buttons"
    '~~ Assemble the matrix of buttons as collection for  the argument buttons
    Set cllMatrix = New Collection
    For i = 1 To 7 ' rows
        For j = 1 To 7 ' row buttons
            If (j * i) < 49 Then
                cllMatrix.Add "Button" & vbLf & i & "-" & j
            Else
                cllMatrix.Add vbOKOnly
            End If
        Next j
        If i < 7 Then cllMatrix.Add vbLf
    Next i
    
    Do
        '~~ Assign initial - and as the test repeats the changed - values (contraints)
        '~~ for this test to the UserForm's properties
        With fMsg
            .MinButtonWidth = 40
            .MinMsgWidthPts = lMinFormWidth
            .MaxMsgWidthPrcntgOfScreenSize = lMaxMsgWidth    ' for this demo to enforce a vertical scrollbar
            .MaxMsgHeightPrcntgOfScreenSize = lMaxMsgHeight  ' for this demo to enbforce a vertical scrollbar for the message section
'            .DsplyFrmsWthBrdrsTestOnly = True
'            .DsplyFrmsWthCptnTestOnly = True
        End With
                         
'        mMsg.Dsply dsply_title:=sTitle _
'                 , dsply_msg:=Message _
'                 , dsply_buttons:=cllMatrix _
'                 , dsply_reply_with_index:=True _
'                 , dsply_modeless:=False
'        Select Case mMsg.RepliedWith
'            Case vbOK: Exit Do
'            Case 49: Exit Do
'        End Select
    
        Select Case mMsg.Dsply(dsply_title:=sTitle _
                             , dsply_msg:=Message _
                             , dsply_buttons:=cllMatrix _
                             , dsply_reply_with_index:=True _
                             , dsply_button_default:="Ok" _
                             , dsply_modeless:=True)
            Case vbOK: Exit Do
            Case 49: Exit Do
        End Select
    Loop

xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

Public Function Test_10_ButtonScrollBarVertical()
    Const PROC = "Test_10_ButtonScrollBarVertical"
    
    Dim i, j        As Long
    Dim cll         As New Collection
    
    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
        .MaxMsgHeightPrcntgOfScreenSize = 60 ' enforce vertical scrollbar
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The number of the used reply ""buttons"", their specific order respectively exceeds " & _
                     "the specified maximum forms height (for this test limited to " & _
                     fMsg.MaxMsgHeightPrcntgOfScreenSize & "% of the screen height"
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The height for the vertically ordered buttons is reduced to fit the specified " & _
                     "maximum message form heigth and a vertical scrollbar is applied."
    End With
    With Message.Section(3)
        .Label.Text = "Finish test:"
        .Text.Text = "This test is repeated with any button clicked othe than the ""Ok"" button"
    End With
    For i = 1 To 5
        For j = 0 To 1
            cll.Add "Reply" & vbLf & "Button" & vbLf & i + j
        Next j
        cll.Add vbLf
    Next i
    cll.Add "Ok"
    
    While mMsg.Dsply( _
                   dsply_title:=Readable(PROC), _
                   dsply_msg:=Message, _
                   dsply_buttons:=cll _
                  ) <> "Ok"
    Wend

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_11_ButtonScrollBarHorizontal()

    Const PROC = "ButtonScrollBarHorizontal"
    Const INIT_WIDTH = 50
    Const CHANGE_WIDTH = 10
    
    On Error GoTo eh
    Dim Bttn10Plus  As String
    Dim Bttn10Minus As String
    Dim BttnOk      As Variant: BttnOk = vbOKOnly
    Dim WidthMax    As Long
    
    WidthMax = INIT_WIDTH

repeat:
    Unload fMsg                                         ' Ensures a message starts from scratch
    fMsg.MaxMsgWidthPrcntgOfScreenSize = WidthMax ' enforce horizontal scrollbar
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The width, number, amd order of the displayed reply buttons exceeds " & _
                     "the specified maximum forms width, which for this test is limited to " & _
                     fMsg.MaxMsgWidthPrcntgOfScreenSize & "% of the screen width."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "When the maximum form width is enlarged the message will be displayed without a scrollbar."
    End With
    With Message.Section(3)
        .Label.Text = "Finish test:"
        .Text.Text = "This test is repeated with any button clicked othe than the ""Ok"" button"
    End With
    
    Bttn10Plus = "Repeat with maximum form width" & vbLf & "extended by " & CHANGE_WIDTH & "% to " & WidthMax + CHANGE_WIDTH & "%"
    Bttn10Minus = "Repeat with maximum form width" & vbLf & "reduced by " & CHANGE_WIDTH & "% to " & WidthMax - CHANGE_WIDTH & "%"
        
    With fMsg
'        .DsplyFrmsWthBrdrsTestOnly = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 5
'        .HmarginFrames = 6
    End With

    Select Case mMsg.Dsply( _
                           dsply_title:=Readable(PROC) _
                         , dsply_msg:=Message _
                         , dsply_buttons:=mMsg.Buttons(Bttn10Plus, Bttn10Minus, BttnOk))
        Case Bttn10Minus:   WidthMax = WidthMax - CHANGE_WIDTH:     GoTo repeat
        Case Bttn10Plus:    WidthMax = WidthMax + CHANGE_WIDTH:     GoTo repeat
        Case BttnOk:                                                GoTo xt
    End Select

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_13_ButtonByValue()

    Const PROC  As String = "Test_13_ButtonByValue"
    
    On Error GoTo eh
    Unload fMsg                     ' Ensures a message starts from scratch
    
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 5
'        .HmarginFrames = 6
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as VB MsgBox value vbYesNo."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in one row"
    End With
    Test_13_ButtonByValue = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & PROC & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=vbOKOnly _
            )
            
xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_14_ButtonByString()

    Const PROC  As String = "Test_14_ButtonByString"
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_14_ButtonByString = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:="Yes," & vbLf & ",No" _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_15_ButtonByCollection()

    Const PROC  As String = "Test_15_ButtonByCollection"
    Dim cll     As New Collection
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    cll.Add "Yes"
    cll.Add "No"
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    Test_15_ButtonByCollection = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=cll _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_16_ButtonByDictionary()
' -----------------------------------------------
' The buttons argument is provided as Dictionary.
' -----------------------------------------------
    Const PROC  As String = "Test_16_ButtonByDictionary"
    Dim dct     As New Collection
    
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test description:"
        .Text.Text = "The ""buttons"" argument is provided as string expression."
    End With
    With Message.Section(2)
        .Label.Text = "Expected result:"
        .Text.Text = "The buttons ""Yes"" an ""No"" are displayed centered in two rows"
    End With
    dct.Add "Yes"
    dct.Add "No"
    
    Test_16_ButtonByDictionary = _
    mMsg.Dsply( _
             dsply_title:="Test: Button by value (" & ErrSrc(PROC) & ")", _
             dsply_msg:=Message, _
             dsply_buttons:=dct _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Function Test_17_MessageAsString() As Variant
' ------------------------------------------------------------------------------
'
' ------------------------------------------------------------------------------
    Const PROC  As String = "Test_17_Box_MessageAsString"
        
    Unload fMsg                     ' Ensures a message starts from scratch
    With fMsg
'        .TestFrameWithBorders = True
'        .TestFrameWithCaptions = True
'        .VmarginFrames = 2
'        .HmarginFrames = 5
    End With
    
    If RegressionTest _
    Then Set vbuttons = mMsg.Buttons(BTTN_PREVIOUS, BTTN_TERMINATE, BTTN_NEXT) _
    Else Set vbuttons = mMsg.Buttons(BTTN_FINISH)
        
    sMsgTitle = Readable(PROC)
    
    Test_17_MessageAsString = _
    mMsg.Box( _
             box_title:=sMsgTitle _
           , box_msg:="This is a message provided as a simple string argument!" _
           , box_buttons:=vbuttons _
            )

xt: Exit Function

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Function

Public Sub Test_20_Progress_Messages()

    Dim i As Long
    Dim Msg As TypeMsg
    With Msg.Section(1).Text
        .Text = " No. Status  Step"
        .Monospaced = True
    End With
    For i = 1 To 30
        With Msg.Section(2).Text
            .Text = Align(i, 4, AlignRight, " ") & Align("Passed", 8, AlignCentered, " ") & Align(i, 2, AlignRight) & ". Follow-Up line"
            .Monospaced = True
        End With
             
        mMsg.Progress prgrs_title:="Follow-Up-Test" _
                    , prgrs_msg:=Msg _
                    , prgrs_section:=2
    Next i
    
End Sub

Public Sub Test_99_Individual()
' ---------------------------------------------------------------------------------
' Test 1 (optional arguments are used in conjunction with the Regression test only)
' ---------------------------------------------------------------------------------
    Const PROC = "Test_99_Individual"
    
    On Error GoTo eh
    With fMsg
'        .MinMsgWidthPts = 300
        .DsplyFrmsWthBrdrsTestOnly = True
'        .FramesWithCaption = True
    End With
    
    MessageInit ' set test-global message specifications
    With Message.Section(1)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text" & vbLf & _
                     "A short message text"
    End With
    With Message.Section(2)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text"
        .Text.Monospaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Test label extra long for this specific test:"
        .Text.Text = "A short message text"
    End With
    mMsg.Dsply dsply_title:="This title is rather short" _
             , dsply_msg:=Message _
             , dsply_buttons:=mMsg.Buttons("Button-1", "Button-2") _
             , dsply_button_default:="Button-1"
    
xt: Exit Sub

eh: ErrMsg ErrSrc(PROC)
#If Test Then
    Stop: Resume
#End If
End Sub

