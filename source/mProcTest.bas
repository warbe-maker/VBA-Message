Attribute VB_Name = "mProcTest"
Option Explicit

Public Sub Test_SizeWidthAndHeight_Width_Limited()
    Const PROC = "Test_SizeWidthAndHeight_Width_Limited"
    
    Dim TestWidth   As Single
    Dim TestHeight  As Single
    Dim i           As Long
    Dim iFrom       As Long:    iFrom = 400
    Dim iTo         As Long:    iTo = 100
    Dim iStep       As Long:    iStep = -100
    
again:
    With fProcTest
        .Top = 0
        .Left = 0
        .Show False
        TestHeight = 200
        For TestWidth = iFrom To iTo Step iStep
            i = i + 1
            .Caption = PROC
            .frm.Width = TestWidth + 3
            .frm.Left = 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .tbx.Left = 0
            .tbx.Top = 0
            
            .SizeWidthAndHeight as_tbx:=.tbx _
                              , as_width:=TestWidth _
                              , as_height:=TestHeight _
                              , as_text:="This " & i & ". test is with a width limited to " & TestWidth & _
                                         ". The width is regarded fixed and thus WordWrap = True is used to maintain/gurantee it. " & _
                                         "A provided height is regarded the minimum even when not used." _
                              , as_append:=True
            
            With .tbxTestAndResult
                With .Font
                    .Name = "Courier New"
                    .Size = 8
                End With
                .Top = 5
                .Width = 250
                .WordWrap = True
                .AutoSize = False
                .MultiLine = True
            End With
            .tbxTestAndResult = "tbx.Width  = " & .tbx.Width & vbLf & _
                                "tbx.Height = " & .tbx.Height & vbLf & _
                                "TestHeight = " & TestHeight
            .tbxTestAndResult.Height = 40
            
            .frm.Top = .tbxTestAndResult.Top + .tbxTestAndResult.Height + 5
            .frm.Height = .tbx.Top + .tbx.Height + .VspaceFrame(.frm)
            .Height = .frm.Top + .frm.Height + (.Height - .InsideHeight) + 5
            If TestWidth <> iTo Then
                If MsgBox(Title:="Continue? > Yes, Terminate? > No", Buttons:=vbYesNo, Prompt:=vbNullString) = vbNo Then
                    Unload fProcTest
                    Exit Sub
                End If
            Else
                If MsgBox(Title:="Done? > Yes, Repeat? > No", Buttons:=vbYesNo, Prompt:=vbNullString) = vbYes Then
                    Unload fProcTest
                    Exit Sub
                Else
                    Unload fProcTest
                    GoTo again
                End If
            End If
        Next TestWidth
    End With

End Sub

Public Sub Test_SizeWidthAndHeight_Width_Unlimited()
    Const PROC = "Test_SizeWidthAndHeight_Width_Unlimited"
    
    Dim TestWidth   As Single
    Dim TestHeight  As Single
    Dim i           As Long
    Dim iFrom       As Long:    iFrom = 1
    Dim iTo         As Long:    iTo = 5
    Dim iStep       As Long:    iStep = 1
    
again:
    With fProcTest
        .Top = 0
        .Left = 0
        TestHeight = 200
        For i = iFrom To iTo Step iStep
            .Caption = PROC
            .frm.Width = TestWidth + 3
            .frm.Left = 5
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .tbx.Left = 0
            .tbx.Top = 0
            
            .SizeWidthAndHeight as_tbx:=.tbx _
                              , as_width:=TestWidth _
                              , as_height:=TestHeight _
                              , as_text:="This " & i & ". test is with an unlimited width. " & _
                                         "The width is determined by the longest text line and WordWrap = False. " & _
                                         "A provided height is regarded the minimum even when not used." _
                              , as_append:=True
            
            With .tbxTestAndResult
                With .Font
                    .Name = "Courier New"
                    .Size = 8
                End With
                .Top = 5
                .Width = 250
                .WordWrap = True
                .AutoSize = False
                .MultiLine = True
            End With
            .tbxTestAndResult = "tbx.Width  = " & .tbx.Width & vbLf & _
                                "tbx.Height = " & .tbx.Height & vbLf & _
                                "TestHeight = " & TestHeight
            .tbxTestAndResult.Height = 40
            
            .frm.Top = .tbxTestAndResult.Top + .tbxTestAndResult.Height + 5
            .frm.Height = .tbx.Top + .tbx.Height + .VspaceFrame(.frm)
            .frm.Width = .tbx.Left + .tbx.Width + 3
            .Width = .frm.Left + .frm.Width + (.Width - .InsideWidth) + 5
            .Height = .frm.Top + .frm.Height + (.Height - .InsideHeight) + 5
            .Show False
            
            If TestWidth <> iTo Then
                If MsgBox(Title:="Continue? > Yes, Terminate? > No", Buttons:=vbYesNo, Prompt:=vbNullString) = vbNo Then
                    Unload fProcTest
                    Exit Sub
                End If
            Else
                If MsgBox(Title:="Done? > Yes, Repeat? > No", Buttons:=vbYesNo, Prompt:=vbNullString) = vbYes Then
                    Unload fProcTest
                    Exit Sub
                Else
                    Unload fProcTest
                    GoTo again
                End If
            End If
        Next i
    End With

End Sub


