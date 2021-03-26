Attribute VB_Name = "mDemo"
Option Explicit

Public Sub FirstTry()
          
    With fMsg
        .MsgTitle = "Any title"
        .MsgText(1) = "Any message"
        .MsgButtons = vbYesNoCancel
        .Setup
        .show
        Select Case .ReplyValue ' obtaining it unloads the form !
            Case vbYes:     MsgBox "Button ""Yes"" clicked"
            Case vbNo:      MsgBox "Button ""No"" clicked"
            Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
        End Select
   End With
   
End Sub

Public Sub Demo_Dsply()

    Dim sTitle  As String
    Dim tMsg    As tMsg
    Dim cll     As New Collection
    Dim i, j    As Long
    
    With fMsg
        .MaxFormWidthPrcntgOfScreenSize = 45    ' for this demo to enforce a vertical scroll bar
        .MaxFormHeightPrcntgOfScreenSize = 75   ' for this demo to enbforce a vertical scroll bar for the message section
    End With
   
    sTitle = "Usage demo: Full featured multiple choice message"
    tMsg.Section(1).sLabel = "1. Demonstration:"
    tMsg.Section(1).sText = "Use of all 3 message sections, all with a label and use of all 7 reply buttons, in a 2-2-2-1  order."
    tMsg.Section(2).sLabel = "2. Demonstration:"
    tMsg.Section(2).sText = "The impact of the specified maximimum message form with, which for this test has been reduced to " _
                          & fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size (the default is 80%)." & vbLf & vbLf _
                          & "Because this message section is very tall (for this demo specifically) the total message " _
                          & "area's height exceeds the specified maximum message form height." & vbLf _
                          & "When it is reduced to its limit the whole message area is provided with a vertical scroll bar." & vbLf & vbLf & _
                            "By this, the alternative MsgBox has in fact no message size limit."
    tMsg.Section(3).sLabel = "3. Demonstration:"
    tMsg.Section(3).sText = "This part of the message demonstrates the mono-spaced option and the impact it " _
                          & "has on the width of the message form, which is determined by its longest line " _
                          & "because mono-spaced message sections are not ""word wrapped"". However, because " _
                          & "the specified maximum message form width is exceed a vertical scroll bar is applied " _
                          & "- in practice it hardly will ever happen. I.e. even for a mono-spaced text section " _
                          & "there is no width limit."
    tMsg.Section(4).sLabel = "Attention!"
    tMsg.Section(4).sText = "The result is re-displayed until the ""Ok"" button is clicked!"
   
   '~~ Prepare the buttons collection
   For j = 1 To 3
        For i = 1 To 3
            cll.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
        cll.Add vbLf
    Next j
    cll.Add "Ok"
    
    While mMsg.Dsply(msg_title:=sTitle, msg:=tMsg, msg_buttons:=cll, msg_min_width:=600) <> cll(cll.Count)
    Wend
    
End Sub

Public Sub Test_Dsply()
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
    
    Dim vMsg    As tMsg                      ' structure of the message
    Dim cll     As New Collection                   ' specification of the displayed buttons
    Dim vReturn As Variant
           
    ' Preparing the message
    With vMsg.Section(1)
        .sLabel = "Any label 1"
        .sText = "Any section text 1"
    End With
    With vMsg.Section(2)
        .sLabel = "Any label 2"
        .sText = "Any section 2 text"
        .bMonspaced = True ' Just to demonstrate
    End With
    With vMsg.Section(3)
        .sLabel = "Any label 3"
        .sText = "Any section text 3"
   End With
       
   vReturn = Dsply(msg_title:="Any title", _
                   msg:=vMsg, _
                   msg_buttons:=mMsg.Buttons(vbAbortRetryIgnore, vbLf, B1, B2, B3, vbLf, B4, B5, B6, vbLf, B7) _
                  )
   MsgBox "Button """ & mMsg.ReplyString(vReturn) & """ had been clicked"
   
End Sub
