Attribute VB_Name = "mDemo"
Option Explicit

Public Sub Demo_Msg()

   Dim sTitle   As String
   Dim tMsg     As tMessage
   Dim cll      As New Collection
    Dim i       As Long
    
    With fMsg
        .MaxFormWidthPrcntgOfScreenSize = 45    ' for this demo to enforce a vertical scroll bar
        .MaxFormHeightPrcntgOfScreenSize = 75   ' for this demo to enbforce a vertical scroll bar for the message section
    End With
   
    sTitle = "Usage demo: Full featured multiple choice message"
    tMsg.section(1).sLabel = "1. Demonstration:"
    tMsg.section(1).sText = "Use of all 3 message sections, all with a label and use of all 7 reply buttons, in a 2-2-2-1  order."
    tMsg.section(2).sLabel = "2. Demonstration:"
    tMsg.section(2).sText = "The impact of the specified maximimum message form with, which for this test has been reduced to " & _
                            fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size (the default is 80%)." & vbLf & vbLf & _
                            "Because this message section is very tall" & vbLf & _
                            "(for this demo specifically)" & vbLf & _
                            "the total message area's height exceeds the" & vbLf & _
                            "specified maximum message form height." & vbLf & _
                            "When it is reduced to its limit" & vbLf & _
                            "the whole message area is provided with a vertical scroll bar." & vbLf & vbLf & _
                            "By this, the alternative MsgBox has in fact no message size limit."
    tMsg.section(3).sLabel = "3. Demonstration:"
    tMsg.section(3).sText = "This part of the message demonstrates the mono-spaced option and" & vbLf & _
                            "the impact it has on the width of the message form, which is" & vbLf & _
                            "determined by its longest line because mono-spaced message sections " & vbLf & _
                            "are not ""word wrapped"". However, because the specified maximum message form width is exceed" & vbLf & _
                            "a vertical scroll bar is applied - in practice it hardly will ever happen." & vbLf & _
                            "I.e. even for a mono-spaced text section there is no width limit." & vbLf & vbLf & _
                            "Attention: The result is redisplayed until the ""Ok"" button is clicked!"
   
   '~~ Prepare the buttons collection
    For i = 1 To 6
        cll.Add "Multiline reply button caption" & vbLf & "Button-" & i
        cll.Add vbLf
    Next i
    cll.Add vbLf: cll.Add "Ok"
    
   Do
      With fMsg
'        .TestFrameWithBorders = True
      End With
      If mMsg.Msg( _
         title:=sTitle, _
        message:=tMsg, _
         buttons:=cll) _
      = cll(cll.Count) _
      Then Exit Do
   Loop
   
End Sub

