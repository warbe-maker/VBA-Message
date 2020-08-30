Attribute VB_Name = "mDemo"
Option Explicit

Public Sub Demo_Msg()

   Dim sTitle   As String
   Dim sLabel1  As String
   Dim sText1   As String
   Dim sLabel2  As String
   Dim sText2   As String
   Dim sLabel3  As String
   Dim sText3   As String
   Dim sButtons As String
   Dim sButton1 As String
   Dim sButton2 As String
   Dim sButton3 As String
   Dim sButton4 As String
   Dim sButton5 As String
   Dim sButton6 As String
   Dim sButton7 As String

   fMsg.MaxFormWidthPrcntgOfScreenSize = 45 ' for this demo to enforce a vertical scroll bar
   
   sTitle = "Usage demo: Full featured multiple choice message"
   sLabel1 = "1. Demonstration:"
   sText1 = "Use of all 3 message sections, all with a label and use of all 7 reply buttons, in a 2-2-2-1  order."
   sLabel2 = "2. Demonstration:"
   sText2 = "The impact of the specified maximimum message form with, which for this test has been reduced to " & _
            fMsg.MaxFormWidthPrcntgOfScreenSize & "% of the screen size (the default is 80%)."
   sLabel3 = "3. Demonstration:"
   sText3 = "This part of the message demonstrates the mono-spaced option and" & vbLf & _
            "the impact it has on the width of the message form, which is" & vbLf & _
            "determined by its longest line because mono-spaced message sections " & vbLf & _
            "are not ""word wrapped"". However, because the specified maximum message form width is exceed" & vbLf & _
            "a vertical scroll bar is applied - in practice it hardly will ever happen." & vbLf & _
            "I.e. even for a mono-spaced text section there is no width limit." & vbLf & vbLf & _
            "Attention: The result is redisplayed until the ""Ok"" button is clicked!"
   sButton1 = "Multiline reply button caption" & vbLf & "Button-1"
   sButton2 = "Multiline reply button caption" & vbLf & "Button-2"
   sButton3 = "Multiline reply button caption" & vbLf & "Button-3"
   sButton4 = "Multiline reply button caption" & vbLf & "Button-4"
   sButton5 = "Multiline reply button caption" & vbLf & "Button-5"
   sButton6 = "Multiline reply button caption" & vbLf & "Button-6"
   sButton7 = "Ok"
   '~~ Assemble the buttons argument string
   sButtons = _
   sButton1 & "," & sButton2 & "," & vbLf & "," & _
   sButton3 & "," & sButton4 & "," & vbLf & "," & _
   sButton4 & "," & sButton5 & "," & vbLf & "," & sButton7 _

   Do
      With fMsg
        .TestFrameWithBorders = True
      End With
      If mMsg.Msg( _
         title:=sTitle, _
         label1:=sLabel1, text1:=sText1, _
         label2:=sLabel2, text2:=sText2, _
         label3:=sLabel3, text3:=sText3, _
         monospaced3:=True, _
         buttons:=sButtons) _
      = sButton7 _
      Then Exit Do
   Loop
   
End Sub

