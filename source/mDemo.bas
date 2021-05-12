Attribute VB_Name = "mDemo"
Option Explicit

Public Sub Demos()
    DemoMsgDsplyService_1
    DemoMsgDsplyService_2
End Sub

Public Sub DemoMsgDsplyService_1()
    Const MAX_WIDTH     As Long = 50
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
                   & "for this demo has been specified " & MAX_WIDTH & "% of the sreen width (the default would be 80%)" & vbLf _
                   & "the text is displayed with a horizontal scrollbar. There is no message size limit for the display despite the" & vbLf & vbLf _
                   & "limit of VBA for text strings  which is about 1GB!"
        .Text.Monospaced = True
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
                   , dsply_max_width:=MAX_WIDTH _
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
            .Monospaced = True ' Just to demonstrate
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

