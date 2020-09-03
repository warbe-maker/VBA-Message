Attribute VB_Name = "mUsage"
Option Explicit

Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       Section(1 To 3) As tSection  ' sections
End Type

Public Function Box( _
                    ByVal title As String, _
                    ByVal prompt As String, _
           Optional ByVal buttons As Variant = vbOKOnly _
                   ) As Variant
          
   With fMsg
      .ApplTitle = title
      .ApplText(1) = prompt
      .ApplButtons = buttons
      .Setup
      .Show
      Box = .ReplyValue ' obtaining the reply value unloads the form !
   End With
   
End Function

Public Sub FirstTry()
          
    With fMsg
        .ApplTitle = "Any title"
        .ApplText(1) = "Any message"
        .ApplButtons = vbYesNoCancel
        .Setup
        .Show
        Select Case .ReplyValue ' obtaining the reply value unloads the form !
            Case vbYes:     MsgBox "Button ""Yes"" clicked"
            Case vbNo:      MsgBox "Button ""No"" clicked"
            Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
        End Select
   End With
   
End Sub

Public Function Msg( _
                    ByVal title As String, _
                    ByRef message As tMessage, _
           Optional ByVal buttons As Variant = vbOKOnly _
                   ) As Variant

   With fMsg
      .ApplTitle = title
      .ApplMsg = message
      .ApplButtons = buttons
      .Setup
      .Show
      Msg = .ReplyValue ' obtaining the reply value unloads the form !
   End With

End Function

Public Sub Test_Box()
    Select Case Box("Any title", "Any message", buttons:=vbYesNoCancel)
        Case vbYes:     MsgBox "Button ""Yes"" clicked"
        Case vbNo:      MsgBox "Button ""No"" clicked"
        Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
    End Select
End Sub

Public Sub Usage_Msg()
' ---------------------------------------------------------
' Displays a message with 3 sections, each with a label and
' 7 reply buttons ordered in rows 3-3-1
' ---------------------------------------------------------
    
    Dim tMsg    As tMessage                         ' structure of the message
    Dim cll     As New Collection                   ' specification of the button
    Dim iB1, iB2, iB3, iB4, iB5, iB6, iB7 As Long   ' indices for the return value
    
    ' Note that because of the "row breaks" the button indices are not equal the position in the collection
    cll.Add "Caption Button 1": iB1 = cll.Count
    cll.Add "Caption Button 2": iB2 = cll.Count
    cll.Add "Caption Button 3": iB3 = cll.Count
    cll.Add vbLf ' button row break (also with vbCr or vbCrLf)
    cll.Add "Caption Button 4": iB4 = cll.Count
    cll.Add "Caption Button 5": iB5 = cll.Count
    cll.Add "Caption Button 6": iB6 = cll.Count
    cll.Add vbLf ' button row break
    cll.Add "Caption Button 7": iB7 = cll.Count
       
    tMsg.Section(1).sLabel = "Label section 1"
    tMsg.Section(1).sText = "Message section 1 text"
    tMsg.Section(2).sLabel = "Label section 2"
    tMsg.Section(2).sText = "Message section 2 text"
    tMsg.Section(2).bMonspaced = True ' Just to demostrate
    tMsg.Section(3).sLabel = "Label section 3"
    tMsg.Section(3).sText = "Message section 3 text"
           
    Select Case Msg(title:="Any title", _
                   message:=tMsg, _
                   buttons:=cll)
        Case cll(iB1): MsgBox "Button with caption """ & cll(iB1) & """ clicked"
        Case cll(iB2): MsgBox "Button with caption """ & cll(iB2) & """ clicked"
        Case cll(iB3): MsgBox "Button with caption """ & cll(iB3) & """ clicked"
        Case cll(iB4): MsgBox "Button with caption """ & cll(iB4) & """ clicked"
        Case cll(iB5): MsgBox "Button with caption """ & cll(iB5) & """ clicked"
        Case cll(iB6): MsgBox "Button with caption """ & cll(iB6) & """ clicked"
        Case cll(iB7): MsgBox "Button with caption """ & cll(iB7) & """ clicked"
    End Select
   
End Sub
