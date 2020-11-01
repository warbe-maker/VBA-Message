# MsgBox Alternative

[Abstract](#abstract)<br>
[Why an alternative MsgBox](#why-an-alternative-msgbox)<br>[Installation](#installation)<br>[Properties of the _fMsg_ UserForm](#properties-of-the-fmsg-userform)<br>[Usage](#usage)<br>[Interfaces](#interfaces)

### Abstract
Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked.

### Why an alternative MsgBox?
The alternative implementation addresses many of the MsgBox's deficiencies - without re-implementing it to 100%.

|MsgBox|Alternative|
|------|-----------|
| The message width and height is limited and cannot be altered | The&nbsp;maximum&nbsp;width and&nbsp;height&nbsp;is&nbsp;specified as&nbsp;a percentage of the screen&nbsp;size&nbsp; which&nbsp;defaults&nbsp;to: 80%&nbsp;width and  90%&nbsp;height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it) may be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a satisfying result | Up&nbsp;to&nbsp;3&nbsp; _Message&nbsp;Sections_ each with an optional _Message Text Label_ and a _Monospaced_ option allow an appealing design without any extra  effort |
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order (=49 buttons in total) |
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| Specifying the default button | (yet) not implemented |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

### Installation
1. Download <a href="https://www.dropbox.com/s/h91lcqa52qrdl5f/fMsg.frm?dl=1">fMsg.frm</a> and the <a href="https://www.dropbox.com/s/h91lcqa52qrdl5f/fMsg.frm?dl=1">fMsg.frx</a>.
2. Import _fMsg.frm_ to a VBA project
3. In the VBE add a Reference to "Microsoft Scripting Runtime"
4. Copy the following code into a standard module's global declarations section:
5. Copy the following into a standard module:<br>
```
Public Enum StartupPosition         ' -------------------
    Manual = 0                      ' Used to position
    CenterOwner = 1                 ' the message window
    CenterScreen = 2                ' horizontally and
    WindowsDefault = 3              ' vertically centered
End Enum                            ' -------------------

Public Type tSection                ' --------------
       sLabel As String             ' Structure of
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' (fMsg) message
End Type                            ' area with its
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type

```
6. Before start have a look at the [UserForm's properties](#properties-of-the-fmsg-userform)
7. Either continue with [Usage step by step](#usage-step-by-step) or start directly using the prepared [Interfaces](#Interfaces) in module _mMsg_.  

#### Properties of the _fMsg_ UserForm

| Property | Meaning |
|----------|---------|
| _MsgTitle_| Mandatory. String expression. Applied in the message window's handle bar|
| _Msg_     | Optional. User defined type. Structure of the UserForm's message area. May alternatively to the properties _MsgLable_, _MsgText_, and _MsgMonoSpaced_ be used to pass a complete message.<br>See .... |
| _MsgLabel(n)_ | Optional. String expression with _n__ as a numeric expression 1 to 3. Applied as a descriptive label above a below message text. Not displayed (even when provided) when no corresponding _MsgText_ is provided |
| _MsgText(n)_ | Optional.String expression with _n__ as a numeric expression 1 to 3). Applied as message text of section _n_.|
| _MsgMonospaced(n)_ | Optional. Boolean expression. Defaults to False when omitted. When True, the text in section _n_ is displayed mono-spaced.|
| _MsgButtons_ | Optional. Defaults to vbOkOnly.<br>A MsgBox buttons value,<br>a comma delimited String expression,<br>a Collection,<br>or a dictionary,<br>with each item specifying a displayed command button's caption or a button row break (vbLf, vbCr, or vbCrLf)|
| _ReplyValue_ | Read only. The clicked button's caption string or [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>). When there is more than one button the form is unloaded when the clicked buttons value is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click.|
| _ReplyIndex_ | Read only. The clicked button's index. When there is more than one button the form is unloaded when the clicked button's index is fetched. When there is just one button this value will not be available since the form is immediately unloaded with the button click. |

See [Additional properties for advanced usage](<Implementation.md#public-properties-for-advanced-usage-of-the-message-form>) to create application specific messages.
### Usage
#### Usage step-by-step
##### A very first try
```
Public Sub FirstTry()
          
    With fMsg
        .ApplTitle = "Any title"
        .ApplText(1) = "Any message"
        .ApplButtons = vbYesNoCancel
        .Setup
        .Show
        Select Case .ReplyValue ' obtaining it unloads the form !
            Case vbYes:     MsgBox "Button ""Yes"" clicked"
            Case vbNo:      MsgBox "Button ""No"" clicked"
            Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
        End Select
   End With
End Sub
```
This example seems not being worth using the alternative.
However, when encapsulated in a function things look much better. Copy the following into a standard module:
```
Public Function Box(ByVal title As String, _
                    ByVal prompt As String, _
           Optional ByVal buttons As Variant = vbOKOnly) As Variant
          
   With fMsg
      .ApplTitle = title
      .ApplText(1) = prompt
      .ApplButtons = buttons
      .Setup
      .Show
      Box = .Reply ' obtaining the reply value unloads the form !
   End With
   
End Function
```
Displaying the message now looks pretty much the same as using MsgBox:
```
Public Sub Test_Box()
    Select Case Box(title:="Any title", _
                    prompt:="Any message", _
                    buttons:=vbYesNoCancel)
        Case vbYes:     MsgBox "Button ""Yes"" clicked"
        Case vbNo:      MsgBox "Button ""No"" clicked"
        Case vbCancel:  MsgBox "Button ""Cancel"" clicked"
    End Select
End Sub
```

#### The full buttons flexibility
The following general purpose message function provides all what the alternative offers.
```
Public Function Msg(ByVal title As String, _
                    ByRef message As tMessage, _
           Optional ByVal buttons As Variant = vbOKOnly, _
           Optional ByVal returnindex As Boolean = False) As Variant

   With fMsg
      .ApplTitle = title
      .ApplMsg = message
      .ApplButtons = buttons
      .Setup
      .Show
      '~~ Obtaining the reply value or index unloads the form !
      If returnindex Then Msg = .ReplyIndex Else Msg = .ReplyValue
   End With

End Function
```
Displaying a full featured message now looks as follows:
```
Public Sub Test_Msg()
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
       
    With tMsg.Section(1)
        .sLabel = "Label section 1"
        .sText = "Message section 1 text"
    End With
    With tMsg.Section(2)
        .sLabel = "Label section 2"
        .sText = "Message section 2 text"
        .bMonspaced = True ' Just to demonstrate
    End With
    With tMsg.Section(3)
        .sLabel = "Label section 3"
        .sText = "Message section 3 text"
   End With
       
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
```

### Interfaces
The [downloaded module _mMsg_](https://www.dropbox.com/s/ld30m8bz7gj9jgq/mMsg.bas?dl=1) provides three interfaces to the UserForm _fMsg_ and return the clicked reply _Button_ value to the caller.

#### _Box_ (see [example](#simple-message))
Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 7 reply _buttons_ in up to 7 rows in any order.
##### Syntax
```
mMsg.Box prompt[, buttons][, title]
```
or alternatively when the clicked reply button matters:
```
Select Case mMsg.Box(prompt[, buttons][, title])
   Case ....
   Case ....
End Select
```
The _Box_ function syntax has these named arguments:

| Part | Description | Corresponding _fMsg_ Property |
| ---- |-----------| --- |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. | Title |
| prompt | String expression displayed as message. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line. | Text(1) |
| [buttons](#syntax-of-the-buttons-argument) | Optional. Defaults to vbOkOnly when omitted. Variant expression, either a [VB MsgBox value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>), a comma delimited string, a collection of string expressions, or a dictionary of string expressions. In case of a string, a collection, or a dictionary, each item either specifies a button's caption (up to 7) or a reply button row break (vbLf, vbCr, or vbCrLf). | Buttons |

#### _Msg_ (see [example](#common-message))
Displays a message in up to 3 sections, each with an optional label and optionally monospaced and up to 7 buttons in up to 7 rows in any order.
##### Syntax
```
mMsg.Msg(title _
[[, label1][, text1][, monospaced1]] _
[[, label2][, text2][, monospaced2]] _
[[, label3][, text3][, monospaced3]] _
[,buttons])
```
The _Msg_ function syntax has these named arguments:

| Part | Description | Corresponding _fMsg_ Property |
| ---- |-------------| ----------------------------- |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. | Title |
| label1<br>label2<br>label3 | Optional. String expression displayed as label above the corresponding text_ | Label(section) |
| text1<br>text2<br>text3 | Optional.  String expression displayed as message section. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line. | Text(section) |
| monospaced1<br>monospaced2<br>monospaced3 | Optional. Defaults to False. When True,  the corresponding text is displayed with a mono-spaced font see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) | Monospaced(section)
| [buttons](#syntax-of-the-buttons-argument) | Optional. Defaults to vbOkOnly when omitted. Variant expression, either a [VB MsgBox value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>), a comma delimited string, a collection of string expressions, or a dictionary of string expressions. In case of a string, a collection, or a dictionary, each item either specifies a button's caption (up to 7) or a reply button row break (vbLf, vbCr, or vbCrLf). | Buttons |

### _ErrMsg_ (see [example](#error-message))
Displays an appealingly designed error message. This function is pretty specific because it is used by a common, nevertheless elaborated, error handler (yet not available on GitHub) with 
#### Syntax
```
mMsg.ErrMsg(errnumber _
[, errsource][, errdescription][, errline][, errtitle][, errpath[, errinfo]
```
| Part | Description | Corresponding _fMsg_ Property |
| ---- |-----------| --- |
| errnumber | Optional. Defaults to 0. A number expression. |
| errsource | Optional. Defaults to vbNullString. String expression indicating the fully qualified name of the procedure where the error occoured. | - |
|<small>errdescription</small>| String expression displayed as top message section with an above label "Error Description". There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line. | Text(1) |
| errline | Optional. Defaults to vbNullString. String expression indicating the line number within the error causing procedure's module where the error occured or had bee raised. | - |
| errtitle | Optional. String expression displayed in the title bar of the dialog box. When not provided, the title is assembled by using errnumber, errsource, and errline. | AppTitle |
| errpath | Optional. The "call stack" from the entry procedure down to the error source procedure. Displayed mono-spaced in order to allow a properly indented layout | Text(2), Monospaced(2) see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) |
| errinfo | Optional. Defaults to vbNullString. String expression providing an additional information about the error. Displayed under a label "Additional information". When not provided, the string is optionally extracted from the errdescription: When the string contains a "\|\|" it is split into errdescription and errinfo.| Text(3) |

### Syntax of the _buttons_ argument
```
button:=string|value[, rowbreak][, button2][, rowbreak][, button3][, rowbreak][, button4][, rowbreak][, button5][, rowbreak][, button6][, rowbreak][, button7]
```
| | |
|-|-|
string, button2 ... button7| captions for the buttons 1 to 7|
|value|the VB MsgBox argument for 1 to 3 buttons all in one row|
|rowbreak| vbLf or Chr(10). Indicates that the next button is displayed in the row below|

## Installation

1. Download fMsg.frm and fmsg.frx and import them to your project
2. Download mMsg.bas and import it to your project<br>or<br>alternatively copy the desired code from the Usage section to any standard module

## Usage

### Simple message
The following code copied to any standard module:
```
Public Function Box( _
       Optional ByVal title As String = vbNullString, _
       Optional ByVal prompt As String = vbNullString, _
       Optional ByVal buttons As Variant = vbOKOnly) As Variant

    With fMsg
        .AppTitle = title
        .SectionText(1) = prompt
        .ApplButtons = buttons
        .Show
        Box = .ReplyValue
    End With
    Unload fMsg
    
End Function
```
displays with:
```
Box title:="Any title" _
  , prompt:="Any message text" _
  ' buttons:=vbOkOnly
```
the message:
image

### Error message
The following code copied to any standard module:

displays the error message:

image

### Common message
When the UserForm <a id="raw-url" href="https://www.dropbox.com/s/h91lcqa52qrdl5f/fMsg.frm?dl=1">fMsg.frm</a> and the <a id="raw-url" href="https://www.dropbox.com/s/0m5eggyy3vx3126/fMsg.frx?dl=1">fMsg.frx</a> had been downloaded and imported to your project and the following code is copied into any standard module:
```
Public Function Msg(ByVal title As String, _
                    ByRef message As tMessage, _
           Optional ByVal buttons As Variant = vbOKOnly, _
           Optional ByVal returnindex As Boolean = False) As Variant
' ------------------------------------------------------------------
' General purpose MsgBox alternative message. By default returns
' the clicked reply buttons value
' ------------------------------------------------------------------
    
    With fMsg
        .MsgTitle = title
        .Msg = message
        .MsgButtons = buttons
         
        '+--------------------------------------------------------------------------+
        '|| Setup prior showing the form is a true performance improvement as it    ||
        '|| avoids a flickering message window when the setup is performed when    ||
        '|| the message window is already displayed, i.e. with the Activate event. ||
        '|| For testing however it may be appropriate to comment the Setup here in ||
        '|| order to have it performed along with the UserForm_Activate event.     ||
        .Setup '                                                                   ||
        '+--------------------------------------------------------------------------+
        
        .show
        On Error Resume Next    ' Just in case the user has terminated the dialog without clicking a reply button
        '~~ Fetching the clicked reply buttons value (or index) unloads the form.
        '~~ In case there were only one button to be clicked, the form will have been unloaded already -
        '~~ and a return value/index will not be available
        If returnindex Then Msg = .ReplyIndex Else Msg = .ReplyValue
    End With
    Unload fMsg

End Function

```
the following code
```
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
```             
re-displays the following message until the Ok button is clicked:

![](images/demo-1.png)

### Examples Summary
The examples above demonstrate  the use of the UserForm _fMsg_. Considering the [Common Public Properties](<Implementation.md#common-public-properties>) of the UserForm some can implement any similar kind of application specific message. 

## Proportional versus Mono-Spaced

#### _Monospaced_ = True

Because the text is ++not++  "wrapped" the width of the _Message Form_ is determined by the longest text line (up to the _Maximum Form Width_ specified). When the maximum width is exceeded a vertical scroll bar is applied.<br>Note: The title and the broadest _Button Row_ May still determine an even broader final _Message Form_.

#### _Monospaced_ = False (default)
Because the text is "wrapped" the width of a proportional-spaced text is determined by the current form width.<br>Note: When a message is displayed exclusively proportional-spaced the _Message Form_ width is determined by the length of the title, the required space for the broadest _Buttons Row_ and the specified _Minimum Form Width_.

<script src="https://utteranc.es/client.js"
        repo="https://github.com/warbe-maker/VBA-MsgBox-Alternative"
        issue-term="pathname"
        theme="github-light"
        crossorigin="anonymous"
        async>
</script>