# MsgBox Alternative

[Abstract](#abstract)<br>
[Why an alternative MsgBox](#why-an-alternative-msgbox)<br>[Interfaces](#interfaces)<br>[Installation](#installation)<br>[Usage](#usage)

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
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order |
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| Specifying the default button | (yet) not implemented |
| Display of an ?, !, etc. image | (yet) not implemented |

### Interfaces
The alternative implementation  comes with three functions (in module _mMsg_) which are the interface to the UserForm _fMsg_ and return the clicked reply _Button_ value to the caller.

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
| [buttons](#syntax-of-the-buttons-argument) | Optional. Defaults to vbOkOnly when omitted. Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 7 reply buttons. | Buttons |

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
| ---- |-----------| --- |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. | Title |
| label1<br>label2<br>label3 | Optional. String expression displayed as label above the corresponding text_ | Label(section) |
| text1<br>text2<br>text3 | Optional.  String expression displayed as message section. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (vbCr or Chr(13), a linefeed character (vbLf or Chr(10)), or carriage return - linefeed character combination (vbCrLf or Chr(13) & Chr(10)) between each line. | Text(section) | monospaced1<br>monospaced2<br>monospaced3 | Optional. Defaults to False. When True,  the corresponding text is displayed with a mono-spaced font see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) | Monospaced(section)
| [buttons](#syntax-of-the-buttons-argument) | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 7 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). | Buttons |

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
| errinfo | Optional. Defaults to vbNullString. String expression providing an additional information about the error. Displayed under a label "Additional information". When not provided, the string is extracted from the errdescription which follows an "||" indication.| Text(3) |

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
When the following code is copied into any standard module:
```
Public Function Msg(ByVal title As String, _
           Optional ByVal section1label As String = vbNullString, _
           Optional ByVal section1text As String = vbNullString, _
           Optional ByVal section1monospaced As Boolean = False, _
           Optional ByVal section2label As String = vbNullString, _
           Optional ByVal section2text As String = vbNullString, _
           Optional ByVal section2monospaced As Boolean = False, _
           Optional ByVal section3label As String = vbNullString, _
           Optional ByVal section3text As String = vbNullString, _
           Optional ByVal section3monospaced As Boolean = False, _
           Optional ByVal monospacedfontsize As Long = 0, _
           Optional ByVal buttons As Variant = vbOKOnly) As Variant
    
    With fMsg
        .AppTitle = title
        
        .SectionLabel(1) = section1label
        .SectionText(1) = section1text
        .SectionMonoSpaced(1) = section1monospaced
        
        .SectionLabel(2) = section2label
        .SectionText(2) = section2text
        .SectionMonoSpaced(2) = section2monospaced
        
        .SectionLabel(3) = section3label
        .SectionText(3) = section3text
        .SectionMonoSpaced(3) = section3monospaced

        .ApplButtons = buttons
        .Show
        Msg = .ReplyValue
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
   Dim sButton1 As String
   Dim sButton2 As String
   Dim sButton3 As String
   Dim sButton4 As String
   Dim sButton5 As String
   Dim sButton6 As String
   Dim sButton7 As String

   sTitle = "Usage demo: Full featured multiple choice message"
   sLabel1 = "Demo 1:"
   sText1 = "Use of all 3 message sections, all with a label"
   sLabel2 = "Demo 2"
   sText2 = "Use of all 7 reply buttons, in a 2-2-2-1  order."
   sLabel3 = "Demo 3:"
   sText3 = "This part of the message just demonstrates the mono-spaced option." &vbLf & _
   "Specifically the result it has on the message width," & vbLf & _
   "which it determines through its longest line."
   sButton1 = "Multiline reply button text" & vbLf & "Button-1"
   sButton2 = "Multiline reply button text" & vbLf & "Button-2"
   sButton3 = "Multiline reply button text" & vbLf & "Button-3" 
   sButton4 = "Multiline reply button text" & vbLf & "Button-4"
   sButton5 = "Multiline reply button text" & vbLf & "Button-5"
   sButton6 = "Multiline reply button text" & vbLf & "Button-6"
   sButton7 ="Ok"
   '~~ Assemble the buttons argument string            
   sButtons = _
   sButton1 & "," & sButton2 & "," & vbLf & "," & _   
   sButton3 & "," & sButton4 & "," & vbLf & "," & _   
   sButton4 & "," & sButton5 & "," & vbLf & "," & sButton7 _

   Do
      If Msg( _
         title:=sTitle, _
         label1:=sLabel1, text1:=sText1, _
         label2:=sLabel2, text2:=sText2, _
         label3:=sLabel3, text3:=sText3, _
         monospaced3:=True)
      = sButton7 _
      Then Exit Do
   Loop
   
End Sub
             

re-displays the following message until the Ok button is clicked:

<a href="/images/demo-1.jpg">

### Examples Summary
The examples above demonstrate  the use of the UserForm _fMsg_. Considering the [Common Public Properties](<Implementation.md#common-public-properties>) of the UserForm some can implement any similar kind of application specific message. 

## Proportional versus Mono-Spaced

#### _Monospaced_ = True

Because the text is ++not++  "wrapped" the width of the _Message Form_ is determined by the longest text line (up to the _Maximum Form Width_ specified). When the maximum width is exceeded a vertical scroll bar is applied.<br>Note: The title and the broadest _Button Row_ May still determine an even broader final _Message Form_.

#### _Monospaced_ = False (default)
Because the text is "wrapped" the width of a proportional-spaced text is determined by the current form width.<br>Note: When a message is displayed exclusively proportional-spaced the _Message Form_ width is determined by the length of the title, the required space for the broadest _Buttons Row_ and the specified _Minimum Form Width_.
