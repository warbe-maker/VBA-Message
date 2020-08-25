# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked.

### Why alternative MsgBox?
The alternative implementation addresses nany of the MsgBox deficiencies.

| VB MsgBox | Alternative |
| ------ | ---- |
| The message width and height is limited and cannot be altered | The maximum width and height is specified as a percentage of the screen size which defaults 80% width and  90% height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it) may be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a good result | With up to 3 _Message Sections_ each with an optional _Message Text Label_ and a _Monospaced_ option an appealing design is effortless |
| The maximum reply _Buttons_) is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order |
| The caption) of the reply _Buttons_ is based on a value (vbYesNo, vbOkOnly, etc. and result in untranslated native English! - terms (Ok, Yes, No, Ignore, Cancel) | The caption of the reply _Buttons_ may specified by those values known from the VB MsgBox but additionally any multi-line text may be specified |
| Specifying the default button | (yet) not implemented |
| Display of an alert image like a ?, !, etc. | (yet) not implemented |

## Interfaces
The alternative implementation  comes with three functions (in module _mMsg_). These are interfaces to the UserForm _fMsg_ and return the clicked reply _Button_ value to the caller.

### _Box_
Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 7 reply _buttons_ in up to 7 rows in any order.

#### Syntax
```
mMsg.Box(prompt[ ,buttons][ ,title ])
```
The _Box_ function syntax has these named arguments:

| Part | Description | Corresponding _fMsg_ Property |
| ---- |-----------| --- |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. | Title |
| prompt | String expression displayed as message. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. | Text(1) |
| buttons | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 7 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). | Buttons |

### _Msg_
Displays a message in up to 3 sections, each with an optional label and optionally monospaced and up to 7 buttons in up to 7 rows in any order.
#### Syntax
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
| text1<br>text2<br>text3 | Optional.  String expression displayed as message section. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. | Text(section) | monospaced1<br>monospaced2<br>monospaced3 | Optional. Defaults to False. When True,  the corresponding text is displayed with a mono-spaced font see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) | Monospaced(section)
| buttons | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 7 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). | Buttons |

### _ErrMsg_
Displays a well designed error message.
#### Syntax
```
mMsg.ErrMsg(errornumber _
, errordescription[, errorpath[, errorline]
```
| Part | Description | Corresponding _fMsg_ Property |
| ---- |-----------| --- |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. | Title |
| label1<br>label2<br>label3 | Optional. String expression displayed as label above the corresponding text_ | Label(section) |
| text1<br>text2<br>text3 | Optional.  String expression displayed as message section. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. Lines may be separated by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. | Text(section) | monospaced1<br>monospaced2<br>monospaced3 | Optional. Defaults to False. When True,  the corresponding text is displayed with a mono-spaced font see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) | Monospaced(section)
| buttons | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 7 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). | Buttons |

### Syntax of the buttons argument
```
button:=string|value[, rowbreak][, button2][, rowbreak][, button3][, rowbreak][, button4][, rowbreak][, button5][, rowbreak][, button6][, rowbreak][, button7]
```
| | |
|-|-|
string, button2 ... button7| captions for the buttons 1 to 7|
|value|the VB MsgBox argument for 1 to 3 buttons all in one row|
|rowbreak| vbLf or Chr(10). Indicates that the next button is displayed in the row below|

## Installation

- Download
  - fMsg.frm
  - fmsg.frx
  - mMsg.bas
- Import
  - fMsf.frm
  - mMsg.bas

## Usage, Examples

### Simple message

Mainly for the compatibility with MsgBox it is displayed with

```MsgBox (prompt, [ buttons, ] [ title, ]
```
or alternatively when the clicked reply matters:

```
Select Case mMsg.Box(title:=..., prompt:=...,buttons:=...)
   Case vbYes
   Case vbNo
End Select
```

The MsgBox function syntax has these named arguments:


image

### Error message
The error message below (my standard one) is displayed with

```
mMsg.ErrMsg errtitle:=..., errnumber:=.., errdescription:=..., errline:=.., errpath:="....", errinfo:="..."
```
and makes use of all 3 _Message Sections_ each with a _Message Section Label_,
a _Monospaced_ font is used for the  _errorsource_ to be displayed properly indented.

image

### Common decision message

```
Select Case mMsg.Box(title:=..., prompt:=...,buttons:=...)
   Case ...
   Case ...
End Select
```

image


### Examples Summary
The examples above illustrate the use of the 3 functions (interfaces) in the module _mMsg_ using the UserForm _fMsg_: _Box_, _Msg_, _ErrMsg_

Considering the [Common Public Properties](<Implementation.md#common-public-properties>) of the UserForm and the mechanism to receive the return value of the clicked reply button some can go ahead without the installation of the _mMsg_ module and implement his/her own application specific message function using those already implemented as examples only.

## Proportional versus Mono-Spaced
The result of the two differs significantly
- _Monospaced_ = True <br> The width of the _Message Form_ is determined by the longest text line (up to the maximum form width specified)  because the text is ++not++  "wrapped"
- _Monospaced_ = False (default) <br> The width of a proportional-spaced text is determined by the form width because it is "wrapped". For a message which is exclusively displayed proportioal-spaced it is pretty likely that the specified _Minimum Form Width_ ist used - unless the length of the title determines a wider form width.