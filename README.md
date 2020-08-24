# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked.
Not sure whether it is worth it? See [Why another MsgBox](<Why%20another%20MsgBox.md>)

## Interfaces
The alternative comes with 3 functions (interfaces) which use the  UserForm _fMsg_ for the message display and do receive the return value of the clicked reply _button_.

### _Msg_
#### Syntax
```
 - msg(title _<br>[[, label1][, text1][, monospaced1]]
 _
[[, label2][, text2][, monospaced2]] _
[[, label3][, text3][, monospaced3]] _
[,buttons])
```
#### Description


### _Box_
#### Syntax

```
Box(title, prompt[,buttons])
```
#### Description
Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 7 reply _buttons_ ordered in up to 7 rows.


### _Msg_
#### Syntax
```
Msg(title _
[[, label1][, text1][, monospaced1]] _
[[, label2][, text2][, monospaced2]] _
[[, label3][, text3][, monospaced3]] _
[,buttons])
```
#### Description
Displays up to 3 message sections, each with an optional label and optionally monospaced and up to 7 buttons in up to 7 rows.

### _ErrMsg_
#### Syntax
```
ErrMsg(errornumber _
, errordescription[,errorpath[,errorline]
```
#### Description
Uses the UserForm _fMsg_ to display a well designed error message.

### Named arguments

| Part | Function | Description |
| ---- | ----------- | --- |
| title | Box<br>Msg | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. |
| label1<br>label2<br>label3 | Msg | Optional. String expression displayed as label above the corresponding text_ |
| text1<br>text2<br>text3 | Msg | Optional.  String expression displayed as message section. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. |
| monospaced1<br>monospaced2<br>monospaced3 | Msg | Optional. Defaults to False. When True,  the corresponding text is displayed with a mono-spaced font see [Proportional- versus Mono-spaced](#proportional-versus-mono-spaced) |
| buttons | Box<br>Msg | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 5 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). |

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

```mMsg.Box title:=..., prompt:=...,buttons:=vbOkOnly
```
or alternatively when the clicked reply matters:

```
Select Case mMsg.Box(title:=..., prompt:=...,buttons:=...)
   Case vbYes
   Case vbNo
End Select
```
Disregarding the buttons argument - which is much more flexible - the arguments are identical with those of the VB MsgBox.

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

## Proportional- versus Mono-Spaced
The result of the two differs significantly
- _Monospaced_ = True <br> The width of the _Message Form_ is determined by the longest text line (up to the maximum form width specified)  because the text is ++not++  "wrapped"
- _Monospaced_ = False (default) <br> The width of a proportional-spaced text is determined by the form width because it is "wrapped". For a message which is exclusively displayed proportioal-spaced it is pretty likely that the specified _Minimum Form Width_ ist used - unless the length of the title determines a wider form width.