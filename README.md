# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked.
Not sure whether it is worth it? See [Why another MsgBox](<Why%20another%20MsgBox.md>)

## Interfaces
Below are the 3 functions/interfaces using the UserForm _fMsg_:

| Function | Description |
| -------- | ----------- |
| _Box_ | Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 7 reply _buttons_ ordered in up to 7 rows. |
| _Msg_ | Displays<br>up to 3 message sections, each with an optional label and optionally monospaced<br>up to 7 buttons in up to 7 rows |
| _ErrMsg_ | Makes use of the advantages of **msg** in displaying a well designed error message. |


## Syntax
`msg(title, label1, text1, monospaced 1 label2, text2, monosoaced2, label3, text3, monospaced3 [,buttons])`

The MsgBox function syntax has these named arguments:

| Named argument | Description |
| ---- | ----------- |
| title | Required  String |
| label1 | Optional. String expression displayed as label above text1 |
| text1 | Optional.  String expression displayed as message section in the dialog box. There is no length limit. When the maximum height or width is exceeded a vertical and/or horizontal scrollbars is displayed. consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. |
| buttons | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 5 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. |

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

mMsg.Box title:=..., prompt:=...,buttons:=vbOkOnly
or alternatively when the clicked reply matters:

Select Case mMsg.Box(title:=..., prompt:=...,buttons:=vbYesNo)
   Case vbYes
   Case vbNo
End Select
Disregarding the buttons which is much more flexible all the rest behaves exactly like the VB MsgBox.

image

Error message
The error message below (my standard one) is displayed with

mMsg.ErrMsg errtitle:=..., errnumber:=.., errdescription:=..., errline:=.., errpath:="....", errinfo:="..."
and makes use of:

Three Message Sections each with a Message Section Label
A Monospaced font option for the errsource which displays a proper indented "path to the error"
image

Common decision message

image


### Summary
The examples above illustrate the use of the 3 functions in the module _mMsg_ which are in fact just interfaces to the UserForm (_fMsg_):
- _Box_
- _Msg_
- _ErrMsg_

Considering the [Common Public Properties](<Implementation.md#common-public-properties>) of the UserForm and the mechanism to receive the return value of the clicked reply button some can go ahead without the installation of the _mMsg_ module and implement his/her own application specific message function using those already implemented as examples only.