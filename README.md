# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked.
See also [Why another Msg box?](#Why%20another%20MsgBox.md)

## Interfaces
All interfaces use the UserForm _fMsg_ (_Message Form_) which may be used to create any application specific message interfaces.

| Function | Description |
| -------- | ----------- |
| _Box_ | Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 7 reply _buttons_ ordered in up to 7 rows. |
| _Msg_ | Displays<br>up to 3 message sections, each with an optional label and optionally monospaced<br>up to 7 buttons in up to 7 rows |
| _ErrMsg_ | Makes use of the advantages of **msg** in displaying a well designed error message. |


## Syntax
`msg(msgtitle, msgtext [,replies])`
The MsgBox function syntax has these named arguments:

| Part | Description |
| ---- | ----------- |
| msgtext | Required.  String expression displayed as the message in the dialog box. The maximum length of prompt is approximately 1024 characters, depending on the width of the characters used. If prompt consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. |
| replies | Optional.  Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 5 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. |

## Installation

- Download
  - fMsg.frm
  - fmsg.frx
  - mMsg.bas
- Import
  - fMsf.frm
  - mMsg.bas