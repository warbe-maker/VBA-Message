# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked. This MsgBox alternative comes in three flavors

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

Download: fMsg.frm, fmsg.frx, mMsg.bas
Import into your project: fMsf.frm and mMsg.bas (fmsg.frx is automatically imported together with fmsg.frm)

Note: For the following examples the above files are imported into the Workbook Msg.xlsm in which also all testing is prepared.



## Usage examples
| Example | Syntax | Displayed message |
| ------- | ------ | ----------------- |
| Simple message | `mMsg.msg1 _ msgtitle:="any", _ msgtext:="any"` | |
| Elaborated user decision dialog | | |
| Elaborated error message | | |

See the wiki for further details