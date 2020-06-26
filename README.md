# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked. This MsgBox alternative comes in three flavors

| Function | Description |
| -------- | ----------- |
| **msg1** | Pretty MsgBox alike, displays a single message with any number of line breaks, with up to 5 free reply buttons. |
| **msg** | Still pretty MsgBox alike, displays up to 3 message junks, each optionally monospaced and with an optional label and up to 5 free reply buttons. |
| **errmsg** | Makes use of the advantages of **msg** in displaying a well designed error message. |


## Syntax
**msg**(msgtitle, msgtext [,replies])
The MsgBox function syntax has these named arguments:

| Part | Description |
| ---- | ----------- |
| msgtext | Required. String expression displayed as the message in the dialog box. The maximum length of prompt is approximately 1024 characters, depending on the width of the characters used. If prompt consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line. |
| replies | Optional. Variant expression, either MsgBox values like vbOkOnly, vbYesNo, etc. or a comma delimited string specifying the caption of up to 5 reply buttons. If omitted, the default value for buttons is 0 (vbOkOnly). |
| title | Optional. String expression displayed in the title bar of the dialog box. When omitted, the application name is placed in the title bar. |

## Installation

Download: fMsg.frm, fmsg.frx, mMsg.bas
Import into your project: fMsf.frm and mMsg.bas (fmsg.frx is automatically imported together with fmsg.frm)



## Usage examples
### Simple message
- mMsg.Msg sTitle:="any", sMsgText:="any", vReplies:=vbOkOnly
  displays the following message

### Simple user decision message

### Elaborated Error message

See the wiki for further details
