# MsgBox Alternative

Displays a message in a dialog box, waits for the user to click a button, and returns a variant indicating which button the user clicked. The re-engineered MsgBox comes in two flavors
- Function **msg** is very MsgBox alike but with some significant enhancements
- Function **msg3** is still pretty MsgBox alike but with parameters allowing a very designed user decision message or an error message 
See the .... for a complete description

## Syntax
msg(sTitle, sMsgText[¥,vReplies])
The MsgBox function syntax has these named arguments:

Part | Description
---- | -----------
sMsgText | Required. String expression displayed as the message in the dialog box. The maximum length of prompt is approximately 1024 characters, depending on the width of the characters used. If prompt consists of more than one line, you can separate the lines by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line.
vReplies | Optional. Variant expression, either the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for buttons is 0.
title | Required. String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.

## Installation

Download: fMsg.frm, fmsg.frx, mMsg.bas
Import to your project: fMsf.frm, mMsg.bas



## Usage examples
### Simple message
- mMsg.Msg sTitle:="any", sMsgText:="any", vReplies:=vbOkOnly
  displays the following message

### Simple user decision message

### Elaborated Error message

See the wiki for further details
