## Why another MsgBox?

The shortest and possibly already convincing answer may be the [Examples,  Demonstrations](#examples-demonstrations).

The idea has a long history. Implementing a general error handling I was looking for a well designed, maximum user friendly, and possibly more appealing **error message**. It should display
- the description of the error
-  the **path to the error**, preferably with a mono-spaced font
- some optionally additional information.
This was the birth of three message sections, each with an optional label.

Another feel for the need occurred when I tried to implement a more complex "decision message" with more choices than just Yes, No, Cancel, etc.. 
This was the birth of the idea to have reply buttons not only fully compatible with the MsgBox but some more and all with any meaningful caption text of any length (to replace a lengthy message above explaining when to click which button.

### Comparison

| The VB MsgBox | The Alternative "_Message Form_" |
| ------ | ---- |
| The message width and height is limited and cannot be altered | The maximum _Message Form_ width and height is specified as a percentage of the screen size. The width defaults to 80% the height defaults to  90% |
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum _Message Form_ size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may optionally be displayed with a mono-spaced font |
| Composing a fair designed message is time consuming and it is difficult to come up with a good result | With up to 3 _Message Sections_ each with an optional _Message Text Label_ and a _Monospaced_ option a good design is effortless |
| The maximum reply _Buttons_) is 3 | Up to 7 reply _Buttons_ may be displayed in any order in up to 7 reply _Button Rows_   |
| The content (caption) of the reply buttons is a limited number of - native English! - terms (Ok, Yes, No, Ignore, Cancel) | The caption of the reply _Buttons_ may be those known from MsgBox and additionally any multi-line text |
| Specifying the default button | (yet) not implemented |
| Display of an alert image like a ?, !, etc. | (yet) not implemented |


### Examples, Demonstrations
#### Simple message
Mainly for the compatibility with MsgBox it is displayed with
```vbscript
mMsg.Box title:=..., prompt:=...,buttons:=vbOkOnly
```
or alternatively when the clicked reply matters:
```vbscript
Select Case mMsg.Box(title:=..., prompt:=...,buttons:=vbYesNo)
   Case vbYes
   Case vbNo
End Select
```
Disregarding the _buttons_ which is much more flexible all the rest behaves exactly like the VB MsgBox.

image

#### Error message

The error message below (my standard one) is displayed with

```vbscript
mMsg.ErrMsg errtitle:=..., errnumber:=.., errdescription:=..., errline:=.., errpath:="....", errinfo:="..."
```
and makes use of:
* Three _Message Sections_ each with a _Message Section Label_ 
* A _Monospaced_ font option for the _errsource_ which displays a proper indented "path to the error"

image

#### Common decision message

image

### Summary
The above examples not only illustrate the benefit of the _Alternative MsgBox_ but also demonstrate the use of the 3 functions in the module _mMsg_ which are the interface to the UserForm (_fMsg_):
- _Box_
- _Msg_
- _ErrMsg_

Considering the public properties in the [Implementation, Public Properties](<Implementation.md>) of the UserForm and the mechanism to receive the return value of the clicked reply button some can go ahead without the installation of the _mMsg_ module and implement his/her own application specific message function using those already implemented as examples only.