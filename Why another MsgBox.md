# Why another MsgBox?

The shortest and possibly already convincing answer may be the [Examples,  Demonstrations](#examples-demonstrations).

The idea has a long history. Implementing a general error handling I was looking for a well designed, maximum user friendly, and possibly more appealing **error message**. It should display
- the description of the error
-  the **path to the error**, preferably with a mono-spaced font
- some optionally additional information.
This was the birth of three message sections, each with an optional label.

Another feel for the need occurred when I tried to implement a more complex "decision message" with more choices than just Yes, No, Cancel, etc.. 
This was the birth of the idea to have reply buttons not only fully compatible with the MsgBox but some more and all with any meaningful caption text of any length (to replace a lengthy message above explaining when to click which button.

The result is now not 100% MsgBox equivalent because some features are excluded. The compare is as follows:

| The VB MsgBox | The Alternative "_Message Form_" |
| ------ | ---- |
| The message width and height is limited and cannot be altered | The maximum _Message Form_ width and height is specified as a percentage of the screen size. The width defaults to 80% the height defaults to  90% |
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum _Message Form_ size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may optionally be displayed with a mono-spaced font |
| Composing a fair designed message is time consuming and it is difficult to come up with a good result | With up to 3 _Message Sections_ each with an optional _Message Text Label_ and a _Monospaced_ option a good design is effortless |
| The maximum _Reply Buttons_) is 3 | Up to 7 _Reply Buttons_ allow a multiple choice dialog and they may be displayed in various orders (1 to 7 in 1 row, 7 rows with 1 in each  and many variants in between) |
| The content (caption) of the reply buttons is a limited number of - native English! - terms (Ok, Yes, No, Ignore, Cancel) | The caption of the _Reply Buttons_ may be those known from MsgBox but in addition any multi-line text is possible |
| Specifying the default button | (yet) not implemented |
| Display of an alert image like a ?, !, etc. | (yet) not implemented |

The implementation comprises of:
- a Standard Module _mMsg_ with the functions
  - _Box_ is mainly for the backwards compatibility with the _MsgBox_
  - _ErrMsg_ became my standard error message
  - _Msg_ allows full use of all design fetures
- UserForm _fMsg_

Beside these three functions any "application specific" message may be implemented analogously by making use of the public properties of the _Message Form_ (see [Implementation](#Implementation.md))

### Examples, Demonstrations
The examples below not only illustrate the major enhancements but also the 3 implemented functions in the module _mMsg_ which do use the UserForm _fMsg_: _Box_, _Msg_, _ErrMsg_

#### Simple message with  _Box_
Mainly for the compatibility with MsgBox it is displayed with
```
mMsg.Box title:=..., prompt:=...,buttons:=vbYesNo
```
The _buttons_ parameter however is much more flexible as it is handed over to the public _Message Form_ property _Replies_ (see [Implementation](#implementation.md)).

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

#### "Common" message

image