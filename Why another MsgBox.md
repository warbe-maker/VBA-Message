# Why another MsgBox?

The shortest answer to the question may be the [Examples,  Demonstrations](#examples-demonstrations).

The idea dates back when I implemented my general error handling and was looking for a well designed and more appealing error message, not only displaying the description of the error but also the source if the error as a kind of call stack, preferably with a mono-spaced font and optionally some additional information.
Another feel for the need occurred when I tried to implement a more complex "decision message" with more choices than just Yes, No, Cancel, etc.. The idea was to have reply buttons with the most meaningful caption text possible in order to save wording in the above message text.

The result is now not 100% MsgBox equivalent because some features are excluded. The compare is as follows:

| The VB MsgBox | The Alternative "_Message Form_" |
| ------ | ---- |
| Limited message width | The maximum _Message Form_ width is specified as a percentage of the screen width and defaults to 80% |
| Limited message height |The maximum _Message Form_ height is specified as a percentage of the screen height and defaults to 90% |
| A message which exceeds the (hard to tell) size limit is truncated | When the maximum _Message Form_ size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message can optionally be displayed with a mono-spaced font |
| To display a well designed message is time consuming and no satisfactory result can be expected | There are up to 3 _Message Sections_ each with an optional _Message Text Label_ and each with a _Monospaced_ option |
| The maximum reply options (reply command buttons) is 3 | Up to 7 _Reply Buttons_ are available for being used and they may be displayed in 1 or up to 7 _Reply Rows_ (one in each row or all underneath) |
| The content (caption) of the reply buttons is a limited amount of terms (Ok, Yes, No, Ignore, Cancel) | The caption of the _Reply Buttons_ may be those known from MsgBox but in addition any multi-line text is possible |
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
The examples below not only illustrate the major enhancements but also the 3 implemented functions in the module _mMsg_ which do use the UserForm _fMsg_.

#### Simple message
The simple message implemented by the _Box_ function in module _mMsg_ is mainly for the compatibility with MsgBox. The example is displayed with
```
mMsg.Box msgtitle:=..., msgtext:=...,replies:=vbYesNo
```

image

#### Error message

The error message below (my standard one) is provided by the _ErrMsg_ function in the _mMsg_ Module and makes use of:
* Three _Message Sections_ each with a _Message Section Label_ 
* A _Monospaced_ font option for the error source which displays a proper indented "call stack"
* The _Re-plies Area_ with one fixed *Ok* button - common for VB error messages.

image

#### "Common" message

image