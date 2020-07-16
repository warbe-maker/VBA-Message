

## VB MsgBox Alternative
Addresses the following limitations and flexibility flaws:
* Limited width
* Title truncation
* Limited message space
* Inflexible reply button options (limited in number and caption text)
* No monospace font option

## Alternative message setup and designs

The following examples should illustrate the effect

# VB MsgBox Alternative
The alternative VB MsgBox is not a 100% equivalent but comes with the main limitations eliminated
* limited window width, resulting in a truncated title
* limited message text space
* limited reply button options (number caption text)
* no monospaced text
Things not implemented yet are
* specifying the default button
* display of an image like a ?, !, etc.

## Examples illustrating the major enhancements 

### Simple message pretty analogous to Msgbox
image

### A "pimped" Error message

The error message below is displayed with my error handler module which will be a future repo.
image

### An enhanced decision requesting message
image

## Specification of the alternative MsgBox
* Up to 3 message sections  
  * optionally monospaced. 
  * optionally with a label
* Up to 5 reply buttons either exactly like Msgbox offers them but additionally with any multiline caption text whereby the replied value korresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
* Flexible message window width by considering the following facts and parameters
  * title width
  * the longest monospaced text line - if any
  * the number and width of the displayed reply buttons
  * minimum window width in pt
  * maximum window width (specified as percentage of the screen width)
* Flexible message window height by considering the following facts an parameters
  * maximum window height (specified as percentage of the screen height)
  * adjusted up to the screen height
  - Message paragraphs which had to be limited in their height show a vertical scroll bar

image

### A complex decision requesting dialog 
image

## Specification
### Basics
* Up to 3 message sections
  * optionally monospaced (not word wrapped!)
  * optionally with a label
* Up to 5 reply buttons. 
either exactly like the VB MsgBox and additionally with any multiline caption text. 
The replied value corresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
* The message window width considers
  * the title width (avoiding truncation)
  * the longest monospaced text line - if any
  * the number and width of the displayed reply buttons
  * the specified minimum window width
  * the specified maximum message window width (as a % of the screen width)
* The message window height considers
  * the space required for the message sections and the reply buttons
  * the specified maximum message window height (as a % if the screen height)

### Handling of an exceeded width or height limits
* when the specified maximum width is exceeded either by a monospaced message section (proportional spaced sections are word wrapped and thus cannot exceed the maximum width) or by the number and width of the reply buttons, a horizontal scroll bar is displayed.
* when the specified maximum height is exceeded, the highest message section's height is reduced to fit and a vertical scroll bar is displayed.




## Installation
See ReadMe

## Usage

## Examples

## Parameters
There are much more parameters available than the ones obviously required for any kind of message. The additional parameters allow the implementation of VB project specific message procedures.

### Basic

| Parameter | applicable for (procedure in mMsg module) | meaning |
| ------- | -------- | ---------- |
| msgtitle | msg, box | The text displayed in the handle bar |
| msgtext | box | The one and only text displayed |
| msg1label | msg| label for the first message section |
| msg1text | msg | text for the first message section |
| msg1monospaced | msg | optional, defaults to False |
| msg2label | msg| label for the first message section |
| msg2text | msg | text for the first message section |
| msg2monospaced | msg| optional, defaults to False |
| msg3label | msg| label for the first message section |
| msg3text | msg | text for the first message section |
| msg3monospaced | msg | optional, defaults to False || vReplies | msg, msg3 | The number and content of the reply buttons (see Table below), defaults to __vbOkOnly__ |
| replies | msg, box | specifies the to be displayed reply buttons, optional, defaults to vbOkOnly |


#### Parameter replies
| Value | Result |
| ----- | -------------------- |
| vbOkOnly, vbYesNo, etc. analogous MsgBox | Up to 3 VB MsgBox alike reply buttons |
| Up to five comma delimited text strings | Each string is displayed as a reply button |
| | Example: | 
| | replies:="Yes,No,Cancel". |
| | is the eequivalent of. |
| | replies:=vbYesNoCancel |

## Development and Test

The Excel Workbook Msg xlsm is for development and testing. The module mTest provides all means for a proper regression test. The implemented tests are available via the test Worksheet Test/wsMsgTest. The test procedures in the mTest module are designed for a compact and complete test of all functions, options and boundaries and in that not necessarily usefully usage examples. For usage examples the procedures in the mExamples module may preferably consulted.
Performing a regression test should be obligatory for anyone contributing by code modifications for any purpose or reason. See Contributing.

## Design and implementation
### UserForm
The Userform uses a hierachy of frames, each dedicated to a specific operation
* MessageSections:  
 .Top = T_MARGIN.  
Collection of MessageSection (cllMsgSections).
  * MessageSection. (. = 1-3)  
Property Get MessageSection(Optional ByVal section As Long) As MsForms.Frame
    * SectionLabel. 
Property Get MsgLabel(Optional ByVal section As Long) As MsForns.Label
    * SectionFrame. 
Property Get MsgFrame(Optional ByVal section As Long) As MsForns.Frame
      * SectionText. 
Property Get MsgText(Optional ByVal section As Long) As MsForns.TextBox
* RepliesSection:  
Bottom frame. 
.Top = MessageSections.Top + MessageSections.Height + V_MARGIN. 
Collection of RepliesRow (cllReplyRows). 
Property Get RepliesRow(Optional ByVal row As Long) As MsForms.Frame
  * RepliesRow. (. = 1 - 6)
    * RepliesFrame. (. = 1 - 6)
Collection of ReplyButton.
      * ReplyButton. (. = 1-6)

The UserForm is prepared for 6 reply button which may appear as follows
* Row 1: 1 to 6 buttons
* Row 2: 0 to 3 buttons
* Row 3: 0 to 2 buttons
* Row 4 to 6: 0 to 1 button

The order depends on the specified maximum message form width and the width of the largest button - wich defines the width for all the other buttons. When the specified maximum height is exceeded by the reply buttons the all used rows surrounding frame is reduced to fit the form and a vertical scroll bar is applied. The visible height will be at least one and a half button row. When the form will still exceed ist's maximum width, the greatest message section will be processed the same way.

Private Property Get ReplyButton(Optional ByVal row As Long, Optional ByVal button As Long) As MsForms.CommandButton

## Implementation
The hierarchy of elements (message section labels 1 to n, message section text frames 1to n),  message section textboxes 1to n, and reply rows commandbuttons 1 to n) is obtained without the use of any control names. The number of message sections and reply buttons is not limited by the design since missing elements are created dynamically.

## Public Properties
### Commonly used properties

| Property | R/W | Purpose |
| -------------- | ------- | ------------- |
|


### Special properties
Additional special properties are available for the modification of the message appearance and last but not least for the implementation of dedicated message functions for specific needs in a VB project. As an example, some of them are used by the test procedures.

| Property | R/W | Purpose |
| -------------- | ------- | ------------- |
|

### Constants
The following constants are initialization values or directly used for the layout and appearance of a displayed message. Some of the initial values may be modified through the special properties.

| Constant | Specifies | Default |
| --------------- | --------------- | ------------ |
| MIN_FORM_WIDTH | minimum width of the message window  | 300 pt |
| MIN_REPLY_HEIGHT | 30 pt |
| MIN_REPLY_WIDTH | minimum width of a reply buttons  | 50 pt |
| MAX_FORM_HEIGHT_POW | maximum message width as percentage of the screen height | 80 % |
| MAX_FORM_WIDTH | maximum percentage space used of the screen height | 80 % |
| T_MARGIN | top margin | 5 or |
| B_MARGIN | bottom margin | 40 pt |
| L_MARGIN | left margin | 0 PT |
| R_MARGIN | right margin | 5 or |
    



