# VB MsgBox Alternative
## Possible message layouts
The alternative message box allows the display of a wide range of information, e.g. even the content of text files which should be illustrated below. The main deficiencies  eliminated are:

- The very limited window width, resulting in a truncated title
- The very limited message text space
- The very limited reply buttons in number and the possible caption text as well
- The inability to display monospaced text

### Simple message pretty analogous to Msgbox
image

### "Pimped" Error message
image

### New kinds of reply buttons
image

## Specification of a re-engineered message box

- Up to 3 message paragraphs
  optionally monospaced and optionally labelled
- Up to 5 reply buttons either exactly like Msgbox offers them but additionally with any multiline caption text whereby the replied value korresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
- Flexible window width considering
  - The title width
  - The longest monospaced text line - if any
  - The number and width of the displayed reply buttons
  - A specified minimum window width
- Flexible window height
  - Adjusted up to the screen height
  - Message paragraphs which had to be limited in their height show a vertical scroll bar

## Installation

## Usage

## Examples

## Parameters

| Parameter | applicable for | meaning |
| ------- | -------- | ---------- |
| sTitle | msg, msg3 | The text displayed in the handle bar |
| sMsgText | msg | The one and only text displayed |
| vReplies | msg, msg3 | The number and content of the reply buttons (see Table below), defaults to __vbOkOnly__ |
| sText1, sText2, sText3 | msg3 | Message paragraphs |
| sLabel1, sLabel2, sLabel3 | msg3 | Label corresponding to the message paragraphs |
| bMonospace1, bMonospace2, bMonospace3 | msg3 | True = Message paragraph monospaced |

#### Parameter vReplies
| Value | Meaning |
| ------------- | ------- |
| vbOkOnly, vbYesNo, etc. analogous MsgBox | MsgBox alike reply buttons (up to 3) |
| Any comma delimited text string (up to 5 strings) which may include line breaks for multiline reply button text | Will be displayed in as many buttons |

Example: A parameter vReplies:="Yes,No,Cancel" results in the same reply buttons as a parameter vReplies:=vbYesNoCancel

## Development and Test

The Excel Workbook Msg xlsm is for development and testing. The module mTest provides all means for a proper regression test. The implemented tests are available via the test Worksheet Test/wsMsgTest. The test procedures in the mTest module are designed for a compact and complete test of all functions, options and boundaries and in that not necessarily usefully usage examples. For usage examples the procedures in the mExamples module may preferably consulted.
Performing a regression test should be obligatory for anyone contributing by code modifications for any purpose or reason. See Contributing.

# UserForm
## Design
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
    



