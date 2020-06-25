# VB MsgBox re-engineered
## Capabilities

This re-engineered message box allows the display of a wide range of information and user dialogues.
- Simple message - pretty analogous to the Msgbox

  Insert image

- Error message

  Insert image

- User decision dialogue

  Insert image

- Text files content (e.g. .ini or .cfg files)

  Insert image

## Msgbox deficiencies addressed

- Limited window width, resulting in a truncated title
- Limited message text space
- Limited reply buttons in number and also regarding their possible caption text
- Inability to display monospaced text

## Specification details

- Up to 3 message paragraphs
  - optionally monospaced
  - optionally labelled
- Up to 5 reply buttons
  - Either up to 3 analogous to Msgbox
  - Up to 5 with any kind of multiline caption text
    whereby the replied value corresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
- Adjusted window/form width considering
  - The title width
  - The longest monospaced text line - if any
  - The number and width of the displayed reply buttons
  - A minimum window width
  - A maximum window width, specified as percentage of the screen width
- Adjusted window/form height considering
  - Any displayed elements
  - The height of the message paragraphs
  displaying a vertical scroll bar when limited in the height
  - A maximum window height, specified as percentage of the screen heght

## Installation
Copy the code in file msg.bas into any common or basic standard module
```
```

Download the files msg.frm and msg.frx and import them

Add the following at the module level
```
```

## Usage

## Examples

## Parameters for the function msg

| Parameter | meaning |
| ------- | ---------- |
| sTitle | The text displayed in the window's handle/title bar |
| sMsgText | The one and only text displayed |
| vReplies | The number and content of the reply buttons (see Table below), defaults to __vbOkOnly__ |


## Parameters for the function msg3

| Parameter | meaning |
| ------- | ---------- |
| sTitle | The text displayed in the window's handle/title bar |
| sText1, sText2, sText3 | Message paragraphs |
| sLabel1, sLabel2, sLabel3 | Label corresponding to the message paragraphs |
| bMonospace1, bMonospace2, bMonospace3 | True = Message paragraph monospaced |
| vReplies | The number and content of the reply buttons (see Table below), defaults to __vbOkOnly__ |

#### Parameter vReplies
| Value | Meaning |
| ------------- | ------- |
| vbOkOnly, vbYesNo, etc. analogous MsgBox | MsgBox alike reply buttons (up to 3) |
| Any comma delimited text string (up to 5 strings) which may include line breaks for multiline reply button text | Will be displayed in as many buttons |

Example: A parameter vReplies:="Yes,No,Cancel" results in the same reply buttons as a parameter vReplies:=vbYesNoCancel






## Constants specified at module level
| Constant | Specifies | Format | Default |
| --------------- | --------------- | ------------ | ------------ |
| MIN_FORM_WIDTH | minimum width of the message window | Single | 250 |
| MIN_REPLY_WIDTH | minimum width of reply buttons | Single | 50 |
| MAX_FORM_HEIGHT | maximum percentage space used of the screen height | Long | 80 |
| MAX_FORM_WIDTH | maximum percentage space used of the screen height | Long | 80 |





