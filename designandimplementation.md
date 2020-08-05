## VB MsgBox Alternative (the idea)
Not a 100% equivalent implementation but without the following limitations and flexibility flaws
* limited _Message Form_ size (truncated title and limited message text space)
* limited reply button options (max 3 with predefined caption text
* no mono-spaced text option

Things not (again/yet) implemented:
* specifying the default button
* display of an alert image like a ?, !, etc.

## Examples
The examples below not only  illustrate the major enhancements but are also examples of the three "flavors" of this MsgBox alternative. These 3 make use of the UseForm fMsg. Any other kind of "application specific" message may be implemented by making use of the public properties of the fMsg UserForm

### Simple message
The simple message implemented by the _Box_ function provides:
* A _Message Area_ with one _Message Text Section_
* A _Replies Area_ with up to 7 _Reply Buttons_ ordered all in one row or  underneath  in up to 7 rows with any Multiline caption text.

Despite the _replies_ parameter all others are pretty MsgBox alike.

image

### Error message

The error message below (my standard one) uses:
* The _Message Area_ with (all) 3 _Message Text Sections_ and each with the optional _Message Label_ and one with the _Monospaced_ font option
* The _Re-plies Area_ with one fixed *Ok* button.

image


### Decision requesting message
The below example uses most of the advantages 

image

## Specification of the alternative MsgBox
* A _Message Area_ with up to 3 _Message Sections_  
  * optionally _Mono-spaced_. 
  * optionally with a _Message Section Label_
* Up to 7 reply buttons in up to 7 _Reply Rows_. 
They first 3 may be used exactly like MsgBox offers them or for all of them with   any multi-line caption text (the replied value corresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
* Flexible message window width by considering the following facts and parameters
  * title width
  * the longest mono-spaced text line - if any
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
  * optionally mono-spaced (not word wrapped!)
  * optionally with a label
* Up to 5 reply buttons. 
either exactly like the VB MsgBox and additionally with any multi-line caption text. 
The replied value corresponds with the button content. I e. it is either vbOk, vbYe, vbNo, vbCancel, etc. or the button's caption text
* The message window width considers
  * the title width (avoiding truncation)
  * the longest mono-spaced text line - if any
  * the number and width of the displayed reply buttons
  * the specified minimum window width
  * the specified maximum message window width (as a % of the screen width)
* The message window height considers
  * the space required for the message sections and the reply buttons
  * the specified maximum message window height (as a % if the screen height)

### Handling of an exceeded width or height limits
* when the specified maximum width is exceeded either by a mono-spaced message section (proportional spaced sections are word wrapped and thus cannot exceed the maximum width) or by the number and width of the reply buttons, a horizontal scroll bar is displayed.
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
The UserForm uses a hierarchy of frames, each dedicated to a specific operation. On the UserForm level these are the *MessageArea* and the *RepliesArea* frames, both used for the assignment of the Top property.  
* *MessageArea*
  * ImageFrame
  * MessageSection1 to ...3  
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
    


## Design and implementation of the Message/UserForm
### General
- The implementation of the message form is strictly design driven. I.e. the number of available **Message Sections**, the number of **Reply Rows**, and the number of **Reply (Command) Buttons** is only a matter of the design and does not require any code change.
- The implementation does not make use of any of the control's object name but relies on the hierarchical order of the frames (see below).

### Message/UserForm design
The message form is organized in a hierarchy of frames as follows.

    +----Message Area (Frame)----------------+
    | +---Message Section (Frame)----------+ |
    | | Message Section Label (Label)      | |
    | | +--Message Section Text (Frame)--+ | |
    | | | Message Section (TextBox)      | | |
    | | +--------------------------------+ | |
    | +------------------------------------+ |
    +----------------------------------------+
    +----Reply Area (Frame)------------------+
    | +----Replies Row (Frame)-------------+ |
    | | Replies (CommandButtons)           | |
    | +------------------------------------+ |
    +----------------------------------------+
 
The controls (frames, textboxes, and command buttons) are collected with the message form's initialization and used throughout the implementation.

```vbscript
' Returns all controls of type (ctltype) which do have a parent (fromparent)
' as collection (into) by assigning the an initial height (ctlheight) and width (ctlwidth).
' -----------------------------------------------------------------------------------------
Private Sub Collect(ByRef into As Collection, _
                    ByVal fromparent As Variant, _
                    ByVal ctltype As String, _
                    ByVal ctlheight As Single, _
                    ByVal ctlwidth As Single)

    Dim ctl As MSForms.Control
    Dim v   As Variant
     
    On Error GoTo on_error
    
    Set into = Nothing: Set into = New Collection
    Select Case TypeName(fromparent)
        Case "Collection"
            '~~ Parent is each frame in the collection
            For Each v In fromparent
                For Each ctl In Me.Controls
                    If TypeName(ctl) = ctltype And ctl.Parent Is v Then
                        With ctl
                            Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Ctl: " & ctl.Name
                            .Visible = False
                            .Height = ctlheight
                            .width = ctlwidth
                        End With
                        into.Add ctl
                    End If
               Next ctl
            Next v
        Case Else
            For Each ctl In Me.Controls
                If TypeName(ctl) = ctltype And ctl.Parent Is fromparent Then
                    With ctl
                        Debug.Print "Parent: " & ctl.Parent.Name & ", Type: " & TypeName(ctl) & ", Ctl: " & ctl.Name
                        .Visible = False
                        .Height = ctlheight
                        .width = ctlwidth
                    End With
                    into.Add ctl
                End If
            Next ctl
    End Select
exit_proc:
    Exit Sub
    
on_error:
    Debug.Print Err.Description: Stop: Resume Next
End Sub
```


## Width Adjustment
The message form is initialized with the specified minimum message form width. Width expansion may be  triggered by the setup (in the outlined sequence) of the following width determining elements:
  1. **Title**  
When the **Title** exceeds the specified  maximum message form width some text will be truncated. However, with a default maximum message form width of 80 % of the screen width that will happen pretty unlikely.

  2. **Mono-spaced message section** followed by **Replies Rows**  
When either of the two exceeds the maximum message form width it will get a horizontal scroll bar.

  3. **Proportional spaced message sections**  
are setup at last and adjusted to the (by then) final message form width.



    ' Re-adjust width of message section text and
    ' adjust frames height accordingly
    ' ---------------------------------------------
    Private Sub MsgSectionAdjustHeightToAvailableWidth( _
            ByVal section As Long, _
            ByVal newwidth As Single)

        Dim s As String
        Dim siNewHeight As Single
     
        With MsgSectionText(section)
            s = .Value
            .Value = vbNullstring
             .AutoSize = False
            .Width = newwidth
            .MultiLine = True
            .AutoSize = True
            .Value = s
            MsgSectionTextFrame.Height = .Height + F_MARGIN
            MsgSectionTextFrame.Width = .Width + F_MARGIN
        End With
        
    End Sub

## Height Adjustment
### Height Increment
Height size increment is done along with
- the setup of a message section by the subsequent repositioning of all below displayed elements' top position
- the setup of the reply buttons (reply button rows respectively).

These height increments are done without considering  the specified maximum message form height.

### Height Decrement
When all elements are setup and the message form exceeds the maximum specified height the form height the message area and/or the reply area are adjusted. When the areas' height relation is 50/50 to 65/35 both areas will get a vertical scroll bar and the height is decremented by the corresponding relation. Otherwise only the taller area is reduced by the exceeding amount and gets a vertical scroll bar. The width of the scrollbar is the height before the reduction 

    '   
    Private Function MsgAreaHeight() As Single
    
    End Function


## Vertical Repositioning
Adjusting the top position of displayed elements is due initially when an element had need setup and subsequently whenever an element's height changed because of a width adjustment. Together with the adjustment of the top position of the bottommost element the new height of the message form is set.

Note: This top repositioning may be done just once when all elements had initially been  setup. However, for testing it is more appropriate to be performed immediately after setup of each individual element.

    Private Sub RepositionTop()
        ReposTopMsgSections
        ReposTopPosReplyRows
        ReposTopAreas
    End Sub


