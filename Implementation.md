# Implementation
## General
- The implementation of the _Message Form_ (the UserForm _fMsg) is mostly design driven. I.e. the number of available _Message Sections_, the number of _Button Rows_, and the number of reply _Buttons_ in _Buttons Rows_ is primarily a matter of the design and requires only moderate code change.
As a consequence the implementation relies on the hierarchical order of the frames and controls therein. Controls object names  are used only where unavoidable as is for the click events of the reply _Buttons_ (at least I havenˋt found a way to avoid this).
```vbscipt
Private Sub cmbReply11_Click():  ReplyClicked 1, 1:   End Sub
Private Sub cmbReply12_Click():  ReplyClicked 1, 2:   End Sub
Private Sub cmbReply13_Click():  ReplyClicked 1, 3:   End Sub
Private Sub cmbReply14_Click():  ReplyClicked 1, 4:   End Sub
Private Sub cmbReply15_Click():  ReplyClicked 1, 5:   End Sub
Private Sub cmbReply16_Click():  ReplyClicked 1, 6:   End Sub
Private Sub cmbReply17_Click():  ReplyClicked 1, 7:   End Sub
Private Sub cmbReply21_Click():  ReplyClicked 2, 1:   End Sub
Private Sub cmbReply22_Click():  ReplyClicked 2, 2:   End Sub
Private Sub cmbReply23_Click():  ReplyClicked 2, 3:   End Sub
Private Sub cmbReply24_Click():  ReplyClicked 2, 4:   End Sub
Private Sub cmbReply31_Click():  ReplyClicked 3, 1:   End Sub
Private Sub cmbReply32_Click():  ReplyClicked 3, 2:   End Sub
Private Sub cmbReply33_Click():  ReplyClicked 3, 3:   End Sub
Private Sub cmbReply41_Click():  ReplyClicked 4, 1:   End Sub
Private Sub cmbReply42_Click():  ReplyClicked 4, 2:   End Sub
Private Sub cmbReply51_Click():  ReplyClicked 5, 1:   End Sub
Private Sub cmbReply61_Click():  ReplyClicked 6, 1:   End Sub
Private Sub cmbReply71_Click():  ReplyClicked 7, 1:   End Sub
```

- In order to keep the number of to-be-installed modules at minimum the implementation is without Class Modules by intention. Possibly Type declarations may have been an laternative for a more elegant implementation but aren't used.

- The controls (frames, text boxes, and command buttons) are collected at the _Message Form's_ initialization and these collections are used throughout the implementation.

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
## Width Adjustments
The _Message Form_ is initialized with the specified minimum message form width (see [Default Value Constants](#default-value-constants) which may be modified via the public property _MaxFormWidthPrcntgOfSceenSize_ (see [Public Properties of the _Message Form_](#public-properties-of-the-message-form)). A width expansion may be triggered by the setup (in the outlined sequence) of the following the width determining elements:
  1. **Title**  
When the **Title** exceeds the specified  maximum message form width some text will be truncated. However, with a default maximum message form width of 80 % of the screen width that will happen pretty unlikely.
  2. **Mono-spaced message section** followed by **Replies Rows**  
When either of the two exceeds the maximum message form width it will get a horizontal scroll bar.
  3. **Proportional spaced message sections**  
are setup at last and adjusted to the (by then) final message form width.

```vbscript
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
```

## Height Adjustments
### Height Increments
The height of the _Message Form_ is incremented along with the setup of the _Message Sections_ and the setup of the _Reply Buttons_ and the _Reply Button Rows_ respectively.

These height increments are at first done without considering any specified maximum _Message Form_ height.

### Height Decrements
When all elements are setup and the message form exceeds the specified maximum height (see [Default Value Constants](#default-value-constants) which may be modified via the [Public Property of the _Message Form_](#public-properties-of-the-message-form) _MaxFormHeightPrcntgOfSceenSize_, the height of the _Message Area_ frame and/or the _Buttons Area_ frame is reduced to fit and provided with a vertical scroll bar.

|Case|Action|
|----|------|
| The _Message Area_ takes 60% or more of the total areas height | The _Message Area_ gets a vertical scrollbar|
|The _Buttons Area_ takes 60% or more if the total areas height|The _Message Area_ gets a vertical scrollbar |
| Otherwise | Both areas get a vertical scroll bar|
 

## Width decrease
When the _Message Area's_ plus the _Button Area's_ height take more than the available maximum specified _Message Form_ height vertical scroll bars come into play. However, the space required for the scroll bar requires an reduction of the _Proportional Spaced_ Message sections and/or the _Button Rows_. In the first case this additionally - increases the _Message Section's_ height, in the second case the Button Rows may require a horizontal scroll bar.

## Vertical Re-positioning  
Adjusting the top position of displayed elements is due when all elements are setup. Subsequent re-positioning is due when the [width is decreased](#width-decrease) in order to make space for a vertical scroll bar. Together with the adjustment of the top position of the bottom-most element the new height of the message form is set - limited by the specified maximum form height.

Note: This re-positioning may be done at any time for testing in order to display the current setup state.

## Default Value Constants 

| Constant | Meaning |
| -------- | ------- |
| MONOSPACED_FONT_NAME | Default Font Name for mono-spaced message section text |
| MONOSPACED_FONT_SIZE | Default Font Size for mono-spaced message section text|
| FORM_WIDTH_MIN | Minimum _Message Form_ wird in pt|            | FORM_WIDTH_MAX_POW | Maximum _Message Form_ width as % of the screen size |
| FORM_HEIGHT_MAX_POW | |
| MIN_WIDTH_REPLY_BUTTON | Minimum width of a _Reply Button_ |


## Common Public Properties
| Property | R/W | Meaning, Usage | Comment |
| ---------- | --- | ------- | ----- |
| Title      | W |         |
| Label      | W | `fMsg.Label(section) = "...."` | section = 1,2,or 3 |
| Text       | W | `fMsg.Text(section) = "...."` | section = 1,2,or 3)  |
| MonoSpaced | W | `fMsg.MonoSpaced(section) = True` | section = 1,2,3<br>defaults to False when not provided for a section |        |
| Buttons    | W |         | |


## Public Size Properties
 
| Property | R/W | Default |
| -------- | --- | ------- |
| MaxFormHeightPrcntgOfScreenSize | R/W |  90 % |
| MaxFormWidthPrcntgOfScreenSize  | R/W |  80 % |
| MinimumFormWidth                |     |  300 pt | 

## Public Properties for test only
| Property | R/W | Meaning |
| -------- | --- | ------- |
| FramesWithCaption | W | Defaults to False. Frames are displayed with their "test" caption |
| FramesWithBorder  | W |  Defaults to False. Frames are displayed with a visible border |