# Common VBA Message Form and Display services

Supplements the README and the [Common-VBA-Message-Services](https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html) post focusing on technical/implementation aspects.

The implementation addresses the major shortcomings of the VBA.MsgBox thereby providing a sort of common message services. Addressed shortcomings:
* limited window width, resulting in title truncation
* very limited message text space
* reply button limitations in number and caption text
* no font options like mono-spaced, bold, italic, underlined, color, name, and size


## Design
### Message Area
Up to 4 message sections either proportional- or mono-spaced, each with an optional label (header)
### Buttons Area
49 buttons, 7 rows by 7 buttons.
### Implementation of the design
The UserForm uses four kinds of controls:
- Frame
- Label
- TextBox
- Command button
ordered in a hierarchy of frames:
- Message-Area frame
  - Message-Section frame
    - Label
    - Message-Text frame
      - TextBox
- Buttons-Area frame
  - Button-Row frame (7)
    - CommandButton (7)


## Implementation of the display services
### Message window width adjustment
The initial setup process focuses on the determination of the final message window width limited by a maximum width specified as a percentage of the available display width. The setup starts with the minimum width specified as pt and continues with ...
1. **Title setup**<br>
The pattern<br> `eeeeeeeeeeeetttttttttaaaaaaaaooooooonnnnnniiiiiihhhhhhssssssrrrrrlllldddduuuccmmwwyyffggppbvkjxqz` reflects about the average number of characters if the English language. It is used to decide for the factor which comes close to an effective average title length.<br>
`(font,size(string).width) * factor`

1. **Monospaced message  sections serup**<br>
The width of monospaced sections is defined by the longest line ending with a line break. When a monospaced section exceeds the maximum width it is reduced to it and a horizontal scrollbar is added.

1. **Reply buttons setup**<br>
The width of the widest reply buttons row defines the message window width. When the button area's width  exceeds the maximum width it is reduced to it and a horizontal scrollbar is added.


### Message window width adjustment
With the setup of the title, all monospaced sections and all buttons had been setup the message window's width has becomes final. When the width exceeds the specified maximum, all _[Monospaced](#property-monospaced)_ message height it is reduced to it and the reduction determines the height reduction of the message and/or the buttons area.

#### Message area height adjustment
- When the message area occupies less than 70% of the overall areas' height both, the message area and the buttons area is height reduced proportionally and will get a vertical scrollbar.
- When the message area occupies 70% or more of the overall areas' height and none of the message sections occupies 60% or more if the message area's height the whole message area's height it is reduced to fit and a vertical scrollbars is added.
- When one of the message sections occupies 60% or more of the message area's height only this one is reduced and gets a vertical scrollbar.

### Vertical frames adustment
- When the height of all message sections has become final their top positions are adjusted correspondingly.
- The reply buttons' height is adjusted to the maximum buttons height
- The reply button rows are height adjusted and top positioned
- At last the buttons area's top position is adjusted

### Scrollbar application and adjustment
Frames are the means for the application of scrollbars, becoming  required when the [content's height](#property-framecontentheight) or [width](#property-framecontentwidth) expands or when the frame size shrinks. Scrollbars may become obsolete when the content's height/width shrinks or the frame size expands.
When scrollbars are already applied only the width/height is adjusted to the content's width/height. When the content's width/height is no longer greater than the surrounding frame the scrollbars are removed.
  

## Properties
### Property Monospaced
Let/Get property with the argument _ctrl_ of type UserForms.Control which may be a TextBox, a TextBox-Frame, or a Section-Frame, defaults to False, may be assigned True when a monospaced section is setup.

Syntax: `Monospaced(ctrl) = True`

### Property _TextBoxWidth_
A Let-only property of a TextBox, when changed  triggers the property _[FrameWidth](#property-framewidth)_

### Property _FrameWidth_
A Let-only property with the arguments  _frame\_object_ and _child\_width_.

Syntax: `FrameWidth(frame_object, child_width) = new_frame`<br>Triggers the application or removal of a horizontal scrollbar depending on the content's width.

### Property _FrameHeight_
A Let-only property with the frame as argument.
Syntax: `FrameHeight(frame_object) = new_frame_height`<br>Triggers the application or removal of a vertical scrollbar depending on the content's height.

### Property _FrameContentHeight_
Get-only property of a _MsForm.Frame_ object, returns the height of it's content defined by:<br>`Max(FrameContentHeight, ctl.Top + ctl.Height)` which is the bottommost control.

Syntax: `FrameContentHeight(frame_object)`

### Property _FrameContentWidth_
Get-only property of an _MsForm.Frame_ object, returns the width of it's content defined by: <br>`Max(FrameContentWidth, ctl.Left + ctl.Width)`

Syntax: `FrameContentHeight(frame_object)`

## Procedures, code snippets, etc.
### Autosize Height only
```vbs
codePublic Sub AutoSizeHeight( _
                    ByRef as_ctl As Variant, _
                    ByVal as_text As String, _
           Optional ByVal as_width As Single = 0, _
           Optional ByVal as_height As Single = 0, _
           Optional ByVal as_append As Boolean = False)
' ------------------------------------------------------------------------------
' Autosizes the control (ctl) a TextBoxe or Label. When a width (as_width)
' is provided, the width is maintained and the height varies.
' When a height (as_height) is provided, the height is maintained and
' the width varies
' Note: A provided height is ignored when a width is provided!.
' ------------------------------------------------------------------------------
    Dim tbx As My forms.TextBox
    Dim lbl As My forms.Label
    
    Select TypeName(ctl)
        Case  "TextBox"
            Set tbx = ctl
            With as_tbx
                .MultiLine = True
                .WordWrap = True
                .AutoSize = False
                If as_width > 0 Then
                    .Width = as_width
                Else if as_height > 0 Then
                    .Height = as_height
                End If
                If Not as_append Then
                    .Value = as_text
                Else
                EndIf
                .AutoSize = True
            End With
        Case "Label"
            Set lbl = ctl
        Case Else
    End Select

End Sub
```
### Align vertical position to grid
When a TextBox or a Label is vertically misplaced the text may not appear as it should - which results in an ugly display. Aligning to the grid avoids this problem.

```vbs
Public Function VgridPos(ByVal si As Single) As Single
' --------------------------------------------------------------
' Returns a value which is the next position vertically down 
' which is aligned to a UserForm's grid.
' Background: A controls content is only properly displayed
' when the top position of it is aligned to such a position.
' --------------------------------------------------------------
    Dim i As Long
    
    For i = 0 To 6
        If Int(si) = 0 Then
            VgridPos = 0
        Else
            If Int(si) < 6 Then si = 6
            If (Int(si) + i) Mod 6 = 0 Then
                VgridPos = Int(si) + i
                Exit For
            End If
        End If
    Next i

End Function

code
```

## Test, test environment
### Test Worksheet
#### Test-Procedures Command button
for the mandatory test procedure
#### Regression-Test Command button
for the consecutive execution of test procedures marked for the Regression-Test

#### Initial test values
#### Display Frames Option
Shows the message with visible frame boundaries
#### Display Modeless Option
Allows interaction with the displayed message form in the Immediate Window - with the little disadvantage that pressed reply buttons are not recognized (each of them just closes the window).
### Explore Control