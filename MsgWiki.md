# Common VBA Message Form and Display services

Supplements the README and the [Common-VBA-Message-Services](https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html) post focusing on technical/implementation aspects.

### General

The design of the _Message Form_ consists of
- 4 message sections, no matter which one is used
- 7 rows each with 7 reply _Buttons_ allowing to display them in any desired design, from all 7 in one row to 7 rows with one button.

### The design of the message form
The message form is organized in a hierarchy of frames as follows.
````
    +----------- Message Area Frame -------------+
    | +------- Message Section 1 Frame --------+ |
    | |   Message Section 1 Label              | |
    | | +--- Message Section 1 Text Frame ---+ | |
    | | | Message Section 1 TextBox          | | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    |                    .                       |
    |                    .                       |
    |                    .                       |
    | +------- Message Section 4 Frame --------+ |
    | |   Message Section 4 Label              | |
    | | +--- Message Section 4 Text Frame ---+ | |
    | | | Message Section 4 TextBox          | | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    +--------------------------------------------+
    +----------- Buttons Area Frame -------------+
    | +--------- Button Rows Frame ------------+ |
    | | +------- Replies Row 1 Frame --------+ | |
    | | |    Button 1, 2, 3, 4, 5, 6, 7      | | |
    | | +------------------------------------+ | |   
    | |                    .                   | |
    | |                    .                   | |
    | |                    .                   | |
    | | +------- Replies Row 7 Frame --------+ | |
    | | |    Button 1, 2, 3, 4, 5, 6, 7      | | |
    | | +------------------------------------+ | |   
    | +--------------------------------------+ | |
    +--------------------------------------------+
````    
The implementation is merely design driven. I.e. with the exception of the button's click events and the Label's click events none of the controls is called by its name. This is managed by collecting all controls (Frames, Labels, TextBoxes, and CommandButtons) in Dictionaries by relying on the hierarchical design. As a consequence, addiing message sections is mostly a matter of a design change and requires very little code change.

### Label and TextBox controls for each section
It turned out that a mixture (or combination) of both has the advantage that a Label can have a Click event and this is used to open optionally specified anything to be opened: a html page, a file, or even an application. 

## Implementation of the display services
### Message window width adjustment
The setup of the message form starts width the default or explicitely provided _Minimum Message Width_ and continues by expanding it up to the default or explicitely provided _Maximum Message Width_ by performing the following steps.
1. **Title setup**<br>
The width of the message form is adjusted to the width required to display the title untruncted

1. **Monospaced message  sections serup**<br>
The width of the message form is triggered by the line lenght of monospaced sections. When a monospaced section exceeds the maximum width it is provided with a horizontal scroll-bar.

1. **Reply buttons setup**<br>
The width of the buttons area (and so the message form widht) is triggered by the widest reply buttons row. When the button area's width exceeds the maximum width it is provided with a horizontal scroll-bar.

1. **Proportially spaced message sections**<br>
are adjusted with their width to the final message form width which resulted from the above setup steps.

#### Message area height adjustment
- When the message area occupies less than 70% of the overall areas' height both, the message area and the buttons area is height reduced proportionally and will get a vertical scrollbar.
- When the message area occupies 70% or more of the overall areas' height and none of the message sections occupies 60% or more if the message area's height the whole message area's height it is reduced to fit and a vertical scrollbars is added.
- When one of the message sections occupies 60% or more of the message area's height only this one is reduced and gets a vertical scrollbar.


## Procedures, code snippets, etc.
### Autosize TextBox width or height
```vb
Public Sub AutoSizeTextBox(ByRef as_tbx As MSForms.TextBox, _
                           ByVal as_text As String, _
                  Optional ByVal as_width_limit As Single = 0, _
                  Optional ByVal as_append As Boolean = False, _
                  Optional ByVal as_append_margin As String = vbNullString)
' ------------------------------------------------------------------------------
' Common AutoSize service for an MsForms.TextBox providing a width and height
' for the TextBox (as_tbx) by considering:
' - When a width limit is provided (as_width_limit > 0) the width is regarded a
'   fixed maximum and thus the height is auto-sized by means of WordWrap=True.
' - When no width limit is provided (the default) WordWrap=False and thus the
'   width of the TextBox is determined by the longest line.
' - When a maximum width is provided (as_width_max > 0) and the parent of the
'   TextBox is a frame a horizontal scrollbar is applied for the parent frame.
' - When a maximum height is provided (as_heightmax > 0) and the parent of the
'   TextBox is a frame a vertical scrollbar is applied for the parent frame.
' - When a minimum width (as_width_min > 0) or a minimum height (as_height_min
'   > 0) is provided the size of the textbox is set correspondingly. This
'   option is specifically usefull when text is appended to avoid much flicker.
'
' Uses: AdjustToVgrid
'
' W. Rauschenberger Berlin April 2022
' ------------------------------------------------------------------------------
    
    With as_tbx
        .MultiLine = True
        If as_width_limit > 0 Then
            '~~ AutoSize the height of the TextBox considering the limited width
            '~~ (applied for proportially spaced text where the width determines the height)
            .WordWrap = True
            .AutoSize = False
            .Width = as_width_limit - 2 ' the readability space is added later
            
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & as_append_margin & vbLf & as_text
                End If
            End If
            .AutoSize = True
        Else
            '~~ AutoSize the height and width of the TextBox
            '~~ (applied for mono-spaced text where the longest line defines the width)
            .MultiLine = True
            .WordWrap = False ' the means to limit the width
            .AutoSize = True
            If Not as_append Then
                .Value = as_text
            Else
                If .Value = vbNullString Then
                    .Value = as_text
                Else
                    .Value = .Value & vbLf & as_text
                End If
            End If
        End If
        .Width = .Width + 2 ' readability space
        .Height = AdjustToVgrid(.Height, 0)
    End With
        
xt: Exit Sub

End Sub
```
### Adjust a value to the vertical grid of the message form
The value, when used for the top position and/or the height of a control which contains text ensures a correct display of it. The font may not be displayed in its correct size otherwise.

```vb
Public Function AdjustToVgrid(ByVal atvg_si As Single, _
                     Optional ByVal atvg_threshold As Single = 1.5, _
                     Optional ByVal atvg_grid As Single = 6) As Single
' -------------------------------------------------------------------------------
' Returns the value (atvg_si) as a Single value which is a multiple of the grid
' value (atvg_grid), which defaults to 6. To avoid irritating vertical spacing
' a certain threshold (atvg_threshold) is considered which defaults to 1.5.
' The returned value can be used to vertically align a control's top position to
' the grid or adjust its height to the grid.
' Examples for the function of the threshold:
'  7.5 < si >= 0   results to 6
' 13.5 < si >= 7.5 results in 12
' -------------------------------------------------------------------------------
    AdjustToVgrid = (Int((atvg_si - atvg_threshold) / atvg_grid) * atvg_grid) + atvg_grid
End Function
```

## Test an test environment
A test worksheet is designed to execute single tests or all consecutively as a regression test. The result of a regression test is saved in the form of an execution trace by using the [Common VBA Execution Trace Service][1]


[1]:https://github.com/warbe-maker/Common-VBA-Execution-Trace-Service