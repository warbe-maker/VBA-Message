# Common VBA Message Form and Display services

Supplements the README and the [Common-VBA-Message-Services](https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html) post focusing on technical aspects.

This implementation of a kind of unwith major flaws of the VBA.MsgBox eliminated:
* limited window width, resulting in a truncated title
* limited message text space
* limited reply button options (number caption text)
* no  mono-spaced, bold, italic, underlined, color, name, and size don't iptions


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
### Width Adjustment I
1. Minimum width
2. Width expansion by Title up to the specified maximum message form width
3. Width expansion by monospaced sections up to the specified maximum message form width. When a monospaced section exceeds the maximum width it is reduced to it and a horizontal scrollbar is added to it
4. Width expansion by reply buttons area. When the buttons area's width  exceeds the maximum width it is reduced to it and a horizontal scrollbar is added to it

With this first width adjustment the message form's width has become **final**.

### Height Adjustment
1. Any proportional spaced message sections use the final form width but still determine the overall height of the message window
2. When the overall message window height exceeds the maximum specified, the message area and/or the buttons area is reduced in its height and a vertical scrollbars is added. When one of the areas occupies 70% or more if the total height only this area is reduced, else both. When the message area is about to be height reduced any section occupying 65% or more will be reduced and a vertically scrollbars added, else the scrollbars is applied for the message area as a whole.

