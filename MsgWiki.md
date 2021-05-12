# Common VBA Message Form and Display services

Supplements the README and the [Common-VBA-Message-Services](https://warbe-maker.github.io/vba/common/2020/11/17/Common-VBA-Message-Services.html) post focusing on technical aspects.

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


### Message window height adjustment
With the setup of the title, all monospaced sections, and the buttons the message window width has become final. With the setup of all proportional spaced message sections the message window's height becomes final. When the height exceeds the specifyed maximum message height it is reduced to it and the reduction determines the height reduction of the message and/or the buttons area.

#### Message area height adjustment
- When the message area occupies less than 70% of the overall areas' height both, the message area and the buttons area is height reduced proportionally and will get a vertical scrollbar.
- When the message area occupies 70% or more of the overall areas' height and none of the message sections occupies 60% or more if the message area's height the whole message area's height it is reduced to fit and a vertical scrollbars is added.
- When one of the message sections occupies 60% or more of the message area's height only this one is reduced and gets a vertical scrollbar.

### Vertical frames adustment
- When the height of all message sections has become final their top positions are adjusted correspondingly.
- The reply buttons' height is adjusted to the maximum buttons height
- The reply button rows are height adjusted and top positioned
- At last the buttons area's top position is adjusted
