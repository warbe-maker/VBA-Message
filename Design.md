## Design of the Message/UserForm

The design of the _Message Form_ allows 3 message sections and 7 _Reply Buttons_. The _Reply Buttons_ may all be ordered in _Replies Row 1_ or all 7 underneath, each in one of the 7 rows - or any kind of meaningful order in between.

The message form is organized in a hierarchy of frames as follows.

    +-- Message Area (Frame) --------------------+
    | +-- Message Section 1 (Frame) -----------+ |
    | | Message Section 1 Label (Label)        | |
    | | +-- Message Section 1 Text (Frame) --+ | |
    | | |  Message Section 1 (TextBox)       | | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    | +-- Message Section 2 (Frame) -----------+ |
    | | Message Section 2 Label (Label)        | |
    | | +-- Message Section 2 Text (Frame) --+ | |
    | | |  Message Section 2 (TextBox)       | | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    | +-- Message Section 3 (Frame) -----------+ |
    | | Message Section 3 Label (Label)        | |
    | | +-- Message Section 3 Text (Frame) --+ | |
    | | |  Message Section 3 (TextBox)       | | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    +--------------------------------------------+
    +-- Reply Area (Frame)   --------------------+
    | +-- Replies Row 1 (Frame) ---------------+ |
    | | Reply Row 1 Button 1 (CommandButtons)  | |
    | | Reply Row 1 Button 2 (CommandButtons)  | |
    | | Reply Row 1 Button 3 (CommandButtons)  | |
    | | Reply Row 1 Button 4 (CommandButtons)  | |
    | | Reply Row 1 Button 5 (CommandButtons)  | |
    | | Reply Row 1 Button 6 (CommandButtons)  | |
    | | Reply Row 1 Button 7 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 2 (Frame) ---------------+ |
    | | Reply Row 2 Button 1 (CommandButtons)  | |
    | | Reply Row 2 Button 2 (CommandButtons)  | |
    | | Reply Row 2 Button 3 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 3 (Frame) ---------------+ |
    | | Reply Row 3 Button 1 (CommandButtons)  | |
    | | Reply Row 3 Button 2 (CommandButtons)  | |
    | | Reply Row 3 Button 3 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 4 (Frame) ---------------+ |
    | | Reply Row 4 Button 1 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 5 (Frame) ---------------+ |
    | | Reply Row 5 Button 1 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 6 (Frame) ---------------+ |
    | | Reply Row 6 Button 1 (CommandButtons)  | |
    | +----------------------------------------+ |
    | +-- Replies Row 7 (Frame) ---------------+ |
    | | Reply Row 7 Button 1 (CommandButtons)  | |
    | +----------------------------------------+ |
    +--------------------------------------------+
 
The controls (frames, text boxes, and command buttons) are collected with the message form's initialization and used throughout the [Implementation](#implementation.md). I.e. the whole approach is merely design driven.

The design allows the following orders og the _Reply Buttons_:

| Scheme |      |
| - | - |
| 1 2 3 4 5 6 7 | 1 - 7 in one row |
| 1 2 3 4<br> &nbsp; 5 6 7 | All in two rows|
| 1 2 3<br>4 5 6<br>&nbsp;&nbsp;&nbsp;7 | All in 3 rows|
| 1 2<br>3 4<br>5 6<br>&nbsp;7 | All in 4 rows |
| 1<br>2<br>3<br>4<br>5<br>6<br>7 | 1 - 7 underneath |
