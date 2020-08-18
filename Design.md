## Design of the Message/UserForm

### General

The design of the _Message Form_ allows 3 message sections and 7 reply _Buttons_. The reply _Buttons_ may all be ordered in _Replies Row_ 1 or 1 in each of the 7 rows, and any other desired approach between these two.

### Organization/Design of the controls

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
 
The [Implementation](#Implementation.md) is merely design driven and avoids using any control's name. This is achieved by collection of the controls (frames, text boxes, and command buttons) by the design rather than by controls name. As a consequence, additional message sections and additional reply buttons are primarily a matter of design change.