## Design of the Message/UserForm

### General

The design of the _Message Form_ consists of
- 3 message sections, no matter which one is used
- 7 rows each with 7 reply _Buttons_ allowing any display from all in one to one in each row.

### Design of the controls

The message form is organized in a hierarchy of frames as follows.
````
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
    | | |  Message Section 3 (TextBox)  --     | |
    | | +------------------------------------+ | |
    | +----------------------------------------+ |
    +--------------------------------------------+
    +-- Buttons Area (Frame)  -------------------+
    | +-- Button Rows (Frame) -----------------+ |
    | | +-- Replies Row 1 (Frame) -----------+ | |
    | | | Reply Row 1 Button 1               | | |
    | | | Reply Row 1 Button 2               | | |
    | | | Reply Row 1 Button 3               | | |
    | | | Reply Row 1 Button 4               | | |
    | | | Reply Row 1 Button 5               | | |
    | | | Reply Row 1 Button 6               | | |
    | | | Reply Row 1 Button 7               | | |
    | +--------------------------------------+ | |
    |                      .                   | |
    |                      .                   | |
    |                      .                   | |
    | | +-- Replies Row 7 (Frame) -----------+ | |
    | | | Reply Row 7 Button 1               | | |
    | | | Reply Row 7 Button 2               | | |
    | | | Reply Row 7 Button 3               | | |
    | | | Reply Row 7 Button 4               | | |
    | | | Reply Row 7 Button 5               | | |
    | | | Reply Row 7 Button 6               | | |
    | | | Reply Row 7 Button 7               | | |
    | | +------------------------------------+ | |
    | +--------------------------------------+ | |    +--------------------------------------------+
````    
The [Implementation](#Implementation.md) is merely design driven. Not using control's name is achieved by storing  all controls (frames, text boxes, and command buttons) in collections by relying on the design rather than on control names. As a consequence, additional message sections and additional reply buttons are primarily a matter of a design change and require a minimum code change.