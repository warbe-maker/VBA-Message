
# Common VBA Message Service (a MsgBox Alternative)

[Abstract](#abstract)<br>
[Why an alternative MsgBox](#why-an-alternative-msgbox)<br>[Installation](#installation)<br>[Properties of the _fMsg_ UserForm](#properties-of-the-fmsg-userform)<br>[Usage](#usage)<br>[Interfaces](#interfaces)

### Abstract
A flexible and powerful VBA MsgBox alternative coming in four flavors. **[Dsply](#the-dsply-service)** is for any common message, **[ErrMsg](#the-errmsg-service)** provides a comprehensive error message, **[Box](#the-box-service)** is a very much VBA MsgBox like service and **[Progress](#the-progress-service)** is a service to display the progress of a process.

### Why an alternative MsgBox?
The alternative implementation addresses many of the MsgBox's deficiencies - without re-implementing it to 100%.

|MsgBox|Alternative|
|------|-----------|
| The message width and height is limited and cannot be altered | The&nbsp;maximum&nbsp;width and&nbsp;height&nbsp;is&nbsp;specified as&nbsp;a percentage of the screen&nbsp;size&nbsp; which&nbsp;defaults&nbsp;to: 80%&nbsp;width and  90%&nbsp;height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it) may be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a satisfying result | Up&nbsp;to&nbsp;3&nbsp; _Message&nbsp;Sections_ each with an optional _Message Text Label_ and a _Monospaced_ option allow an appealing design without any extra  effort |
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order (=49 buttons in total) |
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| vbApplicationModal or vbSystemModal, no vbModeless option | The message can be displayed both ways which _modal_ (the default) or _modeless_. _modal_ equals to vbApplicationModal, there is (yet) no vbSystemModal option.|
| Specifying the default button | The default button may be specified as index or as the displayed caption. However, it cannot be specified as vbOk, vbYes, vbNo, etc. |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

### Installation
1. Download [fMsg.frm][1], [fMsg.frx][2], and [mMsg.bas][3] .
2. Import _fMsg.frm_ and _mMsg.bas_ to your VB-Project
4. In the VBE add a Reference to _Microsoft Scripting Runtime_

### The Dsply service
The service provides all features which make the difference to the VBA.MsgBox.
#### Syntax
`mMsg.Dsply(dsply_title, dsply_msg[, dsply_buttons][, dsply_button_default][, dsply_reply_with_index][, dsply_modeless][, dsply_min_width][, dsply_max_width][, dsply_max_height][, dsply_min_button_width])`

The _Dsply_ service has these named arguments:

| Part          | Description                 |
|-------------------|-------------------------|
| dsply_title       | Required. String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.|
|dsply_msg            | Required. [UDT _TypeMsg_ ][#syntax-of-the-typemsgMsg-udt] expression providing 4 message sections, each with a label and the message text, displayed as the message in the dialog box. The maximum length of each of the four possible message text strings is only limited by the system's limit for string expressions which is about 1GB!. When one of the 4 message text strings consists of more than one line, they can be separated by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line.|
|dsply_buttons         | Optional. Variant expression. Defaults to vbOkOnly. May be provided as a comma delimited String, a Collection, or a Dictionary, with each item specifying a displayed command button's caption or a button row break (vbLf, vbCr, or vbCrLf). Any of the items may be a string or a classic VBA.MsgBox values (see [The VBA.MsgBox buttons argument settings][4]. Items exceeding 49 captions are ignored, when no row breaks are specified max 7 buttons are displayed in a row.|
|dsply_button_default  |_Long_ expression. Defaults to 1, specifies the index of the button which is the default button |
|dsply_reply_with_index|_Boolean_ expression. Defaults to False. When True the index if the pressed button is returned rather than its caption |
|dsply_modeless        |_Boolean_ expression. Defaults to False                   |
|dsply_min_width       |_Single_ expression. Defaults to 300 pt                   |
|dsply_max_width       |_Single_ expression. Defaults to 80% of the screen height |
|dsply_max_height      |_Single_ expression. Defaults to 75% of the screen height |
|dsply_min_button_width| _Single_ expression. Defaults to 70 pt when not specified|

#### Syntax of the _TypeMsg_ UDT
The syntax is described best as a code snippet using all options 
```
Dim Message As TypeMsg
With Message.Section(n)
    With .Label
        .FontBold = True
        .FontColor = rgbRed
        .FontItalic = True
        .FontName = "Tahoma"
        .FontSize = 9
        .FontUnderline = True
        .Monospaced = True ' FontName will be ignored and default to "Courier New"
        .Text As String
    End With
    With .Text
         .FontBold = True
        .FontColor = rgbRed
        .FontItalic = True
        .FontName = "Tahoma"
        .FontSize = 9
        .FontUnderline = True
        .Monospaced = True ' FontName will be ignored and default to "Courier New"
        .Text As String
```
Going with the defaults the minimum message text assignment (without a label) would be `Message.Section(1).Text.Text = "......"`

### The ErrMsg service
Provides the display of a well designed error message by integrating a debugging option which supports the Resume of the code line which caused the error.

#### Using the _mMsg.ErrMsg_ service

The following is a coding example which my personal standard. It uses an ErrSrc function which is module specific and returns '\<modulename>.\<procedurename>'.
```VB
Public Sub Test_ErrMsg_Service()
    Const PROC = "Test_ErrMsg_Service"
    
    On Error GoTo eh
    Dim i As Long
    i = i / 0
    
xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub
```

Displays:

![](images/ErrorMessage.jpg)

Only when the Conditional Compile Argument 'Debugging = 1' the ErrMsg is displayed with Yes/No buttons and thus may return vbYes which means that the line which caused the error may be resumed by F8, F8.


### The Box service
The _Box_ service mimicks the VBA.MsgBox. In contrast to the _Dsply_ service the messsage text is a simple string expression just like the VBA.MsgBoc _Prompt_ argument. All other arguments are identical with the _Dsply_ service - just prefixed with box_ instead of dsply_.

#### Using the _mMsg.Box_ service
```
Public Sub Test_Box()
    Select Case mMsg.Box(title:="Any title" _
                       , prompt:="Any message" _
                       , buttons:=vbYesNoCancel)
        Case vbYes:     MsgBox "Button 'Yes' clicked"
        Case vbNo:      MsgBox "Button 'No' clicked"
        Case vbCancel:  MsgBox "Button 'Cancel' clicked"
    End Select

    ' or alternatively 
    mMsg.Box title:="Any title" _
           , prompt:="Any message" _
           , buttons:=vbYesNoCancel
    Select Case mMsg.RepliedWith
        Case vbYes:     MsgBox "Button 'Yes' clicked"
        Case vbNo:      MsgBox "Button 'No' clicked"
        Case vbCancel:  MsgBox "Button 'Cancel' clicked"
    End Select

End Sub
```

#### Using the mMsg.Dsply_ service
The following is a demonstration of how to use many of the features.
```
Public Sub DemoMsgDsplyService_1()
    Const MAX_WIDTH     As Long = 50
    Const MAX_HEIGHT    As Long = 60
    
    Dim cll             As New Collection
    Dim i, j            As Long
    Dim Message         As TypeMsg
   
    With Message.Section(1)
        .Label.Text = "Demonstration overview:"
        .Label.FontColor = rgbBlue
        .Text.Text = "- Use of all 4 message sections" & vbLf _
                   & "- All sections with a label" & vbLf _
                   & "- One section monospaced exceeding the specified maximum message form width" & vbLf _
                   & "- Use of some of the 7x7 reply buttons in a 4-4-1 order" & vbLf _
                   & "- An an example for available text font options all labels in blue"
    End With
    With Message.Section(2)
        .Label.Text = "Unlimited message width!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section's text is mono-spaced (which is not word-wrapped) and the maximimum message form width" & vbLf _
                   & "for this demo has been specified " & MAX_WIDTH & "% of the sreen width (the default would be 80%)" & vbLf _
                   & "the text is displayed with a horizontal scrollbar. There is no message size limit for the display despite the" & vbLf & vbLf _
                   & "limit of VBA for text strings  which is about 1GB!"
        .Text.Monospaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section lext has many lines (line breaks)" & vbLf _
                   & "the default word-wrapping for this proportional-spaced text" & vbLf _
                   & "has not the otherwise usuall effect. The message area thus" & vbLf _
                   & "exeeds the for this demo specified " & MAX_HEIGHT & "% of the screen size" & vbLf _
                   & "(defaults to 80%) it is displayed with a vertical scrollbar." & vbLf _
                   & "So even a proportional spaced text's size - which usually is word-wrapped -" & vbLf _
                   & "is only limited by the system's limit for a String which is abut 1GB !!!"
    End With
    With Message.Section(4)
        .Label.Text = "Great reply buttons flexibility:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 49 possible reply buttons (7 rows by 7 buttons). " _
                   & "It also shows that a reply button can have any caption text and the buttons can be " _
                   & "displayed in any order within the 7 x 7 limit. Of cource the VBA.MsgBox classic " _
                   & "vbOkOnly, vbYesNoCancel, etc. are also possible - even in a mixture." & vbLf & vbLf _
                   & "By the way: This demo ends only with the Ok button clicked and loops with all the ohter."
    End With
    '~~ Prepare the buttons collection
    For j = 1 To 2
        For i = 1 To 4
            cll.Add "Multiline reply" & vbLf & "button caption" & vbLf & "Button-" & j & "-" & i
        Next i
        cll.Add vbLf
    Next j
    cll.Add vbOKOnly ' The reply when clicked will be vbOK though
    
    While mMsg.Dsply(dsply_title:="Usage demo: Full featured multiple choice message" _
                   , dsply_msg:=Message _
                   , dsply_buttons:=cll _
                   , dsply_max_height:=MAX_HEIGHT _
                   , dsply_max_width:=MAX_WIDTH _
                    ) <> vbOK
    Wend
    
End Sub
```
which displays:

![](images/demo-1.png)

#### Proportional versus Mono-Spaced
##### _Monospaced_ = True

Because the text is ++not++  "wrapped" the width of the _Message Form_ is determined by the longest text line (up to the _Maximum Form Width_ specified). When the maximum width is exceeded a vertical scroll bar is applied.<br>Note: The title and the broadest _Button Row_ May still determine an even broader final _Message Form_.

##### _Monospaced_ = False (default)
Because the text is "wrapped" the width of a proportional-spaced text is determined by the current form width.<br>Note: When a message is displayed exclusively proportional-spaced the _Message Form_ width is determined by the length of the title, the required space for the broadest _Buttons Row_ and the specified _Minimum Form Width_.


[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/edit/master/source/fMsg.frm
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/edit/master/source/fMsg.frx
[3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/edit/master/source/mMsg.bas
[4]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function
