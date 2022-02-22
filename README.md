# Common VBA Message Service (a MsgBox Alternative)

<!-- Start Document Outline -->

* [Summary](#summary)
* [Why an alternative MsgBox?](#why-an-alternative-msgbox)
* [Installation](#installation)
* [Usage](#usage)
	* [The Box service](#the-box-service)
		* [Syntax](#syntax)
		* [Using the Box service](#using-the-box-service)
	* [The Dsply service](#the-dsply-service)
		* [Syntax](#syntax-1)
		* [Syntax of the TypeMsg UDT](#syntax-of-the-typemsg-udt)
		* [Using the Dsply service](#using-the-dsply-service)
	* [The ErrMsg service](#the-errmsg-service)
		* [Syntax](#syntax-2)
		* [Usage example](#usage-example)
	* [The Monitor service](#the-monitor-service)
		* [Usage of the Monitor service](#usage-of-the-monitor-service)
	* [The Buttons service](#the-buttons-service)
* [Other aspects](#other-aspects)
	* [Proportional versus Mono-spaced](#proportional-versus-mono-spaced)
	* [Unambiguous procedure name](#unambiguous-procedure-name)
	* [Multiple Monitor instances](#multiple-monitor-instances)

<!-- End Document Outline -->

## Summary
A flexible and powerful `VBA.MsgBox` alternative providing four specific services:
- ***[Box](#the-box-service)*** as a 'VBA.MsgBox` alike service with extended flexibility and no title and string length limits
- ***[Dsply](#the-dsply-service)*** as a multi purpose message display service
- ***[ErrMsg](#the-errmsg-service)*** for the display of a well designed error message
- ***[Monitor](#the-monitor-service)*** as a service to display the ongoing progress of a process.

## Why an alternative MsgBox?
The alternative implementation addresses many of the MsgBox's deficiencies - without re-implementing it to 100%.

|MsgBox|Alternative|
|------|-----------|
| The message width and height is limited and cannot be altered | The&nbsp;maximum&nbsp;width and&nbsp;height&nbsp;is&nbsp;specified as&nbsp;a percentage of the screen&nbsp;size&nbsp; which&nbsp;defaults&nbsp;to: 80%&nbsp;width and  90%&nbsp;height (hardly ever used)|
| When a message exceeds the (hard to tell) size limit it is truncated | When the maximum size is exceeded a vertical and/or a horizontal scroll bar is applied
| The message is displayed with a proportional font | A message may (or part of it) may be displayed mono-spaced |
| Composing a fair designed message is time consuming and it is difficult to come up with a satisfying result | Up&nbsp;to&nbsp;3&nbsp; _Message&nbsp;Sections_ each with an optional _Message Text Label_ and a _Mono-spaced_ option allow an appealing design without any extra  effort |
| The maximum reply _Buttons_ is 3 | Up to 7 reply _Buttons_ may be displayed in up to 7 reply _Button Rows_ in any order (=49 buttons in total) |
| The caption of the reply _Buttons_ is specified by a [value](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) which results in 1 to 3 reply _Buttons_ with corresponding untranslated! native English captions | The caption of the reply _Buttons_ may be specified by the [VB MsgBox values](<https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function#settings>) **and** additionally by any multi-line text (see [Syntax of the _buttons_ argument](#syntax-of-the-buttons-argument) |
| vbApplicationModal or vbSystemModal, no vbModeless option | The message can be displayed both ways which _modal_ (the default) or _modeless_. _modal_ equals to vbApplicationModal, there is (yet) no vbSystemModal option.|
| Specifying the default button | The default button may be specified as index or as the displayed caption. However, it cannot be specified as vbOk, vbYes, vbNo, etc. |
| Display of an alert image (?, !, etc.) | (yet) not implemented |

## Installation
1. Download [fMsg.frm][1], [fMsg.frx][2], and [mMsg.bas][3] .
2. Import _fMsg.frm_ and _mMsg.bas_ to your VB-Project
4. In the VBE add a Reference to _Microsoft Scripting Runtime_
Note: The

## Usage
### The _Box_ service
The _Box_ service mimics the _VBA.MsgBox_ by displaying a single message string like the _VBA.MsgBox Prompt_ argument. However, due to the use of the _fMsg_ form there is no limit in the length of the message string but the systems limit which is about 1GB. With the exception of the box_msg argument all other arguments are identical w\ith the __Dsply_ service - just prefixed with box_ instead of dsply_.

#### Syntax
The _Box_ service has these named arguments:

| Argument               | Meaning                                              |
| ---------------------- | -----------------------------------------------------|
| `box_title`            | String expression displayed in the window handle bar |
| `box_msg`              | String expression displayed |
| `box_monospaced`       | Boolean expression, defaults to False, displays the `box_msg` with a monospaced font |
| `box_buttons`           | Optional. Variant expression. Defaults to vbOkOnly. May be provided as a comma delimited String, a Collection, or a Dictionary, with each item specifying a displayed command button's caption or a button row break (vbLf, vbCr, or vbCrLf). Any of the items may be a string or a classic VBA.MsgBox values (see [The VBA.MsgBox buttons argument settings][4]. Items exceeding 49 captions are ignored, when no row breaks are specified max 7 buttons are displayed in a row. |
| `box_button_default`    | Optional, numeric expression, defaults to 1, identifies the default button, i.e. the button which has the focus
| `box_returnindex`       | Optional, boolean expression, default to False, indicates that the return value for the clicked button will be the index rather than its caption string.
| `box_width_min`         | Optional, numeric expression, defaults to 400, the minimum width in pt for the display of the message. A value < 100 is interpreted as % of the screen size, a value > 100 as pt
| `box_width_max`         | Optional, numeric expression, defaults to 80, specifies the maximum message window width as % of the screen size. A value < 100 is interpreted as % of the screen size, a value > 100 as pt
| `box_height_max`        | Optional, numeric expression, defaults to 70, specifies the maximum message window height of the screen size. A value < 100 is interpreted as % of the screen size, a value > 100 as pt
| `box_buttons_width_min` | Optional, numeric expression, defaults to 70, specifies the minimum button width in pt |

#### Using the _Box_ service
```
Public Sub Demo_Box_Service()
    Const PROC          As String = "Demo_Box_service"
    Const BTTN_1        As String = "Button-1 caption"
    Const BTTN_2        As String = "Button-2 caption"
    Const BTTN_3        As String = "Button-3 caption"
    Const BTTN_4        As String = "Button-4 caption"
    Const DEMO_TITLE    As String = "Demonstration of the Box service"
    
    On Error GoTo eh
    Dim DemoMessage     As String
    
    DemoMessage = "The message : The ""Box"" service displays one string just like the VBA MsgBox. However, the mono-spaced" & vbLf & _
                  "              option allows a better layout for an indented text like this one for example. It should also be noted" & vbLf & _
                  "              that there is in fact no message width limit." & vbLf & _
                  "The buttons : 7 buttons in 7 rows are possible each with any caption string or a VBA MsgBox value. The latter may" & vbLf & _
                  "              result in more than one button, e.g. vbYesNoCancel." & vbLf & _
                  "The window  : When the message exceeds the specified maximum width a horizontal scroll-bar, when it exceeds" & vbLf & _
                  "              the specified maximum height a vertical scroll.bar is displayed  the message is displayed with a horizontal scroll-bar." & vbLf
    
    Select Case mMsg.Box( _
             box_title:=DEMO_TITLE _
           , box_msg:=DemoMessage _
           , box_monospaced:=True _
           , box_width_max:=50 _
           , box_buttons:=mMsg.Buttons(BTTN_1, BTTN_2, BTTN_3, BTTN_4, vbLf, vbYesNoCancel) _
           , box_button_default:=5 _
            )
        Case BTTN_1:    MsgBox """" & BTTN_1 & """ pressed"
        Case BTTN_2:    MsgBox """" & BTTN_2 & """ pressed"
        Case BTTN_3:    MsgBox """" & BTTN_3 & """ pressed"
        Case BTTN_4:    MsgBox """" & BTTN_4 & """ pressed"
        Case vbYes:     MsgBox """ Yes"" pressed"
        Case vbNo:      MsgBox """No"" pressed"
        Case vbCancel:  MsgBox """Cancel"" pressed"
    End Select

xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub
```
The above code displays
![](images/Demo-Box-Service.jpg)

### The _Dsply_ service
The service provides all features which make the difference to the VBA.MsgBox.

#### Syntax
`mMsg.Dsply(dsply_title, dsply_msg[, dsply_buttons][, dsply_button_default][, dsply_reply_with_index][, dsply_modeless][, dsply_min_width][, dsply_max_width][, dsply_max_height][, dsply_min_button_width])`

The _Dsply_ service has these named arguments:

| Part                        | Description             |
|-----------------------------|-------------------------|
| `dsply_title`             | Required. String expression displayed in the title bar of the dialog box. If you omit title, the application name is placed in the title bar.|
| `dsply_msg`               | Required. [UDT _TypeMsg_ ][#syntax-of-the-typemsgMsg-udt] expression providing 4 message sections, each with a label and the message text, displayed as the message in the dialog box. The maximum length of each of the four possible message text strings is only limited by the system's limit for string expressions which is about 1GB!. When one of the 4 message text strings consists of more than one line, they can be separated by using a carriage return character (Chr(13)), a linefeed character (Chr(10)), or carriage return - linefeed character combination (Chr(13) & Chr(10)) between each line.|
| `dsply_buttons`           | Optional. Variant expression. Defaults to vbOkOnly. May be provided as a comma delimited String, a Collection, or a Dictionary, with each item specifying a displayed command button's caption or a button row break (vbLf, vbCr, or vbCrLf). Any of the items may be a string or a classic VBA.MsgBox values (see [The VBA.MsgBox buttons argument settings][4]. Items exceeding 49 captions are ignored, when no row breaks are specified max 7 buttons are displayed in a row.|
| `dsply_button_default`   | Optional, _Long_ expression, defaults to 1, specifies the index of the button which is the default button. |
| `dsply_reply_with_index` | Optional, _Boolean_ expression, defaults to False. When True the index if the pressed button is returned rather than its caption. |
| `dsply_modeless`          | Optional, _Boolean_ expression, defaults to False. When True the message is displayed modeless.  |
| `dsply_width_min`        | Optional, _Single_ expression, defaults to 300 which interpreted as pt.                   |
| `dsply_width_max`        | Optional, _Single_ expression, Defaults to 80 which interpreted as % of the screen's width. |
| `dsply_height_max`       | Optional, _Single_ expression, defaults to 75 which is interpreted as % of the screen's height.|
| `dsply_button_width_min` | Optional,  _Single_ expression, defaults to 70 pt. Specifies the minimum width of the reply buttons, i.e. even when the displayed string is just Ok, Yes, etc. which would result in a button with much less width. |

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

#### Using the _Dsply_ service
The below code demonstrates most of the available features and message window properties.
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
                   & "- One section mono-spaced exceeding the specified maximum message form width" & vbLf _
                   & "- Use of some of the 7 x 7 reply buttons in a 4-4-1 order" & vbLf _
                   & "- An an example for available text font options all labels in blue"
    End With
    With Message.Section(2)
        .Label.Text = "Unlimited message width!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section's text is mono-spaced (which is not word-wrapped) and the maximum message form width" & vbLf _
                   & "for this demo has been specified " & MAX_WIDTH & "% of the screen width (the default would be 80%)" & vbLf _
                   & "the text is displayed with a horizontal scroll-bar. There is no message size limit for the display despite the" & vbLf & vbLf _
                   & "limit of VBA for text strings  which is about 1GB!"
        .Text.Monospaced = True
    End With
    With Message.Section(3)
        .Label.Text = "Unlimited message height!:"
        .Label.FontColor = rgbBlue
        .Text.Text = "Because this section text has many lines (line breaks)" & vbLf _
                   & "the default word-wrapping for this proportional-spaced text" & vbLf _
                   & "has not the otherwise usual effect. The message area thus" & vbLf _
                   & "exceeds the for this demo specified " & MAX_HEIGHT & "% of the screen size" & vbLf _
                   & "(defaults to 80%) it is displayed with a vertical scroll-bar." & vbLf _
                   & "So even a proportional spaced text's size - which usually is word-wrapped -" & vbLf _
                   & "is only limited by the system's limit for a String which is abut 1GB !!!"
    End With
    With Message.Section(4)
        .Label.Text = "Great reply buttons flexibility:"
        .Label.FontColor = rgbBlue
        .Text.Text = "This demo displays only some of the 49 possible reply buttons (7 rows by 7 buttons). " _
                   & "It also shows that a reply button can have any caption text and the buttons can be " _
                   & "displayed in any order within the 7 x 7 limit. Of course the VBA.MsgBox classic " _
                   & "vbOkOnly, vbYesNoCancel, etc. are also possible - even in a mixture." & vbLf & vbLf _
                   & "By the way: This demo ends only with the Ok button clicked and loops with all the other."
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
                   , dsply_height_max:=MAX_HEIGHT _
                   , dsply_width_max:=MAX_WIDTH _
                    ) <> vbOK
    Wend
    
End Sub
```
which displays:

![](images/demo-1.png)

### The _ErrMsg_ service
Provides the display of a well designed error message by supporting a debugging option enabled with _Conditional Compile Argument_  `Debugging = 1` which displays an extra ***Resume Error Line*** button.
#### Syntax
`mMsg.ErrMsg(proc-name)`
Note: All other information about the error is obtained from the `err` object.

#### Usage example
```VB
Public Sub Test_ErrMsg_Service()
    Const PROC = "Test_ErrMsg_Service"
    
    On Error GoTo eh
    Dim i As Long
    i = i / 0
    
xt: Exit Sub

eh: Select Case mMsg.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      Goto xt
    End Select
End Sub
```

Displays:<br>
![](images/Demo-ErrMsg-Service.jpg)

Only when the Conditional Compile Argument 'Debugging = 1' the ErrMsg is displayed with Yes/No buttons and thus may return vbYes which means that the line which caused the error may be resumed by F8, F8.

### The _Monitor_ service
The _Monitor_ service has the following named arguments

| Part                  | Description             |
|-----------------------|-------------------------|
| `mntr_title`          | _String_ expression, displayed as title of the message window. |
| `mntr_msg`            | _String_ expression, displayed as the message/information. |
| `mntr_header`         | _String_ expression, optional, defaults to `vbNullString`, displayed abovr ther `mntr_msg`,  |
| `mntr_footer`         | _String_ expression, defaults to "Process in progress! Please wait.", displayed below `mntr_msg` |
| `mntr_msg_append`     | _Boolean_ expression, defaults to True. Appends the `mntr_msg` to the current displayed message string. |
| `mntr_msg_monospaced` | _Boolean_ expression, defaults to False. When True the message string is displayed with a mono-spaced font. |
| `mntr_width_min`      | _Long_ expression, defaults to 400, which interpreted as pt. |
| `mntr_width_max`      | _Long_ expression, defaults to 80, which is interpreted as % of the screen's width. |
| `mntr_height_max`     | |

#### Usage of the _Monitor_ service
The code below
```vb
Public Sub Demo_Monitor_Service()
    Const PROC              As String = "Demo_Monitor_Service"
    Const MONITOR_HEADER    As String = " No. Status   Step"
    Const MONITOR_FOOTER    As String = "Process finished! Close this window"
    Const PROCESS_STEPS     As Long = 12
    
    On Error GoTo eh
    Dim i               As Long
    Dim lWait           As Long
    Dim MonitorTitle    As String
    Dim ProgressStep    As String
    
    MonitorTitle = "Demonstration of the monitoring of a process step by step"
    mMsg.Form MonitorTitle, frm_unload:=True ' Ensure there is no process monitoring with this title still displayed
        
    For i = 1 To PROCESS_STEPS
        '~~ Preparing a process step message string
        ProgressStep = mBasic.Align(i, 4, AlignRight, " ") & _
                   mBasic.Align("Passed", 8, AlignCentered, " ") & _
                   Repeat(repeat_n_times:=Int(((i - 1) / 10)) + 1, repeat_string:="  " & _
                   mBasic.Align(i, 2, AlignRight) & _
                   ".  Follow-Up line after " & _
                   Format(lWait, "0000") & _
                   " Milliseconds.")
        
        If i < PROCESS_STEPS Then
            '~~ Steps 1 to n - 1
            mMsg.Monitor mntr_title:=MonitorTitle _
                       , mntr_msg:=ProgressStep _
                       , mntr_msg_monospaced:=True _
                       , mntr_header:=MONITOR_HEADER
            
            '~~ Simmulation of a process
            lWait = 100 * i
            DoEvents
            Sleep 200
        
        Else
            '~~ The last step, separated in order to display the footer along with it
            mMsg.Monitor mntr_title:=MonitorTitle _
                       , mntr_msg:=ProgressStep _
                       , mntr_header:=MONITOR_HEADER _
                       , mntr_footer:=MONITOR_FOOTER
        End If
    Next i
    
xt: Exit Sub

eh: If mMsg.ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Sub
```
displays:<br>
![](images/Demo-Monitor-Service.gif)

### The _Buttons_ service
Eases the provision of any number of buttons by allowing to specify them through a ParamArray which may be a mixture of  strings and numeric (VBA.MsgBox) values. The service allows to specify buttons as well as to add buttons to specified ones. The service ensures a maximum of 7 buttons in 7 rows by ignoring any exceeding button without notice. When no row break items (vbLf, vbCrLf, or vbCr) are included the service includes those break after each 7 buttons in a row.  
```vb
Dim cll As Collection
mMsg.Buttons cll, "A", "B", vbOkOnly
```
returns the the items "A", "B", and vbOkOnly in the Collection which results in the buttons "A", "B", "Ok" in the displayed message.
```
Dim cll As Collection
mMsg.Buttons cll, "A", "B", vbOkOnly
Set cll = mMsg.Buttons(cll, "C", "D") ' returns the buttons "C", "D" added to the buttons "A", "B", vbOkOnly
```
Same as above but the items "C" and "D" are added.

### The _MsgInstance_ service
All services create an instance of the _fMsg_ userForm with the title as the key.<br>
Syntax: `MsgInstance(title, unload)`<br>
`unload` defaults to False. When True an already existing instance is unloaded.
The instance object is kept in a Dictionary with the title as the key and the instance object as the item. When no item with the given title exists the instance is created, stored in the Dictionary and returned.

If an item exists in the Dictionary which is no longer loaded it is removed from the Dictionary.

Example:<br>`Set msg = mMsg.MsgInstasnce("This title")`<br> returns an already existing  _fMsg_ object, when none exists a new created one.

## Miscellaneous aspects
### Min/Max Message Width/Height
A value less than 100 is interpreted as percentage of the screen size, a value equal or greater 100 is interpreted as pt - and re-calculated as percentage to check for the specifiable range. The specifiable width ranges from 25 to 98. I.e. a with less than 25 is set to 25, a width greater than 98 is set to 98.

### Proportional versus Mono-spaced
- ***Monospaced***: Because the text is ++not++  "wrapped" the width of the _Message Form_ is determined by the longest text line (up to the _Maximum Form Width_ specified). When the maximum width is exceeded a vertical scroll bar is applied.<br>Note: The title and the broadest _Button Row_ May still determine an even broader final _Message Form_.
- ***Proportional spaced (default)***: Because the text is "wrapped" the width of a proportional-spaced text is determined by the current form width.<br>Note: When a message is displayed exclusively proportional-spaced the _Message Form_ width is determined by the length of the title, the required space for the broadest _Buttons Row_ and the specified _Minimum Form Width_.

### Unambiguous procedure name
The _ErrSrc_ function provides the procedure-name prefixed by the module-name.
```VB
Private Function ErrSrc(ByVal proc_name As String) As String
    ErrSrc = "<the name of the module goes here>." & proc_name
End Function
```

### Multiple _Monitor_ instances
Because the _Monitor_ service displays the progress message mode-less there may be any number of instances displayed at the same time. See demo<br>
![](images/DemoMsgMonitorInstances.gif)



[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/edit/master/source/fMsg.frm
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/edit/master/source/fMsg.frx
[3]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Message-Service/edit/master/source/mMsg.bas
[4]:https://docs.microsoft.com/de-DE/office/vba/Language/Reference/User-Interface-Help/msgbox-function
