Attribute VB_Name = "mMsg"
Option Explicit
#Const AlternateMsgBox = 1 ' = 1 for ErrMsg use the fMsg UserForm instead of the MsgBox
' --------------------------------------------------------------------------------------------
' Standard Module mMsg  Alternative MsgBox
'          Procedures, methods, functions, etc. for displaying a message with a user response.
'
' Methods:
' - AppErr              Converts a positive number into a negative error number
'                       ensuring it not conflicts with a VB error. A negative error
'                       number is turned back into the original positive Application
'                       Error Number.
' - Msg                 Displays a message with any possible 4 replies and the
'                       message either with a foxed or proportional font.
' - Msg3                Displays a message with any possible 4 replies and 3
'                       message sections each either with a foxed or proportional
'                       font.
' - ErrMsg              Displays a common error message either by means of the
'                       VB MsgBox or by means of the common method Msg.
'
' lScreenWidth. Rauschenberger Berlin June 2020
' -------------------------------------------------------------------------------

' Declarations for the function MakeFormResizable (yet unused)
'Private Declare PtrSafe Function GetForegroundWindow Lib "User32.dll" () As Long
'
'Private Declare PtrSafe Function GetWindowLongPtr Lib "User32.dll" Alias "GetWindowLongA" () As LongPtr
'    (ByVal hwnd As LongPtr, ByVal nIndex As Long)
'
'Private Declare PtrSafe Function SetWindowLongPtr Lib "User32.dll" Alias "SetWindowLongA" () As LongPtr
'   (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr)
'
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000
Public Const WS_THICKFRAME = &H40000

Private vMsgReply As Variant

Public Enum StartupPosition         ' ---------------------------
    Manual = 0                      ' Used to position the
    CenterOwner = 1                 ' final setup message form
    CenterScreen = 2                ' horizontally and vertically
    WindowsDefault = 3              ' centered on the screen
End Enum                            ' ---------------------------

Public Type tSection                ' ------------------
       sLabel As String             ' Structure of the
       sText As String              ' UserForm's
       bMonspaced As Boolean        ' message area which
End Type                            ' consists of
Public Type tMessage                ' three message
       section(1 To 3) As tSection  ' sections
End Type

#If AlternateMsgBox Then
' Elaborated error message using fMsg which supports the display of
' up to 3 message sections, optionally monospaced (here used for the
' error path) and each optionally with a label (here used to specify
' the message sections).
' Note: The error title is automatically assembled.
' -------------------------------------------------------------------
Public Sub ErrMsg(Optional ByVal errnumber As Long = 0, _
                  Optional ByVal errsource As String = vbNullString, _
                  Optional ByVal errdescription As String = vbNullString, _
                  Optional ByVal errline As String = vbNullString, _
                  Optional ByVal errtitle As String = vbNullString, _
                  Optional ByVal errpath As String = vbNullString, _
                  Optional ByVal errinfo As String = vbNullString)

    Const PROC      As String = "ErrMsg"
    Dim sIndicate   As String
    Dim sErrText    As String   ' May be a first part of the errdescription

    If errnumber = 0 _
    Then MsgBox "Apparently there is no exit statement line above the error handling! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"

    '~~ Error line info in case one had been provided - additionally integrated in the assembled error title
    If errline = vbNullString Or errline = "0" Then
        sIndicate = vbNullString
    Else
        sIndicate = " (at line " & errline & ")"
    End If

    If errtitle = vbNullString Then
        '~~ When no title is provided, one is assembled by the provided info
        errtitle = errtitle & sIndicate
        '~~ Distinguish between VBA and Application error
        Select Case errnumber
            Case Is > 0:    errtitle = "VBA Error " & errnumber
            Case Is < 0:    errtitle = "Application Error " & AppErr(errnumber)
        End Select
        errtitle = errtitle & " in:  " & errsource & sIndicate
    End If

    If errinfo = vbNullString Then
        '~~ When no error information is provided one may be within the error description
        '~~ which is only possible with an application error raised by Err.Raise
        If InStr(errdescription, "||") <> 0 Then
            sErrText = Split(errdescription, "||")(0)
            errinfo = Split(errdescription, "||")(1)
        Else
            sErrText = errdescription
            errinfo = vbNullString
        End If
    Else
        sErrText = errdescription
    End If

    '~~ Display error message by UserForm fErrMsg
    With fMsg
        .MsgTitle = errtitle
        .MsgLabel(1) = "Error Message/Description:"
        .MsgText(1) = sErrText
        If errpath <> vbNullString Then
            .MsgLabel(2) = "Error path (call stack):"
            .MsgText(2) = errpath
            .MsgMonoSpaced(2) = True
        End If
        If errinfo <> vbNullString Then
            .MsgLabel(3) = "Info:"
            .MsgText(3) = errinfo
        End If
        .MsgButtons = vbOKOnly
        
        '~~ Setup prior activating/displaying the message form is essential!
        '~~ To aviod flickering, the whole setup process must be done before the form is displayed.
        '~~ This  m u s t  be the method called after passing the arguments and before .show
        .Setup
        .show
    End With

End Sub

#Else

' Common error message using MsgBox.
' ---------------------------------------------
Public Sub ErrMsg(ByVal errnumber As Long, _
                  ByVal errsource As String, _
                  ByVal errdescription As String, _
                  ByVal errline As String, _
         Optional ByVal errpath As String = vbNullString)
    
    Const PROC          As String = "ErrMsg"
    Dim sMsg            As String
    Dim sMsgTitle       As String
    Dim sDescription    As String
    Dim sInfo           As String

    If errnumber = 0 _
    Then MsgBox "Exit statement before error handling missing! Error number is 0!", vbCritical, "Application error in " & ErrSrc(PROC) & "!"

    '~~ Prepare Title
    If errnumber < 0 Then
        sMsgTitle = "Application Error " & AppErr(errnumber)
    Else
        sMsgTitle = "VB Error " & errnumber
    End If
    sMsgTitle = sMsgTitle & " in " & errsource
    If errline <> 0 Then sMsgTitle = sMsgTitle & " (at line " & errline & ")"

    '~~ Prepare message
    If InStr(errdescription, "||") <> 0 Then
        '~~ Split error description/message and info
        sDescription = Split(errdescription, "||")(0)
        sInfo = Split(errdescription, "||")(1)
    Else
        sDescription = errdescription
    End If
    sMsg = "Description: " & vbLf & sDescription & vbLf & vbLf & _
           "Source:" & vbLf & errsource
    If errline <> 0 Then sMsg = sMsg & " (at line " & errline & ")"
    If errpath <> vbNullString Then
        sMsg = sMsg & vbLf & vbLf & _
               "Path:" & vbLf & errpath
    End If
    If sInfo <> vbNullString Then
        sMsg = sMsg & vbLf & vbLf & _
               "Info:" & vbLf & sInfo
    End If
    MsgBox sMsg, vbCritical, sMsgTitle

End Sub
#End If

Public Function Max(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = 0, _
           Optional ByVal v4 As Variant = 0, _
           Optional ByVal v5 As Variant = 0, _
           Optional ByVal v6 As Variant = 0, _
           Optional ByVal v7 As Variant = 0, _
           Optional ByVal v8 As Variant = 0, _
           Optional ByVal v9 As Variant = 0) As Variant
' -----------------------------------------------------
' Returns the maximum (biggest) of all provided values.
' -----------------------------------------------------
Dim dMax As Double
    dMax = v1
    If v2 > dMax Then dMax = v2
    If v3 > dMax Then dMax = v3
    If v4 > dMax Then dMax = v4
    If v5 > dMax Then dMax = v5
    If v6 > dMax Then dMax = v6
    If v7 > dMax Then dMax = v7
    If v8 > dMax Then dMax = v8
    If v9 > dMax Then dMax = v9
    Max = dMax
End Function

Public Function AppErr(ByVal lNo As Long) As Long
' ---------------------------------------------------------------------------
' Converts a positive (programmed "application") error number into a negative
' number by adding vbObjectError. Converts a negative number back into a
' positive i.e. the original programmed application error number.
' Usage example:
'    Err.Raise AppErr(1), .... ' when an application error is detected
'    If Err.Number < 0 Then    ' when the error is displayed
'       MsgBox "Application error " & AppErr(Err.Number)
'    Else
'       MsgBox "VB error " & Err.Number
'    End If
' ---------------------------------------------------------------------------
    AppErr = IIf(lNo < 0, AppErr = lNo - vbObjectError, AppErr = vbObjectError + lNo)
End Function
' MsgBox alternative providing up to 5 reply buttons, specified either
' by MsgBox vbOkOnly (the default), vbYesNo, etc. or a comma delimited
' string specifying the used button's caption. The function uses the
' UserForm fMsg and returns the clicked reply button's caption or its
' corresponding vb variable (vbOk, vbYes, vbNo, etc.).
' Note: This is a simplified version of the Msg function.
' --------------------------------------------------------------------
Public Function Box( _
           Optional ByVal title As String = vbNullString, _
           Optional ByVal MsgSectionText As String = vbNullString, _
           Optional ByVal MsgMonoSpaced As Boolean = False, _
           Optional ByVal MinFormWidth As Single = 0, _
           Optional ByVal buttons As Variant = vbOKOnly) As Variant
    
'    Dim siHeight    As Single

    With fMsg
        .MsgTitle = title
        .MsgText(1) = MsgSectionText
        .MsgMonoSpaced(1) = MsgMonoSpaced
        .MsgButtons = buttons
        
        '~~ Setup prior activating/displaying the message form is essential!
        '~~ To aviod flickering, the whole setup process must be done before the form is displayed.
        '~~ This  m u s t  be the method called after passing the arguments and before .show
        .Setup
        .show
        
        Box = .ReplyValue
    End With
    Unload fMsg

    
End Function


Public Function Min(ByVal v1 As Variant, _
                    ByVal v2 As Variant, _
           Optional ByVal v3 As Variant = Nothing, _
           Optional ByVal v4 As Variant = Nothing, _
           Optional ByVal v5 As Variant = Nothing, _
           Optional ByVal v6 As Variant = Nothing, _
           Optional ByVal v7 As Variant = Nothing, _
           Optional ByVal v8 As Variant = Nothing, _
           Optional ByVal v9 As Variant = Nothing) As Variant
' ------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ------------------------------------------------------
Dim dMin As Double
    dMin = v1
    If v2 < dMin Then dMin = v2
    If TypeName(v3) <> "Nothing" Then If v3 < dMin Then dMin = v3
    If TypeName(v4) <> "Nothing" Then If v4 < dMin Then dMin = v4
    If TypeName(v5) <> "Nothing" Then If v5 < dMin Then dMin = v5
    If TypeName(v6) <> "Nothing" Then If v6 < dMin Then dMin = v6
    If TypeName(v7) <> "Nothing" Then If v7 < dMin Then dMin = v7
    If TypeName(v8) <> "Nothing" Then If v8 < dMin Then dMin = v8
    If TypeName(v9) <> "Nothing" Then If v9 < dMin Then dMin = v9
    Min = dMin
End Function

Public Function Msg(ByVal title As String, _
                    ByRef message As tMessage, _
           Optional ByVal buttons As Variant = vbOKOnly, _
           Optional ByVal returnindex As Boolean = False) As Variant
' ------------------------------------------------------------------
' General purpose MsgBox alternative message. By default returns
' the clicked reply buttons value
' ------------------------------------------------------------------
    
    With fMsg
        .MsgTitle = title
        .Msg = message
        .MsgButtons = buttons
         
        '+--------------------------------------------------------------------------+
        '|| Setup prior showing the form is a true performance improvement as it    ||
        '|| avoids a flickering message window when the setup is performed when    ||
        '|| the message window is already displayed, i.e. with the Activate event. ||
        '|| For testing however it may be appropriate to comment the Setup here in ||
        '|| order to have it performed along with the UserForm_Activate event.     ||
        .Setup '                                                                   ||
        '+--------------------------------------------------------------------------+
        
        .show
        On Error Resume Next    ' Just in case the user has terminated the dialog without clicking a reply button
        '~~ Fetching the clicked reply buttons value (or index) unloads the form.
        '~~ In case there were only one button to be clicked, the form will have been unloaded already -
        '~~ and a return value/ibdex will not be available
        If returnindex Then Msg = .ReplyIndex Else Msg = .ReplyValue
    End With
    Unload fMsg

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg" & "." & sProc
End Function

