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

#If Resizable Then
Public Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Public Type MessageSection
    sLabel As String
    sText As String
    bMonospaced As Boolean
End Type

Private vMsgReply As Variant

Public Enum StartupPosition
    Manual = 0
    CenterOwner = 1
    CenterScreen = 2
    WindowsDefault = 3
End Enum

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
        .ApplTitle = errtitle
        .ApplLabel(1) = "Error Message/Description:"
        .ApplText(1) = sErrText
        If errpath <> vbNullString Then
            .ApplLabel(2) = "Error path (call stack):"
            .ApplText(2) = errpath
            .ApplMonoSpaced(2) = True
        End If
        If errinfo <> vbNullString Then
            .ApplLabel(3) = "Info:"
            .ApplText(3) = errinfo
        End If
        .ApplButtons = vbOKOnly
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

' Used with Err.Raise AppErr() to convert a positive application error number
' into a negative number to avoid any conflict with a VB error. Used when the
' error is displayed with ErrMsg to turn the negative number back into the
' original positive application number.
' The function ensures that a programmed (application) error numbers never
' conflicts with VB error numbers by adding vbObjectError which turns it
' into a negative value. In return, translates a negative error number
' back into an Application error number. The latter is the reason why this
' function must never be used with a true VB error number.
' ------------------------------------------------------------------------
Public Function AppErr(ByVal lNo As Long) As Long
    If lNo < 0 Then
        AppErr = lNo - vbObjectError
    Else
        AppErr = vbObjectError + lNo
    End If
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
           Optional ByVal msgmonospaced As Boolean = False, _
           Optional ByVal minformwidth As Single = 0, _
           Optional ByVal buttons As Variant = vbOKOnly) As Variant
    
'    Dim siHeight    As Single

    With fMsg
        .ApplTitle = title
        .ApplText(1) = MsgSectionText
        .ApplMonoSpaced(1) = msgmonospaced
        .ApplButtons = buttons
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

' MsgBox alternative providing three message sections, each optionally
' monospaced and with an optional label/haeder. The function uses the
' UserForm fMsg and returns the clicked reply button's caption or its
' corresponding vb variable (vbOk, vbYes, vbNo, etc.).
' ------------------------------------------------------------------
Public Function Msg(ByVal title As String, _
           Optional ByVal label1 As String = vbNullString, _
           Optional ByVal text1 As String = vbNullString, _
           Optional ByVal monospaced1 As Boolean = False, _
           Optional ByVal label2 As String = vbNullString, _
           Optional ByVal text2 As String = vbNullString, _
           Optional ByVal monospaced2 As Boolean = False, _
           Optional ByVal label3 As String = vbNullString, _
           Optional ByVal text3 As String = vbNullString, _
           Optional ByVal monospaced3 As Boolean = False, _
           Optional ByVal monospacedfontsize As Long = 0, _
           Optional ByVal buttons As Variant = vbOKOnly) As Variant
    
    With fMsg
        .ApplTitle = title
        
        .ApplLabel(1) = label1
        .ApplText(1) = text1
        .ApplMonoSpaced(1) = monospaced1
        
        .ApplLabel(2) = label2
        .ApplText(2) = text2
        .ApplMonoSpaced(2) = monospaced2
        
        .ApplLabel(3) = label3
        .ApplText(3) = text3
        .ApplMonoSpaced(3) = monospaced3

        .ApplButtons = buttons
        .show
        On Error Resume Next ' Just in case the user has terminated the dialog without clicking a reply button
        Msg = .ReplyValue
    End With
    Unload fMsg

End Function
#If Resizable Then
Public Sub ResizeWindowSettings(frm As Object, show As Boolean)

    Dim windowStyle As Long
    Dim windowHandle As Long

    'Get the references to window and style position within the Windows memory
    windowHandle = FindWindowA(vbNullString, frm.caption)
    windowStyle = GetWindowLong(windowHandle, GWL_STYLE)
    
    'Determine the style to apply based
    If show = False Then
        windowStyle = windowStyle And (Not WS_THICKFRAME)
    Else
        windowStyle = windowStyle + (WS_THICKFRAME)
    End If
    
    'Apply the new style
    SetWindowLong windowHandle, GWL_STYLE, windowStyle
    
    'Recreate the UserForm window with the new style
    DrawMenuBar windowHandle

End Sub
#End If
Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg" & "." & sProc
End Function

