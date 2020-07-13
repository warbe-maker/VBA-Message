Attribute VB_Name = "mMsg"
Option Explicit
#Const AlternateMsgBox = 1 ' = 1 for ErrMsg use the fMsg UserForm instead of the MsgBox
' --------------------------------------------------------------------------------------------
' Standard Module mMsg
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
Private vMsgReply As Variant

Public Enum StartupPosition
    Manual = 0
    CenterOwner = 1
    CenterScreen = 2
    WindowsDefault = 3
End Enum

Public Property Let MsgReply(ByVal v As Variant):   vMsgReply = v:          End Property
Public Property Get MsgReply() As Variant:          MsgReply = vMsgReply:   End Property

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
        '~~ When no msgtitle is provided, one is assembled by the provided info
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
        .Title = errtitle
        .MsgSection1Label = "Error Message/Description:"
        .MsgSection1Text = sErrText
        .MsgSection1Monospaced = False
        If errpath <> vbNullString Then
            .MsgSection2Label = "Error path (call stack):"
            .MsgSection2Text = errpath
            .MsgSection2Monospaced = True
        End If
        If errinfo <> vbNullString Then
            .MsgSection3Label = "Info:"
            .MsgSection3Text = errinfo
        End If
        .Replies = vbOKOnly
        .FormFinalPositionOnScreen
        .Show
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

' MsgBox alternative providing up to 5 reply buttons, specified either
' by MsgBox vbOkOnly (the default), vbYesNo, etc. or a comma delimited
' string specifying the used button's caption. The function uses the
' UserForm fMsg and returns the clicked reply button's caption or its
' corresponding vb variable (vbOk, vbYes, vbNo, etc.).
' Note: This is a simplified version of the Msg function.
' --------------------------------------------------------------------
Public Function Msg1( _
           Optional ByVal msgtitle As String = vbNullString, _
           Optional ByVal MsgSectionText As String = vbNullString, _
           Optional ByVal msgmonospaced As Boolean = False, _
           Optional ByVal msgminformwidth As Single = 0, _
           Optional ByVal msgreplies As Variant = vbOKOnly) As Variant
    
'    Dim siHeight    As Single

    With fMsg
        .Title = msgtitle
        .MsgSection1Text = MsgSectionText
        .MsgSection1Monospaced = msgmonospaced
        .Replies = msgreplies
        .Show
    End With

    Msg1 = vMsgReply
    
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
Public Function Msg(ByVal msgtitle As String, _
           Optional ByVal msg1label As String = vbNullString, _
           Optional ByVal msg1text As String = vbNullString, _
           Optional ByVal msg1monospaced As Boolean = False, _
           Optional ByVal msg2label As String = vbNullString, _
           Optional ByVal msg2text As String = vbNullString, _
           Optional ByVal msg2monospaced As Boolean = False, _
           Optional ByVal msg3label As String = vbNullString, _
           Optional ByVal msg3text As String = vbNullString, _
           Optional ByVal msg3monospaced As Boolean = False, _
           Optional ByVal msgtitlefontsize As Long = 0, _
           Optional ByVal msgminformwidth As Single = 0, _
           Optional ByVal msgreplies As Variant = vbOKOnly) As Variant
    
'    Dim siHeight        As Single

    With fMsg
        .Title = msgtitle
        
        .MsgSection1Label = msg1label
        .MsgSection1Text = msg1text
        .MsgSection1Monospaced = msg1monospaced
        
        .MsgSection2Label = msg2label
        .MsgSection2Text = msg2text
        .MsgSection2Monospaced = msg2monospaced
        
        .MsgSection3Label = msg3label
        .MsgSection3Text = msg3text
        .MsgSection3Monospaced = msg3monospaced

        .Replies = msgreplies
        .FormFinalPositionOnScreen
        .Show
    End With

    Msg = vMsgReply

End Function

'' This part is from Leith Ross                                              |
'' Found this Code on:                                                       |
'' https://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html |
''                                                                           |
'' All credits belong to him                                                 |
'' ---------------------------------------------------------------------------
'Private Sub MakeFormResizable()
'Const WS_THICKFRAME = &H40000
'Const GWL_STYLE As Long = (-16)
'Dim lStyle As LongPtr
'Dim hwnd As LongPtr
'Dim RetVal
'
'    hwnd = GetForegroundWindow
'
'    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE Or WS_THICKFRAME)
'    RetVal = SetWindowLongPtr(hwnd, GWL_STYLE, lStyle)
'End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg" & "." & sProc
End Function

Public Function AppErr(ByVal lNo As Long) As Long
' -------------------------------------------------------------------------------
' Attention: This function is dedicated for being used with Err.Raise AppErr()
'            in conjunction with the common error handling module mErrHndlr when
'            the call stack is supported. The error number passed on to the entry
'            procedure is interpreted when the error message is displayed.
' The function ensures that a programmed (application) error numbers never
' conflicts with VB error numbers by adding vbObjectError which turns it into a
' negative value. In return, translates a negative error number back into an
' Application error number. The latter is the reason why this function must never
' be used with a true VB error number.
' -------------------------------------------------------------------------------
    If lNo < 0 Then
        AppErr = lNo - vbObjectError
    Else
        AppErr = vbObjectError + lNo
    End If
End Function

