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
' W. Rauschenberger Berlin June 2020
' -------------------------------------------------------------------------------
Public Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
' Functions to get the displays DPI
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Const LOGPIXELSX = 88               ' Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72  ' A point is defined as 1/72 inches
Private Declare PtrSafe Function GetForegroundWindow _
  Lib "User32.dll" () As Long

Private Declare PtrSafe Function GetWindowLongPtr _
  Lib "User32.dll" Alias "GetWindowLongA" _
() '    (ByVal hwnd As LongPtr, _
     ByVal nIndex As Long) _
  As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr _
  Lib "User32.dll" Alias "SetWindowLongA" _
() '    (ByVal hwnd As LongPtr, _
     ByVal nIndex As LongPtr, _
     ByVal dwNewLong As LongPtr) _
  As LongPtr

Private vMsgReply As Variant

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
        .msg1Label = "Error Message/Description:"
        .Msg1TextPrprtional = sErrText
        If errpath <> vbNullString Then
            .msg2label = "Error path (call stack):"
            .Msg2TextPrprtional = errpath
        End If
        If errinfo <> vbNullString Then
            .msg3label = "Info:"
            .Msg3TextPrprtional = errinfo
        End If
        .Replies = vbOKOnly
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
           Optional ByVal msgtext As String = vbNullString, _
           Optional ByVal msgmonospaced As Boolean = False, _
           Optional ByVal msgminformwidth As Single = 0, _
           Optional ByVal msgreplies As Variant = vbOKOnly) As Variant
    Dim w           As Long
    Dim h           As Long
    Dim siHeight    As Single

    w = GetSystemMetrics32(0) ' Screen Resolution width in points
    h = GetSystemMetrics32(1) ' Screen Resolution height in points

    With fMsg
        .Title = msgtitle

        If msgtext <> vbNullString Then
            If msgmonospaced = True _
            Then .Msg1TextMonospaced = msgtext _
            Else .Msg1TextPrprtional = msgtext
        End If

        .Replies = msgreplies
        If msgminformwidth <> 0 Then .Width = Max(.Width, msgminformwidth)
        .StartUpPosition = 1
        .Width = w * PointsPerPixel * 0.85 'Userform width= Width in Resolution * DPI * 85%
        siHeight = h * PointsPerPixel * 0.2
        .Height = Min(.Height, siHeight)

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
           Optional ByVal msg1Label As String = vbNullString, _
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
    
    Dim w           As Long
    Dim h           As Long
    Dim siHeight    As Single

    w = GetSystemMetrics32(0) ' Screen Resolution width in points
    h = GetSystemMetrics32(1) ' Screen Resolution height in points

    With fMsg
        .Title = msgtitle

        If msg1text <> vbNullString Then
            If msg1monospaced = True _
            Then .Msg1TextMonospaced = msg1text _
            Else .Msg1TextPrprtional = msg1text
            .msg1Label = msg1Label
        End If

        If msg2text <> vbNullString Then
            If msg2monospaced = True _
            Then .Msg2TextMonospaced = msg2text _
            Else .Msg2TextPrprtional = msg2text
            .msg2label = msg2label
        End If

        If msg3text <> vbNullString Then
            If msg3monospaced = True _
            Then .Msg3TextMonospaced = msg3text _
            Else .Msg3TextPrprtional = msg3text
            .msg3label = msg3label
        End If

        .Replies = msgreplies
        If msgminformwidth <> 0 Then .Width = Max(.Width, msgminformwidth)
        .StartUpPosition = 1
        .Width = w * PointsPerPixel * 0.85 'Userform width= Width in Resolution * DPI * 85%
        siHeight = h * PointsPerPixel * 0.2
        .Height = Min(.Height, siHeight)

        .Show
    End With

    Msg = vMsgReply

End Function

Private Sub MakeFormResizable()
' ---------------------------------------------------------------------------
' This part is from Leith Ross                                              |
' Found this Code on:                                                       |
' https://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html |
'                                                                           |
' All credits belong to him                                                 |
' ---------------------------------------------------------------------------
Const WS_THICKFRAME = &H40000
Const GWL_STYLE As Long = (-16)
Dim lStyle As LongPtr
Dim hwnd As LongPtr
Dim RetVal

    hwnd = GetForegroundWindow

    lStyle = GetWindowLongPtr(hwnd, GWL_STYLE Or WS_THICKFRAME)
    RetVal = SetWindowLongPtr(hwnd, GWL_STYLE, lStyle)
End Sub

Private Function PointsPerPixel() As Double
' ----------------------------------------
' Return DPI
' ----------------------------------------
Dim hDC             As Long
Dim lDotsPerInch    As Long
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC
End Function

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

