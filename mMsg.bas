Attribute VB_Name = "mMsg"
Option Explicit
#Const AlternateMsgBox = 1 ' = 1 for ErrMsg use the fMsg UserForm instead of the MsgBox
' --------------------------------------------------------------------------------------------
' Standard Module mMsg  Alternative MsgBox
'          Procedures, methods, functions, etc. for displaying a message with a user response.
'
' Methods: Dsply        Displays a message with any possible 4 replies and the
'                       message either with a foxed or proportional font.
'
' W. Rauschenberger, Berlin Nov 2020
' -------------------------------------------------------------------------------

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
       section(1 To 4) As tSection  ' sections
End Type

Public Function Max(ParamArray va() As Variant) As Variant
' ------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ------------------------------------------------------
    Dim v   As Variant
    
    On Error Resume Next
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function
'
'Public Function AppErr(ByVal lNo As Long, _
'              Optional ByRef sError As String = vbNullString) As Variant
'' ---------------------------------------------------------------------------
'' Usage example when a programmed application error occurs:
''    If ..... Then Err.Raise AppErr(1), ....
'' Usage example when the error message is displayed, e.g. by means of MsgBox:
''   AppErr Err.Number, sErrTitle
''   MsgBox dsply_title:=sErrTitle
'' ---------------------------------------------------------------------------
'    If lNo < 0 Then
'        '~~ This is an application error number which had turned into a negative number
'        '~~ in order to avoid any conflict with a VB error. The function returns the
'        '~~ original positive application error number and a corresponding title
'        AppErr = lNo - vbObjectError
'        If Not IsMissing(sError) Then sError = "Application error " & AppErr
'    Else
'        '~~ This is a positive error number regarded as a programmed application error
'        '~~ The function returns a negative number in order to avoid any conflict with
'        '~~ a VB error.
'        AppErr = vbObjectError + lNo
'        If Not IsMissing(sError) Then sError = "Microsoft Visual Basic runtime error " & lNo
'    End If
'End Function

Public Function Dsply(ByVal dsply_title As String, _
                      ByRef dsply_message As tMessage, _
             Optional ByVal dsply_buttons As Variant = vbOKOnly, _
             Optional ByVal dsply_returnindex As Boolean = False, _
             Optional ByVal dsply_min_width As Long = 200) As Variant
' ------------------------------------------------------------------------------------
' Common VBA Message Display: A service using the Common VBA (alternative) MsgBox.
' See: https://warbe-maker.github.io/vba/common/2020/10/19/Alternative-VBA-MsgBox.html
'
' W. Rauschenberger, Berlin, Nov 2020
' ------------------------------------------------------------------------------------

    With fMsg
        .MinFormWidth = dsply_min_width
        .MsgTitle = dsply_title
        .Msg = dsply_message
        .MsgButtons = dsply_buttons
        '+------------------------------------------------------------------------+
        '|| Setup prior showing the form improves the performance significantly  ||
        '|| and avoids any flickering message window with its setup.             ||
        '|| For testing purpose it may be appropriate to out-comment the Setup.  ||
        .Setup '                                                                 ||
        '+------------------------------------------------------------------------+
        .show
    End With
    
    ' -----------------------------------------------------------------------------
    ' Obtaining the reply value/index is only possible when more than one button is
    ' displayed! When the user had a choice the form is hidden when the button is
    ' pressed and the UserForm is unloade when the return value/index (either of
    ' the two) is obtained!
    ' -----------------------------------------------------------------------------
    Dsply = IIf(dsply_returnindex, fMsg.ReplyIndex, fMsg.ReplyValue)

End Function

Public Function Min(ParamArray va() As Variant) As Variant
' ---------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' ---------------------------------------------------------
   Dim v As Variant
   
   Min = va(LBound(va))
   On Error Resume Next
   For Each v In va
      If v < Min Then Min = v
   Next v
End Function


Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mMsg" & "." & sProc
End Function

