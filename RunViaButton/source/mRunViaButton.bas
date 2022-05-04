Attribute VB_Name = "mRunViaButton"
Option Explicit
' --------------------------------------------------------------------------
' Standard Module mRunViaButton: Prototype as Proof of concept. Invokes a
'                                service provided by a running (open)
'                                Workbook (this one).
'
' Requires: Reference to the "Microsoft Scripting Runtime".
'
' W. Rauschenberger, Berlin May 2022
' --------------------------------------------------------------------------
' Timer means used for the creation of UserForm instances to avoid conflicts
Private Declare PtrSafe Function getFrequency Lib "kernel32" _
                Alias "QueryPerformanceFrequency" (TimerSystemFrequency As Currency) As Long
Private Declare PtrSafe Function getTickCount Lib "kernel32" _
                Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long

Private Function MsgInstance(ByVal fi_key As String, _
                    Optional ByVal fi_unload As Boolean = False) As fRunViaButton
' -------------------------------------------------------------------------
' Returns an instance of the UserForm fRunViaButton which is uniquely
' identified (fi_key) which usually is the title of the displayed message
' bu may be anything else as well (including an object). An already
' existing or new created instance is maintained in a static Dictionary
' with (fi_key) as the key and returned to the caller. When (fi_unload) is
' TRUE a possibly already existing Userform instance is unloaded.
' -------------------------------------------------------------------------
    Const PROC = "MsgInstance"
    
    Static cyStart      As Currency
    Static Instances    As Dictionary    ' Collection of (possibly still)  active form instances
    Dim MsecsElapsed    As Currency
    Dim MsecsWait       As Long
    
    If Instances Is Nothing Then Set Instances = New Dictionary
    
    If fi_unload Then
        If Instances.Exists(fi_key) Then
            On Error Resume Next
            Unload Instances(fi_key) ' The instance may be already unloaded
            Instances.Remove fi_key
        End If
        Exit Function
    End If
    
    If Not Instances.Exists(fi_key) Then
        '~~ When there is no evidence of an already existing instance a new one is established.
        '~~ In order not to interfere with any prior established instance a minimum wait time
        '~~ of 10 milliseconds is maintained.
        MsecsElapsed = (TicksCount() - cyStart) / CDec(TicksFrequency)
        MsecsWait = 10 - MsecsElapsed
        If MsecsWait > 0 Then
            Sleep MsecsWait
        End If
        cyStart = TicksCount()
        Set MsgInstance = New fRunViaButton
        Instances.Add fi_key, MsgInstance
    Else
        '~~ An instance identified by fi_key exists in the Dictionary.
        '~~ It may however have already been unloaded.
        On Error Resume Next
        Set MsgInstance = Instances(fi_key)
        Select Case Err.Number
            Case 0
            Case 13
                If Instances.Exists(fi_key) Then
                    Instances.Remove fi_key ' the no longer existing instance is removed from the Dictionary
                End If
                Set MsgInstance = New fRunViaButton
                Instances.Add fi_key, MsgInstance
            Case Else
                VBA.MsgBox "Unexpected error in mRunViaButton.MsgInstance", , "Unexpectd error"
        End Select
        On Error GoTo -1
    End If

End Function

Private Sub MsgInstanceTest()
    Const TEST_TITLE = "Test-Title"
    Dim i As Long
    
    For i = 1 To 5
        MsgInstance TEST_TITLE & i ' sleep for 10 milliseconds between the creation of two instances
    Next i
    
    For i = 1 To 5
        MsgInstance TEST_TITLE & i, True ' remove/unload all instances
    Next i

End Sub

Private Sub RunViaButtonTest()
' -------------------------------------------------------------------------
' When the button in the UserForm is clicked the mMsg.Box service is called
' to display a message.
' Note: With Application.Run all arguments are "by position"! I.e. it is
'       not possible to skip optional arguments by means of naming them.
' -------------------------------------------------------------------------
    Dim fRun        As fRunViaButton
    Dim RunTitle    As String: RunTitle = "Application.Run via Button"
    
    Set fRun = MsgInstance(RunTitle)
    With fRun
        '~~ Arguments for the button with the caption "Run via Button"
        .ApplicationRunArgsLet "Run via Button" _
                             , "RunViaButton.xlsm!mMsg.Box" _
                             , "This has been invoked via a message button" _
                             , vbOKOnly _
                             , "Run via button Test" _
                             , False _
                             , 1 _
                             , False _
                             , False _
                             , 300 _
                             , 85 _
                             , 20 _
                             , 85 _
                             , "200;150"
        .Show False
    End With
    
End Sub

Private Function TicksCount() As Currency:      getTickCount TicksCount:        End Function

Private Function TicksFrequency() As Currency:  getFrequency TicksFrequency:    End Function

