VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fRunViaButton 
   Caption         =   "Application.Run via Button"
   ClientHeight    =   3015
   ClientLeft      =   2115
   ClientTop       =   2460
   ClientWidth     =   4560
   OleObjectBlob   =   "fRunViaButton.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "fRunViaButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' --------------------------------------------------------------------------
' UserForm fRunViaButton: Prototype as Proof of concept. Invoke a service
'                         provided by a running (open) Workbook.
'
' W. Rauschenberger, Berlin May 2022
' --------------------------------------------------------------------------
Private dctApplicationRunArgs As Dictionary

Private Function ApplicationRunArgsGet(ByVal gra_button As String) As Collection
    Set ApplicationRunArgsGet = dctApplicationRunArgs(gra_button)
End Function

Private Sub cmbRunViaButton_Click()
    ApplicationRunViaButton Me.cmbRunViaButton.Caption ' does noting when no args had been provided
End Sub

Public Sub ApplicationRunArgsLet(ParamArray args() As Variant)
' --------------------------------------------------------------------------
' Store Application.Run arguments in a Dictionary with the first argument
' (args) as the key and all following arguments as item as Collection.
' Note1: The first argument must be the ment buttons caption property.
' Note2: There is (currently) no solution for optional arguments not
'        provided other than provide them all (i.e. use defaults is case)
' --------------------------------------------------------------------------
    Dim v       As Variant
    Dim cll     As Collection
    Dim vKey    As Variant
    
    If dctApplicationRunArgs Is Nothing Then Set dctApplicationRunArgs = New Dictionary
    For Each v In args
        If cll Is Nothing Then
            '~~ This is the first argument which becomes the key
            If IsObject(v) Then Set vKey = v Else vKey = v
            Set cll = New Collection
        Else
            If TypeName(v) = "Error" Then
                cll.Add Nothing ' optional argument not provided
            Else
                cll.Add v
            End If
        End If
    Next v
    dctApplicationRunArgs.Add vKey, cll
    Set cll = Nothing
    
End Sub

Private Sub ApplicationRunViaButton(ByVal button_caption As String)
' --------------------------------------------------------------------------
' Performs an Application.Run for a button's caption provided run arguments
' hd been provided via ApplicationRunArgsLet.
' --------------------------------------------------------------------------
    Dim cll As Collection
    
    If dctApplicationRunArgs.Exists(button_caption) Then
        Set cll = ApplicationRunArgsGet(button_caption)
        Select Case cll.Count
            Case 1: Application.Run cll(1)
            Case 2: Application.Run cll(1), cll(2)
            Case 3: Application.Run cll(1), cll(2), cll(3)
            Case 4: Application.Run cll(1), cll(2), cll(3), cll(4)
            Case 5: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5)
            Case 6: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6)
            Case 7: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7)
            Case 8: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8)
            Case 9: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9)
            Case 10: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10)
            Case 11: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11)
            Case 12: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12)
            Case 13: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13)
            Case 14: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14)
            Case 15: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14), cll(15)
            Case 16: Application.Run cll(1), cll(2), cll(3), cll(4), cll(5), cll(6), cll(7), cll(8), cll(9), cll(10), cll(11), cll(12), cll(13), cll(14), cll(15), cll(16)
            End Select
    End If
    
End Sub
