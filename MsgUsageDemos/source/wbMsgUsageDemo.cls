VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbMsgUsageDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' -------------------------------------------------------------------------------
' Workbook MsgUsageDemo.xlsm
'
' All code here will have no effect when this Workbook is opened - unless
' is is opened within the environment it had been implemented. I this
' environment the Common VBA Excel Component Management Services
' (https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services)
' are used to automatically update the modules demonstzrated here when they
' had been update.
'
' Conclusion: Can be ignored!
'
' W. Rauschenberger, Berlin April 2022
' -------------------------------------------------------------------------------
Private Const HOSTED_RAWS = ""

Private Sub Workbook_Open()
   mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub


