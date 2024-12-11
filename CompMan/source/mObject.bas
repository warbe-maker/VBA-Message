Attribute VB_Name = "mObject"
Option Explicit
Option Compare Text
' -----------------------------------------------------------------------------------
' Standard  Module mObject Checks the existence of objects.
'
' Methods:
' - ComponentExists     Returns TRUE when the object exists
' - CustomViewExists    Returns TRUE when the object exists
' - FileExists          Returns TRUE when the object exists
' - ProcedureExists     Returns TRUE when the object exists
' - RangeExists         Returns TRUE when the object exists
' - ReferenceExists     Returns TRUE when the object exists
' - WorkbookExists      Returns TRUE when the object exists
' - WorksheetExists     Returns TRUE when the object exists
' - OpenWb              Returns a Workbook object identified by its name (sName)
'                       regardless in which application instance it is opened.
'                       Returns Nothing when a Workbook named is not open.
'                       The name may be a Workbook's full or short name
' - OpenWbs             Returns a Distionary of all open Workbooks in any application
'                       instance with the Workbook's name as the key and the Workbook
'                       object a item.
'
' Uses:     Standard Module mErrHndlr
'
' Requires: Reference to "Microsoft Scripting Runtine"
'           Reference to "Microsoft Visual Basic for Applications Extensibility ..."
' Note:     When the existence checks for Component, Procedure, and Reference are not
'           needed they may be out-commented and the reference to the "Microsoft Visual
'           Basic for Applications Extensibility ..." will then become obsolete.
'
' W. Rauschenberger, Berlin August 2019
' -----------------------------------------------------------------------------------
#Const VBE = 1              ' Requires a Reference to "Microsoft Visual Basis Extensibility ..."
' --- Begin of declarations to get all Workbooks of all running Excel instances
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As LongPtr, ByRef riid As UUID, ByRef ppvObject As Object) As LongPtr

Type UUID 'GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Const OBJID_NATIVEOM As LongPtr = &HFFFFFFF0
' --- End of declarations to get all Workbooks of all running Excel instances
' --- Error declarations
Const ERR_EXISTS_CMP01 = "The Component (parameter vComp) for the Component's existence check is neihter a Component object nor a string (a Component's name)!"
Const ERR_EXISTS_CVW01 = "The CustomView (parameter vCv) for the CustomView's existence check is neither a string (CustomView's name) nor a CustomView object!"
Const ERR_EXISTS_FLE01 = "The File (parameter vFile) for the File's existence check is neither a full path/file name nor a file object!"
Const ERR_EXISTS_OWB01 = "The Workbook (parameter vWb) is not open (it may have been open and already closed)!"
Const ERR_EXISTS_OWB02 = "A Workbook named '<>' is not open in any application instance!"
Const ERR_EXISTS_OWB03 = "The Workbook (parameter vWb) of which the open object is requested is ""Nothing"" (neither a Workbook object nor a Workbook's name or fullname)!"
Const ERR_EXISTS_PRC01 = "The item (parameter v) for the Procedure's existence check is neither a Component object nor a CodeModule object!"
Const ERR_EXISTS_RNG01 = "The Worksheet (parameter vWs) for the Range's existence check does not exist in Workbook (vWb)!"
Const ERR_EXISTS_RNG02 = "The Range (parameter vRange) for the Range's existence check is ""Nothing""!"
Const ERR_EXISTS_REF01 = "The Reference (parameter vRef) for the Reference's existence check is neither a valid GUID (a string enclosed in { } ) nor a Reference object!"
Const ERR_EXISTS_WBK01 = "The Workbook (parameter vWb) is neither a Workbook object nor a Workbook's name or fullname)!"
Const ERR_EXISTS_WSH01 = "The Worksheet (parameter vWs) for the Worksheet's existence check is neither a Worksheet object nor a Worksheet's name or modulename!"
Const ERR_EXISTS_GOW01 = "A Workbook (parameter vWb) named '<>' is not open!"
Const ERR_EXISTS_GOW02 = "A Workbook with the provided name (parameter vWb) is open. However it's location is '<>1' and not '<>2'!"
Const ERR_EXISTS_GOW03 = "A Workbook named '<>' (parameter vWb) is not open. A full name must be provided to get it opened!"
Const ERR_EXISTS_GOW04 = "The Workbook (parameter vWb) is a Workbook object not/no longer open!"
Const ERR_EXISTS_GOW05 = "The Workbook (parameter vWb) is neither a Workbook object nor a string (name or fullname)!"
Const ERR_EXISTS_GOW06 = "A Workbook file named '<>' (parameter vWb) does not exist!"

Public Function ArrayExists(arr As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the array (Arr) exist.
' ------------------------------------------------------------------------------
Dim LB As Long
Dim UB As Long

    Err.Clear
    On Error Resume Next
    If IsArray(arr) = False Then
        ' we weren't passed an array, return True
        ArrayExists = False
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    UB = UBound(arr, 1)
    If (Err.Number <> 0) Then
        ArrayExists = False
    Else
        ''''''''''''''''''''''''''''''''''''''''''
        ' On rare occassion, under circumstances I
        ' cannot reliably replictate, Err.Number
        ' will be 0 for an unallocated, empty array.
        ' On these occassions, LBound is 0 and
        ' UBound is -1.
        ' To accomodate the weird behavior, test to
        ' see if LB > UB. If so, the array is not
        ' allocated.
        ''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        LB = LBound(arr)
        If LB > UB Then
            ArrayExists = False
        Else
            ArrayExists = True
        End If
    End If

End Function

Public Function ProcedureExists(ByVal v As Variant, _
                                ByVal sName As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Procedure named (sName) exists in the CodeModule (vbcm).
' ------------------------------------------------------------------------------
    Const PROC = "ProcedureExists"

    On Error GoTo eh
    Dim vbcm        As CodeModule
    Dim iLine       As Long             ' For the existence check of a VBA procedure in a CodeModule
    Dim sLine       As String           ' For the existence check of a VBA procedure in a CodeModule
    Dim vbProcKind  As vbext_ProcKind   ' For the existence check of a VBA procedure in a CodeModule
    
    ProcedureExists = False

    If Not TypeName(v) = "Nothing" Then
        If TypeOf v Is VBComponent Then
            Set vbcm = v.CodeModule
            With vbcm
                For iLine = 1 To .CountOfLines
                    If .ProcOfLine(iLine, vbProcKind) = sName Then
                        ProcedureExists = True
                        GoTo xt
                    End If
                Next iLine
                GoTo xt
            End With
        ElseIf TypeOf v Is CodeModule Then
            Set vbcm = v
            With vbcm
                For iLine = 1 To .CountOfLines
                    If .ProcOfLine(iLine, vbProcKind) = sName Then
                        ProcedureExists = True
                        GoTo xt
                    End If
                Next iLine
                GoTo xt
            End With
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_PRC01

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function


Private Function ErrMsg(ByVal err_source As String, _
          Optional ByVal err_no As Long = 0, _
          Optional ByVal err_dscrptn As String = vbNullString, _
         Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Common, minimum VBA error handling providing the means to resume the error
' line when the Conditional Compile Argument Debugging=1.
' Usage: When this procedure is copied into any desired module the statement
'        If ErrMsg(ErrSrc(PROC) = vbYes Then: Stop: Resume
'        is appropriate
'        The caller provides the source of the error through ErrSrc(PROC) where
'        ErrSrc is a procedure available in the module using this ErrMsg and
'        PROC is the constant identifying the procedure
' Uses: AppErr to translate a negative programmed application error into its
'              original positive number
' ------------------------------------------------------------------------------
    Dim ErrNo       As Long
    Dim ErrDesc     As String
    Dim ErrType     As String
    Dim ErrLine     As Long
    Dim AtLine      As String
    Dim ArgButtons  As Long
    
    If err_no = 0 Then err_no = Err.Number
    If err_no < 0 Then
        ErrNo = AppErr(err_no)
        ErrType = "Applicatin error "
    Else
        ErrNo = err_no
        ErrType = "Runtime error "
    End If
    
    If err_line = 0 Then ErrLine = Erl
    If err_line <> 0 Then AtLine = " at line " & err_line
    
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error message available ---"
    ErrDesc = "Error: " & vbLf & err_dscrptn & vbLf & vbLf & "Source: " & vbLf & err_source & AtLine

    
#If Debugging = 1 Then
    ArgButtons = vbYesNo
    ErrDesc = ErrDesc & vbLf & vbLf & "Debugging: Yes=Resume error line, No=Continue"
#Else
    ArgButtons = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrType & ErrNo & " in " & err_source _
                  , Prompt:=ErrDesc _
                  , Buttons:=ArgButtons)
End Function

Public Function ErrNo(ByVal l As Long) As Long
    Const ERR_BASE = 600    ' may be any number above 512 to fit into the vb project
    ErrNo = ERR_BASE + l
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mObject" & "." & sProc
End Function

Public Function FileExists(ByVal vFile As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the file (vFile) - which may be a file object or a file's
' full name - exists.
' ------------------------------------------------------------------------------
    Const PROC = "FileExists"
    
    On Error GoTo eh
    Dim sTest   As String
    
    FileExists = False

    If Not TypeName(vFile) = "Nothing" Then
        If TypeOf vFile Is File Then
            With New FileSystemObject
                On Error Resume Next
                sTest = vFile.Name
                FileExists = Err.Number = 0
                GoTo xt
            End With
        ElseIf VarType(vFile) = vbString Then
            With New FileSystemObject
                FileExists = .FileExists(vFile)
                GoTo xt
            End With
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_FLE01

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function WorksheetExists(ByVal vWb As Variant, _
                                ByRef vWs As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Worksheet (vWs) - which may be a Worksheet object or a
' Worksheet's name - exists in the Workbook (vWb).
' ------------------------------------------------------------------------------
    Const PROC = "WorksheetExists"
    
    On Error GoTo eh
    Dim sTest   As String
    Dim wsTest  As Worksheet
    Dim wb      As Workbook
    Dim ws      As Worksheet
    
    WorksheetExists = False
    Set wb = GetOpenWorkbook(vWb) ' raises an error when not open
    
    If Not TypeName(vWs) = "Nothing" Then
        If TypeOf vWs Is Worksheet Then
            Set ws = vWs
            For Each wsTest In wb.Worksheets
                WorksheetExists = wsTest Is ws
                If WorksheetExists Then GoTo xt
            Next wsTest
            GoTo xt
        ElseIf VarType(vWs) = vbString Then
            For Each wsTest In wb.Worksheets
                WorksheetExists = wsTest.Name = vWs
                If WorksheetExists Then
                    Set vWs = wsTest
                    GoTo xt
                End If
            Next wsTest
            GoTo xt
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_WSH01
        
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function WsShapeExists(ByVal wse_ws As Worksheet, _
                              ByVal wse_shape_name) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when a shape named (wse_shape_name) exist in Worksheet wse_ws.
' ------------------------------------------------------------------------------
    Const PROC = "WsShapeExists"
    
    On Error GoTo eh
    Dim shp As Shape
    
    For Each shp In wse_ws.Shapes
        If shp.Name = wse_shape_name Then
            WsShapeExists = True
            GoTo xt
        End If
    Next shp
    
xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Public Function ComponentExists(ByVal vWb As Variant, _
                                ByRef vComp As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE and the Component object (vComp) when the Component named (vComp)
' - which may be a Component object or a Component's name - exists in the
' Workbook (vWb) - which may be a Workbook object or a Workbook's name or
' fullname of an open Workbook.
' ------------------------------------------------------------------------------
    Const PROC = "ComponentExists"

    On Error GoTo eh
    Dim wb      As Workbook
    Dim sTest   As String
    Dim sName   As String
    Dim vbc     As VBComponent
    
    ComponentExists = False
    Set wb = GetOpenWorkbook(vWb)   ' raises an error when not open

    If Not TypeName(vComp) = "Nothing" Then
        If TypeOf vComp Is VBComponent Then
            Set vbc = vComp
            sName = vbc.Name
            On Error Resume Next
            sTest = wb.VBProject.VBComponents(sName).Name
            ComponentExists = Err.Number = 0
            GoTo xt
        ElseIf VarType(vComp) = vbString Then
            sName = vComp
            On Error Resume Next
            sTest = wb.VBProject.VBComponents(sName).Name
            ComponentExists = Err.Number = 0
            GoTo xt
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_CMP01

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function CustomViewExists(ByVal vWb As Variant, _
                                 ByVal vCv As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the CustomView (vCv) - may be a CustomView object or a
' CustoView's name - exists in Workbook (wb). If vCv is provided as CustomView
' object, only its name is used to check the existence in Workbook (wb).
' ------------------------------------------------------------------------------
    Const PROC = "CustomViewExists"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim sTest   As String
    
    CustomViewExists = False
    If TypeName(vWb) = "Workbook" Then Set wb = vWb Else Set wb = GetOpenWorkbook(vWb)   ' raises an error when not open
    If Not TypeName(vCv) = "Nothing" Then
        If TypeOf vCv Is CustomView Then
            On Error Resume Next
            sTest = vCv.Name
            CustomViewExists = Err.Number = 0
            GoTo xt
        End If
    End If
    If VarType(vCv) = vbString Then
        On Error Resume Next
        sTest = wb.CustomViews(vCv).Name
        CustomViewExists = Err.Number = 0
        GoTo xt
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_CVW01
        
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function WorkbookIsOpen(ByRef vWb As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE and the open Workbook object in (vWb) when the Workbook (vWb)
' - which may be a Workbook object, a Workbook's name or fullname - is open in
' whichever Excel Application instance. If a fullname is provided and the file
' does not exist but is open from another location, the Workbook is regarded as
' having been moved to the other location and thus retunred as oben object.
' ------------------------------------------------------------------------------
    Const PROC = "WorkbookIsOpen"

    On Error GoTo eh
    
    If Not TypeName(vWb) = "Nothing" Then
        If TypeOf vWb Is Workbook Then
            On Error Resume Next
            Set vWb = GetOpenWorkbook(vWb)
            WorkbookIsOpen = Err.Number = 0
            GoTo xt
        ElseIf VarType(vWb) = vbString Then
            On Error Resume Next
            Set vWb = GetOpenWorkbook(vWb)
            WorkbookIsOpen = Err.Number = 0
            GoTo xt
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_WBK01
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function ReferenceExists(ByVal vWb As Variant, _
                                ByVal vRef As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Reference (vRef) - which may be a Reference object or a
' Refernece's GUID - exists in the VBProject of the Workbook (vWb) - which may
' be a Workbook object or a Workbook's name or fullname. When vRef is provided
' as object, only its GUID is used for the existence check in Workbook (vWb).
' ------------------------------------------------------------------------------
    Const PROC = "ReferenceExists"

    On Error GoTo eh
    Dim ref     As Reference
    Dim refTest As Reference
    Dim wb      As Workbook
    
    ReferenceExists = False
    Set wb = GetOpenWorkbook(vWb)

    If Not TypeName(vRef) = "Nothing" Then
        If TypeOf vRef Is Reference Then
            Set refTest = vRef
            For Each ref In wb.VBProject.References
                If ref.GUID = refTest.GUID Then
                    ReferenceExists = True
                    GoTo xt
                End If
            Next ref
            GoTo xt
        ElseIf VarType(vRef) = vbString Then
            If Left$(vRef, 1) = "{" And Right$(vRef, 1) = "}" Then ' valid Reference GUID
                For Each ref In wb.VBProject.References
                    If ref.GUID = vRef Then
                        ReferenceExists = True
                        GoTo xt
                    End If
                Next ref
                GoTo xt
            End If
        End If
    End If
    Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_REF01

xt: Exit Function

eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function RangeExists(ByVal vWb As Variant, _
                            ByVal vWs As Variant, _
                            ByVal vRange As Variant) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the Range (vRange) - which may be range object or a range's
' name - exists in the Worksheet (ws) of the Workbook (wb).
' ------------------------------------------------------------------------------
    Const PROC = "RangeExists"

    On Error GoTo eh
    Dim sTest   As String
    Dim ws      As Worksheet
    Dim wb      As Workbook
    Dim rg      As Range
    
    RangeExists = False
    Set wb = GetOpenWorkbook(vWb)   ' raises an error when not open
    
    If Not TypeName(vRange) = "Nothing" Then
        If TypeOf vRange Is Range Then
            Set rg = vRange
            On Error Resume Next
            sTest = rg.Address
            RangeExists = Err.Number = 0
            GoTo xt
        ElseIf VarType(vRange) = vbString Then
            If WorksheetExists(wb, vWs) Then
                Set ws = vWs
                On Error Resume Next
                sTest = ws.Range(vRange).Address
                RangeExists = Err.Number = 0
                GoTo xt
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_RNG01
            End If
        End If
    End If
    Err.Raise AppErr(2), ErrSrc(PROC), ERR_EXISTS_RNG02
            
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function OpenWb(ByVal vWb As Variant) As Workbook
' ------------------------------------------------------------------------------
' Returns the Workbook object named (vWb) - which may be a Workbook opbject or a
' Workbook's name or fullname - regardless in which application instance it is
' opened. Raises an error when the Workbook is not open. When (vWb) is a
' Workbook's fullname and a Workbook with the name is open the object is only
' returned when also the fullname is identical. A specific error is raised when
' the name is equal but the path is different.
' ------------------------------------------------------------------------------
    Const PROC = "OpenWb"
    
    On Error GoTo eh
    Dim i       As Long
    Dim dct     As Dictionary
    Dim sTest   As String
    Dim wb      As Workbook
    Dim sName   As String
    
    Set OpenWb = Nothing
    
    If Not TypeName(vWb) = "Nothing" Then
        
        If TypeOf vWb Is Workbook Then
            '~~ The provided paramenter is a Workbook which may be open or not
            Set wb = vWb
            On Error Resume Next
            sName = wb.Name
            If Err.Number <> 0 Then
                On Error GoTo eh
                Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_OWB01
            End If
            With OpenWbs
                If .Exists(sName) Then
                    Set OpenWb = .Item(sName)
                    GoTo xt
                End If
            End With ' OpenWbs
        ElseIf VarType(vWb) = vbString Then
            sName = Split(vWb, "\")(UBound(Split(vWb, "\")))    ' Unstrip the Workbook name
            With OpenWbs
                If .Exists(sName) Then
                    Set OpenWb = .Item(sName)
                    GoTo xt
                Else
                    Err.Raise AppErr(2), ErrSrc(PROC), Replace$(ERR_EXISTS_OWB02, "<>", CStr(vWb))
                End If
            End With ' OpenWbs
        End If
    End If
    Err.Raise AppErr(3), ErrSrc(PROC), ERR_EXISTS_OWB03
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function OpenWbs() As Dictionary
' ------------------------------------------------------------------------------
' Returns a Dictionary of all Workbooks open in any running excel instance with
' the Workbook's name as the key and the Workbook object as item.
' ------------------------------------------------------------------------------
    Const PROC = "OpenWbs"

    On Error GoTo eh
#If Win64 Then
    Dim hWndMain As LongPtr
#Else
    Dim hWndMain As Long
#End If
    Dim N       As Long
    Dim wbk     As Workbook
    Dim aApps() As Application
    Dim app     As Variant
    Dim dct     As Dictionary
    Dim i       As Long
    
    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)
    N = 0

    '~~ Collect all runing Excel instances as Application
    '~~ in the array aApps
    Do While hWndMain <> 0
        Set app = GetExcelObjectFromHwnd(hWndMain)
        If Not (app Is Nothing) Then
            If N = 0 Then
                N = 1
                ReDim aApps(1 To 1)
                Set aApps(N) = app
            ElseIf checkHwnds(aApps, app.hwnd) Then
                N = N + 1
                ReDim Preserve aApps(1 To N)
                Set aApps(N) = app
            End If
        End If
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

    '~~ Collect all open Workbooks in a Dictionary and return it
    If dct Is Nothing Then Set dct = New Dictionary
    With dct
        .CompareMode = TextCompare
        For Each app In aApps
            For Each wbk In app.Workbooks
                dct.Add wbk.Name, wbk
            Next wbk
        Next app
    End With
    Set OpenWbs = dct

xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

#If Win64 Then
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As LongPtr) As Application
#Else
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As Long) As Application
#End If

#If Win64 Then
    Dim hWndDesk As LongPtr
    Dim hwnd As LongPtr
#Else
    Dim hWndDesk As Long
    Dim hwnd As Long
#End If
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    Dim sText   As String
    Dim lRet    As Long
    Dim iid     As UUID
    Dim ob      As Object
    
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then
        hwnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Do While hwnd <> 0
            sText = String$(100, Chr$(0))
            lRet = CLng(GetClassName(hwnd, sText, 100))
            If Left$(sText, lRet) = "EXCEL7" Then
                Call IIDFromString(StrPtr(IID_IDispatch), iid)
                If AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, iid, ob) = 0 Then 'S_OK
                    Set GetExcelObjectFromHwnd = ob.Application
                    GoTo xt
                End If
            End If
            hwnd = FindWindowEx(hWndDesk, hwnd, vbNullString, vbNullString)
        Loop
        
    End If
    
xt:
End Function

#If Win64 Then
    Private Function checkHwnds(ByRef xlApps() As Application, hwnd As LongPtr) As Boolean
#Else
    Private Function checkHwnds(ByRef xlApps() As Application, hwnd As Long) As Boolean
#End If
' -----------------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------------
    Const PROC = "checkHwnds"

    On Error GoTo eh
    Dim i       As Long
    
    If UBound(xlApps) = 0 Then GoTo xt

    For i = LBound(xlApps) To UBound(xlApps)
        If xlApps(i).hwnd = hwnd Then
            checkHwnds = False
            GoTo xt
        End If
    Next i

    checkHwnds = True
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Public Function GetOpenWorkbook(ByVal vWb As Variant) As Workbook
' ------------------------------------------------------------------------------
' Returns an open Workbook object or raises an error. If vWb is a full path-file
' name, the file exists but is not open it is opened. A desired ReadOnly mode
' has to be set by the caller.
' ------------------------------------------------------------------------------
    Const PROC = "GetOpenWorkbook"

    On Error GoTo eh
    Dim sTest   As String
    Dim sName   As String
    Dim sPath   As String
    Dim wb      As Workbook
    
    Set GetOpenWorkbook = Nothing
    
    If Not TypeName(vWb) = "Nothing" Then
        If TypeOf vWb Is Workbook Then
            On Error Resume Next
            sTest = vWb.Name
            If Err.Number = 0 Then
                Set GetOpenWorkbook = vWb
            Else
                On Error GoTo eh
                Err.Raise AppErr(1), ErrSrc(PROC), ERR_EXISTS_GOW04
            End If
            On Error GoTo eh
        ElseIf VarType(vWb) = vbString Then
            If InStr(vWb, "\") <> 0 Then
                '~~ A Workbook's full name is provided
                sName = Split(vWb, "\")(UBound(Split(vWb, "\")))
                sPath = Replace$(vWb, sName, vbNullString)
                With OpenWbs
                    If .Exists(sName) Then
                        '~~ A Workbook with the same name is open
                        Set wb = .Item(sName)
                        If wb.FullName <> vWb Then
                            '~~ The open Workook with the same name is from a different location
                            If FileExists(vWb) Then
                                '~~ The file still exists on the provided location
                                Err.Raise AppErr(2), ErrSrc(PROC), Replace(Replace$(ERR_EXISTS_GOW02, "<>1", wb.Path), "<>2", sPath)
                            Else
                                '~~ The Workbook file does not/no longer exists at the provivded location.
                                '~~ The open one is apparenty the ment Workbook just moved to the new location.
                                Set GetOpenWorkbook = wb
                            End If
                        Else
                            '~~ The open Workook is the one indicated by the provided full name
                            Set GetOpenWorkbook = wb
                        End If
                    Else
                        '~~ The Workbook is yet not open
                        If FileExists(vWb) Then
                            Set GetOpenWorkbook = Workbooks.Open(vWb)
                        Else
                            Err.Raise AppErr(3), ErrSrc(PROC), Replace(ERR_EXISTS_GOW06, "<>", CStr(vWb))
                        End If
                    End If
                End With
            Else
                '~~ Only a Workbook's name is provided
                With OpenWbs
                    If .Exists(vWb) Then
                        Set GetOpenWorkbook = .Item(vWb)
                    Else
                        Err.Raise AppErr(4), ErrSrc(PROC), Replace$(ERR_EXISTS_GOW03, "<>", vWb)
                    End If
                End With
            End If
        Else
            '~~ Parameter vWb is neither a Workbook object nor a string (name or full name)
            Err.Raise AppErr(5), ErrSrc(PROC), ERR_EXISTS_GOW05
        End If
    Else
        Err.Raise AppErr(6), ErrSrc(PROC), ERR_EXISTS_GOW05
    End If
    
xt: Exit Function
    
eh: If ErrMsg(ErrSrc(PROC)) = vbYes Then: Stop: Resume
End Function

Private Function TestSheet(ByVal wb As Workbook, _
                           ByVal vWs As Variant) As Worksheet
' -----------------------------------------------------------
' Returns the Worksheet object (vWs) - which may be a Work-
' sheet object or a Worksheet's name - of the Workbook (wb).
' Precondition: The Worksheet exists.
' -----------------------------------------------------------
    If VarType(vWs) = vbString Then
        Set TestSheet = wb.Worksheets(vWs)
    ElseIf TypeOf vWs Is Worksheet Then
        Set TestSheet = vWs
    End If
End Function
 
