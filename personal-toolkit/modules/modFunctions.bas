Attribute VB_Name = "modFunctions"
Option Explicit


' Make a separate common functions add-in


' ====================================================================================================
'''
' https://docs.microsoft.com/en-us/office/vba/api/Excel.Application.MacroOptions
'''

Public Function MAKELIST(ByVal rValues As Range) As String
Attribute MAKELIST.VB_Description = "Converts a range to a SQL list\nMAKELIST(<range>)"
Attribute MAKELIST.VB_ProcData.VB_Invoke_Func = " \n14"
    Let MAKELIST = "'" & WorksheetFunction.TextJoin("'," & Chr(10) & "'", True, rValues) & "'"
End Function

Sub RegisterUDF()
    Dim s As String
    s = "Converts a range to a SQL list" & vbLf _
    & "MAKELIST(<range>)"

    Application.MacroOptions Macro:="MAKELIST", Description:=s, Category:=14
End Sub

Sub UnregisterUDF()
    Application.MacroOptions Macro:="MAKELIST", Description:=Empty, Category:=Empty
End Sub


' ====================================================================================================
Public Function OpenWorkbook(ByVal FileName As String, ByVal FileFolder As String, Optional ByVal bReadOnly As Boolean = False) As Workbook
    On Error GoTo OpenBook
        Set OpenWorkbook = Workbooks(FileName)
        Exit Function

OpenBook:
    Set OpenWorkbook = Workbooks.Open(FileFolder & GetDelimiter(FileFolder) & FileName, ReadOnly = bReadOnly)
End Function


Private Function GetDelimiter(ByVal sUrl As String) As String
    If InStr(sUrl, "/") > 0 Then
        Let GetDelimiter = "/"
    ElseIf InStr(sUrl, "\") > 0 Then
        Let GetDelimiter = "\"
    Else
        Let GetDelimiter = ""
    End If
End Function


'  https://stackoverflow.com/a/50563889/8213085
Public Function IsLoaded(FormName As String) As Boolean
    Dim Frm As Object

    For Each Frm In VBA.UserForms
        If Frm.Name = FormName Then
            Let IsLoaded = True
            Exit Function
        End If
    Next Frm

    Let IsLoaded = False
End Function


Public Function IsInArray(StringToFind As String, vArray As Variant) As Boolean
    ' https://stackoverflow.com/a/38268261/8213085
    Dim i
    For i = LBound(vArray) To UBound(vArray)
        If vArray(i) = StringToFind Then
            Let IsInArray = True
            Exit Function
        End If
    Next i
    Let IsInArray = False
End Function


Private Sub PrintArrayEntries(vArray As Variant)
    Dim i
    For i = LBound(vArray) To UBound(vArray)
        Debug.Print vArray(i)
    Next i
End Sub


' https://superuser.com/a/1505673
Public Function UniqueStr(sStr As String, sDelim As String) As String
    Dim dic      As Object
    Dim strArr() As String
    Dim strPart  As Variant
    Dim temp     As String
    Dim key      As Variant

    Set dic = CreateObject("Scripting.Dictionary")
    Let strArr = Split(sStr, sDelim)
    Let temp = ""

    For Each strPart In strArr
        On Error Resume Next
            dic.Add Trim(strPart), Trim(strPart)
        On Error GoTo 0
    Next strPart

    For Each key In dic
        Let temp = temp & key & sDelim
    Next key

    Let UniqueStr = Left(temp, Len(temp) - Len(sDelim))
End Function


'====================IFERROR VARIANTS===================='

Public Function tlkt_IfValueBase(Value As Variant, AlternativeValue As Variant, ValueToCheck As Variant) As Variant
    If Value = ValueToCheck Then
        Let tlkt_IfValueBase = AlternativeValue
    Else
        Let tlkt_IfValueBase = Value
    End If
End Function


Public Function tlkt_IfZero(Value As Variant, ValueIfZero As Variant) As Variant
    Let tlkt_IfZero = tlkt_IfValueBase(Value, ValueIfZero, 0)
End Function
Public Function tlkt_IfBlank(Value As Variant, ValueIfBlank As Variant) As Variant
    Let tlkt_IfBlank = tlkt_IfValueBase(Value, ValueIfBlank, "")
End Function


'====================FILE IN USE===================='

Public Function tlkt_ExcelFileInUseBy(OpenWorkbook As String) As String
    '  Author: Selson Muniappen
    If Workbooks(OpenWorkbook).Name = Workbooks(OpenWorkbook).FullName Then
        tlkt_ExcelFileInUseBy = Application.userName
        Exit Function
    End If

    Dim FilePath                As String
    Dim strTempFile             As String
    Dim iPos                    As Integer
    Dim iRetVal                 As Integer
    Dim objFSO                  As Object
    Dim objWMIService           As Object
    Dim objFileSecuritySettings As Object
    Dim objSD                   As Object

    Let FilePath = Workbooks(OpenWorkbook).FullName
    Let iPos = InStrRev(FilePath, "\")
    Let strTempFile = Left(FilePath, iPos - 1) & "\~$" & Mid(FilePath, iPos + 1)
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If objFSO.FileExists(strTempFile) Then
'        On Error GoTo ErrHandle
        Set objWMIService = GetObject("winmgmts:")
        Set objFileSecuritySettings = objWMIService.Get("Win32_LogicalFileSecuritySetting='" & strTempFile & "'")
        Let iRetVal = objFileSecuritySettings.GetSecurityDescriptor(objSD)

        If iRetVal = 0 Then
            tlkt_ExcelFileInUseBy = tlkt_GetUserFullName(objSD.Owner.Name)
        Else
            tlkt_ExcelFileInUseBy = "unknown"
        End If
    Else
'ErrHandle:
        tlkt_ExcelFileInUseBy = vbNullString
    End If
End Function


Public Function tlkt_GetUserFullName(userName As String) As String

    Dim WSHnet     As Variant
    Dim UserDomain As Variant
    Dim objUser    As Variant

    On Error GoTo ErrorHandler
    Set WSHnet = CreateObject("WScript.Network")
    Let UserDomain = WSHnet.UserDomain
    Set objUser = GetObject("WinNT://" & UserDomain & "/" & userName & ",user")

    tlkt_GetUserFullName = objUser.FullName
    Exit Function

ErrorHandler:
    tlkt_GetUserFullName = ""

End Function


'====================IS WORKBOOK OPEN===================='

Public Function tlkt_IsWorkbookOpen(FileName As String)
    '  https://stackoverflow.com/a/9373914
    Dim ff    As Long
    Dim ErrNo As Long

    On Error Resume Next
        ff = FreeFile()
        Open FileName For Input Lock Read As #ff
        Close ff
        ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
        Case 0:    tlkt_IsWorkbookOpen = False
        Case 70:   tlkt_IsWorkbookOpen = True
        Case Else: Error ErrNo
    End Select
End Function


'====================DOES WORKSHEET EXIST===================='

Public Function tlkt_WorksheetExists(SheetName As String, Optional wBook As Workbook) As Boolean
    Dim wSheet As Worksheet

    If IsMissing(wBook) Then Set wBook = ActiveWorkbook

    On Error Resume Next
        Set wSheet = wBook.Sheets(SheetName)
    On Error GoTo 0

    tlkt_WorksheetExists = Not wSheet Is Nothing
End Function


'====================CHECKDIGIT===================='

'  Compare FLOOR_PRECISE to INT
Private Function tlkt_Floor(arg As Double)
    tlkt_Floor = Application.Floor_Precise(arg, 1)
End Function
Public Function tlkt_CheckDigit(numpersona As Long) As String
    tlkt_CheckDigit = Format(CInt(Right(tlkt_Floor(Application.Product(numpersona, 830000) / 97), 2)), "00")
End Function


'====================UNIVERSAL NAMING CONVENTION PATH -- MAPPED NETWORK DRIVE===================='

Public Function tlkt_GetNetworkPath(ByVal DriveName As String) As String
    '  http://www.excelfox.com/forum/showthread.php/134-VBA-to-Return-the-UNC(Universal-Naming-Convention)-Path-of-a-Mapped-Network-Drive
    Dim objNtWork   As Object
    Dim objDrives   As Object
    Dim lngLoop     As Long

    Set objNtWork = CreateObject("WScript.Network")
    Set objDrives = objNtWork.EnumNetworkDrives

    For lngLoop = 0 To objDrives.Count - 1 Step 2
        If UCase(objDrives.Item(lngLoop)) = UCase(DriveName) Then
            tlkt_GetNetworkPath = objDrives.Item(lngLoop + 1)
            Exit For
        End If
    Next
End Function

Private Function tlkt_GetNetworkPath_test()
    MsgBox tlkt_GetNetworkPath("Z:")
End Function

