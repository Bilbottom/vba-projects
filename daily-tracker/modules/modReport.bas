Attribute VB_Name = "modReport"
Option Explicit
Option Private Module

'''
' Subroutines to send the Dashboard to managers.
'
' Currently not in use, needs to be re-jigged.
'''
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function CopyEnhMetaFileA Lib "Gdi32" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Private Declare PtrSafe Function DeleteEnhMetaFile Lib "Gdi32" (ByVal hdc As Long) As Long


Public Sub MakeReportEmail()
    '''
    ' Email the dashboard to management.
    '
    ' Currently disabled.
    '''
    Exit Sub

    Application.ScreenUpdating = False

    Const PathEMF As String = "C:\Users\bilbottom\Downloads\"
    Const FileEMF As String = "Week_Ending_Report.emf"
    Dim sToday As String
    Let sToday = Format(Date, "yyyy-mm-dd")

    Call ExportObjectToEMF( _
        Obj:=wsDashboard.Range("F3:W18"), _
        FilePath:=PathEMF & FileEMF _
    )

    With CreateObject("Outlook.Application").CreateItem(olMailItem)
        .Attachments.Add PathEMF & FileEMF, olByValue, 0
        .Display
        .To = ""
        .CC = ""
        .Subject = "Working Report " & sToday
        .HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>" _
            & "Hi," & "<br>" & "<br>" _
            & "Find my working report for week ending " & sToday & " below." & "<br>" & "<br>" & "<br>" _
            & "<p style=""margin-left: 40px""><img src=""cid:Week_Ending_Report.emf""></p>" & "<br>" _
            & "Regards"
'            & .HTMLBody
    End With

    Application.ScreenUpdating = False
End Sub


Private Sub HideGraph()
    '''
    ' Make the size of the selected Chart too small to see.
    '''
    With Selection
        .Top = 0
        .Left = 0
        .Width = 0.01
        .Height = 0.01
    End With
End Sub


'====================================================================================================

Private Sub ExportToImageEMF()
    '  Example
    Call ExportObjectToEMF(ActiveSheet.Range("A1:C3"), "C:\Users\bilbottom\Downloads\emf_file.emf")
'    MsgBox "Export finished."
End Sub


Public Sub ExportObjectToEMF(Obj As Object, FilePath As String)
    If Dir(FilePath) <> "" Then Kill FilePath
    If ExportToEMF(Obj, FilePath) = "" Then MsgBox "An error occurred when exporting the file.", vbCritical
End Sub


Private Function ExportToEMF(xObj As Object, FileName As String, Optional Appearance, Optional Format, Optional Size) As String
    Let ExportToEMF = FileName

    If TypeName(xObj.Parent) = "Chart" Then
        xObj.Parent.CopyPicture Appearance, Format, Size
    ElseIf TypeName(xObj) <> "Chart" Then
        xObj.CopyPicture Appearance, Format
    Else
        xObj.CopyPicture Appearance, Format, Size
    End If

    OpenClipboard 0
        If DeleteEnhMetaFile(CopyEnhMetaFileA(GetClipboardData(14), ExportToEMF)) = 0 Then ExportToEMF = ""
    CloseClipboard
End Function
