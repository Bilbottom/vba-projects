Attribute VB_Name = "modWorkbookSubs"
Option Explicit
Option Private Module

Private Const tlktLocation As String = "C:\Users\bilbottom\Documents\Excel\Add-Ins\"
Private Const sAddInName   As String = "Personal Toolkit"


Public Sub tlkt_SaveAsAddIn()
    ThisWorkbook.Save

    AddIns(sAddInName).Installed = False
        Application.DisplayAlerts = False

            On Error GoTo ErrorHandler
            Workbooks(sAddInName & ".xlsm").SaveAs _
                FileName:=tlktLocation & sAddInName & ".xlam", _
                FileFormat:=xlOpenXMLAddIn

        Application.DisplayAlerts = True
    AddIns(sAddInName).Installed = True

    Call MsgBox(Prompt:="Add-In saved successfully.")
'    Call ExportAllModules(ThisWorkbook)
    ThisWorkbook.Save
    ThisWorkbook.Close
    Exit Sub

ErrorHandler:
    Call MsgBox( _
        Prompt:="Add-In could not be saved. Do you have it enabled elsewhere?", _
        Buttons:=vbCritical _
    )
End Sub
