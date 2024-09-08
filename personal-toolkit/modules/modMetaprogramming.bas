Attribute VB_Name = "modMetaprogramming"
Option Explicit
Option Private Module


' Microsoft Visual Basic for Applications Extensibility M.m

Private Sub AddCodeToModule(wBook As Workbook, vbModule As String, sTextToAdd As String)

    Dim xPro  As VBIDE.VBProject
    Dim xCom  As VBIDE.VBComponent
    Dim xMod  As VBIDE.CodeModule

    With wBook
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents(vbModule)
        Set xMod = xCom.CodeModule

        With xMod
            .InsertLines .CountOfLines + 1, sTextToAdd
        End With
    End With

End Sub


Private Sub AddCode()

    Const sBook As String = "Book.xlsx"
    Const sModule As String = "Sheet1"

    Dim wBook As Workbook
'    Set wBook = ThisWorkbook
    Set wBook = Workbooks(sBook)

'    On Error Resume Next
    AddCodeToModule wBook, sModule, ""
    AddCodeToModule wBook, sModule, "Private Sub ToggleVisibility()"
    AddCodeToModule wBook, sModule, "    With Me"
    AddCodeToModule wBook, sModule, "        Select Case .Visible"
    AddCodeToModule wBook, sModule, "            Case xlSheetVisible"
    AddCodeToModule wBook, sModule, "                .Visible = xlSheetVeryHidden"
    AddCodeToModule wBook, sModule, "            Case Else"
    AddCodeToModule wBook, sModule, "                .Visible = xlSheetVisible"
    AddCodeToModule wBook, sModule, "        End Select"
    AddCodeToModule wBook, sModule, "    End With"
    AddCodeToModule wBook, sModule, "End Sub"

End Sub
