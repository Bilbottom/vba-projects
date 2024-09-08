Attribute VB_Name = "Pending"
Option Explicit
Option Private Module


'==========DELETE FORMULA NAMES=========='

Private Sub DeleteFormulaNames()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Dim ContinueLoop As Boolean

RestartLoop:

    Let ContinueLoop = DeleteLoop(Workbooks(""))
    If ContinueLoop Then GoTo RestartLoop

    Application.ScreenUpdating = True

End Sub


Private Function DeleteLoop(wBook As Workbook) As Boolean

    Dim i            As Long
    Dim OldNameCount As Long
    Dim NewNameCount As Long
    Dim ContinueLoop As Boolean

    With wBook

        Let OldNameCount = .Names.Count
        Let ContinueLoop = False

        For i = 1 To .Names.Count

            On Error GoTo ExitLoop
                .Names(i).Delete
                Let ContinueLoop = True

        Next i

        Let NewNameCount = .Names.Count

    End With

ExitLoop:

    On Error GoTo 0
    Debug.Print OldNameCount & " to " & NewNameCount, ContinueLoop

    Let DeleteLoop = ContinueLoop

End Function


'==========REMOVE HYPERLINKS=========='

Private Sub RemoveAllHyperlinks()

    Dim wSheet As Worksheet

    For Each wSheet In Workbooks("").Worksheets

        Debug.Print wSheet.Name
        wSheet.Cells.Hyperlinks.Delete

    Next wSheet

End Sub



'==========CHECK FOR FORMULAS (SPECIFICALLY BEFORE COPYING DATA ETC)=========='

Private Sub CheckForFormulas()
    Dim rSelect As Range
    Dim iCell   As Range

    For Each iCell In Selection
        If iCell.HasFormula Then
            If rSelect Is Nothing Then
                Set rSelect = iCell
            Else
                Set rSelect = Union(rSelect, iCell)
            End If
        End If
    Next iCell

    If rSelect Is Nothing Then Exit Sub
    rSelect.Select
End Sub



'==========CONVERT FORMULAS AS VALUES TO FORMULAS=========='

Private Sub ConvertIntoFormulas()
    Dim rCell As Range

    For Each rCell In Selection
        rCell.Formula = rCell.Formula
    Next rCell
End Sub



'==========ADD ALL COMBINATIONS TO WORKSHEET=========='

Private Sub MakeTFList()
    Dim i  As Integer
    Dim vO As Variant
    Dim vP As Variant
    Dim vR As Variant

    Const sToSplit = "TRUE,FALSE"

    Let i = 0
    For Each vO In Split(sToSplit, ",")
        For Each vP In Split(sToSplit, ",")
            For Each vR In Split(sToSplit, ",")
'                Debug.Print "vO: " & vO, "vP: " & vP, "vR: " & vR

                With ActiveCell
                    .Offset(i, 0).Value = vO
                    .Offset(i, 1).Value = vP
                    .Offset(i, 2).Value = vR
                End With

                Let i = i + 1
            Next vR
        Next vP
    Next vO
End Sub
