Attribute VB_Name = "modListObjectStuff"
Option Explicit
Option Private Module


Public Sub EmptyTableExceptFirstRow(loTable As ListObject)
    Application.ScreenUpdating = False

    With loTable
        With .DataBodyRange
            If .Rows.Count = 1 Then Exit Sub
            .Resize(.Rows.Count - 1).Offset(1).Clear
        End With

        .Resize .DataBodyRange.Resize(2).Offset(-1)
    End With

    Application.ScreenUpdating = True
End Sub


Public Sub EmptyTable(loTable As ListObject)
    With loTable
        .DataBodyRange.Clear
        .Resize .DataBodyRange.Resize(2).Offset(-1)
    End With
End Sub


Function IsActiveCellInTable() As Boolean
    Dim rngActiveCell
    Set rngActiveCell = ActiveCell

    On Error Resume Next
        rngActiveCell = (rngActiveCell.ListObject.Name <> "")
    On Error GoTo 0

    Let IsActiveCellInTable = rngActiveCell
End Function
