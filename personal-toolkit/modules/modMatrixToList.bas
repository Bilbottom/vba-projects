Attribute VB_Name = "modMatrixToList"
Option Explicit
Option Private Module


' https://www.extendoffice.com/documents/excel/3539-excel-convert-array-to-list.html
Private Sub ConvertTableToList()
    Dim CalcState   As Integer
    Dim MatrixRange As Range
    Dim ListCell    As Range

    With Application
        Let CalcState = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    ' Get user input
    On Error Resume Next
        ' Get matrix range
        Set MatrixRange = Application.InputBox( _
            Prompt:="Select the matrix to transform, with the headers.", _
            Title:="Convert Matrix to Table", _
            Default:=Selection.Address, _
            Type:=8 _
        )

        If MatrixRange Is Nothing Then Exit Sub

        ' Get paste top-left
        Set ListCell = Application.InputBox( _
            Prompt:="Select the cell to put the list.", _
            Title:="Convert Matrix to Table", _
            Default:=Selection.Address, _
            Type:=8 _
        ).Cells(1)

        If ListCell Is Nothing Then Exit Sub

        ' Get RowsToColumns or ColumnsToRows

    On Error GoTo 0

    ' Transform the data
    Call ConvertTableToListBase( _
        rMatrixRange:=MatrixRange, _
        rListCell:=ListCell, _
        eDecomposeType:=RowsByColumns _
    )

    With Application
        .Calculation = CalcState
        .ScreenUpdating = True
    End With
End Sub





