VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MatrixConverter 
   Caption         =   "Matrix to List"
   ClientHeight    =   2064
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   6045
   OleObjectBlob   =   "UF_MatrixConverter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_MatrixConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Enum DecomposeType
    RowsByColumns = 0
    ColumnsByRows = 1
End Enum


Private Sub UserForm_Initialize()
    Me.Height = 132
    Me.Width = 312

    On Error Resume Next
    Me.RefEditMatrix.Value = "=" & Selection.Address
End Sub


Private Sub UserForm_Terminate()
    Unload Me
End Sub


Private Sub CommandButtonConvert_Click()
    ' https://www.extendoffice.com/documents/excel/3539-excel-convert-array-to-list.html

    Dim CalcState   As Integer
    Dim MatrixRange As Range
    Dim ListCell    As Range
    Dim DecType     As DecomposeType

    With Application
        Let CalcState = .Calculation
        .Calculation = xlCalculationManual
    End With

    If Me.RefEditMatrix.Value = vbNullString Then
        MsgBox "You need to supply a matrix range."
        Exit Sub
    End If

    If Me.RefEditCell.Value = vbNullString Then
        MsgBox "You need to supply the table destination."
        Exit Sub
    End If

    Set MatrixRange = Range(Me.RefEditMatrix.Value)
    Set ListCell = Range(Me.RefEditCell.Value)

    Select Case True
        Case Me.OptionButtonRowsByColumns.Value
            DecType = RowsByColumns
        Case Me.OptionButtonColumnsByRows.Value
            DecType = ColumnsByRows
    End Select

    Me.Hide

    ' Transform the data
    Call ConvertTableToListBase( _
        rMatrixRange:=MatrixRange, _
        rListCell:=ListCell, _
        eDecomposeType:=DecType, _
        bIncludeCorner:=Me.CheckBoxIncludeCorner.Value _
    )

    Application.Calculation = CalcState

End Sub


Public Sub ConvertTableToListBase(rMatrixRange As Range, rListCell As Range, Optional eDecomposeType As DecomposeType = 0, Optional bIncludeCorner As Boolean = False)

    Application.ScreenUpdating = False

    Dim C As Long 'Counter
    Dim i As Long
    Dim j As Long

    Let C = 0

    If eDecomposeType = RowsByColumns Then
        ' Transform the data - Rows | Columns | Value
        For i = 2 To rMatrixRange.Rows.Count
            For j = 2 To rMatrixRange.Columns.Count

                If bIncludeCorner Then
                    rListCell.Offset(C, 0).Value = rMatrixRange.Cells(1, 1).Value
                    rListCell.Offset(C, 1).Value = rMatrixRange.Cells(i, 1).Value
                    rListCell.Offset(C, 2).Value = rMatrixRange.Cells(1, j).Value
                    rListCell.Offset(C, 3).Value = rMatrixRange.Cells(i, j).Value
                Else
                    rListCell.Offset(C, 0).Value = rMatrixRange.Cells(i, 1).Value
                    rListCell.Offset(C, 1).Value = rMatrixRange.Cells(1, j).Value
                    rListCell.Offset(C, 2).Value = rMatrixRange.Cells(i, j).Value
                End If

                Let C = C + 1

            Next j
        Next i
    Else
        ' Transform the data - Columns | Rows | Value
        For i = 2 To rMatrixRange.Columns.Count
            For j = 2 To rMatrixRange.Rows.Count

                If bIncludeCorner Then
                    rListCell.Offset(C, 0).Value = rMatrixRange.Cells(1, 1).Value
                    rListCell.Offset(C, 1).Value = rMatrixRange.Cells(1, i).Value
                    rListCell.Offset(C, 2).Value = rMatrixRange.Cells(j, 1).Value
                    rListCell.Offset(C, 3).Value = rMatrixRange.Cells(j, i).Value
                Else
                    rListCell.Offset(C, 0).Value = rMatrixRange.Cells(1, i).Value
                    rListCell.Offset(C, 1).Value = rMatrixRange.Cells(j, 1).Value
                    rListCell.Offset(C, 2).Value = rMatrixRange.Cells(j, i).Value
                End If

                Let C = C + 1

            Next j
        Next i
    End If

    Application.ScreenUpdating = True

End Sub

