Attribute VB_Name = "modDataParser"
Option Explicit
Option Private Module


' Usage code to copy and paste
Private Sub GetData()
'    sFromDir As String
'    sFromBook As String
'    sFromSheet As String
'    sFromCell As String
'    sToBook As String
'    sToSheet As String
'    sToCell As String
'    bOverwrite As Boolean
'    Optional bCopyHeaders As Boolean = False
'    Optional bCloseFrom As Boolean = True

    Application.Run _
        "'Personal Toolkit.xlam'!CopyValuesFromRangeToTable", _
        "", _
        "", _
        "", _
        "", _
        "", _
        "", _
        "", _
        True
End Sub


' ========== Subroutine definition

Public Sub CopyValuesFromRangeToTable( _
    ByVal sFromDir As String, _
    ByVal sFromBook As String, _
    ByVal sFromSheet As String, _
    ByVal sFromCell As String, _
    ByVal sToBook As String, _
    ByVal sToSheet As String, _
    ByVal sToCell As String, _
    ByVal bOverwrite As Boolean, _
    Optional ByVal bCopyHeaders As Boolean = False, _
    Optional ByVal bCloseFrom As Boolean = True _
)

    Dim iCalcState As Integer
    Dim wbFromBook As Workbook
    Dim rCopyRange As Range

    With Application
        Let iCalcState = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With


    ' Open source book
    Set wbFromBook = OpenWorkbook(sFromBook, sFromDir, True)

    ' Check the sFromCell
    If InStr(wbFromBook.Worksheets(sFromSheet).Range(sFromCell).CurrentRegion.Address, ":") = 0 Then
        Err.Raise _
            Number:=513, _
            Source:="Personal Toolkit.xlam!modDataParser.CopyValuesFromRangeToTable", _
            Description:="Cannot copy a single cell. Specify a range consisting of more than one cell."
    End If

    ' Get copy range
    With wbFromBook.Worksheets(sFromSheet)
        If bCopyHeaders Then
            Set rCopyRange = .Range(sFromCell).CurrentRegion
        Else
            With .Range(sFromCell)
                Set rCopyRange = .CurrentRegion.Resize(.CurrentRegion.Rows.Count - 1).Offset(1, 0)
            End With
        End If
    End With

    ' Check the stoCell
    With Workbooks(sToBook).Worksheets(sToSheet)
        If InStr(.Range(sToCell).CurrentRegion.Address, ":") = 0 Then
            If MsgBox( _
                Prompt:="The paste cell has no surrounding values. The subroutine might not function as expected. Do you want to continue?", _
                Buttons:=vbYesNo + vbExclamation, _
                Title:="Copy Range Warning" _
            ) <> vbYes Then Exit Sub
        End If
        If Not IsCellInTable(.Range(sToCell)) Then
            Call MsgBox( _
                Prompt:="The paste range is not a table. The subroutine might not function as expected.", _
                Buttons:=vbExclamation, _
                Title:="Copy Range Warning" _
            )
        End If
    End With

    ' Copy in the range
    If bOverwrite Then
        ' Remove existing data
        With Workbooks(sToBook).Worksheets(sToSheet)
            If IsCellInTable(.Range(sToCell)) Then
                Call modListObjectStuff.EmptyTableExceptFirstRow(.Range(sToCell).ListObject)
            End If
        End With

        ' Add new data
        With rCopyRange
            Workbooks(sToBook).Worksheets(sToSheet).Range(sToCell).Resize(.Rows.Count, .Columns.Count).Value = .Value
        End With
    Else
        ' Append new data -- assumes no data beneath the paste range
        With Workbooks(sToBook).Worksheets(sToSheet)
            .Cells(2 ^ 20, .Range(sToCell).Column).End(xlUp).Offset(1, 0).Resize(rCopyRange.Rows.Count, rCopyRange.Columns.Count).Value = rCopyRange.Value
        End With
    End If

    ' Close and garbage collect
    If bCloseFrom Then wbFromBook.Close SaveChanges:=False
    Set wbFromBook = Nothing
    Set rCopyRange = Nothing


    With Application
        .Calculation = iCalcState
        .ScreenUpdating = True
    End With
End Sub


Private Function IsCellInTable(ByVal rCell As Range) As Boolean
    IsCellInTable = False
    On Error Resume Next
        IsCellInTable = (rCell.ListObject.Name <> "")
    On Error GoTo 0
End Function


