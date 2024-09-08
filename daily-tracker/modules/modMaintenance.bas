Attribute VB_Name = "modMaintenance"
Option Explicit
Option Private Module

'''
' Subroutines to help with the maintenance of the Workbook.
'
' Currently not in use, needs to be re-jigged.
'''

'''
' There are few options for setting custom PivotTable dates.
'
' The approach taken in this module is to enumerate through the dates and toggle
' them to `Visible=True` if we want to see them, and `Visible=False` otherwise.
'
' However, the PivotTable will error if all of the items are `Visible=False`, so
' we first make all of them visible and only then go through and toggle the ones
' we don't want to `Visible=False`.
'''

' Find and Replace text
Public Sub FindAndReplace()
    Dim rColsForReplace As Range
    Set rColsForReplace = GetColsForReplace()

    If rColsForReplace Is Nothing Then Exit Sub

    rColsForReplace.Replace _
        What:=wsMaintenance.[sReplace].Value, _
        Replacement:=wsMaintenance.[sWith].Value, _
        LookAt:=xlWhole, _
        MatchCase:=True

    ThisWorkbook.RefreshAll
End Sub


Private Function GetColsForReplace() As Range
    Dim cbxTask As Object
    Dim cbxDetail As Object

    Set cbxTask = wsMaintenance.Shapes("cbxReplaceTask").OLEFormat.Object
    Set cbxDetail = wsMaintenance.Shapes("cbxReplaceDetail").OLEFormat.Object

    If cbxTask = 1 And cbxDetail = 1 Then
        Set GetColsForReplace = Union( _
            wsData.ListObjects("Data").ListColumns("Task").DataBodyRange, _
            wsData.ListObjects("Data").ListColumns("Detail").DataBodyRange _
        )
    ElseIf cbxTask = 1 Then
        Set GetColsForReplace = wsData.ListObjects("Data").ListColumns("Task").DataBodyRange
    ElseIf cbxDetail = 1 Then
        Set GetColsForReplace = wsData.ListObjects("Data").ListColumns("Detail").DataBodyRange
    Else
        Set GetColsForReplace = Nothing
    End If
End Function


Public Sub SetMaintenanceDateFilter()
    '''
    ' Uses the `wsMaintenance` ranges `lFrom` and `lTo` to filter the
    ' `wsMaintenance` PivotTable.
    '''
    Dim CalcState As XlCalculation
    With Application
        Let CalcState = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Dim i As Long

    Call SetAllDatesVisible

    With wsMaintenance.PivotTables("MaintenancePivot").PivotFields("Date")
        On Error Resume Next
            For i = DateSerial(2019, 4, 23) To Date
                If i < wsMaintenance.[lFrom] Or i > wsMaintenance.[lTo] Then
                    .PivotItems(Format(i, "yyyy-mm-dd")).Visible = False
                End If
            Next i
        On Error GoTo 0
    End With

    With Application
        .Calculation = CalcState
        .ScreenUpdating = True
    End With
End Sub


Private Sub SetAllDatesVisible()
    '''
    ' Toggle all of the PivotTable dates to `Visible=True`.
    '''
    Dim CalcState As XlCalculation
    With Application
        Let CalcState = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

    Dim i As Long

    With wsMaintenance.PivotTables("MaintenancePivot").PivotFields("Date")
        On Error Resume Next
            For i = DateSerial(2019, 4, 23) To Date
                .PivotItems(Format(i, "yyyy-mm-dd")).Visible = True
            Next i
        On Error GoTo 0
    End With

    With Application
        .Calculation = CalcState
        .ScreenUpdating = True
    End With
End Sub
