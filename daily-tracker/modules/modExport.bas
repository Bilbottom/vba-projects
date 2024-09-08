Attribute VB_Name = "modExport"
Option Explicit
Option Private Module


Public Sub WriteToCSV(ByVal wsData As Worksheet, ByVal sOutFile As String)
    '''
    ' Write the contents of `wsData` to the CSV file `sOutFile`.
    '''
    Dim i      As Integer
    Dim j      As Integer
    Dim sWrite As String

    Open sOutFile For Output As #1

    With wsData.Cells(1).CurrentRegion
        For i = 1 To .Rows.Count
            For j = 1 To .Columns.Count
                Let sWrite = Replace(.Cells(i, j).Value, ",", "")

                If j = .Columns.Count Then
                    Write #1, sWrite
                Else
                    Write #1, sWrite,
                End If
            Next j
        Next i
    End With

    Close #1
End Sub


Private Sub ExportTrackerDataForPython()
    Const sOutFile As String = _
        "C:\Users\billiam\Downloads\daily-tracker-data.csv"

    Call WriteToCSV( _
        wsData:=wsExport, _
        sOutFile:=sOutFile _
    )
End Sub
