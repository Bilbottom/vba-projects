Attribute VB_Name = "modRibbonSubs"
Option Explicit
Option Private Module


' ========== Live

' Delete objects in the ActiveSheet
Public Sub DeleteObjects()
    Dim cObject As Object

    With ActiveSheet
        For Each cObject In .ChartObjects
            cObject.Delete
        Next cObject
    End With
End Sub


' Unhide all sheets in ActiveWorkbook
Public Sub UnhideAllSheets()
    Application.ScreenUpdating = False

    Dim wSheet As Worksheet

    For Each wSheet In ActiveWorkbook.Worksheets
        wSheet.Visible = xlSheetVisible
    Next wSheet

    Application.ScreenUpdating = True
End Sub


' Colour text in selection to code highlighting
Public Sub ApplyLogicColour()
    Application.ScreenUpdating = False

    Call ColourCellsWithPCO(Selection)

    Application.ScreenUpdating = True
End Sub


' Set grouping location for ActiveSheet
Public Sub SetGroupingLocation()
    On Error Resume Next
        With ActiveSheet.Outline
            .SummaryRow = xlAbove
            .SummaryColumn = xlLeft
        End With
    On Error GoTo 0
End Sub


' Create project directory
Public Sub MakeProjectDir()
    Dim sInput As Variant
    Dim sDef   As String

    On Error Resume Next
        Let sDef = Trim(CStr(ActiveCell.Value))
    On Error GoTo 0

    Let sInput = InputBox( _
        Prompt:="Enter project name:", _
        Title:="Make Project Directory", _
        Default:=sDef _
    )
    Call MakeProjectDirectory(sProjectName:=CStr(sInput))
End Sub


' Set preferred number format
Public Sub NumberFormatDD(ByVal sFormatID As String)
    Dim sFormat As String

    Select Case sFormatID
        ' ddFormatCurrency
        Case "itmFormatComma"
            Let sFormat = "#,##0"
        Case "itmFormatCurrencyK"
            Let sFormat = """£""0.00,""k"""
        Case "itmFormatCurrencyM"
            Let sFormat = """£""0.00,,""m"""
        Case "itmFormatBlank"
            Let sFormat = ";;;"
        ' ddFormatDatetime
        Case "itmFormatYYYYMMDDHHMM"
            Let sFormat = "yyyy-mm-dd hh:mm"
        Case "itmFormatYYYYMMDD"
            Let sFormat = "yyyy-mm-dd"
        Case "itmFormatDD-MMM"
            Let sFormat = "dd-mmm"
        Case "itmFormatMMM-YY"
            Let sFormat = "mmm-yy"
        Case "itmFormatHHMM"
            Let sFormat = "hh:mm"
        Case "itmFormatHHMMSS"
            Let sFormat = "hh:mm:ss"
        ' ddFormatColours
'        Case ""
'            Let sFormat = ""
        ' else
        Case Else
            Let sFormat = ""
    End Select

    If sFormat <> "" Then Selection.NumberFormat = sFormat
End Sub


' Set Alt button behaviour
Public Sub SetAltButtonBehaviour(ByVal sID As String)
    Select Case sID
        Case "itmButtonAltValues"
            Call AssignPlusMinusToAlt
        Case "itmButtonAltSelection"
            Call AssignAltKeyForSelection
        Case Else
            Err.Raise _
                Number:=513, _
                Description:="Unrecognised control.id in SetAltButtonBehaviour"
    End Select
End Sub


' Convert text to hyperlink
Public Sub MakeTextHyperlink()
    Dim rCell As Range

    For Each rCell In Selection
        rCell.Parent.Hyperlinks.Add _
            Anchor:=rCell, _
            Address:=rCell.Value, _
            TextToDisplay:=rCell.Value
    Next rCell
End Sub


' ========== Pending

' Center Picture
Private Sub CenterPictures()
    '  https://www.extendoffice.com/documents/excel/4924-excel-picture-center-cell.html
    Const Cell As String = "F4"

    With ActiveSheet.Shapes(Selection.Name)
        .Top = Range(Cell).Top + (Range(Cell).Height - .Height) / 2
        .Left = Range(Cell).Left + (Range(Cell).Width - .Width) / 2
    End With
End Sub


' Make Cells Empty
Private Sub EmptyCells()
    Dim Cell As Range

    For Each Cell In Selection
        If Cell.Value = "" Then Cell.ClearContents
    Next Cell
End Sub


' Julian and Gregorian
Private Function ConvertJulianToGregorian(ByVal sYYYYDDD As String) As String
    Let ConvertJulianToGregorian = Format( _
        Expression:=DateSerial(CLng(Left(sYYYYDDD, 4)), 1, CLng(Right(sYYYYDDD, 3))), _
        Format:="yyyymmdd" _
    )
End Function
Private Function ConvertGregorianToJulian(ByVal sYYYYMMDD As String) As String
    Let ConvertGregorianToJulian = "" _
        & Left(sYYYYMMDD, 4) _
        & Format(DatePart( _
            Date:=DateSerial(CLng(Left(sYYYYMMDD, 4)), CLng(Mid(sYYYYMMDD, 5, 2)), CLng(Right(sYYYYMMDD, 2))), _
            Interval:="y" _
        ), "000")
End Function


Private Sub Conversion()
    Debug.Print ConvertGregorianToJulian("20210118")
    Debug.Print ConvertJulianToGregorian("2021018")
End Sub
