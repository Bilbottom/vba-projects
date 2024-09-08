Attribute VB_Name = "ColourPCO"
Option Explicit
Option Private Module


Public Sub ColourCellsWithPCO(rSelection As Range)
    Dim iCell As Range

    For Each iCell In rSelection
        iCell.Font.Color = 0

        If Left(iCell.Value, 2) = "//" Then
            iCell.Font.Color = -8355712
        Else
            FixApostrophes iCell
            ColourKeywords iCell
            ColourStringLiterals iCell
            ColourVariableNames iCell
        End If
    Next iCell
End Sub


Private Sub FormatExistingSegmentation()
    Const sExistSeg As String = "Existing Segmentation"
    Dim rRangeToFormat As Range
    Dim iCell          As Range

    With ThisWorkbook.Worksheets(sExistSeg)
        Set rRangeToFormat = .Range(.Cells(2, 2), .Cells(.Cells(.Rows.Count, 2).End(xlUp).Row, 2))
    End With

    For Each iCell In rRangeToFormat
        If iCell.Offset(-1, 0).Value = "" Then
            '  Header cell
            iCell.Font.Bold = True
        ElseIf iCell.Offset(-1, 0).Font.Bold Then
            '  First logic cell
            If iCell.Offset(1, 0) = "" Then
                '  Only one logic cell
                ApplyBorderAround iCell
            Else
                '  At least two logic cells
                With iCell
                    ApplyBorderAround .Resize(1 + .End(xlDown).Row - .Row)
                End With
            End If
        End If

        If Left(iCell.Value, 2) = "//" Then
            iCell.Font.Color = -8355712
        Else
            FixApostrophes iCell
            ColourKeywords iCell
            ColourStringLiterals iCell
            ColourVariableNames iCell
        End If
    Next iCell
End Sub


' ========== Underlying subroutines
Private Sub ApplyBorderAround(rSelection As Range)
    With rSelection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone

        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End With
End Sub


Private Sub FixApostrophes(rCell As Range)
    If Left(rCell.Value, 6) = "DV SDS" Then
        rCell.Value = "''" & rCell.Value
    ElseIf Left(rCell.Value, 20) = "Current Account Data" Then
        rCell.Value = "''" & rCell.Value
    End If
End Sub


Private Sub ColourKeywords(rCell As Range)
    '  Assumes that each keywords appears at most once per line

    Dim FindPos     As Long
    Dim iCounter    As Long
    Dim sKeywords() As String

    Let sKeywords = Split("AND;OR;NOT;IF;THEN;ELSE;ENDIF;INLIST;INCLUDES;EXCLUDES;ISNULL", ";")

    For iCounter = LBound(sKeywords) To UBound(sKeywords)
        Let FindPos = InStr(1, rCell.Value, sKeywords(iCounter))
        If FindPos > 0 Then
            rCell.Characters(Start:=FindPos, Length:=Len(sKeywords(iCounter))).Font.Color = -65536
        End If
    Next iCounter
End Sub


Private Sub ColourStringLiterals(rCell As Range)
    Dim FindPos As Long
    Dim NextPos As Long

    If (Len(rCell.Value) - Len(Replace(rCell.Value, """", ""))) Mod 2 = 1 Then
        Debug.Print "Unbalance quotation marks in cell " & rCell.Address(0, 0)
        Exit Sub
    End If

    Let FindPos = InStr(1, rCell.Value, """")

    While FindPos > 0
        Let NextPos = InStr(FindPos + 1, rCell.Value, """")
        rCell.Characters(Start:=FindPos, Length:=(1 + NextPos - FindPos)).Font.Color = -8388480
        Let FindPos = InStr(NextPos + 1, rCell.Value, """")
    Wend
End Sub


Private Sub ColourVariableNames(rCell As Range)
    Dim FindPos As Long
    Dim NextPos As Long

    If (Len(rCell.Value) - Len(Replace(rCell.Value, "'", ""))) Mod 2 = 1 Then
        Debug.Print "Unbalance apostrophe marks in cell " & rCell.Address(0, 0)
        Exit Sub
    End If

    Let FindPos = InStr(1, rCell.Value, "'")

    While FindPos > 0
        Let NextPos = InStr(FindPos + 1, rCell.Value, "'")
        rCell.Characters(Start:=FindPos, Length:=(1 + NextPos - FindPos)).Font.Color = -16744448
        Let FindPos = InStr(NextPos + 1, rCell.Value, "'")
    Wend
End Sub


