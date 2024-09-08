Attribute VB_Name = "modOnKeys"
Option Explicit
Option Private Module


'==========ASSIGN KEYS WHEN BOOK IS OPENED=========='
Public Sub AssignKeys()
    '  https://stackoverflow.com/questions/9818047/assign-keyboard-shortcut-to-run-procedure
    '  https://docs.microsoft.com/en-us/office/vba/api/excel.application.onkey

    '  Shift  +
    '  Ctrl   ^
    '  Alt    %

    With Application
        .OnKey "%{UP}", "PlusOneToSelection"
        .OnKey "%{DOWN}", "MinusOneToSelection"

        .OnKey "+{F4}", "MakeReferenceTypeWithUF"
        .OnKey "^+{C}", "ColourSelection"

        .OnKey "^+{v}", "PasteAsValues"
    End With
End Sub


'==========PLUS/MINUS ONE TO SELECTION=========='

'  https://stackoverflow.com/questions/28587154/how-do-i-add-1-to-all-cells-in-a-user-selected-range
Public Sub AssignPlusMinusToAlt()
    With Application
        .OnKey "%{UP}", "PlusOneToSelection"
        .OnKey "%{DOWN}", "MinusOneToSelection"
        .OnKey "%{LEFT}", ""
        .OnKey "%{RIGHT}", ""
    End With
End Sub


Private Sub PlusToSelectionBase(iAddedValue As Double)
    Dim CalcState As Integer
    Dim Cell      As Range

    With Application
        Let CalcState = .Calculation
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With

    For Each Cell In Selection
        With Cell
            If (.Value = "True" Or .Value = "False") Then
                .Value = Not .Value
            Else
                .Value = .Value + iAddedValue
            End If
        End With
    Next

    With Application
        .ScreenUpdating = True
        .Calculation = CalcState
    End With
End Sub


Private Sub PlusOneToSelection()
    Call PlusToSelectionBase(1)
End Sub
Private Sub MinusOneToSelection()
    Call PlusToSelectionBase(-1)
End Sub


'==========ALT TO MOVE SELECTION=========='

Public Sub AssignAltKeyForSelection()
    With Application
        .OnKey "%{UP}", "MoveSelectionUp"
        .OnKey "%{DOWN}", "MoveSelectionDown"
        .OnKey "%{LEFT}", "MoveSelectionLeft"
        .OnKey "%{RIGHT}", "MoveSelectionRight"
    End With
End Sub


Private Sub MoveSelectionBase(lRows As Long, lCols As Long)
    On Error Resume Next
        Selection.Offset(lRows, lCols).Select
    On Error GoTo 0
End Sub

Private Sub MoveSelectionUp()
    MoveSelectionBase -1, 0
End Sub
Private Sub MoveSelectionDown()
    MoveSelectionBase 1, 0
End Sub
Private Sub MoveSelectionLeft()
    MoveSelectionBase 0, -1
End Sub
Private Sub MoveSelectionRight()
    MoveSelectionBase 0, 1
End Sub


'==========COLOUR SELECTION=========='

Private Sub ColourSelection()
    Dim LightGreen  As Long
    Dim LightOrange As Long
    Dim Pink        As Long

    LightGreen = 11854022
    LightOrange = 11389944
    Pink = 13551615

    With Selection
        If .Cells(1, 1).Interior.ColorIndex = xlNone Then
            .Interior.Color = LightGreen
        ElseIf .Cells(1, 1).Interior.Color = LightGreen Then
            .Interior.Color = LightOrange
        ElseIf .Cells(1, 1).Interior.Color = LightOrange Then
            .Interior.Color = Pink
        ElseIf .Cells(1, 1).Interior.Color = Pink Then
            .Interior.Color = xlNone
        Else
            .Interior.Color = xlNone
        End If
    End With
End Sub


'==========PASTE VALUES=========='

Private Sub PasteAsValues()
    On Error GoTo ErrorHandler
        Selection.PasteSpecial Paste:=xlPasteValues
        Exit Sub
ErrorHandler:
    Selection.PasteSpecial Paste:=xlPasteAll
End Sub
