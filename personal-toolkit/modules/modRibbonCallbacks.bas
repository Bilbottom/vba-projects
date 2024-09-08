Attribute VB_Name = "modRibbonCallbacks"
Option Explicit
Option Private Module

' https://stackoverflow.com/questions/57146927/vba-populate-custom-ribbon-drop-down-list-box

Private sDropDownCurrencySelection As String
Private sDropDownDatetimeSelection As String
Private sDropDownColoursSelection  As String


'========== Helper Box

'Callback for btnHelperBox onAction
Sub RibbonHelperBoxUF(control As IRibbonControl)
    Call HelperBoxUF
End Sub


'========== General Stuff

'Callback for btnMatrixToList onAction
Sub RibbonMatrixToListUF(control As IRibbonControl)
    Call MatrixToListUF
End Sub

'Callback for btnDeleteObjects onAction
Sub RibbonDeleteObjects(control As IRibbonControl)
    Call DeleteObjects
End Sub

'Callback for btnUnhideSheets onAction
Sub RibbonUnhideAllSheets(control As IRibbonControl)
    Call UnhideAllSheets
End Sub

'Callback for btnExportModules onAction
Sub RibbonExportModules(control As IRibbonControl)
    Call ExportAllModules(ActiveWorkbook)
End Sub

'Callback for btnMakeProjectDir onAction
Sub RibbonMakeProjectDir(control As IRibbonControl)
'    MsgBox "This button is currently disabled", vbCritical
    Call MakeProjectDir
End Sub

'Callback for btnMakeTextHyperlink onAction
Sub RibbonMakeTextHyperlink(control As IRibbonControl)
    Call MakeTextHyperlink
End Sub

'Callback for btnClearTable onAction
Sub RibbonClearTable(control As IRibbonControl)
    If IsActiveCellInTable() Then
        Call EmptyTableExceptFirstRow(ActiveCell.ListObject)
    End If
End Sub


'========== Formatting

'Callback for btnFormatCode onAction
Sub RibbonApplyLogicColour(control As IRibbonControl)
    Call ApplyLogicColour
End Sub

'Callback for btnSetGroupDirection onAction
Sub RibbonSetGroupingLocation(control As IRibbonControl)
    Call SetGroupingLocation
End Sub


'========== Number Formatting

'Callback for ddNumberFormat onAction
Sub RibbonNumberFormatDD(control As IRibbonControl, id As String, index As Integer)
    Select Case control.id
        Case "ddFormatCurrency"
            Let sDropDownCurrencySelection = id
        Case "ddFormatDatetime"
            Let sDropDownDatetimeSelection = id
        Case "ddFormatColours"
            Let sDropDownColoursSelection = id
        Case Else
            Err.Raise _
                513, _
                Source:="'Personal Toolkit.xlam'!RibbonCallbacks.RibbonNumberFormatDD", _
                Description:="Unrecognised control.id"
    End Select

    Call NumberFormatDD(id)
End Sub

'Callback for btnSelectedNumberFormat onAction
Sub RibbonRepeatSelectionOnDD(control As IRibbonControl)
    Dim sDropDownSelection As String

    Select Case control.id
        Case "btnSelectedCurrencyFormat"
            Let sDropDownSelection = sDropDownCurrencySelection
        Case "btnSelectedDatetimeFormat"
            Let sDropDownSelection = sDropDownDatetimeSelection
        Case "btnSelectedColoursFormat"
            Let sDropDownSelection = sDropDownColoursSelection
        Case Else
            Err.Raise _
                513, _
                Source:="'Personal Toolkit.xlam'!RibbonCallbacks.RibbonRepeatSelectionOnDD", _
                Description:="Unrecognised control.id"
    End Select

    If sDropDownSelection = "" Then
        MsgBox "You need to select a new option."
    Else
        Call NumberFormatDD(sDropDownSelection)
    End If
End Sub


'========== Button Behaviour

'Callback for ddButtonBehaviour
Sub RibbonButtonAltDD(control As IRibbonControl, id As String, index As Integer)
    'Default is Plus-Minus One
    Call SetAltButtonBehaviour(id)
End Sub


'========== Development

'Callback for ebxEntry getText
Sub RibbonEntryGetText(control As IRibbonControl, ByRef returnedVal)
    ' Prompted when the ribbon tab is opened for the first time in the session
    Let returnedVal = "Enter..."
End Sub

'Callback for ebxEntry onChange
Sub RibbonEntryOnChange(control As IRibbonControl, text As String)
    If text <> "" Then MsgBox "You input " & text
End Sub

