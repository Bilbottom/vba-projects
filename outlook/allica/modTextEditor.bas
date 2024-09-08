Attribute VB_Name = "modTextEditor"
Option Explicit
Option Private Module


Public Sub SetTextUppercase()
    '''
    ' https://docs.microsoft.com/en-us/office/vba/api/word.font
    '''
    With Outlook.Application.ActiveExplorer.Selection.Item(1).GetInspector.WordEditor
        .Application.Selection.Range = UCase(.Application.Selection.Range)
    End With
End Sub


Public Sub SetTextLowercase()
    '''
    ' https://docs.microsoft.com/en-us/office/vba/api/word.font
    '''
    With Outlook.Application.ActiveExplorer.Selection.Item(1).GetInspector.WordEditor
        .Application.Selection.Range = LCase(.Application.Selection.Range)
    End With
End Sub

