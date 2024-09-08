Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

'''
' Subroutines for the Ribbon and Quick Access Toolbar across different Outlook
' windows.
'''

'----------------------------
' Outlook Application ribbon
'----------------------------
Public Sub RibbonTidyDataInbox()
    Call modDataInbox.TidyDataInbox
End Sub


Public Sub RibbonClearEmailNotification()
    Call modGeneral.ClearEmailNotification
End Sub


Public Sub RibbonAddMonthlyFolders()
    Call modGeneral.AddMonthlyFolders
End Sub


Public Sub RibbonCreateMailToMe()
    Call CreateMailToMe
End Sub


Public Sub RibbonCreateAnnualLeaveItem()
    Call CreateAnnualLeaveItem
End Sub


'---------------------
' New MailItem ribbon
'---------------------

Public Sub EmailRibbonAddJiraTicketLink()
    '''
    ' Want the MS Word Object Library:
    '
    '     http://www.vbaexpress.com/forum/showthread.php?52985-VBA-get-selected-text-from-Outlook-email-body-and-use-in-Excel
    '''
    Const sJiraTicketAddress As String = "https://your-cloud-domain.atlassian.net/browse/"

    With Outlook.Application.ActiveExplorer.Selection.Item(1).GetInspector.WordEditor
        .Hyperlinks.Add _
            Anchor:=.Application.Selection.Range, _
            Address:=sJiraTicketAddress & Trim(.Application.Selection.Range.Text)
    End With
End Sub


Public Sub EmailRibbonChangeFontToCode()
    '''
    ' https://docs.microsoft.com/en-us/office/vba/api/word.font
    '''
    With Outlook.Application.ActiveExplorer.Selection.Item(1).GetInspector.WordEditor
        .Application.Selection.Range.Font.Name = "Consolas"
    End With
End Sub


Public Sub EmailRibbonSetTextUppercase()
    modTextEditor.SetTextUppercase
End Sub


Public Sub EmailRibbonSetTextLowercase()
    modTextEditor.SetTextLowercase
End Sub
