Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

'''
' Subroutines for the Ribbon and Quick Access Toolbar across different Outlook windows
'''

'----------------------------
' Outlook Application ribbon
'----------------------------
Public Sub RibbonClearEmailNotification()
    Call modGeneral.ClearEmailNotification
End Sub


Public Sub RibbonAddMonthlyFolders()
    Call modGeneral.AddMonthlyFolders
End Sub


Public Sub RibbonTidyMI()
    Call modTMS.TidyAnalyticsInbox
    Call modSeeds.LoopToMoveSEED
    Call modKeyfact.LoopToSaveAndMoveKeyfact
    Call modGeneral.ClearEmailNotification
End Sub


Public Sub RibbonCreateMailToMe()
    Call CreateMailToMe
End Sub


Public Sub RibbonCreateAnnualLeaveItem()
    Call modMeetingItems.CreateAnnualLeaveItem
End Sub


Public Sub RibbonRunOpsVBS()
    Call modScripting.RunOpsVBS
End Sub


'---------------------
' New MailItem ribbon
'---------------------

Public Sub EmailRibbonAddJiraTicketLink()
    '''
    ' Want the MS Word Object Library:
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


' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Public Sub SetAutofilter()
    Exit Sub ' currently not working -- can't find Application_Startup
    Call Application_Startup
End Sub
