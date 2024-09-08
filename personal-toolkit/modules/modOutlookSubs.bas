Attribute VB_Name = "modOutlookSubs"
Option Explicit
Option Private Module


'====================ADD MONTHLY FOLDERS===================='

'  This should be automated

Private Sub AddMonthlyFolders()
    Dim ThisMonth   As String
    Dim OutlookApp  As Outlook.Application
    Dim OutlookMAPI As Outlook.Namespace
    Dim MyInbox     As Outlook.Folder

    Let ThisMonth = Format(Date, "yyyy-mm")
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMAPI = OutlookApp.GetNamespace("MAPI")
    Set MyInbox = OutlookMAPI.GetDefaultFolder(olFolderInbox)

    On Error GoTo ErrMonthFolderExists
        MyInbox.Folders.Add (ThisMonth)
    On Error GoTo 0

    With MyInbox.Folders(ThisMonth).Folders
        .Add "Meetings"
        .Add "Notifications"
        .Add "Tickets"
    End With

    Exit Sub

ErrMonthFolderExists:
    MsgBox "This month's folder has already been added."
End Sub


'====================CONSTRUCT NEW EMAIL===================='

Private Sub WriteNewEmail()
    Dim OutlookApp As Outlook.Application
    Dim OutlookMail As Outlook.MailItem

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(olMailItem)

    With OutlookMail
        .BodyFormat = olFormatHTML
        .Display

        .HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>" & "Text here" & "<br>" & "<br>" & "More text" & .HTMLBody

'        .To = "someone@email.domain"
'        .CC = "another-person@email.domain"
'        .Subject = "From VBA"

'        .Attachments.Add = ThisWorkbook.FullName

'        .VotingOptions = "Yes;No;Maybe"

'        .Send
    End With
End Sub
