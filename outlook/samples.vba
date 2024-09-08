'''
' Running code when an email arrives
' ----------------------------------
'
' There is some freedom in how you define your objects to run code when emails
' land, but you must:
'
'   1. declare a global object of type `Outlook.Items` using `WithEvents`
'   2. give the object a value in a subroutine (a good choice is in `Application_Startup`)
'   3. write a subroutine whose name is your object’s name suffixed with `_ItemAdd`
'''

'''
' Personal inbox
'''
Option Explicit

Private WithEvents olItems As Outlook.Items


' This will run when Outlook starts
Private Sub Application_Startup()
    Set olItems = _
        Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items
End Sub


' This will run whenever an email lands in the inbox
Private Sub olItems_ItemAdd(ByVal oItem As Object)
    Dim olMail  As Outlook.MailItem
    Dim olInbox As Outlook.Folder
    Set olInbox = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    ' MailItem is an actual email, rather than a meeting invite or update
    If TypeName(oItem) = "MailItem" Then
        Set olMail = oItem
        Call DoSomething(olMail)
    End If

    Set olMail = Nothing
    Set olInbox = Nothing
End Sub


'''
' Shared inbox
'''
Option Explicit

Private WithEvents olSharedItems As Outlook.Items

Const SharedEmail As String = "shared.inbox@email.domain"


Private Sub Application_Startup()
    Dim olNs    As Outlook.NameSpace
    Dim olRecip As Outlook.Recipient

    Set olNs = Outlook.Application.GetNamespace("MAPI")
    Set olRecip = olNs.CreateRecipient(SharedEmail)

    Set olSharedItems = olNs.GetSharedDefaultFolder(olRecip, olFolderInbox).Items
End Sub


Private Sub olSharedItems_ItemAdd(ByVal oItem As Object)
    Dim olNs          As Outlook.NameSpace
    Dim olRecip       As Outlook.Recipient
    Dim olSharedInbox As Outlook.MAPIFolder
    Dim olMail        As Outlook.MailItem

    Set olNs = Application.GetNamespace("MAPI")
    Set olRecip = olNs.CreateRecipient(SharedEmail)
    Set olSharedInbox = olNs.GetSharedDefaultFolder(olRecip, olFolderInbox)

    If TypeName(oItem) = "MailItem" Then
        Set olMail = oItem
        Call DoSomethingForSharedInbox(olMail)
    End If

    Set olMail = Nothing
    Set olNs = Nothing
    Set olRecip = Nothing
    Set olSharedInbox = Nothing
End Sub


'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'''
' Sending an email (the common example)
' -------------------------------------
'
' You can send emails from any of the Office applications that have VBA, but if
' you're using any that aren’t Outlook then you’ll want to enable the *Microsoft
' Outlook 16.0 Object Library* reference.
'''

Public Sub SendEmail()
    Dim olMail As Outlook.MailItem
    Set olMail = Outlook.Application.CreateItem(olMailItem)

    With olMail
        .BodyFormat = olFormatHTML
        .Display

        .HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>" _
            & "Hi," & "<br><br>" _
            & "This is a test email." & "<br><br>" _
            & .HTMLBody

        .To = "joe.bloggs@gmail.com;john.smith@hotmail.com"
        .CC = "jane.smith@hotmail.com"

        .Subject = "Test Email"
        .Attachments.Add "C:\Documents\Example-File.xlsx"

        .Send
    End With

    Set olMail = Nothing
End Sub
