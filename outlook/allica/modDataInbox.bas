Attribute VB_Name = "modDataInbox"
Option Explicit
Option Private Module

'''
' Subroutines to run on the Data inbox.
'
' Should convert this into some classes when I have the time.
'''

Private Function sFilePath() As String
    '''
    ' Dynamic file path for this module.
    '''
    Let sFilePath = GetUserPath() & "Company\SharePoint Site Name - Documents\path"
End Function


Public Sub TidyDataInbox()
    '''
    ' Loop through the files in the Data inbox.
    '''
    Dim olDataInbox As Outlook.MAPIFolder
    Dim olItem      As Object
    Dim iCounter    As Long

    With Application.GetNamespace("MAPI")
        Set olDataInbox = .GetSharedDefaultFolder(.CreateRecipient("data@company.domain"), olFolderInbox)
    End With

    With olDataInbox
        For iCounter = 1 To .Items.Count
            On Error GoTo ExitEarly
                Set olItem = .Items(iCounter)
            On Error GoTo 0

            Call DoOutlookActions(olItem:=olItem, olInbox:=olDataInbox, lCounter:=iCounter)
        Next iCounter
    End With
ExitEarly:

    Set olDataInbox = Nothing
    Set olItem = Nothing
End Sub


Private Sub DoOutlookActions(ByVal olItem As Object, ByVal olInbox As Outlook.MAPIFolder, ByRef lCounter As Long)
    '''
    ' Middle-man between the looping and the doing.
    '
    ' Need to decrement the counter when moving items since it re-shuffles their
    ' places.
    '''
    If Left(olItem.Subject, 2) = "RE" Or Left(olItem.Subject, 2) = "FW" Then
        ' Skip -- probably something to look at
    ElseIf IsMailFromLooker(olItem) Then
        Call DoActionsForLookerReport(olItem, olInbox)
        Let lCounter = lCounter - 1
    ElseIf IsMailFromWilton(olItem) Then
        Call DoActionsForWiltonReport(olItem, olInbox)
        Let lCounter = lCounter - 1
    ElseIf IsMailFromVAS(olItem) Then
        Call DoActionsForVASReport(olItem, olInbox)
        Let lCounter = lCounter - 1
    ElseIf IsMailMetabaseInvite(olItem) Then
        Call DoActionsForMetabaseInvite(olItem, olInbox)
        Let lCounter = lCounter - 1
    End If
End Sub


'---------------------
' Mail-specific logic
'---------------------

Private Function IsMailFromLooker(ByVal olMail As Variant) As Boolean
    '''
    ' Check whether the sender is Looker.
    '''
    If TypeName(olMail) <> "MailItem" Then
        Let IsMailFromLooker = False
        Exit Function
    End If

    Let IsMailFromLooker = True _
        And olMail.SenderEmailAddress = "noreply@lookermail.com" _
        And olMail.Attachments.Count = 1
End Function

Private Sub DoActionsForLookerReport(ByVal olMail As MailItem, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Move the Looker email to its folder and mark as read.
    '''
    With olMail
        .UnRead = False
        .Move olInbox.Parent.Folders("Reports").Folders("Looker Reports")
    End With
End Sub

Private Function IsMailFromWilton(ByVal olMail As Variant) As Boolean
    '''
    ' Check whether the sender is Landbay and that the mail is the Wilton
    ' report.
    '''
    If TypeName(olMail) <> "MailItem" Then
        Let IsMailFromWilton = False
        Exit Function
    End If

    Let IsMailFromWilton = True _
        And olMail.SenderEmailAddress = "data@landbay.co.uk" _
        And olMail.Attachments.Count = 1 _
        And InStr(olMail.Subject, "Wilton") > 0
End Function

Private Sub DoActionsForWiltonReport(ByVal olMail As MailItem, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Move the Wilton email to its folder and mark as read.
    '''
    With olMail
        .UnRead = False
        .Move olInbox.Parent.Folders("Reports").Folders("Wilton Reports")
    End With
End Sub


Private Function IsMailFromVAS(ByVal olMail As Variant) As Boolean
    '''
    ' Check whether the sender is VAS.
    '''
    If TypeName(olMail) <> "MailItem" Then
        Let IsMailFromVAS = False
        Exit Function
    End If

    Let IsMailFromVAS = True _
        And olMail.SenderEmailAddress = "sender@vas-group.co.uk" _
        And olMail.Attachments.Count = 2
End Function

Private Sub DoActionsForVASReport(ByVal olMail As MailItem, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Move the VAS email to its folder and mark as read.
    '''
    With olMail
        .UnRead = False
        .Move olInbox.Parent.Folders("Reports").Folders("VAS Reports")
    End With
End Sub


Private Function IsMailMetabaseInvite(ByVal olMail As Variant) As Boolean
    '''
    ' Check whether the mail is just a Metabase invite acceptance notification.
    '''
    If TypeName(olMail) <> "MailItem" Then
        Let IsMailMetabaseInvite = False
        Exit Function
    End If

    Let IsMailMetabaseInvite = True _
        And olMail.SenderName = "Metabase SMTP" _
        And olMail.Attachments.Count = 0 _
        And Right(olMail.Subject, 30) = "accepted their Metabase invite"
End Function

Private Sub DoActionsForMetabaseInvite(ByVal olMail As MailItem, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Move the Metabase invite acceptance notification email to its folder and
    ' mark as read.
    '''
    With olMail
        .UnRead = False
        .Move olInbox.Parent.Folders("Notifications").Folders("Metabase Invites")
    End With
End Sub


'----------------
' Save for Later
'----------------

Private Sub MakeDirIfNotExist(ByVal sFullpath As String)
    '''
    ' Make the supplied directory if it doesn't already exists.
    '''
    Dim FSO As New FileSystemObject
    If Not FSO.FolderExists(sFullpath) Then FSO.CreateFolder sFullpath
End Sub
