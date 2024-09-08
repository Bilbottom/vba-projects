Attribute VB_Name = "modGeneral"
Option Explicit
Option Private Module

'''
' General subroutines.
'''


Private Const sMyTeam As String = "" _
    & "forename.surname@company.domain;" _
    & "my-name@company.domain;"


Public Sub CreateAnnualLeaveItem()
    '''
    ' Create an Annual Leave meeting invite.
    '''
    Dim olInbox   As Outlook.Folder
    Dim olMeeting As Outlook.AppointmentItem
    With Outlook.Application
        Set olInbox = .GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
        Set olMeeting = .CreateItem(olAppointmentItem)
    End With

    With olMeeting
        .BodyFormat = olFormatHTML
        .MeetingStatus = Outlook.olMeeting
        .Display

        .Subject = "Billiam AL"
        .RequiredAttendees = sMyTeam
        .Categories = "Annual Leave"
        .BusyStatus = olFree
        .AllDayEvent = True
        .ResponseRequested = False
        .ReminderSet = False
        .ReminderMinutesBeforeStart = 0
    End With

    Set olMeeting = Nothing
End Sub


Public Sub CreateMailToMe()
    '''
    ' Create a mail item to my personal email address.
    '''
    Dim olMail As Outlook.MailItem
    Set olMail = Outlook.Application.CreateItem(olMailItem)

    With olMail
        .BodyFormat = olFormatHTML
        .Display

        .HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>Remove all PII!"
        .To = "my.personal@email.address"
        .Subject = "From Work"
    End With

    Set olMail = Nothing
End Sub


Public Sub ClearEmailNotification()
    '''
    ' Clear the persistent email notification (a bug my Outlook client had).
    '''
    Dim oInbox As Outlook.Folder
    Set oInbox = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    With oInbox.Items(1)
        .Display
        .Close olDiscard
    End With

    Set oInbox = Nothing
End Sub


Public Sub AddMonthlyFolders()
    '''
    ' Add the usual monthly folders.
    '''
    Dim ThisMonth As String
    Dim olInbox   As Outlook.Folder

    Let ThisMonth = Format(Date, "yyyy-mm")
    Set olInbox = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    On Error GoTo ErrMonthFolderExists
        olInbox.Folders.Add ThisMonth
    On Error GoTo 0

    With olInbox.Folders(ThisMonth).Folders
        .Add "Meetings"
        .Add "Notifications"
        .Add "Tickets"
    End With
    Set olInbox = Nothing

    Exit Sub

ErrMonthFolderExists:
    Set olInbox = Nothing
    MsgBox _
        Prompt:="This month's folders have already been added.", _
        buttons:=vbInformation
End Sub
