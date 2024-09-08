Attribute VB_Name = "modTesting"
Option Explicit
Option Private Module

'''
' Module just for testing things -- you can ignore this.
'''


Public Sub PrintTopMailValues()
    Dim olMail As Variant
    Dim i      As Long

    With Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
        For i = 1 To .Items.Count
            On Error GoTo ExitEarly
                Set olMail = .Items(i)
            On Error GoTo 0

            Debug.Print """" & olMail.Sender & """", """" & olMail.SenderName & """", """" & olMail.Subject & """", """" & olMail.Attachments.Count & """"
        Next i
    End With
    Set olMail = Nothing

ExitEarly:
End Sub


Public Sub PrintTopMailValuesData()
    Dim olDataInbox As Outlook.MAPIFolder
    Dim olMail      As Object
    Dim iCounter    As Long

    With Application.GetNamespace("MAPI")
        Set olDataInbox = .GetSharedDefaultFolder(.CreateRecipient("data@company.domain"), olFolderInbox)
    End With

    With olDataInbox
        For iCounter = .Items.Count To .Items.Count - 20 Step -1
            On Error GoTo ExitEarly
                Set olMail = .Items(iCounter)
            On Error GoTo 0

            Debug.Print """" & olMail.SenderName & """", """" & olMail.Subject & """", """" & olMail.Attachments.Count & """"
        Next iCounter
    End With
    Set olMail = Nothing

ExitEarly:
End Sub


Private Function AppendToFilename(ByVal sFilename As String, ByVal sAppend As String) As String
    Dim lMatch As Long
    Let lMatch = InStrRev(sFilename, ".")

    If lMatch = 0 Then
        Let AppendToFilename = sFilename
    Else
        Let AppendToFilename = "" _
            & Left(sFilename, lMatch - 1) _
            & sAppend _
            & Right(sFilename, Len(sFilename) - lMatch + 1)
    End If
End Function


Private Sub DemoFindNext()
    '''
    ' https://stackoverflow.com/a/62959430/8213085
    ' https://docs.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences
    '''

    Dim myNameSpace As Outlook.NameSpace
    Dim tdystart As Date
    Dim tdyend As Date
    Dim myAppointments As Outlook.Items
    Dim currentAppointment As Outlook.AppointmentItem

    Set myNameSpace = Application.GetNamespace("MAPI")
    Let tdystart = VBA.Format(Now, "Short Date")
    Let tdyend = VBA.Format(Now + 1, "Short Date")
    Set myAppointments = myNameSpace.GetDefaultFolder(olFolderCalendar).Items
    myAppointments.Sort "[Start]"
    myAppointments.IncludeRecurrences = True
    Set currentAppointment = myAppointments.Find("[Start] >= """ & tdystart & """ and [Start] <= """ & tdyend & """")

    While TypeName(currentAppointment) <> "Nothing"
        Debug.Print currentAppointment.Subject, InStr(1, currentAppointment.Categories, "Planned Work")
        ' MsgBox currentAppointment.Subject
        Set currentAppointment = myAppointments.FindNext
    Wend

End Sub


Private Sub TidyMeeting(olMail As MeetingItem)
    'https://stackoverflow.com/a/9770371
    Const ToSubFolder As String = "Meetings"

    With olMail
        Debug.Assert .Parent.Name = "Inbox"
        .UnRead = False
        .Move .Parent.Folders(Format(.ReceivedTime, "yyyy-mm")).Folders(ToSubFolder)
    End With
End Sub


Private Sub CreateMeetingFromEmail()
' need this in the Outlook ribbon as well as the MailItem ribbon
' pass in a MailItem object to run this code

    Dim olInbox   As Outlook.Folder
    Dim olMeeting As Outlook.MeetingItem
    Set olInbox = Outlook.Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set olMeeting = Outlook.Application.CreateItem(olmeetingItem)

    With olMeeting
'        .BodyFormat = olFormatHTML
'        .Display
'
'        .HTMLBody = "<BODY style=font-size:11pt;font-family:Calibri>Remove all PII!"
'        .To = "personal@email.address"
'        .Subject = "From Work"
    End With

    Set olMeeting = Nothing
End Sub
