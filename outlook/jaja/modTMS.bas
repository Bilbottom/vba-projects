Attribute VB_Name = "modTMS"
Option Explicit
Option Private Module

'''
' Subroutines to run on the Analytics inbox.
'''


Private Function sFilePath() As String
    '''
    ' Dynamic file path for this module.
    '''
    Let sFilePath = GetUserPath() & "Company\SharePoint Site Name - Documents\Reporting\Operations\Input\"
End Function


Public Sub TidyAnalyticsInbox()
    '''
    ' Loop through the files in the Analytics Inbox and deal with the JAX and
    ' JAY emails.
    '''
    Dim olSharedInbox As Outlook.MAPIFolder
    Dim olNs          As Outlook.NameSpace
    Dim olRecip       As Outlook.Recipient
    Dim olItem        As Object
    Dim iCounter      As Long

    Set olNs = Application.GetNamespace("MAPI")
    Set olRecip = olNs.CreateRecipient("team@company.domain")
    Set olSharedInbox = olNs.GetSharedDefaultFolder(olRecip, olFolderInbox)

    With olSharedInbox
        For iCounter = 1 To .Items.Count
            On Error GoTo ExitEarly
                Set olItem = .Items(iCounter)
            On Error GoTo 0

            If IsMailFromTMS(olItem) Then
                Call SaveAndMoveTMS(olItem, olSharedInbox)
                Let iCounter = iCounter - 1
            End If
        Next iCounter
    End With
ExitEarly:

    Set olNs = Nothing
    Set olRecip = Nothing
    Set olSharedInbox = Nothing
    Set olItem = Nothing
End Sub


Public Function IsMailFromTMS(ByVal olMail As Variant) As Boolean
    '''
    ' Check whether the sender is TMS, there is an attachment, and the subject
    ' starts with JAX or JAY.
    '''
    If TypeName(olMail) <> "MailItem" Then
        Let IsMailFromTMS = False
        Exit Function
    End If

    Let IsMailFromTMS = True _
        And (False _
            Or olMail.SenderEmailAddress = "mi@tsys.com" _
            Or olMail.SenderEmailAddress = "DoNotReply@tsys.com" _
        ) _
        And olMail.Attachments.Count = 1 _
        And (False _
            Or Left(Trim(olMail.Subject), 3) = "JAX" _
            Or Left(Trim(olMail.Subject), 3) = "JAY" _
        )
End Function


Public Sub SaveAndMoveTMS(ByVal olMail As MailItem, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Analytics inbox subroutine to save and move the TMS emails/reports.
    '''
    Dim sReport   As String
    Dim sFileDate As String
    Let sReport = Left(Trim(olMail.Subject), 6)
    Let sFileDate = GetFileDate(Trim(olMail.Subject))

'    Debug.Print sReport, sFileDate
    Call SaveTMSReport(olMail, sReport, sFileDate)
    Call MoveTMSReport(olMail, sReport, olInbox)
End Sub


Private Sub SaveTMSReport(ByVal olMail As MailItem, ByVal sReport As String, ByVal sFileDate As String)
    '''
    ' Inner subroutine for `SaveAndMoveTMS` to save the report to the OneDrive.
    '''
    With olMail
        ' JAX reports
        If Left(sReport, 3) = "JAX" Then
            With .Attachments.Item(1)
                ' Save the version for the workflows
                .SaveAsFile sFilePath() & .FileName

                ' Save the dated version for history
                .SaveAsFile sFilePath() & "\Reports\" & sReport & "\" & sFileDate & " - " & .FileName
            End With

        ' JAY reports
        ElseIf Left(sReport, 3) = "JAY" Then
            With .Attachments.Item(1)
                ' Save to our location
                .SaveAsFile sFilePath() & "\Reports\" & sReport & "\" & .FileName
                ' Save to Connor's location
                Call ConorsSaveHourlyFileFromTMS(.FileName)
            End With

            ' Run the Operations VBScript
            If InStr(.Attachments.Item(1).FileName, "_03") > 0 Then
                Call modScripting.RunOpsVBS
            End If

        ' Something else
        Else
            Err.Raise Number:=513, Description:="Error in report name at modTMS.SaveTMSReport"
        End If
    End With
End Sub


Private Sub MoveTMSReport(ByVal olMail As MailItem, ByVal sReport As String, ByVal olInbox As Outlook.MAPIFolder)
    '''
    ' Inner subroutine for `SaveAndMoveTMS` to move the email to its folder.
    '''
    With olMail
        .UnRead = False
        .Move olInbox.Parent.Folders("TMS Reports").Folders(sReport)
    End With
End Sub


Private Function GetFileDate(ByVal sSubject As String) As String
    '''
    ' Extract the date from the filename where the date is in the dd/mm/yyyy
    ' format.
    '''
    Dim regEx As New RegExp
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = ".*(\d{2})/(\d{2})/(\d{4}).*"
    End With

    If regEx.Test(sSubject) Then
        Let GetFileDate = regEx.Replace(sSubject, "$3-$2-$1")
    Else
        Let GetFileDate = ""
    End If
End Function


' ---------- Conor Subs --------------------------------------------------------

Private Sub ConorsSaveHourlyFileFromTMS(ByVal sFilename As String)
    '''
    ' Conor has been saving the hourly files to SharePoint manually.
    '
    ' This is to automate what he does.
    '''
    Dim sFromPath As String
    Dim sConorTMS As String
    Let sFromPath = sFilePath() & "\Reports\JAY222\"
    Let sConorTMS = "" _
        & GetUserPath() & "\Company\SharePoint Site Name - Documents\" _
        & "Reporting\MI\"

    Dim FSO        As New FileSystemObject
    Dim sDate      As String
    Dim lDate      As Long
    Dim sYearPart  As String
    Dim sMonthPart As String
    Dim sDayPart   As String
    Dim sFullpath  As String

    Let sDate = Mid(sFilename, 32, 8)
    Let lDate = DateSerial(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
    Let sYearPart = Format(lDate, "yyyy")
    Let sMonthPart = Format(lDate, "mmmm")
    Let sDayPart = Format(lDate, "ddmmyyyy")

    Call MakeDirIfNotExist(sConorTMS & sYearPart & "\")
    Call MakeDirIfNotExist(sConorTMS & sYearPart & "\" & sMonthPart & "\")
    Call MakeDirIfNotExist(sConorTMS & sYearPart & "\" & sMonthPart & "\" & sDayPart & "\")

    Let sFullpath = sConorTMS & sYearPart & "\" & sMonthPart & "\" & sDayPart & "\"

'    Debug.Print sFromPath & sFilename & vbCrLf & sFullpath & sFilename
    Call PythonCopyFile(sFrom:=sFromPath & sFilename, sTo:=sFullpath & sFilename)
End Sub


Private Sub MakeDirIfNotExist(ByVal sFullpath As String)
    '''
    ' Make the supplied directory if it doesn't already exists.
    '''
    Dim FSO As New FileSystemObject
    If Not FSO.FolderExists(sFullpath) Then FSO.CreateFolder sFullpath
End Sub
