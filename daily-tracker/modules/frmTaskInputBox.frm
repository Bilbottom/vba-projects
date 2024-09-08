VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTaskInputBox
   Caption         =   "Interval Tracker at hh:mm (int)"
   ClientHeight    =   2280
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "frmTaskInputBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTaskInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'''
' This `frmTaskInputBox` form writes to the Tracker sheet to record current work.
' This is also pushed into a Slack channel and, if it's a Jira ticket, adds
' details to the worklog.
'''


'====================================================================================================
' Fields
' ------
'
' Declaring a variable without a Property Let/Get makes it a field rather than a property which only
' enforces a type.
'
' Good practice is for in-memory properties to use private fields prefixed with the letter `p`.
'====================================================================================================
'Private pInterval     As Long
'Private pScheduleTime As String
Private pCounter As Long  ' Just for the cbxProject_Change event
Private pClicks  As Long  ' ust for the btnOK_Click event


'====================================================================================================
' Properties
' ----------
'
' This is better than using fields as this given better control over the reading and writing of
' properties (namely with the Let/Set and Get methods).
'====================================================================================================

' -----------
' In-Memory
' -----------

' Interval Minutes
'Public Property Let Interval(ByVal v As Long)
'    Let pInterval = v
'End Property
'Public Property Get Interval() As Long
'    Let Interval = pInterval
'End Property
Public Property Get Interval() As Long
    Let Interval = wsVariables.[lCurrentInterval].Value
End Property

' Schedule Time
'Public Property Let ScheduleTime(ByVal v As String)
'    pScheduleTime = v
'End Property
'Public Property Get ScheduleTime() As String
'    Let ScheduleTime = pScheduleTime
'End Property
Public Property Get ScheduleTime() As String
    Let ScheduleTime = wsVariables.[sCurrentTime].Value
End Property


' ---------
' wsOptions
' ---------

' Use Calendar Defaults
Public Property Get UseCalendarDefaults() As Boolean
    Let UseCalendarDefaults = wsOptions.[bUseCalendarDefaults].Value
End Property

' Post to Slack
Public Property Get PostToSlack() As Boolean
    Let PostToSlack = wsOptions.[bPostToSlack].Value
End Property

' Check Active Jira Sprint
Public Property Get CheckJiraSprint() As Boolean
    Let CheckJiraSprint = wsOptions.[bCheckJiraSprint].Value
End Property

' Post to Jira
Public Property Get PostToJira() As Boolean
    Let PostToJira = wsOptions.[bPostToJira].Value
End Property

' Project Drop-Down List
Public Property Get ProjectDropDownList() As String
    '''
    ' Drop-down box for Task based on default values and what we've seen so far,
    ' delimited by commas.
    '''
    Dim sDefault  As String
    Dim sRecent   As String
    Dim sUpcoming As String

    ' Get the default values
    With wsCategories
        Let sDefault = WorksheetFunction.TextJoin( _
            ",", _
            True, _
            .Range(.[rDefault], .[rDefault].End(xlDown)) _
        )
    End With

    ' Get the recent values
    With wsVariables
        Let sRecent = WorksheetFunction.TextJoin( _
            ",", _
            True, _
            .Range(.[rRecentProjects], .[rRecentProjects].End(xlDown)) _
        )
    End With

    ' Get the upcoming values
    With wsVariables
        Let sUpcoming = WorksheetFunction.TextJoin( _
            ",", _
            True, _
            .Range(.[rUpcomingProjects], .[rUpcomingProjects].End(xlDown)) _
        )
    End With

    ' Return the unique list of the union of the values
    Let ProjectDropDownList = UniqueStr( _
        sStr:=sDefault & "," & SortStr(sStr:=sRecent & "," & sUpcoming, sDelim:=","), _
        sDelim:="," _
    )
End Property

' Latest Detail for cbxProject
Public Property Get LatestProjectDetail() As String
    '''
    ' The latest Detail for the `cbxProject`.
    '''
    Dim rLookup As Range
    With wsVariables
        Set rLookup = .Range(.[rRecentProjects], .[rRecentProjects].End(xlDown))
    End With

    Let LatestProjectDetail = WorksheetFunction.XLookup(Me.cbxProject, rLookup, rLookup.Offset(0, 1), "")
    Set rLookup = Nothing
End Property


'====================================================================================================
' Functions
' ---------
'
' Functions to support instances of `clsTracker`.
'====================================================================================================

Private Function GetDefaultTaskAndDetail(ByVal sTime As String) As String
    '''
    ' Get default values for the input box, delimited by a comma.
    '''
    Dim sAppointment As String
    Let sAppointment = GetCurrentAppointment(sTime:=sTime)

    ' Outlook meetings, then last input
    If Me.UseCalendarDefaults And sAppointment <> "" Then
        If IsAppointmentException(sAppointment) Then
            Let GetDefaultTaskAndDetail = ConvertAppointmentException(sAppointment)
'        ElseIf InStr(1, sAppointment.Categories, "Catch Ups") = 0 Then
'            Let GetDefaultTaskAndDetail = "Catch Ups," & Trim(sAppointment)
'        ElseIf InStr(1, sAppointment.Categories, "Personal Development") = 0 Then
'            Let GetDefaultTaskAndDetail = "Personal Development," & Trim(sAppointment)
        Else
            Let GetDefaultTaskAndDetail = "Meetings," & Trim(sAppointment)
        End If
    Else
        Let GetDefaultTaskAndDetail = "" _
            & wsData.Cells(2 ^ 20, 5).End(xlUp).Value & "," _
            & wsData.Cells(2 ^ 20, 6).End(xlUp).Value
    End If
End Function


Private Function GetCurrentAppointment(ByVal sTime As String) As String
    '''
    ' Get the current meeting from Outlook, if one exists.
    '
    ' Might experience performance drop-off when calendar gets bigger.
    ' Only returns something if exactly one appointment is found.
    ' Ignores all-day appointments and Planned Work category.
    '
    ' https://stackoverflow.com/q/1927799/8213085
    '''
    Dim olCalItems As Outlook.Items
    Dim olFiltered As Outlook.Items
    Dim olItem     As Outlook.AppointmentItem
    Dim sTimeFilt  As String
    Dim sFilter    As String
    Dim i          As Long

    Let sTimeFilt = Format(Now, "yyyy-mm-dd") & " " & sTime
    Let sFilter = "[Start] <= '" & sTimeFilt & "' AND [End] > '" & sTimeFilt & "'"

    Set olCalItems = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    olCalItems.IncludeRecurrences = True
    olCalItems.Sort "[Start]"
    Set olFiltered = olCalItems.Restrict(sFilter)

    For Each olItem In olFiltered
        If i > 0 Then GoTo ExitEarly

        If (Not olItem.AllDayEvent) And (InStr(1, olItem.Categories, "Planned Work") = 0) Then
            Let GetCurrentAppointment = olItem.Subject
            Let i = i + 1
        End If
    Next olItem

    Exit Function

ExitEarly:
    Let GetCurrentAppointment = ""
End Function


Private Function IsAppointmentException(ByVal sSubject As String) As Boolean
    '''
    ' Exceptions list -- bad idea to do it this way.
    '''
    Const sExceptions As String = "" _
        & "Daily Jira Call," _
        & "Jira Scrum," _
        & "Jira Scrum / Team Meeting," _
        & "Planning," _
        & "Nik & Bill Catch Up," _
        & "Nik & Bill Developmental Catch Up," _
        & "1-2-1 Bill & Juliana," _
        & "Manage Jira Tickets," _
        & "Personal Development on Alteryx," _
        & "Personal Development on Tableau"

    Let IsAppointmentException = IsInArray(sSubject, Split(sExceptions, ","))
End Function
Private Function ConvertAppointmentException(ByVal sSubject As String) As String
    '''
    ' Exceptions list -- bad idea to do it this way
    '''
    Select Case sSubject
        Case "Daily Jira Call"
            Let ConvertAppointmentException = "Catch Ups,Daily Jira Call"
        Case "Jira Scrum"
            Let ConvertAppointmentException = "Catch Ups,Jira Scrum"
        Case "Jira Scrum / Team Meeting"
            Let ConvertAppointmentException = "Catch Ups,Jira Scrum"
        Case "Planning"
            Let ConvertAppointmentException = "Catch Ups,Planning"
        Case "Nik & Bill Catch Up"
            Let ConvertAppointmentException = "Catch Ups,Nik & Bill"
        Case "Nik & Bill Developmental Catch Up"
            Let ConvertAppointmentException = "Catch Ups,Nik & Bill"
        Case "1-2-1 Bill & Juliana"
            Let ConvertAppointmentException = "Catch Ups,Juliana & Bill"
        Case "Manage Jira Tickets"
            Let ConvertAppointmentException = "Housekeeping,Manage Jira Tickets"
        Case "Personal Development on Alteryx"
            Let ConvertAppointmentException = "Personal Development,Alteryx"
        Case "Personal Development on Tableau"
            Let ConvertAppointmentException = "Personal Development,Tableau"
        Case Else
            Err.Raise Number:=513, Description:="Unrecognised Appointment Exception"
    End Select
End Function

'====================================================================================================
' Methods
' -------
'
' Subroutines to call on instances of `frmTaskInputBox`.
'====================================================================================================

' ------------
' Init and Del
' ------------

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call SaveSetting(AppName:="Userform Positioning", Section:=ThisWorkbook.FullName & "-" & Me.Name, key:="Left", Setting:=Me.Left)
    Call SaveSetting(AppName:="Userform Positioning", Section:=ThisWorkbook.FullName & "-" & Me.Name, key:="Top", Setting:=Me.Top)
End Sub


Private Sub UserForm_Initialize()
    '''
    ' Create the pop-up box and add the defaults to the boxes.
    '''
'    On Error Resume Next
'        Me.Left = GetSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Left", 0)
'        Me.Top = GetSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Top", 0)
'    On Error GoTo 0

    Dim sDefaults As String
    Let pCounter = 0

    With Me
        Let sDefaults = GetDefaultTaskAndDetail(sTime:=.ScheduleTime)

        ' Position and size
        .Left = 0: .Top = 0: .Height = 144: .Width = 240

        ' Caption
        .Caption = "Interval Tracker at " & Me.ScheduleTime & " (" & Me.Interval & ")"

        ' Update the drop-down
        Call ManageUpcomingTicketsInDropDown

        ' Input boxes
        .tbxDetail.Value = Mid(sDefaults, 1 + InStr(sDefaults, ","))
        With Me.cbxProject
            .List = Split(Me.ProjectDropDownList, ",")
            .Value = Split(sDefaults, ",")(0)

            ' Select the text within the Project box upon loading
            .SelStart = 0
            .SelLength = Len(.text)
        End With
    End With
End Sub


Private Sub UserForm_Terminate()
    Unload Me
End Sub


' -----------
' Form Events
' -----------

Private Sub btnCancel_Click()
    Call UserForm_Terminate
End Sub


Private Sub btnOK_Click()
    '''
    ' Clicking OK prompts a number of actions.
    '''
    Dim lRow As Long
    Let pClicks = pClicks + 1
    If pClicks <> 1 Then
        Debug.Print "Too many clicks"
        Exit Sub
    End If

    ' Write to tracker
    With wsData
        Let lRow = .Cells(.Rows.Count, 3).End(xlUp).Row + 1

        .Cells(lRow, 3).Value = Date
        With .Cells(lRow, 4)
            .Value = TimeValue(Me.ScheduleTime)
            .NumberFormat = "hh:mm"
        End With
        .Cells(lRow, 5).Value = Trim(Me.cbxProject)
        .Cells(lRow, 6).Value = Trim(Me.tbxDetail)
        .Cells(lRow, 7).Value = Me.Interval
    End With

    ' External actions
    If Me.PostToSlack Then Call PostMessageToSlack
    If Me.PostToJira Then Call PostLogToJira

    ' Export data on the hour
    If Right(Me.ScheduleTime, 2) = "00" Then Call Me.WriteTrackerDataToCSV

'    ' Update the drop-down
'    Call ManageUpcomingTicketsInDropDown

    ' Close the form
    Let pClicks = 0
    Call UserForm_Terminate
End Sub


Private Sub cbxProject_Change()
    '''
    ' Set the Detail to the last Detail for the selected Project.
    '''
    If pCounter <> 0 Then Me.tbxDetail = Me.LatestProjectDetail
    Let pCounter = pCounter + 1
End Sub


' ----------------
' Specific Actions
' ----------------

Private Sub PostMessageToSlack()
    '''
    ' Execute the Python script which posts a message to a dedicated Slack
    ' channel.
    '''
    Const sSlackPy As String = _
        "C:\Users\billiam\Repos\Slack-API\main.py"

    If Trim(Me.tbxDetail) = "" Then
        Call modPython.RunPython(sFile:=sSlackPy, vArg1:=Trim(Me.cbxProject))
    Else
        Call modPython.RunPython(sFile:=sSlackPy, vArg1:="*" & Trim(Me.cbxProject) & "*: " & Trim(Me.tbxDetail))
    End If
End Sub


Private Sub PostLogToJira()
    '''
    ' Execute the API call which logs time against the corresponding ticket in
    ' Jira.
    '''
    Const sJiraKey = "^([A-Z][\w\d]{1,9}-\d+).*"
    Dim jcJira As clsJiraConnector
    Set jcJira = New clsJiraConnector

    If RegexpMatch(Trim(Me.cbxProject), sJiraKey) Then
        jcJira.AddWorklog _
            sIssueKey:=RegexpReplace(Trim(Me.cbxProject), sJiraKey, "$1"), _
            sLogDetail:=Trim(Me.tbxDetail), _
            sLogDatetime:=Format(Date, "yyyy-mm-dd") & "T" & Me.ScheduleTime & ":00.000+0000", _
            lLogDuration:=Me.Interval * 60
    End If

    Set jcJira = Nothing
End Sub


Public Sub WriteTrackerDataToCSV()
    '''
    ' Write the tracker data to a CSV file.
    '''
    Const sOutFile As String = _
        "C:\Users\billiam\Downloads\daily-tracker-data.csv"

    Call WriteToCSV( _
        wsData:=wsExport, _
        sOutFile:=sOutFile _
    )
End Sub


Private Function GetTicketsInSprint() As String
    '''
    ' Get the tickets (Key & Summary) in the current sprint for the current user.
    '''
    Dim sTickets  As String
    Dim oResponse As Scripting.Dictionary
    Dim vItem     As Variant
    Dim jcJira    As clsJiraConnector

    Let sTickets = ""
    Set jcJira = New clsJiraConnector
    Set oResponse = ParseJson(jcJira.SearchForIssuesUsingJQL(sParams:="fields=summary,duedate,assignee&jql=sprint IN openSprints() AND assignee IN (currentUser())"))

    On Error Resume Next
        For vItem = 1 To oResponse("issues").Count
            sTickets = sTickets & oResponse("issues")(vItem)("key") & " " & oResponse("issues")(vItem)("fields")("summary") & ";"
        Next vItem
    On Error GoTo 0

    Let GetTicketsInSprint = Left(sTickets, Len(sTickets) - 1)

    Set jcJira = Nothing
    Set oResponse = Nothing
End Function


Private Sub ManageUpcomingTicketsInDropDown()
    '''
    '
    '''
    Const sJiraKey = "^([A-Z][\w\d]{1,9}-\d+).*"
    Dim sTicket As Variant
    Dim sKey    As String
    Dim iCell   As Range
    Dim bMatch  As Boolean

    ' Clear the 'pending' tickets
    With wsVariables.Range("rUpcomingProjects")
        If .Offset(1, 0) = "" Then
            ' Skip this part
        ElseIf .Offset(2, 0) = "" Then
            .Offset(1, 0).Clear
        Else
            .Parent.Range(.Offset(1, 0), .Offset(1, 0).End(xlDown)).Clear
        End If
    End With

    ' Add the 'pending' tickets
    For Each sTicket In Split(GetTicketsInSprint, ";")
        Let sKey = RegexpReplace(sTicket, sJiraKey, "$1")
'        Debug.Print sKey, sTicket

        ' If the rRecentProjects list is empty, don't loop through it
        Let bMatch = (wsVariables.Range("rRecentProjects").Offset(1, 0) = "")

        With wsVariables
            Let bMatch = False

            ' Only add it to the rUpcomingProjects list if it isn't in the rRecentProjects list
            For Each iCell In .Range(.Range("rRecentProjects"), .Range("rRecentProjects").End(xlDown))
                If Not bMatch Then bMatch = (sKey = RegexpReplace(iCell.Value, sJiraKey, "$1"))
            Next iCell

            If Not bMatch Then .Range("rUpcomingProjects").Offset(-1, 0).End(xlDown).Offset(1, 0).Value = sTicket
        End With
    Next sTicket
End Sub


' ====================================================================================================

Private Sub ListUpcomingAppointments()
    '''
    ' Not used in the form, just kept for reference.
    '
    ' https://stackoverflow.com/q/1927799/8213085
    '''
    Dim olCalItems As Outlook.Items
    Dim olFiltered As Outlook.Items
    Dim olItem     As Outlook.AppointmentItem

    Const sFilter As String = "[Start] >= '14/12/2020 00:00' AND [End] <= '18/12/2020 23:59'"

    Set olCalItems = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    olCalItems.IncludeRecurrences = True
    olCalItems.Sort "[Start]"
    Set olFiltered = olCalItems.Restrict(sFilter)

    For Each olItem In olFiltered
        Debug.Print olItem.Start, olItem.End, olItem.Duration, olItem.Subject
    Next olItem

    Set olCalItems = Nothing
    Set olFiltered = Nothing
    Set olItem = Nothing
End Sub


Private Function ThisWorkbook_GetTicketsInSprint() As String
    Dim sTickets  As String
    Dim oResponse As Scripting.Dictionary
    Dim vItem     As Variant
    Dim jcJira    As clsJiraConnector

    Let sTickets = ""
    Set jcJira = New clsJiraConnector
    Set oResponse = ParseJson(jcJira.SearchForIssuesUsingJQL(sParams:="fields=summary,duedate,assignee&jql=sprint IN openSprints() AND assignee IN (currentUser())"))

    On Error Resume Next
        For vItem = 1 To oResponse("issues").Count
            sTickets = sTickets & oResponse("issues")(vItem)("key") & " " & oResponse("issues")(vItem)("fields")("summary") & ";"
        Next vItem
    On Error GoTo 0

    Let GetTicketsInSprint = Left(sTickets, Len(sTickets) - 1)

    Set jcJira = Nothing
    Set oResponse = Nothing
End Function


Private Sub ThisWorkbook_AddTicketsToDropDown()
    Const sJiraKey = "^([A-Z][\w\d]{1,9}-\d+).*"
    Dim sTicket As Variant
    Dim sKey    As String
    Dim iCell   As Range
    Dim bMatch  As Boolean

    For Each sTicket In Split(GetTicketsInSprint, ";")
        Let sKey = RegexpReplace(sTicket, sJiraKey, "$1")
        Debug.Print sKey, sTicket

        With wsVariables
            Let bMatch = False
            For Each iCell In .Range(.Range("rRecentProjects"), .Range("rRecentProjects").End(xlDown))
                If Not bMatch Then bMatch = (sKey = RegexpReplace(iCell.Value, sJiraKey, "$1"))
            Next iCell
            If Not bMatch Then .Range("rRecentProjects").End(xlDown).Offset(1, 0).Value = sTicket
        End With
    Next sTicket
End Sub
