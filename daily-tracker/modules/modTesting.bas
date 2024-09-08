Attribute VB_Name = "modTesting"
Option Explicit
Option Private Module

'''
' Space to mess around.
'''

Private Sub TestingDict()
    '''
    ' https://snb-vba.eu/VBA_Dictionary_en.html
    '''
    Dim vKey  As Variant
    Dim dDict As Scripting.Dictionary
    Set dDict = New Scripting.Dictionary

    With dDict
        .Add "Key1", "Item1"
        .Add "Key2", "Item2"
        .Add "Key3", "Item3"
        .Add "Key4", "Item4"
    End With

    For Each vKey In dDict.Keys()
        Debug.Print dDict.Item(vKey)
    Next vKey

'    Debug.Print dDict.Count, dDict.Item("Key1"), Join(Split("1,2,3,4", ","), "")
End Function


Private Sub JsonTest()
    Dim oResponse As Scripting.Dictionary
    Dim jcJira    As New clsJiraConnector
    Dim vKey      As Variant
    Dim vValue    As Variant

    Set oResponse = ParseJson(jcJira.GetIssue("DDB-1620"))

    On Error Resume Next
        For Each vKey In oResponse.Keys
            Let vValue = oResponse(vKey)
            Debug.Print VarType(vKey), vKey, VarType(vValue), vValue
        Next vKey
    On Error GoTo 0

'    On Error Resume Next
'        For Each vKey In oResponse("fields")
'            Debug.Print vKey, oResponse("fields")(vKey)
'        Next vKey
'    On Error GoTo 0

    Set oResponse = Nothing
    Set jcJira = Nothing
End Sub


' ----------


Private Sub ListUpcomingAppointments()
    Dim olCalItems As Outlook.Items
    Dim olFiltered As Outlook.Items
    Dim olItem     As Outlook.AppointmentItem

'    Const sFilter As String = "[Start] >= '2022-02-03 00:00' AND [End] <= '2022-02-03 23:59'"

    Dim sTimeFilt As String
    Dim sFilter   As String
    Let sTimeFilt = Format(Now, "yyyy-mm-dd") & " " & Format(Now, "hh:mm")
    Let sFilter = "[Start] <= '" & sTimeFilt & "' AND [End] > '" & sTimeFilt & "'"
    Debug.Print sTimeFilt
    Debug.Print sFilter

    Set olCalItems = CreateObject("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
    olCalItems.IncludeRecurrences = True
    olCalItems.Sort "[Start]"
    Set olFiltered = olCalItems.Restrict(sFilter)

    For Each olItem In olFiltered
        Debug.Print olItem.Start, olItem.End, olItem.Duration, olItem.Subject
    Next olItem
End Sub
