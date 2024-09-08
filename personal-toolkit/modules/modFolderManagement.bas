Attribute VB_Name = "modFolderManagement"
Option Explicit
Option Private Module


Private Const sDefDir  As String = "C:\Users\bilbottom\Company\SharePoint Site Name - Documents\"
Private Const sJiraURL As String = "https://your-cloud-domain.atlassian.net/browse/"


Private Sub TestMake()
'    Call MakeProjectDirectory(sProjectName:="This is a test")
    Call MakeProjectDirectory(sProjectName:="DATA-0000 Test Project")
End Sub


Public Sub MakeProjectDirectory(ByVal sProjectName As String, Optional ByVal sAltDir As String)
    ' Take sProjectName without a numeric prefix
    Dim FSO        As New FileSystemObject
    Dim sNewDir    As String
    Dim vSubDir    As Variant
    Dim sTicketRef As String

    Let sTicketRef = GetJiraRef(sProjectName)  ' Should pop this in as an arg or create a class with it

    ' Check project name
    If IsNameTaken(Trim(sProjectName)) _
        Then Err.Raise Number:=513, Description:="This project name is already taken."

    ' Make dir name
    If sAltDir = "" Then
        Let sNewDir = sDefDir & Trim(sProjectName) & "\"
    Else
        Err.Raise Number:=513, Description:="This subroutine isn't configured for alternative directories yet."
    End If

    ' Create directory
    Call FSO.CreateFolder(Path:=sNewDir)

    ' Create shortcut
    Call CreateCustomShortcut( _
        sProjectName:=Trim(sProjectName), _
        sFullDir:=sNewDir _
    )

    ' Create files
    Call FSO.CreateTextFile(sNewDir & "\" & sTicketRef & ".py")
    Call FSO.CreateTextFile(sNewDir & "\" & sTicketRef & ".sql")

    ' Clean up memory
    Set FSO = Nothing
End Sub


Private Function IsNameTaken(ByVal sName As String) As Boolean
    Dim vStr As Variant
    Dim sRef As String

    Let sRef = GetJiraRef(sName)
    Let IsNameTaken = False

    If sRef = "" Then
        Err.Raise Number:=513, Description:="The project name should be prefixed with a Jira ticket ref."
    Else
        For Each vStr In Split(GetCurrentNames(), ",")
            Let IsNameTaken = (sRef = GetJiraRef(vStr))
            If IsNameTaken Then Exit Function
        Next vStr
    End If
End Function


Private Function GetCurrentNames() As String
    Dim FSO  As New FileSystemObject
    Dim vFol As Variant
    Dim sStr As String

    Let sStr = ""
    For Each vFol In FSO.GetFolder(sDefDir).SubFolders
        Let sStr = sStr & "," & vFol.Name
    Next vFol

    Let GetCurrentNames = Mid(sStr, 2)
End Function


Public Sub CreateCustomShortcut(ByVal sProjectName As String, ByVal sFullDir As String)
    Dim sTicketRef As String
    Dim oShortcut  As Object

    Let sTicketRef = GetJiraRef(sProjectName)

    If sTicketRef <> "" Then
        Set oShortcut = CreateObject("WScript.Shell").CreateShortCut( _
            sFullDir & "\" & sTicketRef & ".url" _
        )

        With oShortcut
            .TargetPath = sJiraURL & sTicketRef
            .Save
        End With
    End If
End Sub


Private Function GetJiraRef(ByVal sString As String) As String
    Dim reRegex    As New RegExp
    Dim sTicketRef As String

    With reRegex
        .Global = True
        .MultiLine = False
        .IgnoreCase = False
        .Pattern = "(^[A-Za-z]+-[0-9]+).*"
    End With

    If reRegex.Test(sString) Then
        Let GetJiraRef = reRegex.Replace(sString, "$1")
    Else
        Let GetJiraRef = ""
    End If
End Function


' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------


Private Function Greatest(ByVal x As Variant, ByVal y As Variant) As Variant
    Let Greatest = IIf(x > y, x, y)
End Function


Private Function Least(ByVal x As Variant, ByVal y As Variant) As Variant
    Let Least = IIf(x < y, x, y)
End Function
