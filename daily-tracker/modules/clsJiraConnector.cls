VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJiraConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''
' This class uses the following references:
'     - Microsoft XML, v 3.0
'     - Microsoft VBScript Regular Expressions 5.5
'
' This `clsJiraConnector` class uses the Jira API to automate and streamline some tasks.
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
Private pBaseUrl  As String
Private pUsername As String
Private pPassword As String


'====================================================================================================
' Properties
' ----------
'
' This is better than using fields as this given better control over the reading and writing of
' properties (namely with the Let/Set and Get methods).
'====================================================================================================
' Base URL
Public Property Get BaseURL() As String
    Let BaseURL = pBaseUrl
End Property

' Basic Authentication
Public Property Get AuthBasic() As String
    Let AuthBasic = "Basic " & EncodeBase64(pUsername & ":" & pPassword)
End Property


'====================================================================================================
' Methods
' -------
'
' Subroutines to call on instances of `frmTaskInputBox`.
'
' https://www.appsloveworld.com/solved-how-to-call-rest-api-from-excel-macros-vba-and-parse-json/
'====================================================================================================

' ------------
' Init and Del
' ------------
Private Sub Class_Initialize()
    Let pBaseUrl = "https://your-cloud-domain.atlassian.net/rest/api/3/"
    Let pUsername = "some-username"
    Let pPassword = "some-password"
End Sub

Private Sub Class_Terminate()

End Sub


' ------------
' API Requests
' ------------
Private Function SendHTTP( _
    ByVal sMethod As String, _
    ByVal sURL As String, _
    Optional ByVal sPayload As String = "" _
) As String
    '''
    ' Helper function for sending HTTP requests.
    '''
    Dim oXMLHTTP  As MSXML2.ServerXMLHTTP
    Set oXMLHTTP = New MSXML2.ServerXMLHTTP

    With oXMLHTTP
        .Open sMethod, sURL, True
        .SetRequestHeader "Content-Type", "application/json"
        .SetRequestHeader "Accept", "application/json"
        .SetRequestHeader "Authorization", Me.AuthBasic
        .Send sPayload

        While .ReadyState <> 4
            DoEvents
        Wend
        Let SendHTTP = .ResponseText
    End With

    Set oXMLHTTP = Nothing
End Function


Public Function GetIssue(ByVal sIssueKey As String) As String
    '''
    ' https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issues/#api-rest-api-3-issue-issueidorkey-get
    '''
    Dim sEndpoint As String
    Let sEndpoint = Me.BaseURL & "issue/" & sIssueKey

    Let GetIssue = SendHTTP(sMethod:="GET", sURL:=sEndpoint)
End Function


Public Sub AddWorklog( _
    ByVal sIssueKey As String, _
    ByVal sLogDetail As String, _
    ByVal sLogDatetime As String, _
    ByVal lLogDuration As Long _
)
    '''
    ' https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issue-worklogs/#api-rest-api-3-issue-issueidorkey-worklog-post
    '''
    Dim sEndpoint As String
    Dim sBody     As String

    Let sEndpoint = Me.BaseURL & "issue/" & sIssueKey & "/worklog?notifyUsers=false"
    Let sBody = "" _
        & "{" _
        & "    ""timeSpentSeconds"": " & lLogDuration & "," _
        & "    ""comment"": {" _
        & "        ""type"": ""doc""," _
        & "        ""version"": 1," _
        & "        ""content"": [" _
        & "            {" _
        & "                ""type"": ""paragraph""," _
        & "                ""content"": [" _
        & "                    {" _
        & "                        ""text"": """ & sLogDetail & """," _
        & "                        ""type"": ""text""" _
        & "                    }" _
        & "                ]" _
        & "            }" _
        & "        ]" _
        & "    }," _
        & "    ""started"": """ & sLogDatetime & """" _
        & "}"

    SendHTTP sMethod:="POST", sURL:=sEndpoint, sPayload:=sBody
'    Debug.Print SendHTTP(sMethod:="POST", sURL:=sEndpoint, sPayload:=sBody)
End Sub


Public Function SearchForIssuesUsingJQL(ByVal sParams As String) As String
    '''
    ' https://developer.atlassian.com/cloud/jira/platform/rest/v3/api-group-issue-search/#api-rest-api-3-search-get
    '''
    Dim sEndpoint As String
    Let sEndpoint = Me.BaseURL & "search?" & sParams

    Let SearchForIssuesUsingJQL = SendHTTP(sMethod:="Get", sURL:=sEndpoint)
End Function


Public Sub CreateIssue(ByVal sProjectId As String, ByVal sSummary As String, ByVal sDescription As String)
    Debug.Print sProjectId, sSummary, sDescription
End Sub
