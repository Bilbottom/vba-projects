Attribute VB_Name = "modWebControl"
Option Explicit
Option Private Module


Public Sub OpenMyWorkJira()
    Call ThisWorkbook.FollowHyperlink(Address:="https://your-cloud-domain.atlassian.net/jira/your-work")
End Sub
