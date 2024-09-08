Attribute VB_Name = "modScripting"
Option Explicit
Option Private Module

'''
' Subroutines to run any script by "double-clicking" on it.
'''

Public Sub RunOpsVBS()
    '''
    ' Run the main operations MI script.
    '''
    Call RunScript(GetUserPath() & "Company\SharePoint Site Name - Documents\Parsers\run-ops.vbs")
End Sub


Private Sub RunScript(ByVal sFilePath As String)
    '''
    ' Run a script as if you had double-clicked the file in Windows Explorer.
    '''
    Shell "C:\WINDOWS\explorer.exe """ & sFilePath & """"
'    Debug.Print "C:\WINDOWS\explorer.exe """ & sFilePath & """"
End Sub


'Public Sub RunBatFile()
'    ' Causes malicious file warning and shuts down Outlook
'    Shell "C:\Users\bilbottom\Desktop\aws-rtt.bat"
'End Sub
