Attribute VB_Name = "modPython"
Option Explicit
Option Private Module

'''
' Subroutine to run a python script. Assumes using the same executable.
'
' Can add up to 9 parameters.
'''

Private Const pyExe As String = "C:\Users\bilbottom\AppData\Local\Programs\Python\Python310\python.exe"


Public Sub RunPython( _
    ByVal sFile As String, _
    Optional ByVal vArg1 As Variant = Null, _
    Optional ByVal vArg2 As Variant = Null, _
    Optional ByVal vArg3 As Variant = Null, _
    Optional ByVal vArg4 As Variant = Null, _
    Optional ByVal vArg5 As Variant = Null, _
    Optional ByVal vArg6 As Variant = Null, _
    Optional ByVal vArg7 As Variant = Null, _
    Optional ByVal vArg8 As Variant = Null, _
    Optional ByVal vArg9 As Variant = Null, _
    Optional ByVal bVerbose As Boolean = False _
)
    '''
    ' Run a Python script with up to 9 parameters.
    '''
    Dim sRunString As String
    Let sRunString = "" _
        & pyExe _
        & " """ & sFile & """" _
        & IIf(IsNull(vArg1), "", " """ & vArg1 & """") _
        & IIf(IsNull(vArg2), "", " """ & vArg2 & """") _
        & IIf(IsNull(vArg3), "", " """ & vArg3 & """") _
        & IIf(IsNull(vArg4), "", " """ & vArg4 & """") _
        & IIf(IsNull(vArg5), "", " """ & vArg5 & """") _
        & IIf(IsNull(vArg6), "", " """ & vArg6 & """") _
        & IIf(IsNull(vArg7), "", " """ & vArg7 & """") _
        & IIf(IsNull(vArg8), "", " """ & vArg8 & """") _
        & IIf(IsNull(vArg9), "", " """ & vArg9 & """")

    If bVerbose Then Debug.Print sRunString
    Shell sRunString

'    Dim sReturn  As Integer
'    Dim wshShell As Object
'    Set wshShell = VBA.CreateObject("WScript.Shell")
'    Let sReturn = wshShell.Run(sRunString, 1, True)
'    Debug.Print sReturn
End Sub


Public Sub PythonCopyFile(ByVal sFrom As String, ByVal sTo As String)
    '''
    ' Useful for copying things to OneDrive/SharePoint.
    '''
    Call RunPython( _
        sFile:=GetUserPath() & "\Company\SharePoint Site Name - Documents\Scripts\Push to OneDrive\main.py", _
        vArg1:=sFrom, _
        vArg2:=sTo, _
        bVerbose:=False _
    )
End Sub
