Attribute VB_Name = "modFunctions"
Option Explicit
Option Private Module

'''
' Functions for general use.
'''


Public Function GetUserPath() As String
    '''
    ' Dynamic user path -- useful for absolute OneDrive links.
    '''
    Let GetUserPath = "C:\Users\" & Environ("username") & "\"
End Function
