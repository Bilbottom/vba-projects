Attribute VB_Name = "modModuleManipulator"
Option Explicit
Option Private Module


Private Sub TestModuleExport()
    Call ExportAllModules(ThisWorkbook)
End Sub


Public Sub ExportAllModules(ByVal wbBook As Workbook)
    If Not wbBook.HasVBProject Then
        Call MsgBox(Prompt:="There is nothing to export.", Buttons:=vbCritical)
    Else
        Call ExportAllModulesBase(wbBook)
    End If
End Sub


Private Sub ExportAllModulesBase(ByVal wbBook As Workbook)
    Dim sDir      As String
    Dim VBComp    As VBComponent
    Dim bExport   As Boolean
    Dim sFileName As String

    Let sDir = MakeDir(wbBook)

    For Each VBComp In wbBook.VBProject.VBComponents
        Let bExport = True
        Let sFileName = VBComp.Name

        ' Concatenate the correct filename for export
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                Let sFileName = sFileName & ".cls"
            Case vbext_ct_MSForm
                Let sFileName = sFileName & ".frm"
            Case vbext_ct_StdModule
                Let sFileName = sFileName & ".bas"
            Case vbext_ct_Document
                ' This is a worksheet or workbook object, don't try to export
                Let bExport = False
        End Select

        If bExport Then VBComp.Export ConcatPath(sPath:=sDir, sSubPath:=sFileName)
    Next VBComp
End Sub


Private Function MakeDir(ByVal wbBook As Workbook) As String
    Dim FSO       As New FileSystemObject
    Dim sFolder   As String
    Dim sFullPath As String

    If IsInOneDrive(wbBook.Path) Then
        Let sFolder = DropFileExtension(ConvertOneDrivePath(wbBook.Name)) & "_modules"
        Let sFullPath = ConcatPath(sPath:=ConvertOneDrivePath(wbBook.Path), sSubPath:=sFolder)
    Else
        Let sFolder = DropFileExtension(wbBook.Name) & "_modules"
        Let sFullPath = ConcatPath(sPath:=wbBook.Path, sSubPath:=sFolder)
    End If

    If Not FSO.FolderExists(sFullPath) Then
          FSO.CreateFolder sFullPath
    End If

    Let MakeDir = sFullPath
End Function


' This exists within the toolkit already
Private Function GetDelimiter(ByVal sPath As String, Optional ByVal bReturnFullPath As Boolean = False) As String
    Dim sDelimiter As String

    If InStr(sPath, "/") > 0 Then 'URL
        Let sDelimiter = "/"
    ElseIf InStr(sPath, "\") > 0 Then 'Directory
        Let sDelimiter = "\"
    Else
        Let sDelimiter = ""
    End If

    If bReturnFullPath Then
        Let GetDelimiter = sPath & sDelimiter
    Else
        Let GetDelimiter = sDelimiter
    End If
End Function

Private Function PathNeedsDelimiter(ByVal sPath) As Boolean
    Let PathNeedsDelimiter = Not (False _
        Or Right(sPath, 1) = "/" _
        Or Right(sPath, 1) = "\" _
    )
End Function

Private Function ConcatPath(ByVal sPath As String, ByVal sSubPath As String) As String
    If PathNeedsDelimiter(sPath) Then
        Let ConcatPath = GetDelimiter(sPath, True) & sSubPath
    Else
        Let ConcatPath = sPath & sSubPath
    End If
End Function

Private Function DropFileExtension(ByVal sPath As String) As String
    Let DropFileExtension = _
        Left( _
            String:=sPath, _
            Length:=-1 + InStrRev(StringCheck:=sPath, StringMatch:=".") _
        )
End Function

Private Function IsInOneDrive(ByVal sPath As String) As Boolean
    Let IsInOneDrive = (Trim(Left(sPath, 68)) = "https://fmbholding-my.sharepoint.com/personal/bilbottom_domain/")
End Function

Private Function ConvertOneDrivePath(ByVal sPath As String) As String
    Dim sTemp As String
    Let sTemp = sPath
    Let sTemp = Replace(Expression:=sTemp, Find:="https://fmbholding-my.sharepoint.com/personal/bilbottom_domain/Documents/", Replace:="C:/Users/bilbottom/Company/SharePoint Site Name - Documents/")
    Let sTemp = Replace(Expression:=sTemp, Find:="%20", Replace:=" ")

    Let ConvertOneDrivePath = sTemp
End Function
