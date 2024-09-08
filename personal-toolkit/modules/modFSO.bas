Attribute VB_Name = "modFSO"
Option Explicit


Private Sub ListFilesInFolderBase(ByVal sFolderpath As String, Optional ByVal sFileStart As String = "", Optional ByVal sFileEnd As String = "")
    Dim iFile As Variant
    Dim FSO   As FileSystemObject
    Set FSO = New FileSystemObject

    For Each iFile In FSO.GetFolder(sFolderpath).Files
        If sFileStart = "" And sFileEnd = "" Then
            Debug.Print iFile.Name
        Else
            If True _
                And Left(iFile.Name, Len(sFileStart)) = sFileStart _
                And Right(iFile.Name, Len(sFileEnd)) = sFileEnd _
            Then Debug.Print iFile.Name
        End If
    Next iFile
End Sub


Private Sub ListFilesInFolder()
    Call ListFilesInFolderBase( _
        sFolderpath:="", _
        sFileStart:="", _
        sFileEnd:="" _
    )
End Sub

