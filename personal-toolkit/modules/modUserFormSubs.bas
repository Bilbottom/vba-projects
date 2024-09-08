Attribute VB_Name = "modUserFormSubs"
Option Explicit
Option Private Module


'====================CHECK IF USERFORM IS LOADED===================='

'  https://stackoverflow.com/a/50563889/8213085
Public Function IsLoaded(FormName As String) As Boolean
    Dim Frm As Object

    For Each Frm In VBA.UserForms
        If Frm.Name = FormName Then
            Let IsLoaded = True
            Exit Function
        End If
    Next Frm

    Let IsLoaded = False
End Function


'====================LOAD HELPER BOX===================='

'  https://www.excelforum.com/excel-programming-vba-macros/505725-embed-folder-treeview-or-windows-explorer-on-userform.html
Public Sub HelperBoxUF()
    Dim HelperBoxUF As UF_HelperBox

    If IsLoaded("UF_HelperBox") Then Exit Sub

    Set HelperBoxUF = New UF_HelperBox

    HelperBoxUF.Show vbModeless
End Sub


'==========CHANGE REFERENCE TYPE OF ADDRESSES=========='
Private Sub MakeReferenceTypeWithUF()
    '  https://answers.microsoft.com/en-us/msoffice/forum/all/apply-absolute-reference-to-multiple-cells-at-once/5e943e84-b809-4f3c-959b-1aeddfa7e46b
    '  http://www.excelforum.com/excel-general/372383-making-multiple-cells-absolute-at-once.html
    '  https://powerspreadsheets.com/create-excel-userforms/

    Dim MyUserForm    As UF_MakeReferenceType
    Set MyUserForm = New UF_MakeReferenceType

    MyUserForm.Show vbModal

    Unload MyUserForm
    Set MyUserForm = Nothing
End Sub


'==========MATRIX TO LIST CONVERTER=========='
Public Sub MatrixToListUF()
    Dim MatrixToListUF As UF_MatrixConverter

    If IsLoaded("UF_MatrixConverter") Then Exit Sub

    Set MatrixToListUF = New UF_MatrixConverter

    ' RefEdit not supported in modeless
'    MatrixToListUF.Show vbModeless

    MatrixToListUF.Show vbModal
    Unload MatrixToListUF
    Set MatrixToListUF = Nothing
End Sub
