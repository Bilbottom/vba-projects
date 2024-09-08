VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_MakeReferenceType 
   Caption         =   "Edit Formula References"
   ClientHeight    =   2280
   ClientLeft      =   30
   ClientTop       =   135
   ClientWidth     =   4440
   OleObjectBlob   =   "UF_MakeReferenceType.frx":0000
End
Attribute VB_Name = "UF_MakeReferenceType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'  https://bettersolutions.com/vba/userforms/positioning.htm

Private Sub UserForm_Initialize()
    Me.Height = 144
    Me.Width = 240

    Me.Left = GetSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Left", 0)
    Me.Top = GetSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Top", 0)
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   Call SaveSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Left", Me.Left)
   Call SaveSetting("Userform Positioning", ThisWorkbook.FullName & "-" & Me.Name, "Top", Me.Top)
End Sub


'==XlReferenceType Enumeration=='
'  1  xlAbsolute
'  2  xlAbsRowRelColumn
'  4  xlRelative
'  3  xlRelRowAbsColumn
    
Private Sub ButtonOK_Click()
    Dim Cell    As Range
    Dim RefType As Integer

    With Me
        Select Case True
            Case .AbsoluteButton.Value
                RefType = xlAbsolute
            Case .AbsRowRelColButton.Value
                RefType = xlAbsRowRelColumn
            Case .RelativeButton.Value
                RefType = xlRelative
            Case .RelRowAbsColButton.Value
                RefType = xlRelRowAbsColumn
        End Select
    End With

    Me.Hide
    If RefType = 0 Then Exit Sub

    Application.ScreenUpdating = False

        For Each Cell In Selection
            If Cell.HasFormula Then
                Cell.Formula = Application.ConvertFormula( _
                        Formula:=Cell.Formula, _
                        FromReferenceStyle:=xlA1, _
                        ToReferenceStyle:=xlA1, _
                        ToAbsolute:=RefType _
                    )
            End If
        Next Cell

    Application.ScreenUpdating = True
End Sub
