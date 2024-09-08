VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_HelperBox
   Caption         =   "Bill's Helper Box"
   ClientHeight    =   6150
   ClientLeft      =   -30
   ClientTop       =   -105
   ClientWidth     =   6015
   OleObjectBlob   =   "UF_HelperBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UF_HelperBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



'==============================
'  Contents
'------------------------------
'  DAILY INFO PAGE.........01
'  OPEN WORKBOOKS PAGE.....02
'  ADD-INS PAGE............03
'  CONVERSIONS PAGE........04
'  SHORTCUTS PAGE..........05
'  CONTROLS PAGE...........06
'  EXCEL PAGE..............07
'  OUTLOOK PAGE............08
'  POWERPOINT PAGE.........09
'  INFO PAGE...............10
'
'==============================


'==========Userform Initialise=========='

Private Sub UserForm_Initialize()

    Me.Height = 340
    Me.Width = 316

    Call MakeMainInfoBox
    Call MakeUserformInformationBox

End Sub


'==========Page Changes=========='

Private Sub MultiPageOverall_Change()

    If MultiPageOverall.SelectedItem.index = 1 Then

        Call GetBookList

    ElseIf MultiPageOverall.SelectedItem.index = 2 Then

        Call AddInPageInitialise

    End If

End Sub


'==================================================DAILY INFO PAGE=================================================='

Private Sub MakeMainInfoBox()

    Let Me.LabelMainInfo = "" _
        & "Today: " & vbTab & Format(Date, "yyyy-mm-dd") & vbNewLine _
        & "Week:  " & vbTab & WorksheetFunction.WeekNum(Date) & vbNewLine

End Sub


'==================================================OPEN WORKBOOKS PAGE=================================================='

Private Sub ButtonRefreshBooks_Click()

    Call GetBookList

End Sub


'  https://bettersolutions.com/vba/controls/listbox.htm
Private Sub GetBookList()

    Dim wBook As Object
    Dim i     As Integer

    With Me.ListBoxWorkbooks

        For i = .ListCount - 1 To 0 Step -1
            .RemoveItem i
        Next i

'        For Each wBook In Application.Workbooks
'            .AddItem wBook.Name & ", " & tlkt_ExcelFileInUseBy(Workbooks(i).Name)
'        Next wBook

        For i = 1 To Application.Workbooks.Count
            .AddItem
            .List(i - 1, 0) = Workbooks(i).Name
            .List(i - 1, 1) = "Calculating..."
            .List(i - 1, 1) = tlkt_ExcelFileInUseBy(Workbooks(i).Name)
        Next i

    End With

End Sub


'==================================================ADD-INS PAGE=================================================='

Private Sub AddInPageInitialise()

    Dim oAddIn As Object

    For Each oAddIn In Application.COMAddIns

        If oAddIn.Description = "SAS Add-In 7.1 for Microsoft Office" Then CheckBoxCOMAddInSAS.Value = oAddIn.Connect

    Next oAddIn

End Sub

Private Sub CheckBoxCOMAddInSAS_Click()

    Dim oAddIn As Object

    For Each oAddIn In Application.COMAddIns

        If oAddIn.Description = "SAS Add-In 7.1 for Microsoft Office" Then oAddIn.Connect = CheckBoxCOMAddInSAS.Value

    Next oAddIn

End Sub


'==================================================CONVERSIONS PAGE=================================================='

'----------Check Digit----------'

Private Sub BoxNumpersona_Enter()

    With Me.BoxNumpersona
        If .Value = "Put in the numpersona..." Then .Value = ""
    End With

End Sub
Private Sub BoxCheckDigit_Enter()

    With Me.BoxCheckDigit
        If .Value = "The Check Digit is..." Then .Value = ""
    End With

End Sub


Private Sub ButtonCheckDigit_Click()

    Dim TempLong As Long
    On Error GoTo ErrorHandler
        Let TempLong = CLng(Me.BoxNumpersona.Value)
    On Error GoTo 0

    Me.BoxCheckDigit.Value = tlkt_CheckDigit(TempLong)
    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="A valid numpersona is required.", _
        Buttons:=vbCritical, _
        Title:="Numpersona Conversion"

End Sub


'----------I-Number----------'

Private Sub BoxINumber_Enter()

    With Me.BoxINumber
        If .Value = "Put in the I-Number..." Then .Value = ""
    End With

End Sub
Private Sub BoxINumberName_Enter()

    With Me.BoxINumberName
        If .Value = "This number belongs to..." Then .Value = ""
    End With

End Sub


Private Sub ButtonINumberName_Click()

    Dim TempLong As Long
    On Error GoTo ErrorHandler
        If UCase(Left(Me.BoxINumber.Value, 1)) <> "I" Then GoTo ErrorHandler
        Let TempLong = CLng(Mid(Me.BoxINumber.Value, 2))
    On Error GoTo 0

    Me.BoxINumberName.Value = tlkt_GetUserFullName(Me.BoxINumber.Value)
    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="A valid I-Number is required.", _
        Buttons:=vbCritical, _
        Title:="I-Number Conversion"

End Sub


'==================================================SHORTCUTS PAGE=================================================='







'==================================================CONTROLS PAGE=================================================='

Private Sub ButtonOpenSnippingTool_Click()

    Shell "C:\WINDOWS\explorer.exe ""C:\Windows\System32\SnippingTool.exe""", vbNormalFocus

End Sub


'==================================================EXCEL PAGE=================================================='

'----------File Name----------'

Private Sub BoxExcelName_DropButtonClick()

    Dim wBook As Workbook
    Dim i     As Integer

    With Me.BoxExcelName

        For i = .ListCount - 1 To 0 Step -1
            .RemoveItem i
        Next i

        For Each wBook In Application.Workbooks
                .AddItem wBook.Name
        Next wBook

    End With

End Sub


Private Sub ButtonUnhideAllSheets_Click()

    Dim wBook  As Workbook
    Dim wSheet As Worksheet

    With Me.BoxExcelName

        On Error GoTo ErrorHandler
            Set wBook = Workbooks(.Value)
        On Error GoTo 0

        For Each wSheet In wBook.Worksheets
            wSheet.Visible = xlSheetVisible
        Next wSheet

    End With

    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="A valid Excel file name is required.", _
        Buttons:=vbCritical, _
        Title:="Excel Error"

End Sub


Private Sub CommandButtonConvertMatrix_Click()
    Call MatrixToListUF
End Sub


'==================================================OUTLOOK PAGE=================================================='

Private Sub ButtonUnsecMastEmail_Click()

    Dim OutlookApp As Outlook.Application
    Dim OutlookMail As Outlook.MailItem

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(olMailItem)

    With OutlookMail

        .BodyFormat = olFormatHTML
        .Display
        If Not CheckBoxUnsecuredMaster Then .HTMLBody = ""
        .To = "" _
            & "person-1@email.domain; " _
            & "person-2@email.domain; " _
            & ""
        If CheckBoxUnsecuredMaster Then
            .Subject = "Unsecured Master - "
        Else
            .Subject = "Unsecured Master - Code has been run"
        End If

    End With

End Sub


Private Sub ButtonACQTeamEmail_Click()

    Dim OutlookApp As Outlook.Application
    Dim OutlookMail As Outlook.MailItem

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(olMailItem)

    With OutlookMail

        .BodyFormat = olFormatHTML
        .Display
        .To = "" _
            & "person-1@email.domain; " _
            & "person-2@email.domain; " _
            & ""

    End With

End Sub


'==================================================POWERPOINT PAGE=================================================='

'----------File Name----------'

Private Sub BoxPowerPointName_DropButtonClick()

    Dim pptApp  As PowerPoint.Application
    Dim pptPres As Object
    Dim i       As Integer

    Set pptApp = CreateObject("PowerPoint.Application")

    With Me.BoxPowerPointName

        For i = .ListCount - 1 To 0 Step -1
            .RemoveItem i
        Next i

        For Each pptPres In pptApp.Presentations
                .AddItem pptPres.Name
        Next pptPres

    End With

End Sub


'----------Default Controls----------'

Private Sub ButtonBreakLinks_Click()

    Dim pptApp  As PowerPoint.Application
    Dim TempPpt As Object

    Set pptApp = CreateObject("PowerPoint.Application")

    With Me.BoxPowerPointName

        On Error GoTo ErrorHandler
            Set TempPpt = pptApp.Presentations(.Value)
        On Error GoTo 0
        Call BreakPowerPointLinksBase(.Value)
        MsgBox "The links for " & .Value & " have been broken."
        Exit Sub

    End With

ErrorHandler:
    MsgBox _
        Prompt:="A valid PowerPoint file name is required.", _
        Buttons:=vbCritical, _
        Title:="PowerPoint Error"

End Sub


Private Sub ButtonPublishPack_Click()

    Dim pptApp  As PowerPoint.Application
    Dim TempPpt As Object

    Set pptApp = CreateObject("PowerPoint.Application")

    With Me.BoxPowerPointName

        On Error GoTo ErrorHandler
            Set TempPpt = pptApp.Presentations(.Value)
        On Error GoTo 0
        Call PublishPowerPointBase(pptName:=.Value, BreakLinks:=True)
        MsgBox .Value & " has been published with broken links and the date suffix."
        Exit Sub

    End With

ErrorHandler:
    MsgBox _
        Prompt:="A valid PowerPoint file name is required.", _
        Buttons:=vbCritical, _
        Title:="PowerPoint Error"

End Sub


'----------Publishing Controls----------'

Private Sub ButtonPublishPackAs_Click()

    Dim pptApp  As PowerPoint.Application
    Dim TempPpt As Object
    Dim sNewPpt As String
    Dim oNewPpt As Object

    Set pptApp = CreateObject("PowerPoint.Application")

    With Me.BoxPowerPointName

        On Error GoTo ErrorHandler
            Set TempPpt = pptApp.Presentations(.Value)
        On Error GoTo 0

        sNewPpt = Application.GetSaveAsFilename( _
            InitialFileName:=.Value, _
            FileFilter:="PowerPoint Presentation (*.pptx), *.pptx", _
            Title:="Save " & .Value & " as..." _
        )

    End With

    TempPpt.SaveCopyAs FileName:=sNewPpt

    Set oNewPpt = pptApp.Presentations.Open(sNewPpt)

    With oNewPpt

        If CheckBoxBreakLinks Then

            Call BreakPowerPointLinksBase(.Name)
            .Save

        End If

        If Not CheckBoxKeepCopyOpen Then .Close

    End With

    Exit Sub

ErrorHandler:
    MsgBox _
        Prompt:="A valid PowerPoint file name is required.", _
        Buttons:=vbCritical, _
        Title:="PowerPoint Error"

End Sub


'==================================================INFO PAGE=================================================='


Private Sub MakeUserformInformationBox()

    Let Me.LabelUserformInformation = "" _
        & Space(4) & "This userform is designed to assist in business-as-usual type activities." _
        & vbNewLine _
        & vbNewLine _
        & "Version 0.1" & Space(5) & "2019-11-25"

End Sub
