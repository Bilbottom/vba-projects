Attribute VB_Name = "modPowerPointSubs"
Option Explicit
Option Private Module


'  Set the link to the Object Library:
'  Tools -> References -> Microsoft PowerPoint x.xx Object Library


'==========PRINT OPEN POWERPOINTS=========='

Private Sub GetOpenPowerPoints()

    Dim pptApp  As PowerPoint.Application
    Dim pptPres As Object

    Set pptApp = CreateObject("PowerPoint.Application")

    For Each pptPres In pptApp.Presentations
        Debug.Print pptPres.Name
    Next pptPres

End Sub


'====================GET LINKS===================='

Private Sub GetPowerPointLinks()

    Application.ScreenUpdating = False

    Dim pptApp          As PowerPoint.Application
    Dim pptPresentation As Object
    Dim pptSlide        As Object
    Dim pptShape        As Object
    Dim pptName         As String

    Let pptName = ".pptx"
    Set pptApp = CreateObject("PowerPoint.Application"): pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations(pptName)

    For Each pptSlide In pptPresentation.Slides
        For Each pptShape In pptSlide.Shapes

            If pptShape.Type = msoLinkedPicture Or pptShape.Type = msoLinkedOLEObject Then

                Debug.Print pptShape.LinkFormat.SourceFullName

            End If

        Next pptShape
    Next pptSlide

    Application.ScreenUpdating = True

End Sub


'====================EDIT LINKS===================='

Private Sub EditPowerPointLinks()

    Application.ScreenUpdating = False

    Dim pptApp          As PowerPoint.Application
    Dim pptPresentation As Object
    Dim pptSlide        As Object
    Dim pptShape        As Object
    Dim pptName         As String

    Let pptName = ".pptx"
    Set pptApp = CreateObject("PowerPoint.Application"): pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations(pptName)

    For Each pptSlide In pptPresentation.Slides
        For Each pptShape In pptSlide.Shapes

            If pptShape.Type = msoLinkedPicture Or pptShape.Type = msoLinkedOLEObject Then

                With pptShape.LinkFormat

                    ' Change graph links
                    .SourceFullName = Replace( _
                        Expression:=.SourceFullName, _
                        Find:="Text to find", _
                        Replace:="Text to replace" _
                    )

                End With

            End If

        Next pptShape
    Next pptSlide

    Application.ScreenUpdating = True

End Sub


'====================BREAK LINKS===================='

Public Sub BreakPowerPointLinksBase(pptName As String)

    Application.ScreenUpdating = False

    Dim pptApp          As PowerPoint.Application
    Dim pptPresentation As Object
    Dim pptSlide        As Object
    Dim pptShape        As Object

    Set pptApp = CreateObject("PowerPoint.Application")

    With pptApp
        .Visible = True
        Set pptPresentation = .Presentations(pptName)
    End With

    For Each pptSlide In pptPresentation.Slides
        For Each pptShape In pptSlide.Shapes

            If (pptShape.Type = msoLinkedPicture Or pptShape.Type = msoLinkedOLEObject) Then pptShape.LinkFormat.BreakLink

        Next pptShape
    Next pptSlide

    Application.ScreenUpdating = True

End Sub


'====================MAKE POWERPOINT COPY===================='

Public Sub PublishPowerPointBase(pptName As String, Optional BreakLinks As Boolean = True)

    Application.ScreenUpdating = False

    Dim pptApp          As PowerPoint.Application
    Dim pptPresentation As Object
    Dim RepMonth        As String
    Dim NewFileName     As String
    Dim pptNew          As Object

    Let RepMonth = Format(DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1)), "yyyymm")
    Set pptApp = CreateObject("PowerPoint.Application")

    With pptApp
        .Visible = True
        Set pptPresentation = pptApp.Presentations(pptName)
    End With

    With pptPresentation
        Let NewFileName = Replace(.FullName, ".pptx", "") & " - " & RepMonth & ".pptx"
        .SaveCopyAs (NewFileName)
    End With

    Set pptNew = pptApp.Presentations.Open(NewFileName)

    Call BreakPowerPointLinksBase(pptNew.Name)
    pptNew.Save

    Application.ScreenUpdating = True

End Sub


'====================TOGGLE AUTO UPDATE PROPERTY===================='

'  Not working yet

Private Sub TogglePowerPointLinksBase(pptName As String, AutoUpdateState As PpUpdateOption)

    Application.ScreenUpdating = False

    Dim pptApp          As PowerPoint.Application
    Dim pptPresentation As Object
    Dim pptSlide        As Object 'PowerPoint.Slide
    Dim pptShape        As Object 'PowerPoint.Shape

    Set pptApp = CreateObject("PowerPoint.Application"): pptApp.Visible = True
    Set pptPresentation = pptApp.Presentations(pptName)

    For Each pptSlide In pptPresentation.Slides

        For Each pptShape In pptSlide.Shapes

            If (pptShape.Type = msoLinkedPicture Or pptShape.Type = msoLinkedOLEObject) Then pptShape.LinkFormat.AutoUpdate = AutoUpdateState

        Next pptShape

    Next pptSlide

    Application.ScreenUpdating = True

End Sub

Private Sub TogglePowerPointLinks()

    Call TogglePowerPointLinksBase(".pptx", ppUpdateOptionManual)

End Sub
