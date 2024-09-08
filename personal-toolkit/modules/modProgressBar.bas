Attribute VB_Name = "modProgressBar"
Option Explicit
Option Private Module

Dim ProgressStatus As New frmProgressStatusDialogue
Const ProgressStatusInitialising As String = "Initialising progress bar, please wait..."
Const ProgressStatusProcessing   As String = "Updating calculations, please wait..."
Const ProgressStatusStop         As String = "Stop"
Const ProgressStatusOf           As String = " of "
Const ProgressStatusDash         As String = " - "


Public Sub ProgressBarExample()
    Dim TitleOfProgressBar As String
    Dim NumberOfProgressLoops As Double
    Dim CurrentProgressLoop As Double

    Let TitleOfProgressBar = "Progress Bar Title Example"
    Let NumberOfProgressLoops = 10000

    Call DisplayProgressBar(ProgressCounterMax:=NumberOfProgressLoops, ProgressStatusTitle:=TitleOfProgressBar)
    Call StartProgressBar(ProgressCounter:=0, ProgressCounterMax:=NumberOfProgressLoops)

    For CurrentProgressLoop = 1 To NumberOfProgressLoops
        Call UpdateProgressBar(ProgressCounter:=CurrentProgressLoop, ProgressCounterMax:=NumberOfProgressLoops)
    Next CurrentProgressLoop

    Call CloseProgressBar
End Sub


Public Sub DisplayProgressBar(ByRef ProgressCounterMax As Double, ByRef ProgressStatusTitle As String)
    Dim ProgressCounter As Double
    Let ProgressCounter = 0

    With ProgressStatus
        .Configure _
            Title:=ProgressStatusTitle, _
            Status:=ProgressStatusInitialising, _
            Min:=0, _
            Max:=ProgressCounterMax, _
            CancelButtonText:=ProgressStatusStop, _
            optShowTimeElapsed:=True, _
            optShowTimeRemaining:=True
        .Show
        .SetValue Value:=0
    End With
End Sub


Public Sub StartProgressBar(ByRef ProgressCounter As Double, ByRef ProgressCounterMax As Double)
    With ProgressStatus
        .SetStatus Status:=ProgressStatusProcessing & CStr(ProgressCounter) & ProgressStatusOf & CStr(ProgressCounterMax)
        .Show
    End With
End Sub


Public Sub UpdateProgressBar(ByRef ProgressCounter As Double, ByRef ProgressCounterMax As Double)
    With ProgressStatus
        .SetStatus Status:=ProgressStatusProcessing & CStr(ProgressCounter) & ProgressStatusOf & CStr(ProgressCounterMax)
        .SetValue Value:=ProgressStatus.GetValue + 1
        .Show
    End With
End Sub


Public Sub CloseProgressBar()
'    ProgressStatus.Hide
    Unload ProgressStatus
End Sub


