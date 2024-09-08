VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressStatusDialogue 
   Caption         =   "Progress"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6135
   OleObjectBlob   =   "frmProgressStatusDialogue.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressStatusDialogue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Cancelled    As Boolean
Dim showTime     As Boolean
Dim showTimeLeft As Boolean
Dim StartTime    As Long
Dim BarMin       As Long
Dim BarMax       As Long
Dim BarVal       As Long


' Title will be the title of the dialogue
' Status will be the label above the progress bar, and can be changed with SetStatus
' Min is the progress bar minimum value, only set by calling configure
' Max is the progress bar maximum value, only set by calling configure
' CancelButtonText is the caption of the cancel button. If set to vbNullString, it is hidden
' optShowTimeElapsed controls whether the progress bar computes and displays the time elapsed
' optShowTimeRemaining controls whether the progress bar estimates and displays the time remaining

' calling Configure sets the current value equal to Min
' calling Configure resets the current run time


Public Sub Configure( _
    ByVal Title As String, _
    ByVal Status As String, _
    ByVal Min As Long, _
    ByVal Max As Long, _
    Optional ByVal CancelButtonText As String = "Cancel", _
    Optional ByVal optShowTimeElapsed As Boolean = True, _
    Optional ByVal optShowTimeRemaining As Boolean = True, _
    Optional ByVal PercentageColour As Double = -1, _
    Optional ByVal ProgressBarColour As Double = -1, _
    Optional ByVal ProgressBarBackGroundColour As Double = -1 _
)
    Let Cancelled = False
    Let showTime = optShowTimeElapsed
    Let showTimeLeft = optShowTimeRemaining
    Let StartTime = Now()
    Let BarMin = Min
    Let BarMax = Max
    Let BarVal = Min

    Me.Caption = Title
    lblStatus.Caption = Status
    CancelButton.Visible = Not (CancelButtonText = vbNullString)
    CancelButton.Caption = CancelButtonText
    lblRunTime.Caption = ""
    lblRemainingTime.Caption = ""
    
    If PercentageColour >= 0 Then lblPercent.ForeColor = PercentageColour
    If ProgressBarColour >= 0 Then ProgressBar.BackColor = ProgressBarColour
    If ProgressBarBackGroundColour >= 0 Then ProgressBarBG.BackColor = ProgressBarBackGroundColour
End Sub


' Set the value of the status bar, a long which is snapped to a value between Min and Max
Public Sub SetValue(ByVal Value As Long)
    Dim Progress
    Dim PredictedTimeDouble As Double
    Dim RunTime             As Long

    Let BarVal = WorksheetFunction.Min(WorksheetFunction.Max(Value, BarMin), BarMax)
    Let Progress = (BarVal - BarMin) / (BarMax - BarMin)
    Let RunTime = GetRunTime()

    ProgressBar.Width = 292 * Progress
    lblPercent = Int(Progress * 1000) / 10 & "%"

    If showTime Then lblRunTime.Caption = "Time Elapsed: " & GetRunTimeString(RunTime, False)
    If showTimeLeft And Progress > 0 Then
        Let PredictedTimeDouble = RunTime * ((1 - Progress) / Progress)
        If PredictedTimeDouble > 36000000 Then
            lblRemainingTime.Caption = "Est. Time Left: More than 10 hours"
        Else
            lblRemainingTime.Caption = "Est. Time Left: " & GetRunTimeString(CLng(PredictedTimeDouble), False)
        End If
    End If
    DoEvents
End Sub


' Formats a time in milliseconds as hours, minutes, seconds.milliseconds
' Milliseconds are excluded if showMsecs is set to false
Private Function GetRunTimeString(ByVal RunTime As Long, Optional ByVal showMsecs As Boolean = True) As String
    Dim msecs As Long
    Dim hrs   As Long
    Dim mins  As Long
    Dim secs  As Double

    Let msecs = RunTime
    Let hrs = Int(msecs / 3600000)
    Let mins = Int(msecs / 60000) - 60 * hrs
    Let secs = msecs / 1000 - 60 * (mins + 60 * hrs)
    
    GetRunTimeString = "" _
        & IIf(hrs > 0, hrs & " hours ", "") _
        & IIf(mins > 0, mins & " minutes ", "") _
        & IIf(secs > 0 And mins = 0, IIf(showMsecs, secs, Int(secs + 0.5)) & " seconds", "")
End Function


Private Sub CancelButton_Click()
    Cancelled = True
    lblStatus.Caption = "Cancelled By User. Please Wait."
End Sub


Public Sub SetStatus(ByVal Status As String)
    lblStatus.Caption = Status
    DoEvents
End Sub


Public Function GetRunTime() As Long
    GetRunTime = Now() - StartTime
End Function

Public Function GetFormattedRunTime() As String
    GetFormattedRunTime = GetRunTimeString(Now() - StartTime)
End Function

Public Function GetValue() As Long
    GetValue = BarVal
End Function

Public Function cancelIsPressed() As Boolean
    cancelIsPressed = Cancelled
End Function

Private Sub UserForm_Click()

End Sub
