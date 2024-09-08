Attribute VB_Name = "modController"
Option Explicit

'''
' This utilises `clsTracker` to schedule events by clicking the Dashboard button.
'
' The `clsTracker` object is just a conduit for writing to sheets and scheduling
' events, so don't need the same instance to persist between pop-ups.
'''
'Private tTracker As New clsTracker

Private Enum ButtonState
    btnStart = 0
    btnEnd = 1
End Enum


Public Sub ButtonControl_Click()
    '''
    ' Dashboard button which toggles between "Start Tracker" and "End Tracker".
    '''
    With wsDashboard.Shapes("btnToggleTracker").TextFrame2.TextRange.Characters
        If .text = "Start Tracker" Then
            Call StartIntervalTracker
        ElseIf .text = "End Tracker" Then
            Call EndIntervalTracker
        Else
            Err.Raise Number:=513, Description:="The Tracker button has unrecognised text"
        End If
    End With
End Sub


Public Sub ExecuteSchedule()
    '''
    ' Callback for the scheduler as we can't schedule class methods.
    '''
    ThisWorkbook.Activate
    Dim tTracker As New clsTracker

    Call tTracker.ExecuteSchedule

    Set tTracker = Nothing
End Sub


' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Public Sub StartIntervalTracker()
    '''
    ' Start the tracker by setting the default values and then scheduling the
    ' first event.
    '
    ' Needs to be public for `ThisWorkbook` subroutines.
    '''
    Dim tTracker As New clsTracker

    Call tTracker.SetDefaultValues
    Call tTracker.ScheduleFirst

    If tTracker.Running Then
        Call ToggleButton(eButtonState:=btnEnd)
    Else
        Err.Raise Number:=513, Description:="The bRunning cell in the variables sheet should be TRUE"
    End If

    Set tTracker = Nothing
End Sub

Public Sub EndIntervalTracker()
    '''
    ' End the tracker by cancelling the upcoming event.
    '
    ' Needs to be public for `ThisWorkbook` subroutines.
    '''
    Dim tTracker As New clsTracker

    Call tTracker.CancelNext

    If tTracker.Running = False Then
        Call ToggleButton(eButtonState:=btnStart)
    Else
        Err.Raise Number:=513, Description:="The bRunning cell in the variables sheet should be FALSE"
    End If

    Set tTracker = Nothing
End Sub


' ----------------------------------------------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------

Private Sub ToggleButton(ByVal eButtonState As ButtonState)
    '''
    ' Toggle the button to 'start' and 'end'.
    '''
    Select Case eButtonState
        Case btnStart
            ' Green with "Start Tracker" text
            Call ToggleButtonBase( _
                vColourRGB:=RGB(197, 224, 180), _
                sText:="Start Tracker" _
            )
        Case btnEnd
            ' Orange with "End Tracker" text
            Call ToggleButtonBase( _
                vColourRGB:=RGB(248, 203, 173), _
                sText:="End Tracker" _
            )
        Case Else
            Err.Raise Number:=513, Description:="Bad value passed to modController.ToggleButton()"
    End Select
End Sub


Private Sub ToggleButtonBase(ByVal vColourRGB As Variant, ByVal sText As String)
    '''
    ' Change the Dashboard button's colour and text.
    '''
    With wsDashboard.Shapes("btnToggleTracker")
        .Fill.ForeColor.RGB = vColourRGB
        .TextFrame2.TextRange.Characters.text = sText
    End With
End Sub
