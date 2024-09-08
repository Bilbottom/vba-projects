VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''
' This `clsTracker` class generates the pop-up box frmTaskInputBox to record current work.
'''


'====================================================================================================
' Fields
' ------
'
' Declaring a variable without a Property Let/Get makes it a field rather than a property which only
' enforces a type.
'
' Good practice is for in-memory properties to use private fields prefixed with the letter `p`.
'====================================================================================================


'====================================================================================================
' Properties
' ----------
'
' This is better than using fields as this given better control over the reading and writing of
' properties (namely with the Let/Set and Get methods).
'
' Since in-memory values are lost between events, the values we want to persist are written to the
' `wsVariables` sheet.
'
' The values in `wsOptions` are also picked up by property getters to make them easier to use.
'====================================================================================================

' -----------
' wsVariables
' -----------

' Is Tracker Running
Public Property Let Running(ByVal v As Boolean)
    wsVariables.[bRunning].Value = v
End Property
Public Property Get Running() As Boolean
    Let Running = wsVariables.[bRunning].Value
End Property

' Next Event Time
Public Property Let ScheduleTime(ByVal v As String)
    wsVariables.[sScheduleTime].Value = v
End Property
Public Property Get ScheduleTime() As String
    Let ScheduleTime = wsVariables.[sScheduleTime].Value
End Property

' Current Time -- allows users to toggle the public Interval option without interfering with scheduled events
Public Property Let CurrentTime(ByVal v As String)
    wsVariables.[sCurrentTime].Value = v
End Property
Public Property Get CurrentTime() As String
    Let CurrentTime = wsVariables.[sCurrentTime].Value
End Property

' Current Interval -- allows users to toggle the public Interval option without interfering with scheduled events
Public Property Let CurrentInterval(ByVal v As Long)
    wsVariables.[lCurrentInterval].Value = v
End Property
Public Property Get CurrentInterval() As Long
    Let CurrentInterval = wsVariables.[lCurrentInterval].Value
End Property


' ---------
' wsOptions
' ---------

' Interval Minutes
Public Property Get Interval() As Long
    Let Interval = wsOptions.[lInterval].Value
End Property

' Run On Opening
Public Property Get RunOnOpening() As Boolean
    Let RunOnOpening = wsOptions.[bRunOnOpening].Value
End Property

' Save After Input
Public Property Get SaveAfterInput() As Boolean
    Let SaveAfterInput = wsOptions.[bSaveAfterInput].Value
End Property

' Minimise After Input
Public Property Get MinimiseAfterInput() As Boolean
    Let MinimiseAfterInput = wsOptions.[bMinimiseAfterInput].Value
End Property

' Use Calendar Defaults
Public Property Get UseCalendarDefaults() As Boolean
    Let UseCalendarDefaults = wsOptions.[bUseCalendarDefaults].Value
End Property


'====================================================================================================
' Functions
' ---------
'
' Functions to support instances of `clsTracker`.
'====================================================================================================

Private Function GetNextInterval(ByVal sStartTime As String, ByVal lInterval As Long) As String
    '''
    ' Convert the input time into the next schedulable time.
    '''
    Dim dTime As Double
    Let dTime = TimeValue(sStartTime)

    Let GetNextInterval = Format( _
        TimeSerial( _
            Hour:=Hour(dTime), _
            Minute:=lInterval + Minute(dTime) - (Minute(dTime) Mod lInterval), _
            Second:=0 _
        ), _
        "hh:mm" _
    )
End Function


'====================================================================================================
' Methods
' -------
'
' Subroutines to call on instances of `clsTracker`.
'====================================================================================================

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub


Public Sub SetDefaultValues()
    '''
    ' Set the default values when the workbook is opened or events are cancelled.
    '''
    With Me
        .Running = False
        .ScheduleTime = ""
        .CurrentTime = ""
        .CurrentInterval = .Interval
    End With
End Sub


Private Sub PrintCurrentVals()
    '''
    ' Check the values of the worksheet variables at run time.
    '''
    With Me
        Debug.Print "Running:" & vbTab & .Running
        Debug.Print "ScheduleTime:" & vbTab & .ScheduleTime
        Debug.Print "CurrentTime:" & vbTab & .CurrentTime
        Debug.Print "CurrentInterval:" & vbTab & .CurrentInterval

        Debug.Print "Interval:" & vbTab & .Interval
        Debug.Print "RunOnOpening:" & vbTab & .RunOnOpening
        Debug.Print "SaveAfterInput:" & vbTab & .SaveAfterInput
        Debug.Print "MinimiseAfterInput:" & vbTab & .MinimiseAfterInput
        Debug.Print "UseCalendarDefaults:" & vbTab & .UseCalendarDefaults
    End With
End Sub


Private Sub BaseScheduler(ByVal bSchedule As Boolean)
    '''
    ' Schedule the next event to call `ExecuteSchedule`.
    ' This will be scheduled at the current `ScheduleTime`.
    '''
    With Me
        .CurrentTime = .ScheduleTime
        .CurrentInterval = .Interval
    End With

    Application.OnTime _
        EarliestTime:=Date + TimeValue(Me.ScheduleTime), _
        Procedure:="ExecuteSchedule", _
        Schedule:=bSchedule
End Sub


Public Sub ScheduleFirst()
    '''
    ' Schedule the first pop-up event.
    '''
    If Me.Running Then Exit Sub

    Let Me.Running = True
    Let Me.ScheduleTime = GetNextInterval( _
        sStartTime:=Format(Now(), "hh:mm"), _
        lInterval:=Me.Interval _
    )

    Call BaseScheduler(bSchedule:=True)
End Sub


Private Sub ScheduleNext()
    '''
    ' Schedule the next pop-up event.
    '''
    If Me.ScheduleTime = "" Then Exit Sub

    Let Me.ScheduleTime = GetNextInterval( _
        Me.ScheduleTime, _
        Me.Interval _
    )

    Call BaseScheduler(bSchedule:=True)
End Sub


Public Sub CancelNext()
    '''
    ' Cancel the upcoming pop-up event.
    '''
    If Me.ScheduleTime = "" Then Exit Sub

    On Error Resume Next
        Call BaseScheduler(bSchedule:=False)
    On Error GoTo 0

    Call SetDefaultValues
End Sub


Public Sub ExecuteSchedule()
    '''
    ' Execute the event actions and schedule the next one using `CurrentTime`
    ' and `CurrentInterval`.
    '
    ' This needs to be called in a Module since you can't schedule class methods.
    '''
    Dim CalcState As XlCalculation
    With Application
        Let CalcState = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

'    If Format(Time, "hh:mm:ss") <> Me.ScheduleTime Then Call BaseScheduler(bSchedule:=False)  ' For when the schedule is called ahead of time  -- this is breaking the next OnTime event
    Call RunTrackerActions(sTime:=Me.CurrentTime, lInterval:=Me.CurrentInterval)
    Call ScheduleNext

    With Application
        .Calculation = CalcState
        .ScreenUpdating = True
    End With
End Sub


Private Sub RunTrackerActions(ByVal sTime As String, ByVal lInterval As Long)
    '''
    ' Launch the pop-up event and then save and/or minimise the workbook
    ' depending on the options.
    '''
    Dim TrackerBox   As frmTaskInputBox
    Set TrackerBox = New frmTaskInputBox
    With TrackerBox
'        .Interval = lInterval
'        .ScheduleTime = sTime
        .Show vbModeless
    End With

    ' Queue pop-ups instead of stack
    While IsLoaded("frmTaskInputBox")
        ' Stops responding without DoEvents
        DoEvents
    Wend

    ' Save and minimise workbook
    If Me.SaveAfterInput Then ThisWorkbook.Save
    On Error Resume Next
        If Me.MinimiseAfterInput Then ThisWorkbook.Windows(1).WindowState = xlMinimized
    On Error GoTo 0
End Sub
