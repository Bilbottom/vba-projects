VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pScreenUpdating As Boolean
Private pCalculation    As XlCalculation


Public Property Let SetApplicationValues(ByRef Application As Application)
    With Application
        Let pScreenUpdating = .ScreenUpdating
        Let pCalculation = .Calculation

        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
End Property


Public Sub ResetApplicationValues()
    With Application
        .ScreenUpdating = pScreenUpdating
        .Calculation = pCalculation
    End With
End Sub
