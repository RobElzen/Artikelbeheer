Attribute VB_Name = "M90_DAF_Speed"
Option Explicit
Public glb_origCalculationMode As Integer
''''SpeedOn   ''Application.xxx=False
Sub SpeedOn() '(Optional StatusBarMsg As String = "Running macro... ")
On Error Resume Next

    glb_origCalculationMode = Application.Calculation
    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
       '.Cursor = xlWait
       '.StatusBar = StatusBarMsg
       '.EnableCancelKey = xlErrorHandler
    End With
End Sub
''''SpeedOn   ''Application.xxx=False

''''SpeedOff  ''Application.xxx=True
Sub SpeedOff()
    glb_origCalculationMode = Application.Calculation
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .CalculateBeforeSave = True
        .Cursor = xlDefault
       '.StatusBar = False
       '.EnableCancelKey = xlInterrupt
    End With
End Sub

