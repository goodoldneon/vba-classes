Attribute VB_Name = "Optimize"
' This module will speed up VBA by disabling screen updating, events, auto calculation, and page breaks.
' The states of events, calculation, and page breaks are saved and restored after optimizeEnd() is called.


Public eventState As Boolean
Public calcState As Long
Public pageBreakState As Boolean


Public Sub optimizeStart()
    application.ScreenUpdating = False
    
    eventState = application.EnableEvents
    application.EnableEvents = False
    
    calcState = application.Calculation
    application.Calculation = xlCalculationManual
    
    pageBreakState = ActiveSheet.DisplayPageBreaks
    ActiveSheet.DisplayPageBreaks = False
End Sub


Public Sub optimizeEnd()
    ActiveSheet.DisplayPageBreaks = pageBreakState
    application.Calculation = calcState
    application.EnableEvents = eventState
    application.ScreenUpdating = True
End Sub

