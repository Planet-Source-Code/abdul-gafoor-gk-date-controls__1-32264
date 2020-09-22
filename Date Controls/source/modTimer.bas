Attribute VB_Name = "modTimer"
'// ================================================================================
'// Copyright © 2001-2002 by Abdul Gafoor.GK
'// ================================================================================
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public gcolTimers As Collection

Public Sub TimerProc(ByVal lngHwnd As Long, ByVal lngMsg As Long, _
                     ByVal lngTimerID As Long, ByVal lngTime As Long)
    Dim clsTimer    As cTimer
    Dim lngPtr      As Long
    
    lngPtr = gcolTimers(GetTimerKey(lngTimerID))
    Call CopyMemory(clsTimer, lngPtr, 4&)
    Call clsTimer.Tick
    Call CopyMemory(clsTimer, 0&, 4&)
End Sub

Public Function GetTimerKey(ByVal lngTimerID As Long) As String
    GetTimerKey = "TMR" & Trim$(CStr(lngTimerID))
End Function
