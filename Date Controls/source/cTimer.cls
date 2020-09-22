VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// ================================================================================
'// Copyright © 2001-2002 by Abdul Gafoor.GK
'// ================================================================================
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_ENGLISH = &H9

Private mlngInterval    As Long
Private mblnEnabled     As Boolean
Private mlngTimerID     As Long

Public Event Tick()

Property Get Interval() As Long
Attribute Interval.VB_Description = "Gets/sets the interval at which the timer fires.  Set to zero to stop the timer."
    Interval = mlngInterval
End Property

Property Let Interval(ByVal vData As Long)
    mlngInterval = vData
    Call StartTimer
End Property

Public Sub Tick()
Attribute Tick.VB_MemberFlags = "40"
    If mblnEnabled Then RaiseEvent Tick
End Sub

Friend Property Get Enabled() As Boolean
    Enabled = mblnEnabled
End Property

Friend Property Let Enabled(ByVal vData As Boolean)
    mblnEnabled = vData
    Call StartTimer
End Property

Private Sub Class_Initialize()
    '// Set default value(s)
    mblnEnabled = False
    '// Initialize the collection, if not done already
    If gcolTimers Is Nothing Then Set gcolTimers = New Collection
End Sub

Private Sub Class_Terminate()
    '// Destroy timer, if any
    Call StopTimer
    '// Destroy the collection if no more item exists in the collection
    If (gcolTimers.Count > 0) Then Set gcolTimers = Nothing
End Sub

Private Function StartTimer() As Boolean
    Dim lngReturn   As Long
    Dim strError    As String
    
    If (mlngInterval > 0) And mblnEnabled Then
        '// Destroy the existing timer if any
        Call StopTimer
        '// Create new timer
        lngReturn = SetTimer(0&, 0&, mlngInterval, AddressOf modTimer.TimerProc)
        '// Check whether successfully created or not
        If (lngReturn > 0) Then
            '// Add the pointer to current timer object in a collection
            gcolTimers.Add ObjPtr(Me), GetTimerKey(lngReturn)
            mlngTimerID = lngReturn
            '// Return true
            StartTimer = True
        Else
            '// Show error, if not successful
            Call ShowError
        End If
    Else
        Call StopTimer
    End If
End Function

Private Function StopTimer() As Boolean
    If (mlngTimerID > 0) Then
        '// Destroy timer
        StopTimer = KillTimer(0&, mlngTimerID)
        '// Remove it from collection
        gcolTimers.Remove GetTimerKey(mlngTimerID)
        mlngTimerID = 0
        '// Return true
        StopTimer = True
    End If
End Function

Private Sub ShowError()
    Dim lngReturn   As Long
    Dim strError    As String
    
    lngReturn = GetLastError()
    strError = Space(255)
    Call FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lngReturn, LANG_ENGLISH, strError, 255, ByVal 0&)
    strError = Left(strError, InStr(1, strError, vbNullChar) - 1)
    
    Err.Raise vbObjectError + lngReturn, App.EXEName & "Timer", strError
End Sub
