Attribute VB_Name = "Timer"
Option Explicit

' Win32 timer for the timed sweep. VBA has no native timer, so this uses
' SetTimer from user32.dll. The callback must be a Public Sub in a standard
' module (not a class module) so that AddressOf can reference it.
'
' Conditional compilation covers 32-bit Outlook (VBA6) and 64-bit (VBA7).

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private m_TimerID As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private m_TimerID As Long
#End If

' 5 minutes between sweeps. Adjust as needed.
Private Const SWEEP_INTERVAL_MS As Long = 300000

Public Sub StartSweepTimer()
    If m_TimerID <> 0 Then StopSweepTimer
    m_TimerID = SetTimer(0, 0, SWEEP_INTERVAL_MS, AddressOf SweepTimerCallback)
    If m_TimerID = 0 Then
        LogError "Timer: SetTimer failed - sweep will not run automatically"
    Else
        LogInfo "Timer: sweep timer started (" & SWEEP_INTERVAL_MS / 1000 & "s interval)"
    End If
End Sub

Public Sub StopSweepTimer()
    If m_TimerID <> 0 Then
        KillTimer 0, m_TimerID
        m_TimerID = 0
        LogInfo "Timer: stopped"
    End If
End Sub

' Win32 callback - signature must match exactly. Called by Windows on each tick.
#If VBA7 Then
Public Sub SweepTimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                               ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub SweepTimerCallback(ByVal hwnd As Long, ByVal uMsg As Long, _
                               ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    RunSweep
End Sub
