Attribute VB_Name = "ControlPanel"
Option Explicit

' Launcher and auto-refresh timer for frmControlPanel.
'
' The Win32 timer callback (PanelTimerCallback) must live in a standard module
' because AddressOf cannot reference methods on class/form modules.
' This module therefore owns the timer lifecycle and the form object reference.
'
' Usage: run ShowControlPanel from the VBE Immediate Window to open the panel.
' The panel auto-refreshes its monitor counts approximately every 3 seconds.

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private m_PanelTimerID As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private m_PanelTimerID As Long
#End If

' Interval between display refreshes (milliseconds).
Private Const PANEL_REFRESH_MS As Long = 3000

' The live panel instance. Nothing when the panel is not open.
Private m_Panel As frmControlPanel

' --- Public API -------------------------------------------------------------

' Open the control panel as a modeless form. If it is already open, bring it
' to the front instead of creating a second instance.
Public Sub ShowControlPanel()
    If Not m_Panel Is Nothing Then
        ' Try to bring the existing panel to the front.
        ' If the reference is stale (form closed without firing QueryClose),
        ' the Show call raises an error and we fall through to recreate.
        On Error Resume Next
        m_Panel.Show
        Dim alreadyOpen As Boolean
        alreadyOpen = (Err.Number = 0)
        On Error GoTo 0
        If alreadyOpen Then Exit Sub
        ' Stale reference — clean up before recreating.
        Set m_Panel = Nothing
        StopPanelTimer
    End If

    Set m_Panel = New frmControlPanel
    m_Panel.Show vbModeless
    StartPanelTimer
End Sub

' Called by frmControlPanel.UserForm_QueryClose to stop the refresh timer
' and clear the module-level reference.
Public Sub OnPanelClosed()
    StopPanelTimer
    Set m_Panel = Nothing
End Sub

' Win32 callback fired every PANEL_REFRESH_MS. Signature must match exactly.
' Refreshes the monitor section of the open panel without disrupting other
' Outlook activity.
#If VBA7 Then
Public Sub PanelTimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                               ByVal idEvent As LongPtr, ByVal dwTime As Long)
#Else
Public Sub PanelTimerCallback(ByVal hwnd As Long, ByVal uMsg As Long, _
                               ByVal idEvent As Long, ByVal dwTime As Long)
#End If
    If Not m_Panel Is Nothing Then
        On Error Resume Next
        m_Panel.RefreshDisplay
        On Error GoTo 0
    End If
End Sub

' --- Private helpers --------------------------------------------------------

Private Sub StartPanelTimer()
    If m_PanelTimerID <> 0 Then StopPanelTimer
    m_PanelTimerID = SetTimer(0, 0, PANEL_REFRESH_MS, AddressOf PanelTimerCallback)
    If m_PanelTimerID = 0 Then
        LogError "ControlPanel: SetTimer failed - monitor will not auto-refresh"
    End If
End Sub

Private Sub StopPanelTimer()
    If m_PanelTimerID <> 0 Then
        KillTimer 0, m_PanelTimerID
        m_PanelTimerID = 0
    End If
End Sub
