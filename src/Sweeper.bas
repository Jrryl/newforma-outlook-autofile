Attribute VB_Name = "Sweeper"
Option Explicit

Private m_SweepRunning As Boolean

' Called by the timer tick. Re-scans the Inbox (backstop for missed
' ItemAdd events) then releases qualifying items from staging folders.
Public Sub RunSweep()
    If m_SweepRunning Then
        LogWarn "Sweeper: already running - skipping this tick"
        Exit Sub
    End If

    m_SweepRunning = True
    On Error GoTo Cleanup

    LogInfo "Sweeper: sweep started"
    RouteInboxItems
    ReleaseQualifyingItems
    LogInfo "Sweeper: sweep complete"

Cleanup:
    m_SweepRunning = False
    If Err.Number <> 0 Then LogError "Sweeper: " & Err.Description
End Sub

' --- Private helpers --------------------------------------------------------

' Walk every active+enabled project's staging folder and move items that are
' read and not actively flagged to the project's Newforma destination folder.
Private Sub ReleaseQualifyingItems()
    On Error GoTo Fail

    Dim rows As Collection
    Set rows = LoadConfig(ConfigFilePath())
    If rows Is Nothing Then
        LogWarn "Sweeper: config unreadable - skipping release"
        Exit Sub
    End If

    Dim cfg As ProjectConfig
    For Each cfg In rows
        If LCase(cfg.Status) = "active" And LCase(cfg.Enabled) = "true" Then
            ReleaseFromStaging cfg
        End If
    Next cfg
    Exit Sub
Fail:
    LogError "Sweeper: ReleaseQualifyingItems - " & Err.Description
End Sub

Private Sub ReleaseFromStaging(cfg As ProjectConfig)
    On Error GoTo Fail

    Dim stagingFolder As Outlook.Folder
    Set stagingFolder = ResolveFolder(cfg.StagingPath)
    If stagingFolder Is Nothing Then Exit Sub

    Dim destFolder As Outlook.Folder
    Set destFolder = ResolveFolder(cfg.DestinationPath)
    If destFolder Is Nothing Then
        LogWarn "Sweeper: destination not found for " & cfg.ProjectNo
        Exit Sub
    End If

    ' Walk backwards so that moving items does not disturb the index.
    Dim i As Integer
    For i = stagingFolder.Items.Count To 1 Step -1
        Dim obj As Object
        Set obj = stagingFolder.Items(i)
        If TypeOf obj Is Outlook.MailItem Then
            Dim mail As Outlook.MailItem
            Set mail = obj
            ' Release if read AND not carrying an active follow-up flag.
            ' FlagStatus: 0 = no flag, 1 = complete, 2 = marked (olFlagMarked).
            ' Both 0 and 1 are safe to release; only 2 is held back.
            If Not mail.UnRead And mail.FlagStatus <> olFlagMarked Then
                LogInfo "Sweeper: releasing """ & mail.Subject & """ for " & cfg.ProjectNo
                mail.Move destFolder
            End If
        End If
    Next i
    Exit Sub
Fail:
    LogError "Sweeper: ReleaseFromStaging (" & cfg.ProjectNo & ") - " & Err.Description
End Sub
