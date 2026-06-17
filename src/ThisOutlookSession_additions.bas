' =============================================================================
' INSTRUCTIONS: Do NOT import this file as a new module.
' Open the Outlook VBA editor (Alt+F11), expand "Microsoft Outlook Objects",
' and paste the code below into the existing "ThisOutlookSession" module.
' =============================================================================
Option Explicit

' Keeps the FolderWatcher alive for the lifetime of the Outlook session.
Private m_FolderWatcher As FolderWatcher

Private Sub Application_Startup()
    LogInfo "ThisOutlookSession: Outlook started"
    InitFolderWatcher
    RunReconciliation
End Sub

' Wire up the FolderAdd event on the Newforma parent so that a new project
' folder triggers reconciliation immediately, without waiting for the next
' Outlook restart.
Private Sub InitFolderWatcher()
    On Error GoTo Fail

    Dim parent As Outlook.Folder
    Set parent = FindNewformaParent()

    If parent Is Nothing Then
        LogWarn "InitFolderWatcher: Newforma parent not found - folder-add event not wired"
        Exit Sub
    End If

    Set m_FolderWatcher = New FolderWatcher
    m_FolderWatcher.Watch parent.Folders
    LogInfo "InitFolderWatcher: watching " & parent.FolderPath
    Exit Sub

Fail:
    LogError "InitFolderWatcher: " & Err.Description
End Sub
