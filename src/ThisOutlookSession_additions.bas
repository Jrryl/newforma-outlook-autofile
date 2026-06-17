' =============================================================================
' INSTRUCTIONS: Paste into the existing ThisOutlookSession module.
' Do NOT import this as a new module.
' This is the complete final state of the module (Workflows 1 and 2).
' =============================================================================
Option Explicit

Private m_FolderWatcher As FolderWatcher
Private WithEvents m_InboxItems As Outlook.Items

Private Sub Application_Startup()
    LogInfo "ThisOutlookSession: Outlook started"
    InitFolderWatcher
    InitInboxWatcher
    StartSweepTimer
    RunReconciliation
End Sub

Private Sub Application_Quit()
    StopSweepTimer
End Sub

' --- Workflow 1: folder detection --------------------------------------------

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

' --- Workflow 2: arrival routing ---------------------------------------------

Private Sub InitInboxWatcher()
    On Error GoTo Fail
    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")
    Set m_InboxItems = ns.GetDefaultFolder(olFolderInbox).Items
    LogInfo "InitInboxWatcher: inbox watcher active"
    Exit Sub
Fail:
    LogError "InitInboxWatcher: " & Err.Description
End Sub

Private Sub m_InboxItems_ItemAdd(ByVal item As Object)
    If TypeOf item Is Outlook.MailItem Then
        Dim rows As Collection
        Set rows = LoadConfig(ConfigFilePath())
        If Not rows Is Nothing Then RouteItem item, rows
    End If
End Sub
