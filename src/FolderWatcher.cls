VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Holds a WithEvents reference to the Newforma parent's Folders collection so
' that adding a new project folder triggers an immediate reconciliation without
' waiting for the next Outlook restart.
'
' The startup reconciliation is the reliable path (folder-add events are missed
' across restarts). This watcher is a responsiveness improvement only.

Private WithEvents m_Folders As Outlook.Folders

' Call once at startup with newformaParent.Folders.
Public Sub Watch(folders As Outlook.Folders)
    Set m_Folders = folders
End Sub

Private Sub m_Folders_FolderAdd(ByVal folder As MAPIFolder)
    LogInfo "FolderWatcher: folder added (" & folder.Name & ") — triggering reconciliation"
    RunReconciliation
End Sub
