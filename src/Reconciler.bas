Attribute VB_Name = "Reconciler"
Option Explicit

' Main entry point for Workflow 1.
'
' configPath        - full path to the pipe-delimited config file; defaults to
'                     ConfigFilePath() when omitted.
' stagingParentPath - Outlook FolderPath of the staging root folder; when
'                     omitted the function resolves or creates STAGING_PARENT_NAME
'                     at the root of the default data store.
Public Sub RunReconciliation(Optional configPath As String = "", _
                              Optional stagingParentPath As String = "")

    Dim newformaParent  As Outlook.Folder
    Dim stagingParent   As Outlook.Folder
    Dim liveFolders     As Collection
    Dim existingRows    As Collection
    Dim merged          As Collection
    Dim live            As ProjectConfig
    Dim ex              As ProjectConfig
    Dim existing        As ProjectConfig
    Dim newRow          As ProjectConfig
    Dim row             As ProjectConfig
    Dim sf              As Outlook.Folder
    Dim ns              As Outlook.NameSpace
    Dim root            As Outlook.Folder
    Dim subName         As String
    Dim today           As String

    If configPath = "" Then configPath = ConfigFilePath()
    LogInfo "Reconciler: run started"

    ' --- Step 1: Resolve the Newforma parent ---------------------------------
    Set newformaParent = FindNewformaParent()
    If newformaParent Is Nothing Then
        LogError "Reconciler: Newforma parent folder not found - aborting"
        Exit Sub
    End If

    ' --- Resolve (or create) the staging parent ------------------------------
    If stagingParentPath <> "" Then
        Set stagingParent = ResolveFolder(stagingParentPath)
    End If

    If stagingParent Is Nothing Then
        Set ns = Application.GetNamespace("MAPI")
        Set root = ns.GetDefaultFolder(6).Parent
        Set stagingParent = EnsureSubfolder(root, STAGING_PARENT_NAME)
    End If

    If stagingParent Is Nothing Then
        LogError "Reconciler: cannot resolve or create staging parent - aborting"
        Exit Sub
    End If

    ' --- Step 2: Enumerate live folders --------------------------------------
    Set liveFolders = ScanProjectFolders(newformaParent, stagingParent)
    If liveFolders Is Nothing Then
        LogError "Reconciler: folder scan failed - aborting"
        Exit Sub
    End If

    ' --- Load existing config ------------------------------------------------
    Set existingRows = LoadConfig(configPath)
    If existingRows Is Nothing Then
        LogWarn "Reconciler: config file locked or unreadable - skipping run"
        Exit Sub
    End If

    ' --- False-stale guard (spec s.6, step 2) --------------------------------
    ' A live scan returning zero folders against a populated config is almost
    ' certainly a failed or mid-sync enumeration. Treating zero children as
    ' "every project archived" would disable the entire file in one pass.
    If liveFolders.Count = 0 And CountActiveRows(existingRows) > 0 Then
        LogWarn "Reconciler: live scan returned 0 folders against " & _
                CountActiveRows(existingRows) & " active rows - " & _
                "treating as failed enumeration, aborting"
        Exit Sub
    End If

    ' --- Steps 3-5: Three-way merge ------------------------------------------
    Set merged = New Collection
    today = Format(Date, "yyyy-mm-dd")

    ' Pass A - live folders: add new rows, refresh known rows.
    For Each live In liveFolders
        Set existing = FindByProjectNo(existingRows, live.ProjectNo)

        If existing Is Nothing Then
            ' New project: seed a fresh row with user-owned fields at defaults.
            Set newRow = New ProjectConfig
            newRow.ProjectNo       = live.ProjectNo
            newRow.ProjectName     = live.ProjectName
            newRow.MatchTerms      = live.ProjectNo & ";" & live.ProjectName
            newRow.StagingPath     = live.StagingPath
            newRow.DestinationPath = live.DestinationPath
            newRow.Status          = "active"
            newRow.Enabled         = "TRUE"
            newRow.LastSeen        = today
            merged.Add newRow
            LogInfo "Reconciler: new project - " & newRow.ProjectNo & " " & newRow.ProjectName

        Else
            ' Known project, folder present: update detector-owned fields;
            ' leave match_terms and enabled exactly as the user left them.
            existing.ProjectName     = live.ProjectName
            existing.StagingPath     = live.StagingPath
            existing.DestinationPath = live.DestinationPath
            existing.Status          = "active"
            existing.LastSeen        = today
            merged.Add existing
            LogInfo "Reconciler: refreshed - " & existing.ProjectNo
        End If
    Next live

    ' Pass B - existing rows absent from the live scan: mark archived.
    For Each ex In existingRows
        If FindByProjectNo(merged, ex.ProjectNo) Is Nothing Then
            ex.Status = "archived"
            merged.Add ex
            LogInfo "Reconciler: archived - " & ex.ProjectNo & " " & ex.ProjectName
        End If
    Next ex

    ' --- Step 6: Create staging sub-folders for active projects --------------
    For Each row In merged
        If row.Status = "active" Then
            subName = row.ProjectNo & " - " & row.ProjectName
            Set sf = EnsureSubfolder(stagingParent, subName)
            If sf Is Nothing Then
                LogWarn "Reconciler: could not create staging folder for " & row.ProjectNo
            End If
        End If
    Next row

    ' --- Step 7: Write back --------------------------------------------------
    If SaveConfig(configPath, merged) Then
        LogInfo "Reconciler: complete - " & merged.Count & " rows written"
    Else
        LogError "Reconciler: failed to write config file"
    End If
End Sub
