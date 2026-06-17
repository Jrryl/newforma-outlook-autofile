Attribute VB_Name = "FolderScanner"
Option Explicit

' Return the "Newforma – Items to File" folder from the default store,
' or Nothing if it cannot be found.
Public Function FindNewformaParent() As Outlook.Folder
    On Error GoTo Fail

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim st As Outlook.Store
    Dim f As Outlook.Folder
    For Each st In ns.Stores
        For Each f In st.GetRootFolder().Folders
            If f.Name = NEWFORMA_PARENT_NAME Then
                Set FindNewformaParent = f
                Exit Function
            End If
        Next f
    Next st

Fail:
    Set FindNewformaParent = Nothing
End Function

' Scan live child folders under the Newforma parent and return a Collection
' of partially-populated ProjectConfig objects (ProjectNo, ProjectName,
' DestinationPath, StagingPath). Folders that do not match the
' "<no> - <name>" pattern are silently skipped.
'
' Returns Nothing on error; returns an empty Collection when the parent has
' no conforming children.
Public Function ScanProjectFolders(newformaParent As Outlook.Folder, _
                                   stagingParent As Outlook.Folder) As Collection
    On Error GoTo Fail

    Dim result As Collection
    Set result = New Collection

    Dim f As Outlook.Folder
    Dim projNo As String
    Dim projName As String
    Dim cfg As ProjectConfig

    For Each f In newformaParent.Folders
        If TryParseFolderName(f.Name, projNo, projName) Then
            Set cfg = New ProjectConfig
            cfg.ProjectNo       = projNo
            cfg.ProjectName     = projName
            cfg.DestinationPath = f.FolderPath
            cfg.StagingPath     = stagingParent.FolderPath & "\" & f.Name
            result.Add cfg
        End If
    Next f

    Set ScanProjectFolders = result
    Exit Function
Fail:
    Set ScanProjectFolders = Nothing
End Function

' Return the number of rows in rows whose Status is "active".
Public Function CountActiveRows(rows As Collection) As Integer
    Dim cfg As ProjectConfig
    Dim n As Integer
    n = 0
    For Each cfg In rows
        If LCase(cfg.Status) = "active" Then n = n + 1
    Next cfg
    CountActiveRows = n
End Function

' --- Private helpers --------------------------------------------------------

' Split "<no> - <name>" on the FIRST " - " only, so project names that
' themselves contain " - " are preserved intact.
' Returns False if the separator is absent or the number part is empty.
Private Function TryParseFolderName(folderName As String, _
                                    ByRef projNo As String, _
                                    ByRef projName As String) As Boolean
    Const SEP As String = " - "
    Dim pos As Integer
    pos = InStr(1, folderName, SEP)

    If pos < 2 Then
        TryParseFolderName = False
        Exit Function
    End If

    projNo   = Left(folderName, pos - 1)
    projName = Mid(folderName, pos + Len(SEP))
    TryParseFolderName = (Len(Trim(projNo)) > 0 And Len(Trim(projName)) > 0)
End Function
