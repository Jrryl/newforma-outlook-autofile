Attribute VB_Name = "Common"
Option Explicit

' En-dash (U+2013) as used in the Newforma add-in folder name.
Public Const STAGING_PARENT_NAME As String = "Project Staging"
Private Const APP_FOLDER As String = "OutlookAutoFile"

Public Const NEWFORMA_PARENT_NAME As String = "Newforma - Items to File"

' --- Paths ------------------------------------------------------------------

Public Function ConfigFilePath() As String
    ConfigFilePath = AppDataFolder() & "\projects.psv"
End Function

Public Function LogFilePath() As String
    LogFilePath = AppDataFolder() & "\autofile.log"
End Function

Private Function AppDataFolder() As String
    Dim p As String
    p = Environ("APPDATA") & "\" & APP_FOLDER
    If Dir(p, vbDirectory) = "" Then MkDir p
    AppDataFolder = p
End Function

' --- Logging ----------------------------------------------------------------

Public Sub LogInfo(msg As String)
    LogLine "INFO", msg
End Sub

Public Sub LogWarn(msg As String)
    LogLine "WARN", msg
End Sub

Public Sub LogError(msg As String)
    LogLine "ERROR", msg
End Sub

Private Sub LogLine(level As String, msg As String)
    On Error GoTo Fail
    Dim fnum As Integer
    fnum = FreeFile
    Open LogFilePath() For Append As #fnum
    Print #fnum, Format(Now, "yyyy-mm-dd hh:mm:ss") & " [" & level & "] " & msg
    Close #fnum
Fail:
    ' Cannot log a logging failure — silent exit.
End Sub

' --- Folder helpers ---------------------------------------------------------

' Navigate to a folder by its Outlook FolderPath string
' (e.g. "\\Store Name\Parent\Child"). Returns Nothing if any segment is
' missing or the store cannot be matched.
Public Function ResolveFolder(folderPath As String) As Outlook.Folder
    On Error GoTo Fail

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' FolderPath starts with "\\" so Split on "\" yields two leading empty strings.
    Dim parts() As String
    parts = Split(folderPath, "\")

    Dim i As Integer
    For i = 0 To UBound(parts)
        If parts(i) <> "" Then Exit For
    Next i
    If i > UBound(parts) Then GoTo Fail

    ' Match the store by display name (the first non-empty segment).
    Dim current As Outlook.Folder
    Dim st As Outlook.Store
    For Each st In ns.Stores
        If st.DisplayName = parts(i) Then
            Set current = st.GetRootFolder()
            Exit For
        End If
    Next st
    If current Is Nothing Then GoTo Fail

    ' Descend through the remaining path segments.
    Dim segFound As Boolean
    Dim child As Outlook.Folder
    i = i + 1
    Do While i <= UBound(parts)
        If parts(i) <> "" Then
            segFound = False
            For Each child In current.Folders
                If child.Name = parts(i) Then
                    Set current = child
                    segFound = True
                    Exit For
                End If
            Next child
            If Not segFound Then GoTo Fail
        End If
        i = i + 1
    Loop

    Set ResolveFolder = current
    Exit Function
Fail:
    Set ResolveFolder = Nothing
End Function

' Return a named sub-folder of parent, creating it if it does not exist.
' Returns Nothing on any failure.
Public Function EnsureSubfolder(parent As Outlook.Folder, _
                                subName As String) As Outlook.Folder
    On Error GoTo Fail
    Dim f As Outlook.Folder
    For Each f In parent.Folders
        If f.Name = subName Then
            Set EnsureSubfolder = f
            Exit Function
        End If
    Next f
    Set EnsureSubfolder = parent.Folders.Add(subName)
    Exit Function
Fail:
    Set EnsureSubfolder = Nothing
End Function
