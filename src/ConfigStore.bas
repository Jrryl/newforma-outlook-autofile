Attribute VB_Name = "ConfigStore"
Option Explicit

Private Const HEADER As String = _
    "project_no|project_name|match_terms|staging_path|destination_path|status|enabled|last_seen"

Private Const COL_COUNT As Integer = 8

' --- Public API -------------------------------------------------------------

' Read the config file and return a Collection of ProjectConfig objects.
' Returns an empty Collection when the file does not exist yet (first run).
' Returns Nothing when the file exists but cannot be read (e.g. locked in
' Excel), so the caller can distinguish "no rows" from "cannot read".
Public Function LoadConfig(filePath As String) As Collection
    On Error GoTo ReadFail

    Dim result As Collection
    Set result = New Collection

    If Dir(filePath) = "" Then
        Set LoadConfig = result
        Exit Function
    End If

    Dim fnum As Integer
    fnum = FreeFile
    Open filePath For Input As #fnum

    Dim line As String
    Dim isHeader As Boolean
    isHeader = True

    Do While Not EOF(fnum)
        Line Input #fnum, line
        line = Trim(line)
        If isHeader Then
            isHeader = False
        ElseIf line <> "" Then
            Dim cfg As ProjectConfig
            Set cfg = ParseRow(line)
            If Not cfg Is Nothing Then result.Add cfg
        End If
    Loop

    Close #fnum
    Set LoadConfig = result
    Exit Function

ReadFail:
    If fnum > 0 Then Close #fnum
    Set LoadConfig = Nothing
End Function

' Write the full config to disk. Returns True on success.
Public Function SaveConfig(filePath As String, rows As Collection) As Boolean
    On Error GoTo WriteFail

    Dim fnum As Integer
    fnum = FreeFile
    Open filePath For Output As #fnum
    Print #fnum, HEADER

    Dim cfg As ProjectConfig
    For Each cfg In rows
        Print #fnum, FormatRow(cfg)
    Next cfg

    Close #fnum
    SaveConfig = True
    Exit Function

WriteFail:
    If fnum > 0 Then Close #fnum
    SaveConfig = False
End Function

' Return the first row whose ProjectNo matches, or Nothing if not found.
Public Function FindByProjectNo(rows As Collection, projNo As String) As ProjectConfig
    Dim cfg As ProjectConfig
    For Each cfg In rows
        If cfg.ProjectNo = projNo Then
            Set FindByProjectNo = cfg
            Exit Function
        End If
    Next cfg
    Set FindByProjectNo = Nothing
End Function

' --- Private helpers --------------------------------------------------------

Private Function ParseRow(line As String) As ProjectConfig
    Dim parts() As String
    parts = Split(line, "|")

    If UBound(parts) < COL_COUNT - 1 Then
        Set ParseRow = Nothing
        Exit Function
    End If

    Dim cfg As ProjectConfig
    Set cfg = New ProjectConfig
    cfg.ProjectNo       = Trim(parts(0))
    cfg.ProjectName     = Trim(parts(1))
    cfg.MatchTerms      = Trim(parts(2))
    cfg.StagingPath     = Trim(parts(3))
    cfg.DestinationPath = Trim(parts(4))
    cfg.Status          = Trim(parts(5))
    cfg.Enabled         = Trim(parts(6))
    cfg.LastSeen        = Trim(parts(7))
    Set ParseRow = cfg
End Function

Private Function FormatRow(cfg As ProjectConfig) As String
    FormatRow = cfg.ProjectNo & "|" & _
                cfg.ProjectName & "|" & _
                cfg.MatchTerms & "|" & _
                cfg.StagingPath & "|" & _
                cfg.DestinationPath & "|" & _
                cfg.Status & "|" & _
                cfg.Enabled & "|" & _
                cfg.LastSeen
End Function
