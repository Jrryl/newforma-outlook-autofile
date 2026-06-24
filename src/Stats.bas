Attribute VB_Name = "Stats"
Option Explicit

' Persistent counters for emails processed by Route and Sweep.
' File: %APPDATA%\OutlookAutoFile\stats.psv
' Schema (header + single data row):
'   date|today_routed|today_swept|life_routed|life_swept
'
' The single data row stores counts for today plus running lifetime totals.
' When the stored date differs from today, the today_* counters are zeroed on
' the next write (date rollover) while lifetime totals carry forward.

Private Const STATS_HEADER As String = "date|today_routed|today_swept|life_routed|life_swept"
Private Const COL_COUNT    As Integer = 5

Public Type StatsSnapshot
    TodayRouted As Long
    TodaySwept  As Long
    LifeRouted  As Long
    LifeSwept   As Long
End Type

' --- Public API -------------------------------------------------------------

' Call from Router.RouteItem after each successful mail.Move to staging.
Public Sub RecordRouted()
    IncrementStat "routed"
End Sub

' Call from Sweeper.ReleaseFromStaging after each successful mail.Move to dest.
Public Sub RecordSwept()
    IncrementStat "swept"
End Sub

' Read and return the current counters without writing.
' If the stored date is not today, today_* values are returned as 0 (stale).
' Returns all-zeros when the file does not yet exist.
Public Function ReadStats() As StatsSnapshot
    Dim snap As StatsSnapshot
    Dim stored As StatsSnapshot
    Dim storedDate As String
    LoadStatsFile storedDate, stored

    If storedDate = Format(Date, "yyyy-mm-dd") Then
        snap.TodayRouted = stored.TodayRouted
        snap.TodaySwept  = stored.TodaySwept
    End If
    snap.LifeRouted = stored.LifeRouted
    snap.LifeSwept  = stored.LifeSwept
    ReadStats = snap
End Function

' --- Private helpers --------------------------------------------------------

' Increment one counter, rolling over today_* if the date has changed.
' Wrapped entirely in On Error so a stats failure never disrupts a mail.Move.
Private Sub IncrementStat(kind As String)
    On Error GoTo Fail

    Dim storedDate As String
    Dim snap As StatsSnapshot
    LoadStatsFile storedDate, snap

    Dim today As String
    today = Format(Date, "yyyy-mm-dd")

    If storedDate <> today Then
        ' New day — reset today's counters and record the new date.
        snap.TodayRouted = 0
        snap.TodaySwept  = 0
        storedDate = today
    End If

    Select Case kind
        Case "routed"
            snap.TodayRouted = snap.TodayRouted + 1
            snap.LifeRouted  = snap.LifeRouted  + 1
        Case "swept"
            snap.TodaySwept = snap.TodaySwept + 1
            snap.LifeSwept  = snap.LifeSwept  + 1
    End Select

    SaveStatsFile storedDate, snap
    Exit Sub
Fail:
    ' Cannot let a stats failure surface to the caller's mail.Move — silent exit.
End Sub

' Load stats.psv into out-params. Leaves out-params at zero/empty defaults
' if the file is missing or cannot be parsed.
Private Sub LoadStatsFile(outDate As String, outSnap As StatsSnapshot)
    On Error GoTo Fail

    Dim filePath As String
    filePath = StatsFilePath()
    If Dir(filePath) = "" Then Exit Sub

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
            Dim parts() As String
            parts = Split(line, "|")
            If UBound(parts) >= COL_COUNT - 1 Then
                outDate            = Trim(parts(0))
                outSnap.TodayRouted = CLng(Trim(parts(1)))
                outSnap.TodaySwept  = CLng(Trim(parts(2)))
                outSnap.LifeRouted  = CLng(Trim(parts(3)))
                outSnap.LifeSwept   = CLng(Trim(parts(4)))
            End If
        End If
    Loop

    Close #fnum
    Exit Sub
Fail:
    If fnum > 0 Then Close #fnum
End Sub

' Overwrite stats.psv with the supplied values.
Private Sub SaveStatsFile(storedDate As String, snap As StatsSnapshot)
    On Error GoTo Fail

    Dim fnum As Integer
    fnum = FreeFile
    Open StatsFilePath() For Output As #fnum
    Print #fnum, STATS_HEADER
    Print #fnum, storedDate & "|" & _
                 snap.TodayRouted & "|" & _
                 snap.TodaySwept  & "|" & _
                 snap.LifeRouted  & "|" & _
                 snap.LifeSwept
    Close #fnum
    Exit Sub
Fail:
    If fnum > 0 Then Close #fnum
End Sub
