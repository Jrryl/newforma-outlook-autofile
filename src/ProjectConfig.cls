VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProjectConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Detector-owned (written by Workflow 1; never edited by hand).
Public ProjectNo       As String
Public ProjectName     As String
Public StagingPath     As String
Public DestinationPath As String
Public Status          As String   ' "active" | "archived"
Public LastSeen        As String   ' ISO 8601 date, e.g. "2026-06-17"

' User-owned (seeded on first detection; never overwritten by the engine).
Public MatchTerms      As String   ' semicolon-separated; seeded with ProjectNo
Public Enabled         As String   ' "TRUE" | "FALSE"
