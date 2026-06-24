Attribute VB_Name = "Router"
Option Explicit

' Route all qualifying MailItems currently in the Inbox to their staging
' folders. Called on ItemAdd and by the sweep backstop.
Public Sub RouteInboxItems()
    On Error GoTo Fail

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    Dim inbox As Outlook.Folder
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    Dim rows As Collection
    Set rows = LoadConfig(ConfigFilePath())
    If rows Is Nothing Then
        LogWarn "Router: config unreadable - skipping"
        Exit Sub
    End If

    ' Walk backwards so moving an item does not disturb the index.
    Dim i As Integer
    For i = inbox.Items.Count To 1 Step -1
        Dim obj As Object
        Set obj = inbox.Items(i)
        If TypeOf obj Is Outlook.MailItem Then
            RouteItem obj, rows
        End If
    Next i
    Exit Sub
Fail:
    LogError "Router: RouteInboxItems - " & Err.Description
End Sub

' Check one MailItem against all active+enabled projects and move it to the
' first matching project's staging folder. First match in file order wins.
Public Sub RouteItem(mail As Outlook.MailItem, rows As Collection)
    On Error GoTo Fail

    Dim cfg As ProjectConfig
    For Each cfg In rows
        If LCase(cfg.Status) = "active" And LCase(cfg.Enabled) = "true" Then
            If SubjectMatches(mail.Subject, cfg.MatchTerms) Then
                Dim dest As Outlook.Folder
                Set dest = ResolveFolder(cfg.StagingPath)
                If Not dest Is Nothing Then
                    LogInfo "Router: moving """ & mail.Subject & """ to staging for " & cfg.ProjectNo
                    mail.Move dest
                    RecordRouted
                Else
                    LogWarn "Router: staging folder not found for " & cfg.ProjectNo
                End If
                Exit Sub   ' First match wins - stop checking further projects.
            End If
        End If
    Next cfg
    Exit Sub
Fail:
    LogError "Router: RouteItem - " & Err.Description
End Sub

' --- Private helpers --------------------------------------------------------

' Return True if subject contains any semicolon-separated term in matchTerms
' (case-insensitive substring match).
Private Function SubjectMatches(subject As String, matchTerms As String) As Boolean
    Dim terms() As String
    terms = Split(matchTerms, ";")

    Dim lSubject As String
    lSubject = LCase(subject)

    Dim term As Variant
    For Each term In terms
        Dim t As String
        t = LCase(Trim(CStr(term)))
        If t <> "" Then
            If InStr(1, lSubject, t) > 0 Then
                SubjectMatches = True
                Exit Function
            End If
        End If
    Next term

    SubjectMatches = False
End Function
