# Outlook-Newforma Email Filing Automation

A VBA automation that routes project email into the correct Newforma folders using dynamically created matching terms (project names and numbers), eliminating the need to create new rules through the traditional Outlook rule manager for each project. 

## Overview

Two workflows:

- 1 — Folder Detection: Scans the Newforma "Items to File" folder structure, reconciles it against a configuration file, and keeps the config in sync as projects are added or archived.
- 2 — Staging & Release: Routes incoming mail to a per-project staging folder, then releases it to the Newforma destination once the item has been read and any follow-up flag cleared.

## Workflow

1. Email arrives in the Inbox.
2. The script checks the subject line against active projects in the config file.
3. A match routes the email to that project's staging folder.
4. A timed sweep checks staged items — once an item is **read** and **not actively flagged**, it moves to the project's Newforma "Items to File" folder.
5. Newforma files it to the project record on the next Synchronize.

## Configuration

Project routing is controlled by a pipe-delimited text file (`|`). Key columns:

| Column | Description |
|---|---|
| `project_no` | Project number — the default match key |
| `project_name` | Project name |
| `match_terms` | Semicolon-separated terms searched in the subject line (edit to add aliases) |
| `staging_path` | Path to the project's staging folder |
| `destination_path` | Path to the Newforma "Items to File" folder |
| `status` | `active` / `archived` — set automatically by folder presence |
| `enabled` | `TRUE` / `FALSE` — manual on/off switch, never overwritten by the engine |
| `last_seen` | Last date the source folder was observed (ISO 8601) |

Mail is only routed for projects where `status = active` **and** `enabled = TRUE`.

## Requirements

- Outlook Classic (required by the Newforma legacy COM add-in)
- Newforma Project Center with the legacy add-in installed
- VBA macros enabled (confirm macro security policy with IT before deploying)

### Newforma setup

The automation discovers projects through the **Newforma - Items to File** folder that the legacy add-in creates in your mailbox. For a project's staging and routing to work, the project must appear as a sub-folder under that parent, which requires it to be on your **My Projects** tab in Project Center. If a project is not showing up in the config file after a reconciliation run, check that it has been added to My Projects and that Newforma has synchronised, creating its folder in Outlook.

## Module structure

| Module | Responsibility |
|---|---|
| `ConfigStore` | Read/write the pipe-delimited config file |
| `FolderScanner` | Resolve the Newforma parent, enumerate child folders, apply the false-stale guard |
| `Reconciler` | Three-way merge of live folders vs config; creates staging folders |
| `Router` | Arrival routing — matches subject to staging folder |
| `Sweeper` | Timed release of read/unflagged items; Inbox backstop scan |
| `Common` | Shared folder-resolution and file-based logging helpers |
| `Timer` | Win32 SetTimer plumbing that fires the periodic sweep tick (VBA6/VBA7 conditional compile) |
| `ThisOutlookSession` | Entry points: startup, new-mail event, folder-added event, timer tick |

## Limitations

- Subject-line matching only: mail that never names the project in the subject is not routed automatically.
- Similar to traditional Outlook rules; per-project misfires are fixed by editing `match_terms` for that row.
- Only works on your machine.
- Dependent on the Newforma legacy add-in's folder structure; would not work on migration to the Newforma HTML5 add-in.

## Setup
0. Ensure macros are allowed through Outlook Trust Center.
1. In VBA for Applications, use File > Import (Ctrl-M) to import seven files: Common.bas, ConfigStore.bas, FolderScanner.bas, Reconciler.bas, Router.bas, Timer.bas, Sweeper.bas
2. Insert two Class Modules called 'FolderWatcher' and 'ProjectConfig', making sure that the names under properties match. Copy the code from FolderWatcher.cls and ProjectConfig.cls
3. Paste the code from ThisOutlookSession_additions.bas into ThisOutlookSession
4. Restart Outlook. The Reconciler will look for all folders under 'Newforma - Items to File' and create a duplicate structure under 'Project Staging'. Or run "RunReconciliation" in the Immediate Window
5. Configure the match_terms in %AppData%\OutlookAutoFile\projects.psv


## Future improvements

- **Config file UI:** A user-friendly interface for viewing and editing the `.psv` config file — toggling projects on/off, editing match terms, and reviewing routing history — without needing to open a text file directly.
- **Analogous terms expansion:** Automatically seed `match_terms` with common address abbreviation variants when a project name is first detected (e.g. Street → St., Square → Sq., Avenue → Ave., Road → Rd.), so address-based project names match regardless of how correspondents abbreviate them.
- **VSTO port:** The module structure is designed for a straightforward C# VSTO port. The configuration file format is language-agnostic and can be shared across both versions during a transition. Future implementation.
