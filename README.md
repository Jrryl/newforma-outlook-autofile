# Outlook Email Filing Automation

A VBA automation that routes project email into the correct Newforma folders with minimal manual rule maintenance.

## Overview

Two workflows work together:

- **Workflow 1 — Folder Detection:** Scans the Newforma "Items to File" folder structure, reconciles it against a configuration file, and keeps the config in sync as projects are added or archived.
- **Workflow 2 — Staging & Release:** Routes incoming mail to a per-project staging folder, then releases it to the Newforma destination once the item has been read and any follow-up flag cleared.

## How it works

1. Email arrives in the Inbox.
2. The engine checks the subject line against active projects in the config file.
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

## Module structure

| Module | Responsibility |
|---|---|
| `ConfigStore` | Read/write the pipe-delimited config file |
| `FolderScanner` | Resolve the Newforma parent, enumerate child folders, apply the false-stale guard |
| `Reconciler` | Three-way merge of live folders vs config; creates staging folders |
| `Router` | Arrival routing — matches subject to staging folder |
| `Sweeper` | Timed release of read/unflagged items; Inbox backstop scan |
| `Common` | Shared folder-resolution and file-based logging helpers |
| `ThisOutlookSession` | Entry points: startup, new-mail event, folder-added event, timer tick |

## Limitations

- Subject-line matching only — mail that never names the project in the subject is not routed automatically.
- ~90% accuracy accepted; per-project misfires are fixed by editing `match_terms` for that row.
- Single-user, single-machine (a property of VBA).
- Dependent on the Newforma legacy add-in's folder structure; a migration to the Newforma HTML5 add-in would remove that structure.

## Future: VSTO port

The module structure is designed for a straightforward C# VSTO port. The configuration file format is language-agnostic and can be shared across both versions during a transition. See the design specification for full portability notes.
