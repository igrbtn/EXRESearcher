# EXRESearcher

PowerShell GUI for Exchange mailbox content search, eDiscovery, and organization-wide message cleanup. Built for Exchange 2019 SE.

## Features

| Tab | Description |
|-----|-------------|
| **Mailbox Search** | Search-Mailbox with KQL filters (subject, from, to, keywords, attachment, date range). Actions: Estimate, Log, Copy to mailbox, Delete. **Double-click** results to preview individual messages via EWS |
| **Org-Wide Delete** | Search and delete messages across ALL mailboxes. Batch processing, double confirmation, safety checks. **Double-click** to preview messages. For phishing/malware cleanup |
| **eDiscovery** | In-Place eDiscovery management: create, monitor, stop, remove compliance searches |
| **Mailboxes** | Browse mailboxes with filters, view statistics, folder breakdown, quick "Use in Search" |
| **Folder Cleanup** | Folder-level search & delete with filters (age, sender, subject, size, attachments). Dumpster purge. Duplicate detection & backup-and-delete |
| **Audit Log** | Full audit trail of all search and delete operations |

**Async Architecture** — all Exchange operations run in PowerShell runspaces. GUI never freezes.

**EMS Auto-Detection** — when launched from Exchange Management Shell, automatically discovers servers and connects.

## Quick Start

### Requirements
- Windows PowerShell 5.1+
- Exchange 2019 SE (Mailbox role)
- RBAC roles: `Mailbox Search`, `Discovery Management`, `Mailbox Import Export`
- For message preview: `ApplicationImpersonation` role (EWS)

### Launch
```powershell
# From Exchange Management Shell (recommended — auto-detects servers)
.\EXRESearcher.ps1

# Or from regular PowerShell (manual server entry)
.\EXRESearcher.ps1
```

**If launched from EMS**: automatically discovers Exchange servers, picks the local one, and connects — ready to use immediately.

**If launched from regular PowerShell**: enter Exchange server FQDN → click **Connect**.

1. Go to **Mailbox Search** tab
2. Fill in filters (Subject, From, Date range, etc.)
3. Enter mailbox(es) in **Scope** field
4. Click **Estimate** to see how many items match
5. **Double-click** a result row to preview individual messages (Subject, From, To, Date, Size)
6. Use **Search + Log**, **Copy to Mailbox**, or **Search + Delete**

### Check Permissions
Click **Check Permissions** button to verify your RBAC roles and find the Discovery Mailbox.

## Use Cases

### 1. Phishing Cleanup
```
1. Org-Wide Delete tab
2. Subject: "Your account has been compromised"
3. From: attacker@evil.com
4. Date: today
5. Click "Estimate All" → see affected mailboxes
6. Click "DELETE FROM ALL" → double confirmation → done
```

### 2. eDiscovery / Legal Hold
```
1. eDiscovery tab
2. Name: "Case-2026-001"
3. Query: from:"suspect@company.com" AND received>=2026-01-01
4. Check "All Mailboxes" or enter specific ones
5. Create Search → monitor status
```

### 3. Find & Export Evidence
```
1. Mailbox Search tab
2. Subject: "confidential" Keywords: "project alpha"
3. Scope: user@company.com
4. Target Mailbox: discovery@company.com
5. Click "Copy to Mailbox" → results copied to target
```

### 4. Mailbox Forensics
```
1. Mailboxes tab → Load → select mailbox
2. "Get Stats" → item counts, sizes, last logon
3. "Folder Stats" → per-folder breakdown
4. "Use in Search" → jump to Search tab with mailbox pre-filled
```

### 5. Folder Cleanup
```
1. Folder Cleanup tab
2. Enter mailbox → "Load Folders" → select folder
3. Set filters: Subject, From, Older than X days, Size, Has Attachment
4. Click "Estimate" → see matching items
5. Click "Delete" → confirmation → items removed
```

### 6. Dumpster Purge
```
1. Folder Cleanup tab
2. Enter mailbox
3. Click "Purge Dumpster" → Yes to estimate, No to delete immediately
4. Permanently removes recoverable items
```

### 7. Duplicate Detection & Cleanup
```
1. Folder Cleanup tab
2. Enter mailbox → Load Folders
3. Click "Find Duplicates" → scans folders for possible duplicates
4. Select folder + enter Target mailbox
5. "Backup Folder" → copies content to target (safe)
6. "Backup + Delete" → copies then deletes from source (requires confirmation)
```

## Search Query (KQL) Syntax

| Filter | KQL | Example |
|--------|-----|---------|
| Subject | `subject:"text"` | `subject:"invoice"` |
| From | `from:"email"` | `from:"user@domain.com"` |
| To | `to:"email"` | `to:"ceo@company.com"` |
| Attachment | `attachment:"name"` | `attachment:"report.xlsx"` |
| MessageId | `messageid:"id"` | `messageid:"<abc@domain>"` |
| Date range | `received>=date` | `received>=2026-01-01` |
| Folder | `folder:"path"` | `folder:"Inbox"` |
| Size | `size>=value` | `size>=10MB` |
| Has attachment | `hasattachment:true` | `hasattachment:true` |
| Keywords | free text | `confidential secret` |
| Combined | `AND` | `subject:"test" AND from:"user@dom"` |

## Architecture

```
┌──────────────────────────────────────────────────────┐
│ WinForms UI Thread                                    │
│  ┌──────────┐  ┌───────────┐  ┌──────────┐          │
│  │ Search    │  │ Org-Wide  │  │ eDiscovery│          │
│  │ Filters   │  │ Delete    │  │ Manager   │          │
│  └─────┬─────┘  └─────┬─────┘  └─────┬─────┘        │
│  ┌──────────┐  ┌───────────┐                          │
│  │ Folder   │  │ Mailbox   │                          │
│  │ Cleanup  │  │ Browser   │                          │
│  └─────┬─────┘  └─────┬─────┘                        │
│        │               │                              │
│        v               v                              │
│   Start-AsyncJob  Start-AsyncJob                      │
└────────┬───────────────┬─────────────────────────────┘
         │               │
         v               v
   ┌───────────┐  ┌───────────┐
   │ Runspace  │  │ Runspace  │
   │ Search-   │  │ Folder    │
   │ Mailbox   │  │ Cleanup   │
   │           │  │ Dumpster  │
   └───────────┘  └───────────┘

┌──────────────────────────────────────────────────────┐
│ Job Console (bottom panel)                            │
│ [12:30:01] START  #1 Search (Estimate)               │
│ [12:30:03] DONE   #1 Search (Estimate) (2.1s)       │
│ [12:30:05] START  #2 Folder Delete user@co           │
│ ████████████████░░░░ Running: 1 job(s)               │
└──────────────────────────────────────────────────────┘
```

## Safety Features

- **Wildcard block**: `*` query on Org-Wide tab is rejected
- **Double confirmation**: Org-wide delete requires two Yes/No prompts
- **Folder delete confirmation**: Folder cleanup deletes require confirmation dialog
- **Backup before delete**: Duplicate cleanup always backs up to target mailbox first
- **Audit trail**: Every operation logged to `%APPDATA%\EXRESearcher\`
  - `search-audit.csv` — all searches with queries, scopes, results
  - `operator-log.csv` — all operator actions (connect, delete, etc.)
- **Estimate first**: Always estimate before deleting
- **Batch processing**: Org-wide operations process 50 mailboxes at a time
- **Permission check**: Verify RBAC roles before starting

## Project Structure

```
EXRESearcher/
├── EXRESearcher.ps1        # GUI (WinForms) — 6 tabs, async
├── lib/
│   ├── Core.ps1            # Exchange functions (search, eDiscovery, org-wide, folder cleanup, duplicates)
│   ├── Settings.ps1        # Settings, cache, operator audit log
│   └── AsyncRunner.ps1     # Async framework (runspaces, progress, job console)
├── tests/
│   └── EXRESearcher.Tests.ps1
├── CLAUDE.md
├── .env.example
└── .gitignore
```

## Exchange Cmdlets Used

| Cmdlet | Purpose |
|--------|---------|
| `Search-Mailbox` | Content search with estimate/log/copy/delete |
| `Search-Mailbox -SearchDumpsterOnly` | Recoverable items (dumpster) purge |
| `New-MailboxSearch` | Create In-Place eDiscovery search |
| `Start-MailboxSearch` | Start eDiscovery search |
| `Get-MailboxSearch` | Get search status/results |
| `Stop-MailboxSearch` | Stop running search |
| `Remove-MailboxSearch` | Delete search |
| `Get-Mailbox` | Enumerate mailboxes |
| `Get-MailboxStatistics` | Mailbox size/item counts |
| `Get-MailboxFolderStatistics` | Per-folder breakdown, folder list |
| `Get-MailboxPermission` | Access rights audit |
| `Get-ManagementRoleAssignment` | Check RBAC permissions |
| `Get-ExchangeServer` | Server version info, auto-discovery |
| `Get-DistributionGroupMember` | Expand groups for search scope |
| EWS `FindItem` | Message preview (Subject, From, To, Date, Size) |

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **F5** | Refresh current tab |
| **Ctrl+E** | Export current tab data |
| **Double-click** | Preview messages in search result row |

## Version

**1.2.0** — EMS auto-detection and server discovery; connection guard on all operations; EWS message preview on double-click (Subject, From, To, Date, Size); Exchange snap-in loaded in async runspaces; UTF-8 BOM for PowerShell 5.1 compatibility.

**1.1.0** — Folder cleanup tab: folder-level search & delete with filters (age, sender, subject, size, attachments), dumpster purge (recoverable items), duplicate detection, backup-and-delete workflow. Extended KQL with folder, size, hasattachment filters.

**1.0.0** — Mailbox content search (Search-Mailbox), org-wide search & delete, In-Place eDiscovery, mailbox browser with stats, full audit logging. Async architecture with runspaces.
