# EXRESearcher

PowerShell GUI for Exchange mailbox content search, eDiscovery, and organization-wide message cleanup. Built for Exchange 2019 SE.

## Features

| Tab | Description |
|-----|-------------|
| **Mailbox Search** | Search-Mailbox with KQL filters (subject, from, to, keywords, attachment, date range). Actions: Estimate, Log, Copy to mailbox, Delete |
| **Org-Wide Delete** | Search and delete messages across ALL mailboxes. Batch processing, double confirmation, safety checks. For phishing/malware cleanup |
| **eDiscovery** | In-Place eDiscovery management: create, monitor, stop, remove compliance searches |
| **Mailboxes** | Browse mailboxes with filters, view statistics, folder breakdown, quick "Use in Search" |
| **Audit Log** | Full audit trail of all search and delete operations |

**Async Architecture** — all Exchange operations run in PowerShell runspaces. GUI never freezes.

## Quick Start

### Requirements
- Windows PowerShell 5.1+
- Exchange 2019 SE (Mailbox role)
- RBAC roles: `Mailbox Search`, `Discovery Management`, `Mailbox Import Export`

### Launch
```powershell
.\EXRESearcher.ps1
```

1. Enter Exchange server FQDN → **Connect**
2. Go to **Mailbox Search** tab
3. Fill in filters (Subject, From, Date range, etc.)
4. Enter mailbox(es) in **Scope** field
5. Click **Estimate** to see how many items match
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

## Search Query (KQL) Syntax

| Filter | KQL | Example |
|--------|-----|---------|
| Subject | `subject:"text"` | `subject:"invoice"` |
| From | `from:"email"` | `from:"user@domain.com"` |
| To | `to:"email"` | `to:"ceo@company.com"` |
| Attachment | `attachment:"name"` | `attachment:"report.xlsx"` |
| MessageId | `messageid:"id"` | `messageid:"<abc@domain>"` |
| Date range | `received>=date` | `received>=2026-01-01` |
| Keywords | free text | `confidential secret` |
| Combined | `AND` | `subject:"test" AND from:"user@dom"` |

## Architecture

```
┌──────────────────────────────────────────────────┐
│ WinForms UI Thread                                │
│  ┌──────────┐  ┌───────────┐  ┌──────────┐      │
│  │ Search    │  │ Org-Wide  │  │ eDiscovery│      │
│  │ Filters   │  │ Delete    │  │ Manager   │      │
│  └─────┬─────┘  └─────┬─────┘  └─────┬─────┘    │
│        │               │               │          │
│        v               v               v          │
│   Start-AsyncJob  Start-AsyncJob  Start-AsyncJob  │
└────────┬───────────────┬───────────────┬──────────┘
         │               │               │
         v               v               v
   ┌───────────┐  ┌───────────┐  ┌───────────┐
   │ Runspace  │  │ Runspace  │  │ Runspace  │
   │ Search-   │  │ Batch     │  │ New/Get-  │
   │ Mailbox   │  │ Delete    │  │ Mailbox   │
   │           │  │ (50/batch)│  │ Search    │
   └───────────┘  └───────────┘  └───────────┘

┌──────────────────────────────────────────────────┐
│ Job Console (bottom panel)                        │
│ [12:30:01] START  #1 Search (Estimate)           │
│ [12:30:03] DONE   #1 Search (Estimate) (2.1s)   │
│ [12:30:05] START  #2 OrgWide Estimate            │
│ ████████████████░░░░ Running: 1 job(s)           │
└──────────────────────────────────────────────────┘
```

## Safety Features

- **Wildcard block**: `*` query on Org-Wide tab is rejected
- **Double confirmation**: Org-wide delete requires two Yes/No prompts
- **Audit trail**: Every operation logged to `%APPDATA%\EXRESearcher\`
  - `search-audit.csv` — all searches with queries, scopes, results
  - `operator-log.csv` — all operator actions (connect, delete, etc.)
- **Estimate first**: Always estimate before deleting
- **Batch processing**: Org-wide operations process 50 mailboxes at a time
- **Permission check**: Verify RBAC roles before starting

## Project Structure

```
EXRESearcher/
├── EXRESearcher.ps1        # GUI (WinForms) — 5 tabs, async
├── lib/
│   ├── Core.ps1            # Exchange functions (search, eDiscovery, org-wide)
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
| `New-MailboxSearch` | Create In-Place eDiscovery search |
| `Start-MailboxSearch` | Start eDiscovery search |
| `Get-MailboxSearch` | Get search status/results |
| `Stop-MailboxSearch` | Stop running search |
| `Remove-MailboxSearch` | Delete search |
| `Get-Mailbox` | Enumerate mailboxes |
| `Get-MailboxStatistics` | Mailbox size/item counts |
| `Get-MailboxFolderStatistics` | Per-folder breakdown |
| `Get-MailboxPermission` | Access rights audit |
| `Get-ManagementRoleAssignment` | Check RBAC permissions |
| `Get-ExchangeServer` | Server version info |
| `Get-DistributionGroupMember` | Expand groups for search scope |

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| **F5** | Refresh current tab |
| **Ctrl+E** | Export current tab data |

## Version

**1.0.0** — Mailbox content search (Search-Mailbox), org-wide search & delete, In-Place eDiscovery, mailbox browser with stats, full audit logging. Async architecture with runspaces.
