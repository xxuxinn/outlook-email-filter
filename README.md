# Outlook Email Filter v2.0

A VBA-based email filtering system for Microsoft Outlook that automatically classifies and filters emails using rule-based patterns with optional LLM integration.

## What's New in v2.0

- **Automated Installation**: One-click PowerShell installer or VBA self-installer
- **External Configuration**: All settings in `settings.ini` — no more editing VBA code
- **Readable Folder Names**: Configurable folder names (e.g., "Protected" instead of "II")
- **LLM Email Tools**: Summarize emails and draft replies using Azure OpenAI
- **Migration Helper**: Automatically rename old v1.x folders to v2.0 names
- **Dashboard UI** (optional): 4-tab graphical UserForm for all filter operations
- **Portability**: Export all VBA modules to Desktop with one click

## Features

- **Real-time filtering**: Automatically process new emails as they arrive
- **Batch processing**: Filter existing emails in bulk
- **Rule-based classification**: Fast, offline-capable pattern matching
- **Self-improving**: Learns from your manual sorting (drag to LearnKeep/LearnDelete/LearnSubjectDelete folders)
- **LLM integration**: Optional Azure OpenAI for ambiguous emails, summarization, and reply drafting
- **Dry-run mode**: Preview decisions before taking action
- **Server rule export**: Push learned DELETE rules to Exchange for 24/7 filtering
- **Highly configurable**: Edit `settings.ini` with any text editor

## Quick Start

### Automated Installation (Recommended) ⚡

**Option 1: PowerShell Installer** (easiest)
```powershell
.\Install-OutlookFilter.ps1
```
Or double-click `install.bat` in Windows Explorer.

**Option 2: VBA Self-Installer**
1. Import `src/Installer.bas` into Outlook VBA (Alt+F11 -> File -> Import)
2. Run `InstallEmailFilter` in the Immediate Window (Ctrl+G)

See [INSTALL-QUICK.md](INSTALL-QUICK.md) for details.

### Manual Installation

1. **Enable macros** in Outlook (File -> Options -> Trust Center -> Macro Settings)
2. **Import VBA modules** (Alt+F11 -> File -> Import): Config.bas, Utilities.bas, EmailFilter.bas, BatchFilter.bas
3. **Paste ThisOutlookSession** code into the built-in module (see [docs/INSTALL.md](docs/INSTALL.md))
4. **Compile**: Debug -> Compile Project -> Ctrl+S
5. **Restart Outlook** — `settings.ini` is auto-created with defaults
6. **Create learning folders** under Inbox: LearnKeep, LearnDelete, LearnSubjectDelete
7. **Test**: In Immediate Window (Ctrl+G), type `FilterExistingDryRun`

See [docs/INSTALL.md](docs/INSTALL.md) for detailed manual instructions, or [docs/USER_MANUAL.md](docs/USER_MANUAL.md) for the complete reference.

**Upgrading from v1.x?** Run `DetectAndMigrateOldFolders` to rename I/II/III/IIII/V to readable names.

## Classification Rules

| Priority | Rule | Action |
|----------|------|--------|
| 0 | **Learned sender rule** (self-improving) | Keep or Delete |
| 0.5 | **Learned subject rule** (self-improving) | Delete |
| 1 | Protected domain (e.g., substack.com) | Move to Protected |
| 2 | Personally addressed (name/greeting) | Keep |
| 3 | Organizational tags ([MM], [HRO]) | Keep |
| 4 | VIP keywords (thesis, deadline) | Keep |
| 5 | Reply/Forward (RE:, FW:) | Keep |
| 6 | Known spam senders | Delete |
| 7 | Spam sender patterns (noreply) | Delete |
| 8 | Spam subject keywords (newsletter) | Delete |
| 9 | No match | Review folder (or LLM) |

## Files

```
outlook-email-filter/
├── src/
│   ├── Config.bas                 # Default constants and runtime variables
│   ├── Utilities.bas              # Helpers, caching, INI reader/writer
│   ├── EmailFilter.bas            # Classification logic + LLM tools
│   ├── BatchFilter.bas            # Bulk processing + macro launcher
│   ├── ThisOutlookSession.bas     # Event handlers (copy-paste, not import)
│   ├── frmFilterDashboard.frm     # Dashboard UserForm (optional)
│   └── frmDraftReply.frm          # LLM draft reply UserForm (optional)
├── docs/
│   ├── INSTALL.md                 # Installation guide
│   ├── USER_MANUAL.md             # Complete user manual
│   ├── PATTERNS.md                # Pattern configuration guide
│   └── UPDATE_GUIDE_*.md          # Update guides for previous versions
├── CLAUDE.md                      # Developer reference
└── README.md
```

## Configuration

All settings are stored in `%APPDATA%\OutlookEmailFilter\settings.ini` (auto-created on first run).

Edit with any text editor. To reset: delete settings.ini and restart Outlook.

### Key Macros (Immediate Window, Ctrl+G)

| Macro | What It Does |
|-------|-------------|
| `FilterExistingDryRun` | Preview what the filter would do (no changes) |
| `FilterExistingEmails` | Filter all Inbox emails |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation |
| `FilterCurrentFolder` | Filter current folder with confirmation |
| `DetectAndMigrateOldFolders` | Rename v1.x folders to v2.0 names |
| `ShowVersionInfo` | Display version, paths, and status |

See [docs/USER_MANUAL.md](docs/USER_MANUAL.md) for the full macro list.

## Self-Improving Filter

The filter learns from your manual sorting:

1. **Drag an email to "LearnKeep"** -> Always **keep** future emails from that sender
2. **Drag an email to "LearnDelete"** -> Always **delete** future emails from that sender
3. **Drag an email to "LearnSubjectDelete"** -> Always **delete** future emails with matching subjects

Learned rules are stored in `%APPDATA%\OutlookEmailFilter\` and persist across Outlook restarts.

## LLM Integration (Optional)

Configure Azure OpenAI in `settings.ini` under `[LLM]` section:
- **Classification**: Ambiguous emails are classified by LLM instead of going to Review folder
- **Summarize**: Run `SummarizeSelectedEmail` on a selected email
- **Draft Reply**: Run `DraftReplyToSelected` to get an AI-drafted response

## Dashboard (Optional)

A 4-tab graphical UserForm is available for users who prefer a GUI. Since UserForms require manual creation in the VBA Editor, this is optional — all functionality is accessible via macros and `settings.ini`.

See [docs/INSTALL.md](docs/INSTALL.md) for UserForm setup instructions.

## Requirements

- Microsoft Outlook (desktop)
- Outlook 2016, 2019, 2021, or Microsoft 365
- Windows (VBA macros not supported on Mac Outlook)

## License

MIT License - Use freely, modify as needed.
