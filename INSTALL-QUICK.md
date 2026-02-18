# Quick Installation Guide

This guide shows the easiest ways to install the Outlook Email Filter v2.0.

## Method 1: PowerShell Installer (Recommended) ⚡

The fastest installation method using PowerShell automation.

### Steps:

1. **Close Outlook** (very important!)

2. **Open PowerShell** in this folder:
   - Right-click the folder in Explorer
   - Choose "Open in Terminal" or "Open PowerShell window here"

3. **Run the installer:**
   ```powershell
   .\Install-OutlookFilter.ps1
   ```

4. **Follow the prompts** - The script will:
   - ✓ Check Outlook version
   - ✓ Configure Trust Center settings
   - ✓ Import all VBA modules
   - ✓ Create necessary folders
   - ✓ Initialize settings.ini

5. **Open Outlook**, then:
   - Press `Alt+F11` to open VBA Editor
   - Go to `Debug` → `Compile Project`
   - Press `Ctrl+S` to save
   - Close VBA Editor
   - Restart Outlook

6. **Test**: Press `Alt+F11` → `Ctrl+G` → type `FilterExistingDryRun`

**That's it!** 🎉

### Alternative: Run from Windows Explorer

Double-click: **`install.bat`** (runs the PowerShell script with execution policy bypass)

---

## Method 2: VBA Self-Installer

If you prefer staying in VBA or the PowerShell method doesn't work for you.

### Steps:

1. **Enable macros** in Outlook:
   - File → Options → Trust Center → Trust Center Settings
   - Macro Settings → Choose "Enable all macros"
   - **Check**: "Trust access to the VBA project object model" ⚠️ (critical!)
   - Click OK twice

2. **Restart Outlook**

3. **Open VBA Editor**: Press `Alt+F11`

4. **Import installer**:
   - File → Import File
   - Navigate to `src/Installer.bas`
   - Click Open

5. **Run the installer**:
   - Press `Ctrl+G` to open Immediate Window
   - Type: `InstallEmailFilter`
   - Press Enter

6. **Select source folder**:
   - Browse to the `src` folder in this repository
   - Click OK

7. **Wait for completion** - The installer will:
   - Import all modules
   - Set up ThisOutlookSession
   - Create folders
   - Initialize settings.ini

8. **Compile and save**:
   - Debug → Compile Project
   - Press `Ctrl+S`
   - Close VBA Editor

9. **Restart Outlook**

10. **Test**: `Alt+F11` → `Ctrl+G` → `FilterExistingDryRun`

---

## Method 3: Manual Installation

See [docs/INSTALL.md](docs/INSTALL.md) for the traditional step-by-step manual installation.

---

## Troubleshooting

### PowerShell says "execution policy" error

Run PowerShell as Administrator and execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Or use the provided `install.bat` which bypasses this automatically.

### "Trust access to VBA project object model" error

1. Open Outlook
2. File → Options → Trust Center → Trust Center Settings
3. Macro Settings tab
4. **Check** "Trust access to the VBA project object model"
5. Click OK, restart Outlook, try again

### Outlook is running error

Close Outlook completely (check system tray for hidden Outlook icon). Then run the installer again.

### Compile errors after installation

Go to VBA Editor → Tools → References:
- Uncheck any items marked "MISSING"
- Ensure "Microsoft Outlook XX.X Object Library" is checked

---

## Uninstallation

### PowerShell method:
```powershell
.\Install-OutlookFilter.ps1 -Uninstall
```

### VBA method:
In Immediate Window, type: `UninstallEmailFilter`

**Note**: Data files (settings.ini, learned rules) are NOT deleted during uninstall.

---

## What Gets Installed?

### VBA Modules (imported into Outlook VBA project):
- `Config.bas` - Configuration constants and runtime variables
- `Utilities.bas` - Helper functions, caching, INI reader/writer
- `EmailFilter.bas` - Core classification engine
- `BatchFilter.bas` - Bulk operations and macros
- `ThisOutlookSession` - Event handlers (auto-filter on new mail)
- `frmFilterDashboard.frm` - Optional Dashboard UI
- `frmDraftReply.frm` - Optional LLM reply UI

### Folders (created under Inbox):
- **Review** - Ambiguous emails for manual review
- **Protected** - Protected domain emails
- **LearnKeep** - Drag here to teach "always keep this sender"
- **LearnDelete** - Drag here to teach "always delete this sender"
- **LearnSubjectDelete** - Drag here to teach "delete by subject match"

### Data Files (created at `%APPDATA%\OutlookEmailFilter\`):
- `settings.ini` - All configuration (patterns, folders, LLM settings)
- `learned_senders.txt` - Self-improving sender rules (auto-created when you use learning folders)
- `learned_subjects.txt` - Self-improving subject rules (auto-created when you use learning folders)

---

## Next Steps

After installation:

1. **Customize patterns**: Edit `%APPDATA%\OutlookEmailFilter\settings.ini`
   - Update `NamePatterns` with your name
   - Add your important domains to `ProtectedDomains`
   - Adjust VIP keywords, spam patterns, etc.

2. **Test filtering**: In VBA Immediate Window (`Ctrl+G`):
   ```vba
   FilterExistingDryRun  ' Preview what would happen (no changes)
   FilterSelectedEmail   ' Test on one selected email
   FilterExistingEmails  ' Actually filter your Inbox
   ```

3. **Add Quick Access buttons** (optional):
   - File → Options → Quick Access Toolbar
   - Choose "Macros" from dropdown
   - Add: `Project1.FilterSelectedEmails`, `Project1.FilterCurrentFolder`

4. **Enable real-time filtering** (optional):
   - In VBA Editor, open `ThisOutlookSession`
   - Find `Private Sub inboxItems_ItemAdd` (commented out)
   - Uncomment the entire Sub (remove `'` from each line)
   - Save and restart Outlook

5. **Configure LLM** (optional):
   - Edit settings.ini `[LLM]` section
   - Set `UseLLMAPI=true`
   - Configure Azure OpenAI endpoint and API key

---

## Support

For issues and questions:
- Check [docs/USER_MANUAL.md](docs/USER_MANUAL.md) for detailed documentation
- Review [docs/INSTALL.md](docs/INSTALL.md) for troubleshooting
- See [CLAUDE.md](CLAUDE.md) for developer reference
