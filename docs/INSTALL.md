# Installation Guide v2.0

This guide walks you through setting up the Outlook Email Filter VBA macros.

## Prerequisites

- Microsoft Outlook (desktop application, not web)
- Microsoft 365 or Outlook 2016/2019/2021
- Administrator rights (for enabling macros)

## Phase 1: Enable Macros

1. Open **Outlook**
2. Go to **File -> Options -> Trust Center**
3. Click **Trust Center Settings...**
4. Select **Macro Settings** from the left panel
5. Choose one of these options:
   - **"Enable all macros"** (easiest, less secure)
   - **"Notifications for digitally signed macros"** (if you plan to sign your code)
6. Click **OK** twice to close
7. **Restart Outlook**

> **Security Note**: "Enable all macros" is convenient for personal use but not recommended in shared/corporate environments.

## Phase 2: Import VBA Code

### Step 1: Open VBA Editor

1. In Outlook, press **Alt + F11** to open the VBA Editor
2. Or go to **Developer -> Visual Basic** (if Developer tab is visible)

### Step 2: Import Modules

1. In the VBA Editor, go to **File -> Import File...**
2. Navigate to `outlook-email-filter/src/`
3. Import these files (one at a time):
   - `Config.bas`
   - `Utilities.bas`
   - `EmailFilter.bas`
   - `BatchFilter.bas`

### Step 3: Set Up ThisOutlookSession

This module **cannot** be imported -- it must be copy-pasted:

1. In Project Explorer (left panel), find **ThisOutlookSession** under "Microsoft Outlook Objects"
2. Double-click to open it
3. Press **Ctrl + A** -> **Delete** to clear existing code
4. Open `src/ThisOutlookSession.bas` in a text editor
5. Copy everything from `Option Explicit` onward
6. Paste into the ThisOutlookSession code window

### Step 4: Verify Import

Your Project Explorer should show:
```
VBAProject (filename.otm)
├── Microsoft Outlook Objects
│   └── ThisOutlookSession (with your code)
└── Modules
    ├── Config
    ├── Utilities
    ├── EmailFilter
    └── BatchFilter
```

### Step 5: Compile and Save

1. Go to **Debug -> Compile Project** (fix any errors)
2. Press **Ctrl + S** to save
3. Close the VBA Editor

## Phase 3: First Run

1. **Restart Outlook** completely (check system tray)
2. The filter automatically creates `settings.ini` with default values at `%APPDATA%\OutlookEmailFilter\`
3. Check the Immediate Window (Ctrl+G in VBA Editor) for:
   ```
   ... [INFO] Settings loaded from: C:\Users\...\AppData\Roaming\OutlookEmailFilter\settings.ini
   ... [INFO] Email Filter v2.0.0 initialized - real-time filtering active
   ```

## Phase 4: Create Folders

Create the following folders manually under Inbox (right-click Inbox -> New Folder):

| Folder | Purpose | Auto-created? |
|--------|---------|---------------|
| **Review** | Ambiguous emails for manual review | Yes (on first filter) |
| **Protected** | Protected domain emails | Yes (on first filter) |
| **LearnKeep** | Drag here to always KEEP from that sender | No -- create manually |
| **LearnDelete** | Drag here to always DELETE from that sender | No -- create manually |
| **LearnSubjectDelete** | Drag here to always DELETE by subject match | No -- create manually |

> Folder names are configurable in settings.ini under the `[Folders]` section.

### Migrating from v1.x

If you have the old folder names (I, II, III, IIII, V), run this macro in the Immediate Window:
```
DetectAndMigrateOldFolders
```
It will detect old folders and offer to rename them to the new defaults. Restart Outlook afterward.

## Phase 5: Configure Settings

Open `%APPDATA%\OutlookEmailFilter\settings.ini` in any text editor (Notepad, VS Code, etc.).

The file has 4 sections:

- **`[General]`** -- logging, self-improving toggle, batch sizes
- **`[Folders]`** -- folder names (Protected, Review, LearnKeep, etc.)
- **`[Patterns]`** -- all filter patterns (comma-separated)
- **`[LLM]`** -- Azure OpenAI configuration

Changes take effect immediately after saving (no restart needed for pattern changes). Folder name changes require an Outlook restart.

To reset all settings to defaults: delete `settings.ini` and restart Outlook.

## Phase 6: Test the Filter

### Test 1: Dry Run (Preview Only)

1. Open VBA Editor (Alt+F11) -> Immediate Window (Ctrl+G)
2. Type: `FilterExistingDryRun`
3. Check the Immediate Window for classification results
4. No emails are moved or deleted

### Test 2: Single Email Test

1. Select an email in your Inbox
2. In Immediate Window: `FilterSelectedEmail`
3. A dialog shows the classification -- you choose whether to execute

### Test 3: Version Info

In the Immediate Window:
```
ShowVersionInfo
```

## Phase 7: Add Quick Access Toolbar Buttons

For one-click access without opening the VBA Editor:

1. In Outlook, go to **File -> Options -> Quick Access Toolbar**
2. In the **"Choose commands from"** dropdown, select **Macros**
3. Recommended macros to add:
   - **Project1.FilterSelectedEmails** -- filter selected email(s)
   - **Project1.FilterCurrentFolder** -- filter current folder
   - **Project1.FilterExistingDryRun** -- dry run preview
4. Click **Add >>** -> (optional) **Modify...** for icon -> **OK**

## Phase 8: Enable Real-Time Filtering (Optional)

To automatically filter new emails as they arrive:

1. Open **ThisOutlookSession** in VBA Editor
2. Find the commented section `Private Sub inboxItems_ItemAdd`
3. **Uncomment** the entire Sub (remove the `'` from each line)
4. Save and restart Outlook

---

## Dashboard UserForm (Optional)

A graphical 4-tab Dashboard is available but requires manual creation in the VBA Editor since UserForms are binary files (.frm + .frx) that cannot be imported from text alone.

All Dashboard functionality is available via macros and settings.ini editing, so the UserForm is **not required** for normal operation.

If you want to set up the Dashboard:

### frmFilterDashboard

1. In VBA Editor: **Insert -> UserForm**
2. Set form properties in the Properties window:
   - **Name**: `frmFilterDashboard`
   - **Caption**: `Email Filter Dashboard`
   - **Width**: 620
   - **Height**: 480
3. Add a **MultiPage** control (from the Toolbox) filling the form
   - Set **Name**: `mpTabs`
   - Add 4 pages: "Filter Actions", "Patterns", "Settings", "Learned Rules"
4. Add controls to each page as listed in the `.frm` file comments
5. Open `src/frmFilterDashboard.frm` in a text editor, copy the VBA code (everything after `Option Explicit`), and paste into the UserForm's code-behind (double-click the form background)

### frmDraftReply

1. In VBA Editor: **Insert -> UserForm**
2. Set properties:
   - **Name**: `frmDraftReply`
   - **Caption**: `Draft Reply`
   - **Width**: 400
   - **Height**: 350
3. Add controls:
   - **TextBox**: `txtDraft` (MultiLine=True, ScrollBars=2, large area)
   - **CommandButton**: `cmdCopy` (Caption="Copy to Clipboard")
   - **CommandButton**: `cmdCreateReply` (Caption="Create Reply Email")
   - **CommandButton**: `cmdClose` (Caption="Close")
4. Paste the code from `src/frmDraftReply.frm`

See `src/frmFilterDashboard.frm` for the complete list of controls needed on each tab.

---

## Signing Your Macros (Optional)

For security, you can digitally sign your VBA code:

### Create a Self-Signed Certificate

1. Find `selfcert.exe` in your Office installation:
   - Usually: `C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE`
2. Run it and create a certificate (e.g., "OutlookEmailFilter")

### Sign Your Code

1. In VBA Editor, go to **Tools -> Digital Signature...**
2. Click **Choose...**
3. Select your certificate
4. Click **OK**

---

## Migrating to Another PC

To install the filter on a new PC with your existing settings and learned rules:

### Step 1: Copy data files from the old PC

Copy these 3 files from `%APPDATA%\OutlookEmailFilter\` on the old PC:

| File | Contains |
|------|----------|
| `settings.ini` | All your configured settings, patterns, folder names |
| `learned_senders.txt` | All learned sender rules (KEEP/DELETE) |
| `learned_subjects.txt` | All learned subject rules (DELETE) |

Put them somewhere accessible (USB drive, cloud, etc.).

### Step 2: Install the VBA code on the new PC

Follow Phases 1-5 above (Enable macros → Import modules → Paste ThisOutlookSession → Compile → Restart Outlook).

### Step 3: Replace the auto-generated settings

After the first restart, Outlook creates default files at `%APPDATA%\OutlookEmailFilter\`. Replace them:

1. Close Outlook on the new PC
2. Navigate to `%APPDATA%\OutlookEmailFilter\`
3. Replace `settings.ini` with your copy from the old PC
4. Copy `learned_senders.txt` and `learned_subjects.txt` into the same folder
5. Restart Outlook

### Step 4: Create folders and verify

1. Create learning folders under Inbox: **LearnKeep**, **LearnDelete**, **LearnSubjectDelete** (or whatever names are in your settings.ini)
2. In Immediate Window (Ctrl+G), verify:
   ```
   ShowVersionInfo
   ShowLearnedSenders
   FilterExistingDryRun
   ```

All your learned rules and settings are now active on the new PC.

> **Tip**: Run `ExportAllModules` on the old PC to save all `.bas` files to the Desktop for easy transfer.

---

## Troubleshooting

### "Macros are disabled"
- Check Trust Center settings (Phase 1)
- Restart Outlook after changing settings

### "Compile error: Can't find project or library"
- Go to **Tools -> References** in VBA Editor
- Uncheck any items marked "MISSING"
- Ensure "Microsoft Outlook XX.X Object Library" is checked

### Emails not being filtered automatically
- Ensure the `inboxItems_ItemAdd` event is uncommented in ThisOutlookSession
- Run `ReinitializeFilter` in Immediate Window
- Restart Outlook

### Need to undo deletions?
- Deleted emails go to **Deleted Items** folder
- Run `RestoreDeletedKeepEmails` to rescue wrongly deleted emails
- Run `RestoreFromReview` to restore Review folder emails to Inbox

### settings.ini not being created
- Check that `%APPDATA%\OutlookEmailFilter\` directory exists
- If not, create it manually and restart Outlook
