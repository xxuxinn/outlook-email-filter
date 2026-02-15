# Outlook Email Filter — Self-Improving Update Guide

## What This Update Adds

The filter now **learns from your manual sorting**. When you drag an email into folder **III** (keep) or **IIII** (delete), the sender is remembered permanently. Future emails from that sender are classified at **Rule 0** — the highest priority — overriding all static patterns in Config.bas.

Additionally, **two new quick-filter macros** are available for one-click filtering:
- **`FilterSelectedEmails`** — classify and act on selected emails
- **`FilterCurrentFolder`** — classify and act on all emails in the current folder

These can be added to the Quick Access Toolbar (File → Options → Quick Access Toolbar → Macros) for toolbar-button access.

---

## Pre-Update Checklist

- [ ] Outlook is open
- [ ] You can access the VBA Editor (Alt+F11)
- [ ] Folders **III** and **IIII** already exist under Inbox (they are not auto-created)

---

## Step 1: Open VBA Editor

Press **Alt + F11** in Outlook.

---

## Step 2: Remove Old Modules

In the **Project Explorer** panel (left side), expand **Modules**. For each of these four modules, do:

1. Right-click the module name → **Remove [ModuleName]...**
2. When asked **"Do you want to export the file before removing it?"** → click **No**

Remove in this order:
- [ ] **Config**
- [ ] **Utilities**
- [ ] **EmailFilter**
- [ ] **BatchFilter**

> After removal, only **ThisOutlookSession** should remain (under Microsoft Outlook Objects).

---

## Step 3: Import Updated Modules

1. Go to **File → Import File...**
2. Navigate to `outlook-email-filter/src/`
3. Import these four files, one at a time:

- [ ] `Config.bas`
- [ ] `Utilities.bas`
- [ ] `EmailFilter.bas`
- [ ] `BatchFilter.bas`

Your Modules folder should now show all four again.

---

## Step 4: Update ThisOutlookSession

This module **cannot** be imported — it must be copy-pasted.

1. In Project Explorer, find **ThisOutlookSession** under **Microsoft Outlook Objects**
2. Double-click to open it
3. Press **Ctrl + A** to select all existing code
4. Press **Delete** to clear it
5. Open the file `src/ThisOutlookSession.bas` in a text editor (Notepad is fine)
6. Copy **everything below the header comment block** (starting from `Option Explicit`)
7. Paste into the ThisOutlookSession code window in VBA Editor

---

## Step 5: Verify Project Structure

Your Project Explorer should look like:

```
VBAProject (filename.otm)
├── Microsoft Outlook Objects
│   └── ThisOutlookSession    ← updated (with learning event handlers)
└── Modules
    ├── Config                ← updated (5 new constants)
    ├── Utilities             ← updated (learned senders cache + file I/O)
    ├── EmailFilter           ← updated (Rule 0 insertion)
    └── BatchFilter           ← updated (new icons + 2 new macros)
```

---

## Step 6: Compile and Save

1. Go to **Debug → Compile Project** in the VBA Editor menu bar
   - If there are errors, check that all 4 modules imported correctly
   - Common issue: a module imported twice (e.g., "Config1") — remove the duplicate
2. Press **Ctrl + S** to save

---

## Step 7: Restart Outlook

1. **Close Outlook completely** (check the system tray — make sure it's fully closed)
2. **Reopen Outlook**

On startup, `Application_Startup` will:
- Connect the inbox watcher (same as before)
- Look for III and IIII folders under Inbox
- If found: attach learning watchers and load the learned senders cache
- If not found: log a warning, self-improving feature stays inactive

---

## Step 8: Import Existing Emails (One-Time)

Your III and IIII folders already have emails in them. This one-time macro reads every email in both folders and records all senders as learned rules.

1. Open VBA Editor (**Alt + F11**)
2. Open the Immediate Window (**Ctrl + G**)
3. Type and press Enter:

```
ImportExistingLearnedFolders
```

4. A dialog will show results like:

```
From 'III': 12 senders → KEEP
From 'IIII': 8 senders → DELETE
Total unique rules now: 18
```

> You only need to run this **once**. After this, new drags into III/IIII are captured automatically by the event handlers.

---

## Step 9: Verify Everything Works

### Test 1: Check learned rules loaded

In the Immediate Window, type:

```
ShowLearnedSenders
```

You should see the rule count matching what `ImportExistingLearnedFolders` reported, plus the file path (`%APPDATA%\OutlookEmailFilter\learned_senders.txt`).

### Test 2: Dry run with learned rule icons

```
FilterExistingDryRun
```

Check the Immediate Window output. Look for:
- **`[+LR]`** — emails kept because of a learned rule (sender was in III)
- **`[xLR]`** — emails deleted because of a learned rule (sender was in IIII)
- **`[DEL]`**, **`[OK]`**, **`[II]`**, **`[???]`** — regular rules (unchanged)

### Test 3: Live learning

1. Find any email in your Inbox
2. **Drag it into III**
3. Check the Immediate Window — you should see:
   ```
   LEARNED KEEP from folder III: sender@example.com (Sender Name)
   ```
4. Run `FilterExistingDryRun` again — that sender's emails should now show `[+LR]`

---

## New Macros Reference

| Macro | When to Use |
|-------|-------------|
| `ImportExistingLearnedFolders` | **Once** after this upgrade — backfills from existing III/IIII emails |
| `ShowLearnedSenders` | Anytime — check how many rules are loaded and where the file is |
| `ReloadLearnedSenders` | After manually editing the `.txt` file — forces cache refresh |

---

## How It Works Going Forward

| You do this... | The filter does this... |
|----------------|------------------------|
| Drag email → **III** | Records sender as **always KEEP** |
| Drag email → **IIII** | Records sender as **always DELETE** |
| Change your mind | Drag another email from that sender into the opposite folder — new rule overwrites |
| Restart Outlook | All rules reload automatically from the text file |

---

## Troubleshooting

### "Self-improving filter not active" in log

The III or IIII folder is missing or misspelled. Check:
- Both folders exist **directly under Inbox** (not nested deeper)
- Names are exactly **III** (3 letters) and **IIII** (4 letters)

### No `[+LR]`/`[xLR]` icons in dry run

- Run `ShowLearnedSenders` to check if rules loaded (count > 0)
- If count is 0, run `ImportExistingLearnedFolders` first
- If count > 0 but no icons, the learned senders may not match any Inbox emails

### Compile error on import

- Go to **Debug → Compile Project**
- Check for duplicate modules (e.g., "Config1" alongside "Config") — remove the duplicate
- Check **Tools → References** — uncheck anything marked "MISSING"

### Want to start fresh

Delete the file at the path shown by `ShowLearnedSenders` and run `ReloadLearnedSenders`. This clears all learned rules.

---

## Files Changed in This Update

| File | What Changed |
|------|-------------|
| `src/Config.bas` | +5 self-improving constants |
| `src/Utilities.bas` | +learned senders cache, file I/O (6 functions), updated FormatStats |
| `src/EmailFilter.bas` | +Rule 0 block, +lastClassifyWasLearned flag |
| `src/BatchFilter.bas` | +`[+LR]`/`[xLR]` dry-run icons, +learned count in report, +ShowLearnedSenders, +ImportExistingLearnedFolders |
| `src/ThisOutlookSession.bas` | +WithEvents for III/IIII, +learning event handlers, updated startup (no auto-create), updated DisableRealTimeFilter |
| `CLAUDE.md` | Updated architecture, rules, macros, dev notes |
| `README.md` | Updated features, rules table, macros, +self-improving section |
| `docs/INSTALL.md` | Updated folder creation, macros table |
| `docs/PATTERNS.md` | Updated priority list, +learned rules section |
