# Outlook Email Filter — Subject Rules & Server Import Update Guide

## What This Update Adds

Two new features:

1. **Learned subject rules (Rule 0.5)** — Drag an email into folder **V** to permanently delete all future emails with matching subjects. Uses case-insensitive substring matching.
2. **Server rule import** — Migrate your existing server-side Outlook Rules into the VBA filter as learned DELETE rules, then delete the server rules to avoid conflicts.

---

## Complete Classification Rules (all rule types)

| Priority | Rule | How It Matches | Action | Config / Source |
|----------|------|---------------|--------|-----------------|
| 0 | **Learned sender** | Exact sender email match from III/IIII folders | KEEP or DELETE | `learned_senders.txt` |
| 0.5 | **Learned subject** | Case-insensitive substring match from V folder | DELETE only | `learned_subjects.txt` |
| 1 | **Protected domain** | Sender's domain in PROTECTED_DOMAINS list | Move to "II" | `Config.bas` |
| 2 | **Personally addressed** | Name in subject/body, or greeting match | KEEP | `Config.bas` NAME_PATTERNS, GREETING_PATTERNS |
| 3 | **Organizational tags** | Subject contains [MM], [HRO], etc. (case-sensitive) | KEEP | `Config.bas` POLYU_TAGS |
| 4 | **VIP keywords** | Subject contains thesis, deadline, etc. | KEEP | `Config.bas` VIP_SUBJECT_KEYWORDS |
| 5 | **Reply chain** | Subject starts with RE: or AW: | KEEP | Built-in |
| 6 | **Forward chain** | Subject starts with FW:, FWD:, or WG: | KEEP | Built-in |
| 7 | **Known spam senders** | Sender name matches DELETE_KNOWN_SENDERS | DELETE | `Config.bas` |
| 8 | **Spam email patterns** | Sender email contains noreply, marketing, etc. | DELETE | `Config.bas` DELETE_SENDER_PATTERNS |
| 9 | **Spam subject keywords** | Subject contains newsletter, unsubscribe, etc. | DELETE | `Config.bas` DELETE_SUBJECT_PATTERNS |
| 10 | **No match** | Nothing matched | Move to "I" (or LLM) | `Config.bas` USE_LLM_API |

**First match wins** — the engine stops at the first rule that matches.

---

## Complete Macros Reference

### Filtering Macros

| Macro | What It Does | Where to Run |
|-------|-------------|--------------|
| `FilterExistingDryRun` | Preview decisions for Inbox emails (no changes made) | Immediate Window |
| `FilterExistingEmails` | Filter all Inbox emails (with confirmation dialog) | Immediate Window |
| `FilterAllFolders` | Filter Inbox + Other + PST archives | Immediate Window |
| `FilterSelectedEmail` | Classify one selected email, ask before acting | Immediate Window |
| `FilterSelectedEmails` | Filter selected email(s) with confirmation | **QAT button** or Immediate Window |
| `FilterCurrentFolder` | Filter all emails in current folder with confirmation | **QAT button** or Immediate Window |
| `FilterLastNDays 7` | Filter emails from last N days only | Immediate Window |
| `FilterByDateRange #1/1/2026#, #2/1/2026#` | Filter emails in a specific date range | Immediate Window |
| `GenerateClassificationReport` | Count classifications without taking action | Immediate Window |

### Learned Sender Macros

| Macro | What It Does | When to Use |
|-------|-------------|-------------|
| `ShowLearnedSenders` | Show learned sender rule count and file path | Anytime — quick status check |
| `ShowLearnedSendersList` | Dump all sender rules to Immediate Window with timestamps | Review all rules in detail |
| `ReloadLearnedSenders` | Force reload sender rules from file | After manually editing the .txt file |
| `CleanLearnedSendersFile` | Remove duplicate entries from learned_senders.txt | Periodic maintenance |
| `ImportExistingLearnedFolders` | Scan III/IIII folders, record all senders | **Once** after first install |

### Learned Subject Macros (NEW)

| Macro | What It Does | When to Use |
|-------|-------------|-------------|
| `ShowLearnedSubjectsList` | Dump all subject rules to Immediate Window with timestamps | Review all subject rules |
| `CleanLearnedSubjectsFile` | Remove duplicate entries from learned_subjects.txt | Periodic maintenance |
| `ImportExistingLearnedSubjectFolder` | Scan V folder, record all subjects as DELETE | **Once** after first install |

### Server Rule Import / Export

| Macro | What It Does | When to Use |
|-------|-------------|-------------|
| `ImportServerRules` | Import enabled server-side rules as learned DELETE rules | **Once** to migrate server rules |
| `ExportLearnedRulesToServer` | Export learned DELETE rules as server-side Outlook Rules | After building up learned rules, for 24/7 filtering |

### Bulk Operations

| Macro | What It Does |
|-------|-------------|
| `BulkDeleteBySender "pattern"` | Delete all emails from senders matching a pattern |
| `MoveProtectedSources` | Move all protected domain emails to "II" |

### Undo / Recovery

| Macro | What It Does |
|-------|-------------|
| `RestoreFromReview` | Move emails from "I" folder back to Inbox |
| `RestoreDeletedKeepEmails` | Scan Deleted Items, move KEEP/MOVE_II emails back |

### System Control

| Macro | What It Does |
|-------|-------------|
| `ReinitializeFilter` | Restart all event handlers (equivalent to Outlook restart) |
| `EnableRealTimeFilter` | Turn on automatic new-mail filtering |
| `DisableRealTimeFilter` | Turn off automatic filtering + disconnect learning folders |

---

## Learning Folders

| Folder | What It Learns | Rule Priority | Data File |
|--------|---------------|---------------|-----------|
| **III** | Always KEEP from that sender | Rule 0 | `learned_senders.txt` |
| **IIII** | Always DELETE from that sender | Rule 0 | `learned_senders.txt` |
| **V** | Always DELETE emails with matching subject | Rule 0.5 | `learned_subjects.txt` |

Data files are stored at `%APPDATA%\OutlookEmailFilter\` and persist across Outlook restarts.

---

## Pre-Update Checklist

- [ ] Outlook is open
- [ ] You can access the VBA Editor (Alt+F11)
- [ ] Folder **V** exists under Inbox (create it if not — right-click Inbox → New Folder → name it **V**)

---

## Step 1: Open VBA Editor

Press **Alt + F11** in Outlook.

---

## Step 2: Remove Old Modules

In the **Project Explorer** panel (left side), expand **Modules**. For each of these four modules:

1. Right-click the module name → **Remove [ModuleName]...**
2. When asked **"Do you want to export the file before removing it?"** → click **No**

Remove:
- [ ] **Config**
- [ ] **Utilities**
- [ ] **EmailFilter**
- [ ] **BatchFilter**

---

## Step 3: Import Updated Modules

1. Go to **File → Import File...**
2. Navigate to `outlook-email-filter/src/`
3. Import these four files, one at a time:

- [ ] `Config.bas`
- [ ] `Utilities.bas`
- [ ] `EmailFilter.bas`
- [ ] `BatchFilter.bas`

---

## Step 4: Update ThisOutlookSession

This module **cannot** be imported — it must be copy-pasted.

1. In Project Explorer, find **ThisOutlookSession** under **Microsoft Outlook Objects**
2. Double-click to open it
3. Press **Ctrl + A** to select all existing code
4. Press **Delete** to clear it
5. Open the file `src/ThisOutlookSession.bas` in a text editor
6. Copy **everything below the header comment block** (starting from `Option Explicit`)
7. Paste into the ThisOutlookSession code window

---

## Step 5: Compile and Save

1. Go to **Debug → Compile Project**
   - If there are errors, check for duplicate modules (e.g., "Config1") — remove the duplicate
2. Press **Ctrl + S** to save

---

## Step 6: Restart Outlook

1. **Close Outlook completely** (check system tray)
2. **Reopen Outlook**
3. Open VBA Editor (Alt+F11) → Immediate Window (Ctrl+G)
4. You should see:
   ```
   ... [INFO] Sender learning active (N learned sender rules)
   ... [INFO] Subject learning active (0 learned subject rules)
   ... [INFO] Email Filter initialized - real-time filtering active
   ```

---

## Step 7: Import Server-Side Rules (Optional)

If you have server-side Outlook Rules you want to migrate:

1. In the Immediate Window, type:
   ```
   ImportServerRules
   ```
2. A dialog shows how many rules were found — click **Yes**
3. Summary shows how many senders and subjects were imported
4. Verify with:
   ```
   ShowLearnedSendersList
   ShowLearnedSubjectsList
   ```
5. **Manually delete server rules** via: Home → Rules → Manage Rules & Alerts

---

## Step 8: Test

### Test 1: Check subject learning is active

```
ShowLearnedSubjectsList
```

If you just imported server rules, you should see the imported subject keywords.

### Test 2: Dry run

```
FilterExistingDryRun
```

Look for:
- **`[xLS]`** — emails deleted by learned subject rule (NEW)
- **`[+LR]`** — emails kept by learned sender rule
- **`[xLR]`** — emails deleted by learned sender rule
- **`[DEL]`**, **`[OK]`**, **`[II]`**, **`[???]`** — regular rules

### Test 3: Live subject learning

1. Find a junk email in your Inbox
2. **Drag it into the V folder**
3. Check the Immediate Window:
   ```
   LEARNED SUBJECT DELETE from folder V: Weekly Marketing Update (SomeSender)
   ```
4. Run `FilterExistingDryRun` — matching emails now show `[xLS]`

---

## Dry-Run Icon Reference

| Icon | Meaning |
|------|---------|
| `[DEL]` | Will be deleted (static rule 7/8/9) |
| `[xLR]` | Will be deleted (learned sender rule) |
| `[xLS]` | Will be deleted (learned subject rule) |
| `[II]` | Will be moved to "II" folder (protected domain) |
| `[OK]` | Will stay in Inbox (static keep rule) |
| `[+LR]` | Will stay in Inbox (learned keep rule) |
| `[???]` | Will be moved to "I" folder (or LLM review) |

---

## Files Changed in This Update

| File | What Changed |
|------|-------------|
| `src/Config.bas` | +2 constants: FOLDER_LEARN_SUBJECT_DELETE, LEARNED_SUBJECTS_FILE |
| `src/Utilities.bas` | +learned subjects cache (6 functions), updated FormatStats |
| `src/EmailFilter.bas` | +Rule 0.5 block, +lastClassifyWasLearnedSubject flag |
| `src/BatchFilter.bas` | +`[xLS]` dry-run icon, +3 subject macros, +ImportServerRules, +LoadLearnedSubjects calls |
| `src/ThisOutlookSession.bas` | +WithEvents for V folder, +subject learning event handler, updated startup/disable |
| `CLAUDE.md` | Updated architecture, rules, macros, dev notes |
| `README.md` | Updated rules table, macros, self-improving section, file tree |
| `docs/INSTALL.md` | Updated folder creation list, macros table |

---

## Troubleshooting

### "Learning folder (V) not found" in log

The V folder is missing or misspelled. Check:
- Folder exists **directly under Inbox** (not nested deeper)
- Name is exactly **V** (single uppercase letter)

### No `[xLS]` icons in dry run

- Run `ShowLearnedSubjectsList` to check if any subject rules exist (count > 0)
- If count is 0, either drag emails into V or run `ImportServerRules`
- Remember: subject matching is **substring** — "Marketing" will match "Weekly Marketing Update"

### ImportServerRules shows errors

The Outlook Rules COM API is fragile. Partial imports are expected — the macro continues past errors and reports what it successfully imported. Check the Immediate Window for details on which rules were processed.

### Want to clear all learned subject rules

Delete the file at `%APPDATA%\OutlookEmailFilter\learned_subjects.txt` and run `ReloadLearnedSenders` (which triggers a fresh start). Or restart Outlook.
