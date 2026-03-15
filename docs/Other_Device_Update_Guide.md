# Updating VBA Modules on Another Device

Step-by-step guide for syncing the latest Git changes and updating Outlook VBA macros.

---

## Part 1: Pull Latest Code

Open a terminal in the project folder and run:

```bash
git pull
```

If you have uncommitted local changes, stash them first:

```bash
git stash
git pull
git stash pop
```

---

## Part 2: Update VBA Modules in Outlook

### Step 1: Open VBA Editor

- Press **Alt+F11** in Outlook to open the VBA Editor
- Press **Ctrl+G** to open the Immediate Window

### Step 2: Stop Timers and Event Handlers

In the Immediate Window, run these two commands (one at a time, press Enter after each):

```
StopCommandPollerStd
ThisOutlookSession.DisableRealTimeFilter
```

> **Why?** The Win32 command poller fires every 2 seconds. If you remove a module while the timer is active, the callback points to deallocated memory and Outlook will crash. Always stop timers first.

### Step 3: Remove Old Modules

In the Project Explorer (left panel), right-click each of these modules and choose **Remove**. When prompted "Do you want to export the file before removing it?", click **No**:

- `Config`
- `Utilities`
- `EmailFilter` *(only if changed in this update)*
- `EmailAgent` *(only if changed in this update)*
- `BatchFilter` *(only if changed in this update)*

> **Tip:** Check which `.bas` files changed by running `git diff --stat HEAD~1` in the terminal. Only remove/reimport modules that actually changed.

### Step 4: Import Updated Modules

- Go to **File → Import File** (or press **Ctrl+M**)
- Navigate to the `src/` folder in the project directory
- Import each `.bas` file you removed in Step 3:
  - `src/Config.bas`
  - `src/Utilities.bas`
  - `src/EmailFilter.bas` *(if changed)*
  - `src/EmailAgent.bas` *(if changed)*
  - `src/BatchFilter.bas` *(if changed)*

### Step 5: Update ThisOutlookSession (if changed)

`ThisOutlookSession.bas` is **not** a regular module — it cannot be imported. Instead:

1. Open `src/ThisOutlookSession.bas` in a text editor (Notepad, VS Code, etc.)
2. Select all content **below** the comment block at the top (starting from `Option Explicit`)
3. In the VBA Editor, find **ThisOutlookSession** under **Microsoft Outlook Objects** in the Project Explorer
4. Double-click to open it
5. Select all existing code (**Ctrl+A**) and paste the new code (**Ctrl+V**)

### Step 6: Save

Press **Ctrl+S** to save the VBA project.

### Step 7: Restart

In the Immediate Window:

```
ThisOutlookSession.ReinitializeFilter
```

> **Note:** You must use the full qualified name `ThisOutlookSession.ReinitializeFilter` — just `ReinitializeFilter` alone won't work from the Immediate Window because it lives in a document module, not a standard module.

---

## Part 3: Configure New Settings (First Time Only)

If the update introduces new settings (e.g., Cloud Sync), configure them in the Immediate Window:

### Cloud Sync Setup

```
WriteINISetting "Sync", "EnableCloudSync", "True"
WriteINISetting "Sync", "CloudSyncPath", "C:\Users\xxuxinn\OneDrive - The Hong Kong Polytechnic University"
```

Then reload settings:

```
ThisOutlookSession.ReinitializeFilter
```

> **Note:** `CloudSyncPath` should point to the same OneDrive root on both devices. If the Windows username differs, adjust the path accordingly.

---

## Part 4: Verify

In the Immediate Window, run these checks:

```
? RuntimeClassifyBodyChars
? RuntimeUseLLM
? RuntimeLLMProvider
? RuntimeEnableCloudSync
? RuntimeCloudSyncPath
? Len(GetAPIKey())
```

Then test with:

```
FilterExistingDryRun
```

This previews classifications without making changes.

For cloud sync, run:

```
SyncLearnedRules
```

You should see a MsgBox summarizing the merge results for learned senders, subjects, and replies.

---

## Quick Reference: Which Modules to Update

| Module | How to Update | When |
|--------|---------------|------|
| `Config.bas` | Remove → Import | When constants or Runtime vars change |
| `Utilities.bas` | Remove → Import | When helpers, LLM, sync, or bridge logic changes |
| `EmailFilter.bas` | Remove → Import | When classification rules change |
| `EmailAgent.bas` | Remove → Import | When reply drafting or agent features change |
| `BatchFilter.bas` | Remove → Import | When batch macros change |
| `ThisOutlookSession` | Copy-paste into built-in object | When event handlers or startup/quit logic changes |
| `Installer.bas` | Remove → Import | When setup automation changes (rare) |

---

## Troubleshooting

### "Sub or Function not defined" when calling ReinitializeFilter
Use the full qualified name: `ThisOutlookSession.ReinitializeFilter`

### Outlook crashes after reimporting modules
You forgot to stop the command poller timer. Restart Outlook and run `StopCommandPollerStd` before removing any modules next time.

### Cloud sync says "path does not exist"
Verify that OneDrive is running and synced. Check the path in settings.ini matches the actual OneDrive folder path on this device.

### New settings not taking effect
Make sure you ran `ThisOutlookSession.ReinitializeFilter` after importing modules. This reloads `settings.ini` into the Runtime variables.
