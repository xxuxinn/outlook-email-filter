# Update Guide: Self-Improving Filter

> **Historical document** — This guide covers the upgrade to the self-improving feature (originally v1.x → v2.0-era).
> If you are doing a **fresh install or upgrading to v3.0**, use [INSTALL.md](INSTALL.md) instead.

---

## What This Feature Adds

The self-improving filter **learns from your manual sorting**. When you drag an email into the **LearnKeep** (always keep) or **LearnDelete** (always delete) folder, the sender is remembered permanently. Future emails from that sender are classified at **Rule 0** — the highest priority — overriding all static patterns.

**Additional macros:**
- **`FilterSelectedEmails`** — classify and act on selected emails
- **`FilterCurrentFolder`** — classify and act on all emails in the current folder

---

## Installing / Upgrading

1. Open VBA Editor (Alt+F11) → File → Import File → import each `.bas` file from `src/`
2. Paste `src/ThisOutlookSession.bas` into the built-in ThisOutlookSession module
3. Compile (Debug → Compile Project) → Save (Ctrl+S) → Restart Outlook

For full details, see [INSTALL.md](INSTALL.md) Part 2.

---

## One-Time Backfill

If you already have emails in your learning folders and want to bulk-import all their senders:

1. Open Immediate Window (Ctrl+G)
2. Type: `ImportExistingLearnedFolders`

This scans the LearnKeep and LearnDelete folders and records all senders as learned rules. Run this **once** after the first install. After this, new drags are captured automatically.

---

## How It Works Day-to-Day

| You do this... | The filter does this... |
|----------------|------------------------|
| Drag email → **LearnKeep** | Records sender as **always KEEP** |
| Drag email → **LearnDelete** | Records sender as **always DELETE** |
| Change your mind | Drag another email from that sender to the opposite folder — new rule overwrites |
| Restart Outlook | All rules reload from file automatically |

---

## Verification

In the Immediate Window:

```
ShowLearnedSenders          ' check rule count and file path
FilterExistingDryRun        ' look for [+LR] and [xLR] icons
```

| Icon | Meaning |
|------|---------|
| `[+LR]` | Kept by learned sender rule |
| `[xLR]` | Deleted by learned sender rule |

---

## Troubleshooting

**"Self-improving filter not active" in log**: The learning folders are missing or their names don't match what's in `settings.ini` under `[Folders]`. Check that `LearnKeep` and `LearnDelete` exist directly under Inbox.

**No `[+LR]`/`[xLR]` icons in dry run**: Run `ShowLearnedSenders` to check count. If 0, run `ImportExistingLearnedFolders`. If count > 0 but no icons, those learned senders may not have emails in the current Inbox.

**Want to start fresh**: Delete `%APPDATA%\OutlookEmailFilter\learned_senders.txt` and run `ReloadLearnedSenders`.
