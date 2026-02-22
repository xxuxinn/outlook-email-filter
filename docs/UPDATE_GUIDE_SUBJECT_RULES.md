# Update Guide: Subject Rules & Server Import

> **Historical document** — This guide covers the upgrade that added learned subject rules and server rule import/export.
> If you are doing a **fresh install or upgrading to v3.0**, use [INSTALL.md](INSTALL.md) instead.

---

## What This Feature Adds

1. **Learned subject rules (Rule 0.5)** — Drag an email into the **LearnSubjectDelete** folder to permanently delete all future emails with matching subjects. Uses case-insensitive substring matching.
2. **Server rule import** — Migrate existing server-side Outlook Rules into the VBA agent as learned DELETE rules.
3. **Server rule export** — Push learned DELETE rules back to Exchange for 24/7 server-side filtering.

---

## Installing / Upgrading

1. Open VBA Editor (Alt+F11) → File → Import File → import each `.bas` file from `src/`
2. Paste `src/ThisOutlookSession.bas` into the built-in ThisOutlookSession module
3. Compile (Debug → Compile Project) → Save (Ctrl+S) → Restart Outlook

For full details, see [INSTALL.md](INSTALL.md) Part 2.

---

## One-Time Backfill

If you already have emails in the LearnSubjectDelete folder:

1. Open Immediate Window (Ctrl+G)
2. Type: `ImportExistingLearnedSubjectFolder`

To migrate server-side Outlook Rules:

1. Type: `ImportServerRules`
2. Confirm when prompted
3. Manually delete the server rules afterward: Home → Rules → Manage Rules & Alerts

---

## How Subject Learning Works

| You do this... | The filter does this... |
|----------------|------------------------|
| Drag email → **LearnSubjectDelete** | Records subject as always DELETE |
| Email arrives with matching subject | DELETE (Rule 0.5 — substring match) |

Subject matching is **case-insensitive substring** — dragging an email with subject "Weekly Marketing Update" will delete all future emails whose subject contains "Weekly Marketing Update".

---

## Classification Rules (Full List)

| Priority | Rule | Source |
|----------|------|--------|
| 0 | Learned sender | `learned_senders.txt` |
| 0.5 | Learned subject | `learned_subjects.txt` |
| 1 | Protected domain | settings.ini |
| 2 | Personally addressed | settings.ini |
| 3 | Organizational tags | settings.ini |
| 4 | VIP keywords | settings.ini |
| 5 | Reply chain (RE:, AW:) | Built-in |
| 6 | Forward chain (FW:, FWD:) | Built-in |
| 7 | Known spam senders | settings.ini |
| 8 | Spam sender email | settings.ini |
| 9 | Spam subject keywords | settings.ini |
| 10 | No match | Review (or LLM) |

---

## Verification

In the Immediate Window:

```
ShowLearnedSubjectsList     ' check all subject rules
FilterExistingDryRun        ' look for [xLS] icons
```

| Icon | Meaning |
|------|---------|
| `[xLS]` | Deleted by learned subject rule |
| `[xLR]` | Deleted by learned sender rule |
| `[+LR]` | Kept by learned sender rule |

---

## Troubleshooting

**"Learning folder (LearnSubjectDelete) not found"**: The folder is missing or its name doesn't match the `LearnSubject` key in `settings.ini` under `[Folders]`. Check the folder exists directly under Inbox.

**No `[xLS]` icons in dry run**: Run `ShowLearnedSubjectsList` to check count. If 0, drag an email into LearnSubjectDelete or run `ImportServerRules`. Remember: subject matching is substring — the rule key must be a substring of the email subject to trigger.

**ImportServerRules shows errors**: The Outlook Rules COM API is fragile. Partial imports are expected — the macro continues past errors and reports what was successfully imported.

**Want to clear all subject rules**: Delete `%APPDATA%\OutlookEmailFilter\learned_subjects.txt` and restart Outlook.
