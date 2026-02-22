# Pattern Configuration Guide — Outlook Email Agent v3.0

This document explains how to configure the email classification patterns in `settings.ini`.

> **v3.0 note**: All patterns are now configured in `settings.ini` under `[Patterns]`, not in `Config.bas`. Edit with any text editor — changes take effect immediately, no Outlook restart needed.

---

## How Patterns Work

All patterns are **comma-separated strings** stored as values in `settings.ini`:

```ini
[Patterns]
NamePatterns=Xu Xin,XuXin,Prof Xu,Professor Xu
DeleteSubjectPatterns=offer,digest,newsletter,unsubscribe
```

Matching is **case-insensitive substring** for all patterns — "news" matches "newsletter", "breaking news", "newsroom". There are no exact-match or regex patterns.

---

## Pattern Types

### 1. Protected Domains (`ProtectedDomains`)

Emails from these domains are **NEVER deleted**. They are moved to the Protected folder for later reading.

```ini
ProtectedDomains=substack.com,reddit.com,redditmail.com
```

**How it works**: Extracts domain from sender's email address. Checks if the domain contains any pattern in the list. Because it's a substring match, `substack.com` matches `mail.substack.com` too.

**Use for**: Newsletters you subscribe to, sites with valuable content, domains that send important notifications.

| Sender Email | Domain | Matches? |
|--------------|--------|----------|
| `newsletter@substack.com` | `substack.com` | ✅ Yes |
| `no-reply@mail.substack.com` | `mail.substack.com` | ✅ Yes (contains "substack.com") |
| `info@company.com` | `company.com` | ❌ No |

---

### 2. Name Patterns (`NamePatterns`)

If your **name** appears in the subject line or email body (first 200 chars), the email is kept.

```ini
NamePatterns=Xu Xin,XuXin,Xuxin,Xin Xu,Professor Xu,Prof. Xu,Prof Xu,Dr. Xu,Dr Xu
```

**Include:**
- Full name variations (Chinese order and Western order)
- Title + last name (Prof. Xu, Dr. Xu)
- Title + first name (Prof. Xin)
- Common misspellings or abbreviations

> **Tip**: Use `GenerateAddressingPatterns` macro to let the LLM generate a comprehensive list from your name automatically.

---

### 3. Greeting Patterns (`GreetingPatterns`)

If the email **body starts with** any of these greetings, it is kept.

```ini
GreetingPatterns=Dear Professor Xu,Dear Prof. Xu,Dear Prof Xu,Dear Dr. Xu,Dear Dr Xu,Dear Xin,Hi Xin,Hello Xin,Dear Head,Dear Director
```

**Tips:**
- Include both formal and informal greetings
- Add position-based greetings (Director, Head, Chair) if applicable
- Whitespace at the start of the email body is trimmed automatically

---

### 4. Organizational Tags (`PolyUTags`)

Emails with these **tags in the subject** are kept.

```ini
PolyUTags=[MM],[HRO],[CUS],ToXX
```

Matching is case-insensitive (same as all other patterns). Use for department/institution tags that identify internally-routed important emails.

---

### 5. VIP Keywords (`VIPSubjectKeywords`)

If the **subject** contains any of these keywords, the email is kept.

```ini
VIPSubjectKeywords=thesis,dissertation,supervision,urgent,deadline,review request,paper submission,grant,conference,publication,meeting request,appointment,interview
```

**Academic examples to consider**: thesis, dissertation, defense, publication, manuscript, revision, grant, proposal, deadline, conference, review request, referee, editor

---

### 6. Delete Sender Patterns (`DeleteSenderPatterns`)

If the sender's **email address** contains any of these patterns, the email is deleted.

```ini
DeleteSenderPatterns=notice,noreply,no-reply,notification,marketing,promo,newsletter,digest,campaign,bulk,mailer,broadcast
```

**Use for**: Generic notification addresses, marketing senders, automated system messages.

| Sender Email | Matches Pattern | Action |
|--------------|-----------------|--------|
| `noreply@company.com` | `noreply` | DELETE |
| `marketing@brand.com` | `marketing` | DELETE |
| `john.smith@company.com` | (none) | Continue to next rule |

> **Warning**: Avoid overly broad patterns. `info` would delete emails from `info@anycompany.com`, which is likely not what you want.

---

### 7. Delete Known Senders (`DeleteKnownSenders`)

If the sender's **display name** contains any of these, the email is deleted.

```ini
DeleteKnownSenders=LinkedIn Job Alerts,edX,Cathay Pacific,HKBN,Medium Daily Digest
```

**Use for**: Specific services you don't want emails from, bulk senders with recognizable display names.

---

### 8. Delete Subject Patterns (`DeleteSubjectPatterns`)

If the email **subject** contains any of these keywords, it is deleted.

```ini
DeleteSubjectPatterns=優惠,offer,digest,newsletter,unsubscribe,job alert,weekly roundup,daily digest,promotional,special offer,limited time
```

**Supports multi-language**: Chinese characters work (e.g., "優惠" = discount/offer).

---

## Classification Priority

Patterns are checked in this order. **First match wins.**

| Priority | Rule | Patterns Used |
|----------|------|---------------|
| 0 | Learned sender | `learned_senders.txt` (drag to LearnKeep/LearnDelete) |
| 0.5 | Learned subject | `learned_subjects.txt` (drag to LearnSubjectDelete) |
| 1 | Protected domain | `ProtectedDomains` |
| 2 | Personally addressed | `NamePatterns`, `GreetingPatterns` |
| 3 | Organizational tags | `PolyUTags` |
| 4 | VIP keywords | `VIPSubjectKeywords` |
| 5 | Reply chain | Built-in (RE:, AW:) |
| 6 | Forward chain | Built-in (FW:, FWD:, WG:) |
| 7 | Known spam senders | `DeleteKnownSenders` |
| 8 | Spam sender email | `DeleteSenderPatterns` |
| 9 | Spam subject | `DeleteSubjectPatterns` |
| 10 | No match | Review folder (or LLM if enabled) |

---

## Editing Patterns

Open `%APPDATA%\OutlookEmailFilter\settings.ini` in any text editor:

```ini
[Patterns]
; Before
ProtectedDomains=substack.com,reddit.com

; After — add new domains
ProtectedDomains=substack.com,reddit.com,arxiv.org,nature.com
```

Save the file. Changes take effect immediately — no restart needed.

---

## Testing Patterns

After editing:

1. Open VBA Editor (Alt+F11) → Immediate Window (Ctrl+G)
2. Run `FilterExistingDryRun` to preview results
3. Check Immediate Window for unexpected classifications
4. Adjust and repeat

To test a specific pattern match interactively:

```vba
? ContainsAny("john@noreply.company.com", "noreply,no-reply,donotreply")
' Returns: True

? ContainsAny("Dear Professor Xu, I hope", "Dear Professor Xu,Dear Prof")
' Returns: True
```

---

## Common Mistakes

### Too broad — will delete legitimate emails

```ini
; BAD: "info" matches info@hospital.com, info@university.edu
DeleteSenderPatterns=info
```

```ini
; BETTER: target automated addresses
DeleteSenderPatterns=noreply,no-reply,donotreply
```

### Missing variations — won't catch all addressing forms

```ini
; BAD: won't catch "Dr Xu" (no period)
NamePatterns=Dr. Xu
```

```ini
; GOOD: include both with and without period
NamePatterns=Dr. Xu,Dr Xu,Prof. Xu,Prof Xu,Professor Xu
```

### Greeting too short — will over-match

```ini
; BAD: "Dear" alone matches any email starting with "Dear"
GreetingPatterns=Dear
```

```ini
; GOOD: include the name in the greeting
GreetingPatterns=Dear Professor Xu,Dear Prof. Xu,Dear Xin,Hi Xin
```

---

## Academic Configuration Example

```ini
[Patterns]
ProtectedDomains=arxiv.org,researchgate.net,springer.com,ieee.org,acm.org,nature.com

NamePatterns=Prof. Smith,Professor Smith,Dr. Smith,John Smith,J. Smith

GreetingPatterns=Dear Professor Smith,Dear Prof. Smith,Dear Dr. Smith,Dear John,Hi John,Hello Professor

VIPSubjectKeywords=thesis,dissertation,defense,publication,manuscript,revision,grant,NSF,proposal,deadline,conference,ICML,NeurIPS,review request,referee,editor

DeleteKnownSenders=Beall's List,Predatory Journals,OMICS

DeleteSubjectPatterns=call for papers,invitation to submit,special issue,guest editor,waived APC,open access opportunity,be a reviewer
```

---

## Learned Rules vs. Static Patterns

Unlike the static patterns above, **learned rules** are created by dragging emails into special folders:

| Folder | Effect | Priority |
|--------|--------|----------|
| **LearnKeep** | Always **KEEP** emails from that sender | Rule 0 (highest) |
| **LearnDelete** | Always **DELETE** emails from that sender | Rule 0 (highest) |
| **LearnSubjectDelete** | Always **DELETE** emails with matching subject | Rule 0.5 |

Learned rules override ALL static patterns. If you drag a `noreply@company.com` email into LearnKeep, future emails from that sender will be kept even though "noreply" normally triggers deletion.

Data files: `%APPDATA%\OutlookEmailFilter\learned_senders.txt` and `learned_subjects.txt`

To view all active learned rules:
```
ShowLearnedSendersList
ShowLearnedSubjectsList
```

---

## Backup Your Patterns

Before making significant pattern changes:

1. Copy `settings.ini` from `%APPDATA%\OutlookEmailFilter\` to a backup location
2. Or export your VBA modules manually: in the VBA Editor, right-click each module → **Export File...**

To revert: replace the file (or delete it and restart Outlook to regenerate defaults).
