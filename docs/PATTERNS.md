# Email Filter Patterns Guide

This document explains how to customize the email classification patterns in `Config.bas`.

## Pattern Types

### 1. Protected Domains (`PROTECTED_DOMAINS`)

Emails from these domains are **NEVER deleted**. They're moved to the "II" folder for later reading.

```vba
Public Const PROTECTED_DOMAINS As String = "substack.com,reddit.com,redditmail.com"
```

**Use for:**
- Newsletter services you subscribe to
- Sites with valuable content you want to keep
- Domains with important notifications

**How it works:**
- Extracts domain from sender's email address
- Checks if domain contains any pattern in the list
- Case-insensitive matching

**Examples:**
| Sender Email | Domain | Matches? |
|--------------|--------|----------|
| `newsletter@substack.com` | `substack.com` | ✅ Yes |
| `no-reply@mail.substack.com` | `mail.substack.com` | ✅ Yes (contains "substack.com") |
| `info@reddit.com` | `reddit.com` | ✅ Yes |

---

### 2. Delete Sender Patterns (`DELETE_SENDER_PATTERNS`)

If the sender's **email address** contains any of these, the email is deleted.

```vba
Public Const DELETE_SENDER_PATTERNS As String = "notice,noreply,notification,no-reply,marketing,promo,newsletter,digest,campaign,bulk,mailer,broadcast"
```

**Use for:**
- Generic notification addresses
- Marketing/promotional senders
- Automated system messages

**Examples:**
| Sender Email | Matches Pattern | Action |
|--------------|-----------------|--------|
| `noreply@company.com` | `noreply` | DELETE |
| `marketing@brand.com` | `marketing` | DELETE |
| `john.smith@company.com` | (none) | Continue to next rule |

---

### 3. Delete Known Senders (`DELETE_KNOWN_SENDERS`)

If the sender's **display name** contains any of these, the email is deleted.

```vba
Public Const DELETE_KNOWN_SENDERS As String = "LinkedIn Job Alerts,edX,Cathay Pacific,HKBN,MyLink,WIRED Daily,Coursera,Udemy,Medium Daily Digest"
```

**Use for:**
- Specific services you don't want emails from
- Bulk senders with recognizable names
- Company names that send spam

**Tips:**
- Use exact names as they appear in your inbox
- Can be partial matches (e.g., "LinkedIn" matches "LinkedIn Job Alerts")

---

### 4. Delete Subject Patterns (`DELETE_SUBJECT_PATTERNS`)

If the email **subject** contains any of these keywords, it's deleted.

```vba
Public Const DELETE_SUBJECT_PATTERNS As String = "優惠,offer,digest,newsletter,unsubscribe,job alert,weekly roundup,daily digest,promotional,special offer,limited time,act now,don't miss"
```

**Use for:**
- Promotional keywords
- Digest/roundup emails
- Spam trigger words

**Supports:**
- Multiple languages (e.g., Chinese "優惠" = "discount/offer")
- Case-insensitive matching

---

### 5. Name Patterns (`NAME_PATTERNS`)

If your **name** appears in the subject or email body, the email is kept.

```vba
Public Const NAME_PATTERNS As String = "Xu Xin,XuXin,Xuxin,Xin Xu,Professor Xu,Prof. Xu,Prof Xu,Dr. Xu,Dr Xu,Mr. Xu,Mr Xu"
```

**Include:**
- Full name variations
- Title + name combinations
- Common misspellings

**How it's used:**
1. Checks if subject contains any pattern
2. Checks if body's first 200 characters contain any pattern
3. If found → KEEP the email

---

### 6. Greeting Patterns (`GREETING_PATTERNS`)

If the email **body starts with** any of these greetings, it's kept.

```vba
Public Const GREETING_PATTERNS As String = "Dear Professor Xu,Dear Prof. Xu,Dear Prof Xu,Dear Dr. Xu,Dear Dr Xu,Dear Xin,Hi Xin,Hello Xin,Dear Head,Dear Director"
```

**Tips:**
- Include formal and informal greetings
- Add position-based greetings (Director, Head, Chair)
- Whitespace at start of email is trimmed automatically

---

### 7. Organizational Tags (`POLYU_TAGS`)

Emails with these **tags in the subject** are kept (institutional importance).

```vba
Public Const POLYU_TAGS As String = "[MM],[HRO],[CUS],ToXX"
```

**Common formats:**
- `[DEPT]` - Department codes
- `[URGENT]` - Priority tags
- Custom institutional prefixes

---

### 8. VIP Keywords (`VIP_SUBJECT_KEYWORDS`)

If the **subject** contains any of these keywords, the email is kept.

```vba
Public Const VIP_SUBJECT_KEYWORDS As String = "thesis,dissertation,supervision,urgent,deadline,review request,paper submission,grant,conference,publication,meeting request,appointment,interview"
```

**Categories to consider:**
- Academic: thesis, dissertation, publication, grant
- Professional: deadline, interview, appointment
- Communication: meeting request, urgent, important

---

## Classification Priority

The filter checks patterns in this order (first match wins):

0. **Learned Sender Rule** → KEEP or DELETE (self-improving, highest priority)
1. **Protected Domain** → Move to "II" folder
2. **Personally Addressed** → KEEP
3. **Organizational Tags** → KEEP
4. **VIP Subject Keywords** → KEEP
5. **Reply Chain (RE:)** → KEEP
6. **Forward Chain (FW:)** → KEEP
7. **Known Spam Senders** → DELETE
8. **Sender Email Patterns** → DELETE
9. **Subject Patterns** → DELETE
10. **No Match** → LLM_REVIEW (or "I" folder)

---

## Pattern Syntax

### Comma-Separated Lists

All patterns are comma-separated strings:

```vba
' Correct
Public Const PATTERNS As String = "pattern1,pattern2,pattern3"

' Also correct (with spaces after commas - they're trimmed)
Public Const PATTERNS As String = "pattern1, pattern2, pattern3"

' INCORRECT - don't use newlines
Public Const PATTERNS As String = "pattern1," & _
    "pattern2"  ' This won't work correctly
```

### Case Sensitivity

- Most patterns are **case-insensitive** (e.g., "NEWSLETTER" matches "newsletter")
- `POLYU_TAGS` is checked with case sensitivity for exact tag matching

### Partial Matching

Most patterns use **partial matching** (contains):
- Pattern "news" matches "newsletter", "breaking news", "newsroom"
- Pattern "promo" matches "promotional", "promo code", "company-promo"

---

## Adding New Patterns

### To add a protected domain:

```vba
' Before
Public Const PROTECTED_DOMAINS As String = "substack.com,reddit.com"

' After
Public Const PROTECTED_DOMAINS As String = "substack.com,reddit.com,newsite.com"
```

### To add a delete pattern:

```vba
' Before
Public Const DELETE_SUBJECT_PATTERNS As String = "offer,digest"

' After
Public Const DELETE_SUBJECT_PATTERNS As String = "offer,digest,flash sale,clearance"
```

---

## Testing Your Patterns

After editing patterns:

1. **Save** the Config module (Ctrl+S)
2. Run `FilterExistingDryRun` to preview
3. Check the Immediate Window (Ctrl+G) for results
4. Look for unexpected classifications
5. Adjust patterns as needed
6. Repeat until satisfied

### Check Specific Pattern Matching

In the Immediate Window, you can test patterns:

```vba
? ContainsAny("john@noreply.company.com", DELETE_SENDER_PATTERNS)
' Returns: True

? ContainsAny("important-meeting@company.com", DELETE_SENDER_PATTERNS)
' Returns: False

? IsProtectedDomain("substack.com")
' Returns: True
```

---

## Common Pattern Mistakes

### ❌ Too Broad

```vba
' This will delete legitimate emails from "company-info@..."
Public Const DELETE_SENDER_PATTERNS As String = "info"
```

### ✅ More Specific

```vba
' Better - targets automated addresses
Public Const DELETE_SENDER_PATTERNS As String = "noreply,no-reply,donotreply"
```

---

### ❌ Missing Variations

```vba
' Won't catch "Dr Xu" (no period)
Public Const NAME_PATTERNS As String = "Dr. Xu"
```

### ✅ Include Variations

```vba
' Catches both with and without period
Public Const NAME_PATTERNS As String = "Dr. Xu,Dr Xu"
```

---

## Example: Academic Configuration

```vba
' For a university professor
Public Const PROTECTED_DOMAINS As String = "arxiv.org,researchgate.net,academia.edu,springer.com,ieee.org,acm.org"

Public Const NAME_PATTERNS As String = "Prof. Smith,Professor Smith,Dr. Smith,John Smith,J. Smith"

Public Const GREETING_PATTERNS As String = "Dear Professor Smith,Dear Prof. Smith,Dear Dr. Smith,Dear John,Hi John,Hello Professor"

Public Const VIP_SUBJECT_KEYWORDS As String = "thesis,dissertation,defense,publication,manuscript,revision,grant,NSF,NIH,proposal,deadline,conference,ICML,NeurIPS,review request,referee,editor"

Public Const DELETE_KNOWN_SENDERS As String = "Beall's List,Predatory Journals,OMICS,MDPI"

Public Const DELETE_SUBJECT_PATTERNS As String = "call for papers,invitation to submit,special issue,guest editor,waived APC,open access opportunity"
```

---

## Learned Sender Rules (Self-Improving)

Unlike the static patterns above, learned rules are created by **dragging emails** into special folders:

| Folder | Effect |
|--------|--------|
| **III** | Always **KEEP** emails from that sender |
| **IIII** | Always **DELETE** emails from that sender |

Learned rules have **Rule 0 priority** — they override ALL static patterns in Config.bas. This means if you drag a "noreply@company.com" email into III, future emails from that sender will be kept even though "noreply" normally triggers deletion.

### How It Works

1. You drag an email into III or IIII
2. The sender's email address is recorded in a text file
3. The rule is immediately active in memory (no restart needed)
4. On next Outlook startup, all rules are reloaded from file

### Data File

Rules are stored at: `%APPDATA%\OutlookEmailFilter\learned_senders.txt`

Format (pipe-delimited, append-only):
```
sender@example.com|KEEP|2026-02-08 14:30:22
spammer@junk.com|DELETE|2026-02-08 15:01:05
```

- Last entry per sender wins (you can override a previous decision)
- Lines starting with `#` are treated as comments
- You can edit this file manually with a text editor

### Changing a Learned Rule

To change a rule (e.g., from DELETE to KEEP): simply drag another email from that sender into the opposite folder. The new entry is appended and takes precedence.

To remove all learned rules: delete the file at the path shown by `ShowLearnedSenders`.

### Diagnostics

| Macro | Purpose |
|-------|---------|
| `ShowLearnedSenders` | Display rule count and file path |
| `ReloadLearnedSenders` | Force reload from file (after manual edits) |

In dry-run output, learned rules show special icons:
- `[+LR]` = Kept by learned rule
- `[xLR]` = Deleted by learned rule

---

## Backup Your Configuration

Before making changes, copy your current `Config.bas` to a backup:

1. In VBA Editor, right-click the **Config** module
2. Select **Export File...**
3. Save as `Config_backup_YYYYMMDD.bas`
