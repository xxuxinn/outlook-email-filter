# Diagrams — Outlook Email Agent v3.0

All diagrams are authored in Mermaid (`.mmd` source) with pre-rendered `.svg` output.
The three diagrams in sections 1–3 below are also embedded inline so GitHub renders them directly on this page.

> **v3.1 note**: these diagrams show the v3.0 layout. Since then the command bridge moved
> from `Utilities.bas` into a dedicated `Bridge.bas`, and three modules were added
> (`AgentMemory.bas` — decision log/sender history, `EmailDigest.bas` — daily digest +
> rule mining, plus an `mcp/` server). The classification chain (Rules 0–10) is unchanged,
> except Rule 10's LLM step is now structured JSON with a confidence gate.

## Index

| Diagram | Source | Rendered | Description |
|---------|--------|----------|-------------|
| Full-system architecture | [architecture-full-system.mmd](architecture-full-system.mmd) | [SVG](architecture-full-system.svg) | Complete system: VBA modules, Web UI, data files, cloud sync, LLM providers |
| Email agent architecture | [architecture-email-agent.mmd](architecture-email-agent.mmd) | [SVG](architecture-email-agent.svg) | Module-level view: VBA module dependency chain + Web UI |
| Email arrival flow | [flow-email-arrival.mmd](flow-email-arrival.mmd) | [SVG](flow-email-arrival.svg) | Classification trace for a new email through the 10-rule chain |
| Classification chain | [activity-classification-chain.mmd](activity-classification-chain.mmd) | [SVG](activity-classification-chain.svg) | Compact activity view of the rule chain (Rules 0–10) |
| User action flows | [flow-user-actions.mmd](flow-user-actions.mmd) | [SVG](flow-user-actions.svg) | Learning flows A–E: learn KEEP/DELETE/subject rules, draft auto-reply, learn reply style |
| Command bridge sequence | [sequence-command-bridge.mmd](sequence-command-bridge.mmd) | [SVG](sequence-command-bridge.svg) | Web UI → Flask → JSON file bridge → VBA poller round trip |
| Email lifecycle states | [state-email-lifecycle.mmd](state-email-lifecycle.mmd) | [SVG](state-email-lifecycle.svg) | State diagram of an email from arrival through classification to final folder |

---

## 1. System Architecture

```mermaid
graph TB
    subgraph OUTLOOK["Microsoft Outlook Process"]
        direction TB

        subgraph EVENTS["ThisOutlookSession.bas -- Event Handlers Only"]
            STARTUP["Application_Startup\n1 LoadAllSettings\n2 StartCommandPollerStd\n3 Init folder watchers\n4 SyncLearnedRulesAuto"]
            INBOX_EVT["inboxItems_ItemAdd\nReal-time filter trigger"]
            LK_EVT["learnKeepItems_ItemAdd"]
            LD_EVT["learnDeleteItems_ItemAdd"]
            LS_EVT["learnSubjectDeleteItems_ItemAdd"]
            LR_EVT["learnReplyItems_ItemAdd"]
            QUIT["Application_Quit\nStopPoller + SyncAuto"]
        end

        subgraph CONFIG["Config.bas -- Constants and Variables Only"]
            DEFAULTS["DEFAULT_* constants\nCompile-time fallbacks"]
            RUNTIME["Runtime* variables\nLoaded from settings.ini"]
        end

        subgraph UTILS["Utilities.bas -- Shared Infrastructure"]
            direction LR
            STRING["String helpers\nContainsAny, StartsWithAny\nSanitizeSubject, Truncate"]
            LLM["CallLLM\nMulti-provider router"]
            IO["File I/O\nINI read/write\nLearned rules I/O\nReply pairs I/O"]
            ERR["Error handling\nPushCallStack\nLogError, LogMessage"]
            BRIDGE["Command Bridge\nSetTimer poller\nWriteResultFile"]
        end

        subgraph FILTER["EmailFilter.bas -- Classification Engine"]
            CLASSIFY["ClassifyEmail\n10-rule priority chain"]
            EXEC["ExecuteAction\nKEEP / DELETE /\nMOVE_II / LLM_REVIEW"]
            LLMWRAP["LLM Wrappers\nBuildEmailPrompt\nSummarizeSelectedEmail"]
        end

        subgraph AGENT["EmailAgent.bas -- AI Agent Features"]
            PATTERNS["GenerateAddressingPatterns"]
            REPLY["DraftAutoReply\nFew-shot reply engine"]
            SCAN["ScanSentForReplyPatterns"]
            DRAFT["DraftReplyForSelected"]
        end

        subgraph BATCH["BatchFilter.bas -- Batch Macros"]
            BATCHMAC["FilterExisting / FilterAllFolders\nFilterSelectedEmails\nFilterCurrentFolder\nFilterLastNDays\nDryRun / Report"]
        end
    end

    subgraph WEBUI["Web UI -- Python Flask localhost:5000"]
        direction TB
        SERVER["server.py\nFlask routes"]
        BRIDGEPY["bridge.py\nJSON command files"]
        SETTINGS["settings_manager.py\nINI read/write"]
        CHAT["chat.py\nKeyword to action parser"]
        SPA["static/\nindex.html + app.js\n6-tab SPA"]
    end

    subgraph STORAGE["Data Files -- %APPDATA%\OutlookEmailFilter"]
        INI["settings.ini"]
        SENDERS["learned_senders.txt"]
        SUBJECTS["learned_subjects.txt"]
        REPLIES["learned_replies.txt"]
        ERRORLOG["error.log"]
        CMDS["commands/*.json"]
    end

    subgraph CLOUD["Cloud Sync -- OneDrive"]
        CLOUDFILES["OutlookEmailFilter/\nlearned_senders.txt\nlearned_subjects.txt\nlearned_replies.txt"]
    end

    subgraph LLM_PROVIDERS["LLM Providers"]
        LOCAL["Ollama / LM Studio\nlocalhost OpenAI-compat"]
        AZURE["Azure OpenAI\napi-key auth"]
        CLAUDEAI["Anthropic Claude\nx-api-key auth"]
        OPENAI["OpenAI-compat\nOpenRouter, Groq, etc."]
    end

    STARTUP --> CONFIG
    STARTUP --> UTILS
    CONFIG --> UTILS
    UTILS --> FILTER
    UTILS --> AGENT
    FILTER --> BATCH
    AGENT --> BATCH
    INBOX_EVT --> CLASSIFY
    CLASSIFY --> EXEC
    BATCHMAC --> CLASSIFY

    LLM --> LOCAL
    LLM --> AZURE
    LLM --> CLAUDEAI
    LLM --> OPENAI

    IO --> INI
    IO --> SENDERS
    IO --> SUBJECTS
    IO --> REPLIES
    ERR --> ERRORLOG

    BRIDGE --> CMDS
    BRIDGEPY --> CMDS

    IO --> CLOUDFILES

    SPA --> SERVER
    SERVER --> BRIDGEPY
    SERVER --> SETTINGS
    SERVER --> CHAT
    SETTINGS --> INI

    LK_EVT --> IO
    LD_EVT --> IO
    LS_EVT --> IO
    LR_EVT --> IO

    REPLY --> LLM
    PATTERNS --> LLM
    LLMWRAP --> LLM
```

---

## 2. Email Arrival Flow — Classification Trace

```mermaid
flowchart TD
    START(["New email arrives in Inbox"])
    TYPECHECK{"Is it a MailItem?"}
    CLASSIFY["ClassifyEmail -- EmailFilter.bas"]

    R0{"Rule 0\nLearned sender rule?"}
    R0_KEEP["KEEP\nlearned sender"]
    R0_DEL["DELETE\nlearned sender"]

    R05{"Rule 0.5\nLearned subject DELETE?"}
    R05_DEL["DELETE\nlearned subject"]

    R1{"Rule 1\nProtected domain?"}
    R1_MOVE["MOVE_II\nto Protected folder"]

    R2{"Rule 2\nPersonally addressed?"}
    R2_KEEP["KEEP\nname/greeting match"]

    R3{"Rule 3\nOrg tags?"}
    R3_KEEP["KEEP\norg tag"]

    R4{"Rule 4\nVIP subject keywords?"}
    R4_KEEP["KEEP\nVIP keyword"]

    R5{"Rule 5\nRE: chain?"}
    R5_KEEP["KEEP\nreply chain"]

    R6{"Rule 6\nFW: chain?"}
    R6_KEEP["KEEP\nforwarded"]

    R7{"Rule 7\nKnown spam sender?"}
    R7_DEL["DELETE\nknown sender"]

    R8{"Rule 8\nSender email pattern?"}
    R8_DEL["DELETE\nsender pattern"]

    R9{"Rule 9\nSubject pattern?"}
    R9_DEL["DELETE\nsubject pattern"]

    R10["Rule 10\nNo rule matched\nLLM_REVIEW"]

    ACT_KEEP["KEEP\nLeave in Inbox"]
    ACT_DELETE["DELETE\nmail.Delete"]
    ACT_MOVE["MOVE_II\nto Protected folder"]
    ACT_LLM{"LLM enabled?"}
    ACT_REVIEW["LLM_REVIEW\nto Review folder"]
    ACT_LLMCALL["CallLLM\nClassify via LLM"]

    AUTOREPLY{"AutoReplyOnArrival\nenabled?"}
    DRAFTREPLY["DraftAutoReply\nFew-shot reply to Drafts"]
    LOG["LogMessage\naction + sender + subject"]
    DONE(["Processing complete"])
    SKIP(["Ignored -- not a MailItem"])

    START --> TYPECHECK
    TYPECHECK -- No --> SKIP
    TYPECHECK -- Yes --> CLASSIFY
    CLASSIFY --> R0

    R0 -- KEEP --> R0_KEEP
    R0 -- DELETE --> R0_DEL
    R0 -- "no rule" --> R05

    R05 -- DELETE --> R05_DEL
    R05 -- "no match" --> R1

    R1 -- Yes --> R1_MOVE
    R1 -- No --> R2

    R2 -- Yes --> R2_KEEP
    R2 -- No --> R3

    R3 -- Yes --> R3_KEEP
    R3 -- No --> R4

    R4 -- Yes --> R4_KEEP
    R4 -- No --> R5

    R5 -- Yes --> R5_KEEP
    R5 -- No --> R6

    R6 -- Yes --> R6_KEEP
    R6 -- No --> R7

    R7 -- Yes --> R7_DEL
    R7 -- No --> R8

    R8 -- Yes --> R8_DEL
    R8 -- No --> R9

    R9 -- Yes --> R9_DEL
    R9 -- No --> R10

    R0_KEEP --> ACT_KEEP
    R2_KEEP --> ACT_KEEP
    R3_KEEP --> ACT_KEEP
    R4_KEEP --> ACT_KEEP
    R5_KEEP --> ACT_KEEP
    R6_KEEP --> ACT_KEEP

    R0_DEL --> ACT_DELETE
    R05_DEL --> ACT_DELETE
    R7_DEL --> ACT_DELETE
    R8_DEL --> ACT_DELETE
    R9_DEL --> ACT_DELETE

    R1_MOVE --> ACT_MOVE

    R10 --> ACT_LLM
    ACT_LLM -- Yes --> ACT_LLMCALL
    ACT_LLM -- No --> ACT_REVIEW
    ACT_LLMCALL --> ACT_REVIEW

    ACT_KEEP --> AUTOREPLY
    AUTOREPLY -- Yes --> DRAFTREPLY
    AUTOREPLY -- No --> LOG
    DRAFTREPLY --> LOG

    ACT_DELETE --> LOG
    ACT_MOVE --> LOG
    ACT_REVIEW --> LOG

    LOG --> DONE
```

---

## 3. User Action Flows — Learning and Auto-Reply

### Flow A: Learn DELETE Rule

```mermaid
flowchart TD
    A1(["User drags email to LearnDelete"])
    A2["learnDeleteItems_ItemAdd\nThisOutlookSession.bas"]
    A3["GetSenderEmail -- extract SMTP address"]
    A4["RecordLearnedSender\nemail | DELETE | timestamp\nto learned_senders.txt"]
    A5{"Was it KEEP before?"}
    A6["DeleteSenderFromInbox\nDelete all emails from\nthis sender in Inbox"]
    A7["Log: LEARNED DELETE"]
    A8["Cloud sync on next startup/quit"]
    A9(["Future emails from this sender auto-DELETE"])

    A1 --> A2 --> A3 --> A4 --> A5
    A5 -- "Yes - rule reversal" --> A6 --> A7
    A5 -- "No / new rule" --> A7
    A7 --> A8 --> A9
```

### Flow B: Learn KEEP Rule

```mermaid
flowchart TD
    B1(["User drags email to LearnKeep"])
    B2["learnKeepItems_ItemAdd\nThisOutlookSession.bas"]
    B3["GetSenderEmail -- extract SMTP address"]
    B4["RecordLearnedSender\nemail | KEEP | timestamp\nto learned_senders.txt"]
    B5{"Was it DELETE before?"}
    B6["RestoreSenderFromDeleted\nRescue emails from\nDeleted Items to Inbox"]
    B7["Log: LEARNED KEEP"]
    B8(["Future emails from this sender auto-KEEP"])

    B1 --> B2 --> B3 --> B4 --> B5
    B5 -- "Yes - rule reversal" --> B6 --> B7
    B5 -- "No / new rule" --> B7
    B7 --> B8
```

### Flow C: Learn Subject DELETE Rule

```mermaid
flowchart TD
    C1(["User drags email to LearnSubjectDelete"])
    C2["learnSubjectDeleteItems_ItemAdd\nThisOutlookSession.bas"]
    C3["SanitizeSubject\nStrip CR/LF/pipe/null"]
    C4["RecordLearnedSubject\nsubject | DELETE | timestamp\nto learned_subjects.txt"]
    C5["Log: LEARNED SUBJECT DELETE"]
    C6(["Future emails with matching\nsubject substring auto-DELETE\nat Rule 0.5"])

    C1 --> C2 --> C3 --> C4 --> C5 --> C6
```

### Flow D: Draft Auto-Reply

```mermaid
flowchart TD
    D1(["User runs DraftReplyForSelected\nor auto-triggered on KEEP"])
    D2["LoadRecentReplyPairs\nRead last N entries from\nlearned_replies.txt"]
    D3["Build few-shot prompt\nSystem: reply persona\nExamples: learned pairs\nCurrent: email to reply to"]
    D4["CallLLM\nTemperature: 0.7  MaxTokens: 800"]
    D5{"Provider?"}
    D6["Ollama / LM Studio"]
    D7["Azure OpenAI"]
    D8["Anthropic Claude"]
    D9["OpenRouter / Groq"]
    D10["Create reply MailItem\nmail.Reply then set .Body"]
    D11["Save to Drafts folder"]
    D12(["Draft ready in Outlook Drafts"])

    D1 --> D2 --> D3 --> D4 --> D5
    D5 --> D6 --> D10
    D5 --> D7 --> D10
    D5 --> D8 --> D10
    D5 --> D9 --> D10
    D10 --> D11 --> D12
```

### Flow E: Learn Reply Style

```mermaid
flowchart TD
    E1(["User drags sent reply to LearnReply"])
    E2["learnReplyItems_ItemAdd\nThisOutlookSession.bas"]
    E3["Strip RE:/AW: prefix\nto get originalSubject"]
    E4["ExtractMyReplyFromBody\nText before delimiter"]
    E5["ExtractOriginalFromBody\nText after delimiter, first 500 chars"]
    E6["Get original sender\nmail.Recipients 1 .Name"]
    E7["RecordLearnedReply\nsubject | from | orig_body | reply_body | timestamp\nto learned_replies.txt"]
    E8(["Reply pair saved for few-shot learning"])

    E1 --> E2 --> E3 --> E4 --> E5 --> E6 --> E7 --> E8
```
