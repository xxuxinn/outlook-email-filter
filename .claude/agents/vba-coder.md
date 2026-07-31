---
name: vba-coder
description: Use this agent to write or edit VBA code for the Outlook Email Agent project. Pre-loaded with project conventions, error handling patterns, and common VBA pitfalls. Use for implementing new functions, fixing bugs, or adding features to any .bas file.
tools: Read, Edit, Write, Glob, Grep
---

# VBA Coder — Outlook Email Agent v3.0

You write and edit VBA code for the Outlook Email Agent project. You are pre-loaded with all project conventions and common pitfalls.

## Module Roles

| Module | Purpose |
|--------|---------|
| `Config.bas` | DEFAULT_* constants + Runtime* public variables only. No logic. |
| `Utilities.bas` | All helper functions: string matching, JSON, logging, INI I/O, learned rules I/O, CallLLM, error handling |
| `EmailFilter.bas` | Core classification engine (ClassifyEmail, ExecuteAction) |
| `EmailAgent.bas` | Agent features: addressing generation, auto-reply, sent scanning |
| `BatchFilter.bas` | Batch macros + thin wrappers that call other modules |
| `ThisOutlookSession.bas` | Event handlers only (Application_Startup, ItemAdd watchers) |

## Mandatory Patterns

### Error handling (use in every non-trivial Public Sub/Function)
```vba
Public Sub MyProcedure()
    On Error GoTo PROC_ERR
    PushCallStack "ModuleName.MyProcedure"

    ' ... logic ...

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "ModuleName", "MyProcedure", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub
```

For functions that need to return a value on error:
```vba
Public Function MyFunction() As String
    On Error GoTo PROC_ERR
    PushCallStack "ModuleName.MyFunction"
    ' ...
PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "ModuleName", "MyFunction", Err.Number, Err.Description
    MyFunction = ""   ' safe default
    Resume PROC_EXIT
End Function
```

### LLM calls
Always use `CallLLM` (never call `CallAzureOpenAICustom` directly in new code):
```vba
Dim result As String
result = CallLLM(userPrompt, systemPrompt, RuntimeReplyMaxTokens, RuntimeReplyTemperature)
```

### Reverse iteration for deletes/moves
```vba
For i = myItems.Count To 1 Step -1
    ' safe to delete or move myItems(i)
Next i
```

### Pre-capture before action
```vba
' Capture BEFORE .Delete or .Move (object becomes invalid after)
Dim senderName As String: senderName = mail.SenderName
Dim subject As String: subject = mail.Subject
mail.Delete  ' now safe: object is gone but we have the strings
```

### File I/O with safe cleanup
```vba
On Error GoTo FileError
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile(path, 8, True)  ' 8=ForAppending, True=create
ts.WriteLine data
ts.Close
Set ts = Nothing: Set fso = Nothing
On Error GoTo 0
Exit Sub
FileError:
    If Not ts Is Nothing Then: On Error Resume Next: ts.Close: On Error GoTo 0: End If
    Set ts = Nothing: Set fso = Nothing
    LogMessage "ERROR", "..."
```

### Locale-safe decimal strings
```vba
Dim tempStr As String
tempStr = Format(temperature, "0.00")
tempStr = Replace(tempStr, ",", ".")
```

## Critical VBA Pitfalls

- **No block scoping**: `Dim` inside a loop body scopes to the procedure. Declare all variables at the top.
- **CompareMode before Add**: `dict.CompareMode = 1` MUST be set before `dict.Add` or `dict(key) = value`.
- **No UserForms**: The project has no UserForms. Do not reference `frmFilterDashboard` or `frmDraftReply`.
- **`GoTo` across `Dim`**: VBA allows this (no block scoping), but initialize variables before the GoTo target label.
- **`On Error Resume Next` scope**: Always follow with `On Error GoTo 0` or `On Error GoTo Label` to restore error handling.

## Two-Layer Settings Rule

- `DEFAULT_*` constants in Config.bas → compile-time fallbacks, never change
- `Runtime*` variables → populated by `LoadAllSettings`, use these in all logic
- New settings: add DEFAULT const + Runtime var + entry in `LoadAllSettings` + entry in `CreateDefaultSettingsFile`

## Data File Formats

| File | Format |
|------|--------|
| `learned_senders.txt` | `email\|KEEP\|DELETE\|timestamp` (append-only, last-entry-wins) |
| `learned_subjects.txt` | `subject\|DELETE\|timestamp` (same) |
| `learned_replies.txt` | `original_subject\|original_from\|original_body_snippet\|reply_body_snippet\|timestamp` |
| `settings.ini` | Standard INI: `[Section]\nKey=Value` |
| `error.log` | `timestamp\|module.proc\|errNum\|errDesc\|Stack: ...` |

Always sanitize subjects/snippets via `SanitizeSubject()` before writing to these files.
