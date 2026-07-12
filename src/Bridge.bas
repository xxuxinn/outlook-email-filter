'===============================================================================
' Bridge.bas - Web UI / MCP Command Bridge v3.1
'===============================================================================
' File-based IPC with the Python Flask Web UI and the MCP server.
' Commands are JSON files in %APPDATA%\OutlookEmailFilter\commands\:
'   <id>.json   {"id":"...","macro":"...","args":{...}}   (written by Python)
'   <id>.result {"id":"...","status":"ok|error","output":"..."} (written here)
'
' Extracted from Utilities.bas in v3.1 so the helper module no longer calls
' upward into every layer (the dispatcher references EmailFilter, EmailAgent,
' EmailDigest, and BatchFilter, which sit above Utilities in the layering).
'
' v3.1 fixes over the old dispatcher:
'   - Re-entrancy guard: modal dialogs / DoEvents pump WM_TIMER, so the 2 s
'     poller could previously start a second macro mid-run.
'   - Orphan-timer self-kill: after an unhandled error resets VBA state, the
'     OS timer kept firing into a module that lost its timer id. The callback
'     now kills any timer whose id it no longer recognises.
'   - Honest results: every command returns real counts / "ERROR: ..." from
'     headless *Core functions instead of hardcoded "completed." strings.
'   - Unreadable command files are quarantined (.bad) instead of retried
'     every 2 seconds forever.
'   - Scheduler: piggybacks on the poller to run the daily digest and weekly
'     rule mining without any external cron.
'===============================================================================

Option Explicit

' Poller state
Private pollerRunningFlag As Boolean
Private pollerBusy As Boolean          ' re-entrancy guard for dispatches
Private schedulerTickCounter As Long   ' poller ticks since last scheduler check

' Windows API timer (Outlook has no Application.OnTime)
#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private pollerTimerId As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
    Private pollerTimerId As Long
#End If

'-------------------------------------------------------------------------------
' COMMANDS DIRECTORY AND RESULT FILES
'-------------------------------------------------------------------------------

' Return path to the commands directory (auto-creates it)
Public Function GetCommandsDir() As String
    Dim dir As String
    dir = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\commands"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then
        On Error Resume Next
        fso.CreateFolder dir
        On Error GoTo 0
    End If
    Set fso = Nothing
    GetCommandsDir = dir
End Function

' Write a result JSON file for a completed command (UTF-8 via ADODB.Stream).
' status: "ok" or "error"; output: text to return to the Web UI / MCP client.
Public Sub WriteResultFile(ByVal cmdId As String, ByVal status As String, ByVal output As String)
    Dim stm As Object
    Dim resultPath As String
    Dim jsonLine As String

    resultPath = GetCommandsDir() & "\" & cmdId & ".result"
    jsonLine = "{""id"":""" & EscapeJSON(cmdId) & """,""status"":""" & status & _
               """,""output"":""" & EscapeJSON(output) & """}"

    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2            ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText jsonLine
    stm.SaveToFile resultPath, 2  ' adSaveCreateOverWrite
    stm.Close
    On Error GoTo 0
    Set stm = Nothing
End Sub

'-------------------------------------------------------------------------------
' COMMAND POLLER (Win32 SetTimer, fires every 2 s)
'-------------------------------------------------------------------------------

Public Sub StartCommandPollerStd()
    If pollerRunningFlag Then Exit Sub

    ' Defensive: kill any timer we still know about before starting a new one
    If pollerTimerId <> 0 Then
        KillTimer 0, pollerTimerId
        pollerTimerId = 0
    End If

    pollerRunningFlag = True
    pollerTimerId = SetTimer(0, 0, 2000, AddressOf PollerCallback)
    If pollerTimerId = 0 Then
        pollerRunningFlag = False
        LogMessage "ERROR", "Failed to start command poller timer"
    Else
        LogMessage "INFO", "Web UI command poller started (timer ID: " & pollerTimerId & ")"
    End If
End Sub

Public Sub StopCommandPollerStd()
    pollerRunningFlag = False
    If pollerTimerId <> 0 Then
        KillTimer 0, pollerTimerId
        pollerTimerId = 0
    End If
    LogMessage "INFO", "Web UI command poller stopped"
End Sub

#If VBA7 Then
Private Sub PollerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                           ByVal nIDEvent As LongPtr, ByVal dwTime As Long)
#Else
Private Sub PollerCallback(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal nIDEvent As Long, ByVal dwTime As Long)
#End If
    On Error Resume Next

    ' Orphan-timer recovery: if VBA state was reset (unhandled error -> "End"),
    ' pollerRunningFlag/pollerTimerId are zeroed while the OS timer keeps
    ' firing. Kill any timer we no longer recognise so it cannot leak or
    ' stack with a restarted poller.
    If Not pollerRunningFlag Or nIDEvent <> pollerTimerId Then
        KillTimer 0, nIDEvent
        Exit Sub
    End If

    PollForCommandsTimer
    On Error GoTo 0
End Sub

Public Sub PollForCommandsTimer()
    If Not pollerRunningFlag Then Exit Sub

    ' Re-entrancy guard: dispatched macros can call DoEvents (or, in legacy
    ' paths, show dialogs), which pumps WM_TIMER and re-enters this Sub.
    ' Without the guard, two batch jobs could interleave deletions.
    If pollerBusy Then Exit Sub
    pollerBusy = True
    On Error GoTo PollerExit

    ' Scheduled jobs (daily digest / weekly rule mining) — checked ~every 60 s
    schedulerTickCounter = schedulerTickCounter + 1
    If schedulerTickCounter >= 30 Then
        schedulerTickCounter = 0
        CheckScheduledJobs
    End If

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim cmdDir As String
    Dim cmdId As String
    Dim macroName As String
    Dim output As String
    Dim ts As Object
    Dim content As String
    Dim status As String

    cmdDir = GetCommandsDir()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(cmdDir) Then
        Set fso = Nothing
        GoTo PollerExit
    End If

    Set folder = fso.GetFolder(cmdDir)

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "json" Then
            cmdId = fso.GetBaseName(file.Name)

            On Error Resume Next
            Set ts = fso.OpenTextFile(file.Path, 1)
            If Err.Number <> 0 Then
                ' Unreadable (locked mid-write or corrupt): quarantine after
                ' it has had time to finish being written, so it is not
                ' retried every 2 s forever with no result for the client.
                Err.Clear
                If DateDiff("s", file.DateLastModified, Now) > 10 Then
                    fso.MoveFile file.Path, file.Path & ".bad"
                    ' Only report quarantine if the move actually succeeded —
                    ' otherwise the file is still locked and will be retried
                    If Err.Number = 0 Then
                        WriteResultFile cmdId, "error", "Command file unreadable; quarantined as .bad"
                    Else
                        Err.Clear
                    End If
                End If
                On Error GoTo PollerExit
                GoTo NextFile
            End If
            content = ts.ReadAll
            ts.Close
            Set ts = Nothing
            On Error GoTo PollerExit

            ' Delete before dispatch: prevents double-execution if the macro
            ' itself crashes hard (the command is not retried).
            On Error Resume Next
            fso.DeleteFile file.Path
            On Error GoTo PollerExit

            macroName = ExtractJSONStringValue(content, "macro")
            If Len(macroName) = 0 Then
                WriteResultFile cmdId, "error", "Could not parse macro name from command"
                GoTo NextFile
            End If

            output = DispatchMacroStd(macroName, content)
            status = "ok"
            If Left(output, 6) = "ERROR:" Then status = "error"

            WriteResultFile cmdId, status, output
        End If
NextFile:
    Next file

    Set fso = Nothing

PollerExit:
    pollerBusy = False
End Sub

'-------------------------------------------------------------------------------
' SCHEDULER (daily digest + weekly rule mining, no external cron needed)
'-------------------------------------------------------------------------------

Private Sub CheckScheduledJobs()
    On Error GoTo PROC_ERR
    PushCallStack "Bridge.CheckScheduledJobs"

    Dim todayStr As String
    todayStr = Format(Date, "yyyy-mm-dd")

    ' Daily digest: once per day, after DigestHour
    If RuntimeEnableDailyDigest Then
        If Hour(Now) >= RuntimeDigestHour Then
            If ReadINISetting("Digest", "LastDigestDate", "") <> todayStr Then
                ' Stamp BEFORE running so a failing digest cannot retry-loop
                WriteINISetting "Digest", "LastDigestDate", todayStr
                LogMessage "INFO", "Scheduler: generating daily digest"
                Dim digestResult As String
                digestResult = GenerateDailyDigestCore()
                LogMessage "INFO", "Scheduler digest result: " & Left(digestResult, 120)
            End If
        End If
    End If

    ' Weekly rule mining: every 7+ days, after DigestHour
    If RuntimeEnableRuleMining Then
        If Hour(Now) >= RuntimeDigestHour Then
            Dim lastMining As String
            lastMining = ReadINISetting("Digest", "LastRuleMiningDate", "")
            Dim daysSince As Long
            daysSince = 9999
            If IsDate(lastMining) Then daysSince = DateDiff("d", CDate(lastMining), Date)
            If daysSince >= 7 Then
                WriteINISetting "Digest", "LastRuleMiningDate", todayStr
                LogMessage "INFO", "Scheduler: running weekly rule mining"
                Dim miningResult As String
                miningResult = ProposeRulesCore()
                LogMessage "INFO", "Scheduler mining result: " & Left(miningResult, 120)
            End If
        End If
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Bridge", "CheckScheduledJobs", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' MACRO DISPATCHER (hard allowlist; unknown macros are rejected)
'-------------------------------------------------------------------------------

Private Function DispatchMacroStd(ByVal macroName As String, Optional ByVal rawJson As String = "") As String
    Dim result As String

    On Error GoTo DispatchError

    Select Case macroName
        ' --- Filtering ---
        Case "FilterExistingDryRun"
            result = CaptureFilterDryRunStd()

        Case "FilterExistingEmails"
            result = FilterExistingEmailsCore()

        Case "FilterAllFolders"
            result = FilterAllFoldersCore()

        Case "FilterCurrentFolder"
            result = FilterCurrentFolderCore()

        Case "FilterSelectedEmail", "FilterSelectedEmails"
            result = FilterSelectedEmailsCore()

        Case "FilterLastNDays"
            Dim daysArg As String
            Dim days As Long
            daysArg = ExtractJSONStringValue(rawJson, "days")
            days = 7
            If Len(daysArg) > 0 And IsNumeric(daysArg) Then days = CLng(daysArg)
            If days < 1 Or days > 365 Then
                result = "ERROR: days must be between 1 and 365"
            Else
                result = FilterLastNDaysCore(days)
            End If

        Case "BulkDeleteBySender"
            Dim patternArg As String
            patternArg = ExtractJSONStringValue(rawJson, "pattern")
            If Len(patternArg) >= 3 Then
                result = BulkDeleteBySenderCore(patternArg)
            Else
                result = "ERROR: BulkDeleteBySender requires a sender pattern of at least 3 characters."
            End If

        Case "MoveProtectedSources"
            result = MoveProtectedSourcesCore()

        Case "GenerateClassificationReport"
            result = GenerateClassificationReportCore()

        ' --- Digest / rule mining (v3.1) ---
        Case "GenerateDailyDigest"
            result = GenerateDailyDigestCore()

        Case "ProposeRules"
            result = ProposeRulesCore()

        ' --- Agent tools ---
        Case "ScanSentForReplyPatterns"
            result = ScanSentForReplyPatternsCore()

        Case "DraftReplyForSelected"
            result = DraftReplyForSelectedStd()

        Case "SummarizeSelectedEmail"
            result = SummarizeSelectedEmailStd()

        Case "DraftReplyToSelected"
            result = DraftReplyToSelectedStd()

        Case "GenerateAddressingPatterns"
            Dim nameArg As String, titleArg As String, roleArg As String
            nameArg = ExtractJSONStringValue(rawJson, "name")
            titleArg = ExtractJSONStringValue(rawJson, "title")
            roleArg = ExtractJSONStringValue(rawJson, "role")
            If Len(nameArg) = 0 Then
                result = "ERROR: GenerateAddressingPatterns requires a 'name' argument (plus optional 'title' and 'role')."
            Else
                result = GenerateAddressingPatternsStd(nameArg, titleArg, roleArg)
            End If

        ' --- Learned rules ---
        Case "ShowLearnedSenders"
            result = "Learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
                     "File: " & GetLearnedSendersFilePath()

        Case "ShowLearnedSendersList"
            result = BuildLearnedSendersListStd()

        Case "CleanLearnedSendersFile"
            result = CleanLearnedSendersFileCore()

        Case "ShowLearnedSubjectsList"
            result = BuildLearnedSubjectsListStd()

        Case "CleanLearnedSubjectsFile"
            result = CleanLearnedSubjectsFileCore()

        Case "ImportExistingLearnedFolders"
            result = ImportExistingLearnedFoldersCore()

        Case "ImportExistingLearnedSubjectFolder"
            result = ImportExistingLearnedSubjectFolderCore()

        Case "ReloadLearnedSenders"
            result = ReloadLearnedSendersCore()

        Case "ShowLearnedRepliesSummary"
            result = ShowLearnedRepliesSummaryCore()

        ' --- Server rules ---
        Case "ImportServerRules"
            result = ImportServerRulesCore()

        Case "ExportLearnedRulesToServer"
            result = ExportLearnedRulesToServerCore()

        ' --- Undo / recovery ---
        Case "RestoreFromReview"
            result = RestoreFromReviewCore()

        Case "RestoreDeletedKeepEmails"
            result = RestoreDeletedKeepEmailsCore()

        ' --- System ---
        Case "ShowVersionInfo"
            result = ShowVersionInfoCore()

        Case "ReinitializeFilter"
            ' Reloads settings + learned caches. Event handlers live in
            ' ThisOutlookSession and cannot be restarted from a timer callback;
            ' say so honestly instead of implying a full reinitialise.
            LoadAllSettings
            LoadLearnedSenders True
            LoadLearnedSubjects True
            result = "Settings and learned rules reloaded. (Event handlers unchanged — run " & _
                     "ThisOutlookSession.ReinitializeFilter in the VBA Immediate Window for a full restart.)"

        Case "DetectAndMigrateOldFolders"
            result = DetectAndMigrateOldFoldersCore()

        Case "EnableRealTimeFilter"
            result = "Cannot run from Web UI. In VBA Immediate Window (Ctrl+G), type:" & vbCrLf & _
                     "  ThisOutlookSession.EnableRealTimeFilter"

        Case "DisableRealTimeFilter"
            result = "Cannot run from Web UI. In VBA Immediate Window (Ctrl+G), type:" & vbCrLf & _
                     "  ThisOutlookSession.DisableRealTimeFilter"

        Case "SyncLearnedRules"
            result = SyncLearnedRulesCore()

        Case Else
            result = "ERROR: Unknown macro: " & macroName

    End Select

    DispatchMacroStd = result
    Exit Function

DispatchError:
    DispatchMacroStd = "ERROR: " & macroName & " failed: " & Err.Description
End Function

'-------------------------------------------------------------------------------
' BRIDGE-ONLY HELPERS (headless variants of interactive macros)
'-------------------------------------------------------------------------------

' Headless rule-list dumps for the bridge (the interactive Show*List macros in
' BatchFilter end in MsgBox, which would block the timer callback until someone
' clicks OK at the desktop — see the no-dialogs rule for dispatched commands).
Private Function BuildLearnedSendersListStd() As String
    On Error GoTo StdErr
    Dim dict As Object
    Set dict = GetLearnedSendersCacheCopy()

    Dim out As String
    Dim k As Variant
    Dim n As Long
    out = "Learned sender rules (" & dict.Count & "):" & vbCrLf
    n = 0
    For Each k In dict.keys
        n = n + 1
        If n > 500 Then
            out = out & "... (" & (dict.Count - 500) & " more)" & vbCrLf
            Exit For
        End If
        out = out & k & " -> " & dict(k) & vbCrLf
    Next k
    BuildLearnedSendersListStd = out
    Exit Function
StdErr:
    BuildLearnedSendersListStd = "ERROR: could not list sender rules: " & Err.Description
End Function

Private Function BuildLearnedSubjectsListStd() As String
    On Error GoTo StdErr
    Dim dict As Object
    Set dict = GetLearnedSubjectsCacheCopy()

    Dim out As String
    Dim k As Variant
    Dim n As Long
    out = "Learned subject rules (" & dict.Count & ", all DELETE):" & vbCrLf
    n = 0
    For Each k In dict.keys
        n = n + 1
        If n > 500 Then
            out = out & "... (" & (dict.Count - 500) & " more)" & vbCrLf
            Exit For
        End If
        out = out & k & vbCrLf
    Next k
    BuildLearnedSubjectsListStd = out
    Exit Function
StdErr:
    BuildLearnedSubjectsListStd = "ERROR: could not list subject rules: " & Err.Description
End Function

Private Function CaptureFilterDryRunStd() As String
    On Error GoTo StdErr

    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim output As String
    Dim decision As String
    Dim icon As String
    Dim processCount As Long

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items
    myItems.Sort "[ReceivedTime]", True

    output = "DRY RUN - First " & RuntimeDryRunLimit & " emails:" & vbCrLf
    processCount = 0

    For i = 1 To myItems.Count
        If processCount >= RuntimeDryRunLimit Then Exit For
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            processCount = processCount + 1
            decision = ClassifyEmail(mail)
            Select Case decision
                Case "DELETE"
                    icon = IIf(lastClassifyWasLearned, "[xLR]", IIf(lastClassifyWasLearnedSubject, "[xLS]", "[DEL]"))
                Case "MOVE_II": icon = "[II] "
                Case "LLM_REVIEW": icon = "[???]"
                Case "KEEP": icon = IIf(lastClassifyWasLearned, "[+LR]", "[OK] ")
                Case Else: icon = "[???]"
            End Select
            output = output & icon & " " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & " | " & _
                     Truncate(mail.SenderName, 20) & " | " & Truncate(mail.Subject, 40) & vbCrLf
        End If
    Next i

    output = output & vbCrLf & "Total: " & processCount & " emails previewed."
    CaptureFilterDryRunStd = output
    Exit Function
StdErr:
    CaptureFilterDryRunStd = "ERROR: FilterExistingDryRun failed: " & Err.Description
End Function

' Bridge-friendly SummarizeSelectedEmail — returns result string instead of MsgBox
Private Function SummarizeSelectedEmailStd() As String
    On Error GoTo StdErr
    Dim mail As Outlook.MailItem
    Dim prompt As String
    Dim summary As String
    Dim systemPrompt As String

    If Not RuntimeUseLLM Then
        SummarizeSelectedEmailStd = "ERROR: LLM is not enabled. Set UseLLMAPI=True in settings."
        Exit Function
    End If

    If Application.ActiveExplorer Is Nothing Then
        SummarizeSelectedEmailStd = "ERROR: No active Outlook window (Outlook may be minimised to tray)."
        Exit Function
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        SummarizeSelectedEmailStd = "ERROR: No email selected. Please select an email in Outlook first."
        Exit Function
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        SummarizeSelectedEmailStd = "ERROR: Selected item is not an email."
        Exit Function
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    systemPrompt = "You are a helpful assistant. Summarize the following email concisely in 2-3 bullet points. " & _
                   "Focus on: who sent it, what they want, and any action required."

    prompt = "Summarize this email:" & vbCrLf & _
             "From: " & mail.SenderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & mail.Subject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    summary = CallLLM(prompt, systemPrompt, RuntimeSummarizeMaxTokens)

    If Len(summary) = 0 Then
        SummarizeSelectedEmailStd = "ERROR: LLM returned no response. Check your API configuration."
        Exit Function
    End If

    SummarizeSelectedEmailStd = "Summary of: " & mail.Subject & vbCrLf & vbCrLf & summary
    Exit Function
StdErr:
    SummarizeSelectedEmailStd = "ERROR: SummarizeSelectedEmail failed: " & Err.Description
End Function

' Bridge-friendly DraftReplyToSelected — returns result string, saves draft via DraftAutoReply, no MsgBox
Private Function DraftReplyToSelectedStd() As String
    On Error GoTo StdErr
    Dim mail As Outlook.MailItem

    If Not RuntimeUseLLM Then
        DraftReplyToSelectedStd = "ERROR: LLM is not enabled. Set UseLLMAPI=True in settings."
        Exit Function
    End If

    If Application.ActiveExplorer Is Nothing Then
        DraftReplyToSelectedStd = "ERROR: No active Outlook window (Outlook may be minimised to tray)."
        Exit Function
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        DraftReplyToSelectedStd = "ERROR: No email selected. Please select an email in Outlook first."
        Exit Function
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        DraftReplyToSelectedStd = "ERROR: Selected item is not an email."
        Exit Function
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    If DraftAutoReply(mail) Then
        DraftReplyToSelectedStd = "Draft reply saved to Drafts for: " & mail.Subject
    Else
        DraftReplyToSelectedStd = "ERROR: Could not draft reply. Check LLM configuration and logs."
    End If
    Exit Function
StdErr:
    DraftReplyToSelectedStd = "ERROR: DraftReplyToSelected failed: " & Err.Description
End Function

' Bridge-friendly DraftReplyForSelected — drafts replies for all selected emails, no MsgBox
Private Function DraftReplyForSelectedStd() As String
    On Error GoTo StdErr
    Dim sel As Outlook.Selection
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim draftedCount As Long
    Dim skippedCount As Long

    If Not RuntimeUseLLM Then
        DraftReplyForSelectedStd = "ERROR: LLM is not enabled. Set UseLLMAPI=True in settings."
        Exit Function
    End If

    If Application.ActiveExplorer Is Nothing Then
        DraftReplyForSelectedStd = "ERROR: No active Outlook window (Outlook may be minimised to tray)."
        Exit Function
    End If

    Set sel = Application.ActiveExplorer.Selection

    If sel.Count = 0 Then
        DraftReplyForSelectedStd = "ERROR: No email selected. Please select one or more emails in Outlook first."
        Exit Function
    End If

    draftedCount = 0
    skippedCount = 0

    For i = 1 To sel.Count
        If Not TypeOf sel(i) Is Outlook.MailItem Then
            skippedCount = skippedCount + 1
        Else
            Set mail = sel(i)
            If DraftAutoReply(mail) Then
                draftedCount = draftedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        End If
    Next i

    DraftReplyForSelectedStd = "Draft replies complete. Drafted: " & draftedCount & ", Skipped: " & skippedCount & ". Check your Drafts folder."
    Exit Function
StdErr:
    DraftReplyForSelectedStd = "ERROR: DraftReplyForSelected failed: " & Err.Description
End Function
