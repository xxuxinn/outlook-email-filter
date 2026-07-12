'===============================================================================
' EmailDigest.bas - Daily Triage Digest + Rule Mining v3.1
'===============================================================================
' GenerateDailyDigestCore:
'   Collects the last 24 h from Inbox + Review, batch-summarizes via the LLM
'   (category / urgency / one-line summary / deadline per email), groups by
'   conversation, and renders a ranked markdown digest:
'     Needs action -> Worth a look -> FYI -> Review folder -> filtered activity.
'   Saved to %APPDATA%\OutlookEmailFilter\digests\digest_YYYY-MM-DD.md and
'   optionally sent as a self-addressed email. Optionally creates draft
'   Outlook Tasks for emails with extracted deadlines.
'
' ProposeRulesCore:
'   Mines decision_log.txt + the Review folder for repeat offenders and asks
'   the LLM to propose new sender/subject rules. Proposals are written to
'   rule_proposals.txt as PENDING — a human approves or rejects each in the
'   Web UI before anything becomes a live rule.
'
' Scheduling: Bridge.CheckScheduledJobs runs the digest daily after
' RuntimeDigestHour and rule mining weekly (no external cron needed).
'===============================================================================

Option Explicit

' Collected per-email data for one digest run
Private Type DigestItem
    senderName As String
    senderEmail As String
    subject As String
    bodyPreview As String
    received As Date
    folderName As String
    convTopic As String
    ' Filled by the LLM pass:
    category As String
    urgency As Long
    summary As String
    deadline As String   ' "yyyy-mm-dd" or ""
End Type

' Monotonic suffix so proposal ids minted in the same second stay unique
Private proposalIdCounter As Long

'-------------------------------------------------------------------------------
' PATHS
'-------------------------------------------------------------------------------

Public Function GetDigestsFolderPath() As String
    Dim dir As String
    dir = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & DIGESTS_SUBFOLDER
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then
        On Error Resume Next
        fso.CreateFolder dir
        On Error GoTo 0
    End If
    Set fso = Nothing
    GetDigestsFolderPath = dir
End Function

Public Function GetProposalsFilePath() As String
    GetProposalsFilePath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & PROPOSALS_FILE
End Function

'-------------------------------------------------------------------------------
' DAILY DIGEST
'-------------------------------------------------------------------------------

' Interactive wrapper — run from QAT or the VBA editor
Public Sub GenerateDailyDigest()
    Dim result As String
    result = GenerateDailyDigestCore()
    MsgBox result, IIf(Left(result, 6) = "ERROR:", vbExclamation, vbInformation), "Daily Digest"
End Sub

' Headless digest generation. Returns a one-paragraph summary or "ERROR: ...".
Public Function GenerateDailyDigestCore() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.GenerateDailyDigestCore"

    Dim items() As DigestItem
    Dim itemCount As Long
    Dim inboxCount As Long
    Dim reviewCount As Long

    itemCount = CollectRecentEmails(items, inboxCount, reviewCount)

    If itemCount = 0 Then
        GenerateDailyDigestCore = "No emails in the last 24 hours — no digest generated."
        GoTo PROC_EXIT
    End If

    ' LLM enrichment pass (best-effort; digest still renders without it)
    Dim llmUsed As Boolean
    llmUsed = False
    If RuntimeUseLLM Then
        llmUsed = EnrichItemsViaLLM(items, itemCount)
    End If

    ' Render + persist
    Dim md As String
    md = RenderDigestMarkdown(items, itemCount, inboxCount, reviewCount, llmUsed)

    Dim digestPath As String
    digestPath = GetDigestsFolderPath() & "\digest_" & Format(Date, "yyyy-mm-dd") & ".md"
    WriteTextFileUTF8 digestPath, md

    ' Draft Outlook Tasks for extracted deadlines
    Dim tasksCreated As Long
    tasksCreated = 0
    If RuntimeEnableTaskExtraction And llmUsed Then
        tasksCreated = CreateDeadlineTasks(items, itemCount)
    End If

    ' Self-addressed digest email
    Dim emailSent As Boolean
    emailSent = False
    If RuntimeDigestSendEmail Then
        emailSent = SendDigestEmail(md)
    End If

    GenerateDailyDigestCore = "Digest generated for " & itemCount & " email(s) " & _
        "(Inbox: " & inboxCount & ", Review: " & reviewCount & ")." & _
        IIf(llmUsed, "", " LLM disabled/unavailable - counts only, no summaries.") & _
        IIf(tasksCreated > 0, " Created " & tasksCreated & " deadline task(s).", "") & _
        IIf(emailSent, " Sent to your inbox.", "") & _
        " Saved: " & digestPath

    LogMessage "INFO", "GenerateDailyDigestCore: " & GenerateDailyDigestCore

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "GenerateDailyDigestCore", Err.Number, Err.Description
    GenerateDailyDigestCore = "ERROR: Digest generation failed: " & Err.Description
    Resume PROC_EXIT
End Function

' Collect MailItems from the last 24 h in Inbox + Review (newest first, capped
' at RuntimeDigestMaxEmails). Returns the item count.
Private Function CollectRecentEmails(ByRef items() As DigestItem, _
                                     ByRef inboxCount As Long, ByRef reviewCount As Long) As Long
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.CollectRecentEmails"

    Dim maxItems As Long
    maxItems = RuntimeDigestMaxEmails
    If maxItems < 1 Then maxItems = 50

    ReDim items(1 To maxItems)
    Dim count As Long
    count = 0
    inboxCount = 0
    reviewCount = 0

    Dim cutoff As Date
    cutoff = Now - 1  ' 24 hours

    Dim ns As Outlook.NameSpace
    Set ns = Application.GetNamespace("MAPI")

    ' Inbox
    count = CollectFromFolder(ns.GetDefaultFolder(olFolderInbox), "Inbox", cutoff, items, count, maxItems)
    inboxCount = count

    ' Review folder (may not exist yet)
    Dim reviewFolder As Outlook.Folder
    On Error Resume Next
    Set reviewFolder = ns.GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderReview)
    On Error GoTo PROC_ERR
    If Not reviewFolder Is Nothing Then
        count = CollectFromFolder(reviewFolder, RuntimeFolderReview, cutoff, items, count, maxItems)
    End If
    reviewCount = count - inboxCount

    CollectRecentEmails = count

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "CollectRecentEmails", Err.Number, Err.Description
    CollectRecentEmails = count
    Resume PROC_EXIT
End Function

Private Function CollectFromFolder(ByVal folder As Outlook.Folder, ByVal folderLabel As String, _
                                   ByVal cutoff As Date, ByRef items() As DigestItem, _
                                   ByVal startCount As Long, ByVal maxItems As Long) As Long
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.CollectFromFolder"

    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim count As Long
    count = startCount

    Set myItems = folder.Items
    myItems.Sort "[ReceivedTime]", True  ' newest first

    For i = 1 To myItems.Count
        If count >= maxItems Then Exit For
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            If mail.ReceivedTime < cutoff Then Exit For  ' sorted desc: done

            count = count + 1
            items(count).senderName = mail.senderName
            items(count).senderEmail = GetSenderEmail(mail)
            items(count).subject = SanitizeSubject(mail.subject)
            items(count).bodyPreview = Truncate(mail.Body, 300)
            items(count).received = mail.ReceivedTime
            items(count).folderName = folderLabel
            On Error Resume Next
            items(count).convTopic = mail.ConversationTopic
            On Error GoTo PROC_ERR
            items(count).category = ""
            items(count).urgency = 0
            items(count).summary = ""
            items(count).deadline = ""
        End If
    Next i

    CollectFromFolder = count

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "CollectFromFolder", Err.Number, Err.Description, folderLabel
    CollectFromFolder = count
    Resume PROC_EXIT
End Function

' Enrich items with category/urgency/summary/deadline via batched LLM calls.
' Returns True if at least one batch succeeded.
Private Function EnrichItemsViaLLM(ByRef items() As DigestItem, ByVal itemCount As Long) As Boolean
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.EnrichItemsViaLLM"

    EnrichItemsViaLLM = False

    Dim batchSize As Long
    batchSize = RuntimeLLMBatchSize
    If batchSize < 1 Then batchSize = 10

    Dim batchStart As Long
    Dim batchEnd As Long
    Dim response As String
    Dim anySuccess As Boolean
    anySuccess = False

    Dim systemPrompt As String
    systemPrompt = "You triage emails for a busy professor. For EACH numbered email respond with EXACTLY one JSON " & _
                   "object on its own line, in the same order, no other text: " & _
                   "{""id"":<number>,""category"":""student|colleague|admin|newsletter|external|other""," & _
                   """urgency"":1-5 (5 = must respond today),""summary"":""max 20 words, what they want""," & _
                   """deadline"":""yyyy-mm-dd or empty string if none mentioned""}"

    batchStart = 1
    Do While batchStart <= itemCount
        batchEnd = batchStart + batchSize - 1
        If batchEnd > itemCount Then batchEnd = itemCount

        response = CallLLM(BuildDigestBatchPrompt(items, batchStart, batchEnd), systemPrompt, _
                           200 * (batchEnd - batchStart + 1))

        If Len(response) > 0 Then
            If ParseDigestBatchResponse(response, items, batchStart, batchEnd) Then anySuccess = True
        Else
            LogMessage "WARN", "EnrichItemsViaLLM: empty LLM response for batch " & batchStart & "-" & batchEnd
        End If

        batchStart = batchEnd + 1
    Loop

    EnrichItemsViaLLM = anySuccess

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "EnrichItemsViaLLM", Err.Number, Err.Description
    Resume PROC_EXIT
End Function

Private Function BuildDigestBatchPrompt(ByRef items() As DigestItem, _
                                        ByVal batchStart As Long, ByVal batchEnd As Long) As String
    Dim prompt As String
    Dim i As Long

    prompt = "Triage these " & (batchEnd - batchStart + 1) & " emails:" & vbCrLf & vbCrLf
    For i = batchStart To batchEnd
        prompt = prompt & "--- Email " & i & " ---" & vbCrLf & _
                 "From: " & items(i).senderName & " <" & items(i).senderEmail & ">" & vbCrLf & _
                 "Subject: " & items(i).subject & vbCrLf & _
                 "Body: " & items(i).bodyPreview & vbCrLf & vbCrLf
    Next i

    prompt = prompt & "Respond with one JSON object per line, ids " & batchStart & " to " & batchEnd & "."
    BuildDigestBatchPrompt = prompt
End Function

' Parse one-JSON-object-per-line responses; tolerant of blank/extra lines.
Private Function ParseDigestBatchResponse(ByVal response As String, ByRef items() As DigestItem, _
                                          ByVal batchStart As Long, ByVal batchEnd As Long) As Boolean
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.ParseDigestBatchResponse"

    ParseDigestBatchResponse = False

    Dim lines As Variant
    Dim line As String
    Dim i As Long
    Dim id As Long
    Dim parsed As Long
    parsed = 0

    lines = SplitLines(response)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 2 And InStr(1, line, "{") > 0 Then
            id = CLng(ExtractJSONNumberValue(line, "id", 0))
            If id >= batchStart And id <= batchEnd Then
                items(id).category = ExtractJSONStringValue(line, "category")
                items(id).urgency = CLng(ExtractJSONNumberValue(line, "urgency", 2))
                items(id).summary = ExtractJSONStringValue(line, "summary")
                items(id).deadline = ExtractJSONStringValue(line, "deadline")
                parsed = parsed + 1
            End If
        End If
    Next i

    ParseDigestBatchResponse = (parsed > 0)
    If parsed < (batchEnd - batchStart + 1) Then
        LogMessage "WARN", "ParseDigestBatchResponse: parsed " & parsed & " of " & _
                   (batchEnd - batchStart + 1) & " items in batch"
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "ParseDigestBatchResponse", Err.Number, Err.Description
    Resume PROC_EXIT
End Function

' Render the digest markdown: ranked sections, conversation grouping, and a
' filtered-activity tail from decision_log.txt.
Private Function RenderDigestMarkdown(ByRef items() As DigestItem, ByVal itemCount As Long, _
                                      ByVal inboxCount As Long, ByVal reviewCount As Long, _
                                      ByVal llmUsed As Boolean) As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.RenderDigestMarkdown"

    Dim md As String
    md = "# Daily Email Digest — " & Format(Date, "yyyy-mm-dd") & vbCrLf & vbCrLf
    md = md & "Generated " & Format(Now, "yyyy-mm-dd hh:nn") & ". " & _
         itemCount & " email(s) in the last 24 h (Inbox: " & inboxCount & _
         ", " & RuntimeFolderReview & ": " & reviewCount & ")." & vbCrLf & vbCrLf

    If Not llmUsed Then
        md = md & "> LLM summaries unavailable — enable UseLLMAPI in settings for ranked triage." & vbCrLf & vbCrLf
        md = md & "## All emails" & vbCrLf & vbCrLf
        md = md & RenderItemLines(items, itemCount, -1, 99, False)
    Else
        Dim sectionText As String

        sectionText = RenderItemLines(items, itemCount, 4, 5, True)
        If Len(sectionText) > 0 Then
            md = md & "## Needs action" & vbCrLf & vbCrLf & sectionText & vbCrLf
        End If

        sectionText = RenderItemLines(items, itemCount, 3, 3, True)
        If Len(sectionText) > 0 Then
            md = md & "## Worth a look" & vbCrLf & vbCrLf & sectionText & vbCrLf
        End If

        sectionText = RenderItemLines(items, itemCount, -1, 2, True)
        If Len(sectionText) > 0 Then
            md = md & "## FYI" & vbCrLf & vbCrLf & sectionText & vbCrLf
        End If
    End If

    ' Review folder call-out
    If reviewCount > 0 Then
        md = md & "## Waiting for your verdict in " & RuntimeFolderReview & vbCrLf & vbCrLf
        md = md & "- " & reviewCount & " email(s) the filter was unsure about. " & _
             "Drag to LearnKeep/LearnDelete to teach it." & vbCrLf & vbCrLf
    End If

    ' Filtered activity from the decision log
    md = md & RenderFilteredActivity()

    RenderDigestMarkdown = md

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "RenderDigestMarkdown", Err.Number, Err.Description
    RenderDigestMarkdown = md
    Resume PROC_EXIT
End Function

' Render bullet lines for items whose urgency is within [minUrg, maxUrg].
' Conversations are grouped: only the newest email of each ConversationTopic is
' listed, with a "(n messages)" suffix. Only Inbox items are listed here.
Private Function RenderItemLines(ByRef items() As DigestItem, ByVal itemCount As Long, _
                                 ByVal minUrg As Long, ByVal maxUrg As Long, _
                                 ByVal withSummary As Boolean) As String
    Dim out As String
    Dim i As Long, j As Long
    Dim threadCount As Long
    Dim seenTopics As Object
    Set seenTopics = CreateObject("Scripting.Dictionary")
    seenTopics.CompareMode = 1

    For i = 1 To itemCount
        If items(i).folderName = "Inbox" And items(i).urgency >= minUrg And items(i).urgency <= maxUrg Then
            Dim topicKey As String
            topicKey = Trim(items(i).convTopic)
            If Len(topicKey) = 0 Then topicKey = items(i).subject & "|" & items(i).senderEmail

            If Not seenTopics.Exists(topicKey) Then
                seenTopics(topicKey) = True

                ' Count other messages in the same conversation (any section)
                threadCount = 0
                For j = 1 To itemCount
                    If StrComp(Trim(items(j).convTopic), topicKey, vbTextCompare) = 0 And Len(Trim(items(j).convTopic)) > 0 Then
                        threadCount = threadCount + 1
                    End If
                Next j

                out = out & "- **" & items(i).senderName & "** — " & items(i).subject
                If threadCount > 1 Then out = out & " *(" & threadCount & " messages in thread)*"
                out = out & vbCrLf
                If withSummary And Len(items(i).summary) > 0 Then
                    out = out & "  - " & items(i).summary
                    If Len(items(i).deadline) > 0 Then out = out & " **[deadline: " & items(i).deadline & "]**"
                    If Len(items(i).category) > 0 Then out = out & " _(" & items(i).category & ")_"
                    out = out & vbCrLf
                End If
            End If
        End If
    Next i

    RenderItemLines = out
End Function

' Summarize the last 24 h of decision_log.txt: what got auto-filtered and why.
Private Function RenderFilteredActivity() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.RenderFilteredActivity"

    RenderFilteredActivity = ""

    Dim content As String
    content = ReadTextFileSmart(GetDecisionLogPath())
    If Len(content) = 0 Then GoTo PROC_EXIT

    Dim cutoff As Date
    cutoff = Now - 1

    Dim lines As Variant
    Dim parts() As String
    Dim i As Long
    Dim delCount As Long, keepCount As Long, revCount As Long, protCount As Long
    Dim delSenders As Object
    Set delSenders = CreateObject("Scripting.Dictionary")
    delSenders.CompareMode = 1

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            ' timestamp|sender|subject|source|action|confidence
            If UBound(parts) >= 4 Then
                If IsDate(parts(0)) Then
                    If CDate(parts(0)) >= cutoff Then
                        Select Case UCase(Trim(parts(4)))
                            Case "DELETE"
                                delCount = delCount + 1
                                If delSenders.Exists(parts(1)) Then
                                    delSenders(parts(1)) = delSenders(parts(1)) + 1
                                Else
                                    delSenders(parts(1)) = 1
                                End If
                            Case "KEEP": keepCount = keepCount + 1
                            Case "REVIEW": revCount = revCount + 1
                            Case "MOVE_II": protCount = protCount + 1
                        End Select
                    End If
                End If
            End If
        End If
    Next i

    If delCount + keepCount + revCount + protCount = 0 Then GoTo PROC_EXIT

    Dim md As String
    md = "## Filter activity (last 24 h)" & vbCrLf & vbCrLf
    md = md & "- Kept: " & keepCount & " | Deleted: " & delCount & " | To Review: " & revCount & _
         " | Protected: " & protCount & vbCrLf

    ' Top deleted senders (up to 3)
    If delSenders.Count > 0 Then
        Dim k As Variant, topKey As String, topVal As Long, pass As Long
        Dim listed As Object
        Set listed = CreateObject("Scripting.Dictionary")
        listed.CompareMode = 1
        For pass = 1 To 3
            topVal = 0
            topKey = ""
            For Each k In delSenders.keys
                If delSenders(k) > topVal And Not listed.Exists(CStr(k)) Then
                    topVal = delSenders(k)
                    topKey = CStr(k)
                End If
            Next k
            If Len(topKey) = 0 Then Exit For
            listed(topKey) = True
            md = md & "- Top deleted: " & topKey & " (" & topVal & ")" & vbCrLf
        Next pass
    End If

    RenderFilteredActivity = md & vbCrLf

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "RenderFilteredActivity", Err.Number, Err.Description
    Resume PROC_EXIT
End Function

' Create draft Outlook Tasks (saved, never sent) for items with a parseable deadline.
Private Function CreateDeadlineTasks(ByRef items() As DigestItem, ByVal itemCount As Long) As Long
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.CreateDeadlineTasks"

    Dim created As Long
    Dim i As Long
    Dim task As Object  ' Outlook.TaskItem
    Dim dueDate As Date

    created = 0
    For i = 1 To itemCount
        If Len(items(i).deadline) > 0 And items(i).urgency >= 3 Then
            If TryParseISODate(items(i).deadline, dueDate) Then
                Set task = Application.CreateItem(olTaskItem)
                task.subject = "Reply: " & Truncate(items(i).subject, 80) & " (" & items(i).senderName & ")"
                task.DueDate = dueDate
                task.Body = "From: " & items(i).senderName & " <" & items(i).senderEmail & ">" & vbCrLf & _
                            "Subject: " & items(i).subject & vbCrLf & _
                            "Summary: " & items(i).summary & vbCrLf & vbCrLf & _
                            "(Created by the Email Agent daily digest — deadline extracted by LLM, verify before relying on it.)"
                task.Save
                created = created + 1
                Set task = Nothing
            End If
        End If
    Next i

    CreateDeadlineTasks = created
    If created > 0 Then LogMessage "INFO", "CreateDeadlineTasks: created " & created & " task draft(s)"

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "CreateDeadlineTasks", Err.Number, Err.Description
    CreateDeadlineTasks = created
    Resume PROC_EXIT
End Function

' Locale-safe "yyyy-mm-dd" parser (CDate is locale-dependent)
Private Function TryParseISODate(ByVal s As String, ByRef result As Date) As Boolean
    TryParseISODate = False
    Dim parts() As String
    parts = Split(Trim(s), "-")
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function

    On Error GoTo Fail
    result = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
    ' Ignore absurd deadlines (hallucinated past dates or >1 year out)
    If result < Date - 1 Or result > Date + 366 Then Exit Function
    TryParseISODate = True
Fail:
End Function

' Send the digest to the user's own mailbox. Returns True on success.
Private Function SendDigestEmail(ByVal digestMarkdown As String) As Boolean
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.SendDigestEmail"

    SendDigestEmail = False

    Dim item As Outlook.MailItem
    Set item = Application.CreateItem(olMailItem)
    item.subject = "Daily Email Digest — " & Format(Date, "yyyy-mm-dd")
    item.Body = digestMarkdown
    item.Recipients.Add Application.Session.CurrentUser.Name
    If item.Recipients.ResolveAll Then
        item.Send
        SendDigestEmail = True
    Else
        ' Could not resolve own address — leave as draft instead of failing
        item.Save
        LogMessage "WARN", "SendDigestEmail: could not resolve own address; digest saved to Drafts"
    End If
    Set item = Nothing

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "SendDigestEmail", Err.Number, Err.Description
    SendDigestEmail = False
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' RULE MINING (human-in-the-loop: proposals only, approval in the Web UI)
'-------------------------------------------------------------------------------

' Interactive wrapper
Public Sub ProposeRules()
    Dim result As String
    result = ProposeRulesCore()
    MsgBox result, IIf(Left(result, 6) = "ERROR:", vbExclamation, vbInformation), "Rule Mining"
End Sub

' Mine the decision log + Review folder and ask the LLM to propose new rules.
' Proposals land in rule_proposals.txt with status PENDING.
Public Function ProposeRulesCore() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.ProposeRulesCore"

    If Not RuntimeUseLLM Then
        ProposeRulesCore = "ERROR: Rule mining requires the LLM. Set UseLLMAPI=True in settings."
        GoTo PROC_EXIT
    End If

    ' --- Evidence: repeat senders in Review + review/delete history ---
    Dim evidence As String
    evidence = BuildMiningEvidence()

    If Len(evidence) = 0 Then
        ProposeRulesCore = "No mining evidence yet — the Review folder is empty and the decision log has no recent activity."
        GoTo PROC_EXIT
    End If

    Dim systemPrompt As String
    systemPrompt = "You mine email-filtering rules for a professor's mailbox. Given evidence about repeat senders " & _
                   "and subjects, propose up to 8 NEW rules. Be conservative: only propose a rule when the pattern " & _
                   "is clearly junk (DELETE) or clearly important (KEEP). Subject values must be distinctive multi-word " & _
                   "phrases (never a single common word — rules match by substring). Respond with EXACTLY one JSON " & _
                   "object per line, no other text: " & _
                   "{""type"":""SENDER"" or ""SUBJECT"",""value"":""email address or subject phrase""," & _
                   """action"":""KEEP"" or ""DELETE"",""reason"":""max 15 words""}" & _
                   " SUBJECT rules may only use action DELETE."

    Dim response As String
    response = CallLLM(evidence, systemPrompt, 1000)

    If Len(response) = 0 Then
        ProposeRulesCore = "ERROR: LLM returned no response for rule mining."
        GoTo PROC_EXIT
    End If

    ' --- Parse, validate, and record proposals ---
    Dim added As Long, skipped As Long
    ParseAndRecordProposals response, added, skipped

    ProposeRulesCore = "Rule mining complete: " & added & " new proposal(s)" & _
                       IIf(skipped > 0, ", " & skipped & " skipped (invalid or already covered)", "") & _
                       ". Review and approve them in the Web UI Proposals tab."
    LogMessage "INFO", "ProposeRulesCore: " & ProposeRulesCore

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "ProposeRulesCore", Err.Number, Err.Description
    ProposeRulesCore = "ERROR: Rule mining failed: " & Err.Description
    Resume PROC_EXIT
End Function

' Build the evidence prompt: Review folder samples + repeat senders from the log.
Private Function BuildMiningEvidence() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.BuildMiningEvidence"

    Dim evidence As String

    ' Review folder samples (up to 30)
    Dim ns As Outlook.NameSpace
    Dim reviewFolder As Outlook.Folder
    Set ns = Application.GetNamespace("MAPI")
    On Error Resume Next
    Set reviewFolder = ns.GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderReview)
    On Error GoTo PROC_ERR

    If Not reviewFolder Is Nothing Then
        Dim myItems As Outlook.Items
        Dim mail As Outlook.MailItem
        Dim i As Long, sampled As Long
        Set myItems = reviewFolder.Items
        myItems.Sort "[ReceivedTime]", True

        Dim reviewBlock As String
        sampled = 0
        For i = 1 To myItems.Count
            If sampled >= 30 Then Exit For
            If TypeOf myItems(i) Is Outlook.MailItem Then
                Set mail = myItems(i)
                sampled = sampled + 1
                reviewBlock = reviewBlock & "- " & mail.senderName & " <" & GetSenderEmail(mail) & _
                              "> | " & SanitizeSubject(Truncate(mail.subject, 70)) & vbCrLf
            End If
        Next i
        If sampled > 0 Then
            evidence = "Emails currently sitting in the Review folder (the filter could not decide):" & vbCrLf & _
                       reviewBlock & vbCrLf
        End If
    End If

    ' Repeat senders from the decision log (last 500 lines): senders reviewed
    ' or deleted 3+ times are rule candidates
    Dim content As String
    content = ReadTextFileSmart(GetDecisionLogPath())
    If Len(content) > 0 Then
        Dim lines As Variant
        Dim parts() As String
        Dim senderCounts As Object
        Set senderCounts = CreateObject("Scripting.Dictionary")
        senderCounts.CompareMode = 1

        lines = SplitLines(content)
        Dim startIdx As Long
        startIdx = LBound(lines)
        If UBound(lines) - startIdx + 1 > 500 Then startIdx = UBound(lines) - 499

        Dim j As Long
        For j = startIdx To UBound(lines)
            If Len(Trim(lines(j))) > 0 Then
                parts = Split(lines(j), "|")
                If UBound(parts) >= 4 Then
                    Dim act As String
                    act = UCase(Trim(parts(4)))
                    If act = "REVIEW" Or act = "DELETE" Then
                        Dim sKey As String
                        sKey = Trim(parts(1)) & " -> " & act
                        If senderCounts.Exists(sKey) Then
                            senderCounts(sKey) = senderCounts(sKey) + 1
                        Else
                            senderCounts(sKey) = 1
                        End If
                    End If
                End If
            End If
        Next j

        Dim histBlock As String
        Dim k As Variant
        For Each k In senderCounts.keys
            If senderCounts(k) >= 3 Then
                ' Skip senders that already have a learned rule
                Dim senderOnly As String
                senderOnly = Trim(Split(CStr(k), " -> ")(0))
                If Len(LookupLearnedSender(senderOnly)) = 0 Then
                    histBlock = histBlock & "- " & k & " (" & senderCounts(k) & " times)" & vbCrLf
                End If
            End If
        Next k
        If Len(histBlock) > 0 Then
            evidence = evidence & "Repeat senders from recent history (no learned rule yet):" & vbCrLf & histBlock
        End If
    End If

    BuildMiningEvidence = evidence

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailDigest", "BuildMiningEvidence", Err.Number, Err.Description
    BuildMiningEvidence = evidence
    Resume PROC_EXIT
End Function

' Parse one-object-per-line proposals, validate hard, append PENDING rows.
Private Sub ParseAndRecordProposals(ByVal response As String, ByRef added As Long, ByRef skipped As Long)
    On Error GoTo PROC_ERR
    PushCallStack "EmailDigest.ParseAndRecordProposals"

    added = 0
    skipped = 0

    ' Existing proposal values (avoid duplicates across runs)
    Dim existingValues As Object
    Set existingValues = CreateObject("Scripting.Dictionary")
    existingValues.CompareMode = 1

    Dim existing As String
    existing = ReadTextFileSmart(GetProposalsFilePath())
    If Len(existing) > 0 Then
        Dim exLines As Variant
        Dim exParts() As String
        Dim e As Long
        exLines = SplitLines(existing)
        For e = LBound(exLines) To UBound(exLines)
            exParts = Split(exLines(e), "|")
            ' id|type|value|action|reason|status|timestamp
            If UBound(exParts) >= 5 Then existingValues(Trim(exParts(2))) = True
        Next e
    End If

    Dim lines As Variant
    Dim line As String
    Dim i As Long

    lines = SplitLines(response)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 2 And InStr(1, line, "{") > 0 Then
            Dim ruleType As String, ruleValue As String, ruleAction As String, ruleReason As String
            ruleType = UCase(Trim(ExtractJSONStringValue(line, "type")))
            ruleValue = Trim(ExtractJSONStringValue(line, "value"))
            ruleAction = UCase(Trim(ExtractJSONStringValue(line, "action")))
            ruleReason = Trim(ExtractJSONStringValue(line, "reason"))

            If ValidateProposal(ruleType, ruleValue, ruleAction) And Not existingValues.Exists(ruleValue) Then
                existingValues(ruleValue) = True
                AppendLineUTF8 GetProposalsFilePath(), _
                    GenerateProposalId() & "|" & ruleType & "|" & _
                    SanitizeSubject(ruleValue) & "|" & ruleAction & "|" & _
                    SanitizeSubject(Truncate(ruleReason, 100)) & "|PENDING|" & _
                    Format(Now, "yyyy-mm-dd hh:nn:ss")
                added = added + 1
            ElseIf Len(ruleType) > 0 Then
                skipped = skipped + 1
            End If
        End If
    Next i

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailDigest", "ParseAndRecordProposals", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Hard validation: same guards a human would apply before trusting an LLM rule.
Private Function ValidateProposal(ByVal ruleType As String, ByVal ruleValue As String, _
                                  ByVal ruleAction As String) As Boolean
    ValidateProposal = False

    If ruleAction <> "KEEP" And ruleAction <> "DELETE" Then Exit Function
    If InStr(1, ruleValue, "|") > 0 Then Exit Function

    Select Case ruleType
        Case "SENDER"
            ' Must look like an email address and not already be covered
            If InStr(1, ruleValue, "@") = 0 Then Exit Function
            If Len(ruleValue) < 6 Then Exit Function
            If Len(LookupLearnedSender(ruleValue)) > 0 Then Exit Function
        Case "SUBJECT"
            ' DELETE only; distinctive multi-word phrase (substring matching!)
            If ruleAction <> "DELETE" Then Exit Function
            If Len(ruleValue) < 12 Then Exit Function
            If InStr(1, Trim(ruleValue), " ") = 0 Then Exit Function
            If Len(LookupLearnedSubject(ruleValue)) > 0 Then Exit Function
        Case Else
            Exit Function
    End Select

    ValidateProposal = True
End Function

' Mint an 8-char lowercase hex id (webui validates ^[0-9a-f]{8}$).
' Low 7 hex digits of seconds-since-2020 + a rolling counter nibble. The
' seconds value outgrows 7 hex digits in mid-2028; after that the top digit
' is dropped, so ids can only collide with ones minted a full 8.5-year epoch
' earlier — acceptable for a human-reviewed proposals queue.
Private Function GenerateProposalId() As String
    Dim seconds As Long
    seconds = DateDiff("s", DateSerial(2020, 1, 1), Now)
    proposalIdCounter = (proposalIdCounter + 1) Mod 16
    GenerateProposalId = LCase(Right("0000000" & Hex(seconds), 7) & Hex(proposalIdCounter))
End Function
