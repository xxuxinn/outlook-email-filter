'===============================================================================
' EmailFilter.bas - Main Classification Logic v3.1
'===============================================================================
' This module contains the core email classification functions,
' LLM integration (via multi-provider CallLLM), and LLM email tools.
'
' v3.1: LLM classification is structured (JSON with action/category/urgency/
' confidence), enriched with sender history from AgentMemory, gated by a
' confidence threshold, and every executed decision is recorded to
' decision_log.txt.
'
' All pattern/setting references use Runtime* variables from Config.bas,
' loaded from settings.ini by LoadAllSettings at startup.
'===============================================================================

Option Explicit

' Flag set by ClassifyEmail when a learned rule was applied (used by dry-run/reporting)
Public lastClassifyWasLearned As Boolean

' Flag set by ClassifyEmail when a learned subject rule was applied
Public lastClassifyWasLearnedSubject As Boolean

' Which rule produced the last ClassifyEmail decision (for decision_log.txt)
Public lastClassifySource As String

' Structured outputs of the last LLM classification (see ClassifyViaLLMEx)
Public lastLLMConfidence As Double
Public lastLLMUrgency As Long
Public lastLLMCategory As String
Public lastLLMReason As String

'-------------------------------------------------------------------------------
' MAIN CLASSIFICATION FUNCTION
'-------------------------------------------------------------------------------

' Classify an email and return the action to take
' Returns: "MOVE_II", "DELETE", "KEEP", or "LLM_REVIEW"
Public Function ClassifyEmail(ByVal mail As Outlook.MailItem) As String
    Dim senderEmail As String
    Dim senderName As String
    Dim subject As String
    Dim bodyStart As String
    Dim domain As String

    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.ClassifyEmail"

    ' Extract email properties
    senderEmail = GetSenderEmail(mail)
    senderName = mail.senderName
    subject = mail.subject
    bodyStart = Left(mail.Body, 500)  ' First 500 chars for greeting detection
    domain = GetDomain(senderEmail)

    LogMessage "DEBUG", "Classifying: " & Truncate(subject, 50)

    ' Reset learned-rule flags and decision source
    lastClassifyWasLearned = False
    lastClassifyWasLearnedSubject = False
    lastClassifySource = ""

    ' =========================================================================
    ' RULE 0: LEARNED SENDER RULES (highest priority, from self-improving)
    ' =========================================================================
    If RuntimeEnableSelfImproving Then
        Dim learnedAction As String
        learnedAction = LookupLearnedSender(senderEmail)

        If learnedAction = "KEEP" Then
            ClassifyEmail = "KEEP"
            lastClassifyWasLearned = True
            lastClassifySource = "LEARNED_SENDER"
            LogMessage "DEBUG", "  -> KEEP (learned rule for: " & senderEmail & ")"
            GoTo PROC_EXIT
        ElseIf learnedAction = "DELETE" Then
            ClassifyEmail = "DELETE"
            lastClassifyWasLearned = True
            lastClassifySource = "LEARNED_SENDER"
            LogMessage "DEBUG", "  -> DELETE (learned rule for: " & senderEmail & ")"
            GoTo PROC_EXIT
        End If
        ' Empty string = no learned rule, fall through
    End If

    ' =========================================================================
    ' RULE 0.5: LEARNED SUBJECT RULES (from LearnSubjectDelete folder)
    ' =========================================================================
    If RuntimeEnableSelfImproving Then
        Dim learnedSubjectAction As String
        learnedSubjectAction = LookupLearnedSubject(subject)

        If learnedSubjectAction = "DELETE" Then
            ClassifyEmail = "DELETE"
            lastClassifyWasLearnedSubject = True
            lastClassifySource = "LEARNED_SUBJECT"
            LogMessage "DEBUG", "  -> DELETE (learned subject rule match)"
            GoTo PROC_EXIT
        End If
    End If

    ' =========================================================================
    ' RULE 1: PROTECTED DOMAINS - Move to Protected folder
    ' =========================================================================
    If IsProtectedDomain(domain) Then
        ClassifyEmail = "MOVE_II"
        lastClassifySource = "RULE1_PROTECTED"
        LogMessage "DEBUG", "  -> MOVE_II (protected domain: " & domain & ")"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 2: PERSONALLY ADDRESSED - Keep
    ' =========================================================================
    If IsPersonallyAddressed(subject, bodyStart) Then
        ClassifyEmail = "KEEP"
        lastClassifySource = "RULE2_PERSONAL"
        LogMessage "DEBUG", "  -> KEEP (personally addressed)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 3: ORGANIZATIONAL TAGS - Keep
    ' =========================================================================
    If ContainsAny(subject, RuntimePolyUTags) Then
        ClassifyEmail = "KEEP"
        lastClassifySource = "RULE3_ORGTAG"
        LogMessage "DEBUG", "  -> KEEP (org tag found)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 4: VIP SUBJECT KEYWORDS - Keep
    ' =========================================================================
    If ContainsAny(subject, RuntimeVIPKeywords) Then
        ClassifyEmail = "KEEP"
        lastClassifySource = "RULE4_VIP"
        LogMessage "DEBUG", "  -> KEEP (VIP keyword)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 5: REPLY CHAINS - Keep (RE: subjects)
    ' =========================================================================
    If IsReplyEmail(subject) Then
        ClassifyEmail = "KEEP"
        lastClassifySource = "RULE5_REPLY"
        LogMessage "DEBUG", "  -> KEEP (reply chain)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 6: FORWARD CHAINS - Keep (FW: subjects)
    ' =========================================================================
    If IsForwardEmail(subject) Then
        ClassifyEmail = "KEEP"
        lastClassifySource = "RULE6_FORWARD"
        LogMessage "DEBUG", "  -> KEEP (forwarded)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 7: KNOWN SPAM SENDERS - Delete
    ' =========================================================================
    If ContainsAny(senderName, RuntimeDeleteKnownSenders) Then
        ClassifyEmail = "DELETE"
        lastClassifySource = "RULE7_KNOWN_SENDER"
        LogMessage "DEBUG", "  -> DELETE (known sender: " & senderName & ")"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 8: SENDER EMAIL PATTERNS - Delete
    ' =========================================================================
    If ContainsAny(senderEmail, RuntimeDeleteSenderPatterns) Then
        ClassifyEmail = "DELETE"
        lastClassifySource = "RULE8_SENDER_PATTERN"
        LogMessage "DEBUG", "  -> DELETE (sender pattern)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 9: SUBJECT PATTERNS - Delete
    ' =========================================================================
    If ContainsAny(subject, RuntimeDeleteSubjectPatterns) Then
        ClassifyEmail = "DELETE"
        lastClassifySource = "RULE9_SUBJECT_PATTERN"
        LogMessage "DEBUG", "  -> DELETE (subject pattern)"
        GoTo PROC_EXIT
    End If

    ' =========================================================================
    ' RULE 10: AMBIGUOUS - Send to LLM or Review folder
    ' =========================================================================
    ClassifyEmail = "LLM_REVIEW"
    lastClassifySource = "RULE10_AMBIGUOUS"
    LogMessage "DEBUG", "  -> LLM_REVIEW (no rule matched)"

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailFilter", "ClassifyEmail", Err.Number, Err.Description
    ClassifyEmail = "KEEP"  ' Safe default: keep on error
    lastClassifySource = "ERROR_DEFAULT"
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' HELPER CLASSIFICATION FUNCTIONS
'-------------------------------------------------------------------------------

' Check if domain is in the protected list
Private Function IsProtectedDomain(ByVal domain As String) As Boolean
    IsProtectedDomain = ContainsAny(domain, RuntimeProtectedDomains)
End Function

' Check if email is personally addressed to the user
Private Function IsPersonallyAddressed(ByVal subject As String, ByVal bodyStart As String) As Boolean
    ' Check if subject contains name
    If ContainsAny(subject, RuntimeNamePatterns) Then
        IsPersonallyAddressed = True
        Exit Function
    End If

    ' Check if body starts with personal greeting
    If StartsWithAny(bodyStart, RuntimeGreetingPatterns) Then
        IsPersonallyAddressed = True
        Exit Function
    End If

    ' Check if body contains name in first few lines
    If ContainsAny(Left(bodyStart, 200), RuntimeNamePatterns) Then
        IsPersonallyAddressed = True
        Exit Function
    End If

    IsPersonallyAddressed = False
End Function

' Check if this is a reply email
Private Function IsReplyEmail(ByVal subject As String) As Boolean
    Dim trimmedSubject As String
    trimmedSubject = UCase(Trim(subject))

    IsReplyEmail = (Left(trimmedSubject, 3) = "RE:" Or _
                    Left(trimmedSubject, 4) = "RE: " Or _
                    Left(trimmedSubject, 3) = "AW:" Or _
                    Left(trimmedSubject, 4) = "AW: ")  ' German reply prefix
End Function

' Check if this is a forwarded email
Private Function IsForwardEmail(ByVal subject As String) As Boolean
    Dim trimmedSubject As String
    trimmedSubject = UCase(Trim(subject))

    IsForwardEmail = (Left(trimmedSubject, 3) = "FW:" Or _
                      Left(trimmedSubject, 4) = "FW: " Or _
                      Left(trimmedSubject, 4) = "FWD:" Or _
                      Left(trimmedSubject, 5) = "FWD: " Or _
                      Left(trimmedSubject, 3) = "WG:" Or _
                      Left(trimmedSubject, 4) = "WG: ")  ' German forward prefix
End Function

'-------------------------------------------------------------------------------
' EXECUTE CLASSIFICATION ACTION
'-------------------------------------------------------------------------------

' Execute the classification decision on an email.
' Records every executed decision to decision_log.txt (source from
' lastClassifySource set by ClassifyEmail). resolvedAction (optional out)
' receives the final action after LLM_REVIEW resolution: KEEP / DELETE /
' MOVE_II / REVIEW — callers like the real-time handler use it to trigger
' auto-reply only on resolved KEEPs.
Public Sub ExecuteAction(ByVal mail As Outlook.MailItem, ByVal decision As String, _
                         Optional ByVal stats As Object = Nothing, _
                         Optional ByRef resolvedAction As String)
    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.ExecuteAction"

    ' Capture email info BEFORE any action (mail object becomes invalid after delete/move)
    Dim senderName As String
    Dim subject As String
    Dim senderEmail As String
    Dim source As String
    senderName = mail.senderName
    subject = mail.subject
    senderEmail = GetSenderEmail(mail)
    source = lastClassifySource
    If Len(source) = 0 Then source = "UNKNOWN"

    resolvedAction = decision

    Select Case decision
        Case "MOVE_II"
            LogActionDirect senderName, subject, "MOVED to " & RuntimeFolderProtected
            RecordDecision senderEmail, subject, source, "MOVE_II", 1
            mail.Move GetOrCreateFolder(RuntimeFolderProtected)
            If Not stats Is Nothing Then IncrementStat stats, "MOVE_II"

        Case "DELETE"
            LogActionDirect senderName, subject, "DELETED"
            RecordDecision senderEmail, subject, source, "DELETE", 1
            mail.Delete
            If Not stats Is Nothing Then IncrementStat stats, "DELETE"

        Case "LLM_REVIEW"
            ' Try LLM if configured, otherwise move to Review folder
            Dim llmDecision As String
            llmDecision = ProcessAmbiguousEmail(mail)

            If llmDecision = "DELETE" Then
                LogActionDirect senderName, subject, "DELETED (LLM: " & Truncate(lastLLMReason, 60) & ")"
                RecordDecision senderEmail, subject, "LLM", "DELETE", lastLLMConfidence
                mail.Delete
                resolvedAction = "DELETE"
                If Not stats Is Nothing Then IncrementStat stats, "DELETE"
            ElseIf llmDecision = "KEEP" Then
                LogActionDirect senderName, subject, "KEPT (LLM: " & Truncate(lastLLMReason, 60) & ")"
                RecordDecision senderEmail, subject, "LLM", "KEEP", lastLLMConfidence
                ApplyUrgencyMarkers mail
                resolvedAction = "KEEP"
                If Not stats Is Nothing Then IncrementStat stats, "KEEP"
            Else
                ' Move to Review folder
                LogActionDirect senderName, subject, "MOVED to " & RuntimeFolderReview
                RecordDecision senderEmail, subject, IIf(RuntimeUseLLM, "LLM", "DEFAULT"), "REVIEW", lastLLMConfidence
                mail.Move GetOrCreateFolder(RuntimeFolderReview)
                resolvedAction = "REVIEW"
                If Not stats Is Nothing Then IncrementStat stats, "REVIEW"
            End If

        Case "KEEP"
            ' Do nothing, leave in current folder
            LogActionDirect senderName, subject, "KEPT"
            RecordDecision senderEmail, subject, source, "KEEP", 1
            If Not stats Is Nothing Then IncrementStat stats, "KEEP"

        Case Else
            LogMessage "WARN", "Unknown decision: " & decision
            resolvedAction = "KEEP"
            If Not stats Is Nothing Then IncrementStat stats, "KEEP"
    End Select

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "ExecuteAction", Err.Number, Err.Description
    If Not stats Is Nothing Then IncrementStat stats, "ERROR"
    Resume PROC_EXIT
End Sub

' Mark high-urgency KEEP emails so triage is visible inside Outlook itself:
' urgency 4-5 -> "Urgent" category + high importance; urgency 3 -> "Action" category.
Private Sub ApplyUrgencyMarkers(ByVal mail As Outlook.MailItem)
    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.ApplyUrgencyMarkers"

    If lastLLMUrgency >= 4 Then
        If InStr(1, mail.Categories, "Urgent", vbTextCompare) = 0 Then
            If Len(mail.Categories) > 0 Then
                mail.Categories = mail.Categories & ";Urgent"
            Else
                mail.Categories = "Urgent"
            End If
        End If
        mail.Importance = olImportanceHigh
        mail.Save
    ElseIf lastLLMUrgency = 3 Then
        If InStr(1, mail.Categories, "Action", vbTextCompare) = 0 Then
            If Len(mail.Categories) > 0 Then
                mail.Categories = mail.Categories & ";Action"
            Else
                mail.Categories = "Action"
            End If
            mail.Save
        End If
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "ApplyUrgencyMarkers", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' LLM INTEGRATION
'-------------------------------------------------------------------------------

' Process an ambiguous email using LLM or fallback to Review folder
Private Function ProcessAmbiguousEmail(ByVal mail As Outlook.MailItem) As String
    Dim apiKey As String

    ' Reset structured outputs
    lastLLMConfidence = 0
    lastLLMUrgency = 0
    lastLLMCategory = ""
    lastLLMReason = ""

    ' Check if LLM is enabled and configured
    If Not RuntimeUseLLM Then
        ProcessAmbiguousEmail = "REVIEW"
        Exit Function
    End If

    ' Local provider does not require an API key (normalize like CallLLM does)
    If LCase(Trim(RuntimeLLMProvider)) <> "local" Then
        apiKey = GetAPIKey()
        If Len(apiKey) = 0 Then
            LogMessage "WARN", "LLM enabled but API key not found"
            ProcessAmbiguousEmail = "REVIEW"
            Exit Function
        End If
    End If

    ' Call LLM (structured classification)
    On Error GoTo FallbackToReview
    ProcessAmbiguousEmail = ClassifyViaLLMEx(mail)
    Exit Function

FallbackToReview:
    LogMessage "WARN", "LLM call failed: " & Err.Description
    ProcessAmbiguousEmail = "REVIEW"
End Function

' Build the structured classification prompt: email details + sender history
' evidence (AgentMemory) + past user corrections + strict JSON output format.
Private Function BuildEmailPrompt(ByVal mail As Outlook.MailItem) As String
    Dim prompt As String
    Dim senderContext As String
    Dim corrections As String

    prompt = "Classify this email:" & vbCrLf
    prompt = prompt & "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf
    prompt = prompt & "Subject: " & mail.subject & vbCrLf
    prompt = prompt & "Body preview: " & Truncate(mail.Body, RuntimeClassifyBodyChars) & vbCrLf

    ' Deterministic evidence about this sender (counts from decision_log.txt,
    ' replied-before signal from learned_replies.txt)
    senderContext = GetSenderContext(GetSenderEmail(mail), mail.senderName)
    If Len(senderContext) > 0 Then
        prompt = prompt & vbCrLf & senderContext & vbCrLf
    End If

    ' Few-shot corrections: LLM decisions the user has reversed
    corrections = GetRecentCorrectionsBlock(5)
    If Len(corrections) > 0 Then
        prompt = prompt & vbCrLf & corrections
    End If

    prompt = prompt & vbCrLf & _
        "Respond with ONLY a JSON object, no other text:" & vbCrLf & _
        "{""action"":""KEEP"" or ""DELETE""," & vbCrLf & _
        " ""category"":""student|colleague|admin|newsletter|external|other""," & vbCrLf & _
        " ""urgency"":1-5 (5 = needs response today)," & vbCrLf & _
        " ""confidence"":0.0-1.0 (how sure you are)," & vbCrLf & _
        " ""reason"":""max 15 words""}"

    BuildEmailPrompt = prompt
End Function

' Structured LLM classification. Returns "KEEP", "DELETE", or "REVIEW" and sets
' lastLLMConfidence / lastLLMUrgency / lastLLMCategory / lastLLMReason.
' Safety gates:
'   - JSON parse failure falls back to legacy substring matching at 0.5 confidence
'   - DELETE below RuntimeConfidenceThreshold is demoted to REVIEW (a human
'     looks at anything the model is unsure about deleting)
Public Function ClassifyViaLLMEx(ByVal mail As Outlook.MailItem) As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.ClassifyViaLLMEx"

    Dim result As String
    Dim action As String

    ' The user-configurable system prompt supplies the KEEP/DELETE policy; the
    ' appended override pins the output format regardless of how the legacy
    ' prompt phrases it.
    Dim systemPrompt As String
    systemPrompt = RuntimeLLMSystemPrompt & vbCrLf & _
                   "FORMAT OVERRIDE: You MUST respond with only the requested JSON object."

    result = CallLLM(BuildEmailPrompt(mail), systemPrompt, RuntimeClassifyMaxTokens)

    If Len(result) = 0 Then
        ClassifyViaLLMEx = "REVIEW"
        GoTo PROC_EXIT
    End If

    action = UCase(Trim(ExtractJSONStringValue(result, "action")))

    If action = "KEEP" Or action = "DELETE" Then
        lastLLMConfidence = ExtractJSONNumberValue(result, "confidence", 0.5)
        lastLLMUrgency = CLng(ExtractJSONNumberValue(result, "urgency", 0))
        lastLLMCategory = ExtractJSONStringValue(result, "category")
        lastLLMReason = ExtractJSONStringValue(result, "reason")
    Else
        ' Legacy fallback: model ignored the JSON format — substring match
        If InStr(1, UCase(result), "DELETE", vbTextCompare) > 0 Then
            action = "DELETE"
        ElseIf InStr(1, UCase(result), "KEEP", vbTextCompare) > 0 Then
            action = "KEEP"
        Else
            action = "REVIEW"
        End If
        lastLLMConfidence = 0.5
        lastLLMReason = Truncate(result, 100)
        LogMessage "DEBUG", "ClassifyViaLLMEx: non-JSON response, fell back to substring match"
    End If

    ' Confidence gate: uncertain DELETEs go to Review instead of the bin
    If action = "DELETE" And lastLLMConfidence < RuntimeConfidenceThreshold Then
        LogMessage "INFO", "ClassifyViaLLMEx: DELETE demoted to REVIEW (confidence " & _
                   Format(lastLLMConfidence, "0.00") & " < threshold " & _
                   Format(RuntimeConfidenceThreshold, "0.00") & ")"
        action = "REVIEW"
    End If

    ClassifyViaLLMEx = action

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailFilter", "ClassifyViaLLMEx", Err.Number, Err.Description
    ClassifyViaLLMEx = "REVIEW"
    Resume PROC_EXIT
End Function

' Backwards-compatible wrapper — delegates to CallLLM in Utilities.bas.
' Kept so any external macros calling this directly continue to work.
Public Function CallAzureOpenAICustom(ByVal userPrompt As String, ByVal systemPrompt As String, ByVal maxTokens As Integer) As String
    CallAzureOpenAICustom = CallLLM(userPrompt, systemPrompt, maxTokens)
End Function

'-------------------------------------------------------------------------------
' SINGLE EMAIL FILTER (for manual testing)
'-------------------------------------------------------------------------------

' Filter the currently selected email
Public Sub FilterSelectedEmail()
    Dim mail As Outlook.MailItem
    Dim decision As String

    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.FilterSelectedEmail"

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation
        GoTo PROC_EXIT
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation
        GoTo PROC_EXIT
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    decision = ClassifyEmail(mail)

    Dim response As VbMsgBoxResult
    response = MsgBox("Email: " & mail.subject & vbCrLf & vbCrLf & _
                      "Decision: " & decision & vbCrLf & vbCrLf & _
                      "Execute this action?", vbYesNo + vbQuestion, "Email Filter")

    If response = vbYes Then
        ExecuteAction mail, decision
        MsgBox "Action executed: " & decision, vbInformation
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "FilterSelectedEmail", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' LLM-POWERED EMAIL TOOLS
'-------------------------------------------------------------------------------

' Summarize the currently selected email using the LLM
Public Sub SummarizeSelectedEmail()
    Dim mail As Outlook.MailItem
    Dim prompt As String
    Dim summary As String
    Dim systemPrompt As String

    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.SummarizeSelectedEmail"

    If Not RuntimeUseLLM Then
        MsgBox "LLM is not enabled." & vbCrLf & vbCrLf & _
               "Set UseLLMAPI=True in settings.ini.", _
               vbExclamation, "Summarize Email"
        GoTo PROC_EXIT
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "Summarize Email"
        GoTo PROC_EXIT
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation, "Summarize Email"
        GoTo PROC_EXIT
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    systemPrompt = "You are a helpful assistant. Summarize the following email concisely in 2-3 bullet points. " & _
                   "Focus on: who sent it, what they want, and any action required."

    prompt = "Summarize this email:" & vbCrLf & _
             "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & mail.subject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    summary = CallLLM(prompt, systemPrompt, RuntimeSummarizeMaxTokens)

    If Len(summary) = 0 Then
        MsgBox "LLM returned no response. Check your API configuration.", vbExclamation, "Summarize Email"
        GoTo PROC_EXIT
    End If

    MsgBox "Summary of: " & mail.subject & vbCrLf & vbCrLf & summary, _
           vbInformation, "Email Summary"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "SummarizeSelectedEmail", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Draft a reply to the currently selected email using the few-shot engine.
' Delegates to DraftAutoReply (EmailAgent.bas) which uses learned reply
' examples from learned_replies.txt for style-consistent drafting.
Public Sub DraftReplyToSelected()
    Dim mail As Outlook.MailItem
    Dim sSubject As String

    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.DraftReplyToSelected"

    If Not RuntimeUseLLM Then
        MsgBox "LLM is not enabled." & vbCrLf & vbCrLf & _
               "Set UseLLMAPI=True in settings.ini.", _
               vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    Set mail = Application.ActiveExplorer.Selection(1)
    sSubject = mail.subject

    If DraftAutoReply(mail) Then
        MsgBox "Draft reply saved to Drafts folder for:" & vbCrLf & _
               Left(sSubject, 100), vbInformation, "Draft Reply"
    Else
        MsgBox "Could not draft reply. Check LLM configuration and logs.", _
               vbExclamation, "Draft Reply"
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "DraftReplyToSelected", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub
