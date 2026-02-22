'===============================================================================
' EmailFilter.bas - Main Classification Logic v3.0
'===============================================================================
' This module contains the core email classification functions,
' LLM integration (via multi-provider CallLLM), and LLM email tools.
'
' All pattern/setting references use Runtime* variables from Config.bas,
' loaded from settings.ini by LoadAllSettings at startup.
'===============================================================================

Option Explicit

' Flag set by ClassifyEmail when a learned rule was applied (used by dry-run/reporting)
Public lastClassifyWasLearned As Boolean

' Flag set by ClassifyEmail when a learned subject rule was applied
Public lastClassifyWasLearnedSubject As Boolean

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

    ' Reset learned-rule flags
    lastClassifyWasLearned = False
    lastClassifyWasLearnedSubject = False

    ' =========================================================================
    ' RULE 0: LEARNED SENDER RULES (highest priority, from self-improving)
    ' =========================================================================
    If RuntimeEnableSelfImproving Then
        Dim learnedAction As String
        learnedAction = LookupLearnedSender(senderEmail)

        If learnedAction = "KEEP" Then
            ClassifyEmail = "KEEP"
            lastClassifyWasLearned = True
            LogMessage "DEBUG", "  -> KEEP (learned rule for: " & senderEmail & ")"
            Exit Function
        ElseIf learnedAction = "DELETE" Then
            ClassifyEmail = "DELETE"
            lastClassifyWasLearned = True
            LogMessage "DEBUG", "  -> DELETE (learned rule for: " & senderEmail & ")"
            Exit Function
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
            LogMessage "DEBUG", "  -> DELETE (learned subject rule match)"
            Exit Function
        End If
    End If

    ' =========================================================================
    ' RULE 1: PROTECTED DOMAINS - Move to Protected folder
    ' =========================================================================
    If IsProtectedDomain(domain) Then
        ClassifyEmail = "MOVE_II"
        LogMessage "DEBUG", "  -> MOVE_II (protected domain: " & domain & ")"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 2: PERSONALLY ADDRESSED - Keep
    ' =========================================================================
    If IsPersonallyAddressed(subject, bodyStart) Then
        ClassifyEmail = "KEEP"
        LogMessage "DEBUG", "  -> KEEP (personally addressed)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 3: ORGANIZATIONAL TAGS - Keep
    ' =========================================================================
    If ContainsAny(subject, RuntimePolyUTags) Then
        ClassifyEmail = "KEEP"
        LogMessage "DEBUG", "  -> KEEP (org tag found)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 4: VIP SUBJECT KEYWORDS - Keep
    ' =========================================================================
    If ContainsAny(subject, RuntimeVIPKeywords) Then
        ClassifyEmail = "KEEP"
        LogMessage "DEBUG", "  -> KEEP (VIP keyword)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 5: REPLY CHAINS - Keep (RE: subjects)
    ' =========================================================================
    If IsReplyEmail(subject) Then
        ClassifyEmail = "KEEP"
        LogMessage "DEBUG", "  -> KEEP (reply chain)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 6: FORWARD CHAINS - Keep (FW: subjects)
    ' =========================================================================
    If IsForwardEmail(subject) Then
        ClassifyEmail = "KEEP"
        LogMessage "DEBUG", "  -> KEEP (forwarded)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 7: KNOWN SPAM SENDERS - Delete
    ' =========================================================================
    If ContainsAny(senderName, RuntimeDeleteKnownSenders) Then
        ClassifyEmail = "DELETE"
        LogMessage "DEBUG", "  -> DELETE (known sender: " & senderName & ")"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 8: SENDER EMAIL PATTERNS - Delete
    ' =========================================================================
    If ContainsAny(senderEmail, RuntimeDeleteSenderPatterns) Then
        ClassifyEmail = "DELETE"
        LogMessage "DEBUG", "  -> DELETE (sender pattern)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 9: SUBJECT PATTERNS - Delete
    ' =========================================================================
    If ContainsAny(subject, RuntimeDeleteSubjectPatterns) Then
        ClassifyEmail = "DELETE"
        LogMessage "DEBUG", "  -> DELETE (subject pattern)"
        Exit Function
    End If

    ' =========================================================================
    ' RULE 10: AMBIGUOUS - Send to LLM or Review folder
    ' =========================================================================
    ClassifyEmail = "LLM_REVIEW"
    LogMessage "DEBUG", "  -> LLM_REVIEW (no rule matched)"

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailFilter", "ClassifyEmail", Err.Number, Err.Description
    ClassifyEmail = "KEEP"  ' Safe default: keep on error
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

' Execute the classification decision on an email
Public Sub ExecuteAction(ByVal mail As Outlook.MailItem, ByVal decision As String, Optional ByVal stats As Object = Nothing)
    On Error GoTo PROC_ERR
    PushCallStack "EmailFilter.ExecuteAction"

    ' Capture email info BEFORE any action (mail object becomes invalid after delete/move)
    Dim senderName As String
    Dim subject As String
    senderName = mail.senderName
    subject = mail.subject

    Select Case decision
        Case "MOVE_II"
            LogActionDirect senderName, subject, "MOVED to " & RuntimeFolderProtected
            mail.Move GetOrCreateFolder(RuntimeFolderProtected)
            If Not stats Is Nothing Then IncrementStat stats, "MOVE_II"

        Case "DELETE"
            LogActionDirect senderName, subject, "DELETED"
            mail.Delete
            If Not stats Is Nothing Then IncrementStat stats, "DELETE"

        Case "LLM_REVIEW"
            ' Try LLM if configured, otherwise move to Review folder
            Dim llmDecision As String
            llmDecision = ProcessAmbiguousEmail(mail)

            If llmDecision = "DELETE" Then
                LogActionDirect senderName, subject, "DELETED (LLM)"
                mail.Delete
                If Not stats Is Nothing Then IncrementStat stats, "DELETE"
            ElseIf llmDecision = "KEEP" Then
                LogActionDirect senderName, subject, "KEPT (LLM)"
                If Not stats Is Nothing Then IncrementStat stats, "KEEP"
            Else
                ' Move to Review folder
                LogActionDirect senderName, subject, "MOVED to " & RuntimeFolderReview
                mail.Move GetOrCreateFolder(RuntimeFolderReview)
                If Not stats Is Nothing Then IncrementStat stats, "REVIEW"
            End If

        Case "KEEP"
            ' Do nothing, leave in current folder
            LogActionDirect senderName, subject, "KEPT"
            If Not stats Is Nothing Then IncrementStat stats, "KEEP"

        Case Else
            LogMessage "WARN", "Unknown decision: " & decision
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

'-------------------------------------------------------------------------------
' LLM INTEGRATION
'-------------------------------------------------------------------------------

' Process an ambiguous email using LLM or fallback to Review folder
Private Function ProcessAmbiguousEmail(ByVal mail As Outlook.MailItem) As String
    Dim apiKey As String

    ' Check if LLM is enabled and configured
    If Not RuntimeUseLLM Then
        ProcessAmbiguousEmail = "REVIEW"
        Exit Function
    End If

    ' Local provider does not require an API key
    If RuntimeLLMProvider <> "local" Then
        apiKey = GetAPIKey()
        If Len(apiKey) = 0 Then
            LogMessage "WARN", "LLM enabled but API key not found"
            ProcessAmbiguousEmail = "REVIEW"
            Exit Function
        End If
    End If

    ' Call LLM
    On Error GoTo FallbackToReview
    ProcessAmbiguousEmail = ClassifyViaLLM(BuildEmailPrompt(mail))
    Exit Function

FallbackToReview:
    LogMessage "WARN", "LLM call failed: " & Err.Description
    ProcessAmbiguousEmail = "REVIEW"
End Function

' Build the prompt for LLM classification
Private Function BuildEmailPrompt(ByVal mail As Outlook.MailItem) As String
    Dim prompt As String

    prompt = "Classify this email:" & vbCrLf
    prompt = prompt & "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf
    prompt = prompt & "Subject: " & mail.subject & vbCrLf
    prompt = prompt & "Body preview: " & Truncate(mail.Body, RuntimeClassifyBodyChars) & vbCrLf
    prompt = prompt & vbCrLf
    prompt = prompt & "Respond with ONLY 'DELETE' or 'KEEP' followed by brief reason."

    BuildEmailPrompt = prompt
End Function

' Call LLM for classification — routes through CallLLM in Utilities.bas
Private Function ClassifyViaLLM(ByVal prompt As String) As String
    Dim result As String
    result = CallLLM(prompt, RuntimeLLMSystemPrompt, RuntimeClassifyMaxTokens)

    ' Extract decision from LLM response
    If InStr(1, UCase(result), "DELETE", vbTextCompare) > 0 Then
        ClassifyViaLLM = "DELETE"
    ElseIf InStr(1, UCase(result), "KEEP", vbTextCompare) > 0 Then
        ClassifyViaLLM = "KEEP"
    Else
        ClassifyViaLLM = "REVIEW"
    End If
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

' Draft a reply to the currently selected email using the LLM
Public Sub DraftReplyToSelected()
    Dim mail As Outlook.MailItem
    Dim replyItem As Outlook.MailItem
    Dim prompt As String
    Dim draft As String
    Dim systemPrompt As String
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

    systemPrompt = "You are Professor Xu Xin at PolyU Hong Kong. Draft a professional, concise reply to the following email. " & _
                   "Be polite and to the point. If the email requires a specific action, acknowledge it. " & _
                   "Do not include a subject line in your reply."

    prompt = "Draft a reply to this email:" & vbCrLf & _
             "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & sSubject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    draft = CallLLM(prompt, systemPrompt, RuntimeReplyMaxTokens, RuntimeReplyTemperature)

    If Len(draft) = 0 Then
        MsgBox "LLM returned no response. Check your API configuration.", vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    ' Create the reply draft before showing MsgBox (mail reference is still fresh)
    Set replyItem = mail.Reply
    replyItem.Body = draft & vbCrLf & vbCrLf & replyItem.Body
    replyItem.Save
    Set replyItem = Nothing

    MsgBox "Draft reply saved to Drafts folder." & vbCrLf & vbCrLf & Left(draft, 1000), vbInformation, "Draft Reply"

    LogMessage "INFO", "Draft reply saved to Drafts for: " & Left(sSubject, 50)

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailFilter", "DraftReplyToSelected", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub
