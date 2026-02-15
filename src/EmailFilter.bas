'===============================================================================
' EmailFilter.bas - Main Classification Logic v2.0
'===============================================================================
' This module contains the core email classification functions,
' LLM integration, and LLM-powered email tools (summarize, draft reply).
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

    On Error GoTo ErrorHandler

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
    Exit Function

ErrorHandler:
    LogMessage "ERROR", "ClassifyEmail error: " & Err.Description
    ClassifyEmail = "KEEP"  ' Safe default: keep on error
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
    On Error GoTo ErrorHandler

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

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "ExecuteAction error: " & Err.Description & " | Subject: " & subject
    If Not stats Is Nothing Then IncrementStat stats, "ERROR"
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

    apiKey = GetAPIKey()
    If Len(apiKey) = 0 Then
        LogMessage "WARN", "LLM enabled but API key not found"
        ProcessAmbiguousEmail = "REVIEW"
        Exit Function
    End If

    ' Call LLM
    On Error GoTo FallbackToReview
    ProcessAmbiguousEmail = CallAzureOpenAI(BuildEmailPrompt(mail))
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
    prompt = prompt & "Body preview: " & Truncate(mail.Body, 300) & vbCrLf
    prompt = prompt & vbCrLf
    prompt = prompt & "Respond with ONLY 'DELETE' or 'KEEP' followed by brief reason."

    BuildEmailPrompt = prompt
End Function

' Call Azure OpenAI API (classification wrapper)
Private Function CallAzureOpenAI(ByVal prompt As String) As String
    Dim result As String
    result = CallAzureOpenAICustom(prompt, RuntimeLLMSystemPrompt, RuntimeLLMMaxTokens)

    ' Extract decision from LLM response
    If InStr(1, UCase(result), "DELETE", vbTextCompare) > 0 Then
        CallAzureOpenAI = "DELETE"
    ElseIf InStr(1, UCase(result), "KEEP", vbTextCompare) > 0 Then
        CallAzureOpenAI = "KEEP"
    Else
        CallAzureOpenAI = "REVIEW"
    End If
End Function

' Call Azure OpenAI API with custom system prompt and max tokens.
' Returns the raw content string from the API response.
' Shared by classification, summarization, and reply drafting.
Public Function CallAzureOpenAICustom(ByVal userPrompt As String, ByVal systemPrompt As String, ByVal maxTokens As Integer) As String
    Dim http As Object
    Dim requestBody As String
    Dim response As String
    Dim apiKey As String

    apiKey = GetAPIKey()

    If Len(apiKey) = 0 Then
        CallAzureOpenAICustom = ""
        LogMessage "WARN", "CallAzureOpenAICustom: no API key"
        Exit Function
    End If

    ' Build JSON request body
    requestBody = "{" & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & RuntimeLLMTemperature & _
    "}"

    LogMessage "DEBUG", "Calling Azure OpenAI..."

    ' Create HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", RuntimeLLMEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "api-key", apiKey
    http.send requestBody

    ' Parse response
    If http.Status = 200 Then
        response = ParseJSONContent(http.responseText)
        LogMessage "DEBUG", "LLM response: " & Left(response, 100)
        CallAzureOpenAICustom = response
    Else
        LogMessage "ERROR", "Azure OpenAI error: " & http.Status & " - " & http.statusText
        CallAzureOpenAICustom = ""
    End If

    Set http = Nothing
End Function

'-------------------------------------------------------------------------------
' SINGLE EMAIL FILTER (for manual testing)
'-------------------------------------------------------------------------------

' Filter the currently selected email
Public Sub FilterSelectedEmail()
    Dim mail As Outlook.MailItem
    Dim decision As String

    On Error GoTo ErrorHandler

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation
        Exit Sub
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation
        Exit Sub
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

    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
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

    On Error GoTo ErrorHandler

    If Not RuntimeUseLLM Then
        MsgBox "LLM is not enabled." & vbCrLf & vbCrLf & _
               "Enable it in the Dashboard Settings tab or set UseLLMAPI=True in settings.ini.", _
               vbExclamation, "Summarize Email"
        Exit Sub
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "Summarize Email"
        Exit Sub
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation, "Summarize Email"
        Exit Sub
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    systemPrompt = "You are a helpful assistant. Summarize the following email concisely in 2-3 bullet points. " & _
                   "Focus on: who sent it, what they want, and any action required."

    prompt = "Summarize this email:" & vbCrLf & _
             "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & mail.subject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    summary = CallAzureOpenAICustom(prompt, systemPrompt, 300)

    If Len(summary) = 0 Then
        MsgBox "LLM returned no response. Check your API configuration.", vbExclamation, "Summarize Email"
        Exit Sub
    End If

    MsgBox "Summary of: " & mail.subject & vbCrLf & vbCrLf & summary, _
           vbInformation, "Email Summary"

    Exit Sub

ErrorHandler:
    MsgBox "Error summarizing email: " & Err.Description, vbCritical, "Summarize Email"
End Sub

' Draft a reply to the currently selected email using the LLM
Public Sub DraftReplyToSelected()
    Dim mail As Outlook.MailItem
    Dim prompt As String
    Dim draft As String
    Dim systemPrompt As String

    On Error GoTo ErrorHandler

    If Not RuntimeUseLLM Then
        MsgBox "LLM is not enabled." & vbCrLf & vbCrLf & _
               "Enable it in the Dashboard Settings tab or set UseLLMAPI=True in settings.ini.", _
               vbExclamation, "Draft Reply"
        Exit Sub
    End If

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation, "Draft Reply"
        Exit Sub
    End If

    If Not TypeOf Application.ActiveExplorer.Selection(1) Is Outlook.MailItem Then
        MsgBox "Please select an email (not a meeting or other item).", vbExclamation, "Draft Reply"
        Exit Sub
    End If

    Set mail = Application.ActiveExplorer.Selection(1)

    systemPrompt = "You are Professor Xu Xin at PolyU Hong Kong. Draft a professional, concise reply to the following email. " & _
                   "Be polite and to the point. If the email requires a specific action, acknowledge it. " & _
                   "Do not include a subject line in your reply."

    prompt = "Draft a reply to this email:" & vbCrLf & _
             "From: " & mail.senderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & mail.subject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    draft = CallAzureOpenAICustom(prompt, systemPrompt, 500)

    If Len(draft) = 0 Then
        MsgBox "LLM returned no response. Check your API configuration.", vbExclamation, "Draft Reply"
        Exit Sub
    End If

    ' Try to show in frmDraftReply UserForm (late binding so it compiles without the form)
    Dim frm As Object
    On Error Resume Next
    Set frm = VBA.UserForms.Add("frmDraftReply")
    On Error GoTo ErrorHandler
    If Not frm Is Nothing Then
        frm.Initialize draft, mail
        frm.Show
    Else
        ' Fallback: show in MsgBox if UserForm not installed
        MsgBox "Draft Reply:" & vbCrLf & vbCrLf & Left(draft, 1000), vbInformation, "Draft Reply"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error drafting reply: " & Err.Description, vbCritical, "Draft Reply"
End Sub
