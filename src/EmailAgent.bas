'===============================================================================
' EmailAgent.bas - AI Agent Features v3.0
'===============================================================================
' New agent capabilities added in v3.0:
'   - GenerateAddressingPatterns: LLM-generates personal addressing patterns
'   - DraftAutoReply: Few-shot reply drafting using learned_replies.txt examples
'   - ScanSentForReplyPatterns: Bulk-imports reply pairs from Sent Items
'   - DraftReplyForSelected: Draft replies for selected email(s) via few-shot engine
'
' Depends on:
'   - CallLLM (Utilities.bas) for all LLM calls
'   - RecordLearnedReply / LoadRecentReplyPairs (Utilities.bas) for reply I/O
'   - WriteINISetting / LoadAllSettings (Utilities.bas) for settings updates
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' GENERATE ADDRESSING PATTERNS
'-------------------------------------------------------------------------------

' LLM-generates comprehensive personal addressing patterns and saves them to
' settings.ini. Interactive version: prompts for inputs and confirms before
' overwriting existing patterns. Run once when setting up for a new professor.
Public Sub GenerateAddressingPatterns()
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.GenerateAddressingPatterns"

    If Not RuntimeUseLLM Then
        MsgBox "LLM must be enabled to generate addressing patterns." & vbCrLf & _
               "Set UseLLMAPI=True in settings.ini and configure a provider.", _
               vbExclamation, "Generate Patterns"
        GoTo PROC_EXIT
    End If

    ' Gather inputs
    Dim fullName As String
    fullName = Trim(InputBox("Enter your full name (e.g., Yu Yin):", "Addressing Patterns - Step 1 of 3"))
    If Len(fullName) = 0 Then GoTo PROC_EXIT

    Dim title As String
    title = Trim(InputBox("Your title (e.g., Professor, Dr., Mr.) - leave blank to skip:", "Addressing Patterns - Step 2 of 3"))

    Dim role As String
    role = Trim(InputBox("Your role/position (e.g., Head, Director, Dean) - leave blank to skip:", "Addressing Patterns - Step 3 of 3"))

    Dim namePatterns As String
    Dim greetingPatterns As String
    Dim genError As String
    genError = GeneratePatternsViaLLM(fullName, title, role, namePatterns, greetingPatterns)

    If Len(genError) > 0 Then
        MsgBox genError, vbExclamation, "Generate Patterns"
        GoTo PROC_EXIT
    End If

    ' Show preview and ask for confirmation
    Dim preview As String
    preview = "Generated addressing patterns for: " & fullName & vbCrLf & vbCrLf & _
              "NAME_PATTERNS (" & CountCommas(namePatterns) + 1 & " patterns):" & vbCrLf & _
              WrapText(namePatterns, 70) & vbCrLf & vbCrLf & _
              "GREETING_PATTERNS (" & CountCommas(greetingPatterns) + 1 & " patterns):" & vbCrLf & _
              WrapText(greetingPatterns, 70) & vbCrLf & vbCrLf & _
              "Save these to settings.ini and reload? (This will replace existing patterns.)"

    If MsgBox(preview, vbYesNo + vbQuestion, "Confirm Addressing Patterns") = vbNo Then
        GoTo PROC_EXIT
    End If

    SavePatternsToSettings namePatterns, greetingPatterns

    MsgBox "Addressing patterns saved and loaded." & vbCrLf & vbCrLf & _
           "The filter will now recognise emails addressed to """ & fullName & """.", _
           vbInformation, "Generate Patterns"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailAgent", "GenerateAddressingPatterns", Err.Number, Err.Description
    MsgBox "Error generating patterns: " & Err.Description, vbCritical, "Generate Patterns"
    Resume PROC_EXIT
End Sub

' Bridge variant: takes name/title/role as arguments (no InputBox, no confirm
' dialog — the Web UI supplies the inputs). Returns a summary or "ERROR: ...".
Public Function GenerateAddressingPatternsStd(ByVal fullName As String, _
                                              ByVal title As String, _
                                              ByVal role As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.GenerateAddressingPatternsStd"

    If Not RuntimeUseLLM Then
        GenerateAddressingPatternsStd = "ERROR: LLM must be enabled (UseLLMAPI=True) to generate addressing patterns."
        GoTo PROC_EXIT
    End If

    If Len(Trim(fullName)) = 0 Then
        GenerateAddressingPatternsStd = "ERROR: name argument is required."
        GoTo PROC_EXIT
    End If

    Dim namePatterns As String
    Dim greetingPatterns As String
    Dim genError As String
    genError = GeneratePatternsViaLLM(Trim(fullName), Trim(title), Trim(role), namePatterns, greetingPatterns)

    If Len(genError) > 0 Then
        GenerateAddressingPatternsStd = "ERROR: " & genError
        GoTo PROC_EXIT
    End If

    SavePatternsToSettings namePatterns, greetingPatterns

    GenerateAddressingPatternsStd = "Addressing patterns saved and loaded for """ & fullName & """." & vbCrLf & _
        "NAME_PATTERNS: " & namePatterns & vbCrLf & _
        "GREETING_PATTERNS: " & greetingPatterns

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailAgent", "GenerateAddressingPatternsStd", Err.Number, Err.Description
    GenerateAddressingPatternsStd = "ERROR: pattern generation failed: " & Err.Description
    Resume PROC_EXIT
End Function

' Shared LLM call + parse for both pattern-generation entry points.
' Returns "" on success (out-params filled) or an error message.
Private Function GeneratePatternsViaLLM(ByVal fullName As String, ByVal title As String, _
                                        ByVal role As String, _
                                        ByRef namePatterns As String, _
                                        ByRef greetingPatterns As String) As String
    Dim roleHint As String
    If Len(role) > 0 Then roleHint = " The person also has the role: " & role & "."

    Dim titleHint As String
    If Len(title) > 0 Then titleHint = " Their title is: " & title & "."

    Dim userPrompt As String
    userPrompt = "Given the name """ & fullName & """" & titleHint & roleHint & _
                 ", generate ALL possible ways this person might be addressed in professional and academic emails." & vbCrLf & _
                 "Consider both Western name order (first last) and Eastern name order (last first)." & vbCrLf & _
                 "Include full name variations, title+name, informal greetings, and formal salutations." & vbCrLf & _
                 "If a role was provided, include role-based salutations (e.g., Dear Head, Dear Director)." & vbCrLf & vbCrLf & _
                 "Respond in EXACTLY this format (comma-separated, no quotes, no extra text):" & vbCrLf & _
                 "NAME_PATTERNS: pattern1,pattern2,pattern3,..." & vbCrLf & _
                 "GREETING_PATTERNS: greeting1,greeting2,greeting3,..."

    Dim systemPrompt As String
    systemPrompt = "You are a helpful assistant generating email addressing patterns for academic email filtering software. " & _
                   "Be thorough and include all plausible variations including common misspellings and abbreviations."

    LogMessage "INFO", "Calling LLM to generate addressing patterns for: " & fullName

    Dim llmResponse As String
    llmResponse = CallLLM(userPrompt, systemPrompt, 400)

    If Len(llmResponse) = 0 Then
        GeneratePatternsViaLLM = "LLM returned no response. Check your API configuration."
        Exit Function
    End If

    namePatterns = ExtractPatternList(llmResponse, "NAME_PATTERNS:")
    greetingPatterns = ExtractPatternList(llmResponse, "GREETING_PATTERNS:")

    If Len(namePatterns) = 0 And Len(greetingPatterns) = 0 Then
        GeneratePatternsViaLLM = "Could not parse LLM response. Raw response: " & Left(llmResponse, 400)
        Exit Function
    End If

    GeneratePatternsViaLLM = ""
End Function

' Persist generated patterns and reload settings so they take effect immediately
Private Sub SavePatternsToSettings(ByVal namePatterns As String, ByVal greetingPatterns As String)
    If Len(namePatterns) > 0 Then
        WriteINISetting "Patterns", "NamePatterns", namePatterns
    End If
    If Len(greetingPatterns) > 0 Then
        WriteINISetting "Patterns", "GreetingPatterns", greetingPatterns
    End If
    LoadAllSettings
End Sub

' Extract a comma-separated list from a line like "NAME_PATTERNS: a,b,c"
Private Function ExtractPatternList(ByVal response As String, ByVal label As String) As String
    Dim labelPos As Long
    Dim lineEnd As Long
    Dim lineContent As String

    labelPos = InStr(1, response, label, vbTextCompare)
    If labelPos = 0 Then
        ExtractPatternList = ""
        Exit Function
    End If

    ' Find start of list (after the label)
    Dim listStart As Long
    listStart = labelPos + Len(label)

    ' Find end of line
    lineEnd = InStr(listStart, response, vbCrLf)
    If lineEnd = 0 Then
        lineEnd = InStr(listStart, response, vbLf)
    End If
    If lineEnd = 0 Then
        lineEnd = Len(response) + 1
    End If

    lineContent = Trim(Mid(response, listStart, lineEnd - listStart))
    ExtractPatternList = lineContent
End Function

' Count commas in a string (used to determine pattern count)
Private Function CountCommas(ByVal s As String) As Integer
    Dim i As Integer
    Dim count As Integer
    count = 0
    For i = 1 To Len(s)
        If Mid(s, i, 1) = "," Then count = count + 1
    Next i
    CountCommas = count
End Function

' Wrap long comma-separated text into multiple lines for display
Private Function WrapText(ByVal text As String, ByVal maxWidth As Integer) As String
    Dim parts() As String
    Dim result As String
    Dim currentLine As String
    Dim i As Integer
    Dim part As String

    parts = Split(text, ",")
    result = ""
    currentLine = ""

    For i = 0 To UBound(parts)
        part = Trim(parts(i))
        If Len(currentLine) + Len(part) + 2 > maxWidth Then
            result = result & currentLine & vbCrLf
            currentLine = "  " & part
        Else
            If Len(currentLine) = 0 Then
                currentLine = part
            Else
                currentLine = currentLine & ", " & part
            End If
        End If
    Next i

    If Len(currentLine) > 0 Then result = result & currentLine

    WrapText = result
End Function

'-------------------------------------------------------------------------------
' DRAFT AUTO REPLY (Few-Shot Style)
'-------------------------------------------------------------------------------

' Draft a reply to the given email using learned reply examples as few-shot
' context. Saves the draft to Outlook's Drafts folder.
' Returns True on success, False on failure.
Public Function DraftAutoReply(ByVal mail As Outlook.MailItem) As Boolean
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.DraftAutoReply"

    DraftAutoReply = False

    If Not RuntimeUseLLM Then
        LogMessage "WARN", "DraftAutoReply: LLM not enabled"
        GoTo PROC_EXIT
    End If

    ' Check sender filter (if set, only draft for specified senders)
    If Len(RuntimeAutoReplyForSenders) > 0 Then
        If Not ContainsAny(GetSenderEmail(mail), RuntimeAutoReplyForSenders) And _
           Not ContainsAny(mail.SenderName, RuntimeAutoReplyForSenders) Then
            LogMessage "DEBUG", "DraftAutoReply: sender not in AutoReplyForSenders list, skipping"
            GoTo PROC_EXIT
        End If
    End If

    ' Load reply examples
    Dim examples As Collection
    Set examples = LoadRecentReplyPairs(RuntimeMaxReplyExamples)

    ' Build persona string
    Dim persona As String
    persona = RuntimeReplyPersona
    If Len(Trim(persona)) = 0 Then
        ' Auto-generate from name patterns (first pattern as the name)
        If Len(RuntimeNamePatterns) > 0 Then
            persona = "You are " & Split(RuntimeNamePatterns, ",")(0) & ", a professor."
        Else
            persona = "You are a university professor."
        End If
    End If

    ' Build the few-shot prompt
    Dim userPrompt As String
    userPrompt = BuildReplyPrompt(mail, examples, persona)

    Dim systemPrompt As String
    systemPrompt = "You draft professional academic email replies. Match the tone and style of the provided examples. " & _
                   "Write only the reply body text. Do not include subject line, salutation ""Dear..."", or signature unless the examples show them."

    LogMessage "INFO", "Drafting auto-reply for: " & Left(mail.Subject, 50)

    Dim draftText As String
    draftText = CallLLM(userPrompt, systemPrompt, RuntimeReplyMaxTokens, RuntimeReplyTemperature)

    If Len(draftText) = 0 Then
        LogMessage "WARN", "DraftAutoReply: LLM returned empty response"
        GoTo PROC_EXIT
    End If

    ' Create Outlook reply draft and save to Drafts folder
    Dim replyItem As Outlook.MailItem
    Set replyItem = mail.Reply
    ' Prepend draft text before the quoted original message
    replyItem.Body = draftText & vbCrLf & vbCrLf & replyItem.Body
    replyItem.Save  ' Saves to Drafts folder automatically
    Set replyItem = Nothing

    LogMessage "INFO", "Auto-drafted reply saved to Drafts for: " & Left(mail.Subject, 50)
    DraftAutoReply = True

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailAgent", "DraftAutoReply", Err.Number, Err.Description
    DraftAutoReply = False
    Resume PROC_EXIT
End Function

' Build the few-shot reply prompt from examples and the incoming email
Private Function BuildReplyPrompt(ByVal mail As Outlook.MailItem, _
                                   ByVal examples As Collection, _
                                   ByVal persona As String) As String
    Dim prompt As String

    prompt = "Persona: " & persona & vbCrLf & vbCrLf

    If examples.Count > 0 Then
        prompt = prompt & "Here are examples of how I typically reply to emails:" & vbCrLf & vbCrLf

        Dim i As Integer
        For i = 1 To examples.Count
            Dim parts() As String
            parts = Split(examples(i), "|")
            If UBound(parts) >= 3 Then
                prompt = prompt & "Example " & i & ":" & vbCrLf
                prompt = prompt & "  Original from: " & parts(1) & vbCrLf
                prompt = prompt & "  Original subject: " & parts(0) & vbCrLf
                prompt = prompt & "  Original body: " & parts(2) & vbCrLf
                prompt = prompt & "  My reply: " & parts(3) & vbCrLf & vbCrLf
            End If
        Next i
    End If

    prompt = prompt & "Now draft a reply to this new email:" & vbCrLf & vbCrLf
    prompt = prompt & "From: " & mail.SenderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf
    prompt = prompt & "Subject: " & mail.Subject & vbCrLf
    prompt = prompt & "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf
    prompt = prompt & "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    BuildReplyPrompt = prompt
End Function

'-------------------------------------------------------------------------------
' SCAN SENT ITEMS FOR REPLY PATTERNS
'-------------------------------------------------------------------------------

' Interactive wrapper — shows the scan summary in a MsgBox
Public Sub ScanSentForReplyPatterns()
    Dim result As String
    result = ScanSentForReplyPatternsCore()
    MsgBox result, IIf(Left(result, 6) = "ERROR:", vbCritical, vbInformation), "Scan Sent Items"
End Sub

' Scan Sent Items for reply emails and extract original/reply pairs into
' learned_replies.txt. Uses a simple heuristic: sent emails with RE: subjects
' where the body contains a "From:" / "Sent:" delimiter.
' Headless: returns a summary string ("ERROR: ..." on failure).
Public Function ScanSentForReplyPatternsCore() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.ScanSentForReplyPatternsCore"

    Dim ns As Outlook.NameSpace
    Dim sentFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim scannedCount As Long
    Dim learnedCount As Long
    Dim cutoffDate As Date

    Set ns = Application.GetNamespace("MAPI")
    Set sentFolder = ns.GetDefaultFolder(olFolderSentMail)
    Set myItems = sentFolder.Items
    myItems.Sort "[SentOn]", True  ' Newest first

    cutoffDate = Now - RuntimeScanSentDays
    scannedCount = 0
    learnedCount = 0

    LogMessage "INFO", "ScanSentForReplyPatterns: scanning last " & RuntimeScanSentDays & " days in Sent Items"

    For i = 1 To myItems.Count
        If Not TypeOf myItems(i) Is Outlook.MailItem Then GoTo NextItem

        Set mail = myItems(i)

        ' Stop when we go past the cutoff date
        If mail.SentOn < cutoffDate Then Exit For

        scannedCount = scannedCount + 1

        ' Only process reply emails (RE: prefix)
        If Not IsReplySubject(mail.Subject) Then GoTo NextItem

        ' Extract reply pair from body
        Dim myReplyText As String
        Dim originalBodySnippet As String
        Dim originalSubject As String
        Dim originalFrom As String

        ' Strip RE:/AW: prefix to get the original subject
        originalSubject = StripReplyPrefix(mail.Subject)

        ' The "To" field of the sent reply is the original sender
        If mail.Recipients.Count > 0 Then
            originalFrom = mail.Recipients(1).Name
        Else
            originalFrom = ""
        End If

        ' Split body on common reply delimiter lines
        myReplyText = ExtractReplyTextFromBody(mail.Body)
        originalBodySnippet = ExtractOriginalBodySnippet(mail.Body)

        If Len(Trim(myReplyText)) > 20 Then  ' Skip trivially short replies
            RecordLearnedReply originalSubject, originalFrom, originalBodySnippet, myReplyText
            learnedCount = learnedCount + 1
        End If

NextItem:
    Next i

    ' New reply pairs change the replied-to set used for context enrichment
    If learnedCount > 0 Then InvalidateRepliedToCache

    LogMessage "INFO", "ScanSentForReplyPatterns: scanned " & scannedCount & ", learned " & learnedCount & " reply pairs"

    ScanSentForReplyPatternsCore = "Scan complete. Scanned: " & scannedCount & " sent emails, learned: " & _
                                   learnedCount & " reply pairs." & vbCrLf & _
                                   "File: " & GetLearnedRepliesFilePath()

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailAgent", "ScanSentForReplyPatternsCore", Err.Number, Err.Description
    ScanSentForReplyPatternsCore = "ERROR: Sent Items scan failed: " & Err.Description
    Resume PROC_EXIT
End Function

' Check if subject has a reply prefix
Private Function IsReplySubject(ByVal subject As String) As Boolean
    Dim u As String
    u = UCase(Trim(subject))
    IsReplySubject = (Left(u, 3) = "RE:" Or Left(u, 3) = "AW:")
End Function

' Strip RE: / AW: prefix(es) from subject
Private Function StripReplyPrefix(ByVal subject As String) As String
    Dim result As String
    result = Trim(subject)
    Do While UCase(Left(result, 3)) = "RE:" Or UCase(Left(result, 3)) = "AW:"
        result = Trim(Mid(result, 4))
    Loop
    StripReplyPrefix = result
End Function

' Extract the user's reply text: everything before the first quoted-message delimiter.
' Looks for "From:", "-----Original Message-----", "________________________________" etc.
' Public: shared with ThisOutlookSession's LearnReply folder watcher (the
' previous private copies there had drifted from these).
Public Function ExtractReplyTextFromBody(ByVal body As String) As String
    Dim delimiters(4) As String
    delimiters(0) = vbCrLf & "From:"
    delimiters(1) = vbLf & "From:"
    delimiters(2) = "-----Original Message-----"
    delimiters(3) = "________________________________"
    delimiters(4) = vbCrLf & "Sent:"

    Dim earliest As Long
    Dim pos As Long
    Dim d As Integer

    earliest = 0

    For d = 0 To 4
        pos = InStr(1, body, delimiters(d), vbTextCompare)
        If pos > 0 Then
            If earliest = 0 Or pos < earliest Then
                earliest = pos
            End If
        End If
    Next d

    If earliest > 1 Then
        ExtractReplyTextFromBody = Trim(Left(body, earliest - 1))
    Else
        ExtractReplyTextFromBody = ""
    End If
End Function

' Extract snippet of original message (text after the first delimiter)
' Public: shared with ThisOutlookSession's LearnReply folder watcher.
Public Function ExtractOriginalBodySnippet(ByVal body As String) As String
    Dim delimiters(4) As String
    delimiters(0) = vbCrLf & "From:"
    delimiters(1) = vbLf & "From:"
    delimiters(2) = "-----Original Message-----"
    delimiters(3) = "________________________________"
    delimiters(4) = vbCrLf & "Sent:"

    Dim earliest As Long
    Dim pos As Long
    Dim d As Integer

    earliest = 0

    For d = 0 To 4
        pos = InStr(1, body, delimiters(d), vbTextCompare)
        If pos > 0 Then
            If earliest = 0 Or pos < earliest Then
                earliest = pos
            End If
        End If
    Next d

    If earliest > 0 And earliest + 10 < Len(body) Then
        ExtractOriginalBodySnippet = Truncate(Trim(Mid(body, earliest)), 500)
    Else
        ExtractOriginalBodySnippet = ""
    End If
End Function

'-------------------------------------------------------------------------------
' DRAFT REPLY FOR SELECTED EMAILS
'-------------------------------------------------------------------------------

' Draft few-shot replies for the currently selected email(s).
' Uses DraftAutoReply (learned reply examples from learned_replies.txt).
' Assign to QAT for one-click reply drafting.
Public Sub DraftReplyForSelected()
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.DraftReplyForSelected"

    If Not RuntimeUseLLM Then
        MsgBox "LLM must be enabled for reply drafting." & vbCrLf & _
               "Set UseLLMAPI=True in settings.ini.", _
               vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    Dim sel As Outlook.Selection
    Set sel = Application.ActiveExplorer.Selection

    If sel.Count = 0 Then
        MsgBox "Please select one or more emails first.", vbExclamation, "Draft Reply"
        GoTo PROC_EXIT
    End If

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Draft LLM replies for " & sel.Count & " selected email(s)?" & vbCrLf & _
                     "Drafts will be saved to your Drafts folder.", _
                     vbYesNo + vbQuestion, "Draft Reply")
    If confirm = vbNo Then GoTo PROC_EXIT

    Dim i As Long
    Dim draftedCount As Long
    Dim skippedCount As Long
    Dim mail As Outlook.MailItem

    draftedCount = 0
    skippedCount = 0

    For i = 1 To sel.Count
        If Not TypeOf sel(i) Is Outlook.MailItem Then
            skippedCount = skippedCount + 1
            GoTo NextSelectedItem
        End If

        Set mail = sel(i)

        If DraftAutoReply(mail) Then
            draftedCount = draftedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If

NextSelectedItem:
    Next i

    MsgBox "Done." & vbCrLf & vbCrLf & _
           "Drafted: " & draftedCount & " replies" & vbCrLf & _
           "Skipped: " & skippedCount & vbCrLf & vbCrLf & _
           "Check your Drafts folder.", _
           vbInformation, "Draft Reply"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailAgent", "DraftReplyForSelected", Err.Number, Err.Description
    MsgBox "Error drafting replies: " & Err.Description, vbCritical, "Draft Reply"
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' LEARNED REPLY MANAGEMENT MACROS
'-------------------------------------------------------------------------------

' Show count and file path for learned replies (interactive wrapper)
Public Sub ShowLearnedRepliesSummary()
    MsgBox ShowLearnedRepliesSummaryCore(), vbInformation, "Learned Replies"
End Sub

' Headless summary of the learned reply corpus (used by the Web UI bridge)
Public Function ShowLearnedRepliesSummaryCore() As String
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.ShowLearnedRepliesSummaryCore"

    Dim filePath As String
    Dim lineCount As Long
    Dim lines As Variant
    Dim i As Long

    filePath = GetLearnedRepliesFilePath()

    Dim content As String
    content = ReadTextFileSmart(filePath)

    If Len(content) = 0 Then
        ShowLearnedRepliesSummaryCore = "No learned replies yet. Expected at: " & filePath & vbCrLf & _
            "Drag sent reply emails into the '" & RuntimeFolderLearnReply & "' folder, " & _
            "or run ScanSentForReplyPatterns."
        GoTo PROC_EXIT
    End If

    lineCount = 0
    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 And Left(Trim(lines(i)), 1) <> "#" Then lineCount = lineCount + 1
    Next i

    ShowLearnedRepliesSummaryCore = "Learned reply pairs: " & lineCount & vbCrLf & _
        "File: " & filePath & vbCrLf & _
        "Using top " & RuntimeMaxReplyExamples & " most-recent examples per draft."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "EmailAgent", "ShowLearnedRepliesSummaryCore", Err.Number, Err.Description
    ShowLearnedRepliesSummaryCore = "ERROR: could not read learned replies: " & Err.Description
    Resume PROC_EXIT
End Function
