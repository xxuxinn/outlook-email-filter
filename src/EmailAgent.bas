'===============================================================================
' EmailAgent.bas - AI Agent Features v3.0
'===============================================================================
' New agent capabilities added in v3.0:
'   - GenerateAddressingPatterns: LLM-generates personal addressing patterns
'   - DraftAutoReply: Few-shot reply drafting using learned_replies.txt examples
'   - ScanSentForReplyPatterns: Bulk-imports reply pairs from Sent Items
'   - DraftRepliesForInbox: Batch auto-draft replies for unread KEEP emails
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
' settings.ini. Run this once when setting up for a new professor.
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

    ' Build LLM prompt
    Dim roleHint As String
    roleHint = ""
    If Len(role) > 0 Then
        roleHint = " The person also has the role: " & role & "."
    End If

    Dim titleHint As String
    titleHint = ""
    If Len(title) > 0 Then
        titleHint = " Their title is: " & title & "."
    End If

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
        MsgBox "LLM returned no response. Check your API configuration.", vbExclamation, "Generate Patterns"
        GoTo PROC_EXIT
    End If

    ' Parse the two comma-separated lists
    Dim namePatterns As String
    Dim greetingPatterns As String

    namePatterns = ExtractPatternList(llmResponse, "NAME_PATTERNS:")
    greetingPatterns = ExtractPatternList(llmResponse, "GREETING_PATTERNS:")

    If Len(namePatterns) = 0 And Len(greetingPatterns) = 0 Then
        MsgBox "Could not parse LLM response. Raw response:" & vbCrLf & vbCrLf & Left(llmResponse, 800), _
               vbExclamation, "Generate Patterns"
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

    ' Write to settings.ini
    If Len(namePatterns) > 0 Then
        WriteINISetting "Patterns", "NamePatterns", namePatterns
    End If
    If Len(greetingPatterns) > 0 Then
        WriteINISetting "Patterns", "GreetingPatterns", greetingPatterns
    End If

    ' Reload settings so new patterns take effect immediately
    LoadAllSettings

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

' Scan Sent Items for reply emails and extract original/reply pairs into
' learned_replies.txt. Uses a simple heuristic: sent emails with RE: subjects
' where the body contains a "From:" / "Sent:" delimiter.
Public Sub ScanSentForReplyPatterns()
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.ScanSentForReplyPatterns"

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

    LogMessage "INFO", "ScanSentForReplyPatterns: scanned " & scannedCount & ", learned " & learnedCount & " reply pairs"

    MsgBox "Scan complete." & vbCrLf & vbCrLf & _
           "Scanned: " & scannedCount & " sent emails" & vbCrLf & _
           "Learned: " & learnedCount & " reply pairs" & vbCrLf & vbCrLf & _
           "File: " & GetLearnedRepliesFilePath(), _
           vbInformation, "Scan Sent Items"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailAgent", "ScanSentForReplyPatterns", Err.Number, Err.Description
    MsgBox "Error scanning sent items: " & Err.Description, vbCritical, "Scan Sent Items"
    Resume PROC_EXIT
End Sub

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
Private Function ExtractReplyTextFromBody(ByVal body As String) As String
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
Private Function ExtractOriginalBodySnippet(ByVal body As String) As String
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
' BATCH AUTO-REPLY FOR INBOX
'-------------------------------------------------------------------------------

' Scan Inbox for unread KEEP emails and draft replies for all of them.
' Use this for bulk reply-drafting on a batch of new emails.
Public Sub DraftRepliesForInbox()
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.DraftRepliesForInbox"

    If Not RuntimeEnableAutoReply Then
        MsgBox "Auto-reply is disabled. Set EnableAutoReply=True in settings.ini [Agent] section.", _
               vbExclamation, "Draft Replies"
        GoTo PROC_EXIT
    End If

    If Not RuntimeUseLLM Then
        MsgBox "LLM must be enabled for auto-reply drafting.", vbExclamation, "Draft Replies"
        GoTo PROC_EXIT
    End If

    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim draftedCount As Long
    Dim skippedCount As Long

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = inbox.Items
    myItems.Sort "[ReceivedTime]", True  ' Newest first

    draftedCount = 0
    skippedCount = 0

    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("This will draft LLM replies for all unread KEEP emails in your Inbox." & vbCrLf & _
                     "Drafts will be saved to your Drafts folder." & vbCrLf & vbCrLf & _
                     "Continue?", vbYesNo + vbQuestion, "Draft Replies for Inbox")
    If confirm = vbNo Then GoTo PROC_EXIT

    For i = 1 To myItems.Count
        If Not TypeOf myItems(i) Is Outlook.MailItem Then GoTo NextInboxItem

        Set mail = myItems(i)

        ' Only unread emails
        If Not mail.UnRead Then GoTo NextInboxItem

        ' Only KEEP decisions
        Dim decision As String
        decision = ClassifyEmail(mail)
        If decision <> "KEEP" Then GoTo NextInboxItem

        ' Draft the reply
        If DraftAutoReply(mail) Then
            draftedCount = draftedCount + 1
        Else
            skippedCount = skippedCount + 1
        End If

NextInboxItem:
    Next i

    MsgBox "Done." & vbCrLf & vbCrLf & _
           "Drafted: " & draftedCount & " replies" & vbCrLf & _
           "Skipped: " & skippedCount & vbCrLf & vbCrLf & _
           "Check your Drafts folder.", _
           vbInformation, "Draft Replies for Inbox"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailAgent", "DraftRepliesForInbox", Err.Number, Err.Description
    MsgBox "Error drafting replies: " & Err.Description, vbCritical, "Draft Replies"
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' LEARNED REPLY MANAGEMENT MACROS
'-------------------------------------------------------------------------------

' Show count and file path for learned replies
Public Sub ShowLearnedRepliesSummary()
    On Error GoTo PROC_ERR
    PushCallStack "EmailAgent.ShowLearnedRepliesSummary"

    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim lineCount As Long
    Dim line As String

    filePath = GetLearnedRepliesFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(filePath) Then
        MsgBox "No learned replies file found." & vbCrLf & vbCrLf & _
               "Expected at: " & filePath & vbCrLf & vbCrLf & _
               "Drag sent reply emails into the '" & RuntimeFolderLearnReply & "' folder, " & _
               "or run ScanSentForReplyPatterns.", _
               vbInformation, "Learned Replies"
        GoTo PROC_EXIT
    End If

    ' Count lines
    lineCount = 0
    Set ts = fso.OpenTextFile(filePath, 1)
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then lineCount = lineCount + 1
    Loop

    MsgBox "Learned reply pairs: " & lineCount & vbCrLf & vbCrLf & _
           "File: " & filePath & vbCrLf & vbCrLf & _
           "Using top " & RuntimeMaxReplyExamples & " most-recent examples per draft.", _
           vbInformation, "Learned Replies"

PROC_EXIT:
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close: Set ts = Nothing
    Set fso = Nothing
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "EmailAgent", "ShowLearnedRepliesSummary", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub
