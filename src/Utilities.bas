'===============================================================================
' Utilities.bas - Helper Functions for Email Agent v3.0
'===============================================================================
' This module contains utility functions used across the email filter:
'   - String matching (ContainsAny, StartsWithAny, MatchesAny)
'   - JSON encoding/parsing
'   - Folder management
'   - Logging
'   - Email address extraction
'   - Learned senders/subjects cache (in-memory Dictionary + file I/O)
'   - Settings INI reader/writer
'   - Learned rule deletion
'   - Call stack tracking and centralized error handling
'   - Multi-provider LLM caller (CallLLM)
'   - Reply pair I/O (learned_replies.txt)
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' CALL STACK TRACKING
'-------------------------------------------------------------------------------
' Simulates a call stack using a fixed-size array. PushCallStack is called at
' the start of each instrumented procedure; PopCallStack in its PROC_EXIT label.
' GetCallStack returns a " -> " separated string for use in error log entries.

Private callStack(0 To 19) As String  ' CALL_STACK_MAX_DEPTH = 20
Private callStackDepth As Integer

' Learned senders cache (Self-Improving Filter)
Private learnedSendersCache As Object  ' Scripting.Dictionary (email -> "KEEP"|"DELETE")
Private learnedSendersCacheLoaded As Boolean

' Learned subjects cache (Self-Improving Filter - Subject Rules)
Private learnedSubjectsCache As Object  ' Scripting.Dictionary (subject -> "DELETE")
Private learnedSubjectsCacheLoaded As Boolean

' Web UI command poller state
Private pollerRunningFlag As Boolean

' Windows API timer for command poller (Outlook has no Application.OnTime)
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

Public Sub PushCallStack(ByVal procName As String)
    If callStackDepth < CALL_STACK_MAX_DEPTH Then
        callStack(callStackDepth) = procName
        callStackDepth = callStackDepth + 1
    End If
End Sub

Public Sub PopCallStack()
    If callStackDepth > 0 Then
        callStackDepth = callStackDepth - 1
        callStack(callStackDepth) = ""
    End If
End Sub

Public Function GetCallStack() As String
    Dim i As Integer
    Dim parts() As String
    If callStackDepth = 0 Then
        GetCallStack = "(empty)"
        Exit Function
    End If
    ReDim parts(0 To callStackDepth - 1)
    For i = 0 To callStackDepth - 1
        parts(i) = callStack(i)
    Next i
    GetCallStack = Join(parts, " -> ")
End Function

'-------------------------------------------------------------------------------
' CENTRALIZED ERROR HANDLING
'-------------------------------------------------------------------------------
' LogError writes a structured entry to error.log and optionally shows a MsgBox.
' WriteToLogFile handles the file I/O for all log-to-file operations.

Public Sub LogError(ByVal moduleName As String, ByVal procName As String, _
                    ByVal errNum As Long, ByVal errDesc As String, _
                    Optional ByVal context As String = "")
    Dim entry As String
    Dim stackStr As String

    stackStr = GetCallStack()
    entry = Format(Now, "yyyy-mm-dd hh:nn:ss") & "|" & _
            moduleName & "." & procName & "|" & _
            errNum & "|" & _
            errDesc

    If Len(context) > 0 Then entry = entry & "|" & context
    entry = entry & "|Stack: " & stackStr

    ' Write to log file (best effort - do not raise new error)
    On Error Resume Next
    WriteToLogFile entry
    On Error GoTo 0

    ' Write to Immediate Window
    Debug.Print entry

    ' Optionally show MsgBox in debug mode
    If RuntimeDebugMode Then
        MsgBox "Error in " & moduleName & "." & procName & ":" & vbCrLf & _
               "  " & errDesc & " (Error " & errNum & ")" & vbCrLf & vbCrLf & _
               "Stack: " & stackStr, _
               vbExclamation, "Email Agent Error"
    End If
End Sub

' NOTE: No error handler here — this is called from LogError/LogMessage,
' so adding LogError would cause infinite recursion. Callers use On Error Resume Next.
Public Sub WriteToLogFile(ByVal message As String)
    Dim fso As Object
    Dim ts As Object
    Dim logPath As String

    ' Build path if not yet set (RuntimeErrorLogFile set in LoadAllSettings)
    If Len(RuntimeErrorLogFile) = 0 Then
        logPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & ERROR_LOG_FILE_NAME
    Else
        logPath = RuntimeErrorLogFile
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logPath, 8, True)  ' 8 = ForAppending, True = create if missing
    ts.WriteLine message
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' MULTI-PROVIDER LLM CALLER
'-------------------------------------------------------------------------------
' CallLLM routes to the configured provider (local | azure | claude | openai).
' Returns the raw response content string, or "" on error/no key.
' All callers should use CallLLM instead of CallAzureOpenAICustom (kept for
' backwards-compatibility but now delegates here).

Public Function CallLLM(ByVal userPrompt As String, ByVal systemPrompt As String, _
                        ByVal maxTokens As Integer, _
                        Optional ByVal temperature As Double = -1) As String
    On Error GoTo LLMError
    PushCallStack "Utilities.CallLLM"

    ' Use the per-call temperature if given; otherwise use the classify default
    Dim temp As Double
    If temperature < 0 Then
        temp = RuntimeLLMTemperature
    Else
        temp = temperature
    End If

    Dim provider As String
    provider = LCase(Trim(RuntimeLLMProvider))
    If Len(provider) = 0 Then provider = "azure"

    Select Case provider
        Case "local"
            CallLLM = CallLLMLocal(userPrompt, systemPrompt, maxTokens, temp)
        Case "claude"
            CallLLM = CallLLMClaude(userPrompt, systemPrompt, maxTokens, temp)
        Case "openai"
            CallLLM = CallLLMOpenAI(userPrompt, systemPrompt, maxTokens, temp)
        Case Else  ' "azure" and anything unrecognised
            CallLLM = CallLLMAzure(userPrompt, systemPrompt, maxTokens, temp)
    End Select

PROC_EXIT:
    PopCallStack
    Exit Function
LLMError:
    LogError "Utilities", "CallLLM", Err.Number, Err.Description
    CallLLM = ""
    Resume PROC_EXIT
End Function

' Build OpenAI-compatible chat/completions JSON body (shared by azure & local)
Private Function BuildOpenAIBody(ByVal userPrompt As String, ByVal systemPrompt As String, _
                                  ByVal maxTokens As Integer, ByVal temp As Double) As String
    ' Force "." decimal separator regardless of Windows locale
    Dim tempStr As String
    tempStr = Format(temp, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    BuildOpenAIBody = "{" & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & tempStr & "," & _
        """stream"":false" & _
    "}"
End Function

' Call local OpenAI-compatible endpoint (Ollama, LM Studio, Inferencer, etc.)
Private Function CallLLMLocal(ByVal userPrompt As String, ByVal systemPrompt As String, _
                               ByVal maxTokens As Integer, ByVal temp As Double) As String
    Dim http As Object
    Dim body As String

    ' Inject model name into body for local servers
    Dim tempStr As String
    tempStr = Format(temp, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    body = "{" & _
        """model"":""" & RuntimeLocalModel & """," & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & tempStr & "," & _
        """stream"":false" & _
    "}"

    LogMessage "DEBUG", "Calling local LLM at " & RuntimeLocalEndpoint & " (model: " & RuntimeLocalModel & ")"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", RuntimeLocalEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    ' Local servers typically accept any key or no key; provide a dummy to avoid rejections
    http.setRequestHeader "Authorization", "Bearer local"
    http.send body

    If http.Status = 200 Then
        Dim result As String
        result = ParseJSONContent(http.responseText)
        LogMessage "DEBUG", "Local LLM response: " & Left(result, 100)
        WriteLLMDebugLog "CallLLMLocal", RuntimeLocalEndpoint, RuntimeLocalModel, _
                         maxTokens, temp, body, http.Status, http.responseText, result
        CallLLMLocal = result
    Else
        LogMessage "ERROR", "Local LLM error: " & http.Status & " - " & http.statusText
        WriteLLMDebugLog "CallLLMLocal", RuntimeLocalEndpoint, RuntimeLocalModel, _
                         maxTokens, temp, body, http.Status, http.responseText, ""
        CallLLMLocal = ""
    End If

    Set http = Nothing
End Function

' Call external OpenAI-compatible endpoint (OpenRouter, Groq, Together AI, OpenAI, etc.)
' Unlike CallLLMLocal, this requires a real API key via GetAPIKey().
Private Function CallLLMOpenAI(ByVal userPrompt As String, ByVal systemPrompt As String, _
                                ByVal maxTokens As Integer, ByVal temp As Double) As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.CallLLMOpenAI"

    Dim http As Object
    Dim apiKey As String

    apiKey = GetAPIKey()
    If Len(apiKey) = 0 Then
        LogMessage "WARN", "CallLLMOpenAI: no API key"
        CallLLMOpenAI = ""
        GoTo PROC_EXIT
    End If

    ' Build body with model name (OpenAI-compatible format)
    Dim tempStr As String
    tempStr = Format(temp, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    Dim body As String
    body = "{" & _
        """model"":""" & RuntimeOpenAIModel & """," & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & tempStr & "," & _
        """stream"":false" & _
    "}"

    LogMessage "DEBUG", "Calling OpenAI-compatible API at " & RuntimeOpenAIEndpoint & " (model: " & RuntimeOpenAIModel & ")"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", RuntimeOpenAIEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body

    If http.Status = 200 Then
        Dim result As String
        result = ParseJSONContent(http.responseText)
        LogMessage "DEBUG", "OpenAI-compatible LLM response: " & Left(result, 100)
        WriteLLMDebugLog "CallLLMOpenAI", RuntimeOpenAIEndpoint, RuntimeOpenAIModel, _
                         maxTokens, temp, body, http.Status, http.responseText, result
        CallLLMOpenAI = result
    Else
        LogMessage "ERROR", "OpenAI-compatible API error: " & http.Status & " - " & http.statusText
        WriteLLMDebugLog "CallLLMOpenAI", RuntimeOpenAIEndpoint, RuntimeOpenAIModel, _
                         maxTokens, temp, body, http.Status, http.responseText, ""
        CallLLMOpenAI = ""
    End If

    Set http = Nothing

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "CallLLMOpenAI", Err.Number, Err.Description
    CallLLMOpenAI = ""
    Resume PROC_EXIT
End Function

' Call Azure OpenAI endpoint
Private Function CallLLMAzure(ByVal userPrompt As String, ByVal systemPrompt As String, _
                               ByVal maxTokens As Integer, ByVal temp As Double) As String
    Dim http As Object
    Dim apiKey As String

    apiKey = GetAPIKey()
    If Len(apiKey) = 0 Then
        LogMessage "WARN", "CallLLMAzure: no API key"
        CallLLMAzure = ""
        Exit Function
    End If

    Dim body As String
    body = BuildOpenAIBody(userPrompt, systemPrompt, maxTokens, temp)

    LogMessage "DEBUG", "Calling Azure OpenAI..."

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", RuntimeLLMEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "api-key", apiKey
    http.send body

    If http.Status = 200 Then
        Dim result As String
        result = ParseJSONContent(http.responseText)
        LogMessage "DEBUG", "Azure LLM response: " & Left(result, 100)
        WriteLLMDebugLog "CallLLMAzure", RuntimeLLMEndpoint, "azure", _
                         maxTokens, temp, body, http.Status, http.responseText, result
        CallLLMAzure = result
    Else
        LogMessage "ERROR", "Azure OpenAI error: " & http.Status & " - " & http.statusText
        WriteLLMDebugLog "CallLLMAzure", RuntimeLLMEndpoint, "azure", _
                         maxTokens, temp, body, http.Status, http.responseText, ""
        CallLLMAzure = ""
    End If

    Set http = Nothing
End Function

' Call Anthropic Claude API (/v1/messages - different schema from OpenAI)
Private Function CallLLMClaude(ByVal userPrompt As String, ByVal systemPrompt As String, _
                                ByVal maxTokens As Integer, ByVal temp As Double) As String
    Dim http As Object
    Dim apiKey As String

    apiKey = GetAPIKey()
    If Len(apiKey) = 0 Then
        LogMessage "WARN", "CallLLMClaude: no API key"
        CallLLMClaude = ""
        Exit Function
    End If

    ' Force "." decimal separator regardless of Windows locale
    Dim tempStr As String
    tempStr = Format(temp, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    ' Claude API uses a different JSON schema: system is top-level, model in body
    Dim body As String
    body = "{" & _
        """model"":""" & RuntimeClaudeModel & """," & _
        """system"":""" & EscapeJSON(systemPrompt) & """," & _
        """messages"":[" & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & tempStr & "," & _
        """stream"":false" & _
    "}"

    LogMessage "DEBUG", "Calling Claude API (model: " & RuntimeClaudeModel & ")..."

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", RuntimeClaudeEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "x-api-key", apiKey
    http.setRequestHeader "anthropic-version", "2023-06-01"
    http.send body

    If http.Status = 200 Then
        ' Claude response format: {"content":[{"type":"text","text":"..."}],...}
        ' We reuse ParseJSONContent but must find "text" field inside the content array
        Dim rawResponse As String
        rawResponse = http.responseText

        ' Try to extract the first text block from Claude's response
        Dim textStart As Long
        textStart = InStr(1, rawResponse, """text"":", vbTextCompare)
        If textStart > 0 Then
            textStart = InStr(textStart, rawResponse, ":""") + 2
            Dim textEnd As Long
            textEnd = textStart
            Do
                textEnd = InStr(textEnd + 1, rawResponse, """")
                If textEnd = 0 Then Exit Do
                If Mid(rawResponse, textEnd - 1, 1) <> "\" Then Exit Do
            Loop
            Dim claudeResult As String
            If textEnd > textStart Then
                claudeResult = Mid(rawResponse, textStart, textEnd - textStart)
                claudeResult = Replace(claudeResult, "\n", vbCrLf)
                claudeResult = Replace(claudeResult, "\t", vbTab)
                claudeResult = Replace(claudeResult, "\""", """")
                claudeResult = Replace(claudeResult, "\\", "\")
            End If
        End If

        LogMessage "DEBUG", "Claude LLM response: " & Left(claudeResult, 100)
        WriteLLMDebugLog "CallLLMClaude", RuntimeClaudeEndpoint, RuntimeClaudeModel, _
                         maxTokens, temp, body, http.Status, http.responseText, claudeResult
        CallLLMClaude = claudeResult
    Else
        LogMessage "ERROR", "Claude API error: " & http.Status & " - " & http.statusText
        WriteLLMDebugLog "CallLLMClaude", RuntimeClaudeEndpoint, RuntimeClaudeModel, _
                         maxTokens, temp, body, http.Status, http.responseText, ""
        CallLLMClaude = ""
    End If

    Set http = Nothing
End Function

'-------------------------------------------------------------------------------
' LEARNED SENDERS CACHE (Self-Improving Filter)
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' LEARNED SUBJECTS CACHE (Self-Improving Filter - Subject Rules)
'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------
' STRING MATCHING FUNCTIONS
'-------------------------------------------------------------------------------

' Check if a string contains any item from a comma-separated list
' Returns True if any pattern is found in the text
Public Function ContainsAny(ByVal text As String, ByVal patterns As String) As Boolean
    Dim patternArray() As String
    Dim pattern As Variant
    Dim lowerText As String

    ContainsAny = False

    If Len(text) = 0 Or Len(patterns) = 0 Then Exit Function

    lowerText = LCase(text)
    patternArray = Split(patterns, ",")

    For Each pattern In patternArray
        If Len(Trim(pattern)) > 0 Then
            If InStr(1, lowerText, LCase(Trim(pattern)), vbTextCompare) > 0 Then
                ContainsAny = True
                Exit Function
            End If
        End If
    Next pattern
End Function

' Check if text starts with any pattern from a comma-separated list
Public Function StartsWithAny(ByVal text As String, ByVal patterns As String) As Boolean
    Dim patternArray() As String
    Dim pattern As Variant
    Dim trimmedText As String

    StartsWithAny = False

    If Len(text) = 0 Or Len(patterns) = 0 Then Exit Function

    ' Trim leading whitespace and newlines
    trimmedText = LTrim(text)

    patternArray = Split(patterns, ",")

    For Each pattern In patternArray
        If Len(Trim(pattern)) > 0 Then
            If LCase(Left(trimmedText, Len(Trim(pattern)))) = LCase(Trim(pattern)) Then
                StartsWithAny = True
                Exit Function
            End If
        End If
    Next pattern
End Function

' Check if text exactly matches any item from a comma-separated list
Public Function MatchesAny(ByVal text As String, ByVal patterns As String) As Boolean
    Dim patternArray() As String
    Dim pattern As Variant

    MatchesAny = False

    If Len(text) = 0 Or Len(patterns) = 0 Then Exit Function

    patternArray = Split(patterns, ",")

    For Each pattern In patternArray
        If Len(Trim(pattern)) > 0 Then
            If LCase(Trim(text)) = LCase(Trim(pattern)) Then
                MatchesAny = True
                Exit Function
            End If
        End If
    Next pattern
End Function

'-------------------------------------------------------------------------------
' JSON FUNCTIONS
'-------------------------------------------------------------------------------

' Escape a string for JSON encoding
Public Function EscapeJSON(ByVal text As String) As String
    Dim result As String

    result = text

    ' Escape backslashes first
    result = Replace(result, "\", "\\")

    ' Escape double quotes
    result = Replace(result, """", "\""")

    ' Escape control characters
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")

    EscapeJSON = result
End Function

' Parse a simple JSON response to extract the "content" field
' This is a basic parser - for complex JSON, consider a library
Public Function ParseJSONContent(ByVal jsonText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim content As String

    ' Look for "content": " pattern
    startPos = InStr(1, jsonText, """content"":", vbTextCompare)

    If startPos = 0 Then
        ParseJSONContent = ""
        Exit Function
    End If

    ' Find the opening quote of the value
    startPos = InStr(startPos, jsonText, ":""") + 2

    If startPos < 3 Then
        ParseJSONContent = ""
        Exit Function
    End If

    ' Find the closing quote (handle escaped quotes)
    endPos = startPos
    Do
        endPos = InStr(endPos + 1, jsonText, """")
        If endPos = 0 Then Exit Do
        ' Check if this quote is escaped
        If Mid(jsonText, endPos - 1, 1) <> "\" Then Exit Do
    Loop

    If endPos > startPos Then
        content = Mid(jsonText, startPos, endPos - startPos)
        ' Unescape the content
        content = Replace(content, "\n", vbCrLf)
        content = Replace(content, "\t", vbTab)
        content = Replace(content, "\""", """")
        content = Replace(content, "\\", "\")
        ParseJSONContent = content
    Else
        ParseJSONContent = ""
    End If
End Function

'-------------------------------------------------------------------------------
' FOLDER MANAGEMENT
'-------------------------------------------------------------------------------

' Get or create a folder under the Inbox
Public Function GetOrCreateFolder(ByVal folderName As String) As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim targetFolder As Outlook.Folder

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    On Error Resume Next
    Set targetFolder = inbox.Folders(folderName)
    On Error GoTo 0

    If targetFolder Is Nothing Then
        ' Create the folder
        Set targetFolder = inbox.Folders.Add(folderName)
        LogMessage "INFO", "Created folder: " & folderName
    End If

    Set GetOrCreateFolder = targetFolder
End Function

' Get a folder by path (e.g., "Inbox/Subfolder/Target")
Public Function GetFolderByPath(ByVal folderPath As String) As Outlook.Folder
    Dim ns As Outlook.NameSpace
    Dim pathParts() As String
    Dim currentFolder As Outlook.Folder
    Dim i As Integer

    Set ns = Application.GetNamespace("MAPI")

    pathParts = Split(folderPath, "/")

    ' Start from the first folder in the path
    If LCase(pathParts(0)) = "inbox" Then
        Set currentFolder = ns.GetDefaultFolder(olFolderInbox)
    Else
        ' Try to find it as a root folder
        On Error Resume Next
        Set currentFolder = ns.Folders(pathParts(0))
        On Error GoTo 0
    End If

    If currentFolder Is Nothing Then
        Set GetFolderByPath = Nothing
        Exit Function
    End If

    ' Navigate through subfolders
    For i = 1 To UBound(pathParts)
        On Error Resume Next
        Set currentFolder = currentFolder.Folders(pathParts(i))
        On Error GoTo 0

        If currentFolder Is Nothing Then
            Set GetFolderByPath = Nothing
            Exit Function
        End If
    Next i

    Set GetFolderByPath = currentFolder
End Function

'-------------------------------------------------------------------------------
' LOGGING FUNCTIONS
'-------------------------------------------------------------------------------

' Log a message to the Immediate Window
Public Sub LogMessage(ByVal level As String, ByVal message As String)
    If Not RuntimeSettingsLoaded Then
        ' Before settings are loaded, use default logging behavior
        If Not DEFAULT_ENABLE_LOGGING Then Exit Sub
        Dim defLevel As Integer
        defLevel = GetLogLevelValue(DEFAULT_LOG_LEVEL)
        If GetLogLevelValue(level) < defLevel Then Exit Sub
    Else
        If Not RuntimeEnableLogging Then Exit Sub
        If GetLogLevelValue(level) < GetLogLevelValue(RuntimeLogLevel) Then Exit Sub
    End If

    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & message
End Sub

' Get numeric value for log level comparison
Private Function GetLogLevelValue(ByVal level As String) As Integer
    Select Case UCase(level)
        Case "DEBUG": GetLogLevelValue = 1
        Case "INFO": GetLogLevelValue = 2
        Case "WARN": GetLogLevelValue = 3
        Case "ERROR": GetLogLevelValue = 4
        Case Else: GetLogLevelValue = 2
    End Select
End Function

' Log an email action (requires valid mail object)
Public Sub LogAction(ByVal mail As Outlook.MailItem, ByVal action As String)
    Dim logEntry As String

    logEntry = action & " | From: " & Left(mail.SenderName, 25) & _
               " | Subject: " & Left(mail.subject, 40)

    LogMessage "INFO", logEntry
End Sub

' Log an email action using pre-captured values (use when mail object may become invalid)
Public Sub LogActionDirect(ByVal senderName As String, ByVal subject As String, ByVal action As String)
    Dim logEntry As String

    logEntry = action & " | From: " & Left(senderName, 25) & _
               " | Subject: " & Left(subject, 40)

    LogMessage "INFO", logEntry
End Sub

'-------------------------------------------------------------------------------
' API KEY MANAGEMENT
'-------------------------------------------------------------------------------

' Get the Azure OpenAI API key
Public Function GetAPIKey() As String
    Select Case UCase(RuntimeAPIKeyMethod)
        Case "ENV"
            GetAPIKey = Environ(RuntimeAPIKeyEnvVar)
        Case "HARDCODED"
            GetAPIKey = RuntimeAPIKeyHardcoded
        Case Else
            GetAPIKey = ""
    End Select
End Function

'-------------------------------------------------------------------------------
' EMAIL HELPER FUNCTIONS
'-------------------------------------------------------------------------------

' Get the sender's email address (handles Exchange addresses)
Public Function GetSenderEmail(ByVal mail As Outlook.MailItem) As String
    Dim senderEmail As String

    On Error Resume Next

    ' First try the direct property
    senderEmail = mail.SenderEmailAddress

    ' If it's an Exchange address, try to get the SMTP address
    If InStr(1, senderEmail, "/O=", vbTextCompare) > 0 Then
        ' Exchange internal address format
        If Not mail.Sender Is Nothing Then
            senderEmail = mail.Sender.GetExchangeUser.PrimarySmtpAddress
        End If
    End If

    On Error GoTo 0

    GetSenderEmail = LCase(senderEmail)
End Function

' Get the domain from an email address
Public Function GetDomain(ByVal email As String) As String
    Dim atPos As Long

    atPos = InStr(1, email, "@")

    If atPos > 0 Then
        GetDomain = LCase(Mid(email, atPos + 1))
    Else
        GetDomain = ""
    End If
End Function

' Truncate text to a maximum length with ellipsis
Public Function Truncate(ByVal text As String, ByVal maxLength As Integer) As String
    If Len(text) <= maxLength Then
        Truncate = text
    Else
        Truncate = Left(text, maxLength - 3) & "..."
    End If
End Function

'-------------------------------------------------------------------------------
' STATISTICS DICTIONARY HELPER
'-------------------------------------------------------------------------------

' Initialize a statistics dictionary
Public Function CreateStatsDict() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1  ' Case-insensitive (project convention)

    dict("DELETE") = 0
    dict("MOVE_II") = 0
    dict("KEEP") = 0
    dict("REVIEW") = 0
    dict("ERROR") = 0

    Set CreateStatsDict = dict
End Function

' Increment a statistics counter
Public Sub IncrementStat(ByVal stats As Object, ByVal key As String)
    If stats.Exists(key) Then
        stats(key) = stats(key) + 1
    Else
        stats(key) = 1
    End If
End Sub

' Format statistics for display
Public Function FormatStats(ByVal stats As Object) As String
    Dim result As String

    result = "Filtering Complete!" & vbCrLf & vbCrLf
    result = result & "Deleted: " & stats("DELETE") & vbCrLf
    result = result & "Moved to " & RuntimeFolderProtected & ": " & stats("MOVE_II") & vbCrLf
    result = result & "Moved to " & RuntimeFolderReview & ": " & stats("REVIEW") & vbCrLf
    result = result & "Kept in Inbox: " & stats("KEEP") & vbCrLf

    If stats("ERROR") > 0 Then
        result = result & "Errors: " & stats("ERROR") & vbCrLf
    End If

    If stats.Exists("LEARNED_KEEP") And stats("LEARNED_KEEP") > 0 Then
        result = result & "Learned Keep: " & stats("LEARNED_KEEP") & vbCrLf
    End If
    If stats.Exists("LEARNED_DELETE") And stats("LEARNED_DELETE") > 0 Then
        result = result & "Learned Delete: " & stats("LEARNED_DELETE") & vbCrLf
    End If
    If stats.Exists("LEARNED_SUBJECT_DELETE") And stats("LEARNED_SUBJECT_DELETE") > 0 Then
        result = result & "Learned Subject Delete: " & stats("LEARNED_SUBJECT_DELETE") & vbCrLf
    End If

    FormatStats = result
End Function

'-------------------------------------------------------------------------------
' SETTINGS INI FILE - PATH AND INITIALIZATION
'-------------------------------------------------------------------------------

' Build the full path to the settings.ini file, creating the directory if needed
Public Function GetSettingsFilePath() As String
    Dim folderPath As String
    Dim fso As Object

    folderPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER

    ' Create directory if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        LogMessage "INFO", "Created settings folder: " & folderPath
    End If
    Set fso = Nothing

    GetSettingsFilePath = folderPath & "\" & SETTINGS_FILE_NAME
End Function

' Build the full path to the LLM debug log file, creating the directory if needed
Public Function GetLLMDebugLogPath() As String
    Dim folderPath As String
    Dim fso As Object

    folderPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing

    GetLLMDebugLogPath = folderPath & "\llm_debug.log"
End Function

' Write verbose LLM request/response to llm_debug.log (only when LogLevel=DEBUG).
' Best-effort: failures here never break LLM calls.
Private Sub WriteLLMDebugLog(ByVal provider As String, ByVal endpoint As String, _
                              ByVal model As String, ByVal maxTokens As Integer, _
                              ByVal temperature As Double, ByVal requestBody As String, _
                              ByVal httpStatus As Long, ByVal responseBody As String, _
                              ByVal parsedContent As String)
    On Error Resume Next

    ' Only write when LogLevel is DEBUG
    If UCase(RuntimeLogLevel) <> "DEBUG" Then Exit Sub

    Dim fso As Object
    Dim ts As Object
    Dim logPath As String
    Dim tempStr As String

    logPath = GetLLMDebugLogPath()
    tempStr = Format(temperature, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logPath, 8, True)  ' 8 = ForAppending, True = create if missing
    ts.WriteLine String(80, "=")
    ts.WriteLine Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & provider & " | " & model & " | " & endpoint
    ts.WriteLine "--- REQUEST (max_tokens=" & maxTokens & ", temperature=" & tempStr & ") ---"
    ts.WriteLine requestBody
    ts.WriteLine "--- RESPONSE (HTTP " & httpStatus & ", " & Len(responseBody) & " chars) ---"
    ts.WriteLine responseBody
    ts.WriteLine "--- PARSED (" & Len(parsedContent) & " chars) ---"
    ts.WriteLine parsedContent
    ts.WriteLine String(80, "=")
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

' Load all settings from the INI file into Runtime* variables.
' If the INI file doesn't exist, creates one with defaults.
' MUST be called at the very start of Application_Startup.
Public Sub LoadAllSettings()
    Dim settingsPath As String

    settingsPath = GetSettingsFilePath()

    ' Create default settings file if it doesn't exist
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(settingsPath) Then
        Set fso = Nothing
        CreateDefaultSettingsFile
    Else
        Set fso = Nothing
    End If

    ' --- General ---
    RuntimeEnableLogging = ReadINIBool("General", "EnableLogging", DEFAULT_ENABLE_LOGGING)
    RuntimeLogLevel = ReadINISetting("General", "LogLevel", DEFAULT_LOG_LEVEL)
    RuntimeEnableSelfImproving = ReadINIBool("General", "EnableSelfImproving", DEFAULT_ENABLE_SELF_IMPROVING)
    RuntimeProgressInterval = ReadINIInt("General", "ProgressInterval", DEFAULT_PROGRESS_INTERVAL)
    RuntimeDryRunLimit = ReadINIInt("General", "DryRunLimit", DEFAULT_DRY_RUN_LIMIT)
    RuntimeLLMBatchSize = ReadINIInt("General", "LLMBatchSize", DEFAULT_LLM_BATCH_SIZE)

    ' --- Folders ---
    RuntimeFolderProtected = ReadINISetting("Folders", "Protected", DEFAULT_FOLDER_PROTECTED)
    RuntimeFolderReview = ReadINISetting("Folders", "Review", DEFAULT_FOLDER_REVIEW)
    RuntimeFolderLearnKeep = ReadINISetting("Folders", "LearnKeep", DEFAULT_FOLDER_LEARN_KEEP)
    RuntimeFolderLearnDelete = ReadINISetting("Folders", "LearnDelete", DEFAULT_FOLDER_LEARN_DELETE)
    RuntimeFolderLearnSubject = ReadINISetting("Folders", "LearnSubject", DEFAULT_FOLDER_LEARN_SUBJECT_DELETE)

    ' --- Patterns ---
    RuntimeProtectedDomains = ReadINISetting("Patterns", "ProtectedDomains", DEFAULT_PROTECTED_DOMAINS)
    RuntimeNamePatterns = ReadINISetting("Patterns", "NamePatterns", DEFAULT_NAME_PATTERNS)
    RuntimeGreetingPatterns = ReadINISetting("Patterns", "GreetingPatterns", DEFAULT_GREETING_PATTERNS)
    RuntimePolyUTags = ReadINISetting("Patterns", "PolyUTags", DEFAULT_POLYU_TAGS)
    RuntimeVIPKeywords = ReadINISetting("Patterns", "VIPSubjectKeywords", DEFAULT_VIP_SUBJECT_KEYWORDS)
    RuntimeDeleteSenderPatterns = ReadINISetting("Patterns", "DeleteSenderPatterns", DEFAULT_DELETE_SENDER_PATTERNS)
    RuntimeDeleteKnownSenders = ReadINISetting("Patterns", "DeleteKnownSenders", DEFAULT_DELETE_KNOWN_SENDERS)
    RuntimeDeleteSubjectPatterns = ReadINISetting("Patterns", "DeleteSubjectPatterns", DEFAULT_DELETE_SUBJECT_PATTERNS)

    ' --- LLM ---
    RuntimeUseLLM = ReadINIBool("LLM", "UseLLMAPI", DEFAULT_USE_LLM_API)
    RuntimeLLMProvider = ReadINISetting("LLM", "Provider", DEFAULT_LLM_PROVIDER)
    RuntimeLLMEndpoint = ReadINISetting("LLM", "AzureEndpoint", DEFAULT_AZURE_OPENAI_ENDPOINT)
    RuntimeLocalEndpoint = ReadINISetting("LLM", "LocalEndpoint", DEFAULT_LOCAL_ENDPOINT)
    RuntimeLocalModel = ReadINISetting("LLM", "LocalModel", DEFAULT_LOCAL_MODEL)
    RuntimeClaudeEndpoint = ReadINISetting("LLM", "ClaudeEndpoint", DEFAULT_CLAUDE_ENDPOINT)
    RuntimeClaudeModel = ReadINISetting("LLM", "ClaudeModel", DEFAULT_CLAUDE_MODEL)
    RuntimeOpenAIEndpoint = ReadINISetting("LLM", "OpenAIEndpoint", DEFAULT_OPENAI_COMPAT_ENDPOINT)
    RuntimeOpenAIModel = ReadINISetting("LLM", "OpenAIModel", DEFAULT_OPENAI_COMPAT_MODEL)
    RuntimeAPIKeyMethod = ReadINISetting("LLM", "APIKeyMethod", DEFAULT_API_KEY_METHOD)
    RuntimeAPIKeyEnvVar = ReadINISetting("LLM", "APIKeyEnvVar", DEFAULT_API_KEY_ENV_VAR)
    RuntimeAPIKeyHardcoded = ReadINISetting("LLM", "APIKeyHardcoded", DEFAULT_API_KEY_HARDCODED)
    RuntimeClassifyMaxTokens = ReadINIInt("LLM", "ClassifyMaxTokens", DEFAULT_CLASSIFY_MAX_TOKENS)
    RuntimeSummarizeMaxTokens = ReadINIInt("LLM", "SummarizeMaxTokens", DEFAULT_SUMMARIZE_MAX_TOKENS)
    RuntimeReplyMaxTokens = ReadINIInt("LLM", "ReplyMaxTokens", DEFAULT_REPLY_MAX_TOKENS)
    RuntimeLLMTemperature = ReadINIDouble("LLM", "Temperature", DEFAULT_LLM_TEMPERATURE)
    RuntimeReplyTemperature = ReadINIDouble("LLM", "ReplyTemperature", DEFAULT_REPLY_TEMPERATURE)
    RuntimeLLMSystemPrompt = ReadINISetting("LLM", "SystemPrompt", DEFAULT_LLM_SYSTEM_PROMPT)

    ' --- Agent ---
    RuntimeEnableAutoReply = ReadINIBool("Agent", "EnableAutoReply", DEFAULT_ENABLE_AUTO_REPLY)
    RuntimeAutoReplyOnArrival = ReadINIBool("Agent", "AutoReplyOnArrival", DEFAULT_AUTO_REPLY_ON_ARRIVAL)
    RuntimeFolderLearnReply = ReadINISetting("Agent", "LearnReplyFolder", DEFAULT_FOLDER_LEARN_REPLY)
    RuntimeMaxReplyExamples = ReadINIInt("Agent", "MaxReplyExamples", DEFAULT_MAX_REPLY_EXAMPLES)
    RuntimeReplyPersona = ReadINISetting("Agent", "ReplyPersona", DEFAULT_REPLY_PERSONA)
    RuntimeScanSentItems = ReadINIBool("Agent", "ScanSentItems", DEFAULT_SCAN_SENT_ITEMS)
    RuntimeScanSentDays = ReadINIInt("Agent", "ScanSentDays", DEFAULT_SCAN_SENT_DAYS)
    RuntimeAutoReplyForSenders = ReadINISetting("Agent", "AutoReplyForSenders", DEFAULT_AUTO_REPLY_FOR_SENDERS)

    ' --- Debug / Error handling ---
    RuntimeDebugMode = ReadINIBool("General", "DebugMode", DEFAULT_DEBUG_MODE)
    RuntimeErrorLogFile = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & ERROR_LOG_FILE_NAME

    RuntimeSettingsLoaded = True
    LogMessage "INFO", "Settings loaded from: " & settingsPath
End Sub

' Write a default settings.ini from DEFAULT_* constants
Public Sub CreateDefaultSettingsFile()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object

    filePath = GetSettingsFilePath()

    On Error GoTo FileError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True)  ' True = overwrite

    ts.WriteLine "[General]"
    ts.WriteLine "Version=" & FILTER_VERSION
    ts.WriteLine "EnableLogging=" & IIf(DEFAULT_ENABLE_LOGGING, "True", "False")
    ts.WriteLine "LogLevel=" & DEFAULT_LOG_LEVEL
    ts.WriteLine "EnableSelfImproving=" & IIf(DEFAULT_ENABLE_SELF_IMPROVING, "True", "False")
    ts.WriteLine "DebugMode=" & IIf(DEFAULT_DEBUG_MODE, "True", "False")
    ts.WriteLine "ProgressInterval=" & DEFAULT_PROGRESS_INTERVAL
    ts.WriteLine "DryRunLimit=" & DEFAULT_DRY_RUN_LIMIT
    ts.WriteLine "LLMBatchSize=" & DEFAULT_LLM_BATCH_SIZE
    ts.WriteLine ""
    ts.WriteLine "[Folders]"
    ts.WriteLine "Protected=" & DEFAULT_FOLDER_PROTECTED
    ts.WriteLine "Review=" & DEFAULT_FOLDER_REVIEW
    ts.WriteLine "LearnKeep=" & DEFAULT_FOLDER_LEARN_KEEP
    ts.WriteLine "LearnDelete=" & DEFAULT_FOLDER_LEARN_DELETE
    ts.WriteLine "LearnSubject=" & DEFAULT_FOLDER_LEARN_SUBJECT_DELETE
    ts.WriteLine ""
    ts.WriteLine "[Patterns]"
    ts.WriteLine "ProtectedDomains=" & DEFAULT_PROTECTED_DOMAINS
    ts.WriteLine "NamePatterns=" & DEFAULT_NAME_PATTERNS
    ts.WriteLine "GreetingPatterns=" & DEFAULT_GREETING_PATTERNS
    ts.WriteLine "PolyUTags=" & DEFAULT_POLYU_TAGS
    ts.WriteLine "VIPSubjectKeywords=" & DEFAULT_VIP_SUBJECT_KEYWORDS
    ts.WriteLine "DeleteSenderPatterns=" & DEFAULT_DELETE_SENDER_PATTERNS
    ts.WriteLine "DeleteKnownSenders=" & DEFAULT_DELETE_KNOWN_SENDERS
    ts.WriteLine "DeleteSubjectPatterns=" & DEFAULT_DELETE_SUBJECT_PATTERNS
    ts.WriteLine ""
    ts.WriteLine "[LLM]"
    ts.WriteLine "UseLLMAPI=" & IIf(DEFAULT_USE_LLM_API, "True", "False")
    ts.WriteLine "; Provider: local | azure | claude | openai"
    ts.WriteLine "Provider=" & DEFAULT_LLM_PROVIDER
    ts.WriteLine "AzureEndpoint=" & DEFAULT_AZURE_OPENAI_ENDPOINT
    ts.WriteLine "LocalEndpoint=" & DEFAULT_LOCAL_ENDPOINT
    ts.WriteLine "LocalModel=" & DEFAULT_LOCAL_MODEL
    ts.WriteLine "ClaudeEndpoint=" & DEFAULT_CLAUDE_ENDPOINT
    ts.WriteLine "ClaudeModel=" & DEFAULT_CLAUDE_MODEL
    ts.WriteLine "OpenAIEndpoint=" & DEFAULT_OPENAI_COMPAT_ENDPOINT
    ts.WriteLine "OpenAIModel=" & DEFAULT_OPENAI_COMPAT_MODEL
    ts.WriteLine "APIKeyMethod=" & DEFAULT_API_KEY_METHOD
    ts.WriteLine "APIKeyEnvVar=" & DEFAULT_API_KEY_ENV_VAR
    ts.WriteLine "APIKeyHardcoded=" & DEFAULT_API_KEY_HARDCODED
    ts.WriteLine "ClassifyMaxTokens=" & DEFAULT_CLASSIFY_MAX_TOKENS
    ts.WriteLine "SummarizeMaxTokens=" & DEFAULT_SUMMARIZE_MAX_TOKENS
    ts.WriteLine "ReplyMaxTokens=" & DEFAULT_REPLY_MAX_TOKENS
    ts.WriteLine "Temperature=" & DEFAULT_LLM_TEMPERATURE
    ts.WriteLine "ReplyTemperature=" & DEFAULT_REPLY_TEMPERATURE
    ts.WriteLine "SystemPrompt=" & DEFAULT_LLM_SYSTEM_PROMPT
    ts.WriteLine ""
    ts.WriteLine "[Agent]"
    ts.WriteLine "EnableAutoReply=" & IIf(DEFAULT_ENABLE_AUTO_REPLY, "True", "False")
    ts.WriteLine "AutoReplyOnArrival=" & IIf(DEFAULT_AUTO_REPLY_ON_ARRIVAL, "True", "False")
    ts.WriteLine "LearnReplyFolder=" & DEFAULT_FOLDER_LEARN_REPLY
    ts.WriteLine "MaxReplyExamples=" & DEFAULT_MAX_REPLY_EXAMPLES
    ts.WriteLine "ReplyPersona=" & DEFAULT_REPLY_PERSONA
    ts.WriteLine "ScanSentItems=" & IIf(DEFAULT_SCAN_SENT_ITEMS, "True", "False")
    ts.WriteLine "ScanSentDays=" & DEFAULT_SCAN_SENT_DAYS
    ts.WriteLine "AutoReplyForSenders=" & DEFAULT_AUTO_REPLY_FOR_SENDERS

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    LogMessage "INFO", "Created default settings file: " & filePath
    Exit Sub

FileError:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "CreateDefaultSettingsFile failed: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' SETTINGS INI FILE - READER/WRITER
'-------------------------------------------------------------------------------

' Read a single string value from the INI file.
' Returns defaultValue if the section/key is not found.
Public Function ReadINISetting(ByVal section As String, ByVal key As String, ByVal defaultValue As String) As String
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim currentSection As String
    Dim eqPos As Long

    filePath = GetSettingsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        ReadINISetting = defaultValue
        Exit Function
    End If

    currentSection = ""

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        ' Skip empty lines and comments
        If Len(line) = 0 Then GoTo NextLine
        If Left(line, 1) = ";" Or Left(line, 1) = "#" Then GoTo NextLine

        ' Check for section header
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            currentSection = Mid(line, 2, Len(line) - 2)
            GoTo NextLine
        End If

        ' Check for matching section and key
        If LCase(currentSection) = LCase(section) Then
            eqPos = InStr(1, line, "=")
            If eqPos > 0 Then
                If LCase(Trim(Left(line, eqPos - 1))) = LCase(key) Then
                    ReadINISetting = Trim(Mid(line, eqPos + 1))
                    ts.Close
                    Set ts = Nothing
                    Set fso = Nothing
                    Exit Function
                End If
            End If
        End If

NextLine:
    Loop

    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    ReadINISetting = defaultValue
End Function

' Read a boolean value from the INI file
Public Function ReadINIBool(ByVal section As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim raw As String
    raw = ReadINISetting(section, key, IIf(defaultValue, "True", "False"))
    ReadINIBool = (LCase(Trim(raw)) = "true" Or Trim(raw) = "1")
End Function

' Read an integer value from the INI file
Public Function ReadINIInt(ByVal section As String, ByVal key As String, ByVal defaultValue As Integer) As Integer
    Dim raw As String
    raw = ReadINISetting(section, key, CStr(defaultValue))
    If IsNumeric(Trim(raw)) Then
        ReadINIInt = CInt(Trim(raw))
    Else
        ReadINIInt = defaultValue
    End If
End Function

' Read a double value from the INI file
Public Function ReadINIDouble(ByVal section As String, ByVal key As String, ByVal defaultValue As Double) As Double
    Dim raw As String
    raw = ReadINISetting(section, key, CStr(defaultValue))
    If IsNumeric(Trim(raw)) Then
        ReadINIDouble = CDbl(Trim(raw))
    Else
        ReadINIDouble = defaultValue
    End If
End Function

' Write a single value to the INI file (read-modify-write approach).
' Creates the section if it doesn't exist.
' Creates the key if it doesn't exist in the section.
Public Sub WriteINISetting(ByVal section As String, ByVal key As String, ByVal value As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim allLines As Collection
    Dim line As String
    Dim currentSection As String
    Dim eqPos As Long
    Dim found As Boolean
    Dim sectionFound As Boolean
    Dim lastSectionLine As Long
    Dim i As Long
    Dim j As Long
    Dim tempLines As Collection
    Dim replacedLines As Collection

    filePath = GetSettingsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set allLines = New Collection

    ' Read all existing lines
    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
        Do While Not ts.AtEndOfStream
            allLines.Add ts.ReadLine
        Loop
        ts.Close
        Set ts = Nothing
    End If

    ' Find and update the key, or add it
    currentSection = ""
    found = False
    sectionFound = False
    lastSectionLine = 0

    For i = 1 To allLines.Count
        line = Trim(allLines(i))

        ' Check for section header
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            ' If we were in the target section and didn't find the key,
            ' we need to insert the key before this new section
            If sectionFound And Not found Then
                ' Insert key=value before this section header
                Set tempLines = New Collection
                For j = 1 To i - 1
                    tempLines.Add allLines(j)
                Next j
                tempLines.Add key & "=" & value
                For j = i To allLines.Count
                    tempLines.Add allLines(j)
                Next j
                Set allLines = tempLines
                found = True
                GoTo WriteFile
            End If

            currentSection = Mid(line, 2, Len(line) - 2)
            If LCase(currentSection) = LCase(section) Then
                sectionFound = True
            End If
            lastSectionLine = i
            GoTo NextWriteLine
        End If

        ' Check for matching section and key
        If LCase(currentSection) = LCase(section) Then
            eqPos = InStr(1, line, "=")
            If eqPos > 0 Then
                If LCase(Trim(Left(line, eqPos - 1))) = LCase(key) Then
                    ' Replace this line
                    Set replacedLines = New Collection
                    For j = 1 To allLines.Count
                        If j = i Then
                            replacedLines.Add key & "=" & value
                        Else
                            replacedLines.Add allLines(j)
                        End If
                    Next j
                    Set allLines = replacedLines
                    found = True
                    GoTo WriteFile
                End If
            End If
        End If

NextWriteLine:
    Next i

    ' If section was found but key wasn't, append key at end of section (which is end of file)
    If sectionFound And Not found Then
        allLines.Add key & "=" & value
        found = True
    End If

    ' If section wasn't found, append section and key at end
    If Not found Then
        If allLines.Count > 0 Then allLines.Add ""
        allLines.Add "[" & section & "]"
        allLines.Add key & "=" & value
    End If

WriteFile:
    ' Write all lines back
    On Error GoTo WriteError
    Set ts = fso.CreateTextFile(filePath, True)  ' True = overwrite
    For i = 1 To allLines.Count
        ts.WriteLine allLines(i)
    Next i
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Sub

WriteError:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "WriteINISetting failed: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' LEARNED SENDERS - FILE I/O AND CACHE
'-------------------------------------------------------------------------------

' Build the full path to the learned senders file, creating the directory if needed
Public Function GetLearnedSendersFilePath() As String
    Dim folderPath As String
    Dim fso As Object

    folderPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER

    ' Create directory if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        LogMessage "INFO", "Created learned data folder: " & folderPath
    End If
    Set fso = Nothing

    GetLearnedSendersFilePath = folderPath & "\" & LEARNED_SENDERS_FILE
End Function

' Load learned senders from file into the in-memory cache
' Set forceReload = True to re-read the file even if already loaded
Public Sub LoadLearnedSenders(Optional ByVal forceReload As Boolean = False)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String

    If learnedSendersCacheLoaded And Not forceReload Then Exit Sub

    ' Initialize or clear the cache
    Set learnedSendersCache = CreateObject("Scripting.Dictionary")
    learnedSendersCache.CompareMode = 1  ' vbTextCompare (case-insensitive keys)

    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        ' No file yet - cache is empty but loaded
        learnedSendersCacheLoaded = True
        LogMessage "INFO", "No learned senders file found (will be created on first learn)"
        Set fso = Nothing
        Exit Sub
    End If

    ' Read the file line by line
    Dim email As String
    Dim action As String

    Set ts = fso.OpenTextFile(filePath, 1)  ' 1 = ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        ' Skip empty lines and comments
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                email = LCase(Trim(parts(0)))
                action = UCase(Trim(parts(1)))

                ' Validate action
                If action = "KEEP" Or action = "DELETE" Then
                    ' Last entry per sender wins (overwrite)
                    learnedSendersCache(email) = action
                End If
            End If
        End If
    Loop
    ts.Close

    learnedSendersCacheLoaded = True
    LogMessage "INFO", "Loaded " & learnedSendersCache.Count & " learned sender rules"

    Set ts = Nothing
    Set fso = Nothing
End Sub

' Look up a sender email in the learned cache
' Returns "KEEP", "DELETE", or "" (empty string if not found)
Public Function LookupLearnedSender(ByVal email As String) As String
    ' Ensure cache is loaded
    If Not learnedSendersCacheLoaded Then LoadLearnedSenders

    If learnedSendersCache Is Nothing Then
        LookupLearnedSender = ""
        Exit Function
    End If

    Dim lowerEmail As String
    lowerEmail = LCase(email)

    If learnedSendersCache.Exists(lowerEmail) Then
        LookupLearnedSender = learnedSendersCache(lowerEmail)
    Else
        LookupLearnedSender = ""
    End If
End Function

' Record a learned sender rule: append to file and update cache
Public Sub RecordLearnedSender(ByVal email As String, ByVal action As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim lowerEmail As String
    Dim timestamp As String

    lowerEmail = LCase(Trim(email))

    ' Warn if email has no @ (likely unresolved Exchange address)
    If InStr(1, lowerEmail, "@") = 0 Then
        LogMessage "WARN", "RecordLearnedSender: email has no @ sign (unresolved Exchange?): " & lowerEmail
    End If

    ' Validate action
    If action <> "KEEP" And action <> "DELETE" Then
        LogMessage "ERROR", "RecordLearnedSender: invalid action '" & action & "'"
        Exit Sub
    End If

    ' Ensure cache is loaded
    If Not learnedSendersCacheLoaded Then LoadLearnedSenders

    ' Update in-memory cache
    learnedSendersCache(lowerEmail) = action

    ' Append to file (with error handling to prevent leaked file handles)
    filePath = GetLearnedSendersFilePath()
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    On Error GoTo FileError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending, True = Create if missing
    ts.WriteLine lowerEmail & "|" & action & "|" & timestamp
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0

    LogMessage "INFO", "LEARNED " & action & " recorded for: " & lowerEmail

    Exit Sub

FileError:
    ' Ensure file handle is closed even on error
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "RecordLearnedSender file write failed: " & Err.Description
End Sub

' Restore a specific sender's emails from Deleted Items back to Inbox.
' Called automatically when a DELETE rule is reversed to KEEP.
' Returns the number of emails restored.
Public Function RestoreSenderFromDeleted(ByVal senderEmail As String) As Long
    Dim deletedFolder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim restoredCount As Long
    Dim itemEmail As String

    On Error GoTo PROC_ERR
    PushCallStack "Utilities.RestoreSenderFromDeleted"

    Set deletedFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems)
    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = deletedFolder.Items

    restoredCount = 0

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            itemEmail = GetSenderEmail(mail)

            If LCase(itemEmail) = LCase(senderEmail) Then
                mail.Move inbox
                restoredCount = restoredCount + 1
            End If
        End If
    Next i

    RestoreSenderFromDeleted = restoredCount

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "RestoreSenderFromDeleted", Err.Number, Err.Description
    RestoreSenderFromDeleted = restoredCount
    Resume PROC_EXIT
End Function

' Delete a specific sender's emails from Inbox.
' Called automatically when a KEEP rule is reversed to DELETE.
' Returns the number of emails deleted.
Public Function DeleteSenderFromInbox(ByVal senderEmail As String) As Long
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim deletedCount As Long
    Dim itemEmail As String

    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeleteSenderFromInbox"

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = inbox.Items

    deletedCount = 0

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            itemEmail = GetSenderEmail(mail)

            If LCase(itemEmail) = LCase(senderEmail) Then
                mail.Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next i

    DeleteSenderFromInbox = deletedCount

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "DeleteSenderFromInbox", Err.Number, Err.Description
    DeleteSenderFromInbox = deletedCount
    Resume PROC_EXIT
End Function

' Return the number of learned sender rules in cache
Public Function GetLearnedSendersCount() As Long
    If Not learnedSendersCacheLoaded Then LoadLearnedSenders

    If learnedSendersCache Is Nothing Then
        GetLearnedSendersCount = 0
    Else
        GetLearnedSendersCount = learnedSendersCache.Count
    End If
End Function

' Force reload learned senders from file and show count
Public Sub ReloadLearnedSenders()
    LoadLearnedSenders True  ' forceReload = True

    MsgBox "Learned senders reloaded." & vbCrLf & vbCrLf & _
           "Total rules: " & GetLearnedSendersCount() & vbCrLf & _
           "File: " & GetLearnedSendersFilePath(), _
           vbInformation, "Learned Senders"
End Sub

' Remove duplicate entries from the learned senders file.
' Keeps only the last (most recent) entry per sender email.
' The in-memory cache is reloaded after rewriting.
Public Sub DeduplicateLearnedSenders()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim dedupDict As Object    ' email -> full line (last wins)
    Dim orderList As Object    ' email -> insertion order (to preserve sequence)
    Dim lowerEmail As String
    Dim linesBefore As Long
    Dim linesAfter As Long
    Dim orderIndex As Long

    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        LogMessage "INFO", "DeduplicateLearnedSenders: no file to deduplicate"
        Set fso = Nothing
        Exit Sub
    End If

    ' Pass 1: read all lines, keep last entry per sender
    Set dedupDict = CreateObject("Scripting.Dictionary")
    dedupDict.CompareMode = 1  ' case-insensitive
    Set orderList = CreateObject("Scripting.Dictionary")
    orderList.CompareMode = 1
    orderIndex = 0
    linesBefore = 0

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            linesBefore = linesBefore + 1
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                lowerEmail = LCase(Trim(parts(0)))
                dedupDict(lowerEmail) = line   ' last entry wins
                ' Track order: first appearance sets position
                If Not orderList.Exists(lowerEmail) Then
                    orderList(lowerEmail) = orderIndex
                    orderIndex = orderIndex + 1
                End If
            End If
        End If
    Loop
    ts.Close
    Set ts = Nothing

    linesAfter = dedupDict.Count

    If linesAfter = linesBefore Then
        LogMessage "INFO", "DeduplicateLearnedSenders: no duplicates found (" & linesBefore & " lines)"
        Set fso = Nothing
        Set dedupDict = Nothing
        Set orderList = Nothing
        Exit Sub
    End If

    ' Pass 2: sort by insertion order and rewrite the file
    ' Build an array sorted by order index
    Dim keys As Variant
    Dim sortedLines() As String
    Dim idx As Long
    keys = dedupDict.keys

    ReDim sortedLines(0 To linesAfter - 1)
    Dim k As Variant
    For Each k In keys
        idx = orderList(k)
        sortedLines(idx) = dedupDict(k)
    Next k

    ' Rewrite the file (ForWriting = 2, overwrite)
    Set ts = fso.OpenTextFile(filePath, 2, True)  ' 2 = ForWriting
    Dim j As Long
    For j = 0 To linesAfter - 1
        ts.WriteLine sortedLines(j)
    Next j
    ts.Close

    Set ts = Nothing
    Set fso = Nothing
    Set dedupDict = Nothing
    Set orderList = Nothing

    ' Reload cache from the clean file
    LoadLearnedSenders True

    LogMessage "INFO", "DeduplicateLearnedSenders: " & linesBefore & " -> " & linesAfter & " lines (" & (linesBefore - linesAfter) & " duplicates removed)"
End Sub

'-------------------------------------------------------------------------------
' LEARNED SUBJECTS - FILE I/O AND CACHE
'-------------------------------------------------------------------------------

' Build the full path to the learned subjects file, creating the directory if needed
Public Function GetLearnedSubjectsFilePath() As String
    Dim folderPath As String
    Dim fso As Object

    folderPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER

    ' Create directory if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
        LogMessage "INFO", "Created learned data folder: " & folderPath
    End If
    Set fso = Nothing

    GetLearnedSubjectsFilePath = folderPath & "\" & LEARNED_SUBJECTS_FILE
End Function

' Load learned subjects from file into the in-memory cache
' Set forceReload = True to re-read the file even if already loaded
Public Sub LoadLearnedSubjects(Optional ByVal forceReload As Boolean = False)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String

    If learnedSubjectsCacheLoaded And Not forceReload Then Exit Sub

    ' Initialize or clear the cache
    Set learnedSubjectsCache = CreateObject("Scripting.Dictionary")
    learnedSubjectsCache.CompareMode = 1  ' vbTextCompare (case-insensitive keys)

    filePath = GetLearnedSubjectsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        ' No file yet - cache is empty but loaded
        learnedSubjectsCacheLoaded = True
        LogMessage "INFO", "No learned subjects file found (will be created on first learn)"
        Set fso = Nothing
        Exit Sub
    End If

    ' Read the file line by line
    Dim subj As String
    Dim subjAction As String

    Set ts = fso.OpenTextFile(filePath, 1)  ' 1 = ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)

        ' Skip empty lines and comments
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                subj = Trim(parts(0))
                subjAction = UCase(Trim(parts(1)))

                ' Only DELETE is valid for subject rules
                If subjAction = "DELETE" And Len(subj) > 0 Then
                    ' Last entry per subject wins (overwrite)
                    learnedSubjectsCache(subj) = subjAction
                End If
            End If
        End If
    Loop
    ts.Close

    learnedSubjectsCacheLoaded = True
    LogMessage "INFO", "Loaded " & learnedSubjectsCache.Count & " learned subject rules"

    Set ts = Nothing
    Set fso = Nothing
End Sub

' Look up a subject against all learned subject patterns
' Returns "DELETE" if any cached subject key is a substring of the incoming subject
' Returns "" (empty string) if no match
Public Function LookupLearnedSubject(ByVal subject As String) As String
    ' Ensure cache is loaded
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    If learnedSubjectsCache Is Nothing Then
        LookupLearnedSubject = ""
        Exit Function
    End If

    If learnedSubjectsCache.Count = 0 Then
        LookupLearnedSubject = ""
        Exit Function
    End If

    ' Sanitize incoming subject for consistent matching
    Dim cleanSubject As String
    cleanSubject = SanitizeSubject(subject)

    If Len(cleanSubject) = 0 Then
        LookupLearnedSubject = ""
        Exit Function
    End If

    ' Iterate all cached keys and check for substring match (case-insensitive)
    Dim cachedSubject As Variant
    For Each cachedSubject In learnedSubjectsCache.keys
        If Len(cachedSubject) > 0 Then
            If InStr(1, cleanSubject, CStr(cachedSubject), vbTextCompare) > 0 Then
                LookupLearnedSubject = learnedSubjectsCache(cachedSubject)
                Exit Function
            End If
        End If
    Next cachedSubject

    LookupLearnedSubject = ""
End Function

' Sanitize a subject string for use as a Dictionary key and pipe-delimited file entry.
' Strips newlines, pipe characters, and null characters that would corrupt the data file
' or cause "Invalid procedure call or argument" errors in Scripting.Dictionary.
Public Function SanitizeSubject(ByVal subject As String) As String
    Dim result As String
    result = subject

    ' Strip newlines (subjects can contain embedded CR/LF from Exchange)
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")

    ' Strip pipe character (our file delimiter)
    result = Replace(result, "|", " ")

    ' Strip null characters (cause error 5 in Scripting.Dictionary)
    result = Replace(result, Chr(0), "")

    ' Collapse multiple spaces and trim
    Do While InStr(1, result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    SanitizeSubject = Trim(result)
End Function

' Record a learned subject rule: append to file and update cache
Public Sub RecordLearnedSubject(ByVal subject As String, ByVal action As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim trimmedSubject As String
    Dim timestamp As String

    trimmedSubject = SanitizeSubject(subject)

    ' Validate: only DELETE is allowed for subject rules
    If action <> "DELETE" Then
        LogMessage "ERROR", "RecordLearnedSubject: invalid action '" & action & "' (only DELETE allowed)"
        Exit Sub
    End If

    If Len(trimmedSubject) = 0 Then
        LogMessage "WARN", "RecordLearnedSubject: empty subject, skipping"
        Exit Sub
    End If

    ' Ensure cache is loaded
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    ' Update in-memory cache
    learnedSubjectsCache(trimmedSubject) = action

    ' Append to file (with error handling to prevent leaked file handles)
    filePath = GetLearnedSubjectsFilePath()
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    On Error GoTo FileError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending, True = Create if missing
    ts.WriteLine trimmedSubject & "|" & action & "|" & timestamp
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0

    LogMessage "INFO", "LEARNED SUBJECT " & action & " recorded for: " & Left(trimmedSubject, 50)

    Exit Sub

FileError:
    ' Ensure file handle is closed even on error
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "RecordLearnedSubject file write failed: " & Err.Description
End Sub

' Return the number of learned subject rules in cache
Public Function GetLearnedSubjectsCount() As Long
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    If learnedSubjectsCache Is Nothing Then
        GetLearnedSubjectsCount = 0
    Else
        GetLearnedSubjectsCount = learnedSubjectsCache.Count
    End If
End Function

' Remove duplicate entries from the learned subjects file.
' Keeps only the last (most recent) entry per subject.
' The in-memory cache is reloaded after rewriting.
Public Sub DeduplicateLearnedSubjects()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim dedupDict As Object    ' subject -> full line (last wins)
    Dim orderList As Object    ' subject -> insertion order (to preserve sequence)
    Dim subjKey As String
    Dim linesBefore As Long
    Dim linesAfter As Long
    Dim orderIndex As Long

    filePath = GetLearnedSubjectsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        LogMessage "INFO", "DeduplicateLearnedSubjects: no file to deduplicate"
        Set fso = Nothing
        Exit Sub
    End If

    ' Pass 1: read all lines, keep last entry per subject
    Set dedupDict = CreateObject("Scripting.Dictionary")
    dedupDict.CompareMode = 1  ' case-insensitive
    Set orderList = CreateObject("Scripting.Dictionary")
    orderList.CompareMode = 1
    orderIndex = 0
    linesBefore = 0

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            linesBefore = linesBefore + 1
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                subjKey = Trim(parts(0))
                dedupDict(subjKey) = line   ' last entry wins
                ' Track order: first appearance sets position
                If Not orderList.Exists(subjKey) Then
                    orderList(subjKey) = orderIndex
                    orderIndex = orderIndex + 1
                End If
            End If
        End If
    Loop
    ts.Close
    Set ts = Nothing

    linesAfter = dedupDict.Count

    If linesAfter = linesBefore Then
        LogMessage "INFO", "DeduplicateLearnedSubjects: no duplicates found (" & linesBefore & " lines)"
        Set fso = Nothing
        Set dedupDict = Nothing
        Set orderList = Nothing
        Exit Sub
    End If

    ' Pass 2: sort by insertion order and rewrite the file
    Dim keys As Variant
    Dim sortedLines() As String
    Dim idx As Long
    keys = dedupDict.keys

    ReDim sortedLines(0 To linesAfter - 1)
    Dim k As Variant
    For Each k In keys
        idx = orderList(k)
        sortedLines(idx) = dedupDict(k)
    Next k

    ' Rewrite the file (ForWriting = 2, overwrite)
    Set ts = fso.OpenTextFile(filePath, 2, True)  ' 2 = ForWriting
    Dim j As Long
    For j = 0 To linesAfter - 1
        ts.WriteLine sortedLines(j)
    Next j
    ts.Close

    Set ts = Nothing
    Set fso = Nothing
    Set dedupDict = Nothing
    Set orderList = Nothing

    ' Reload cache from the clean file
    LoadLearnedSubjects True

    LogMessage "INFO", "DeduplicateLearnedSubjects: " & linesBefore & " -> " & linesAfter & " lines (" & (linesBefore - linesAfter) & " duplicates removed)"
End Sub

'-------------------------------------------------------------------------------
' LEARNED RULES - PUBLIC CACHE ACCESSORS (for export and dashboard)
'-------------------------------------------------------------------------------

' Return a Collection of sender emails where the learned action is DELETE.
' Used by ExportLearnedRulesToServer to push DELETE rules to Exchange.
Public Function GetLearnedDeleteSenders() As Collection
    Dim result As Collection
    Dim k As Variant

    Set result = New Collection

    If Not learnedSendersCacheLoaded Then LoadLearnedSenders

    If learnedSendersCache Is Nothing Then
        Set GetLearnedDeleteSenders = result
        Exit Function
    End If

    For Each k In learnedSendersCache.keys
        If learnedSendersCache(k) = "DELETE" And Len(CStr(k)) > 0 Then
            result.Add CStr(k)
        End If
    Next k

    Set GetLearnedDeleteSenders = result
End Function

' Return a Collection of all learned subject keys (all are DELETE action).
' Used by ExportLearnedRulesToServer to push subject DELETE rules to Exchange.
Public Function GetLearnedSubjectKeys() As Collection
    Dim result As Collection
    Dim k As Variant

    Set result = New Collection

    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    If learnedSubjectsCache Is Nothing Then
        Set GetLearnedSubjectKeys = result
        Exit Function
    End If

    For Each k In learnedSubjectsCache.keys
        result.Add CStr(k)
    Next k

    Set GetLearnedSubjectKeys = result
End Function

' Return a dictionary copy of all learned sender rules (email -> action).
' Used by the dashboard to display learned rules in a ListBox.
Public Function GetLearnedSendersCacheCopy() As Object
    Dim result As Object
    Dim k As Variant

    If Not learnedSendersCacheLoaded Then LoadLearnedSenders

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not learnedSendersCache Is Nothing Then
        For Each k In learnedSendersCache.keys
            result(k) = learnedSendersCache(k)
        Next k
    End If

    Set GetLearnedSendersCacheCopy = result
End Function

' Return a dictionary copy of all learned subject rules (subject -> action).
' Used by the dashboard to display learned rules in a ListBox.
Public Function GetLearnedSubjectsCacheCopy() As Object
    Dim result As Object
    Dim k As Variant

    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = 1

    If Not learnedSubjectsCache Is Nothing Then
        For Each k In learnedSubjectsCache.keys
            result(k) = learnedSubjectsCache(k)
        Next k
    End If

    Set GetLearnedSubjectsCacheCopy = result
End Function

'-------------------------------------------------------------------------------
' LEARNED RULES - DELETE INDIVIDUAL RULE
'-------------------------------------------------------------------------------

' Delete a single learned sender rule from both cache and file.
' Uses read-all -> filter -> write-all approach.
Public Sub DeleteLearnedSenderRule(ByVal email As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim allLines As Collection
    Dim lowerEmail As String
    Dim i As Long

    lowerEmail = LCase(Trim(email))

    ' Remove from in-memory cache
    If Not learnedSendersCacheLoaded Then LoadLearnedSenders
    If learnedSendersCache.Exists(lowerEmail) Then
        learnedSendersCache.Remove lowerEmail
    End If

    ' Rewrite the file without this sender's entries
    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        Exit Sub
    End If

    ' Read all lines, keeping only those NOT matching the target email
    Set allLines = New Collection

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 And Left(Trim(line), 1) <> "#" Then
            parts = Split(Trim(line), "|")
            If UBound(parts) >= 1 Then
                If LCase(Trim(parts(0))) <> lowerEmail Then
                    allLines.Add line
                End If
            Else
                allLines.Add line  ' Keep malformed lines as-is
            End If
        Else
            allLines.Add line  ' Keep comments and blank lines
        End If
    Loop
    ts.Close
    Set ts = Nothing

    ' Rewrite the file
    On Error GoTo FileError
    Set ts = fso.CreateTextFile(filePath, True)  ' Overwrite
    For i = 1 To allLines.Count
        ts.WriteLine allLines(i)
    Next i
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    LogMessage "INFO", "Deleted learned sender rule for: " & lowerEmail
    Exit Sub

FileError:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "DeleteLearnedSenderRule file write failed: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' LEARNED REPLIES - FILE I/O
'-------------------------------------------------------------------------------
' Reply pairs are stored as pipe-delimited records in learned_replies.txt:
'   original_subject|original_from|original_body_snippet|reply_body_snippet|timestamp
' Subjects/bodies are sanitized via SanitizeSubject before writing.

' Build the full path to the learned replies file
Public Function GetLearnedRepliesFilePath() As String
    Dim folderPath As String
    Dim fso As Object

    folderPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing

    GetLearnedRepliesFilePath = folderPath & "\" & LEARNED_REPLIES_FILE
End Function

' Sanitize a text snippet for pipe-delimited file storage.
' Strips newlines, pipes, and null chars (same rules as SanitizeSubject).
Private Function SanitizeSnippet(ByVal text As String, ByVal maxLen As Integer) As String
    Dim result As String
    result = Left(text, maxLen)
    result = Replace(result, vbCrLf, " ")
    result = Replace(result, vbCr, " ")
    result = Replace(result, vbLf, " ")
    result = Replace(result, "|", " ")
    result = Replace(result, Chr(0), "")
    SanitizeSnippet = Trim(result)
End Function

' Record a reply pair to learned_replies.txt (append-only)
Public Sub RecordLearnedReply(ByVal originalSubject As String, _
                               ByVal originalFrom As String, _
                               ByVal originalBodySnippet As String, _
                               ByVal replyBodySnippet As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim timestamp As String

    filePath = GetLearnedRepliesFilePath()
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Dim safeSubject As String
    Dim safeFrom As String
    Dim safeOrigBody As String
    Dim safeReplyBody As String

    safeSubject = SanitizeSubject(originalSubject)
    safeFrom = SanitizeSnippet(originalFrom, 200)
    safeOrigBody = SanitizeSnippet(originalBodySnippet, 500)
    safeReplyBody = SanitizeSnippet(replyBodySnippet, 1000)

    On Error GoTo FileError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, 8, True)  ' 8 = ForAppending, create if missing
    ts.WriteLine safeSubject & "|" & safeFrom & "|" & safeOrigBody & "|" & safeReplyBody & "|" & timestamp
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    On Error GoTo 0

    LogMessage "INFO", "Learned reply recorded for: " & Left(safeSubject, 50)
    Exit Sub

FileError:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "RecordLearnedReply file write failed: " & Err.Description
End Sub

' Load the most recent N reply pairs from learned_replies.txt.
' Returns a Collection of pipe-delimited strings (raw lines).
Public Function LoadRecentReplyPairs(ByVal maxPairs As Integer) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim filePath As String
    filePath = GetLearnedRepliesFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        Set LoadRecentReplyPairs = result
        Exit Function
    End If

    ' Read all lines, keep last maxPairs
    Dim allLines As Collection
    Set allLines = New Collection

    Dim ts As Object
    Dim line As String
    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            allLines.Add line
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    ' Return the last maxPairs lines
    Dim startIdx As Long
    startIdx = allLines.Count - maxPairs + 1
    If startIdx < 1 Then startIdx = 1

    Dim i As Long
    For i = startIdx To allLines.Count
        result.Add allLines(i)
    Next i

    Set LoadRecentReplyPairs = result
End Function

' Delete a single learned subject rule from both cache and file.
' Uses read-all -> filter -> write-all approach.
Public Sub DeleteLearnedSubjectRule(ByVal subject As String)
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim parts() As String
    Dim allLines As Collection
    Dim targetSubject As String
    Dim i As Long

    targetSubject = SanitizeSubject(subject)

    ' Remove from in-memory cache
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects
    If learnedSubjectsCache.Exists(targetSubject) Then
        learnedSubjectsCache.Remove targetSubject
    End If

    ' Rewrite the file without this subject's entries
    filePath = GetLearnedSubjectsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        Set fso = Nothing
        Exit Sub
    End If

    ' Read all lines, keeping only those NOT matching the target subject
    Set allLines = New Collection

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 And Left(Trim(line), 1) <> "#" Then
            parts = Split(Trim(line), "|")
            If UBound(parts) >= 1 Then
                ' Case-insensitive comparison (matching Dictionary behavior)
                If LCase(Trim(parts(0))) <> LCase(targetSubject) Then
                    allLines.Add line
                End If
            Else
                allLines.Add line
            End If
        Else
            allLines.Add line
        End If
    Loop
    ts.Close
    Set ts = Nothing

    ' Rewrite the file
    On Error GoTo FileError2
    Set ts = fso.CreateTextFile(filePath, True)  ' Overwrite
    For i = 1 To allLines.Count
        ts.WriteLine allLines(i)
    Next i
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    LogMessage "INFO", "Deleted learned subject rule for: " & Left(targetSubject, 50)
    Exit Sub

FileError2:
    If Not ts Is Nothing Then
        On Error Resume Next
        ts.Close
        On Error GoTo 0
    End If
    Set ts = Nothing
    Set fso = Nothing
    LogMessage "ERROR", "DeleteLearnedSubjectRule file write failed: " & Err.Description
End Sub

'===============================================================================
' WEB UI COMMAND BRIDGE HELPERS (v3.0)
' Support for file-based IPC with the Python Flask Web UI.
' Commands are JSON files in %APPDATA%\OutlookEmailFilter\commands\
'===============================================================================

' Return path to the commands directory (auto-creates it)
Public Function GetCommandsDir() As String
    Dim dir As String
    dir = Environ("APPDATA") & "\OutlookEmailFilter\commands"
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

' Write a result JSON file for a completed command
' status: "ok" or "error"; output: text to return to the Web UI
Public Sub WriteResultFile(ByVal cmdId As String, ByVal status As String, ByVal output As String)
    Dim stm As Object
    Dim resultPath As String
    Dim jsonLine As String

    ' Sanitize output for JSON embedding (escape all JSON-unsafe characters)
    Dim safe As String
    safe = Replace(output, "\", "\\")
    safe = Replace(safe, """", "\""")
    safe = Replace(safe, vbCrLf, "\n")
    safe = Replace(safe, vbCr, "\n")
    safe = Replace(safe, vbLf, "\n")
    safe = Replace(safe, vbTab, "\t")
    safe = Replace(safe, Chr(8), "")    ' backspace
    safe = Replace(safe, Chr(12), "")   ' form feed
    safe = Replace(safe, Chr(0), "")

    resultPath = GetCommandsDir() & "\" & cmdId & ".result"
    jsonLine = "{""id"":""" & cmdId & """,""status"":""" & status & """,""output"":""" & safe & """}"

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

' ---------------------------------------------------------------------------
' WEB UI COMMAND POLLER
' Uses Windows API SetTimer/KillTimer because Outlook VBA does NOT support
' Application.OnTime (that is Excel/Word only).
' ---------------------------------------------------------------------------

Public Sub StartCommandPollerStd()
    If pollerRunningFlag Then Exit Sub
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
    If Not pollerRunningFlag Then Exit Sub
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
    PollForCommandsTimer
    On Error GoTo 0
End Sub

Public Sub PollForCommandsTimer()
    If Not pollerRunningFlag Then Exit Sub

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
        GoTo Reschedule
    End If

    Set folder = fso.GetFolder(cmdDir)

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "json" Then
            cmdId = fso.GetBaseName(file.Name)

            On Error Resume Next
            Set ts = fso.OpenTextFile(file.Path, 1)
            If Err.Number <> 0 Then
                On Error GoTo 0
                GoTo NextFile
            End If
            content = ts.ReadAll
            ts.Close
            Set ts = Nothing
            On Error GoTo 0

            On Error Resume Next
            fso.DeleteFile file.Path
            On Error GoTo 0

            macroName = ExtractJsonStringStd(content, "macro")
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

Reschedule:
    ' No rescheduling needed — Windows API SetTimer repeats automatically
End Sub

Private Function ExtractJsonStringStd(ByVal json As String, ByVal key As String) As String
    Dim pos As Long
    Dim valueStart As Long
    Dim valueEnd As Long

    ' Try "key":"value" first (compact JSON)
    Dim searchKey As String
    searchKey = """" & key & """:"""
    pos = InStr(1, json, searchKey, vbTextCompare)
    If pos > 0 Then
        valueStart = pos + Len(searchKey)
    Else
        ' Try "key": "value" (Python json.dump default — space after colon)
        searchKey = """" & key & """: """
        pos = InStr(1, json, searchKey, vbTextCompare)
        If pos > 0 Then
            valueStart = pos + Len(searchKey)
        Else
            ExtractJsonStringStd = ""
            Exit Function
        End If
    End If

    valueEnd = InStr(valueStart, json, """")
    If valueEnd = 0 Then
        ExtractJsonStringStd = ""
        Exit Function
    End If

    ExtractJsonStringStd = Mid(json, valueStart, valueEnd - valueStart)
End Function

Private Function DispatchMacroStd(ByVal macroName As String, Optional ByVal rawJson As String = "") As String
    Dim result As String

    On Error GoTo DispatchError

    Select Case macroName
        Case "FilterExistingDryRun"
            result = CaptureFilterDryRunStd()

        Case "FilterExistingEmails"
            FilterExistingEmails
            result = "FilterExistingEmails completed."

        Case "ShowVersionInfo"
            result = BuildVersionInfoStd()

        Case "ReinitializeFilter"
            ' Cannot use Application.Run from timer callback context, so call
            ' the reinitialization logic directly from this standard module.
            LoadAllSettings
            LoadLearnedSenders True   ' forceReload
            LoadLearnedSubjects True  ' forceReload
            result = "Settings and learned rules reloaded."

        Case "ScanSentForReplyPatterns"
            EmailAgent.ScanSentForReplyPatterns
            result = "Sent items scan complete."

        Case "DraftRepliesForInbox"
            EmailAgent.DraftRepliesForInbox
            result = "Draft replies complete."

        Case "ShowLearnedSenders"
            result = "Learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
                     "File: " & GetLearnedSendersFilePath()

        Case "ShowLearnedSendersList"
            ShowLearnedSendersList
            result = "Sender rules dumped to VBA Immediate Window (" & GetLearnedSendersCount() & " unique rules)."

        Case "CleanLearnedSendersFile"
            DeduplicateLearnedSenders
            result = "Learned senders deduplicated. Unique rules: " & GetLearnedSendersCount()

        Case "ShowLearnedSubjectsList"
            ShowLearnedSubjectsList
            result = "Subject rules dumped to VBA Immediate Window (" & GetLearnedSubjectsCount() & " unique rules)."

        Case "CleanLearnedSubjectsFile"
            DeduplicateLearnedSubjects
            result = "Learned subjects deduplicated. Unique rules: " & GetLearnedSubjectsCount()

        Case "ImportExistingLearnedFolders"
            ImportExistingLearnedFolders
            result = "Learned folder import complete."

        Case "ImportExistingLearnedSubjectFolder"
            ImportExistingLearnedSubjectFolder
            result = "Learned subject folder import complete."

        Case "ImportServerRules"
            ImportServerRules
            result = "Server rule import complete."

        Case "ExportLearnedRulesToServer"
            ExportLearnedRulesToServer
            result = "Server rule export complete."

        Case "RestoreFromReview"
            RestoreFromReview
            result = "Restore from Review complete."

        Case "RestoreDeletedKeepEmails"
            RestoreDeletedKeepEmails
            result = "Restore deleted KEEP emails complete."

        Case "GenerateAddressingPatterns"
            EmailAgent.GenerateAddressingPatterns
            result = "Addressing patterns generated."

        Case "GenerateClassificationReport"
            GenerateClassificationReport
            result = "Classification report shown."

        Case "SummarizeSelectedEmail"
            result = SummarizeSelectedEmailStd()

        Case "DraftReplyToSelected"
            result = DraftReplyToSelectedStd()

        Case "FilterAllFolders"
            FilterAllFolders
            result = "FilterAllFolders completed."

        Case "FilterSelectedEmail"
            FilterSelectedEmail
            result = "FilterSelectedEmail completed — check Outlook for prompt."

        Case "FilterSelectedEmails"
            FilterSelectedEmails
            result = "FilterSelectedEmails completed."

        Case "FilterCurrentFolder"
            FilterCurrentFolder
            result = "FilterCurrentFolder completed."

        Case "FilterLastNDays"
            Dim daysArg As String
            daysArg = ExtractJsonStringStd(rawJson, "days")
            If Len(daysArg) > 0 And IsNumeric(daysArg) Then
                FilterLastNDays CInt(daysArg)
            Else
                FilterLastNDays 7
            End If
            result = "FilterLastNDays completed."

        Case "BulkDeleteBySender"
            Dim patternArg As String
            patternArg = ExtractJsonStringStd(rawJson, "pattern")
            If Len(patternArg) > 0 Then
                BulkDeleteBySender patternArg
                result = "BulkDeleteBySender completed for: " & patternArg
            Else
                result = "ERROR: BulkDeleteBySender requires a sender pattern argument."
            End If

        Case "MoveProtectedSources"
            MoveProtectedSources
            result = "MoveProtectedSources completed."

        Case "ShowLearnedRepliesSummary"
            EmailAgent.ShowLearnedRepliesSummary
            result = "Learned reply summary shown — check Outlook."

        Case "ReloadLearnedSenders"
            ReloadLearnedSenders
            result = "Learned senders reloaded. Count: " & GetLearnedSendersCount()

        Case "DetectAndMigrateOldFolders"
            DetectAndMigrateOldFolders
            result = "Old folder migration complete."

        Case "EnableRealTimeFilter"
            ' Requires WithEvents setup in ThisOutlookSession — cannot run from bridge.
            result = "Cannot run from Web UI. In VBA Immediate Window (Ctrl+G), type:" & vbCrLf & _
                     "  ThisOutlookSession.EnableRealTimeFilter"

        Case "DisableRealTimeFilter"
            result = "Cannot run from Web UI. In VBA Immediate Window (Ctrl+G), type:" & vbCrLf & _
                     "  ThisOutlookSession.DisableRealTimeFilter"

        Case Else
            result = "ERROR: Unknown macro: " & macroName

    End Select

    DispatchMacroStd = result
    Exit Function

DispatchError:
    DispatchMacroStd = "ERROR: " & macroName & " failed: " & Err.Description
End Function

Private Function BuildVersionInfoStd() As String
    Dim info As String
    info = "Email Agent v" & FILTER_VERSION & " (" & FILTER_VERSION_DATE & ")" & vbCrLf
    info = info & "Settings: " & GetSettingsFilePath() & vbCrLf
    info = info & "Learned senders: " & GetLearnedSendersCount() & vbCrLf
    info = info & "Learned subjects: " & GetLearnedSubjectsCount() & vbCrLf
    info = info & "LLM: " & IIf(RuntimeUseLLM, "ON (" & RuntimeLLMProvider & ")", "OFF") & vbCrLf
    info = info & "Self-improving: " & IIf(RuntimeEnableSelfImproving, "ON", "OFF") & vbCrLf
    info = info & "Auto-reply: " & IIf(RuntimeEnableAutoReply, "ON", "OFF")
    BuildVersionInfoStd = info
End Function

Private Function CaptureFilterDryRunStd() As String
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
            output = output & icon & " " & Format(mail.ReceivedTime, "mm/dd hh:nn") & " | " & _
                     Truncate(mail.SenderName, 20) & " | " & Truncate(mail.Subject, 40) & vbCrLf
        End If
    Next i

    output = output & vbCrLf & "Total: " & processCount & " emails previewed."
    CaptureFilterDryRunStd = output
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

' Bridge-friendly DraftReplyToSelected — returns result string, saves draft, no MsgBox
Private Function DraftReplyToSelectedStd() As String
    On Error GoTo StdErr
    Dim mail As Outlook.MailItem
    Dim replyItem As Outlook.MailItem
    Dim prompt As String
    Dim draft As String
    Dim systemPrompt As String
    Dim sSubject As String

    If Not RuntimeUseLLM Then
        DraftReplyToSelectedStd = "ERROR: LLM is not enabled. Set UseLLMAPI=True in settings."
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
    sSubject = mail.Subject

    systemPrompt = "You are Professor Xu Xin at PolyU Hong Kong. Draft a professional, concise reply to the following email. " & _
                   "Be polite and to the point. If the email requires a specific action, acknowledge it. " & _
                   "Do not include a subject line in your reply."

    prompt = "Draft a reply to this email:" & vbCrLf & _
             "From: " & mail.SenderName & " <" & GetSenderEmail(mail) & ">" & vbCrLf & _
             "Subject: " & sSubject & vbCrLf & _
             "Date: " & Format(mail.ReceivedTime, "yyyy-mm-dd hh:nn") & vbCrLf & _
             "Body:" & vbCrLf & Truncate(mail.Body, 2000)

    draft = CallLLM(prompt, systemPrompt, RuntimeReplyMaxTokens, RuntimeReplyTemperature)

    If Len(draft) = 0 Then
        DraftReplyToSelectedStd = "ERROR: LLM returned no response. Check your API configuration."
        Exit Function
    End If

    ' Save as actual draft in Outlook Drafts folder
    Set replyItem = mail.Reply
    replyItem.Body = draft & vbCrLf & vbCrLf & replyItem.Body
    replyItem.Save
    Set replyItem = Nothing

    LogMessage "INFO", "Draft reply saved to Drafts for: " & Left(sSubject, 50)
    DraftReplyToSelectedStd = "Draft reply saved to Drafts for: " & sSubject & vbCrLf & vbCrLf & draft
    Exit Function
StdErr:
    DraftReplyToSelectedStd = "ERROR: DraftReplyToSelected failed: " & Err.Description
End Function
