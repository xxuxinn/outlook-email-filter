'===============================================================================
' Utilities.bas - Helper Functions for Email Agent v3.1
'===============================================================================
' This module contains utility functions used across the email filter:
'   - String matching (ContainsAny, StartsWithAny, MatchesAny)
'   - JSON encoding/parsing (robust key extraction, escape-aware)
'   - UTF-8 file I/O helpers (settings + learned data files are UTF-8 w/ BOM)
'   - Folder management
'   - Logging (with size-capped rotation)
'   - Email address extraction
'   - Learned senders/subjects cache (in-memory Dictionary + file I/O)
'   - Settings INI reader/writer
'   - Learned rule deletion
'   - Call stack tracking and centralized error handling
'   - Multi-provider LLM caller (CallLLM) with HTTP timeouts
'   - Reply pair I/O (learned_replies.txt)
'
' The Web UI command bridge (poller + dispatcher) lives in Bridge.bas (v3.1).
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
    RotateLogIfOversize fso, logPath, ERROR_LOG_MAX_BYTES
    Set ts = fso.OpenTextFile(logPath, 8, True)  ' 8 = ForAppending, True = create if missing
    ts.WriteLine message
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

' Rotate a log file to <name>.old when it exceeds maxBytes (replacing any
' previous .old). Best-effort: callers already run under On Error Resume Next.
Private Sub RotateLogIfOversize(ByVal fso As Object, ByVal logPath As String, ByVal maxBytes As Long)
    If Not fso.FileExists(logPath) Then Exit Sub
    If fso.GetFile(logPath).Size <= maxBytes Then Exit Sub

    Dim oldPath As String
    oldPath = logPath & ".old"
    If fso.FileExists(oldPath) Then fso.DeleteFile oldPath, True
    fso.MoveFile logPath, oldPath
End Sub

'-------------------------------------------------------------------------------
' MULTI-PROVIDER LLM CALLER
'-------------------------------------------------------------------------------
' CallLLM routes to the configured provider (local | azure | claude | openai).
' Returns the raw response content string, or "" on error/no key.
' All callers should use CallLLM instead of CallAzureOpenAICustom (kept for
' backwards-compatibility but now delegates here).

Public Function CallLLM(ByVal userPrompt As String, ByVal systemPrompt As String, _
                        ByVal maxTokens As Long, _
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

' Build OpenAI-compatible chat/completions JSON body.
' Pass modelName = "" to omit the model field (Azure encodes it in the URL).
Private Function BuildOpenAIBody(ByVal userPrompt As String, ByVal systemPrompt As String, _
                                  ByVal maxTokens As Long, ByVal temp As Double, _
                                  Optional ByVal modelName As String = "") As String
    ' Force "." decimal separator regardless of Windows locale
    Dim tempStr As String
    tempStr = Format(temp, "0.00")
    tempStr = Replace(tempStr, ",", ".")

    Dim modelPart As String
    If Len(modelName) > 0 Then
        modelPart = """model"":""" & EscapeJSON(modelName) & ""","
    Else
        modelPart = ""
    End If

    BuildOpenAIBody = "{" & modelPart & _
        """messages"":[" & _
            "{""role"":""system"",""content"":""" & EscapeJSON(systemPrompt) & """}," & _
            "{""role"":""user"",""content"":""" & EscapeJSON(userPrompt) & """}" & _
        "]," & _
        """max_tokens"":" & maxTokens & "," & _
        """temperature"":" & tempStr & "," & _
        """stream"":false" & _
    "}"
End Function

' Create an HTTP client with real timeouts. ServerXMLHTTP (unlike XMLHTTP)
' supports setTimeouts, so a stalled LLM endpoint can no longer freeze
' the Outlook UI thread indefinitely.
Private Function CreateHTTPClient() As Object
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim receiveMs As Long
    receiveMs = RuntimeLLMTimeoutSeconds * 1000
    If receiveMs <= 0 Then receiveMs = DEFAULT_LLM_TIMEOUT_SECONDS * 1000

    ' resolve, connect, send, receive (milliseconds)
    http.setTimeouts 10000, 10000, 30000, receiveMs
    Set CreateHTTPClient = http
End Function

' Call local OpenAI-compatible endpoint (Ollama, LM Studio, Inferencer, etc.)
Private Function CallLLMLocal(ByVal userPrompt As String, ByVal systemPrompt As String, _
                               ByVal maxTokens As Long, ByVal temp As Double) As String
    Dim http As Object
    Dim body As String

    body = BuildOpenAIBody(userPrompt, systemPrompt, maxTokens, temp, RuntimeLocalModel)

    LogMessage "DEBUG", "Calling local LLM at " & RuntimeLocalEndpoint & " (model: " & RuntimeLocalModel & ")"

    Set http = CreateHTTPClient()
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
                                ByVal maxTokens As Long, ByVal temp As Double) As String
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

    Dim body As String
    body = BuildOpenAIBody(userPrompt, systemPrompt, maxTokens, temp, RuntimeOpenAIModel)

    LogMessage "DEBUG", "Calling OpenAI-compatible API at " & RuntimeOpenAIEndpoint & " (model: " & RuntimeOpenAIModel & ")"

    Set http = CreateHTTPClient()
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
                               ByVal maxTokens As Long, ByVal temp As Double) As String
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

    Set http = CreateHTTPClient()
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
                                ByVal maxTokens As Long, ByVal temp As Double) As String
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

    Set http = CreateHTTPClient()
    http.Open "POST", RuntimeClaudeEndpoint, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "x-api-key", apiKey
    http.setRequestHeader "anthropic-version", "2023-06-01"
    http.send body

    If http.Status = 200 Then
        ' Claude response format: {"content":[{"type":"text","text":"..."}],...}
        Dim claudeResult As String
        claudeResult = ExtractJSONStringValue(http.responseText, "text")

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

' Escape a string for JSON encoding.
' Handles backslash, quote, named control chars, and strips/escapes ALL
' remaining chars < 0x20 (a crafted email with e.g. a vertical tab could
' otherwise produce structurally invalid JSON and break its own classification).
Public Function EscapeJSON(ByVal text As String) As String
    Dim result As String

    result = text

    ' Escape backslashes first
    result = Replace(result, "\", "\\")

    ' Escape double quotes
    result = Replace(result, """", "\""")

    ' Escape named control characters
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    result = Replace(result, Chr(8), "\b")
    result = Replace(result, Chr(12), "\f")

    ' Strip any remaining control characters (0x00-0x1F)
    Dim i As Long
    For i = 0 To 31
        Select Case i
            Case 8, 9, 10, 12, 13  ' already handled above
            Case Else
                If InStr(1, result, Chr(i)) > 0 Then
                    result = Replace(result, Chr(i), "")
                End If
        End Select
    Next i

    EscapeJSON = result
End Function

' Unescape a JSON string value. Uses a placeholder for "\\" so that a literal
' backslash followed by 'n' does not get mangled into a newline (the old
' replace-order bug garbled Windows paths in LLM output).
Public Function UnescapeJSON(ByVal text As String) As String
    Dim result As String
    Dim marker As String
    marker = Chr(1)  ' unused control char as temporary placeholder

    result = Replace(text, "\\", marker)
    result = Replace(result, "\n", vbCrLf)
    result = Replace(result, "\r", "")
    result = Replace(result, "\t", vbTab)
    result = Replace(result, "\""", """")
    result = Replace(result, "\/", "/")
    result = Replace(result, marker, "\")

    UnescapeJSON = result
End Function

' Extract a JSON string value by key, robustly:
'   - tolerates any whitespace between the colon and the opening quote
'   - returns "" for null / non-string values instead of grabbing a later field
'   - escape detection counts consecutive preceding backslashes, so \\" is
'     correctly recognised as escaped-backslash + closing quote
' Searches from startAt (default 1); returns "" if key not found.
Public Function ExtractJSONStringValue(ByVal jsonText As String, ByVal key As String, _
                                       Optional ByVal startAt As Long = 1) As String
    ExtractJSONStringValue = ""

    Dim keyToken As String
    keyToken = """" & key & """"

    Dim keyPos As Long
    keyPos = InStr(startAt, jsonText, keyToken, vbTextCompare)
    If keyPos = 0 Then Exit Function

    ' Find the colon after the key
    Dim pos As Long
    pos = keyPos + Len(keyToken)
    Do While pos <= Len(jsonText) And Mid(jsonText, pos, 1) <> ":"
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid(jsonText, pos, 1)) = 0 Then Exit Function
        pos = pos + 1
    Loop
    If pos > Len(jsonText) Then Exit Function
    pos = pos + 1  ' skip the colon

    ' Skip whitespace after the colon
    Do While pos <= Len(jsonText) And InStr(1, " " & vbTab & vbCr & vbLf, Mid(jsonText, pos, 1)) > 0
        pos = pos + 1
    Loop
    If pos > Len(jsonText) Then Exit Function

    ' Value must be a string; null/number/object -> return ""
    If Mid(jsonText, pos, 1) <> """" Then Exit Function

    Dim valueStart As Long
    valueStart = pos + 1

    ' Scan for the closing quote, counting preceding backslashes
    Dim endPos As Long
    Dim backslashes As Long
    Dim scanPos As Long
    endPos = 0
    scanPos = valueStart
    Do While scanPos <= Len(jsonText)
        If Mid(jsonText, scanPos, 1) = """" Then
            ' Count consecutive backslashes immediately before this quote
            backslashes = 0
            Do While scanPos - backslashes - 1 >= valueStart And _
                     Mid(jsonText, scanPos - backslashes - 1, 1) = "\"
                backslashes = backslashes + 1
            Loop
            If backslashes Mod 2 = 0 Then
                endPos = scanPos
                Exit Do
            End If
        End If
        scanPos = scanPos + 1
    Loop
    If endPos = 0 Then Exit Function

    ExtractJSONStringValue = UnescapeJSON(Mid(jsonText, valueStart, endPos - valueStart))
End Function

' Extract a JSON numeric value by key (e.g. "confidence": 0.85 or "urgency": 3).
' Returns defaultValue if the key is missing or the value is not numeric.
' Locale-safe: JSON always uses "." which CDbl may reject on comma locales,
' so the digits are parsed manually.
Public Function ExtractJSONNumberValue(ByVal jsonText As String, ByVal key As String, _
                                       ByVal defaultValue As Double, _
                                       Optional ByVal startAt As Long = 1) As Double
    ExtractJSONNumberValue = defaultValue

    Dim keyToken As String
    keyToken = """" & key & """"

    Dim keyPos As Long
    keyPos = InStr(startAt, jsonText, keyToken, vbTextCompare)
    If keyPos = 0 Then Exit Function

    Dim pos As Long
    pos = keyPos + Len(keyToken)
    Do While pos <= Len(jsonText) And Mid(jsonText, pos, 1) <> ":"
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid(jsonText, pos, 1)) = 0 Then Exit Function
        pos = pos + 1
    Loop
    If pos > Len(jsonText) Then Exit Function
    pos = pos + 1

    Do While pos <= Len(jsonText) And InStr(1, " " & vbTab & vbCr & vbLf & """", Mid(jsonText, pos, 1)) > 0
        pos = pos + 1
    Loop
    If pos > Len(jsonText) Then Exit Function

    ' Collect number characters
    Dim numStr As String
    Dim ch As String
    Do While pos <= Len(jsonText)
        ch = Mid(jsonText, pos, 1)
        If InStr(1, "0123456789.-+eE", ch) > 0 Then
            numStr = numStr & ch
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    If Len(numStr) = 0 Then Exit Function

    ' Manual locale-safe parse of the common "int.frac" shape
    On Error GoTo ParseFail
    Dim dotPos As Long
    dotPos = InStr(1, numStr, ".")
    If dotPos = 0 Then
        ExtractJSONNumberValue = CDbl(CLng(numStr))
    Else
        Dim intPart As String
        Dim fracPart As String
        intPart = Left(numStr, dotPos - 1)
        fracPart = Mid(numStr, dotPos + 1)
        If Len(intPart) = 0 Or intPart = "-" Then intPart = intPart & "0"
        If Len(fracPart) = 0 Then fracPart = "0"
        Dim sign As Double
        sign = IIf(Left(intPart, 1) = "-", -1, 1)
        ExtractJSONNumberValue = CDbl(CLng(Replace(intPart, "-", ""))) + CDbl(CLng(fracPart)) / (10 ^ Len(fracPart))
        ExtractJSONNumberValue = ExtractJSONNumberValue * sign
    End If
    Exit Function
ParseFail:
    ExtractJSONNumberValue = defaultValue
End Function

' Parse a chat/completions JSON response to extract the "content" field.
' Delegates to the robust escape-aware extractor.
Public Function ParseJSONContent(ByVal jsonText As String) As String
    ParseJSONContent = ExtractJSONStringValue(jsonText, "content")
End Function

'-------------------------------------------------------------------------------
' UTF-8 FILE I/O HELPERS (v3.1)
'-------------------------------------------------------------------------------
' settings.ini and all learned data files are UTF-8 with BOM as of v3.1 so the
' Python Web UI / MCP server and VBA agree on encoding (Chinese patterns like
' 優惠 previously corrupted across the ANSI/UTF-8 boundary).
' Legacy ANSI files are detected by missing BOM and migrated on first write.

' Read an entire text file. If it starts with a UTF-8 BOM, read as UTF-8;
' otherwise read as ANSI (legacy files written before v3.1).
Public Function ReadTextFileSmart(ByVal filePath As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.ReadTextFileSmart"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ReadTextFileSmart = ""
    If Not fso.FileExists(filePath) Then GoTo PROC_EXIT
    If fso.GetFile(filePath).Size = 0 Then GoTo PROC_EXIT

    If FileHasUTF8BOM(filePath) Then
        Dim stm As Object
        Set stm = CreateObject("ADODB.Stream")
        stm.Type = 2            ' adTypeText
        stm.Charset = "utf-8"
        stm.Open
        stm.LoadFromFile filePath
        ReadTextFileSmart = stm.ReadText(-1)  ' adReadAll
        stm.Close
        Set stm = Nothing
    Else
        Dim ts As Object
        Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading, ANSI
        If Not ts.AtEndOfStream Then ReadTextFileSmart = ts.ReadAll
        ts.Close
        Set ts = Nothing
    End If

PROC_EXIT:
    Set fso = Nothing
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "ReadTextFileSmart", Err.Number, Err.Description, filePath
    ReadTextFileSmart = ""
    Resume PROC_EXIT
End Function

' True if the file starts with the UTF-8 BOM (EF BB BF)
Private Function FileHasUTF8BOM(ByVal filePath As String) As Boolean
    On Error GoTo NoBOM

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1  ' adTypeBinary
    stm.Open
    stm.LoadFromFile filePath
    If stm.Size >= 3 Then
        Dim head() As Byte
        head = stm.Read(3)
        FileHasUTF8BOM = (head(0) = 239 And head(1) = 187 And head(2) = 191)
    End If
    stm.Close
    Set stm = Nothing
    Exit Function
NoBOM:
    FileHasUTF8BOM = False
End Function

' Write an entire text file as UTF-8 with BOM (overwrite)
Public Sub WriteTextFileUTF8(ByVal filePath As String, ByVal content As String)
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.WriteTextFileUTF8"

    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2            ' adTypeText
    stm.Charset = "utf-8"   ' ADODB writes the BOM for utf-8
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2  ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "WriteTextFileUTF8", Err.Number, Err.Description, filePath
    Resume PROC_EXIT
End Sub

' Append one line to a UTF-8 file as a TRUE append (binary Put at EOF — no
' full-file read/rewrite, so a crash mid-write cannot destroy existing data
' and the cost stays O(line), not O(file)). If the file exists without a BOM
' (legacy ANSI), it is migrated to UTF-8 once; new/empty files get a BOM so
' all readers detect UTF-8.
Public Sub AppendLineUTF8(ByVal filePath As String, ByVal lineText As String)
    Dim fnum As Integer
    fnum = 0

    On Error GoTo PROC_ERR
    PushCallStack "Utilities.AppendLineUTF8"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim needsBom As Boolean
    needsBom = False

    If fso.FileExists(filePath) Then
        If fso.GetFile(filePath).Size = 0 Then
            needsBom = True
        ElseIf Not FileHasUTF8BOM(filePath) Then
            ' One-time migration of a legacy ANSI file to UTF-8 with BOM
            WriteTextFileUTF8 filePath, ReadTextFileSmart(filePath)
        End If
    Else
        needsBom = True
    End If
    Set fso = Nothing

    Dim bytes() As Byte
    bytes = StringToUtf8Bytes(lineText & vbCrLf)

    fnum = FreeFile
    Open filePath For Binary Access Write As #fnum
    If needsBom Then
        Dim bom(0 To 2) As Byte
        bom(0) = 239: bom(1) = 187: bom(2) = 191
        Put #fnum, LOF(fnum) + 1, bom
    End If
    Put #fnum, LOF(fnum) + 1, bytes
    Close #fnum
    fnum = 0

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "AppendLineUTF8", Err.Number, Err.Description, filePath
    On Error Resume Next
    If fnum <> 0 Then Close #fnum
    GoTo PROC_EXIT
End Sub

' Encode a VBA string as UTF-8 bytes WITHOUT the BOM (ADODB writes one; we
' strip it so appended chunks never inject mid-file BOMs).
Private Function StringToUtf8Bytes(ByVal s As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2            ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1            ' adTypeBinary
    stm.Position = 3        ' skip the BOM
    StringToUtf8Bytes = stm.Read(-1)  ' adReadAll
    stm.Close
    Set stm = Nothing
End Function

' Public rotation helper for data files that grow unbounded (e.g. decision_log.txt)
Public Sub RotateFileIfOversize(ByVal filePath As String, ByVal maxBytes As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    RotateLogIfOversize fso, filePath, maxBytes
    Set fso = Nothing
End Sub

' Split file content into lines, tolerating CRLF and bare LF
Public Function SplitLines(ByVal content As String) As Variant
    Dim normalized As String
    normalized = Replace(content, vbCrLf, vbLf)
    normalized = Replace(normalized, vbCr, vbLf)
    SplitLines = Split(normalized, vbLf)
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
Public Function Truncate(ByVal text As String, ByVal maxLength As Long) As String
    If maxLength <= 3 Then
        ' Guard: Left(text, negative) raises error 5
        Truncate = Left(text, IIf(maxLength < 0, 0, maxLength))
    ElseIf Len(text) <= maxLength Then
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
                              ByVal model As String, ByVal maxTokens As Long, _
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
    RotateLogIfOversize fso, logPath, LLM_DEBUG_LOG_MAX_BYTES
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
' A failing individual setting is logged and skipped (Resume Next) so one bad
' value can no longer abort startup with half-applied configuration.
Public Sub LoadAllSettings()
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.LoadAllSettings"

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
    RuntimeClassifyBodyChars = ReadINIInt("LLM", "ClassifyBodyChars", DEFAULT_CLASSIFY_BODY_CHARS)
    RuntimeClassifyMaxTokens = ReadINIInt("LLM", "ClassifyMaxTokens", DEFAULT_CLASSIFY_MAX_TOKENS)
    RuntimeSummarizeMaxTokens = ReadINIInt("LLM", "SummarizeMaxTokens", DEFAULT_SUMMARIZE_MAX_TOKENS)
    RuntimeReplyMaxTokens = ReadINIInt("LLM", "ReplyMaxTokens", DEFAULT_REPLY_MAX_TOKENS)
    RuntimeLLMTemperature = ReadINIDouble("LLM", "Temperature", DEFAULT_LLM_TEMPERATURE)
    RuntimeReplyTemperature = ReadINIDouble("LLM", "ReplyTemperature", DEFAULT_REPLY_TEMPERATURE)
    RuntimeLLMSystemPrompt = ReadINISetting("LLM", "SystemPrompt", DEFAULT_LLM_SYSTEM_PROMPT)
    RuntimeLLMTimeoutSeconds = ReadINIInt("LLM", "RequestTimeoutSeconds", DEFAULT_LLM_TIMEOUT_SECONDS)
    RuntimeConfidenceThreshold = ReadINIDouble("LLM", "ConfidenceThreshold", DEFAULT_CONFIDENCE_THRESHOLD)

    ' --- Agent ---
    RuntimeEnableAutoReply = ReadINIBool("Agent", "EnableAutoReply", DEFAULT_ENABLE_AUTO_REPLY)
    RuntimeAutoReplyOnArrival = ReadINIBool("Agent", "AutoReplyOnArrival", DEFAULT_AUTO_REPLY_ON_ARRIVAL)
    RuntimeFolderLearnReply = ReadINISetting("Agent", "LearnReplyFolder", DEFAULT_FOLDER_LEARN_REPLY)
    RuntimeMaxReplyExamples = ReadINIInt("Agent", "MaxReplyExamples", DEFAULT_MAX_REPLY_EXAMPLES)
    RuntimeReplyPersona = ReadINISetting("Agent", "ReplyPersona", DEFAULT_REPLY_PERSONA)
    RuntimeScanSentItems = ReadINIBool("Agent", "ScanSentItems", DEFAULT_SCAN_SENT_ITEMS)
    RuntimeScanSentDays = ReadINIInt("Agent", "ScanSentDays", DEFAULT_SCAN_SENT_DAYS)
    RuntimeAutoReplyForSenders = ReadINISetting("Agent", "AutoReplyForSenders", DEFAULT_AUTO_REPLY_FOR_SENDERS)
    RuntimeEnableTaskExtraction = ReadINIBool("Agent", "EnableTaskExtraction", DEFAULT_ENABLE_TASK_EXTRACTION)
    RuntimeEnableContextEnrichment = ReadINIBool("Agent", "EnableContextEnrichment", DEFAULT_ENABLE_CONTEXT_ENRICHMENT)

    ' --- Digest / rule mining ---
    RuntimeEnableDailyDigest = ReadINIBool("Digest", "EnableDailyDigest", DEFAULT_ENABLE_DAILY_DIGEST)
    RuntimeDigestHour = ReadINIInt("Digest", "DigestHour", DEFAULT_DIGEST_HOUR)
    RuntimeDigestMaxEmails = ReadINIInt("Digest", "DigestMaxEmails", DEFAULT_DIGEST_MAX_EMAILS)
    RuntimeDigestSendEmail = ReadINIBool("Digest", "DigestSendEmail", DEFAULT_DIGEST_SEND_EMAIL)
    RuntimeEnableRuleMining = ReadINIBool("Digest", "EnableRuleMining", DEFAULT_ENABLE_RULE_MINING)

    ' --- Sync ---
    RuntimeEnableCloudSync = ReadINIBool("Sync", "EnableCloudSync", DEFAULT_ENABLE_CLOUD_SYNC)
    RuntimeCloudSyncPath = ReadINISetting("Sync", "CloudSyncPath", DEFAULT_CLOUD_SYNC_PATH)

    ' --- Debug / Error handling ---
    RuntimeDebugMode = ReadINIBool("General", "DebugMode", DEFAULT_DEBUG_MODE)
    RuntimeErrorLogFile = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & ERROR_LOG_FILE_NAME

    RuntimeSettingsLoaded = True
    LogMessage "INFO", "Settings loaded from: " & settingsPath

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "LoadAllSettings", Err.Number, Err.Description
    Resume Next  ' Skip the failing setting; ReadINI* fall back to defaults anyway
End Sub

' Format a Double for INI storage with "." decimal separator on any locale
Private Function FormatNumberForINI(ByVal value As Double) As String
    FormatNumberForINI = Replace(Format(value, "0.00"), ",", ".")
End Function

' Write a default settings.ini from DEFAULT_* constants (UTF-8 with BOM)
Public Sub CreateDefaultSettingsFile()
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.CreateDefaultSettingsFile"

    Dim filePath As String
    Dim s As String

    filePath = GetSettingsFilePath()

    s = "[General]" & vbCrLf
    s = s & "Version=" & FILTER_VERSION & vbCrLf
    s = s & "EnableLogging=" & IIf(DEFAULT_ENABLE_LOGGING, "True", "False") & vbCrLf
    s = s & "LogLevel=" & DEFAULT_LOG_LEVEL & vbCrLf
    s = s & "EnableSelfImproving=" & IIf(DEFAULT_ENABLE_SELF_IMPROVING, "True", "False") & vbCrLf
    s = s & "DebugMode=" & IIf(DEFAULT_DEBUG_MODE, "True", "False") & vbCrLf
    s = s & "ProgressInterval=" & DEFAULT_PROGRESS_INTERVAL & vbCrLf
    s = s & "DryRunLimit=" & DEFAULT_DRY_RUN_LIMIT & vbCrLf
    s = s & "LLMBatchSize=" & DEFAULT_LLM_BATCH_SIZE & vbCrLf
    s = s & vbCrLf
    s = s & "[Folders]" & vbCrLf
    s = s & "Protected=" & DEFAULT_FOLDER_PROTECTED & vbCrLf
    s = s & "Review=" & DEFAULT_FOLDER_REVIEW & vbCrLf
    s = s & "LearnKeep=" & DEFAULT_FOLDER_LEARN_KEEP & vbCrLf
    s = s & "LearnDelete=" & DEFAULT_FOLDER_LEARN_DELETE & vbCrLf
    s = s & "LearnSubject=" & DEFAULT_FOLDER_LEARN_SUBJECT_DELETE & vbCrLf
    s = s & vbCrLf
    s = s & "[Patterns]" & vbCrLf
    s = s & "ProtectedDomains=" & DEFAULT_PROTECTED_DOMAINS & vbCrLf
    s = s & "NamePatterns=" & DEFAULT_NAME_PATTERNS & vbCrLf
    s = s & "GreetingPatterns=" & DEFAULT_GREETING_PATTERNS & vbCrLf
    s = s & "PolyUTags=" & DEFAULT_POLYU_TAGS & vbCrLf
    s = s & "VIPSubjectKeywords=" & DEFAULT_VIP_SUBJECT_KEYWORDS & vbCrLf
    s = s & "DeleteSenderPatterns=" & DEFAULT_DELETE_SENDER_PATTERNS & vbCrLf
    s = s & "DeleteKnownSenders=" & DEFAULT_DELETE_KNOWN_SENDERS & vbCrLf
    s = s & "DeleteSubjectPatterns=" & DEFAULT_DELETE_SUBJECT_PATTERNS & vbCrLf
    s = s & vbCrLf
    s = s & "[LLM]" & vbCrLf
    s = s & "UseLLMAPI=" & IIf(DEFAULT_USE_LLM_API, "True", "False") & vbCrLf
    s = s & "; Provider: local | azure | claude | openai" & vbCrLf
    s = s & "Provider=" & DEFAULT_LLM_PROVIDER & vbCrLf
    s = s & "AzureEndpoint=" & DEFAULT_AZURE_OPENAI_ENDPOINT & vbCrLf
    s = s & "LocalEndpoint=" & DEFAULT_LOCAL_ENDPOINT & vbCrLf
    s = s & "LocalModel=" & DEFAULT_LOCAL_MODEL & vbCrLf
    s = s & "ClaudeEndpoint=" & DEFAULT_CLAUDE_ENDPOINT & vbCrLf
    s = s & "ClaudeModel=" & DEFAULT_CLAUDE_MODEL & vbCrLf
    s = s & "OpenAIEndpoint=" & DEFAULT_OPENAI_COMPAT_ENDPOINT & vbCrLf
    s = s & "OpenAIModel=" & DEFAULT_OPENAI_COMPAT_MODEL & vbCrLf
    s = s & "APIKeyMethod=" & DEFAULT_API_KEY_METHOD & vbCrLf
    s = s & "APIKeyEnvVar=" & DEFAULT_API_KEY_ENV_VAR & vbCrLf
    s = s & "APIKeyHardcoded=" & DEFAULT_API_KEY_HARDCODED & vbCrLf
    s = s & "ClassifyBodyChars=" & DEFAULT_CLASSIFY_BODY_CHARS & vbCrLf
    s = s & "ClassifyMaxTokens=" & DEFAULT_CLASSIFY_MAX_TOKENS & vbCrLf
    s = s & "SummarizeMaxTokens=" & DEFAULT_SUMMARIZE_MAX_TOKENS & vbCrLf
    s = s & "ReplyMaxTokens=" & DEFAULT_REPLY_MAX_TOKENS & vbCrLf
    s = s & "Temperature=" & FormatNumberForINI(DEFAULT_LLM_TEMPERATURE) & vbCrLf
    s = s & "ReplyTemperature=" & FormatNumberForINI(DEFAULT_REPLY_TEMPERATURE) & vbCrLf
    s = s & "RequestTimeoutSeconds=" & DEFAULT_LLM_TIMEOUT_SECONDS & vbCrLf
    s = s & "; Min LLM confidence (0-1) required to act on a DELETE decision; below -> Review" & vbCrLf
    s = s & "ConfidenceThreshold=" & FormatNumberForINI(DEFAULT_CONFIDENCE_THRESHOLD) & vbCrLf
    s = s & "SystemPrompt=" & DEFAULT_LLM_SYSTEM_PROMPT & vbCrLf
    s = s & vbCrLf
    s = s & "[Agent]" & vbCrLf
    s = s & "EnableAutoReply=" & IIf(DEFAULT_ENABLE_AUTO_REPLY, "True", "False") & vbCrLf
    s = s & "AutoReplyOnArrival=" & IIf(DEFAULT_AUTO_REPLY_ON_ARRIVAL, "True", "False") & vbCrLf
    s = s & "LearnReplyFolder=" & DEFAULT_FOLDER_LEARN_REPLY & vbCrLf
    s = s & "MaxReplyExamples=" & DEFAULT_MAX_REPLY_EXAMPLES & vbCrLf
    s = s & "ReplyPersona=" & DEFAULT_REPLY_PERSONA & vbCrLf
    s = s & "ScanSentItems=" & IIf(DEFAULT_SCAN_SENT_ITEMS, "True", "False") & vbCrLf
    s = s & "ScanSentDays=" & DEFAULT_SCAN_SENT_DAYS & vbCrLf
    s = s & "AutoReplyForSenders=" & DEFAULT_AUTO_REPLY_FOR_SENDERS & vbCrLf
    s = s & "; Create draft Outlook Tasks from deadlines found by the daily digest" & vbCrLf
    s = s & "EnableTaskExtraction=" & IIf(DEFAULT_ENABLE_TASK_EXTRACTION, "True", "False") & vbCrLf
    s = s & "; Inject sender history (from decision_log.txt) into LLM classification prompts" & vbCrLf
    s = s & "EnableContextEnrichment=" & IIf(DEFAULT_ENABLE_CONTEXT_ENRICHMENT, "True", "False") & vbCrLf
    s = s & vbCrLf
    s = s & "[Digest]" & vbCrLf
    s = s & "EnableDailyDigest=" & IIf(DEFAULT_ENABLE_DAILY_DIGEST, "True", "False") & vbCrLf
    s = s & "; Hour of day (0-23) after which the digest is generated once per day" & vbCrLf
    s = s & "DigestHour=" & DEFAULT_DIGEST_HOUR & vbCrLf
    s = s & "DigestMaxEmails=" & DEFAULT_DIGEST_MAX_EMAILS & vbCrLf
    s = s & "DigestSendEmail=" & IIf(DEFAULT_DIGEST_SEND_EMAIL, "True", "False") & vbCrLf
    s = s & "; Weekly LLM rule-proposal mining (proposals need approval in the Web UI)" & vbCrLf
    s = s & "EnableRuleMining=" & IIf(DEFAULT_ENABLE_RULE_MINING, "True", "False") & vbCrLf
    s = s & "LastDigestDate=" & vbCrLf
    s = s & "LastRuleMiningDate=" & vbCrLf
    s = s & vbCrLf
    s = s & "[Sync]" & vbCrLf
    s = s & "EnableCloudSync=" & IIf(DEFAULT_ENABLE_CLOUD_SYNC, "True", "False") & vbCrLf
    s = s & "; Path to shared cloud folder (OneDrive, Google Drive, etc.)" & vbCrLf
    s = s & "CloudSyncPath=" & DEFAULT_CLOUD_SYNC_PATH & vbCrLf

    WriteTextFileUTF8 filePath, s

    LogMessage "INFO", "Created default settings file: " & filePath

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "CreateDefaultSettingsFile", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' SETTINGS INI FILE - READER/WRITER
'-------------------------------------------------------------------------------

' Read a single string value from the INI file.
' Returns defaultValue if the section/key is not found.
' Reads the whole file via the BOM-aware smart reader (UTF-8 or legacy ANSI).
Public Function ReadINISetting(ByVal section As String, ByVal key As String, ByVal defaultValue As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.ReadINISetting"

    Dim filePath As String
    Dim content As String
    Dim lines As Variant
    Dim line As String
    Dim currentSection As String
    Dim eqPos As Long
    Dim i As Long

    ReadINISetting = defaultValue

    filePath = GetSettingsFilePath()
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then GoTo PROC_EXIT

    lines = SplitLines(content)
    currentSection = ""

    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))

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
                    GoTo PROC_EXIT
                End If
            End If
        End If

NextLine:
    Next i

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "ReadINISetting", Err.Number, Err.Description, section & "." & key
    ReadINISetting = defaultValue
    Resume PROC_EXIT
End Function

' Read a boolean value from the INI file
Public Function ReadINIBool(ByVal section As String, ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim raw As String
    raw = ReadINISetting(section, key, IIf(defaultValue, "True", "False"))
    ReadINIBool = (LCase(Trim(raw)) = "true" Or Trim(raw) = "1")
End Function

' Read an integer value from the INI file.
' Returns Long (the old Integer return overflowed on values > 32,767,
' e.g. ReplyMaxTokens=40000, killing Outlook startup mid-LoadAllSettings).
Public Function ReadINIInt(ByVal section As String, ByVal key As String, ByVal defaultValue As Long) As Long
    Dim raw As String
    raw = ReadINISetting(section, key, CStr(defaultValue))
    On Error GoTo Fallback
    If IsNumeric(Trim(raw)) Then
        ReadINIInt = CLng(Trim(raw))
    Else
        ReadINIInt = defaultValue
    End If
    Exit Function
Fallback:
    ReadINIInt = defaultValue
End Function

' Read a double value from the INI file.
' Values are written with "." decimal (FormatNumberForINI), so parse with the
' locale-invariant Val() instead of CDbl — on a comma-decimal locale
' CDbl("0.60") returns 60 and would break ConfidenceThreshold/Temperature.
Public Function ReadINIDouble(ByVal section As String, ByVal key As String, ByVal defaultValue As Double) As Double
    Dim raw As String
    raw = Trim(ReadINISetting(section, key, ""))

    If Len(raw) = 0 Then
        ReadINIDouble = defaultValue
        Exit Function
    End If

    ' Val is locale-invariant (always treats "." as the decimal separator)
    ' but returns 0 for non-numeric text — verify the first char is number-like
    If InStr(1, "0123456789.-+", Left(raw, 1)) > 0 Then
        ReadINIDouble = Val(raw)
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

    ' Strip CR/LF from the value — a value containing a newline would inject
    ' arbitrary lines (even new sections) into the INI structure
    value = Replace(Replace(value, vbCr, " "), vbLf, " ")

    ' Read all existing lines (BOM-aware: UTF-8 or legacy ANSI)
    If fso.FileExists(filePath) Then
        Dim rawLines As Variant
        rawLines = SplitLines(ReadTextFileSmart(filePath))
        For i = LBound(rawLines) To UBound(rawLines)
            ' Skip a single trailing empty artifact from the final CRLF
            If i = UBound(rawLines) And Len(rawLines(i)) = 0 Then Exit For
            allLines.Add CStr(rawLines(i))
        Next i
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
    ' Write all lines back as UTF-8 with BOM
    On Error GoTo WriteError
    Dim sb As String
    For i = 1 To allLines.Count
        sb = sb & allLines(i) & vbCrLf
    Next i
    WriteTextFileUTF8 filePath, sb
    Set fso = Nothing
    Exit Sub

WriteError:
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
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.LoadLearnedSenders"

    Dim filePath As String
    Dim lines As Variant
    Dim line As String
    Dim parts() As String
    Dim i As Long

    If learnedSendersCacheLoaded And Not forceReload Then GoTo PROC_EXIT

    ' Initialize or clear the cache
    Set learnedSendersCache = CreateObject("Scripting.Dictionary")
    learnedSendersCache.CompareMode = 1  ' vbTextCompare (case-insensitive keys)

    filePath = GetLearnedSendersFilePath()

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then
        ' No file yet - cache is empty but loaded
        learnedSendersCacheLoaded = True
        LogMessage "INFO", "No learned senders file found (will be created on first learn)"
        GoTo PROC_EXIT
    End If

    Dim email As String
    Dim action As String

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))

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
    Next i

    learnedSendersCacheLoaded = True
    LogMessage "INFO", "Loaded " & learnedSendersCache.Count & " learned sender rules"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "LoadLearnedSenders", Err.Number, Err.Description
    learnedSendersCacheLoaded = True  ' Prevent retry loops on a locked/corrupt file
    Resume PROC_EXIT
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
    Dim lowerEmail As String
    Dim timestamp As String

    lowerEmail = LCase(Trim(email))

    ' Skip @-less addresses (unresolved Exchange /O=... DN). A rule keyed to a
    ' DN only matches when SMTP resolution fails again — an intermittent,
    ' confusing rule. The import macros already skip these; now writes do too.
    If InStr(1, lowerEmail, "@") = 0 Then
        LogMessage "WARN", "RecordLearnedSender: skipped non-SMTP address (unresolved Exchange?): " & lowerEmail
        Exit Sub
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

    ' Append to file (UTF-8; migrates legacy ANSI file on first write)
    filePath = GetLearnedSendersFilePath()
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    AppendLineUTF8 filePath, lowerEmail & "|" & action & "|" & timestamp

    LogMessage "INFO", "LEARNED " & action & " recorded for: " & lowerEmail
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

' Force reload learned senders from file — headless core (no dialogs)
Public Function ReloadLearnedSendersCore() As String
    LoadLearnedSenders True  ' forceReload = True
    LoadLearnedSubjects True
    ReloadLearnedSendersCore = "Learned rules reloaded. Senders: " & GetLearnedSendersCount() & _
                               ", Subjects: " & GetLearnedSubjectsCount()
End Function

' Force reload learned senders from file and show count (interactive wrapper)
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
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeduplicateLearnedSenders"

    Dim filePath As String
    filePath = GetLearnedSendersFilePath()

    DeduplicatePipeFile filePath, True, "DeduplicateLearnedSenders"

    ' Reload cache from the clean file
    LoadLearnedSenders True

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "DeduplicateLearnedSenders", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Shared dedup engine for pipe-delimited rule files (key = first field).
' Keeps the LAST entry per key, preserves first-appearance order, rewrites
' the file as UTF-8. lowerKey controls case-normalisation of the key.
Private Sub DeduplicatePipeFile(ByVal filePath As String, ByVal lowerKey As Boolean, ByVal label As String)
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeduplicatePipeFile"

    Dim lines As Variant
    Dim line As String
    Dim parts() As String
    Dim dedupDict As Object    ' key -> full line (last wins)
    Dim orderList As Object    ' key -> insertion order
    Dim ruleKey As String
    Dim linesBefore As Long
    Dim linesAfter As Long
    Dim orderIndex As Long
    Dim i As Long

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then
        LogMessage "INFO", label & ": no file to deduplicate"
        GoTo PROC_EXIT
    End If

    Set dedupDict = CreateObject("Scripting.Dictionary")
    dedupDict.CompareMode = 1  ' case-insensitive
    Set orderList = CreateObject("Scripting.Dictionary")
    orderList.CompareMode = 1
    orderIndex = 0
    linesBefore = 0

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            linesBefore = linesBefore + 1
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                ruleKey = Trim(parts(0))
                If lowerKey Then ruleKey = LCase(ruleKey)
                dedupDict(ruleKey) = line   ' last entry wins
                If Not orderList.Exists(ruleKey) Then
                    orderList(ruleKey) = orderIndex
                    orderIndex = orderIndex + 1
                End If
            End If
        End If
    Next i

    linesAfter = dedupDict.Count

    If linesAfter = linesBefore Then
        LogMessage "INFO", label & ": no duplicates found (" & linesBefore & " lines)"
        GoTo PROC_EXIT
    End If

    ' Sort by insertion order and rewrite the file as UTF-8
    Dim keys As Variant
    Dim sortedLines() As String
    Dim idx As Long
    Dim k As Variant
    keys = dedupDict.keys

    ReDim sortedLines(0 To linesAfter - 1)
    For Each k In keys
        idx = orderList(k)
        sortedLines(idx) = dedupDict(k)
    Next k

    WriteTextFileUTF8 filePath, Join(sortedLines, vbCrLf) & vbCrLf

    LogMessage "INFO", label & ": " & linesBefore & " -> " & linesAfter & " lines (" & (linesBefore - linesAfter) & " duplicates removed)"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "DeduplicatePipeFile", Err.Number, Err.Description, filePath
    Resume PROC_EXIT
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
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.LoadLearnedSubjects"

    Dim filePath As String
    Dim lines As Variant
    Dim line As String
    Dim parts() As String
    Dim i As Long

    If learnedSubjectsCacheLoaded And Not forceReload Then GoTo PROC_EXIT

    ' Initialize or clear the cache
    Set learnedSubjectsCache = CreateObject("Scripting.Dictionary")
    learnedSubjectsCache.CompareMode = 1  ' vbTextCompare (case-insensitive keys)

    filePath = GetLearnedSubjectsFilePath()

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then
        ' No file yet - cache is empty but loaded
        learnedSubjectsCacheLoaded = True
        LogMessage "INFO", "No learned subjects file found (will be created on first learn)"
        GoTo PROC_EXIT
    End If

    Dim subj As String
    Dim subjAction As String

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))

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
    Next i

    learnedSubjectsCacheLoaded = True
    LogMessage "INFO", "Loaded " & learnedSubjectsCache.Count & " learned subject rules"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "LoadLearnedSubjects", Err.Number, Err.Description
    learnedSubjectsCacheLoaded = True  ' Prevent retry loops on a locked/corrupt file
    Resume PROC_EXIT
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

' Extract a generalized pattern from a subject by stripping variable parts
' (unique codes, reference numbers, dates, ticket IDs) so that one learned
' rule matches all future emails with the same template.
'
' Examples:
'   "Funding Application Submitted For Your Information Only (A0061323)"
'     -> "Funding Application Submitted For Your Information Only"
'   "Your Order #WX-98234 Has Been Shipped"
'     -> "Your Order Has Been Shipped"
'   "[TICKET-4521] Server maintenance scheduled 2026-03-14"
'     -> "Server maintenance scheduled"
'   "Re: Invoice INV-2026-0042 Payment Confirmation"
'     -> "Re: Invoice Payment Confirmation"
Public Function ExtractSubjectPattern(ByVal subject As String) As String
    Dim re As Object
    Dim result As String

    result = subject

    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    ' 1. Remove parenthesized codes: (A0061323), (REF-123), (2026-03-14), etc.
    '    Matches ( + alphanumeric/dash/dot/space content + )
    re.Pattern = "\s*\([A-Za-z0-9][A-Za-z0-9\-\./ ]*\)"
    result = re.Replace(result, "")

    ' 2. Remove bracketed codes: [TICKET-4521], [REF:123], [ID-ABC-99], etc.
    '    but preserve org tags like [MM], [HRO], [CUS] (2-4 uppercase letters only)
    re.Pattern = "\s*\[[A-Za-z0-9][A-Za-z0-9\-\.:/ ]*[0-9][A-Za-z0-9\-\.:/ ]*\]"
    result = re.Replace(result, "")

    ' 3. Remove standalone reference codes: INV-2026-0042, WX-98234, REF12345, etc.
    '    Pattern: 1-5 letters + optional separator + 3+ digits + optional suffix
    re.Pattern = "\b[A-Za-z]{1,5}[\-:#]?\d{3,}[\-\.\w]*\b"
    result = re.Replace(result, "")

    ' 4. Remove standalone pure numbers (5+ digits): 1234567, 00987654
    re.Pattern = "\b\d{5,}\b"
    result = re.Replace(result, "")

    ' 5. Remove dates: 2026-03-14, 14/03/2026, 03.14.2026, Mar 14 2026, etc.
    re.Pattern = "\b\d{4}[\-/\.]\d{1,2}[\-/\.]\d{1,2}\b"
    result = re.Replace(result, "")
    re.Pattern = "\b\d{1,2}[\-/\.]\d{1,2}[\-/\.]\d{2,4}\b"
    result = re.Replace(result, "")
    re.Pattern = "\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{1,2},?\s+\d{4}\b"
    result = re.Replace(result, "")

    ' 6. Remove trailing/leading hyphens and colons left by stripped codes
    re.Pattern = "\s+[\-:]+\s+"
    result = re.Replace(result, " ")
    re.Pattern = "[\-:]+\s*$"
    result = re.Replace(result, "")

    ' 7. Collapse multiple spaces
    re.Pattern = "\s{2,}"
    result = re.Replace(result, " ")

    Set re = Nothing

    result = Trim(result)

    ' Guard against over-generalisation: because LookupLearnedSubject matches
    ' by SUBSTRING, a short or single-word pattern (e.g. "Admission" left over
    ' after stripping "PolyU2026") would silently DELETE every future email
    ' containing that word. Require >= 12 chars AND at least two words;
    ' otherwise fall back to the verbatim subject.
    If Len(result) < 12 Or InStr(1, Trim(result), " ") = 0 Then
        result = subject
    End If

    ExtractSubjectPattern = result
End Function

' Record a learned subject rule: extract a generalized pattern from the subject,
' append to file and update cache. Variable parts (codes, IDs, dates) are stripped
' so that one rule matches all future emails with the same template.
Public Sub RecordLearnedSubject(ByVal subject As String, ByVal action As String)
    Dim filePath As String
    Dim trimmedSubject As String
    Dim patternSubject As String
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

    ' Extract generalized pattern (strip unique codes, IDs, dates)
    patternSubject = SanitizeSubject(ExtractSubjectPattern(trimmedSubject))

    ' Ensure cache is loaded
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects

    ' Skip if this pattern is already in cache (avoid duplicate rules)
    If learnedSubjectsCache.Exists(patternSubject) Then
        LogMessage "INFO", "LEARNED SUBJECT pattern already exists: " & Left(patternSubject, 50)
        Exit Sub
    End If

    ' Update in-memory cache with the pattern (not the verbatim subject)
    learnedSubjectsCache(patternSubject) = action

    ' Append to file (UTF-8; migrates legacy ANSI file on first write)
    filePath = GetLearnedSubjectsFilePath()
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    AppendLineUTF8 filePath, patternSubject & "|" & action & "|" & timestamp

    If patternSubject <> trimmedSubject Then
        LogMessage "INFO", "LEARNED SUBJECT " & action & " pattern: " & Left(patternSubject, 50) & _
                   " (from: " & Left(trimmedSubject, 50) & ")"
    Else
        LogMessage "INFO", "LEARNED SUBJECT " & action & " recorded for: " & Left(trimmedSubject, 50)
    End If
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
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeduplicateLearnedSubjects"

    Dim filePath As String
    filePath = GetLearnedSubjectsFilePath()

    DeduplicatePipeFile filePath, False, "DeduplicateLearnedSubjects"

    ' Reload cache from the clean file
    LoadLearnedSubjects True

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "DeduplicateLearnedSubjects", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' CLOUD SYNC - Merge learned rules with a cloud folder (OneDrive, etc.)
'-------------------------------------------------------------------------------

' Parse the timestamp (3rd pipe-delimited field) from a learned rule line.
' Returns "1900-01-01 00:00:00" if no timestamp field is found.
Private Function ParseTimestamp(ByVal pipeLine As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.ParseTimestamp"

    ' Timestamp is always the LAST pipe-delimited field in all formats:
    '   senders:  email|action|timestamp (index 2)
    '   subjects: subject|action|timestamp (index 2)
    '   replies:  subject|from|orig_body|reply_body|timestamp (index 4)
    Dim parts() As String
    parts = Split(pipeLine, "|")
    If UBound(parts) >= 2 Then
        ParseTimestamp = Trim(parts(UBound(parts)))
        If Len(ParseTimestamp) = 0 Then ParseTimestamp = "1900-01-01 00:00:00"
    Else
        ParseTimestamp = "1900-01-01 00:00:00"
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "ParseTimestamp", Err.Number, Err.Description
    ParseTimestamp = "1900-01-01 00:00:00"
    Resume PROC_EXIT
End Function

' Count non-empty, non-comment lines in a file (BOM-aware).
Private Function CountFileLines(ByVal filePath As String) As Long
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.CountFileLines"

    Dim lines As Variant
    Dim line As String
    Dim cnt As Long
    Dim i As Long

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then
        CountFileLines = 0
        GoTo PROC_EXIT
    End If

    cnt = 0
    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 0 And Left(line, 1) <> "#" Then cnt = cnt + 1
    Next i

    CountFileLines = cnt

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "CountFileLines", Err.Number, Err.Description
    CountFileLines = 0
    Resume PROC_EXIT
End Function

' Load a pipe-delimited rule file into a Dictionary (last entry per key wins).
' Key: first field (lowercased for sender files) or subject|from for replies.
Private Function LoadRuleFileDict(ByVal filePath As String, ByVal isSendersFile As Boolean, _
                                  ByVal isRepliesFile As Boolean) As Object
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.LoadRuleFileDict"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1  ' case-insensitive
    Set LoadRuleFileDict = dict

    Dim lines As Variant
    Dim line As String
    Dim parts() As String
    Dim ruleKey As String
    Dim i As Long

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then GoTo PROC_EXIT

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                If isSendersFile Then
                    ruleKey = LCase(Trim(parts(0)))
                ElseIf isRepliesFile Then
                    ' Composite key: subject|from (first two fields)
                    ruleKey = Trim(parts(0)) & "|" & Trim(parts(1))
                Else
                    ruleKey = Trim(parts(0))
                End If
                dict(ruleKey) = line
            End If
        End If
    Next i

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "LoadRuleFileDict", Err.Number, Err.Description, filePath
    Resume PROC_EXIT
End Function

' Sync a single learned rules file between local and cloud.
' Returns a human-readable summary string describing what happened.
Private Function SyncOneFile(ByVal localPath As String, ByVal cloudFolder As String, ByVal fileName As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.SyncOneFile"

    Dim fso As Object
    Dim cloudPath As String
    Dim dictLocal As Object
    Dim dictCloud As Object
    Dim dictMerged As Object
    Dim orderList As Object
    Dim orderIndex As Long
    Dim localCount As Long
    Dim cloudCount As Long
    Dim mergedCount As Long
    Dim newFromCloud As Long
    Dim updatedFromCloud As Long
    Dim newFromLocal As Long
    Dim k As Variant
    Dim keys As Variant
    Dim sortedLines() As String
    Dim idx As Long
    Dim isSendersFile As Boolean
    Dim isRepliesFile As Boolean
    Dim tsLocal As String
    Dim tsCloud As String
    Dim errNum As Long
    Dim errDesc As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    cloudPath = cloudFolder & "\" & fileName
    isSendersFile = (LCase(fileName) = LCase(LEARNED_SENDERS_FILE))
    isRepliesFile = (LCase(fileName) = LCase(LEARNED_REPLIES_FILE))

    ' Create cloud folder if it doesn't exist
    If Not fso.FolderExists(cloudFolder) Then
        fso.CreateFolder cloudFolder
        LogMessage "INFO", "SyncOneFile: created cloud folder: " & cloudFolder
    End If

    Dim localExists As Boolean
    Dim cloudExists As Boolean
    localExists = fso.FileExists(localPath)
    cloudExists = fso.FileExists(cloudPath)

    ' Case 1: neither file exists
    If Not localExists And Not cloudExists Then
        SyncOneFile = fileName & ": no files to sync"
        GoTo PROC_EXIT
    End If

    ' Case 2: local only - upload to cloud
    If localExists And Not cloudExists Then
        fso.CopyFile localPath, cloudPath, True
        localCount = CountFileLines(localPath)
        SyncOneFile = fileName & ": uploaded " & localCount & " rules to cloud"
        LogMessage "INFO", "SyncOneFile: uploaded " & localPath & " -> " & cloudPath & " (" & localCount & " rules)"
        GoTo PROC_EXIT
    End If

    ' Case 3: cloud only - download to local
    If Not localExists And cloudExists Then
        fso.CopyFile cloudPath, localPath, True
        cloudCount = CountFileLines(cloudPath)
        SyncOneFile = fileName & ": downloaded " & cloudCount & " rules from cloud"
        LogMessage "INFO", "SyncOneFile: downloaded " & cloudPath & " -> " & localPath & " (" & cloudCount & " rules)"
        GoTo PROC_EXIT
    End If

    ' Case 4: both exist - merge with timestamp-based conflict resolution
    ' (BOM-aware reads: files may be legacy ANSI or v3.1 UTF-8)
    Set dictLocal = LoadRuleFileDict(localPath, isSendersFile, isRepliesFile)
    localCount = dictLocal.Count

    Set dictCloud = LoadRuleFileDict(cloudPath, isSendersFile, isRepliesFile)
    cloudCount = dictCloud.Count

    ' Build merged dictionary: start with local, merge cloud entries
    Set dictMerged = CreateObject("Scripting.Dictionary")
    dictMerged.CompareMode = 1
    Set orderList = CreateObject("Scripting.Dictionary")
    orderList.CompareMode = 1
    orderIndex = 0
    newFromCloud = 0
    updatedFromCloud = 0
    newFromLocal = 0

    ' Add all local entries first
    keys = dictLocal.keys
    For Each k In keys
        dictMerged(k) = dictLocal(k)
        orderList(k) = orderIndex
        orderIndex = orderIndex + 1
    Next k

    ' Merge cloud entries (timestamp wins for conflicts)
    keys = dictCloud.keys
    For Each k In keys
        If Not dictMerged.Exists(k) Then
            ' New rule from cloud
            dictMerged(k) = dictCloud(k)
            orderList(k) = orderIndex
            orderIndex = orderIndex + 1
            newFromCloud = newFromCloud + 1
        Else
            ' Key exists in both - compare timestamps, keep later one
            tsLocal = ParseTimestamp(dictMerged(k))
            tsCloud = ParseTimestamp(dictCloud(k))
            ' String comparison works because timestamps are always YYYY-MM-DD HH:MM:SS (ISO sortable)
            If tsCloud > tsLocal Then
                dictMerged(k) = dictCloud(k)
                updatedFromCloud = updatedFromCloud + 1
            End If
        End If
    Next k

    ' Count rules that were only in local (new to cloud)
    keys = dictLocal.keys
    For Each k In keys
        If Not dictCloud.Exists(k) Then
            newFromLocal = newFromLocal + 1
        End If
    Next k

    mergedCount = dictMerged.Count

    ' Build sorted output array (preserving insertion order)
    ReDim sortedLines(0 To mergedCount - 1)
    keys = dictMerged.keys
    For Each k In keys
        idx = orderList(k)
        sortedLines(idx) = dictMerged(k)
    Next k

    ' Write merged result to local and cloud files (UTF-8 with BOM)
    Dim mergedContent As String
    mergedContent = Join(sortedLines, vbCrLf) & vbCrLf
    WriteTextFileUTF8 localPath, mergedContent
    WriteTextFileUTF8 cloudPath, mergedContent

    Set dictLocal = Nothing
    Set dictCloud = Nothing
    Set dictMerged = Nothing
    Set orderList = Nothing
    Set fso = Nothing

    SyncOneFile = fileName & ": merged " & localCount & " local + " & cloudCount & " cloud -> " & mergedCount & " unique" & _
                  " (" & newFromCloud & " new from cloud, " & updatedFromCloud & " updated, " & newFromLocal & " new to cloud)"

    LogMessage "INFO", "SyncOneFile: " & SyncOneFile

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    errNum = Err.Number
    errDesc = Err.Description
    LogError "Utilities", "SyncOneFile", errNum, errDesc
    Set fso = Nothing
    Set dictLocal = Nothing
    Set dictCloud = Nothing
    Set dictMerged = Nothing
    Set orderList = Nothing
    SyncOneFile = fileName & ": ERROR - " & errDesc
    Resume PROC_EXIT
End Function

' Headless sync core shared by the interactive macro, startup/quit auto-sync,
' and the Web UI bridge. Returns a multi-line summary; "SKIPPED: ..." when
' sync is disabled/unavailable, "ERROR: ..." on failure.
Public Function SyncLearnedRulesCore() As String
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.SyncLearnedRulesCore"

    SyncLearnedRulesCore = ""

    If Not RuntimeEnableCloudSync Then
        SyncLearnedRulesCore = "SKIPPED: Cloud sync is disabled. Set EnableCloudSync=True under [Sync] in settings.ini."
        GoTo PROC_EXIT
    End If

    If Len(Trim(RuntimeCloudSyncPath)) = 0 Then
        SyncLearnedRulesCore = "SKIPPED: CloudSyncPath is not configured under [Sync] in settings.ini."
        GoTo PROC_EXIT
    End If

    ' Skip gracefully if the cloud folder is unavailable (offline, paused, unmounted)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(RuntimeCloudSyncPath) Then
        Set fso = Nothing
        SyncLearnedRulesCore = "SKIPPED: Cloud path not available (" & RuntimeCloudSyncPath & "). Using local rules only."
        GoTo PROC_EXIT
    End If
    Set fso = Nothing

    Dim cloudFolder As String
    cloudFolder = RuntimeCloudSyncPath & "\" & LEARNED_DATA_FOLDER

    Dim resultSenders As String
    Dim resultSubjects As String
    Dim resultReplies As String

    resultSenders = SyncOneFile(GetLearnedSendersFilePath(), cloudFolder, LEARNED_SENDERS_FILE)
    resultSubjects = SyncOneFile(GetLearnedSubjectsFilePath(), cloudFolder, LEARNED_SUBJECTS_FILE)
    resultReplies = SyncOneFile(GetLearnedRepliesFilePath(), cloudFolder, LEARNED_REPLIES_FILE)

    ' Refresh in-memory caches after sync
    LoadLearnedSenders True
    LoadLearnedSubjects True

    SyncLearnedRulesCore = resultSenders & vbCrLf & resultSubjects & vbCrLf & resultReplies & vbCrLf & _
                           "Cloud folder: " & cloudFolder

    LogMessage "INFO", "SyncLearnedRules complete. " & resultSenders & " | " & resultSubjects & " | " & resultReplies

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "SyncLearnedRulesCore", Err.Number, Err.Description
    SyncLearnedRulesCore = "ERROR: Cloud sync failed: " & Err.Description
    Resume PROC_EXIT
End Function

' Sync learned senders, subjects, and replies with a cloud folder (interactive).
' Merges rules bidirectionally using timestamp-based conflict resolution.
Public Sub SyncLearnedRules()
    Dim summary As String
    summary = SyncLearnedRulesCore()

    If Left(summary, 8) = "SKIPPED:" Then
        MsgBox Mid(summary, 10), vbExclamation, "Cloud Sync"
    ElseIf Left(summary, 6) = "ERROR:" Then
        MsgBox summary, vbCritical, "Cloud Sync"
    Else
        MsgBox "Cloud sync complete:" & vbCrLf & vbCrLf & summary, vbInformation, "Cloud Sync"
    End If
End Sub

' Silent auto-sync: called at Outlook startup and quit.
' No MsgBox prompts. Gracefully skips if cloud sync is disabled or OneDrive is unavailable.
' Returns True if sync was performed, False if skipped/failed.
Public Function SyncLearnedRulesAuto() As Boolean
    Dim summary As String
    summary = SyncLearnedRulesCore()

    If Left(summary, 8) = "SKIPPED:" Then
        LogMessage "INFO", "SyncLearnedRulesAuto: " & Mid(summary, 10)
        SyncLearnedRulesAuto = False
    ElseIf Left(summary, 6) = "ERROR:" Then
        LogMessage "WARN", "SyncLearnedRulesAuto failed, continuing with local rules"
        SyncLearnedRulesAuto = False
    Else
        SyncLearnedRulesAuto = True
    End If
End Function

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
Public Sub DeleteLearnedSenderRule(ByVal email As String)
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeleteLearnedSenderRule"

    Dim lowerEmail As String
    lowerEmail = LCase(Trim(email))

    ' Remove from in-memory cache
    If Not learnedSendersCacheLoaded Then LoadLearnedSenders
    If learnedSendersCache.Exists(lowerEmail) Then
        learnedSendersCache.Remove lowerEmail
    End If

    RemovePipeFileEntries GetLearnedSendersFilePath(), lowerEmail
    LogMessage "INFO", "Deleted learned sender rule for: " & lowerEmail

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "DeleteLearnedSenderRule", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Rewrite a pipe-delimited rule file without the entries whose first field
' matches keyValue (case-insensitive). Comments/blank/malformed lines are kept.
Private Sub RemovePipeFileEntries(ByVal filePath As String, ByVal keyValue As String)
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.RemovePipeFileEntries"

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then GoTo PROC_EXIT

    Dim lines As Variant
    Dim line As String
    Dim parts() As String
    Dim kept As String
    Dim i As Long

    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = lines(i)
        ' Skip the single trailing empty artifact of the final newline
        If i = UBound(lines) And Len(line) = 0 Then Exit For

        If Len(Trim(line)) > 0 And Left(Trim(line), 1) <> "#" Then
            parts = Split(Trim(line), "|")
            If UBound(parts) >= 1 Then
                If LCase(Trim(parts(0))) <> LCase(keyValue) Then
                    kept = kept & line & vbCrLf
                End If
            Else
                kept = kept & line & vbCrLf  ' Keep malformed lines as-is
            End If
        Else
            kept = kept & line & vbCrLf  ' Keep comments and blank lines
        End If
    Next i

    WriteTextFileUTF8 filePath, kept

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "RemovePipeFileEntries", Err.Number, Err.Description, filePath
    Resume PROC_EXIT
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

    AppendLineUTF8 filePath, safeSubject & "|" & safeFrom & "|" & safeOrigBody & "|" & safeReplyBody & "|" & timestamp

    LogMessage "INFO", "Learned reply recorded for: " & Left(safeSubject, 50)
End Sub

' Load the most recent N reply pairs from learned_replies.txt.
' Returns a Collection of pipe-delimited strings (raw lines).
Public Function LoadRecentReplyPairs(ByVal maxPairs As Long) As Collection
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.LoadRecentReplyPairs"

    Dim result As Collection
    Set result = New Collection
    Set LoadRecentReplyPairs = result

    Dim filePath As String
    filePath = GetLearnedRepliesFilePath()

    Dim content As String
    content = ReadTextFileSmart(filePath)
    If Len(content) = 0 Then GoTo PROC_EXIT

    ' Read all lines, keep last maxPairs
    Dim allLines As Collection
    Set allLines = New Collection

    Dim lines As Variant
    Dim line As String
    Dim i As Long
    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        line = Trim(lines(i))
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            allLines.Add line
        End If
    Next i

    ' Return the last maxPairs lines
    Dim startIdx As Long
    startIdx = allLines.Count - maxPairs + 1
    If startIdx < 1 Then startIdx = 1

    For i = startIdx To allLines.Count
        result.Add allLines(i)
    Next i

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "Utilities", "LoadRecentReplyPairs", Err.Number, Err.Description
    Resume PROC_EXIT
End Function

' Delete a single learned subject rule from both cache and file.
Public Sub DeleteLearnedSubjectRule(ByVal subject As String)
    On Error GoTo PROC_ERR
    PushCallStack "Utilities.DeleteLearnedSubjectRule"

    Dim targetSubject As String
    targetSubject = SanitizeSubject(subject)

    ' Remove from in-memory cache
    If Not learnedSubjectsCacheLoaded Then LoadLearnedSubjects
    If learnedSubjectsCache.Exists(targetSubject) Then
        learnedSubjectsCache.Remove targetSubject
    End If

    RemovePipeFileEntries GetLearnedSubjectsFilePath(), targetSubject
    LogMessage "INFO", "Deleted learned subject rule for: " & Left(targetSubject, 50)

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "Utilities", "DeleteLearnedSubjectRule", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

