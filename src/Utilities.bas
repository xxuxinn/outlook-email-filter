'===============================================================================
' Utilities.bas - Helper Functions for Email Filter v2.0
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
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' LEARNED SENDERS CACHE (Self-Improving Filter)
'-------------------------------------------------------------------------------
Private learnedSendersCache As Object  ' Scripting.Dictionary (email -> "KEEP"|"DELETE")
Private learnedSendersCacheLoaded As Boolean

'-------------------------------------------------------------------------------
' LEARNED SUBJECTS CACHE (Self-Improving Filter - Subject Rules)
'-------------------------------------------------------------------------------
Private learnedSubjectsCache As Object  ' Scripting.Dictionary (subject -> "DELETE")
Private learnedSubjectsCacheLoaded As Boolean

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
    RuntimeLLMEndpoint = ReadINISetting("LLM", "Endpoint", DEFAULT_AZURE_OPENAI_ENDPOINT)
    RuntimeAPIKeyMethod = ReadINISetting("LLM", "APIKeyMethod", DEFAULT_API_KEY_METHOD)
    RuntimeAPIKeyEnvVar = ReadINISetting("LLM", "APIKeyEnvVar", DEFAULT_API_KEY_ENV_VAR)
    RuntimeAPIKeyHardcoded = ReadINISetting("LLM", "APIKeyHardcoded", DEFAULT_API_KEY_HARDCODED)
    RuntimeLLMMaxTokens = ReadINIInt("LLM", "MaxTokens", DEFAULT_LLM_MAX_TOKENS)
    RuntimeLLMTemperature = ReadINIDouble("LLM", "Temperature", DEFAULT_LLM_TEMPERATURE)
    RuntimeLLMSystemPrompt = ReadINISetting("LLM", "SystemPrompt", DEFAULT_LLM_SYSTEM_PROMPT)

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
    ts.WriteLine "Endpoint=" & DEFAULT_AZURE_OPENAI_ENDPOINT
    ts.WriteLine "APIKeyMethod=" & DEFAULT_API_KEY_METHOD
    ts.WriteLine "APIKeyEnvVar=" & DEFAULT_API_KEY_ENV_VAR
    ts.WriteLine "APIKeyHardcoded=" & DEFAULT_API_KEY_HARDCODED
    ts.WriteLine "MaxTokens=" & DEFAULT_LLM_MAX_TOKENS
    ts.WriteLine "Temperature=" & DEFAULT_LLM_TEMPERATURE
    ts.WriteLine "SystemPrompt=" & DEFAULT_LLM_SYSTEM_PROMPT

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
                    ReadINISetting = Mid(line, eqPos + 1)
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
                Dim tempLines As Collection
                Set tempLines = New Collection
                Dim j As Long
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
                    Dim replacedLines As Collection
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

    On Error GoTo ErrorHandler

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
    Exit Function

ErrorHandler:
    LogMessage "ERROR", "RestoreSenderFromDeleted error: " & Err.Description
    RestoreSenderFromDeleted = restoredCount
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

    On Error GoTo ErrorHandler

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
    Exit Function

ErrorHandler:
    LogMessage "ERROR", "DeleteSenderFromInbox error: " & Err.Description
    DeleteSenderFromInbox = deletedCount
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
