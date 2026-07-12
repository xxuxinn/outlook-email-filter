'===============================================================================
' AgentMemory.bas - Agent Decision Memory v3.1
'===============================================================================
' Persistent memory of what the filter decided and why, powering:
'   - decision_log.txt   : every classification (sender|subject|source|action|confidence)
'   - Sender history     : "45 of 47 emails from this sender were deleted" —
'                          injected into LLM classification prompts (the single
'                          biggest accuracy lever: evidence instead of guessing)
'   - Replied-to set     : "you have replied to this sender before" from
'                          learned_replies.txt
'   - llm_corrections.txt: user reversals of LLM decisions, fed back into the
'                          classify prompt as few-shot corrections
'
' All files are pipe-delimited, append-only, UTF-8 (see Utilities.AppendLineUTF8).
'===============================================================================

Option Explicit

' Sender stats cache: senderEmail -> "keep|delete|review|lastSource|lastAction"
Private senderStatsCache As Object
Private senderStatsLoaded As Boolean

' Replied-to cache: lowercased recipient names from learned_replies.txt
Private repliedToCache As Object
Private repliedToLoaded As Boolean

' Cap on how many decision-log lines are parsed into the stats cache
Private Const MAX_DECISION_LINES As Long = 5000

'-------------------------------------------------------------------------------
' FILE PATHS
'-------------------------------------------------------------------------------

Public Function GetDecisionLogPath() As String
    GetDecisionLogPath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & DECISION_LOG_FILE
End Function

Public Function GetCorrectionsFilePath() As String
    GetCorrectionsFilePath = Environ("APPDATA") & "\" & LEARNED_DATA_FOLDER & "\" & CORRECTIONS_FILE
End Function

'-------------------------------------------------------------------------------
' DECISION LOG
'-------------------------------------------------------------------------------

' Record one classification decision.
'   source: "LEARNED_SENDER" | "LEARNED_SUBJECT" | "RULE1_PROTECTED" ... | "LLM" | "DEFAULT"
'   action: "KEEP" | "DELETE" | "MOVE_II" | "REVIEW"
'   confidence: 1 for deterministic rules, the model's 0-1 estimate for LLM
Public Sub RecordDecision(ByVal senderEmail As String, ByVal subject As String, _
                          ByVal source As String, ByVal action As String, _
                          ByVal confidence As Double)
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.RecordDecision"

    Dim safeSender As String
    Dim safeSubject As String
    Dim confStr As String

    safeSender = LCase(Trim(Replace(senderEmail, "|", " ")))
    safeSubject = SanitizeSubject(Truncate(subject, 80))
    confStr = Replace(Format(confidence, "0.00"), ",", ".")

    ' The log grows with every classified email — rotate to .old at the cap
    ' (sender stats only parse the most recent MAX_DECISION_LINES anyway)
    RotateFileIfOversize GetDecisionLogPath(), DECISION_LOG_MAX_BYTES

    AppendLineUTF8 GetDecisionLogPath(), _
        Format(Now, "yyyy-mm-dd hh:nn:ss") & "|" & safeSender & "|" & safeSubject & "|" & _
        source & "|" & action & "|" & confStr

    ' Keep the in-memory stats current without re-reading the file
    If senderStatsLoaded Then
        UpdateStatsEntry safeSender, source, action
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "AgentMemory", "RecordDecision", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Load the sender stats cache from decision_log.txt (last MAX_DECISION_LINES)
Public Sub LoadSenderStats(Optional ByVal forceReload As Boolean = False)
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.LoadSenderStats"

    If senderStatsLoaded And Not forceReload Then GoTo PROC_EXIT

    Set senderStatsCache = CreateObject("Scripting.Dictionary")
    senderStatsCache.CompareMode = 1

    Dim content As String
    content = ReadTextFileSmart(GetDecisionLogPath())
    senderStatsLoaded = True
    If Len(content) = 0 Then GoTo PROC_EXIT

    Dim lines As Variant
    Dim i As Long
    Dim startIdx As Long
    Dim parts() As String

    lines = SplitLines(content)
    startIdx = LBound(lines)
    If UBound(lines) - startIdx + 1 > MAX_DECISION_LINES Then
        startIdx = UBound(lines) - MAX_DECISION_LINES + 1
    End If

    For i = startIdx To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            ' timestamp|sender|subject|source|action|confidence
            If UBound(parts) >= 4 Then
                UpdateStatsEntry LCase(Trim(parts(1))), Trim(parts(3)), UCase(Trim(parts(4)))
            End If
        End If
    Next i

    LogMessage "INFO", "AgentMemory: sender stats loaded for " & senderStatsCache.Count & " senders"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "AgentMemory", "LoadSenderStats", Err.Number, Err.Description
    senderStatsLoaded = True  ' Prevent retry loops on a locked/corrupt file
    Resume PROC_EXIT
End Sub

' Increment one sender's counters. Entry format: "keep|delete|review|lastSource|lastAction"
Private Sub UpdateStatsEntry(ByVal senderKey As String, ByVal source As String, ByVal action As String)
    If Len(senderKey) = 0 Then Exit Sub
    If senderStatsCache Is Nothing Then Exit Sub

    Dim keepN As Long, delN As Long, revN As Long
    Dim parts() As String

    If senderStatsCache.Exists(senderKey) Then
        parts = Split(senderStatsCache(senderKey), "|")
        If UBound(parts) >= 2 Then
            keepN = CLng(parts(0))
            delN = CLng(parts(1))
            revN = CLng(parts(2))
        End If
    End If

    Select Case UCase(action)
        Case "KEEP", "MOVE_II": keepN = keepN + 1
        Case "DELETE": delN = delN + 1
        Case Else: revN = revN + 1
    End Select

    senderStatsCache(senderKey) = keepN & "|" & delN & "|" & revN & "|" & source & "|" & UCase(action)
End Sub

' Return "source|action" of the most recent recorded decision for a sender,
' or "" if none. Used to detect user reversals of LLM decisions.
Public Function GetLastDecisionForSender(ByVal senderEmail As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.GetLastDecisionForSender"

    GetLastDecisionForSender = ""
    If Not senderStatsLoaded Then LoadSenderStats

    Dim key As String
    key = LCase(Trim(senderEmail))

    If senderStatsCache.Exists(key) Then
        Dim parts() As String
        parts = Split(senderStatsCache(key), "|")
        If UBound(parts) >= 4 Then
            GetLastDecisionForSender = parts(3) & "|" & parts(4)
        End If
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "AgentMemory", "GetLastDecisionForSender", Err.Number, Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' SENDER CONTEXT (prompt enrichment)
'-------------------------------------------------------------------------------

' Build a short evidence block about this sender for the classification prompt.
' Returns "" when there is nothing useful to say (new sender, no history).
Public Function GetSenderContext(ByVal senderEmail As String, ByVal senderName As String) As String
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.GetSenderContext"

    GetSenderContext = ""
    If Not RuntimeEnableContextEnrichment Then GoTo PROC_EXIT

    If Not senderStatsLoaded Then LoadSenderStats

    Dim context As String
    Dim key As String
    key = LCase(Trim(senderEmail))

    If senderStatsCache.Exists(key) Then
        Dim parts() As String
        parts = Split(senderStatsCache(key), "|")
        If UBound(parts) >= 2 Then
            Dim total As Long
            total = CLng(parts(0)) + CLng(parts(1)) + CLng(parts(2))
            If total > 0 Then
                context = "Sender history: " & total & " previous email(s) from this sender — " & _
                          parts(0) & " kept, " & parts(1) & " deleted, " & parts(2) & " sent to review."
            End If
        End If
    End If

    If HasRepliedTo(senderName) Then
        If Len(context) > 0 Then context = context & " "
        context = context & "The user has personally replied to this sender before (strong KEEP signal)."
    End If

    GetSenderContext = context

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "AgentMemory", "GetSenderContext", Err.Number, Err.Description
    GetSenderContext = ""
    Resume PROC_EXIT
End Function

' True if the sender's display name appears as a recipient in learned_replies.txt
' (i.e. the user has written a reply to them before).
Public Function HasRepliedTo(ByVal senderName As String) As Boolean
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.HasRepliedTo"

    HasRepliedTo = False
    If Len(Trim(senderName)) = 0 Then GoTo PROC_EXIT

    If Not repliedToLoaded Then LoadRepliedToCache

    If Not repliedToCache Is Nothing Then
        HasRepliedTo = repliedToCache.Exists(LCase(Trim(senderName)))
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "AgentMemory", "HasRepliedTo", Err.Number, Err.Description
    HasRepliedTo = False
    Resume PROC_EXIT
End Function

Private Sub LoadRepliedToCache()
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.LoadRepliedToCache"

    Set repliedToCache = CreateObject("Scripting.Dictionary")
    repliedToCache.CompareMode = 1
    repliedToLoaded = True

    Dim content As String
    content = ReadTextFileSmart(GetLearnedRepliesFilePath())
    If Len(content) = 0 Then GoTo PROC_EXIT

    ' learned_replies format: subject|from|orig_body|reply_body|timestamp
    Dim lines As Variant
    Dim parts() As String
    Dim i As Long
    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 And Left(Trim(lines(i)), 1) <> "#" Then
            parts = Split(lines(i), "|")
            If UBound(parts) >= 1 Then
                If Len(Trim(parts(1))) > 0 Then
                    repliedToCache(LCase(Trim(parts(1)))) = True
                End If
            End If
        End If
    Next i

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "AgentMemory", "LoadRepliedToCache", Err.Number, Err.Description
    repliedToLoaded = True
    Resume PROC_EXIT
End Sub

' Invalidate the replied-to cache (called after new reply pairs are learned)
Public Sub InvalidateRepliedToCache()
    repliedToLoaded = False
End Sub

'-------------------------------------------------------------------------------
' LLM CORRECTIONS (user reversals -> few-shot examples)
'-------------------------------------------------------------------------------

' Record that the user reversed an LLM decision (e.g. dragged a deleted email
' into LearnKeep). These become few-shot corrections in future classify prompts.
Public Sub RecordCorrection(ByVal senderEmail As String, ByVal subject As String, _
                            ByVal wrongAction As String, ByVal correctAction As String)
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.RecordCorrection"

    AppendLineUTF8 GetCorrectionsFilePath(), _
        Format(Now, "yyyy-mm-dd hh:nn:ss") & "|" & LCase(Trim(Replace(senderEmail, "|", " "))) & "|" & _
        SanitizeSubject(Truncate(subject, 80)) & "|" & UCase(wrongAction) & "|" & UCase(correctAction)

    LogMessage "INFO", "AgentMemory: correction recorded — LLM said " & wrongAction & _
               ", user says " & correctAction & " for " & senderEmail

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "AgentMemory", "RecordCorrection", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Build a few-shot block of the most recent corrections for the classify prompt.
' Returns "" when there are no corrections yet.
Public Function GetRecentCorrectionsBlock(ByVal maxItems As Long) As String
    On Error GoTo PROC_ERR
    PushCallStack "AgentMemory.GetRecentCorrectionsBlock"

    GetRecentCorrectionsBlock = ""

    Dim content As String
    content = ReadTextFileSmart(GetCorrectionsFilePath())
    If Len(content) = 0 Then GoTo PROC_EXIT

    Dim lines As Variant
    Dim kept As Collection
    Dim parts() As String
    Dim i As Long

    Set kept = New Collection
    lines = SplitLines(content)
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            parts = Split(lines(i), "|")
            ' timestamp|sender|subject|wrong|correct
            If UBound(parts) >= 4 Then kept.Add lines(i)
        End If
    Next i

    If kept.Count = 0 Then GoTo PROC_EXIT

    Dim startIdx As Long
    startIdx = kept.Count - maxItems + 1
    If startIdx < 1 Then startIdx = 1

    Dim block As String
    block = "Past mistakes the user has corrected (do not repeat them):" & vbCrLf
    For i = startIdx To kept.Count
        parts = Split(kept(i), "|")
        block = block & "- From " & parts(1) & ", subject """ & parts(2) & """: you said " & _
                parts(3) & ", the correct answer was " & parts(4) & "." & vbCrLf
    Next i

    GetRecentCorrectionsBlock = block

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "AgentMemory", "GetRecentCorrectionsBlock", Err.Number, Err.Description
    GetRecentCorrectionsBlock = ""
    Resume PROC_EXIT
End Function
