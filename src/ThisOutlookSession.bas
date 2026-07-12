'===============================================================================
' ThisOutlookSession - Event Handlers v3.0
'===============================================================================
' IMPORTANT: This code goes DIRECTLY into ThisOutlookSession (built-in module).
'
' DO NOT import this file! Instead:
' 1. In VBA Editor, find "ThisOutlookSession" under Microsoft Outlook Objects
' 2. Double-click to open it
' 3. Copy everything BELOW this comment block and paste into that window
'===============================================================================

Option Explicit

' Event handler for Inbox items
Public WithEvents inboxItems As Outlook.Items

' Event handlers for self-improving learning folders
Public WithEvents learnKeepItems As Outlook.Items         ' LearnKeep folder
Public WithEvents learnDeleteItems As Outlook.Items        ' LearnDelete folder
Public WithEvents learnSubjectDeleteItems As Outlook.Items ' LearnSubjectDelete folder
Public WithEvents learnReplyItems As Outlook.Items         ' LearnReply folder (v3.0)

'-------------------------------------------------------------------------------
' APPLICATION STARTUP
'-------------------------------------------------------------------------------

' Initialize event handlers when Outlook starts
Private Sub Application_Startup()
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.Application_Startup"

    ' CRITICAL: Load all settings from settings.ini FIRST, before anything else.
    ' This populates all Runtime* variables used by every other module.
    LoadAllSettings

    ' Start the Web UI command poller (polls %APPDATA%\OutlookEmailFilter\commands\ every 2s)
    StartCommandPollerStd

    ' Get reference to Inbox items
    Set inboxItems = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Items

    ' Initialize self-improving learning folder watchers
    ' Learning folders must already exist under Inbox (not auto-created)
    If RuntimeEnableSelfImproving Then
        Dim learnKeepFolder As Outlook.Folder
        Dim learnDeleteFolder As Outlook.Folder

        On Error Resume Next
        Set learnKeepFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderLearnKeep)
        Set learnDeleteFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderLearnDelete)
        On Error GoTo PROC_ERR

        If learnKeepFolder Is Nothing Or learnDeleteFolder Is Nothing Then
            LogMessage "WARN", "Learning folders (" & RuntimeFolderLearnKeep & "/" & RuntimeFolderLearnDelete & ") not found under Inbox - sender learning disabled"
        Else
            Set learnKeepItems = learnKeepFolder.Items
            Set learnDeleteItems = learnDeleteFolder.Items
            LoadLearnedSenders
            DeduplicateLearnedSenders
            LogMessage "INFO", "Sender learning active (" & GetLearnedSendersCount() & " learned sender rules)"
        End If

        ' LearnSubjectDelete folder for subject-based DELETE learning (independent of sender folders)
        Dim learnSubjectDeleteFolder As Outlook.Folder
        On Error Resume Next
        Set learnSubjectDeleteFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderLearnSubject)
        On Error GoTo PROC_ERR

        If learnSubjectDeleteFolder Is Nothing Then
            LogMessage "WARN", "Learning folder (" & RuntimeFolderLearnSubject & ") not found under Inbox - subject learning disabled"
        Else
            Set learnSubjectDeleteItems = learnSubjectDeleteFolder.Items
            LoadLearnedSubjects
            DeduplicateLearnedSubjects
            LogMessage "INFO", "Subject learning active (" & GetLearnedSubjectsCount() & " learned subject rules)"
        End If

        ' LearnReply folder for reply-style learning (v3.0)
        Dim learnReplyFolder As Outlook.Folder
        On Error Resume Next
        Set learnReplyFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderLearnReply)
        On Error GoTo PROC_ERR

        If learnReplyFolder Is Nothing Then
            LogMessage "INFO", "LearnReply folder (" & RuntimeFolderLearnReply & ") not found under Inbox - reply learning disabled (optional)"
        Else
            Set learnReplyItems = learnReplyFolder.Items
            LogMessage "INFO", "Reply style learning active (folder: " & RuntimeFolderLearnReply & ")"
        End If
    End If

    ' Auto-sync learned rules with cloud on startup (silent, skips if unavailable)
    SyncLearnedRulesAuto

    ' Warm the sender-history cache (context enrichment for LLM classification)
    LoadSenderStats

    LogMessage "INFO", "Email Agent v" & FILTER_VERSION & " initialized"

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "Application_Startup", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' APPLICATION QUIT - Auto-sync learned rules to cloud before closing
'-------------------------------------------------------------------------------

Private Sub Application_Quit()
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.Application_Quit"

    ' Stop command poller timer before shutdown
    StopCommandPollerStd

    ' Auto-sync learned rules to cloud (silent, skips if unavailable)
    SyncLearnedRulesAuto

    LogMessage "INFO", "Email Agent v" & FILTER_VERSION & " shutting down"

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "Application_Quit", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' NEW MAIL EVENT HANDLER
'-------------------------------------------------------------------------------

' Fires when a new email arrives in the Inbox — real-time filtering.
' v3.1: routes through ExecuteAction, so ambiguous emails now get the SAME
' LLM classification as batch runs (previously they went straight to Review
' without ever consulting the LLM), and every decision lands in
' decision_log.txt. Auto-reply drafting triggers on resolved KEEPs when
' RuntimeAutoReplyOnArrival = True (drafts only — never auto-sent).
Private Sub inboxItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    ' Only process MailItems (not meetings, tasks, etc.)
    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    Dim mail As Outlook.MailItem
    Set mail = Item

    ' Classify the email
    Dim decision As String
    decision = ClassifyEmail(mail)

    ' Pre-capture before action (object becomes invalid after .Move/.Delete)
    Dim senderName As String, subject As String
    senderName = mail.SenderName
    subject = mail.subject

    ' Execute (handles LLM resolution of LLM_REVIEW + decision logging)
    Dim resolvedAction As String
    ExecuteAction mail, decision, Nothing, resolvedAction

    LogMessage "INFO", "AUTO: " & resolvedAction & " | " & senderName & " | " & subject

    ' Auto-draft reply for resolved KEEPs (mail object is still valid: KEEP
    ' means the email was neither moved nor deleted)
    If resolvedAction = "KEEP" Then
        If RuntimeAutoReplyOnArrival And RuntimeEnableAutoReply Then
            DraftAutoReply mail
        End If
    End If

    Exit Sub

ErrorHandler:
    LogError "ThisOutlookSession", "inboxItems_ItemAdd", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' SELF-IMPROVING LEARNING EVENT HANDLERS
'-------------------------------------------------------------------------------

' Fires when an email is dragged into the LearnKeep folder (learn to KEEP)
Private Sub learnKeepItems_ItemAdd(ByVal Item As Object)
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.learnKeepItems_ItemAdd"

    If Not TypeOf Item Is Outlook.MailItem Then GoTo PROC_EXIT

    Dim mail As Outlook.MailItem
    Set mail = Item

    Dim senderEmail As String
    senderEmail = GetSenderEmail(mail)

    ' Check if this sender previously had a DELETE rule (rule reversal)
    Dim previousAction As String
    previousAction = LookupLearnedSender(senderEmail)

    ' Correction capture: if the LLM most recently DELETEd/REVIEWed this sender
    ' and the user is now teaching KEEP, record it as a few-shot correction
    Dim lastDecision As String
    lastDecision = GetLastDecisionForSender(senderEmail)
    If Left(lastDecision, 4) = "LLM|" Then
        Dim llmAction As String
        llmAction = Mid(lastDecision, 5)
        If llmAction = "DELETE" Or llmAction = "REVIEW" Then
            RecordCorrection senderEmail, mail.subject, llmAction, "KEEP"
        End If
    End If

    RecordLearnedSender senderEmail, "KEEP"
    LogMessage "INFO", "LEARNED KEEP from folder " & RuntimeFolderLearnKeep & ": " & senderEmail & _
               " (" & Left(mail.senderName, 25) & ")"

    ' If reversing a DELETE rule, rescue this sender's emails from Deleted Items
    If previousAction = "DELETE" Then
        Dim restoredCount As Long
        restoredCount = RestoreSenderFromDeleted(senderEmail)
        If restoredCount > 0 Then
            LogMessage "INFO", "Rule reversal: restored " & restoredCount & " email(s) from Deleted Items for " & senderEmail
        End If
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "learnKeepItems_ItemAdd", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Fires when an email is dragged into the LearnDelete folder (learn to DELETE)
Private Sub learnDeleteItems_ItemAdd(ByVal Item As Object)
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.learnDeleteItems_ItemAdd"

    If Not TypeOf Item Is Outlook.MailItem Then GoTo PROC_EXIT

    Dim mail As Outlook.MailItem
    Set mail = Item

    Dim senderEmail As String
    senderEmail = GetSenderEmail(mail)

    ' Check if this sender previously had a KEEP rule (rule reversal)
    Dim previousAction As String
    previousAction = LookupLearnedSender(senderEmail)

    ' Correction capture: if the LLM most recently KEPT this sender and the
    ' user is now teaching DELETE, record it as a few-shot correction
    Dim lastDecision As String
    lastDecision = GetLastDecisionForSender(senderEmail)
    If lastDecision = "LLM|KEEP" Then
        RecordCorrection senderEmail, mail.subject, "KEEP", "DELETE"
    End If

    RecordLearnedSender senderEmail, "DELETE"
    LogMessage "INFO", "LEARNED DELETE from folder " & RuntimeFolderLearnDelete & ": " & senderEmail & _
               " (" & Left(mail.senderName, 25) & ")"

    ' If reversing a KEEP rule, clean up this sender's emails from Inbox
    If previousAction = "KEEP" Then
        Dim deletedCount As Long
        deletedCount = DeleteSenderFromInbox(senderEmail)
        If deletedCount > 0 Then
            LogMessage "INFO", "Rule reversal: deleted " & deletedCount & " email(s) from Inbox for " & senderEmail
        End If
    End If

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "learnDeleteItems_ItemAdd", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Fires when a sent email is dragged into the LearnReply folder (learn reply style)
' The dragged email should be a SENT email (from Sent Items) that is a reply.
' Extracts the original message / reply body pair and appends to learned_replies.txt.
Private Sub learnReplyItems_ItemAdd(ByVal Item As Object)
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.learnReplyItems_ItemAdd"

    If Not TypeOf Item Is Outlook.MailItem Then GoTo PROC_EXIT

    Dim mail As Outlook.MailItem
    Set mail = Item

    Dim subject As String
    subject = mail.Subject

    If Len(Trim(subject)) = 0 Then
        LogMessage "WARN", "learnReplyItems_ItemAdd: empty subject, skipping"
        GoTo PROC_EXIT
    End If

    ' Extract original subject (strip RE: prefix)
    Dim originalSubject As String
    Dim trimmedSubject As String
    trimmedSubject = Trim(subject)
    Do While UCase(Left(trimmedSubject, 3)) = "RE:" Or UCase(Left(trimmedSubject, 3)) = "AW:"
        trimmedSubject = Trim(Mid(trimmedSubject, 4))
    Loop
    originalSubject = trimmedSubject

    ' Extract reply body / original snippet (shared helpers in EmailAgent.bas —
    ' the previous local copies drifted from the EmailAgent versions)
    Dim myReplyText As String
    myReplyText = ExtractReplyTextFromBody(mail.Body)

    Dim originalBodySnippet As String
    originalBodySnippet = ExtractOriginalBodySnippet(mail.Body)

    ' Get the original sender (first recipient of the sent reply)
    Dim originalFrom As String
    If mail.Recipients.Count > 0 Then
        originalFrom = mail.Recipients(1).Name
    Else
        originalFrom = ""
    End If

    If Len(Trim(myReplyText)) < 10 Then
        LogMessage "WARN", "learnReplyItems_ItemAdd: reply text too short to be useful, skipping"
        GoTo PROC_EXIT
    End If

    RecordLearnedReply originalSubject, originalFrom, originalBodySnippet, myReplyText
    InvalidateRepliedToCache  ' AgentMemory: sender is now in the replied-to set
    LogMessage "INFO", "LEARNED REPLY from folder " & RuntimeFolderLearnReply & ": " & Left(originalSubject, 50)

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "learnReplyItems_ItemAdd", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Fires when an email is dragged into the LearnSubjectDelete folder (learn to DELETE by subject)
Private Sub learnSubjectDeleteItems_ItemAdd(ByVal Item As Object)
    On Error GoTo PROC_ERR
    PushCallStack "ThisOutlookSession.learnSubjectDeleteItems_ItemAdd"

    If Not TypeOf Item Is Outlook.MailItem Then GoTo PROC_EXIT

    Dim mail As Outlook.MailItem
    Set mail = Item

    Dim subject As String
    subject = mail.subject

    If Len(Trim(subject)) = 0 Then
        LogMessage "WARN", "learnSubjectDeleteItems_ItemAdd: empty subject, skipping"
        GoTo PROC_EXIT
    End If

    RecordLearnedSubject subject, "DELETE"
    LogMessage "INFO", "LEARNED SUBJECT DELETE from folder " & RuntimeFolderLearnSubject & ": " & Left(subject, 50) & _
               " (" & Left(mail.senderName, 25) & ")"

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "learnSubjectDeleteItems_ItemAdd", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' MANUAL TRIGGER (for testing)
'-------------------------------------------------------------------------------

' Manually reinitialize the event handlers
Public Sub ReinitializeFilter()
    Application_Startup
    MsgBox "Email Filter v" & FILTER_VERSION & " reinitialized.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' ENABLE/DISABLE REAL-TIME FILTERING
'-------------------------------------------------------------------------------

' Disable real-time filtering (disconnect event handler)
Public Sub DisableRealTimeFilter()
    Set inboxItems = Nothing
    Set learnKeepItems = Nothing
    Set learnDeleteItems = Nothing
    Set learnSubjectDeleteItems = Nothing
    Set learnReplyItems = Nothing
    StopCommandPollerStd
    LogMessage "INFO", "Real-time filtering disabled"
    MsgBox "Real-time filtering disabled." & vbCrLf & _
           "New emails will NOT be automatically filtered." & vbCrLf & _
           "All learning folders are also disconnected.", vbInformation
End Sub

' Enable real-time filtering (reconnect event handler)
Public Sub EnableRealTimeFilter()
    Application_Startup
    MsgBox "Real-time filtering enabled." & vbCrLf & _
           "New emails will be automatically filtered.", vbInformation
End Sub

'===============================================================================
' WEB UI COMMAND POLLER
' All poller logic lives in Bridge.bas (standard module) because Win32 timer
' callbacks (AddressOf) cannot target Subs in Document modules like this one.
' See: StartCommandPollerStd, StopCommandPollerStd, PollForCommandsTimer
'===============================================================================
