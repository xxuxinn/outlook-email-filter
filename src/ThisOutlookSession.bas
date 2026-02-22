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

    LogMessage "INFO", "Email Agent v" & FILTER_VERSION & " initialized"

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "Application_Startup", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' NEW MAIL EVENT HANDLER
'-------------------------------------------------------------------------------

' Fires when a new email arrives in the Inbox — real-time filtering.
' Auto-reply drafting is also triggered here when RuntimeAutoReplyOnArrival = True.
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

    ' Execute the action
    Select Case decision
        Case "MOVE_II"
            mail.Move GetOrCreateFolder(RuntimeFolderProtected)
            LogMessage "INFO", "AUTO: MOVED " & senderName & " | " & subject & " to " & RuntimeFolderProtected

        Case "DELETE"
            mail.Delete
            LogMessage "INFO", "AUTO: DELETED " & senderName & " | " & subject

        Case "LLM_REVIEW"
            ' Move to Review folder (or call LLM if configured)
            mail.Move GetOrCreateFolder(RuntimeFolderReview)
            LogMessage "INFO", "AUTO: MOVED " & senderName & " | " & subject & " to " & RuntimeFolderReview

        Case "KEEP"
            ' Leave in Inbox
            LogMessage "INFO", "AUTO: KEPT " & senderName & " | " & subject

            ' Auto-draft reply if enabled (v3.0)
            If RuntimeAutoReplyOnArrival And RuntimeEnableAutoReply Then
                DraftAutoReply mail
            End If
    End Select

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "inboxItems_ItemAdd error: " & Err.Description
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

    ' Extract reply body (text before original message delimiter)
    Dim myReplyText As String
    myReplyText = ExtractMyReplyFromBody(mail.Body)

    ' Extract original body snippet (text after delimiter)
    Dim originalBodySnippet As String
    originalBodySnippet = ExtractOriginalFromBody(mail.Body)

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
    LogMessage "INFO", "LEARNED REPLY from folder " & RuntimeFolderLearnReply & ": " & Left(originalSubject, 50)

PROC_EXIT:
    PopCallStack
    Exit Sub

PROC_ERR:
    LogError "ThisOutlookSession", "learnReplyItems_ItemAdd", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Extract user's reply text from body (before the quoted original message)
Private Function ExtractMyReplyFromBody(ByVal body As String) As String
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
        If pos > 1 Then
            If earliest = 0 Or pos < earliest Then earliest = pos
        End If
    Next d

    If earliest > 1 Then
        ExtractMyReplyFromBody = Trim(Left(body, earliest - 1))
    Else
        ExtractMyReplyFromBody = ""
    End If
End Function

' Extract original message snippet from body (after the quoted original delimiter)
Private Function ExtractOriginalFromBody(ByVal body As String) As String
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
        If pos > 1 Then
            If earliest = 0 Or pos < earliest Then earliest = pos
        End If
    Next d

    If earliest > 0 And earliest + 10 < Len(body) Then
        ExtractOriginalFromBody = Left(Trim(Mid(body, earliest)), 500)
    Else
        ExtractOriginalFromBody = ""
    End If
End Function

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
' All poller logic has been moved to Utilities.bas (standard module) because
' Application.OnTime cannot call Subs in Document modules (ThisOutlookSession).
' See: StartCommandPollerStd, StopCommandPollerStd, PollForCommandsTimer
'===============================================================================
