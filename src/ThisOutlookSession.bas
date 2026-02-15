'===============================================================================
' ThisOutlookSession - Event Handlers v2.0
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
Public WithEvents learnKeepItems As Outlook.Items    ' LearnKeep folder
Public WithEvents learnDeleteItems As Outlook.Items   ' LearnDelete folder
Public WithEvents learnSubjectDeleteItems As Outlook.Items  ' LearnSubjectDelete folder

'-------------------------------------------------------------------------------
' APPLICATION STARTUP
'-------------------------------------------------------------------------------

' Initialize event handlers when Outlook starts
Private Sub Application_Startup()
    On Error GoTo ErrorHandler

    ' CRITICAL: Load all settings from settings.ini FIRST, before anything else.
    ' This populates all Runtime* variables used by every other module.
    LoadAllSettings

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
        On Error GoTo ErrorHandler

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
        On Error GoTo ErrorHandler

        If learnSubjectDeleteFolder Is Nothing Then
            LogMessage "WARN", "Learning folder (" & RuntimeFolderLearnSubject & ") not found under Inbox - subject learning disabled"
        Else
            Set learnSubjectDeleteItems = learnSubjectDeleteFolder.Items
            LoadLearnedSubjects
            DeduplicateLearnedSubjects
            LogMessage "INFO", "Subject learning active (" & GetLearnedSubjectsCount() & " learned subject rules)"
        End If
    End If

    LogMessage "INFO", "Email Filter v" & FILTER_VERSION & " initialized - real-time filtering active"

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "Application_Startup error: " & Err.Description
End Sub

'-------------------------------------------------------------------------------
' NEW MAIL EVENT HANDLER
'-------------------------------------------------------------------------------

' Fires when a new email arrives in the Inbox
' UNCOMMENT THIS SECTION TO ENABLE REAL-TIME FILTERING
'
'Private Sub inboxItems_ItemAdd(ByVal Item As Object)
'    On Error GoTo ErrorHandler
'
'    ' Only process MailItems (not meetings, tasks, etc.)
'    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub
'
'    Dim mail As Outlook.MailItem
'    Set mail = Item
'
'    ' Classify the email
'    Dim decision As String
'    decision = ClassifyEmail(mail)
'
'    ' Execute the action
'    Select Case decision
'        Case "MOVE_II"
'            mail.Move GetOrCreateFolder(RuntimeFolderProtected)
'            LogAction mail, "MOVED to " & RuntimeFolderProtected & " (auto)"
'
'        Case "DELETE"
'            mail.Delete
'            LogAction mail, "DELETED (auto)"
'
'        Case "LLM_REVIEW"
'            ' Move to Review folder (or call LLM if configured)
'            mail.Move GetOrCreateFolder(RuntimeFolderReview)
'            LogAction mail, "MOVED to " & RuntimeFolderReview & " (auto)"
'
'        Case "KEEP"
'            ' Leave in Inbox
'            LogAction mail, "KEPT (auto)"
'    End Select
'
'    Exit Sub
'
'ErrorHandler:
'    LogMessage "ERROR", "inboxItems_ItemAdd error: " & Err.Description
'End Sub

'-------------------------------------------------------------------------------
' SELF-IMPROVING LEARNING EVENT HANDLERS
'-------------------------------------------------------------------------------

' Fires when an email is dragged into the LearnKeep folder (learn to KEEP)
Private Sub learnKeepItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

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

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "learnKeepItems_ItemAdd error: " & Err.Description
End Sub

' Fires when an email is dragged into the LearnDelete folder (learn to DELETE)
Private Sub learnDeleteItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

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

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "learnDeleteItems_ItemAdd error: " & Err.Description
End Sub

' Fires when an email is dragged into the LearnSubjectDelete folder (learn to DELETE by subject)
Private Sub learnSubjectDeleteItems_ItemAdd(ByVal Item As Object)
    On Error GoTo ErrorHandler

    If Not TypeOf Item Is Outlook.MailItem Then Exit Sub

    Dim mail As Outlook.MailItem
    Set mail = Item

    Dim subject As String
    subject = mail.subject

    If Len(Trim(subject)) = 0 Then
        LogMessage "WARN", "learnSubjectDeleteItems_ItemAdd: empty subject, skipping"
        Exit Sub
    End If

    RecordLearnedSubject subject, "DELETE"
    LogMessage "INFO", "LEARNED SUBJECT DELETE from folder " & RuntimeFolderLearnSubject & ": " & Left(subject, 50) & _
               " (" & Left(mail.senderName, 25) & ")"

    Exit Sub

ErrorHandler:
    LogMessage "ERROR", "learnSubjectDeleteItems_ItemAdd error: " & Err.Description
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
    LogMessage "INFO", "Real-time filtering disabled"
    MsgBox "Real-time filtering disabled." & vbCrLf & _
           "New emails will NOT be automatically filtered." & vbCrLf & _
           "Learning folders are also disconnected.", vbInformation
End Sub

' Enable real-time filtering (reconnect event handler)
Public Sub EnableRealTimeFilter()
    Application_Startup
    MsgBox "Real-time filtering enabled." & vbCrLf & _
           "New emails will be automatically filtered.", vbInformation
End Sub
