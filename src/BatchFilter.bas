'===============================================================================
' BatchFilter.bas - Batch Processing Functions v3.0
'===============================================================================
' This module contains functions for filtering existing emails in bulk,
' diagnostics, migration helpers, and server rule import/export.
'
' Structure: each interactive macro (MsgBox confirmations + result dialog) is a
' thin wrapper around a headless Public Function <Name>Core() As String that
' never shows UI, returns a plain-text summary with real counts, and returns
' a string starting with "ERROR: " on failure. Core functions are safe to call
' from the Web UI command bridge.
'
' All pattern/setting references use Runtime* variables from Config.bas,
' loaded from settings.ini by LoadAllSettings at startup.
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' DRY-RUN PREVIEW
'-------------------------------------------------------------------------------

' Preview filtering decisions without making changes
Public Sub FilterExistingDryRun()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim output As String
    Dim decision As String
    Dim icon As String
    Dim processCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterExistingDryRun"

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items
    myItems.Sort "[ReceivedTime]", True  ' Newest first

    output = "=== DRY RUN - Email Filter Preview ===" & vbCrLf
    output = output & "Showing first " & RuntimeDryRunLimit & " emails" & vbCrLf
    output = output & String(50, "=") & vbCrLf & vbCrLf

    processCount = 0

    For i = 1 To myItems.Count
        If processCount >= RuntimeDryRunLimit Then Exit For

        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            processCount = processCount + 1

            decision = ClassifyEmail(mail)

            ' Assign icon based on decision (with learned rule indicators)
            Select Case decision
                Case "DELETE"
                    If lastClassifyWasLearned Then
                        icon = "[xLR]"
                    ElseIf lastClassifyWasLearnedSubject Then
                        icon = "[xLS]"
                    Else
                        icon = "[DEL]"
                    End If
                Case "MOVE_II": icon = "[II] "
                Case "LLM_REVIEW": icon = "[???]"
                Case "KEEP"
                    If lastClassifyWasLearned Then
                        icon = "[+LR]"
                    Else
                        icon = "[OK] "
                    End If
                Case Else: icon = "[???]"
            End Select

            output = output & icon & " " & _
                     Format(mail.ReceivedTime, "mm/dd hh:nn") & " | " & _
                     Truncate(mail.senderName, 20) & " | " & _
                     Truncate(mail.subject, 40) & vbCrLf
        End If
    Next i

    output = output & vbCrLf & String(50, "=") & vbCrLf
    output = output & "Total previewed: " & processCount & " emails" & vbCrLf
    output = output & vbCrLf
    output = output & "Legend:" & vbCrLf
    output = output & "  [DEL] = Will be deleted" & vbCrLf
    output = output & "  [II]  = Will be moved to '" & RuntimeFolderProtected & "' folder" & vbCrLf
    output = output & "  [???] = Will be moved to '" & RuntimeFolderReview & "' folder (or LLM)" & vbCrLf
    output = output & "  [OK]  = Will stay in Inbox" & vbCrLf
    output = output & "  [+LR] = Will stay in Inbox (learned keep rule)" & vbCrLf
    output = output & "  [xLR] = Will be deleted (learned sender rule)" & vbCrLf
    output = output & "  [xLS] = Will be deleted (learned subject rule)" & vbCrLf

    ' Write to Immediate Window
    Debug.Print output

    MsgBox "Dry run complete!" & vbCrLf & vbCrLf & _
           "Check Immediate Window (Ctrl+G in VBA Editor) for results." & vbCrLf & vbCrLf & _
           "Previewed: " & processCount & " emails", vbInformation, "Email Filter Dry Run"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterExistingDryRun", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' FILTER EXISTING EMAILS
'-------------------------------------------------------------------------------

' Filter all existing emails in Inbox (interactive wrapper)
Public Sub FilterExistingEmails()
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterExistingEmails"

    ' Confirm before proceeding
    response = MsgBox("This will filter ALL emails in your Inbox." & vbCrLf & vbCrLf & _
                      "Emails will be:" & vbCrLf & _
                      "  - Deleted (spam/promotional)" & vbCrLf & _
                      "  - Moved to '" & RuntimeFolderProtected & "' (protected sources)" & vbCrLf & _
                      "  - Moved to '" & RuntimeFolderReview & "' (ambiguous)" & vbCrLf & _
                      "  - Kept in Inbox (important)" & vbCrLf & vbCrLf & _
                      "Run dry-run first (FilterExistingDryRun) to preview." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbExclamation, "Email Filter")

    If response <> vbYes Then
        MsgBox "Filtering cancelled.", vbInformation
        GoTo PROC_EXIT
    End If

    resultText = FilterExistingEmailsCore()
    MsgBox resultText, vbInformation, "Email Filter Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterExistingEmails", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: filter all existing emails in Inbox. No UI.
Public Function FilterExistingEmailsCore() As String
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim decision As String
    Dim totalCount As Long
    Dim processedCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterExistingEmailsCore"

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set stats = CreateStatsDict()
    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    ' Sort by received time for predictable processing
    myItems.Sort "[ReceivedTime]", True  ' Newest first

    totalCount = myItems.Count
    processedCount = 0
    LogMessage "INFO", "Starting filter of " & totalCount & " items in Inbox"

    ' Process from end to beginning (CRITICAL for deletions)
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats
            processedCount = processedCount + 1

            ' Progress indicator (log only - no UI in Core)
            If processedCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "Progress: " & processedCount & " / " & totalCount
                DoEvents  ' Allow UI to update
            End If
        End If
    Next i

    LogMessage "INFO", "Filtering complete"
    FilterExistingEmailsCore = SummarizeStats(stats, processedCount)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterExistingEmailsCore", Err.Number, Err.Description
    FilterExistingEmailsCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' FILTER MULTIPLE FOLDERS
'-------------------------------------------------------------------------------

' Filter Inbox, Other folder, and any mounted PST archives (interactive wrapper)
Public Sub FilterAllFolders()
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterAllFolders"

    response = MsgBox("This will filter emails in:" & vbCrLf & _
                      "  - Inbox" & vbCrLf & _
                      "  - 'Other' folder (Focused Inbox)" & vbCrLf & _
                      "  - Any mounted PST archives" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbExclamation, "Email Filter - All Folders")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = FilterAllFoldersCore()
    MsgBox resultText, vbInformation, "Email Filter Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterAllFolders", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: filter Inbox, Other folder, and mounted PST archives. No UI.
Public Function FilterAllFoldersCore() As String
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim otherFolder As Outlook.Folder
    Dim stats As Object
    Dim processedCount As Long
    Dim folderCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterAllFoldersCore"

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set ns = Application.GetNamespace("MAPI")
    Set stats = CreateStatsDict()
    processedCount = 0
    folderCount = 0

    ' 1. Filter default Inbox
    LogMessage "INFO", "Filtering default Inbox..."
    processedCount = processedCount + FilterFolderIntoStats(ns.GetDefaultFolder(olFolderInbox), stats)
    folderCount = folderCount + 1

    ' 2. Filter "Other" folder (Focused Inbox feature)
    Set otherFolder = Nothing
    On Error Resume Next
    Set otherFolder = ns.GetDefaultFolder(olFolderInbox).Folders("Other")
    On Error GoTo PROC_ERR
    If Not otherFolder Is Nothing Then
        LogMessage "INFO", "Filtering 'Other' folder..."
        processedCount = processedCount + FilterFolderIntoStats(otherFolder, stats)
        folderCount = folderCount + 1
    End If

    ' 3. Filter mounted PST archives
    Dim pstInbox As Outlook.Folder
    For Each store In ns.Stores
        If store.ExchangeStoreType = olNotExchange Then
            ' This is a PST file
            LogMessage "INFO", "Filtering PST: " & store.DisplayName
            Set pstInbox = Nothing
            On Error Resume Next
            Set pstInbox = store.GetDefaultFolder(olFolderInbox)
            On Error GoTo PROC_ERR
            If Not pstInbox Is Nothing Then
                processedCount = processedCount + FilterFolderIntoStats(pstInbox, stats)
                folderCount = folderCount + 1
            Else
                LogMessage "WARN", "PST has no Inbox folder: " & store.DisplayName
            End If
        End If
    Next store

    FilterAllFoldersCore = "Processed " & folderCount & " folder(s). " & _
                           SummarizeStats(stats, processedCount)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterAllFoldersCore", Err.Number, Err.Description
    FilterAllFoldersCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

' Filter a specific folder (kept for backwards compatibility - logs its own stats)
Public Sub FilterFolder(ByVal folder As Outlook.Folder)
    Dim stats As Object

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterFolder"

    If folder Is Nothing Then GoTo PROC_EXIT

    Set stats = CreateStatsDict()
    FilterFolderIntoStats folder, stats
    LogMessage "INFO", "Folder complete: " & FormatStats(stats)

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterFolder", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Filter one folder, accumulating into a shared stats dict. Returns emails processed.
Private Function FilterFolderIntoStats(ByVal folder As Outlook.Folder, ByVal stats As Object) As Long
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim decision As String
    Dim processedCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterFolderIntoStats"

    FilterFolderIntoStats = 0
    If folder Is Nothing Then GoTo PROC_EXIT

    Set myItems = folder.Items
    processedCount = 0

    LogMessage "INFO", "Filtering folder: " & folder.Name & " (" & myItems.Count & " items)"

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats
            processedCount = processedCount + 1

            If processedCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "Progress in '" & folder.Name & "': " & processedCount
                DoEvents
            End If
        End If
    Next i

    FilterFolderIntoStats = processedCount

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterFolderIntoStats", Err.Number, Err.Description
    FilterFolderIntoStats = processedCount
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' FILTER BY DATE RANGE
'-------------------------------------------------------------------------------

' Filter emails received within a date range (interactive wrapper)
Public Sub FilterByDateRange(ByVal startDate As Date, ByVal endDate As Date)
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterByDateRange"

    resultText = FilterDateRangeCore(startDate, endDate)
    MsgBox resultText, vbInformation, "Date Range Filter Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterByDateRange", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Filter emails from the last N days (interactive wrapper)
Public Sub FilterLastNDays(ByVal days As Integer)
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterLastNDays"

    resultText = FilterLastNDaysCore(CLng(days))
    MsgBox resultText, vbInformation, "Filter Last " & days & " Days Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterLastNDays", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: filter emails from the last N days. No UI.
Public Function FilterLastNDaysCore(ByVal days As Long) As String
    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterLastNDaysCore"

    If days < 1 Then
        FilterLastNDaysCore = "ERROR: days must be at least 1"
        GoTo PROC_EXIT
    End If

    FilterLastNDaysCore = FilterDateRangeCore(DateAdd("d", -days, Date), Date)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterLastNDaysCore", Err.Number, Err.Description
    FilterLastNDaysCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

' Shared headless date-range filter used by FilterByDateRange and FilterLastNDaysCore.
' Uses locale-safe date literals ("ddddd h:nn AMPM" = VBA locale short date + time,
' the format Microsoft recommends for Jet Restrict) and a strict upper bound at
' midnight of the day AFTER endDate (avoids the old <= endDate + 1 off-by-one).
Private Function FilterDateRangeCore(ByVal startDate As Date, ByVal endDate As Date) As String
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim filter As String
    Dim endDatePlusOne As Date
    Dim processedCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterDateRangeCore"

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set stats = CreateStatsDict()
    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    ' Build locale-safe date filter: [start 00:00, day-after-end 00:00)
    endDatePlusOne = DateAdd("d", 1, DateValue(endDate))
    filter = "[ReceivedTime] >= '" & Format(DateValue(startDate), "ddddd h:nn AMPM") & "' AND " & _
             "[ReceivedTime] < '" & Format(endDatePlusOne, "ddddd h:nn AMPM") & "'"

    Set filteredItems = myItems.Restrict(filter)

    LogMessage "INFO", "Filtering " & filteredItems.Count & " emails from " & _
               Format(startDate, "yyyy-mm-dd") & " to " & Format(endDate, "yyyy-mm-dd")

    processedCount = 0
    For i = filteredItems.Count To 1 Step -1
        If TypeOf filteredItems(i) Is Outlook.MailItem Then
            Set mail = filteredItems(i)
            ExecuteAction mail, ClassifyEmail(mail), stats
            processedCount = processedCount + 1

            If processedCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "Progress: " & processedCount
                DoEvents
            End If
        End If
    Next i

    FilterDateRangeCore = "Date range " & Format(startDate, "yyyy-mm-dd") & " to " & _
                          Format(endDate, "yyyy-mm-dd") & ". " & _
                          SummarizeStats(stats, processedCount)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterDateRangeCore", Err.Number, Err.Description
    FilterDateRangeCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' FILTER BY SENDER PATTERN (BULK DELETE)
'-------------------------------------------------------------------------------

' Delete all emails from senders matching a pattern (interactive wrapper)
Public Sub BulkDeleteBySender(ByVal senderPattern As String)
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.BulkDeleteBySender"

    response = MsgBox("Delete ALL Inbox emails whose sender address or name contains:" & vbCrLf & vbCrLf & _
                      "  " & senderPattern & vbCrLf & vbCrLf & _
                      "Delete all?", vbYesNo + vbExclamation, "Bulk Delete")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = BulkDeleteBySenderCore(senderPattern)
    MsgBox resultText, vbInformation, "Bulk Delete Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "BulkDeleteBySender", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: delete all Inbox emails from senders matching a pattern. No UI.
' Manual scan (no Restrict): Jet "Like" is unsupported by Items.Restrict, and
' [SenderEmailAddress] holds raw /O= Exchange DNs. Matches against the resolved
' SMTP address (GetSenderEmail) and the display name instead.
Public Function BulkDeleteBySenderCore(ByVal senderPattern As String) As String
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim deleteCount As Long
    Dim scannedCount As Long
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.BulkDeleteBySenderCore"

    If Len(Trim(senderPattern)) < 3 Then
        BulkDeleteBySenderCore = "ERROR: pattern too short"
        GoTo PROC_EXIT
    End If
    senderPattern = Trim(senderPattern)

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    deleteCount = 0
    scannedCount = 0

    ' Reverse iteration - deletions invalidate indices
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            scannedCount = scannedCount + 1

            If InStr(1, GetSenderEmail(mail), senderPattern, vbTextCompare) > 0 _
               Or InStr(1, mail.senderName, senderPattern, vbTextCompare) > 0 Then
                ' Pre-capture BEFORE delete (object becomes invalid after)
                senderName = mail.senderName
                subject = mail.subject
                mail.Delete
                deleteCount = deleteCount + 1
                LogActionDirect senderName, subject, "DELETED (bulk: " & senderPattern & ")"
            End If

            If scannedCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "BulkDelete progress: " & scannedCount & " scanned, " & deleteCount & " deleted"
                DoEvents
            End If
        End If
    Next i

    LogMessage "INFO", "BulkDeleteBySender '" & senderPattern & "': scanned " & scannedCount & ", deleted " & deleteCount

    If deleteCount = 0 Then
        BulkDeleteBySenderCore = "No emails found matching pattern: " & senderPattern & _
                                 " (scanned " & scannedCount & " emails)."
    Else
        BulkDeleteBySenderCore = "Deleted " & deleteCount & " of " & scannedCount & _
                                 " emails matching '" & senderPattern & "'."
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "BulkDeleteBySenderCore", Err.Number, Err.Description
    BulkDeleteBySenderCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' MOVE PROTECTED SOURCES (BULK)
'-------------------------------------------------------------------------------

' Move all emails from protected domains to the Protected folder (interactive wrapper)
Public Sub MoveProtectedSources()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.MoveProtectedSources"

    resultText = MoveProtectedSourcesCore()
    MsgBox resultText, vbInformation, "Protected Sources Moved"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "MoveProtectedSources", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: move all Inbox emails from protected domains to Protected folder. No UI.
Public Function MoveProtectedSourcesCore() As String
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim moveCount As Long
    Dim domain As String
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.MoveProtectedSourcesCore"

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    moveCount = 0

    ' Reverse iteration - moves invalidate indices
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            domain = GetDomain(GetSenderEmail(mail))

            If ContainsAny(domain, RuntimeProtectedDomains) Then
                ' Pre-capture BEFORE move (object becomes invalid after)
                senderName = mail.senderName
                subject = mail.subject
                mail.Move GetOrCreateFolder(RuntimeFolderProtected)
                moveCount = moveCount + 1
                LogActionDirect senderName, subject, "MOVED to " & RuntimeFolderProtected & " (bulk protected)"

                If moveCount Mod RuntimeProgressInterval = 0 Then
                    LogMessage "INFO", "MoveProtectedSources progress: " & moveCount & " moved"
                    DoEvents
                End If
            End If
        End If
    Next i

    LogMessage "INFO", "MoveProtectedSources complete: " & moveCount & " moved"
    MoveProtectedSourcesCore = "Moved " & moveCount & " emails to '" & RuntimeFolderProtected & "' folder."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "MoveProtectedSourcesCore", Err.Number, Err.Description
    MoveProtectedSourcesCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' STATISTICS AND REPORTING
'-------------------------------------------------------------------------------

' Generate a classification report without taking action (interactive wrapper)
Public Sub GenerateClassificationReport()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.GenerateClassificationReport"

    resultText = GenerateClassificationReportCore()
    MsgBox resultText, vbInformation, "Classification Report"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "GenerateClassificationReport", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: classification report without taking action. Returns report text.
' NOTE: classify-only - must NOT record decisions (no ExecuteAction, no RecordDecision).
Public Function GenerateClassificationReportCore() As String
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim decision As String
    Dim mailCount As Long
    Dim otherCount As Long
    Dim learnedCount As Long
    Dim learnedSubjectCount As Long
    Dim learnedInfo As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.GenerateClassificationReportCore"

    Set stats = CreateStatsDict()
    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    mailCount = 0
    otherCount = 0
    learnedCount = 0
    learnedSubjectCount = 0

    For i = 1 To myItems.Count
        If TypeOf myItems(i) Is Outlook.MailItem Then
            mailCount = mailCount + 1
            Set mail = myItems(i)
            decision = ClassifyEmail(mail)

            ' Track learned rule matches
            If lastClassifyWasLearned Then learnedCount = learnedCount + 1
            If lastClassifyWasLearnedSubject Then learnedSubjectCount = learnedSubjectCount + 1

            ' Count by decision type
            Select Case decision
                Case "DELETE"
                    IncrementStat stats, "DELETE"
                Case "MOVE_II"
                    IncrementStat stats, "MOVE_II"
                Case "LLM_REVIEW"
                    IncrementStat stats, "REVIEW"
                Case "KEEP"
                    IncrementStat stats, "KEEP"
            End Select

            ' Progress indicator for large mailboxes
            If mailCount Mod 500 = 0 Then
                DoEvents  ' Keep UI responsive
            End If
        Else
            otherCount = otherCount + 1
        End If
    Next i

    If RuntimeEnableSelfImproving Then
        learnedInfo = vbCrLf & "Learned Rules:" & vbCrLf & _
                      "  Matched by learned sender rules: " & learnedCount & vbCrLf & _
                      "  Matched by learned subject rules: " & learnedSubjectCount & vbCrLf & _
                      "  Total learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
                      "  Total learned subject rules: " & GetLearnedSubjectsCount() & vbCrLf
    Else
        learnedInfo = ""
    End If

    GenerateClassificationReportCore = _
           "Classification Report (No Actions Taken)" & vbCrLf & vbCrLf & _
           "Total items in folder: " & myItems.Count & vbCrLf & _
           "  - Emails (MailItems): " & mailCount & vbCrLf & _
           "  - Other (meetings, etc.): " & otherCount & vbCrLf & vbCrLf & _
           "Email Classification:" & vbCrLf & _
           "  Would DELETE: " & stats("DELETE") & vbCrLf & _
           "  Would MOVE to " & RuntimeFolderProtected & ": " & stats("MOVE_II") & vbCrLf & _
           "  Would MOVE to " & RuntimeFolderReview & ": " & stats("REVIEW") & vbCrLf & _
           "  Would KEEP: " & stats("KEEP") & vbCrLf & _
           learnedInfo & vbCrLf & _
           "Verification: " & (stats("DELETE") + stats("MOVE_II") + stats("REVIEW") + stats("KEEP")) & " = " & mailCount

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "GenerateClassificationReportCore", Err.Number, Err.Description
    GenerateClassificationReportCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' UNDO HELPERS
'-------------------------------------------------------------------------------

' Move emails from Review folder back to Inbox (interactive wrapper)
Public Sub RestoreFromReview()
    Dim reviewFolder As Outlook.Folder
    Dim itemCount As Long
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RestoreFromReview"

    ' Peek at the folder (read-only) so the confirmation can show a count
    Set reviewFolder = Nothing
    On Error Resume Next
    Set reviewFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderReview)
    On Error GoTo PROC_ERR

    If reviewFolder Is Nothing Then
        MsgBox "'" & RuntimeFolderReview & "' folder not found.", vbInformation
        GoTo PROC_EXIT
    End If

    itemCount = reviewFolder.Items.Count
    If itemCount = 0 Then
        MsgBox "'" & RuntimeFolderReview & "' folder is empty.", vbInformation
        GoTo PROC_EXIT
    End If

    response = MsgBox("Move " & itemCount & " emails from '" & RuntimeFolderReview & "' back to Inbox?", _
                      vbYesNo + vbQuestion, "Restore from Review")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = RestoreFromReviewCore()
    MsgBox resultText, vbInformation, "Restore Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "RestoreFromReview", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: move all emails from Review folder back to Inbox. No UI.
Public Function RestoreFromReviewCore() As String
    Dim reviewFolder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim i As Long
    Dim restoreCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RestoreFromReviewCore"

    Set reviewFolder = Nothing
    On Error Resume Next
    Set reviewFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderReview)
    On Error GoTo PROC_ERR

    If reviewFolder Is Nothing Then
        RestoreFromReviewCore = "ERROR: '" & RuntimeFolderReview & "' folder not found"
        GoTo PROC_EXIT
    End If

    Set myItems = reviewFolder.Items

    If myItems.Count = 0 Then
        RestoreFromReviewCore = "'" & RuntimeFolderReview & "' folder is empty. Nothing restored."
        GoTo PROC_EXIT
    End If

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    restoreCount = 0

    ' Reverse iteration - moves invalidate indices
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            myItems(i).Move inbox
            restoreCount = restoreCount + 1

            If restoreCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "RestoreFromReview progress: " & restoreCount & " restored"
                DoEvents
            End If
        End If
    Next i

    LogMessage "INFO", "RestoreFromReview complete: " & restoreCount & " restored"
    RestoreFromReviewCore = "Restored " & restoreCount & " emails from '" & RuntimeFolderReview & "' to Inbox."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "RestoreFromReviewCore", Err.Number, Err.Description
    RestoreFromReviewCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' LEARNED SENDERS DIAGNOSTICS
'-------------------------------------------------------------------------------

' Display learned sender rules count and file location
Public Sub ShowLearnedSenders()
    Dim filePath As String
    Dim fso As Object
    Dim fileInfo As String
    Dim fileSize As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ShowLearnedSenders"

    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled." & vbCrLf & _
               "Set EnableSelfImproving=True in settings.ini.", _
               vbInformation, "Learned Senders"
        GoTo PROC_EXIT
    End If

    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        fileSize = fso.GetFile(filePath).Size
        fileInfo = "File size: " & fileSize & " bytes"
    Else
        fileInfo = "File not yet created (no rules learned)"
    End If
    Set fso = Nothing

    MsgBox "Self-Improving Filter Status" & vbCrLf & vbCrLf & _
           "Learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
           "Data file: " & filePath & vbCrLf & _
           fileInfo & vbCrLf & vbCrLf & _
           "How to use:" & vbCrLf & _
           "  Drag email to '" & RuntimeFolderLearnKeep & "' = always KEEP from that sender" & vbCrLf & _
           "  Drag email to '" & RuntimeFolderLearnDelete & "' = always DELETE from that sender" & vbCrLf & vbCrLf & _
           "Macros:" & vbCrLf & _
           "  ReloadLearnedSenders - Force reload from file" & vbCrLf & _
           "  ShowLearnedSenders   - This dialog", _
           vbInformation, "Learned Senders"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ShowLearnedSenders", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Dump all learned sender rules to the Immediate Window for review
Public Sub ShowLearnedSendersList()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim output As String
    Dim parts() As String
    Dim lineNum As Long
    Dim keepCount As Long
    Dim deleteCount As Long
    Dim count As Long
    Dim action As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ShowLearnedSendersList"

    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Senders"
        GoTo PROC_EXIT
    End If

    ' Force reload to get freshest data
    LoadLearnedSenders True

    count = GetLearnedSendersCount()

    If count = 0 Then
        MsgBox "No learned sender rules found.", vbInformation, "Learned Senders"
        GoTo PROC_EXIT
    End If

    ' Read the file directly to show all entries with timestamps
    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "Learned senders file not found at:" & vbCrLf & filePath, vbExclamation
        Set fso = Nothing
        GoTo PROC_EXIT
    End If

    output = "=== LEARNED SENDER RULES ===" & vbCrLf
    output = output & "File: " & filePath & vbCrLf
    output = output & String(60, "=") & vbCrLf & vbCrLf

    keepCount = 0
    deleteCount = 0
    lineNum = 0

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            lineNum = lineNum + 1
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                action = UCase(Trim(parts(1)))
                If action = "KEEP" Then
                    keepCount = keepCount + 1
                    output = output & "[+KEEP]   " & Trim(parts(0))
                ElseIf action = "DELETE" Then
                    deleteCount = deleteCount + 1
                    output = output & "[xDELETE] " & Trim(parts(0))
                Else
                    output = output & "[?????]   " & Trim(parts(0))
                End If
                ' Append timestamp if present
                If UBound(parts) >= 2 Then
                    output = output & "  (" & Trim(parts(2)) & ")"
                End If
                output = output & vbCrLf
            End If
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    output = output & vbCrLf & String(60, "=") & vbCrLf
    output = output & "Total lines in file: " & lineNum & vbCrLf
    output = output & "Unique rules in cache: " & count & vbCrLf
    output = output & "KEEP rules: " & keepCount & "  |  DELETE rules: " & deleteCount & vbCrLf
    output = output & vbCrLf
    output = output & "NOTE: If a sender appears multiple times, the LAST entry wins." & vbCrLf

    Debug.Print output

    MsgBox "Learned senders list printed to Immediate Window (Ctrl+G)." & vbCrLf & vbCrLf & _
           "Unique rules in cache: " & count & vbCrLf & _
           "KEEP: " & keepCount & "  |  DELETE: " & deleteCount, _
           vbInformation, "Learned Senders List"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ShowLearnedSendersList", Err.Number, Err.Description
    ' Safe cleanup: after On Error Resume Next the error state is cleared,
    ' so jump with GoTo (Resume would raise "Resume without error")
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    GoTo PROC_EXIT
End Sub

' Remove duplicate entries from the learned senders file (interactive wrapper)
Public Sub CleanLearnedSendersFile()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.CleanLearnedSendersFile"

    resultText = CleanLearnedSendersFileCore()
    MsgBox resultText, vbInformation, "Clean Learned Senders"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "CleanLearnedSendersFile", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: deduplicate the learned senders file. No UI.
' Counts only rule lines (non-comment, non-blank) so "removed" is accurate.
Public Function CleanLearnedSendersFileCore() As String
    Dim beforeRules As Long
    Dim afterRules As Long
    Dim removed As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.CleanLearnedSendersFileCore"

    If Not RuntimeEnableSelfImproving Then
        CleanLearnedSendersFileCore = "ERROR: self-improving filter is disabled (set EnableSelfImproving=True in settings.ini)"
        GoTo PROC_EXIT
    End If

    beforeRules = CountRuleLines(GetLearnedSendersFilePath())

    ' Run deduplication
    DeduplicateLearnedSenders

    afterRules = CountRuleLines(GetLearnedSendersFilePath())
    removed = beforeRules - afterRules
    If removed < 0 Then removed = 0

    If removed = 0 Then
        CleanLearnedSendersFileCore = "No duplicates found. File has " & afterRules & " rule entries, all unique."
    Else
        CleanLearnedSendersFileCore = "Deduplication complete: " & beforeRules & " -> " & afterRules & _
                                      " rule entries (" & removed & " duplicate(s) removed). " & _
                                      "Unique rules in cache: " & GetLearnedSendersCount() & "."
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "CleanLearnedSendersFileCore", Err.Number, Err.Description
    CleanLearnedSendersFileCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' RESTORE WRONGLY DELETED EMAILS
'-------------------------------------------------------------------------------

' Scan Deleted Items, re-classify, and move KEEP/MOVE_II back (interactive wrapper)
Public Sub RestoreDeletedKeepEmails()
    Dim deletedFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim i As Long
    Dim totalMail As Long
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RestoreDeletedKeepEmails"

    ' Peek (read-only) so the confirmation can show a count
    Set deletedFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems)
    Set myItems = deletedFolder.Items

    totalMail = 0
    For i = 1 To myItems.Count
        If TypeOf myItems(i) Is Outlook.MailItem Then totalMail = totalMail + 1
    Next i

    If totalMail = 0 Then
        MsgBox "No emails in Deleted Items.", vbInformation, "Restore"
        GoTo PROC_EXIT
    End If

    response = MsgBox("Scan " & totalMail & " emails in Deleted Items?" & vbCrLf & vbCrLf & _
                      "Emails classified as KEEP will be moved back to Inbox." & vbCrLf & _
                      "Emails classified as MOVE_II will be moved to '" & RuntimeFolderProtected & "'." & vbCrLf & _
                      "All other emails stay in Deleted Items.", _
                      vbYesNo + vbQuestion, "Restore Deleted Keep Emails")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = RestoreDeletedKeepEmailsCore()
    MsgBox resultText, vbInformation, "Restore Deleted Keep Emails"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "RestoreDeletedKeepEmails", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: re-classify Deleted Items and restore KEEP/MOVE_II emails. No UI.
Public Function RestoreDeletedKeepEmailsCore() As String
    Dim deletedFolder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim decision As String
    Dim i As Long
    Dim restoreCount As Long
    Dim moveIICount As Long
    Dim totalMail As Long
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RestoreDeletedKeepEmailsCore"

    ' Force reload learned caches to use the most up-to-date rules
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders True
        LoadLearnedSubjects True
    End If

    Set deletedFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderDeletedItems)
    Set myItems = deletedFolder.Items

    ' Count MailItems
    totalMail = 0
    For i = 1 To myItems.Count
        If TypeOf myItems(i) Is Outlook.MailItem Then totalMail = totalMail + 1
    Next i

    If totalMail = 0 Then
        RestoreDeletedKeepEmailsCore = "No emails in Deleted Items. Nothing restored."
        GoTo PROC_EXIT
    End If

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    restoreCount = 0
    moveIICount = 0

    ' Reverse iteration - moves invalidate indices
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            ' Pre-capture BEFORE move (object becomes invalid after)
            senderName = mail.senderName
            subject = mail.subject

            decision = ClassifyEmail(mail)

            If decision = "KEEP" Then
                mail.Move inbox
                restoreCount = restoreCount + 1
                LogMessage "INFO", "RESTORED to Inbox: " & Left(senderName, 25) & " | " & Left(subject, 40)
            ElseIf decision = "MOVE_II" Then
                mail.Move GetOrCreateFolder(RuntimeFolderProtected)
                moveIICount = moveIICount + 1
                LogMessage "INFO", "RESTORED to " & RuntimeFolderProtected & ": " & Left(senderName, 25) & " | " & Left(subject, 40)
            End If
            ' DELETE / LLM_REVIEW emails stay in Deleted Items

            If (restoreCount + moveIICount) Mod 50 = 0 Then DoEvents
        End If
    Next i

    LogMessage "INFO", "RestoreDeletedKeepEmails complete: " & restoreCount & " to Inbox, " & moveIICount & " to " & RuntimeFolderProtected

    RestoreDeletedKeepEmailsCore = "Scanned " & totalMail & " deleted emails: " & _
                                   restoreCount & " restored to Inbox, " & _
                                   moveIICount & " moved to " & RuntimeFolderProtected & ", " & _
                                   (totalMail - restoreCount - moveIICount) & " left in Deleted Items."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "RestoreDeletedKeepEmailsCore", Err.Number, Err.Description
    RestoreDeletedKeepEmailsCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' QUICK FILTER ACTIONS (assign to QAT or ribbon button)
'-------------------------------------------------------------------------------

' Filter selected email(s) in the active explorer window (interactive wrapper)
Public Sub FilterSelectedEmails()
    Dim explorer As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim i As Long
    Dim mailCount As Long
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterSelectedEmails"

    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then
        MsgBox "No active Outlook window.", vbExclamation, "Filter Selected"
        GoTo PROC_EXIT
    End If

    Set sel = explorer.Selection

    If sel.Count = 0 Then
        MsgBox "No emails selected.", vbExclamation, "Filter Selected"
        GoTo PROC_EXIT
    End If

    ' Count actual MailItems in selection
    mailCount = 0
    For i = 1 To sel.Count
        If TypeOf sel(i) Is Outlook.MailItem Then mailCount = mailCount + 1
    Next i

    If mailCount = 0 Then
        MsgBox "No email items in selection (meetings/tasks are skipped).", vbExclamation, "Filter Selected"
        GoTo PROC_EXIT
    End If

    ' Confirmation dialog
    response = MsgBox("Filter " & mailCount & " selected email(s)?" & vbCrLf & vbCrLf & _
                      "Actions: Delete / Move to " & RuntimeFolderProtected & " / Move to " & RuntimeFolderReview & " / Keep", _
                      vbYesNo + vbQuestion, "Filter Selected Emails")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = FilterSelectedEmailsCore()
    MsgBox resultText, vbInformation, "Filter Selected Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterSelectedEmails", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: filter the current Outlook selection without confirmation.
' Returns a per-email decision summary plus totals.
Public Function FilterSelectedEmailsCore() As String
    Dim explorer As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim mail As Outlook.MailItem
    Dim stats As Object
    Dim decision As String
    Dim i As Long
    Dim mailCount As Long
    Dim detailText As String
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterSelectedEmailsCore"

    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then
        FilterSelectedEmailsCore = "ERROR: no active Outlook window"
        GoTo PROC_EXIT
    End If

    Set sel = explorer.Selection

    If sel.Count = 0 Then
        FilterSelectedEmailsCore = "ERROR: no emails selected"
        GoTo PROC_EXIT
    End If

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set stats = CreateStatsDict()
    mailCount = 0
    detailText = ""

    ' Reverse iteration - deletions/moves invalidate indices
    For i = sel.Count To 1 Step -1
        If TypeOf sel(i) Is Outlook.MailItem Then
            Set mail = sel(i)

            ' Pre-capture BEFORE action (object becomes invalid after delete/move)
            senderName = mail.senderName
            subject = mail.subject

            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats
            mailCount = mailCount + 1

            detailText = detailText & decision & " | " & _
                         Truncate(senderName, 25) & " | " & _
                         Truncate(subject, 40) & vbCrLf
        End If
    Next i

    If mailCount = 0 Then
        FilterSelectedEmailsCore = "ERROR: no email items in selection (meetings/tasks are skipped)"
        GoTo PROC_EXIT
    End If

    FilterSelectedEmailsCore = detailText & _
                               SummarizeStats(stats, mailCount)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterSelectedEmailsCore", Err.Number, Err.Description
    FilterSelectedEmailsCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

' Filter all emails in the currently displayed folder (interactive wrapper)
Public Sub FilterCurrentFolder()
    Dim explorer As Outlook.Explorer
    Dim folder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim response As VbMsgBoxResult
    Dim isNonInbox As Boolean
    Dim isReviewFolder As Boolean
    Dim itemCount As Long
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterCurrentFolder"

    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then
        MsgBox "No active Outlook window.", vbExclamation, "Filter Folder"
        GoTo PROC_EXIT
    End If

    Set folder = explorer.CurrentFolder

    If folder Is Nothing Then
        MsgBox "No folder is currently active.", vbExclamation, "Filter Folder"
        GoTo PROC_EXIT
    End If

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    isNonInbox = (folder.EntryID <> inbox.EntryID)
    isReviewFolder = (StrComp(folder.Name, RuntimeFolderReview, vbTextCompare) = 0)
    itemCount = folder.Items.Count

    ' Confirmation dialog (different message for non-Inbox folders)
    If isReviewFolder Then
        response = MsgBox("Filter Review folder?" & vbCrLf & _
                          "(" & itemCount & " items)" & vbCrLf & vbCrLf & _
                          "Review folder detected." & vbCrLf & _
                          "Only DELETE rules will be applied." & vbCrLf & _
                          "Non-DELETE emails will stay in Review.", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    ElseIf isNonInbox Then
        response = MsgBox("Filter folder '" & folder.Name & "'?" & vbCrLf & _
                          "(" & itemCount & " items)" & vbCrLf & vbCrLf & _
                          "Non-Inbox folder detected." & vbCrLf & _
                          "KEEP/MOVE_II/REVIEW emails will be moved to Inbox." & vbCrLf & _
                          "DELETE emails will be deleted.", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    Else
        response = MsgBox("Filter folder '" & folder.Name & "'?" & vbCrLf & _
                          "(" & itemCount & " items)" & vbCrLf & vbCrLf & _
                          "Actions: Delete / Move to " & RuntimeFolderProtected & " / Move to " & RuntimeFolderReview & " / Keep", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    End If

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = FilterCurrentFolderCore()
    MsgBox resultText, vbInformation, "Filter '" & folder.Name & "' Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterCurrentFolder", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: filter the currently displayed folder. No UI.
' Keeps Review DELETE-only mode: in the Review folder, non-DELETE emails stay put.
Public Function FilterCurrentFolderCore() As String
    Dim explorer As Outlook.Explorer
    Dim folder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim stats As Object
    Dim mail As Outlook.MailItem
    Dim decision As String
    Dim i As Long
    Dim isNonInbox As Boolean
    Dim isReviewFolder As Boolean
    Dim senderName As String
    Dim subject As String
    Dim processedCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterCurrentFolderCore"

    Set explorer = Application.ActiveExplorer
    If explorer Is Nothing Then
        FilterCurrentFolderCore = "ERROR: no active Outlook window"
        GoTo PROC_EXIT
    End If

    Set folder = explorer.CurrentFolder

    If folder Is Nothing Then
        FilterCurrentFolderCore = "ERROR: no folder is currently active"
        GoTo PROC_EXIT
    End If

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    isNonInbox = (folder.EntryID <> inbox.EntryID)
    isReviewFolder = (StrComp(folder.Name, RuntimeFolderReview, vbTextCompare) = 0)

    Set myItems = folder.Items
    Set stats = CreateStatsDict()
    processedCount = 0

    LogMessage "INFO", "Filtering folder '" & folder.Name & "' (" & myItems.Count & " items)" & _
               IIf(isReviewFolder, " [Review: DELETE-only mode]", IIf(isNonInbox, " [non-Inbox: KEEP->Inbox]", ""))

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            decision = ClassifyEmail(mail)

            If isNonInbox And decision <> "DELETE" Then
                If isReviewFolder Then
                    ' Review folder: leave non-DELETE emails in place
                    senderName = mail.senderName
                    subject = mail.subject
                    LogActionDirect senderName, subject, "KEPT in Review (was " & decision & ")"
                    IncrementStat stats, "KEEP"
                Else
                    ' Other non-Inbox folders: move to Inbox as before
                    senderName = mail.senderName
                    subject = mail.subject
                    mail.Move inbox
                    LogActionDirect senderName, subject, "MOVED to Inbox (was " & decision & ")"
                    IncrementStat stats, "KEEP"
                End If
            Else
                ExecuteAction mail, decision, stats
            End If

            processedCount = processedCount + 1
            If processedCount Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "Progress in '" & folder.Name & "': " & processedCount
                DoEvents
            End If
        End If
    Next i

    LogMessage "INFO", "Folder '" & folder.Name & "' complete: " & FormatStats(stats)

    FilterCurrentFolderCore = "Folder '" & folder.Name & "'" & _
                              IIf(isReviewFolder, " (Review DELETE-only mode)", "") & ". " & _
                              SummarizeStats(stats, processedCount)

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "FilterCurrentFolderCore", Err.Number, Err.Description
    FilterCurrentFolderCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' LEARNED SENDERS: BULK IMPORT
'-------------------------------------------------------------------------------

' One-time import: scan existing emails in LearnKeep and LearnDelete folders
' and record all senders as learned rules. Run once after upgrading. (interactive wrapper)
Public Sub ImportExistingLearnedFolders()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportExistingLearnedFolders"

    resultText = ImportExistingLearnedFoldersCore()
    MsgBox resultText, vbInformation, "Import Learned Folders"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ImportExistingLearnedFolders", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: bulk-import senders from LearnKeep/LearnDelete folders. No UI.
Public Function ImportExistingLearnedFoldersCore() As String
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim keepFolder As Outlook.Folder
    Dim deleteFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim keepCount As Long
    Dim deleteCount As Long
    Dim skipCount As Long
    Dim senderEmail As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportExistingLearnedFoldersCore"

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Find the learning folders (each lookup can legitimately fail)
    Set keepFolder = Nothing
    Set deleteFolder = Nothing
    On Error Resume Next
    Set keepFolder = inbox.Folders(RuntimeFolderLearnKeep)
    Set deleteFolder = inbox.Folders(RuntimeFolderLearnDelete)
    On Error GoTo PROC_ERR

    If keepFolder Is Nothing And deleteFolder Is Nothing Then
        ImportExistingLearnedFoldersCore = "ERROR: neither '" & RuntimeFolderLearnKeep & _
                                           "' nor '" & RuntimeFolderLearnDelete & "' folder found under Inbox"
        GoTo PROC_EXIT
    End If

    keepCount = 0
    deleteCount = 0
    skipCount = 0

    ' Import from LearnKeep
    If Not keepFolder Is Nothing Then
        Set myItems = keepFolder.Items
        For i = 1 To myItems.Count
            If TypeOf myItems(i) Is Outlook.MailItem Then
                Set mail = myItems(i)
                senderEmail = GetSenderEmail(mail)
                If InStr(1, senderEmail, "@") > 0 Then
                    RecordLearnedSender senderEmail, "KEEP"
                    keepCount = keepCount + 1
                Else
                    skipCount = skipCount + 1
                    LogMessage "WARN", "Import skipped (no @): " & senderEmail & " | " & mail.senderName
                End If
            End If
        Next i
    End If

    ' Import from LearnDelete
    If Not deleteFolder Is Nothing Then
        Set myItems = deleteFolder.Items
        For i = 1 To myItems.Count
            If TypeOf myItems(i) Is Outlook.MailItem Then
                Set mail = myItems(i)
                senderEmail = GetSenderEmail(mail)
                If InStr(1, senderEmail, "@") > 0 Then
                    RecordLearnedSender senderEmail, "DELETE"
                    deleteCount = deleteCount + 1
                Else
                    skipCount = skipCount + 1
                    LogMessage "WARN", "Import skipped (no @): " & senderEmail & " | " & mail.senderName
                End If
            End If
        Next i
    End If

    ImportExistingLearnedFoldersCore = "Imported " & keepCount & " KEEP senders from '" & RuntimeFolderLearnKeep & _
                                       "' and " & deleteCount & " DELETE senders from '" & RuntimeFolderLearnDelete & "'." & _
                                       IIf(skipCount > 0, " Skipped (no @ in address): " & skipCount & ".", "") & _
                                       " Total unique rules now: " & GetLearnedSendersCount() & "."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ImportExistingLearnedFoldersCore", Err.Number, Err.Description
    ImportExistingLearnedFoldersCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' LEARNED SUBJECTS DIAGNOSTICS
'-------------------------------------------------------------------------------

' Dump all learned subject rules to the Immediate Window for review
Public Sub ShowLearnedSubjectsList()
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim output As String
    Dim parts() As String
    Dim lineNum As Long
    Dim deleteCount As Long
    Dim count As Long
    Dim action As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ShowLearnedSubjectsList"

    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Subjects"
        GoTo PROC_EXIT
    End If

    ' Force reload to get freshest data
    LoadLearnedSubjects True

    count = GetLearnedSubjectsCount()

    If count = 0 Then
        MsgBox "No learned subject rules found." & vbCrLf & vbCrLf & _
               "Drag emails into the '" & RuntimeFolderLearnSubject & "' folder to learn subject-based DELETE rules.", _
               vbInformation, "Learned Subjects"
        GoTo PROC_EXIT
    End If

    ' Read the file directly to show all entries with timestamps
    filePath = GetLearnedSubjectsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "Learned subjects file not found at:" & vbCrLf & filePath, vbExclamation
        Set fso = Nothing
        GoTo PROC_EXIT
    End If

    output = "=== LEARNED SUBJECT RULES ===" & vbCrLf
    output = output & "File: " & filePath & vbCrLf
    output = output & String(60, "=") & vbCrLf & vbCrLf

    deleteCount = 0
    lineNum = 0

    Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If Len(line) > 0 And Left(line, 1) <> "#" Then
            lineNum = lineNum + 1
            parts = Split(line, "|")
            If UBound(parts) >= 1 Then
                action = UCase(Trim(parts(1)))
                If action = "DELETE" Then
                    deleteCount = deleteCount + 1
                    output = output & "[xDELETE] " & Trim(parts(0))
                Else
                    output = output & "[?????]   " & Trim(parts(0))
                End If
                ' Append timestamp if present
                If UBound(parts) >= 2 Then
                    output = output & "  (" & Trim(parts(2)) & ")"
                End If
                output = output & vbCrLf
            End If
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    output = output & vbCrLf & String(60, "=") & vbCrLf
    output = output & "Total lines in file: " & lineNum & vbCrLf
    output = output & "Unique rules in cache: " & count & vbCrLf
    output = output & "DELETE rules: " & deleteCount & vbCrLf
    output = output & vbCrLf
    output = output & "NOTE: If a subject appears multiple times, the LAST entry wins." & vbCrLf
    output = output & "Matching: case-insensitive substring match against incoming subject." & vbCrLf

    Debug.Print output

    MsgBox "Learned subjects list printed to Immediate Window (Ctrl+G)." & vbCrLf & vbCrLf & _
           "Unique rules in cache: " & count & vbCrLf & _
           "DELETE: " & deleteCount, _
           vbInformation, "Learned Subjects List"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ShowLearnedSubjectsList", Err.Number, Err.Description
    ' Safe cleanup: after On Error Resume Next the error state is cleared,
    ' so jump with GoTo (Resume would raise "Resume without error")
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    GoTo PROC_EXIT
End Sub

' Remove duplicate entries from the learned subjects file (interactive wrapper)
Public Sub CleanLearnedSubjectsFile()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.CleanLearnedSubjectsFile"

    resultText = CleanLearnedSubjectsFileCore()
    MsgBox resultText, vbInformation, "Clean Learned Subjects"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "CleanLearnedSubjectsFile", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: deduplicate the learned subjects file. No UI.
' Counts only rule lines (non-comment, non-blank) so "removed" is accurate.
Public Function CleanLearnedSubjectsFileCore() As String
    Dim beforeRules As Long
    Dim afterRules As Long
    Dim removed As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.CleanLearnedSubjectsFileCore"

    If Not RuntimeEnableSelfImproving Then
        CleanLearnedSubjectsFileCore = "ERROR: self-improving filter is disabled (set EnableSelfImproving=True in settings.ini)"
        GoTo PROC_EXIT
    End If

    beforeRules = CountRuleLines(GetLearnedSubjectsFilePath())

    ' Run deduplication
    DeduplicateLearnedSubjects

    afterRules = CountRuleLines(GetLearnedSubjectsFilePath())
    removed = beforeRules - afterRules
    If removed < 0 Then removed = 0

    If removed = 0 Then
        CleanLearnedSubjectsFileCore = "No duplicates found. File has " & afterRules & " rule entries, all unique."
    Else
        CleanLearnedSubjectsFileCore = "Deduplication complete: " & beforeRules & " -> " & afterRules & _
                                       " rule entries (" & removed & " duplicate(s) removed). " & _
                                       "Unique rules in cache: " & GetLearnedSubjectsCount() & "."
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "CleanLearnedSubjectsFileCore", Err.Number, Err.Description
    CleanLearnedSubjectsFileCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

' One-time import: scan existing emails in LearnSubjectDelete folder
' and record all subjects as learned DELETE rules. (interactive wrapper)
Public Sub ImportExistingLearnedSubjectFolder()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportExistingLearnedSubjectFolder"

    resultText = ImportExistingLearnedSubjectFolderCore()
    MsgBox resultText, vbInformation, "Import Learned Subject Folder"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ImportExistingLearnedSubjectFolder", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: bulk-import subjects from LearnSubjectDelete folder. No UI.
Public Function ImportExistingLearnedSubjectFolderCore() As String
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim subjectFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim importCount As Long
    Dim skipCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportExistingLearnedSubjectFolderCore"

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Find the folder (lookup can legitimately fail)
    Set subjectFolder = Nothing
    On Error Resume Next
    Set subjectFolder = inbox.Folders(RuntimeFolderLearnSubject)
    On Error GoTo PROC_ERR

    If subjectFolder Is Nothing Then
        ImportExistingLearnedSubjectFolderCore = "ERROR: '" & RuntimeFolderLearnSubject & _
                                                 "' folder not found under Inbox (create it first)"
        GoTo PROC_EXIT
    End If

    Set myItems = subjectFolder.Items
    importCount = 0
    skipCount = 0

    For i = 1 To myItems.Count
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            If Len(Trim(mail.subject)) > 0 Then
                RecordLearnedSubject mail.subject, "DELETE"
                importCount = importCount + 1
            Else
                skipCount = skipCount + 1
                LogMessage "WARN", "Import skipped (empty subject): " & mail.senderName
            End If
        End If
    Next i

    ImportExistingLearnedSubjectFolderCore = "Imported " & importCount & " DELETE subjects from '" & _
                                             RuntimeFolderLearnSubject & "'." & _
                                             IIf(skipCount > 0, " Skipped (empty subject): " & skipCount & ".", "") & _
                                             " Total unique subject rules now: " & GetLearnedSubjectsCount() & "."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ImportExistingLearnedSubjectFolderCore", Err.Number, Err.Description
    ImportExistingLearnedSubjectFolderCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' IMPORT SERVER-SIDE RULES AS LEARNED RULES
'-------------------------------------------------------------------------------

' Import server-side Outlook Rules as learned rules (interactive wrapper)
Public Sub ImportServerRules()
    Dim rules As Object  ' Outlook.Rules
    Dim i As Long
    Dim enabledCount As Long
    Dim ruleEnabled As Boolean
    Dim response As VbMsgBoxResult
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportServerRules"

    ' Get server rules (read-only peek for the confirmation dialog)
    Set rules = Application.Session.DefaultStore.GetRules()

    If rules.Count = 0 Then
        MsgBox "No server-side rules found.", vbInformation, "Import Server Rules"
        GoTo PROC_EXIT
    End If

    ' Count enabled rules (narrow OERN: .Enabled can fail on corrupt rules)
    enabledCount = 0
    For i = 1 To rules.Count
        ruleEnabled = False
        On Error Resume Next
        ruleEnabled = rules.Item(i).Enabled
        On Error GoTo PROC_ERR
        If ruleEnabled Then enabledCount = enabledCount + 1
    Next i

    response = MsgBox("Found " & rules.Count & " server rules (" & enabledCount & " enabled)." & vbCrLf & vbCrLf & _
                      "This will import sender addresses and subject keywords from" & vbCrLf & _
                      "ENABLED rules as learned DELETE rules." & vbCrLf & vbCrLf & _
                      "After import, you can manually delete the server rules" & vbCrLf & _
                      "via Home -> Rules -> Manage Rules & Alerts." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Import Server Rules")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = ImportServerRulesCore()
    MsgBox resultText, vbInformation, "Import Server Rules"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ImportServerRules", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: import server rules as learned DELETE rules. No UI.
' Each On Error Resume Next covers exactly one COM property access; failures
' are counted (condErrors) and logged instead of silently swallowed.
Public Function ImportServerRulesCore() As String
    Dim rules As Object  ' Outlook.Rules
    Dim rl As Object     ' Outlook.Rule
    Dim i As Long
    Dim j As Long
    Dim senderCount As Long
    Dim subjectCount As Long
    Dim ruleCount As Long
    Dim skippedCount As Long
    Dim condErrors As Long
    Dim ruleEnabled As Boolean
    Dim ruleName As String
    Dim condObj As Object
    Dim condEnabled As Boolean
    Dim recipCount As Long
    Dim recip As Object
    Dim recipEmail As String
    Dim addrArray As Variant
    Dim addr As Variant
    Dim subjArray As Variant
    Dim subj As Variant
    Dim sobArray As Variant
    Dim sob As Variant
    Dim errText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportServerRulesCore"

    ' Get server rules
    Set rules = Application.Session.DefaultStore.GetRules()

    If rules.Count = 0 Then
        ImportServerRulesCore = "No server-side rules found. Nothing imported."
        GoTo PROC_EXIT
    End If

    senderCount = 0
    subjectCount = 0
    ruleCount = 0
    skippedCount = 0
    condErrors = 0

    For i = 1 To rules.Count
        Set rl = rules.Item(i)

        ' Only process enabled rules (.Enabled can fail on corrupt rules)
        ruleEnabled = False
        On Error Resume Next
        ruleEnabled = rl.Enabled
        If Err.Number <> 0 Then
            errText = Err.Description
            condErrors = condErrors + 1
            LogMessage "WARN", "ImportServerRules: cannot read Enabled on rule #" & i & ": " & errText
        End If
        On Error GoTo PROC_ERR

        If Not ruleEnabled Then
            skippedCount = skippedCount + 1
            GoTo NextRule
        End If

        ruleCount = ruleCount + 1

        ruleName = "(unnamed rule #" & i & ")"
        On Error Resume Next
        ruleName = rl.Name
        On Error GoTo PROC_ERR
        LogMessage "INFO", "Importing rule: " & ruleName

        ' --- Extract SENDER conditions ---

        ' Conditions.From.Recipients (Exchange recipients)
        Set condObj = Nothing
        On Error Resume Next
        Set condObj = rl.Conditions.From
        If Err.Number <> 0 Then
            errText = Err.Description
            condErrors = condErrors + 1
            LogMessage "WARN", "ImportServerRules: Conditions.From unreadable on '" & ruleName & "': " & errText
        End If
        On Error GoTo PROC_ERR

        If Not condObj Is Nothing Then
            condEnabled = False
            On Error Resume Next
            condEnabled = condObj.Enabled
            If Err.Number <> 0 Then
                errText = Err.Description
                condErrors = condErrors + 1
                LogMessage "WARN", "ImportServerRules: From.Enabled unreadable on '" & ruleName & "': " & errText
            End If
            On Error GoTo PROC_ERR

            If condEnabled Then
                recipCount = 0
                On Error Resume Next
                recipCount = condObj.Recipients.Count
                If Err.Number <> 0 Then
                    errText = Err.Description
                    condErrors = condErrors + 1
                    LogMessage "WARN", "ImportServerRules: From.Recipients unreadable on '" & ruleName & "': " & errText
                End If
                On Error GoTo PROC_ERR

                For j = 1 To recipCount
                    Set recip = Nothing
                    On Error Resume Next
                    Set recip = condObj.Recipients.Item(j)
                    If Err.Number <> 0 Then
                        errText = Err.Description
                        condErrors = condErrors + 1
                        LogMessage "WARN", "ImportServerRules: recipient " & j & " unreadable on '" & ruleName & "': " & errText
                    End If
                    On Error GoTo PROC_ERR

                    If Not recip Is Nothing Then
                        recipEmail = ResolveRecipientSmtp(recip)
                        If Len(recipEmail) > 0 And InStr(1, recipEmail, "@") > 0 Then
                            RecordLearnedSender recipEmail, "DELETE"
                            senderCount = senderCount + 1
                            LogMessage "INFO", "  Sender from rule '" & ruleName & "': " & recipEmail
                        ElseIf Len(recipEmail) = 0 Then
                            condErrors = condErrors + 1
                            LogMessage "WARN", "ImportServerRules: could not resolve recipient " & j & " on '" & ruleName & "'"
                        End If
                    End If
                Next j
            End If
        End If

        ' Conditions.SenderAddress.Address (string array of email addresses)
        Set condObj = Nothing
        On Error Resume Next
        Set condObj = rl.Conditions.SenderAddress
        If Err.Number <> 0 Then
            errText = Err.Description
            condErrors = condErrors + 1
            LogMessage "WARN", "ImportServerRules: Conditions.SenderAddress unreadable on '" & ruleName & "': " & errText
        End If
        On Error GoTo PROC_ERR

        If Not condObj Is Nothing Then
            condEnabled = False
            On Error Resume Next
            condEnabled = condObj.Enabled
            If Err.Number <> 0 Then
                errText = Err.Description
                condErrors = condErrors + 1
                LogMessage "WARN", "ImportServerRules: SenderAddress.Enabled unreadable on '" & ruleName & "': " & errText
            End If
            On Error GoTo PROC_ERR

            If condEnabled Then
                addrArray = Empty
                On Error Resume Next
                addrArray = condObj.Address
                If Err.Number <> 0 Then
                    errText = Err.Description
                    condErrors = condErrors + 1
                    LogMessage "WARN", "ImportServerRules: SenderAddress.Address unreadable on '" & ruleName & "': " & errText
                End If
                On Error GoTo PROC_ERR

                If IsArray(addrArray) Then
                    For Each addr In addrArray
                        If Len(addr) > 0 And InStr(1, CStr(addr), "@") > 0 Then
                            RecordLearnedSender LCase(CStr(addr)), "DELETE"
                            senderCount = senderCount + 1
                            LogMessage "INFO", "  SenderAddress from rule '" & ruleName & "': " & CStr(addr)
                        End If
                    Next addr
                End If
            End If
        End If

        ' --- Extract SUBJECT conditions ---

        ' Conditions.Subject.Text (array of subject keywords)
        Set condObj = Nothing
        On Error Resume Next
        Set condObj = rl.Conditions.subject
        If Err.Number <> 0 Then
            errText = Err.Description
            condErrors = condErrors + 1
            LogMessage "WARN", "ImportServerRules: Conditions.Subject unreadable on '" & ruleName & "': " & errText
        End If
        On Error GoTo PROC_ERR

        If Not condObj Is Nothing Then
            condEnabled = False
            On Error Resume Next
            condEnabled = condObj.Enabled
            If Err.Number <> 0 Then
                errText = Err.Description
                condErrors = condErrors + 1
                LogMessage "WARN", "ImportServerRules: Subject.Enabled unreadable on '" & ruleName & "': " & errText
            End If
            On Error GoTo PROC_ERR

            If condEnabled Then
                subjArray = Empty
                On Error Resume Next
                subjArray = condObj.text
                If Err.Number <> 0 Then
                    errText = Err.Description
                    condErrors = condErrors + 1
                    LogMessage "WARN", "ImportServerRules: Subject.Text unreadable on '" & ruleName & "': " & errText
                End If
                On Error GoTo PROC_ERR

                If IsArray(subjArray) Then
                    For Each subj In subjArray
                        If Len(Trim(CStr(subj))) > 0 Then
                            RecordLearnedSubject Trim(CStr(subj)), "DELETE"
                            subjectCount = subjectCount + 1
                            LogMessage "INFO", "  Subject from rule '" & ruleName & "': " & CStr(subj)
                        End If
                    Next subj
                End If
            End If
        End If

        ' Conditions.SubjectOrBody.Text (some rules use "subject or body contains")
        Set condObj = Nothing
        On Error Resume Next
        Set condObj = rl.Conditions.SubjectOrBody
        If Err.Number <> 0 Then
            errText = Err.Description
            condErrors = condErrors + 1
            LogMessage "WARN", "ImportServerRules: Conditions.SubjectOrBody unreadable on '" & ruleName & "': " & errText
        End If
        On Error GoTo PROC_ERR

        If Not condObj Is Nothing Then
            condEnabled = False
            On Error Resume Next
            condEnabled = condObj.Enabled
            If Err.Number <> 0 Then
                errText = Err.Description
                condErrors = condErrors + 1
                LogMessage "WARN", "ImportServerRules: SubjectOrBody.Enabled unreadable on '" & ruleName & "': " & errText
            End If
            On Error GoTo PROC_ERR

            If condEnabled Then
                sobArray = Empty
                On Error Resume Next
                sobArray = condObj.text
                If Err.Number <> 0 Then
                    errText = Err.Description
                    condErrors = condErrors + 1
                    LogMessage "WARN", "ImportServerRules: SubjectOrBody.Text unreadable on '" & ruleName & "': " & errText
                End If
                On Error GoTo PROC_ERR

                If IsArray(sobArray) Then
                    For Each sob In sobArray
                        If Len(Trim(CStr(sob))) > 0 Then
                            RecordLearnedSubject Trim(CStr(sob)), "DELETE"
                            subjectCount = subjectCount + 1
                            LogMessage "INFO", "  SubjectOrBody from rule '" & ruleName & "': " & CStr(sob)
                        End If
                    Next sob
                End If
            End If
        End If

NextRule:
    Next i

    LogMessage "INFO", "ImportServerRules complete: " & ruleCount & " rules, " & senderCount & _
               " senders, " & subjectCount & " subjects, " & condErrors & " read errors"

    ImportServerRulesCore = "Processed " & ruleCount & " enabled rules (skipped " & skippedCount & " disabled): " & _
                            "imported " & senderCount & " senders and " & subjectCount & " subjects as learned DELETE rules." & _
                            IIf(condErrors > 0, " Property read errors: " & condErrors & " (see log).", "") & vbCrLf & _
                            "Total learned sender rules: " & GetLearnedSendersCount() & _
                            ", subject rules: " & GetLearnedSubjectsCount() & "."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ImportServerRulesCore", Err.Number, Err.Description
    ImportServerRulesCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' EXPORT LEARNED RULES TO SERVER-SIDE OUTLOOK RULES
'-------------------------------------------------------------------------------

' Export learned DELETE rules as server-side Outlook Rules (interactive wrapper)
Public Sub ExportLearnedRulesToServer()
    Dim colRules As Object          ' Outlook.Rules
    Dim response As VbMsgBoxResult
    Dim existingCount As Long
    Dim i As Long
    Dim senderCount As Long
    Dim subjectCount As Long
    Dim removeExisting As Boolean
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ExportLearnedRulesToServer"

    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled." & vbCrLf & _
               "Set EnableSelfImproving=True in settings.ini.", _
               vbInformation, "Export Learned Rules"
        GoTo PROC_EXIT
    End If

    ' Load caches so the confirmation can show counts
    LoadLearnedSenders
    LoadLearnedSubjects

    senderCount = GetLearnedDeleteSenders().Count
    subjectCount = GetLearnedSubjectKeys().Count

    If senderCount = 0 And subjectCount = 0 Then
        MsgBox "No learned DELETE rules to export." & vbCrLf & vbCrLf & _
               "Drag emails into '" & RuntimeFolderLearnDelete & "' (sender) or " & _
               "'" & RuntimeFolderLearnSubject & "' (subject) to learn DELETE rules.", _
               vbInformation, "Export Learned Rules"
        GoTo PROC_EXIT
    End If

    response = MsgBox("Export learned DELETE rules to server-side Outlook Rules?" & vbCrLf & vbCrLf & _
                      "Sender DELETE rules: " & senderCount & vbCrLf & _
                      "Subject DELETE rules: " & subjectCount & vbCrLf & vbCrLf & _
                      "Server rules run on Exchange even when Outlook is closed." & vbCrLf & _
                      "Action: Delete (to Deleted Items, recoverable)." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Export Learned Rules to Server")

    If response <> vbYes Then GoTo PROC_EXIT

    ' Check for existing export rules (read-only peek for the second dialog)
    Set colRules = Application.Session.DefaultStore.GetRules()
    existingCount = 0
    For i = colRules.Count To 1 Step -1
        If Left(colRules.Item(i).Name, 18) = "VBA Filter Export " Then
            existingCount = existingCount + 1
        End If
    Next i
    Set colRules = Nothing  ' Core gets its own fresh Rules collection

    removeExisting = True
    If existingCount > 0 Then
        response = MsgBox("Found " & existingCount & " existing 'VBA Filter Export' rule(s)." & vbCrLf & vbCrLf & _
                          "Remove them before creating new ones?" & vbCrLf & _
                          "(Recommended to avoid duplicates)", _
                          vbYesNoCancel + vbQuestion, "Existing Export Rules")

        If response = vbCancel Then GoTo PROC_EXIT
        removeExisting = (response = vbYes)
    End If

    resultText = ExportLearnedRulesToServerCore(removeExisting)
    MsgBox resultText, vbInformation, "Export Learned Rules to Server"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ExportLearnedRulesToServer", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: export learned DELETE rules as server-side Outlook Rules. No UI.
' removeExisting: True (default) removes previous "VBA Filter Export" rules first.
' Rollback safety: after a failed Create nothing is removed; after a failed
' configure, the new rule is removed BY NAME (never positional Remove, which
' could delete an unrelated pre-existing rule).
Public Function ExportLearnedRulesToServerCore(Optional ByVal removeExisting As Boolean = True) As String
    Dim colRules As Object          ' Outlook.Rules
    Dim newRule As Object           ' Outlook.Rule
    Dim deleteSenders As Collection
    Dim subjectKeys As Collection
    Dim senderCount As Long
    Dim subjectCount As Long
    Dim senderRuleCount As Long
    Dim subjectRuleCount As Long
    Dim existingCount As Long
    Dim i As Long
    Dim j As Long
    Dim batchStart As Long
    Dim batchEnd As Long
    Dim batchSize As Long
    Dim totalBatches As Long
    Dim batchNum As Long
    Dim ruleName As String
    Dim idx As Long
    Dim subjectSkipped As Long
    Dim cleanSubjects As Collection
    Dim rawSubj As String
    Dim cleanSubj As String
    Dim rawSenders As Collection
    Dim cleanSender As String
    Dim senderSkipped As Long
    Dim varArray() As Variant
    Dim createFailed As Boolean
    Dim failNum As Long
    Dim failDesc As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ExportLearnedRulesToServerCore"

    ' --- Step 1: Load caches and collect DELETE rules ---
    If Not RuntimeEnableSelfImproving Then
        ExportLearnedRulesToServerCore = "ERROR: self-improving filter is disabled (set EnableSelfImproving=True in settings.ini)"
        GoTo PROC_EXIT
    End If

    LoadLearnedSenders
    LoadLearnedSubjects

    Set subjectKeys = GetLearnedSubjectKeys()

    ' Sanitize sender list
    Set rawSenders = GetLearnedDeleteSenders()
    Set deleteSenders = New Collection
    For j = 1 To rawSenders.Count
        cleanSender = Trim(CStr(rawSenders(j)))
        If Len(cleanSender) > 0 And InStr(1, cleanSender, "@") > 0 Then
            deleteSenders.Add cleanSender
        Else
            If Len(cleanSender) > 0 Then
                LogMessage "WARN", "ExportLearnedRulesToServer: skipped invalid sender: """ & cleanSender & """"
            Else
                LogMessage "WARN", "ExportLearnedRulesToServer: skipped empty sender entry"
            End If
        End If
    Next j

    senderCount = deleteSenders.Count
    subjectCount = subjectKeys.Count

    If senderCount = 0 And subjectCount = 0 Then
        ExportLearnedRulesToServerCore = "No learned DELETE rules to export. Nothing created."
        GoTo PROC_EXIT
    End If

    ' --- Step 2: Get server rules; optionally remove existing exports ---
    Set colRules = Application.Session.DefaultStore.GetRules()

    existingCount = 0
    For i = colRules.Count To 1 Step -1
        If Left(colRules.Item(i).Name, 18) = "VBA Filter Export " Then
            existingCount = existingCount + 1
            If removeExisting Then colRules.Remove i
        End If
    Next i

    If removeExisting And existingCount > 0 Then
        LogMessage "INFO", "Removed " & existingCount & " existing VBA Filter Export rules"
    End If

    senderRuleCount = 0
    subjectRuleCount = 0

    ' --- Step 3: Create sender rules (batches of 50) ---
    senderSkipped = 0

    If senderCount > 0 Then
        batchSize = 50
        totalBatches = Int((senderCount - 1) / batchSize) + 1

        For batchNum = 1 To totalBatches
            batchStart = (batchNum - 1) * batchSize + 1
            batchEnd = batchStart + batchSize - 1
            If batchEnd > senderCount Then batchEnd = senderCount

            ' Build Variant array for this batch
            ReDim varArray(0 To batchEnd - batchStart)
            idx = 0
            For j = batchStart To batchEnd
                varArray(idx) = CStr(deleteSenders(j))
                idx = idx + 1
            Next j

            ' Create the rule with error handling per batch
            If totalBatches = 1 Then
                ruleName = "VBA Filter Export - Senders"
            Else
                ruleName = "VBA Filter Export - Senders (" & batchNum & "/" & totalBatches & ")"
            End If

            On Error Resume Next
            Err.Clear
            Set newRule = Nothing
            Set newRule = colRules.Create(ruleName, 0)  ' 0 = olRuleReceive
            If Not newRule Is Nothing Then
                newRule.Conditions.SenderAddress.Address = varArray
                newRule.Conditions.SenderAddress.Enabled = True
                newRule.Actions.Delete.Enabled = True
                newRule.Enabled = True
            End If
            createFailed = (Err.Number <> 0) Or (newRule Is Nothing)
            failNum = Err.Number
            failDesc = Err.Description
            On Error GoTo PROC_ERR

            If createFailed Then
                LogMessage "WARN", "Sender batch " & batchNum & " FAILED: Error " & failNum & ": " & failDesc
                senderSkipped = senderSkipped + (batchEnd - batchStart + 1)
                ' Roll back ONLY if Create actually added a rule; remove it by name.
                ' Never call Remove when Create itself failed (newRule Is Nothing).
                If Not newRule Is Nothing Then
                    RemoveRuleByName colRules, ruleName
                End If
            Else
                senderRuleCount = senderRuleCount + 1
                LogMessage "INFO", "Created rule: " & ruleName & " (" & (batchEnd - batchStart + 1) & " senders)"
            End If
        Next batchNum
    End If

    ' --- Step 4: Sanitize subject keys ---
    Set cleanSubjects = New Collection
    For j = 1 To subjectKeys.Count
        rawSubj = subjectKeys(j)
        cleanSubj = rawSubj
        cleanSubj = Replace(cleanSubj, vbCrLf, " ")
        cleanSubj = Replace(cleanSubj, vbCr, " ")
        cleanSubj = Replace(cleanSubj, vbLf, " ")
        cleanSubj = Replace(cleanSubj, vbTab, " ")
        cleanSubj = Replace(cleanSubj, Chr(0), "")
        cleanSubj = Replace(cleanSubj, "|", " ")
        cleanSubj = Trim(cleanSubj)
        If Len(cleanSubj) > 255 Then cleanSubj = Left(cleanSubj, 255)
        If Len(cleanSubj) > 0 Then
            cleanSubjects.Add cleanSubj
        Else
            LogMessage "WARN", "ExportLearnedRulesToServer: skipped empty/invalid subject key"
        End If
    Next j

    subjectCount = cleanSubjects.Count

    ' --- Step 5: Create subject rules (one per subject for robustness) ---
    subjectSkipped = 0

    If subjectCount > 0 Then
        For j = 1 To subjectCount
            ruleName = "VBA Filter Export - Subject " & j

            On Error Resume Next
            Err.Clear
            Set newRule = Nothing
            Set newRule = colRules.Create(ruleName, 0)
            If Not newRule Is Nothing Then
                newRule.Conditions.subject.text = Array(cleanSubjects(j))
                newRule.Conditions.subject.Enabled = True
                newRule.Actions.Delete.Enabled = True
                newRule.Enabled = True
            End If
            createFailed = (Err.Number <> 0) Or (newRule Is Nothing)
            failNum = Err.Number
            failDesc = Err.Description

            If createFailed And Not newRule Is Nothing Then
                ' Fallback: try SubjectOrBody on the same (already created) rule
                LogMessage "WARN", "Subject rule FAILED (Subject.Text): """ & Left(cleanSubjects(j), 60) & _
                           """ - Error " & failNum & ": " & failDesc
                Err.Clear
                newRule.Conditions.SubjectOrBody.text = Array(cleanSubjects(j))
                newRule.Conditions.SubjectOrBody.Enabled = True
                newRule.Actions.Delete.Enabled = True
                newRule.Enabled = True
                createFailed = (Err.Number <> 0)
                failNum = Err.Number
                failDesc = Err.Description
            End If
            On Error GoTo PROC_ERR

            If createFailed Then
                LogMessage "WARN", "Subject rule FAILED: """ & Left(cleanSubjects(j), 60) & _
                           """ - Error " & failNum & ": " & failDesc
                subjectSkipped = subjectSkipped + 1
                ' Roll back ONLY if Create actually added a rule; remove it by name.
                If Not newRule Is Nothing Then
                    RemoveRuleByName colRules, ruleName
                End If
            Else
                subjectRuleCount = subjectRuleCount + 1
            End If
        Next j
    End If

    ' --- Step 6: Save all rules to server ---
    On Error Resume Next
    Err.Clear
    colRules.Save
    failNum = Err.Number
    failDesc = Err.Description
    On Error GoTo PROC_ERR

    If failNum <> 0 Then
        LogMessage "ERROR", "colRules.Save failed: Error " & failNum & ": " & failDesc
        ExportLearnedRulesToServerCore = "ERROR: rules were created but Save to server failed: " & failDesc & _
                                         " (sender rules: " & senderRuleCount & ", subject rules: " & subjectRuleCount & _
                                         "). If the Rules dialog is open, close it and try again."
        GoTo PROC_EXIT
    End If

    LogMessage "INFO", "ExportLearnedRulesToServer complete: " & senderRuleCount & " sender rules, " & _
               subjectRuleCount & " subject rules saved"

    ExportLearnedRulesToServerCore = "Created " & senderRuleCount & " sender rule(s) (" & senderCount & " addresses)" & _
                                     IIf(senderSkipped > 0, ", " & senderSkipped & " addresses skipped", "") & _
                                     " and " & subjectRuleCount & " subject rule(s) (" & subjectCount & " keywords)" & _
                                     IIf(subjectSkipped > 0, ", " & subjectSkipped & " skipped", "") & "." & _
                                     IIf(removeExisting And existingCount > 0, _
                                         " Removed " & existingCount & " previous export rule(s).", "") & _
                                     " Verify in: Home -> Rules -> Manage Rules & Alerts."

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ExportLearnedRulesToServerCore", Err.Number, Err.Description
    ExportLearnedRulesToServerCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' FOLDER MIGRATION (v1.x -> v2.0)
'-------------------------------------------------------------------------------

' Detect old-style folder names (I, II, III, IIII, V) and offer to rename them.
' (interactive wrapper)
Public Sub DetectAndMigrateOldFolders()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim oldNames As Variant
    Dim newNames As Variant
    Dim i As Long
    Dim detectedMsg As String
    Dim detectedCount As Long
    Dim response As VbMsgBoxResult
    Dim testFolder As Outlook.Folder
    Dim existingNew As Outlook.Folder
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.DetectAndMigrateOldFolders"

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Old -> New name mapping
    oldNames = Array("I", "II", "III", "IIII", "V")
    newNames = Array(RuntimeFolderReview, RuntimeFolderProtected, RuntimeFolderLearnKeep, RuntimeFolderLearnDelete, RuntimeFolderLearnSubject)

    ' Read-only detection pass for the confirmation dialog
    detectedMsg = ""
    detectedCount = 0

    For i = 0 To UBound(oldNames)
        Set testFolder = Nothing
        On Error Resume Next
        Set testFolder = inbox.Folders(CStr(oldNames(i)))
        On Error GoTo PROC_ERR

        If Not testFolder Is Nothing Then
            ' Check if the new name folder already exists
            Set existingNew = Nothing
            On Error Resume Next
            Set existingNew = inbox.Folders(CStr(newNames(i)))
            On Error GoTo PROC_ERR

            If existingNew Is Nothing Then
                detectedCount = detectedCount + 1
                detectedMsg = detectedMsg & "  '" & oldNames(i) & "' -> '" & newNames(i) & "'" & vbCrLf
            Else
                detectedMsg = detectedMsg & "  '" & oldNames(i) & "' -> SKIPPED ('" & newNames(i) & "' already exists)" & vbCrLf
            End If
        End If
    Next i

    If detectedCount = 0 Then
        MsgBox "No old-style folders detected. Nothing to migrate.", vbInformation, "Folder Migration"
        GoTo PROC_EXIT
    End If

    response = MsgBox("Detected " & detectedCount & " old-style folder(s) to rename:" & vbCrLf & vbCrLf & _
                      detectedMsg & vbCrLf & _
                      "Rename them now?" & vbCrLf & vbCrLf & _
                      "After renaming, restart Outlook to refresh event handlers.", _
                      vbYesNo + vbQuestion, "Folder Migration (v1.x -> v2.0)")

    If response <> vbYes Then GoTo PROC_EXIT

    resultText = DetectAndMigrateOldFoldersCore()
    MsgBox resultText, vbInformation, "Folder Migration"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "DetectAndMigrateOldFolders", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: detect and rename v1.x folders (I/II/III/IIII/V). No UI.
Public Function DetectAndMigrateOldFoldersCore() As String
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim oldNames As Variant
    Dim newNames As Variant
    Dim i As Long
    Dim detectedCount As Long
    Dim renamedCount As Long
    Dim skippedCount As Long
    Dim failedCount As Long
    Dim testFolder As Outlook.Folder
    Dim existingNew As Outlook.Folder
    Dim renameFailed As Boolean
    Dim failDesc As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.DetectAndMigrateOldFoldersCore"

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Old -> New name mapping
    oldNames = Array("I", "II", "III", "IIII", "V")
    newNames = Array(RuntimeFolderReview, RuntimeFolderProtected, RuntimeFolderLearnKeep, RuntimeFolderLearnDelete, RuntimeFolderLearnSubject)

    detectedCount = 0
    renamedCount = 0
    skippedCount = 0
    failedCount = 0

    For i = 0 To UBound(oldNames)
        Set testFolder = Nothing
        On Error Resume Next
        Set testFolder = inbox.Folders(CStr(oldNames(i)))
        On Error GoTo PROC_ERR

        If Not testFolder Is Nothing Then
            detectedCount = detectedCount + 1

            Set existingNew = Nothing
            On Error Resume Next
            Set existingNew = inbox.Folders(CStr(newNames(i)))
            On Error GoTo PROC_ERR

            If existingNew Is Nothing Then
                renameFailed = False
                failDesc = ""
                On Error Resume Next
                testFolder.Name = CStr(newNames(i))
                renameFailed = (Err.Number <> 0)
                failDesc = Err.Description
                On Error GoTo PROC_ERR

                If Not renameFailed Then
                    renamedCount = renamedCount + 1
                    LogMessage "INFO", "Folder renamed: '" & oldNames(i) & "' -> '" & newNames(i) & "'"
                Else
                    failedCount = failedCount + 1
                    LogMessage "WARN", "Failed to rename '" & oldNames(i) & "': " & failDesc
                End If
            Else
                skippedCount = skippedCount + 1
                LogMessage "INFO", "Migration skipped '" & oldNames(i) & "': '" & newNames(i) & "' already exists"
            End If
        End If
    Next i

    If detectedCount = 0 Then
        DetectAndMigrateOldFoldersCore = "No old-style folders detected. Nothing to migrate."
    Else
        DetectAndMigrateOldFoldersCore = "Migration complete: " & renamedCount & " folder(s) renamed" & _
                                         IIf(skippedCount > 0, ", " & skippedCount & " skipped (target exists)", "") & _
                                         IIf(failedCount > 0, ", " & failedCount & " failed (see log)", "") & "." & _
                                         " Restart Outlook to refresh event handlers."
    End If

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "DetectAndMigrateOldFoldersCore", Err.Number, Err.Description
    DetectAndMigrateOldFoldersCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

' Display version information and system status (interactive wrapper)
Public Sub ShowVersionInfo()
    Dim resultText As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ShowVersionInfo"

    resultText = ShowVersionInfoCore()
    MsgBox resultText, vbInformation, "Email Agent Version Info"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ShowVersionInfo", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Headless core: build the full version/status text. No UI.
' The bridge (Bridge.bas DispatchMacroStd, case "ShowVersionInfo") calls this
' directly — it is the single version-info implementation.
Public Function ShowVersionInfoCore() As String
    Dim info As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ShowVersionInfoCore"

    info = "Email Agent v" & FILTER_VERSION & vbCrLf
    info = info & "Version Date: " & FILTER_VERSION_DATE & vbCrLf
    info = info & String(40, "-") & vbCrLf & vbCrLf

    info = info & "Settings file: " & GetSettingsFilePath() & vbCrLf
    info = info & "Learned senders file: " & GetLearnedSendersFilePath() & vbCrLf
    info = info & "Learned subjects file: " & GetLearnedSubjectsFilePath() & vbCrLf
    info = info & "Learned replies file: " & GetLearnedRepliesFilePath() & vbCrLf
    info = info & "Error log: " & RuntimeErrorLogFile & vbCrLf & vbCrLf

    info = info & "Learned sender rules: " & GetLearnedSendersCount() & vbCrLf
    info = info & "Learned subject rules: " & GetLearnedSubjectsCount() & vbCrLf & vbCrLf

    info = info & "Settings:" & vbCrLf
    info = info & "  Logging: " & IIf(RuntimeEnableLogging, "ON (" & RuntimeLogLevel & ")", "OFF") & vbCrLf
    info = info & "  Debug mode: " & IIf(RuntimeDebugMode, "ON", "OFF") & vbCrLf
    info = info & "  Self-improving: " & IIf(RuntimeEnableSelfImproving, "ON", "OFF") & vbCrLf
    info = info & "  LLM API: " & IIf(RuntimeUseLLM, "ON (provider: " & RuntimeLLMProvider & ")", "OFF") & vbCrLf
    info = info & "  Auto-reply: " & IIf(RuntimeEnableAutoReply, "ON", "OFF") & vbCrLf & vbCrLf

    info = info & "Folder names:" & vbCrLf
    info = info & "  Protected: " & RuntimeFolderProtected & vbCrLf
    info = info & "  Review: " & RuntimeFolderReview & vbCrLf
    info = info & "  LearnKeep: " & RuntimeFolderLearnKeep & vbCrLf
    info = info & "  LearnDelete: " & RuntimeFolderLearnDelete & vbCrLf
    info = info & "  LearnSubject: " & RuntimeFolderLearnSubject & vbCrLf
    info = info & "  LearnReply: " & RuntimeFolderLearnReply & vbCrLf

    ShowVersionInfoCore = info

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ShowVersionInfoCore", Err.Number, Err.Description
    ShowVersionInfoCore = "ERROR: " & Err.Description
    Resume PROC_EXIT
End Function

'-------------------------------------------------------------------------------
' PRIVATE HELPERS
'-------------------------------------------------------------------------------

' One-line stats summary for Core function return values.
Private Function SummarizeStats(ByVal stats As Object, ByVal totalCount As Long) As String
    Dim s As String

    s = "Filtered " & totalCount & " emails: " & _
        stats("DELETE") & " deleted, " & _
        stats("REVIEW") & " moved to " & RuntimeFolderReview & ", " & _
        stats("MOVE_II") & " moved to " & RuntimeFolderProtected & ", " & _
        stats("KEEP") & " kept."

    If stats("ERROR") > 0 Then
        s = s & " Errors: " & stats("ERROR") & "."
    End If

    SummarizeStats = s
End Function

' Count rule lines in a learned-rules file: non-blank, non-comment (#) lines only.
' Returns 0 if the file does not exist.
Private Function CountRuleLines(ByVal filePath As String) As Long
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim ruleCount As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.CountRuleLines"

    ruleCount = 0

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1)  ' ForReading
        Do While Not ts.AtEndOfStream
            line = Trim(ts.ReadLine)
            If Len(line) > 0 And Left(line, 1) <> "#" Then
                ruleCount = ruleCount + 1
            End If
        Loop
        ts.Close
        Set ts = Nothing
    End If
    Set fso = Nothing

    CountRuleLines = ruleCount

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "CountRuleLines", Err.Number, Err.Description
    CountRuleLines = 0
    ' Safe cleanup: after On Error Resume Next the error state is cleared,
    ' so jump with GoTo (Resume would raise "Resume without error")
    On Error Resume Next
    If Not ts Is Nothing Then ts.Close
    Set ts = Nothing
    Set fso = Nothing
    GoTo PROC_EXIT
End Function

' Remove a server rule by exact name match (reverse iteration).
' Safer than positional Remove after a failed Create, which could delete an
' unrelated pre-existing rule.
Private Sub RemoveRuleByName(ByVal colRules As Object, ByVal ruleName As String)
    Dim i As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RemoveRuleByName"

    For i = colRules.Count To 1 Step -1
        If StrComp(colRules.Item(i).Name, ruleName, vbTextCompare) = 0 Then
            colRules.Remove i
            LogMessage "INFO", "Rolled back rule: " & ruleName
            Exit For
        End If
    Next i

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "RemoveRuleByName", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Resolve a rule-condition Recipient to an SMTP address. Returns "" on failure.
' Each On Error Resume Next covers a single COM property access.
Private Function ResolveRecipientSmtp(ByVal recip As Object) As String
    Dim result As String
    Dim ae As Object
    Dim exUser As Object
    Dim userType As Long

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ResolveRecipientSmtp"

    result = ""

    Set ae = Nothing
    On Error Resume Next
    Set ae = recip.AddressEntry
    On Error GoTo PROC_ERR

    If Not ae Is Nothing Then
        userType = -1
        On Error Resume Next
        userType = ae.AddressEntryUserType
        On Error GoTo PROC_ERR

        If userType = 0 Then  ' olExchangeUserAddressEntry
            Set exUser = Nothing
            On Error Resume Next
            Set exUser = ae.GetExchangeUser
            On Error GoTo PROC_ERR

            If Not exUser Is Nothing Then
                On Error Resume Next
                result = LCase(exUser.PrimarySmtpAddress)
                On Error GoTo PROC_ERR
            End If
        Else
            On Error Resume Next
            result = LCase(ae.Address)
            On Error GoTo PROC_ERR
        End If
    End If

    ' Fall back to Address property
    If Len(result) = 0 Then
        On Error Resume Next
        result = LCase(recip.Address)
        On Error GoTo PROC_ERR
    End If

    ResolveRecipientSmtp = result

PROC_EXIT:
    PopCallStack
    Exit Function
PROC_ERR:
    LogError "BatchFilter", "ResolveRecipientSmtp", Err.Number, Err.Description
    ResolveRecipientSmtp = ""
    Resume PROC_EXIT
End Function
