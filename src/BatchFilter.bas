'===============================================================================
' BatchFilter.bas - Batch Processing Functions v2.0
'===============================================================================
' This module contains functions for filtering existing emails in bulk,
' diagnostics, migration helpers, and the dashboard launcher.
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
End Sub

'-------------------------------------------------------------------------------
' FILTER EXISTING EMAILS
'-------------------------------------------------------------------------------

' Filter all existing emails in Inbox
Public Sub FilterExistingEmails()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim decision As String
    Dim totalCount As Long
    Dim response As VbMsgBoxResult

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
        Exit Sub
    End If

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
    LogMessage "INFO", "Starting filter of " & totalCount & " items in Inbox"

    ' Process from end to beginning (CRITICAL for deletions)
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats

            ' Progress indicator
            If (totalCount - i + 1) Mod RuntimeProgressInterval = 0 Then
                LogMessage "INFO", "Progress: " & (totalCount - i + 1) & " / " & totalCount
                DoEvents  ' Allow UI to update
            End If
        End If
    Next i

    LogMessage "INFO", "Filtering complete"
    MsgBox FormatStats(stats), vbInformation, "Email Filter Complete"
End Sub

'-------------------------------------------------------------------------------
' FILTER MULTIPLE FOLDERS
'-------------------------------------------------------------------------------

' Filter Inbox, Other folder, and any mounted PST archives
Public Sub FilterAllFolders()
    Dim ns As Outlook.NameSpace
    Dim store As Outlook.Store
    Dim response As VbMsgBoxResult

    response = MsgBox("This will filter emails in:" & vbCrLf & _
                      "  - Inbox" & vbCrLf & _
                      "  - 'Other' folder (Focused Inbox)" & vbCrLf & _
                      "  - Any mounted PST archives" & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbExclamation, "Email Filter - All Folders")

    If response <> vbYes Then Exit Sub

    Set ns = Application.GetNamespace("MAPI")

    ' 1. Filter default Inbox
    LogMessage "INFO", "Filtering default Inbox..."
    FilterFolder ns.GetDefaultFolder(olFolderInbox)

    ' 2. Filter "Other" folder (Focused Inbox feature)
    On Error Resume Next
    Dim otherFolder As Outlook.Folder
    Set otherFolder = ns.GetDefaultFolder(olFolderInbox).Folders("Other")
    If Not otherFolder Is Nothing Then
        LogMessage "INFO", "Filtering 'Other' folder..."
        FilterFolder otherFolder
    End If
    On Error GoTo 0

    ' 3. Filter mounted PST archives
    For Each store In ns.Stores
        If store.ExchangeStoreType = olNotExchange Then
            ' This is a PST file
            LogMessage "INFO", "Filtering PST: " & store.DisplayName
            On Error Resume Next
            FilterFolder store.GetDefaultFolder(olFolderInbox)
            On Error GoTo 0
        End If
    Next store

    MsgBox "All folders processed!", vbInformation, "Email Filter Complete"
End Sub

' Filter a specific folder
Public Sub FilterFolder(ByVal folder As Outlook.Folder)
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim decision As String

    If folder Is Nothing Then Exit Sub

    Set stats = CreateStatsDict()
    Set myItems = folder.Items

    LogMessage "INFO", "Filtering folder: " & folder.Name & " (" & myItems.Count & " items)"

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats
        End If
    Next i

    LogMessage "INFO", "Folder complete: " & FormatStats(stats)
End Sub

'-------------------------------------------------------------------------------
' FILTER BY DATE RANGE
'-------------------------------------------------------------------------------

' Filter emails received within a date range
Public Sub FilterByDateRange(ByVal startDate As Date, ByVal endDate As Date)
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim filter As String

    Set stats = CreateStatsDict()
    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    ' Build date filter
    filter = "[ReceivedTime] >= '" & Format(startDate, "mm/dd/yyyy") & "' AND " & _
             "[ReceivedTime] <= '" & Format(endDate + 1, "mm/dd/yyyy") & "'"

    Set filteredItems = myItems.Restrict(filter)

    LogMessage "INFO", "Filtering " & filteredItems.Count & " emails from " & _
               Format(startDate, "mm/dd/yyyy") & " to " & Format(endDate, "mm/dd/yyyy")

    For i = filteredItems.Count To 1 Step -1
        If TypeOf filteredItems(i) Is Outlook.MailItem Then
            Set mail = filteredItems(i)
            ExecuteAction mail, ClassifyEmail(mail), stats
        End If
    Next i

    MsgBox FormatStats(stats), vbInformation, "Date Range Filter Complete"
End Sub

' Filter emails from the last N days
Public Sub FilterLastNDays(ByVal days As Integer)
    FilterByDateRange DateAdd("d", -days, Date), Date
End Sub

'-------------------------------------------------------------------------------
' FILTER BY SENDER PATTERN (BULK DELETE)
'-------------------------------------------------------------------------------

' Delete all emails from senders matching a pattern
Public Sub BulkDeleteBySender(ByVal senderPattern As String)
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim deleteCount As Long
    Dim response As VbMsgBoxResult

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    ' Use Restrict for faster filtering
    Set filteredItems = myItems.Restrict("[SenderEmailAddress] Like '%" & senderPattern & "%'")

    If filteredItems.Count = 0 Then
        MsgBox "No emails found matching pattern: " & senderPattern, vbInformation
        Exit Sub
    End If

    response = MsgBox("Found " & filteredItems.Count & " emails from senders matching:" & vbCrLf & _
                      senderPattern & vbCrLf & vbCrLf & _
                      "Delete all?", vbYesNo + vbExclamation, "Bulk Delete")

    If response <> vbYes Then Exit Sub

    deleteCount = 0
    For i = filteredItems.Count To 1 Step -1
        If TypeOf filteredItems.Item(i) Is Outlook.MailItem Then
            filteredItems.Item(i).Delete
            deleteCount = deleteCount + 1
        End If
    Next i

    MsgBox "Deleted " & deleteCount & " emails.", vbInformation, "Bulk Delete Complete"
End Sub

'-------------------------------------------------------------------------------
' MOVE PROTECTED SOURCES (BULK)
'-------------------------------------------------------------------------------

' Move all emails from protected domains to the Protected folder
Public Sub MoveProtectedSources()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim moveCount As Long
    Dim domain As String

    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    moveCount = 0

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)
            domain = GetDomain(GetSenderEmail(mail))

            If ContainsAny(domain, RuntimeProtectedDomains) Then
                mail.Move GetOrCreateFolder(RuntimeFolderProtected)
                moveCount = moveCount + 1
            End If
        End If
    Next i

    MsgBox "Moved " & moveCount & " emails to '" & RuntimeFolderProtected & "' folder.", _
           vbInformation, "Protected Sources Moved"
End Sub

'-------------------------------------------------------------------------------
' STATISTICS AND REPORTING
'-------------------------------------------------------------------------------

' Generate a classification report without taking action
Public Sub GenerateClassificationReport()
    Dim myFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim stats As Object
    Dim decision As String
    Dim mailCount As Long
    Dim otherCount As Long

    Set stats = CreateStatsDict()
    Set myFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    Set myItems = myFolder.Items

    mailCount = 0
    otherCount = 0
    Dim learnedCount As Long
    Dim learnedSubjectCount As Long
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

    Dim learnedInfo As String
    If RuntimeEnableSelfImproving Then
        learnedInfo = vbCrLf & "Learned Rules:" & vbCrLf & _
                      "  Matched by learned sender rules: " & learnedCount & vbCrLf & _
                      "  Matched by learned subject rules: " & learnedSubjectCount & vbCrLf & _
                      "  Total learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
                      "  Total learned subject rules: " & GetLearnedSubjectsCount() & vbCrLf
    Else
        learnedInfo = ""
    End If

    MsgBox "Classification Report (No Actions Taken)" & vbCrLf & vbCrLf & _
           "Total items in folder: " & myItems.Count & vbCrLf & _
           "  - Emails (MailItems): " & mailCount & vbCrLf & _
           "  - Other (meetings, etc.): " & otherCount & vbCrLf & vbCrLf & _
           "Email Classification:" & vbCrLf & _
           "  Would DELETE: " & stats("DELETE") & vbCrLf & _
           "  Would MOVE to " & RuntimeFolderProtected & ": " & stats("MOVE_II") & vbCrLf & _
           "  Would MOVE to " & RuntimeFolderReview & ": " & stats("REVIEW") & vbCrLf & _
           "  Would KEEP: " & stats("KEEP") & vbCrLf & _
           learnedInfo & vbCrLf & _
           "Verification: " & (stats("DELETE") + stats("MOVE_II") + stats("REVIEW") + stats("KEEP")) & " = " & mailCount, _
           vbInformation, "Classification Report"
End Sub

'-------------------------------------------------------------------------------
' UNDO HELPERS
'-------------------------------------------------------------------------------

' Move emails from Review folder back to Inbox
Public Sub RestoreFromReview()
    Dim reviewFolder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim restoreCount As Long
    Dim response As VbMsgBoxResult

    On Error Resume Next
    Set reviewFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders(RuntimeFolderReview)
    On Error GoTo 0

    If reviewFolder Is Nothing Then
        MsgBox "'" & RuntimeFolderReview & "' folder not found.", vbInformation
        Exit Sub
    End If

    Set myItems = reviewFolder.Items

    If myItems.Count = 0 Then
        MsgBox "'" & RuntimeFolderReview & "' folder is empty.", vbInformation
        Exit Sub
    End If

    response = MsgBox("Move " & myItems.Count & " emails from '" & RuntimeFolderReview & "' back to Inbox?", _
                      vbYesNo + vbQuestion, "Restore from Review")

    If response <> vbYes Then Exit Sub

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    restoreCount = 0

    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            myItems(i).Move inbox
            restoreCount = restoreCount + 1
        End If
    Next i

    MsgBox "Restored " & restoreCount & " emails to Inbox.", vbInformation, "Restore Complete"
End Sub

'-------------------------------------------------------------------------------
' LEARNED SENDERS DIAGNOSTICS
'-------------------------------------------------------------------------------

' Display learned sender rules count and file location
Public Sub ShowLearnedSenders()
    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled." & vbCrLf & _
               "Set EnableSelfImproving=True in settings.ini.", _
               vbInformation, "Learned Senders"
        Exit Sub
    End If

    Dim filePath As String
    filePath = GetLearnedSendersFilePath()

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim fileInfo As String
    If fso.FileExists(filePath) Then
        Dim fileSize As Long
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
End Sub

' Dump all learned sender rules to the Immediate Window for review
Public Sub ShowLearnedSendersList()
    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Senders"
        Exit Sub
    End If

    ' Force reload to get freshest data
    LoadLearnedSenders True

    Dim count As Long
    count = GetLearnedSendersCount()

    If count = 0 Then
        MsgBox "No learned sender rules found.", vbInformation, "Learned Senders"
        Exit Sub
    End If

    ' Read the file directly to show all entries with timestamps
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim output As String
    Dim parts() As String
    Dim lineNum As Long
    Dim keepCount As Long
    Dim deleteCount As Long

    filePath = GetLearnedSendersFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "Learned senders file not found at:" & vbCrLf & filePath, vbExclamation
        Set fso = Nothing
        Exit Sub
    End If

    output = "=== LEARNED SENDER RULES ===" & vbCrLf
    output = output & "File: " & filePath & vbCrLf
    output = output & String(60, "=") & vbCrLf & vbCrLf

    Dim action As String
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
End Sub

' Remove duplicate entries from the learned senders file and show results
Public Sub CleanLearnedSendersFile()
    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Senders"
        Exit Sub
    End If

    Dim countBefore As Long
    countBefore = GetLearnedSendersCount()

    ' Read line count before dedup
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim linesBefore As Long

    filePath = GetLearnedSendersFilePath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1)
        linesBefore = 0
        Do While Not ts.AtEndOfStream
            ts.ReadLine
            linesBefore = linesBefore + 1
        Loop
        ts.Close
        Set ts = Nothing
    Else
        linesBefore = 0
    End If
    Set fso = Nothing

    ' Run deduplication
    DeduplicateLearnedSenders

    Dim countAfter As Long
    countAfter = GetLearnedSendersCount()
    Dim removed As Long
    removed = linesBefore - countAfter

    If removed = 0 Then
        MsgBox "No duplicates found." & vbCrLf & vbCrLf & _
               "File has " & linesBefore & " entries, all unique.", _
               vbInformation, "Clean Learned Senders"
    Else
        MsgBox "Deduplication complete!" & vbCrLf & vbCrLf & _
               "Lines before: " & linesBefore & vbCrLf & _
               "Lines after:  " & countAfter & vbCrLf & _
               "Removed:      " & removed & " duplicate(s)" & vbCrLf & vbCrLf & _
               "Unique rules: " & countAfter, _
               vbInformation, "Clean Learned Senders"
    End If
End Sub

'-------------------------------------------------------------------------------
' RESTORE WRONGLY DELETED EMAILS
'-------------------------------------------------------------------------------

' Scan Deleted Items, re-classify each email, and move KEEP/MOVE_II emails back
Public Sub RestoreDeletedKeepEmails()
    Dim deletedFolder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim decision As String
    Dim i As Long
    Dim restoreCount As Long
    Dim moveIICount As Long
    Dim totalMail As Long
    Dim response As VbMsgBoxResult
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.RestoreDeletedKeepEmails"

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
        MsgBox "No emails in Deleted Items.", vbInformation, "Restore"
        GoTo PROC_EXIT
    End If

    response = MsgBox("Scan " & totalMail & " emails in Deleted Items?" & vbCrLf & vbCrLf & _
                      "Emails classified as KEEP will be moved back to Inbox." & vbCrLf & _
                      "Emails classified as MOVE_II will be moved to '" & RuntimeFolderProtected & "'." & vbCrLf & _
                      "All other emails stay in Deleted Items.", _
                      vbYesNo + vbQuestion, "Restore Deleted Keep Emails")

    If response <> vbYes Then GoTo PROC_EXIT

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)

    restoreCount = 0
    moveIICount = 0

    ' Reverse iteration - moves invalidate indices
    For i = myItems.Count To 1 Step -1
        If TypeOf myItems(i) Is Outlook.MailItem Then
            Set mail = myItems(i)

            ' Pre-capture in case of move
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

    MsgBox "Restore Complete!" & vbCrLf & vbCrLf & _
           "Moved to Inbox: " & restoreCount & vbCrLf & _
           "Moved to " & RuntimeFolderProtected & ": " & moveIICount & vbCrLf & _
           "Left in Deleted Items: " & (totalMail - restoreCount - moveIICount), _
           vbInformation, "Restore Deleted Keep Emails"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "RestoreDeletedKeepEmails", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' QUICK FILTER ACTIONS (assign to QAT or ribbon button)
'-------------------------------------------------------------------------------

' Filter selected email(s) in the active explorer window
Public Sub FilterSelectedEmails()
    Dim sel As Outlook.Selection
    Dim mail As Outlook.MailItem
    Dim stats As Object
    Dim decision As String
    Dim i As Long
    Dim mailCount As Long
    Dim response As VbMsgBoxResult

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterSelectedEmails"

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set sel = Application.ActiveExplorer.Selection

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

    Set stats = CreateStatsDict()

    ' Reverse iteration - deletions/moves invalidate indices
    For i = sel.Count To 1 Step -1
        If TypeOf sel(i) Is Outlook.MailItem Then
            Set mail = sel(i)
            decision = ClassifyEmail(mail)
            ExecuteAction mail, decision, stats
        End If
    Next i

    MsgBox FormatStats(stats), vbInformation, "Filter Selected Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterSelectedEmails", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

' Filter all emails in the folder currently displayed in the active explorer
Public Sub FilterCurrentFolder()
    Dim folder As Outlook.Folder
    Dim inbox As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim stats As Object
    Dim mail As Outlook.MailItem
    Dim decision As String
    Dim i As Long
    Dim response As VbMsgBoxResult
    Dim isNonInbox As Boolean
    Dim isReviewFolder As Boolean
    Dim senderName As String
    Dim subject As String

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.FilterCurrentFolder"

    ' Ensure learned caches are loaded before classifying
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    End If

    Set folder = Application.ActiveExplorer.CurrentFolder

    If folder Is Nothing Then
        MsgBox "No folder is currently active.", vbExclamation, "Filter Folder"
        GoTo PROC_EXIT
    End If

    Set inbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    isNonInbox = (folder.EntryID <> inbox.EntryID)
    isReviewFolder = (folder.Name = RuntimeFolderReview)

    Set myItems = folder.Items

    ' Confirmation dialog (different message for non-Inbox folders)
    If isReviewFolder Then
        response = MsgBox("Filter Review folder?" & vbCrLf & _
                          "(" & myItems.Count & " items)" & vbCrLf & vbCrLf & _
                          "Review folder detected." & vbCrLf & _
                          "Only DELETE rules will be applied." & vbCrLf & _
                          "Non-DELETE emails will stay in Review.", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    ElseIf isNonInbox Then
        response = MsgBox("Filter folder '" & folder.Name & "'?" & vbCrLf & _
                          "(" & myItems.Count & " items)" & vbCrLf & vbCrLf & _
                          "Non-Inbox folder detected." & vbCrLf & _
                          "KEEP/MOVE_II/REVIEW emails will be moved to Inbox." & vbCrLf & _
                          "DELETE emails will be deleted.", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    Else
        response = MsgBox("Filter folder '" & folder.Name & "'?" & vbCrLf & _
                          "(" & myItems.Count & " items)" & vbCrLf & vbCrLf & _
                          "Actions: Delete / Move to " & RuntimeFolderProtected & " / Move to " & RuntimeFolderReview & " / Keep", _
                          vbYesNo + vbQuestion, "Filter This Folder")
    End If

    If response <> vbYes Then GoTo PROC_EXIT

    Set stats = CreateStatsDict()

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
        End If
    Next i

    LogMessage "INFO", "Folder '" & folder.Name & "' complete: " & FormatStats(stats)
    MsgBox FormatStats(stats), vbInformation, "Filter '" & folder.Name & "' Complete"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "FilterCurrentFolder", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' LEARNED SENDERS: BULK IMPORT
'-------------------------------------------------------------------------------

' One-time import: scan existing emails in LearnKeep and LearnDelete folders
' and record all senders as learned rules. Run once after upgrading.
Public Sub ImportExistingLearnedFolders()
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

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Find the learning folders
    On Error Resume Next
    Set keepFolder = inbox.Folders(RuntimeFolderLearnKeep)
    Set deleteFolder = inbox.Folders(RuntimeFolderLearnDelete)
    On Error GoTo 0

    If keepFolder Is Nothing And deleteFolder Is Nothing Then
        MsgBox "Neither '" & RuntimeFolderLearnKeep & "' nor '" & RuntimeFolderLearnDelete & "' folder found under Inbox.", _
               vbExclamation, "Import Learned Folders"
        Exit Sub
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

    MsgBox "Import complete!" & vbCrLf & vbCrLf & _
           "From '" & RuntimeFolderLearnKeep & "': " & keepCount & " senders -> KEEP" & vbCrLf & _
           "From '" & RuntimeFolderLearnDelete & "': " & deleteCount & " senders -> DELETE" & vbCrLf & _
           IIf(skipCount > 0, "Skipped (no @ in address): " & skipCount & vbCrLf, "") & vbCrLf & _
           "Total unique rules now: " & GetLearnedSendersCount(), _
           vbInformation, "Import Learned Folders"
End Sub

'-------------------------------------------------------------------------------
' LEARNED SUBJECTS DIAGNOSTICS
'-------------------------------------------------------------------------------

' Dump all learned subject rules to the Immediate Window for review
Public Sub ShowLearnedSubjectsList()
    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Subjects"
        Exit Sub
    End If

    ' Force reload to get freshest data
    LoadLearnedSubjects True

    Dim count As Long
    count = GetLearnedSubjectsCount()

    If count = 0 Then
        MsgBox "No learned subject rules found." & vbCrLf & vbCrLf & _
               "Drag emails into the '" & RuntimeFolderLearnSubject & "' folder to learn subject-based DELETE rules.", _
               vbInformation, "Learned Subjects"
        Exit Sub
    End If

    ' Read the file directly to show all entries with timestamps
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim output As String
    Dim parts() As String
    Dim lineNum As Long
    Dim deleteCount As Long

    filePath = GetLearnedSubjectsFilePath()

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "Learned subjects file not found at:" & vbCrLf & filePath, vbExclamation
        Set fso = Nothing
        Exit Sub
    End If

    output = "=== LEARNED SUBJECT RULES ===" & vbCrLf
    output = output & "File: " & filePath & vbCrLf
    output = output & String(60, "=") & vbCrLf & vbCrLf

    Dim action As String
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
End Sub

' Remove duplicate entries from the learned subjects file and show results
Public Sub CleanLearnedSubjectsFile()
    If Not RuntimeEnableSelfImproving Then
        MsgBox "Self-improving filter is disabled.", vbInformation, "Learned Subjects"
        Exit Sub
    End If

    Dim countBefore As Long
    countBefore = GetLearnedSubjectsCount()

    ' Read line count before dedup
    Dim filePath As String
    Dim fso As Object
    Dim ts As Object
    Dim linesBefore As Long

    filePath = GetLearnedSubjectsFilePath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1)
        linesBefore = 0
        Do While Not ts.AtEndOfStream
            ts.ReadLine
            linesBefore = linesBefore + 1
        Loop
        ts.Close
        Set ts = Nothing
    Else
        linesBefore = 0
    End If
    Set fso = Nothing

    ' Run deduplication
    DeduplicateLearnedSubjects

    Dim countAfter As Long
    countAfter = GetLearnedSubjectsCount()
    Dim removed As Long
    removed = linesBefore - countAfter

    If removed = 0 Then
        MsgBox "No duplicates found." & vbCrLf & vbCrLf & _
               "File has " & linesBefore & " entries, all unique.", _
               vbInformation, "Clean Learned Subjects"
    Else
        MsgBox "Deduplication complete!" & vbCrLf & vbCrLf & _
               "Lines before: " & linesBefore & vbCrLf & _
               "Lines after:  " & countAfter & vbCrLf & _
               "Removed:      " & removed & " duplicate(s)" & vbCrLf & vbCrLf & _
               "Unique rules: " & countAfter, _
               vbInformation, "Clean Learned Subjects"
    End If
End Sub

' One-time import: scan existing emails in LearnSubjectDelete folder
' and record all subjects as learned DELETE rules.
Public Sub ImportExistingLearnedSubjectFolder()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim subjectFolder As Outlook.Folder
    Dim myItems As Outlook.Items
    Dim mail As Outlook.MailItem
    Dim i As Long
    Dim importCount As Long
    Dim skipCount As Long

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Find the folder
    On Error Resume Next
    Set subjectFolder = inbox.Folders(RuntimeFolderLearnSubject)
    On Error GoTo 0

    If subjectFolder Is Nothing Then
        MsgBox "'" & RuntimeFolderLearnSubject & "' folder not found under Inbox." & vbCrLf & vbCrLf & _
               "Create it first, then drag emails there to learn subject-based DELETE rules.", _
               vbExclamation, "Import Learned Subject Folder"
        Exit Sub
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

    MsgBox "Import complete!" & vbCrLf & vbCrLf & _
           "From '" & RuntimeFolderLearnSubject & "': " & importCount & " subjects -> DELETE" & vbCrLf & _
           IIf(skipCount > 0, "Skipped (empty subject): " & skipCount & vbCrLf, "") & vbCrLf & _
           "Total unique subject rules now: " & GetLearnedSubjectsCount(), _
           vbInformation, "Import Learned Subject Folder"
End Sub

'-------------------------------------------------------------------------------
' IMPORT SERVER-SIDE RULES AS LEARNED RULES
'-------------------------------------------------------------------------------

' Import server-side Outlook Rules as learned sender/subject DELETE rules.
Public Sub ImportServerRules()
    Dim rules As Object  ' Outlook.Rules
    Dim rl As Object     ' Outlook.Rule
    Dim i As Long
    Dim senderCount As Long
    Dim subjectCount As Long
    Dim ruleCount As Long
    Dim skippedCount As Long
    Dim response As VbMsgBoxResult

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ImportServerRules"

    ' Get server rules
    Set rules = Application.Session.DefaultStore.GetRules()

    If rules.Count = 0 Then
        MsgBox "No server-side rules found.", vbInformation, "Import Server Rules"
        GoTo PROC_EXIT
    End If

    ' Count enabled rules
    Dim enabledCount As Long
    enabledCount = 0
    For i = 1 To rules.Count
        If rules.Item(i).Enabled Then enabledCount = enabledCount + 1
    Next i

    response = MsgBox("Found " & rules.Count & " server rules (" & enabledCount & " enabled)." & vbCrLf & vbCrLf & _
                      "This will import sender addresses and subject keywords from" & vbCrLf & _
                      "ENABLED rules as learned DELETE rules." & vbCrLf & vbCrLf & _
                      "After import, you can manually delete the server rules" & vbCrLf & _
                      "via Home -> Rules -> Manage Rules & Alerts." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Import Server Rules")

    If response <> vbYes Then GoTo PROC_EXIT

    ' Declare loop variables at procedure level (VBA has no block scoping)
    Dim recip As Object
    Dim recipEmail As String
    Dim j As Long
    Dim addrArray As Variant
    Dim addr As Variant
    Dim subjArray As Variant
    Dim subj As Variant
    Dim sobArray As Variant
    Dim sob As Variant

    senderCount = 0
    subjectCount = 0
    ruleCount = 0
    skippedCount = 0

    For i = 1 To rules.Count
        Set rl = rules.Item(i)

        ' Only process enabled rules
        If Not rl.Enabled Then
            skippedCount = skippedCount + 1
            GoTo NextRule
        End If

        ruleCount = ruleCount + 1
        LogMessage "INFO", "Importing rule: " & rl.Name

        ' --- Extract SENDER conditions ---

        ' Conditions.From.Recipients (Exchange recipients)
        On Error Resume Next
        If Not rl.Conditions.From Is Nothing Then
            If rl.Conditions.From.Enabled Then
                For j = 1 To rl.Conditions.From.Recipients.Count
                    Set recip = rl.Conditions.From.Recipients.Item(j)

                    ' Try to resolve Exchange address to SMTP
                    recipEmail = ""
                    If Not recip.AddressEntry Is Nothing Then
                        If recip.AddressEntry.AddressEntryUserType = 0 Then  ' olExchangeUserAddressEntry
                            If Not recip.AddressEntry.GetExchangeUser Is Nothing Then
                                recipEmail = LCase(recip.AddressEntry.GetExchangeUser.PrimarySmtpAddress)
                            End If
                        Else
                            recipEmail = LCase(recip.AddressEntry.Address)
                        End If
                    End If

                    ' Fall back to Address property
                    If Len(recipEmail) = 0 Then
                        recipEmail = LCase(recip.Address)
                    End If

                    If Len(recipEmail) > 0 And InStr(1, recipEmail, "@") > 0 Then
                        RecordLearnedSender recipEmail, "DELETE"
                        senderCount = senderCount + 1
                        LogMessage "INFO", "  Sender from rule '" & rl.Name & "': " & recipEmail
                    End If
                Next j
            End If
        End If
        On Error GoTo PROC_ERR

        ' Conditions.SenderAddress.Address (string array of email addresses)
        On Error Resume Next
        If Not rl.Conditions.SenderAddress Is Nothing Then
            If rl.Conditions.SenderAddress.Enabled Then
                addrArray = rl.Conditions.SenderAddress.Address
                If IsArray(addrArray) Then
                    For Each addr In addrArray
                        If Len(addr) > 0 And InStr(1, CStr(addr), "@") > 0 Then
                            RecordLearnedSender LCase(CStr(addr)), "DELETE"
                            senderCount = senderCount + 1
                            LogMessage "INFO", "  SenderAddress from rule '" & rl.Name & "': " & CStr(addr)
                        End If
                    Next addr
                End If
            End If
        End If
        On Error GoTo PROC_ERR

        ' --- Extract SUBJECT conditions ---

        ' Conditions.Subject.Text (array of subject keywords)
        On Error Resume Next
        If Not rl.Conditions.subject Is Nothing Then
            If rl.Conditions.subject.Enabled Then
                subjArray = rl.Conditions.subject.text
                If IsArray(subjArray) Then
                    For Each subj In subjArray
                        If Len(Trim(CStr(subj))) > 0 Then
                            RecordLearnedSubject Trim(CStr(subj)), "DELETE"
                            subjectCount = subjectCount + 1
                            LogMessage "INFO", "  Subject from rule '" & rl.Name & "': " & CStr(subj)
                        End If
                    Next subj
                End If
            End If
        End If
        On Error GoTo PROC_ERR

        ' Conditions.SubjectOrBody.Text (some rules use "subject or body contains")
        On Error Resume Next
        If Not rl.Conditions.SubjectOrBody Is Nothing Then
            If rl.Conditions.SubjectOrBody.Enabled Then
                sobArray = rl.Conditions.SubjectOrBody.text
                If IsArray(sobArray) Then
                    For Each sob In sobArray
                        If Len(Trim(CStr(sob))) > 0 Then
                            RecordLearnedSubject Trim(CStr(sob)), "DELETE"
                            subjectCount = subjectCount + 1
                            LogMessage "INFO", "  SubjectOrBody from rule '" & rl.Name & "': " & CStr(sob)
                        End If
                    Next sob
                End If
            End If
        End If
        On Error GoTo PROC_ERR

NextRule:
    Next i

    MsgBox "Server Rule Import Complete!" & vbCrLf & vbCrLf & _
           "Rules processed: " & ruleCount & " (skipped " & skippedCount & " disabled)" & vbCrLf & _
           "Senders imported: " & senderCount & " -> learned DELETE" & vbCrLf & _
           "Subjects imported: " & subjectCount & " -> learned DELETE" & vbCrLf & vbCrLf & _
           "Total learned sender rules: " & GetLearnedSendersCount() & vbCrLf & _
           "Total learned subject rules: " & GetLearnedSubjectsCount() & vbCrLf & vbCrLf & _
           "You can now delete the server rules via:" & vbCrLf & _
           "Home -> Rules -> Manage Rules & Alerts", _
           vbInformation, "Import Server Rules"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ImportServerRules", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' EXPORT LEARNED RULES TO SERVER-SIDE OUTLOOK RULES
'-------------------------------------------------------------------------------

' Export learned DELETE rules (senders and subjects) as server-side Outlook Rules.
Public Sub ExportLearnedRulesToServer()
    Dim colRules As Object          ' Outlook.Rules
    Dim newRule As Object           ' Outlook.Rule
    Dim response As VbMsgBoxResult
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
    Dim batchArray() As String
    Dim idx As Long
    Dim subjectSkipped As Long
    Dim cleanSubjects As Collection
    Dim rawSubj As String
    Dim cleanSubj As String
    Dim rawSenders As Collection
    Dim cleanSender As String
    Dim senderSkipped As Long
    Dim varArray() As Variant

    On Error GoTo PROC_ERR
    PushCallStack "BatchFilter.ExportLearnedRulesToServer"

    ' --- Step 1: Load caches and collect DELETE rules ---
    If RuntimeEnableSelfImproving Then
        LoadLearnedSenders
        LoadLearnedSubjects
    Else
        MsgBox "Self-improving filter is disabled." & vbCrLf & _
               "Set EnableSelfImproving=True in settings.ini.", _
               vbInformation, "Export Learned Rules"
        GoTo PROC_EXIT
    End If

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
        MsgBox "No learned DELETE rules to export." & vbCrLf & vbCrLf & _
               "Drag emails into '" & RuntimeFolderLearnDelete & "' (sender) or " & _
               "'" & RuntimeFolderLearnSubject & "' (subject) to learn DELETE rules.", _
               vbInformation, "Export Learned Rules"
        GoTo PROC_EXIT
    End If

    ' --- Step 2: Confirmation dialog ---
    response = MsgBox("Export learned DELETE rules to server-side Outlook Rules?" & vbCrLf & vbCrLf & _
                      "Sender DELETE rules: " & senderCount & vbCrLf & _
                      "Subject DELETE rules: " & subjectCount & vbCrLf & vbCrLf & _
                      "Server rules run on Exchange even when Outlook is closed." & vbCrLf & _
                      "Action: Delete (to Deleted Items, recoverable)." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Export Learned Rules to Server")

    If response <> vbYes Then GoTo PROC_EXIT

    ' --- Step 3: Get server rules and check for existing exports ---
    Set colRules = Application.Session.DefaultStore.GetRules()

    ' Count existing VBA Filter Export rules
    existingCount = 0
    For i = colRules.Count To 1 Step -1
        If Left(colRules.Item(i).Name, 18) = "VBA Filter Export " Then
            existingCount = existingCount + 1
        End If
    Next i

    ' Offer to remove existing export rules
    If existingCount > 0 Then
        response = MsgBox("Found " & existingCount & " existing 'VBA Filter Export' rule(s)." & vbCrLf & vbCrLf & _
                          "Remove them before creating new ones?" & vbCrLf & _
                          "(Recommended to avoid duplicates)", _
                          vbYesNoCancel + vbQuestion, "Existing Export Rules")

        If response = vbCancel Then GoTo PROC_EXIT

        If response = vbYes Then
            ' Remove existing export rules (reverse iteration)
            For i = colRules.Count To 1 Step -1
                If Left(colRules.Item(i).Name, 18) = "VBA Filter Export " Then
                    colRules.Remove i
                End If
            Next i
            LogMessage "INFO", "Removed " & existingCount & " existing VBA Filter Export rules"
        End If
    End If

    senderRuleCount = 0
    subjectRuleCount = 0

    ' --- Step 4: Create sender rules (batches of 50) ---
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

            Set newRule = colRules.Create(ruleName, 0)  ' 0 = olRuleReceive
            newRule.Conditions.SenderAddress.Address = varArray
            newRule.Conditions.SenderAddress.Enabled = True
            newRule.Actions.Delete.Enabled = True
            newRule.Enabled = True

            If Err.Number <> 0 Then
                LogMessage "WARN", "Sender batch " & batchNum & " FAILED: Error " & Err.Number & ": " & Err.Description
                senderSkipped = senderSkipped + (batchEnd - batchStart + 1)
                On Error Resume Next
                colRules.Remove colRules.Count
                On Error Resume Next
            Else
                senderRuleCount = senderRuleCount + 1
                LogMessage "INFO", "Created rule: " & ruleName & " (" & (batchEnd - batchStart + 1) & " senders)"
            End If

            On Error GoTo PROC_ERR
        Next batchNum
    End If

    ' --- Step 5: Sanitize subject keys ---
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

    ' --- Step 6: Create subject rules (one per subject for robustness) ---
    subjectSkipped = 0

    If subjectCount > 0 Then
        For j = 1 To subjectCount
            On Error Resume Next
            Err.Clear

            ruleName = "VBA Filter Export - Subject " & j
            Set newRule = colRules.Create(ruleName, 0)

            newRule.Conditions.subject.text = Array(cleanSubjects(j))
            newRule.Conditions.subject.Enabled = True
            newRule.Actions.Delete.Enabled = True
            newRule.Enabled = True

            If Err.Number <> 0 Then
                LogMessage "WARN", "Subject rule FAILED (Subject.Text): """ & Left(cleanSubjects(j), 60) & """ - Error " & Err.Number & ": " & Err.Description
                Err.Clear

                ' Fallback: try SubjectOrBody
                newRule.Conditions.SubjectOrBody.text = Array(cleanSubjects(j))
                newRule.Conditions.SubjectOrBody.Enabled = True
                newRule.Actions.Delete.Enabled = True
                newRule.Enabled = True

                If Err.Number <> 0 Then
                    LogMessage "WARN", "Subject rule FAILED (SubjectOrBody too): """ & Left(cleanSubjects(j), 60) & """ - Error " & Err.Number & ": " & Err.Description
                    subjectSkipped = subjectSkipped + 1
                    On Error Resume Next
                    colRules.Remove colRules.Count
                    On Error Resume Next
                Else
                    subjectRuleCount = subjectRuleCount + 1
                End If
            Else
                subjectRuleCount = subjectRuleCount + 1
            End If

            On Error GoTo PROC_ERR
        Next j
    End If

    ' --- Step 7: Save all rules to server ---
    On Error Resume Next
    Err.Clear
    colRules.Save

    If Err.Number <> 0 Then
        LogMessage "ERROR", "colRules.Save failed: Error " & Err.Number & ": " & Err.Description
        MsgBox "Rules were created but Save to server failed:" & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               "Sender rules created: " & senderRuleCount & vbCrLf & _
               "Subject rules created: " & subjectRuleCount & vbCrLf & _
               "If the Rules dialog is open, close it and try again.", _
               vbExclamation, "Export Learned Rules"
        On Error GoTo 0
        GoTo PROC_EXIT
    End If
    On Error GoTo 0

    MsgBox "Export Complete!" & vbCrLf & vbCrLf & _
           "Server rules created:" & vbCrLf & _
           "  Sender rules: " & senderRuleCount & " (" & senderCount & " addresses)" & vbCrLf & _
           IIf(senderSkipped > 0, "  Sender addresses skipped (invalid): " & senderSkipped & vbCrLf, "") & _
           "  Subject rules: " & subjectRuleCount & " (" & subjectCount & " keywords)" & vbCrLf & _
           IIf(subjectSkipped > 0, "  Subject rules skipped (invalid): " & subjectSkipped & vbCrLf, "") & vbCrLf & _
           "Verify in: Home -> Rules -> Manage Rules & Alerts", _
           vbInformation, "Export Learned Rules to Server"

PROC_EXIT:
    PopCallStack
    Exit Sub
PROC_ERR:
    LogError "BatchFilter", "ExportLearnedRulesToServer", Err.Number, Err.Description
    Resume PROC_EXIT
End Sub

'-------------------------------------------------------------------------------
' FOLDER MIGRATION (v1.x -> v2.0)
'-------------------------------------------------------------------------------

' Detect old-style folder names (I, II, III, IIII, V) and offer to rename them.
Public Sub DetectAndMigrateOldFolders()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim oldNames As Variant
    Dim newNames As Variant
    Dim i As Long
    Dim detectedMsg As String
    Dim detectedCount As Long
    Dim response As VbMsgBoxResult

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Old -> New name mapping
    oldNames = Array("I", "II", "III", "IIII", "V")
    newNames = Array(RuntimeFolderReview, RuntimeFolderProtected, RuntimeFolderLearnKeep, RuntimeFolderLearnDelete, RuntimeFolderLearnSubject)

    detectedMsg = ""
    detectedCount = 0

    Dim testFolder As Outlook.Folder
    For i = 0 To UBound(oldNames)
        Set testFolder = Nothing
        On Error Resume Next
        Set testFolder = inbox.Folders(CStr(oldNames(i)))
        On Error GoTo 0

        If Not testFolder Is Nothing Then
            ' Check if the new name folder already exists
            Dim existingNew As Outlook.Folder
            Set existingNew = Nothing
            On Error Resume Next
            Set existingNew = inbox.Folders(CStr(newNames(i)))
            On Error GoTo 0

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
        Exit Sub
    End If

    response = MsgBox("Detected " & detectedCount & " old-style folder(s) to rename:" & vbCrLf & vbCrLf & _
                      detectedMsg & vbCrLf & _
                      "Rename them now?" & vbCrLf & vbCrLf & _
                      "After renaming, restart Outlook to refresh event handlers.", _
                      vbYesNo + vbQuestion, "Folder Migration (v1.x -> v2.0)")

    If response <> vbYes Then Exit Sub

    Dim renamedCount As Long
    renamedCount = 0

    For i = 0 To UBound(oldNames)
        Set testFolder = Nothing
        On Error Resume Next
        Set testFolder = inbox.Folders(CStr(oldNames(i)))
        On Error GoTo 0

        If Not testFolder Is Nothing Then
            Dim existNew As Outlook.Folder
            Set existNew = Nothing
            On Error Resume Next
            Set existNew = inbox.Folders(CStr(newNames(i)))
            On Error GoTo 0

            If existNew Is Nothing Then
                On Error Resume Next
                testFolder.Name = CStr(newNames(i))
                If Err.Number = 0 Then
                    renamedCount = renamedCount + 1
                    LogMessage "INFO", "Folder renamed: '" & oldNames(i) & "' -> '" & newNames(i) & "'"
                Else
                    LogMessage "WARN", "Failed to rename '" & oldNames(i) & "': " & Err.Description
                End If
                On Error GoTo 0
            End If
        End If
    Next i

    MsgBox "Migration complete!" & vbCrLf & vbCrLf & _
           "Folders renamed: " & renamedCount & vbCrLf & vbCrLf & _
           "Please restart Outlook to refresh event handlers.", _
           vbInformation, "Folder Migration"
End Sub

' Display version information and system status
Public Sub ShowVersionInfo()
    Dim info As String

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

    MsgBox info, vbInformation, "Email Agent Version Info"
End Sub
