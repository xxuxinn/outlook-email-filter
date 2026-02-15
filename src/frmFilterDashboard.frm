VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFilterDashboard
   Caption         =   "Email Filter Dashboard"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9300
   OleObjectBlob   =   "frmFilterDashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilterDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' frmFilterDashboard - Email Filter Dashboard v2.0
'===============================================================================
' 4-tab dashboard UserForm (Filter Actions, Patterns, Settings, Learned Rules)
'
' SETUP INSTRUCTIONS:
' Since VBA UserForms are binary, you must create the form visually:
'
' 1. In VBA Editor: Insert -> UserForm
' 2. Set form properties: Name=frmFilterDashboard, Caption="Email Filter Dashboard",
'    Width=620, Height=480
' 3. Add a MultiPage control (Name=mpTabs) filling most of the form
'    - Page 0: "Filter Actions"
'    - Page 1: "Patterns"
'    - Page 2: "Settings"
'    - Page 3: "Learned Rules"
' 4. Add controls to each page as described in the comments below
' 5. Paste this entire code into the UserForm's code-behind module
'
' CONTROLS NEEDED:
'
' === Page 0: Filter Actions ===
' CommandButton: cmdDryRun        Caption="Preview (Dry Run)"
' CommandButton: cmdFilterInbox   Caption="Filter Inbox"
' CommandButton: cmdFilterAll     Caption="Filter All Folders"
' CommandButton: cmdFilterSelected Caption="Filter Selected"
' CommandButton: cmdFilterCurrent Caption="Filter Current Folder"
' CommandButton: cmdFilterLastN   Caption="Filter Last N Days"
' TextBox:       txtDays          Text="7" (Width=40)
' CommandButton: cmdImportRules   Caption="Import Server Rules"
' CommandButton: cmdExportRules   Caption="Export to Server"
' CommandButton: cmdSummarize     Caption="Summarize Email"
' CommandButton: cmdDraftReply    Caption="Draft Reply"
' CommandButton: cmdMigrateFolders Caption="Migrate Old Folders"
' TextBox:       txtStatus        MultiLine=True, ScrollBars=2 (Vertical), Height=120
' Label:         lblVersion       Caption="" (set at runtime)
'
' === Page 1: Patterns ===
' ComboBox:      cboCategory      Style=2 (Dropdown List)
' ListBox:       lstPatterns      Height=200
' TextBox:       txtNewPattern
' CommandButton: cmdAddPattern    Caption="Add"
' CommandButton: cmdEditPattern   Caption="Edit"
' CommandButton: cmdRemovePattern Caption="Remove"
' CommandButton: cmdSavePatterns  Caption="Save Patterns"
' Label:         lblPatternHelp   Caption="" (set at runtime)
'
' === Page 2: Settings ===
' CheckBox:      chkLogging       Caption="Enable Logging"
' ComboBox:      cboLogLevel      Style=2
' CheckBox:      chkSelfImproving Caption="Enable Self-Improving"
' CheckBox:      chkUseLLM        Caption="Enable LLM API"
' TextBox:       txtEndpoint
' ComboBox:      cboKeyMethod     Style=2
' TextBox:       txtKeyEnvVar
' TextBox:       txtMaxTokens
' TextBox:       txtTemperature
' TextBox:       txtFolderProtected
' TextBox:       txtFolderReview
' TextBox:       txtFolderLearnKeep
' TextBox:       txtFolderLearnDelete
' TextBox:       txtFolderLearnSubject
' CommandButton: cmdSaveSettings  Caption="Save Settings"
' CommandButton: cmdResetDefaults Caption="Reset to Defaults"
'
' === Page 3: Learned Rules ===
' OptionButton:  optSenders       Caption="Senders" Value=True
' OptionButton:  optSubjects      Caption="Subjects"
' TextBox:       txtSearch
' ListBox:       lstRules         Height=250
' Label:         lblRuleCount     Caption=""
' CommandButton: cmdDeleteRule    Caption="Delete Selected Rule"
' CommandButton: cmdCleanDuplicates Caption="Clean Duplicates"
' CommandButton: cmdRefreshRules  Caption="Refresh"
' CommandButton: cmdImportFolders Caption="Import from Folders"
'===============================================================================

Option Explicit

' Pattern category names and their corresponding Runtime* variable references
Private patternCategories As Variant
Private currentPatternCategory As Long

'-------------------------------------------------------------------------------
' FORM INITIALIZATION
'-------------------------------------------------------------------------------

Private Sub UserForm_Initialize()
    ' Ensure settings are loaded
    If Not RuntimeSettingsLoaded Then LoadAllSettings

    ' Setup pattern categories
    patternCategories = Array("Protected Domains", "Name Patterns", "Greeting Patterns", _
                              "PolyU Tags", "VIP Subject Keywords", "Delete Sender Patterns", _
                              "Delete Known Senders", "Delete Subject Patterns")

    ' Populate pattern category dropdown
    Dim i As Long
    For i = 0 To UBound(patternCategories)
        cboCategory.AddItem patternCategories(i)
    Next i
    cboCategory.ListIndex = 0

    ' Populate log level dropdown
    cboLogLevel.AddItem "DEBUG"
    cboLogLevel.AddItem "INFO"
    cboLogLevel.AddItem "WARN"
    cboLogLevel.AddItem "ERROR"

    ' Populate API key method dropdown
    cboKeyMethod.AddItem "ENV"
    cboKeyMethod.AddItem "HARDCODED"

    ' Load current settings into controls
    LoadSettingsToControls

    ' Load learned rules
    RefreshLearnedRules

    ' Set version label
    lblVersion.Caption = "v" & FILTER_VERSION & " (" & FILTER_VERSION_DATE & ")"

    ' Set initial status
    txtStatus.text = "Dashboard ready." & vbCrLf & _
                     "Learned senders: " & GetLearnedSendersCount() & vbCrLf & _
                     "Learned subjects: " & GetLearnedSubjectsCount()
End Sub

'-------------------------------------------------------------------------------
' SETTINGS -> CONTROLS
'-------------------------------------------------------------------------------

Private Sub LoadSettingsToControls()
    ' General
    chkLogging.value = RuntimeEnableLogging
    SelectComboItem cboLogLevel, RuntimeLogLevel
    chkSelfImproving.value = RuntimeEnableSelfImproving

    ' LLM
    chkUseLLM.value = RuntimeUseLLM
    txtEndpoint.text = RuntimeLLMEndpoint
    SelectComboItem cboKeyMethod, RuntimeAPIKeyMethod
    txtKeyEnvVar.text = RuntimeAPIKeyEnvVar
    txtMaxTokens.text = CStr(RuntimeLLMMaxTokens)
    txtTemperature.text = CStr(RuntimeLLMTemperature)

    ' Folders
    txtFolderProtected.text = RuntimeFolderProtected
    txtFolderReview.text = RuntimeFolderReview
    txtFolderLearnKeep.text = RuntimeFolderLearnKeep
    txtFolderLearnDelete.text = RuntimeFolderLearnDelete
    txtFolderLearnSubject.text = RuntimeFolderLearnSubject
End Sub

Private Sub SelectComboItem(ByRef cbo As MSForms.ComboBox, ByVal value As String)
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If LCase(cbo.List(i)) = LCase(value) Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    ' If not found, add it
    cbo.AddItem value
    cbo.ListIndex = cbo.ListCount - 1
End Sub

'-------------------------------------------------------------------------------
' TAB 0: FILTER ACTIONS
'-------------------------------------------------------------------------------

Private Sub cmdDryRun_Click()
    Me.Hide
    FilterExistingDryRun
    UpdateStatus "Dry run complete."
    Me.Show
End Sub

Private Sub cmdFilterInbox_Click()
    Me.Hide
    FilterExistingEmails
    UpdateStatus "Inbox filter complete."
    Me.Show
End Sub

Private Sub cmdFilterAll_Click()
    Me.Hide
    FilterAllFolders
    UpdateStatus "All folders filter complete."
    Me.Show
End Sub

Private Sub cmdFilterSelected_Click()
    Me.Hide
    FilterSelectedEmails
    UpdateStatus "Selected emails filtered."
    Me.Show
End Sub

Private Sub cmdFilterCurrent_Click()
    Me.Hide
    FilterCurrentFolder
    UpdateStatus "Current folder filtered."
    Me.Show
End Sub

Private Sub cmdFilterLastN_Click()
    Dim days As Integer
    If IsNumeric(txtDays.text) Then
        days = CInt(txtDays.text)
        If days > 0 Then
            Me.Hide
            FilterLastNDays days
            UpdateStatus "Last " & days & " days filtered."
            Me.Show
            Exit Sub
        End If
    End If
    MsgBox "Please enter a valid number of days.", vbExclamation
End Sub

Private Sub cmdImportRules_Click()
    Me.Hide
    ImportServerRules
    UpdateStatus "Server rules imported."
    RefreshLearnedRules
    Me.Show
End Sub

Private Sub cmdExportRules_Click()
    Me.Hide
    ExportLearnedRulesToServer
    UpdateStatus "Rules exported to server."
    Me.Show
End Sub

Private Sub cmdSummarize_Click()
    Me.Hide
    SummarizeSelectedEmail
    Me.Show
End Sub

Private Sub cmdDraftReply_Click()
    Me.Hide
    DraftReplyToSelected
    Me.Show
End Sub

Private Sub cmdMigrateFolders_Click()
    Me.Hide
    DetectAndMigrateOldFolders
    UpdateStatus "Folder migration complete."
    Me.Show
End Sub

Private Sub UpdateStatus(ByVal msg As String)
    txtStatus.text = Format(Now, "hh:nn:ss") & " - " & msg & vbCrLf & txtStatus.text
End Sub

'-------------------------------------------------------------------------------
' TAB 1: PATTERNS
'-------------------------------------------------------------------------------

Private Sub cboCategory_Change()
    LoadPatternsForCategory cboCategory.ListIndex
End Sub

Private Sub LoadPatternsForCategory(ByVal catIndex As Long)
    currentPatternCategory = catIndex
    lstPatterns.Clear

    Dim patterns As String
    patterns = GetPatternString(catIndex)

    If Len(patterns) = 0 Then Exit Sub

    Dim arr() As String
    arr = Split(patterns, ",")
    Dim p As Variant
    For Each p In arr
        If Len(Trim(p)) > 0 Then
            lstPatterns.AddItem Trim(p)
        End If
    Next p

    ' Update help text
    Select Case catIndex
        Case 0: lblPatternHelp.Caption = "Domains whose emails always go to Protected folder (e.g. substack.com)"
        Case 1: lblPatternHelp.Caption = "Name variations to detect personally-addressed emails"
        Case 2: lblPatternHelp.Caption = "Greeting lines at start of body (e.g. Dear Professor Xu)"
        Case 3: lblPatternHelp.Caption = "Organizational tags in subject (case-sensitive, e.g. [MM])"
        Case 4: lblPatternHelp.Caption = "VIP keywords in subject that trigger KEEP (e.g. thesis)"
        Case 5: lblPatternHelp.Caption = "Patterns in sender EMAIL address that trigger DELETE"
        Case 6: lblPatternHelp.Caption = "Known sender NAMES that trigger DELETE"
        Case 7: lblPatternHelp.Caption = "Keywords in subject that trigger DELETE"
    End Select
End Sub

Private Function GetPatternString(ByVal catIndex As Long) As String
    Select Case catIndex
        Case 0: GetPatternString = RuntimeProtectedDomains
        Case 1: GetPatternString = RuntimeNamePatterns
        Case 2: GetPatternString = RuntimeGreetingPatterns
        Case 3: GetPatternString = RuntimePolyUTags
        Case 4: GetPatternString = RuntimeVIPKeywords
        Case 5: GetPatternString = RuntimeDeleteSenderPatterns
        Case 6: GetPatternString = RuntimeDeleteKnownSenders
        Case 7: GetPatternString = RuntimeDeleteSubjectPatterns
        Case Else: GetPatternString = ""
    End Select
End Function

Private Sub SetPatternString(ByVal catIndex As Long, ByVal value As String)
    Select Case catIndex
        Case 0: RuntimeProtectedDomains = value
        Case 1: RuntimeNamePatterns = value
        Case 2: RuntimeGreetingPatterns = value
        Case 3: RuntimePolyUTags = value
        Case 4: RuntimeVIPKeywords = value
        Case 5: RuntimeDeleteSenderPatterns = value
        Case 6: RuntimeDeleteKnownSenders = value
        Case 7: RuntimeDeleteSubjectPatterns = value
    End Select
End Sub

Private Function GetPatternINIKey(ByVal catIndex As Long) As String
    Select Case catIndex
        Case 0: GetPatternINIKey = "ProtectedDomains"
        Case 1: GetPatternINIKey = "NamePatterns"
        Case 2: GetPatternINIKey = "GreetingPatterns"
        Case 3: GetPatternINIKey = "PolyUTags"
        Case 4: GetPatternINIKey = "VIPSubjectKeywords"
        Case 5: GetPatternINIKey = "DeleteSenderPatterns"
        Case 6: GetPatternINIKey = "DeleteKnownSenders"
        Case 7: GetPatternINIKey = "DeleteSubjectPatterns"
        Case Else: GetPatternINIKey = ""
    End Select
End Function

Private Sub cmdAddPattern_Click()
    Dim newPat As String
    newPat = Trim(txtNewPattern.text)
    If Len(newPat) = 0 Then
        MsgBox "Enter a pattern first.", vbExclamation
        Exit Sub
    End If
    lstPatterns.AddItem newPat
    txtNewPattern.text = ""
End Sub

Private Sub cmdEditPattern_Click()
    If lstPatterns.ListIndex < 0 Then
        MsgBox "Select a pattern to edit.", vbExclamation
        Exit Sub
    End If
    Dim newVal As String
    newVal = Trim(txtNewPattern.text)
    If Len(newVal) = 0 Then
        MsgBox "Enter the new value in the text box.", vbExclamation
        Exit Sub
    End If
    lstPatterns.List(lstPatterns.ListIndex) = newVal
    txtNewPattern.text = ""
End Sub

Private Sub cmdRemovePattern_Click()
    If lstPatterns.ListIndex < 0 Then
        MsgBox "Select a pattern to remove.", vbExclamation
        Exit Sub
    End If
    lstPatterns.RemoveItem lstPatterns.ListIndex
End Sub

Private Sub cmdSavePatterns_Click()
    ' Rebuild comma-separated string from ListBox
    Dim result As String
    Dim i As Long
    result = ""
    For i = 0 To lstPatterns.ListCount - 1
        If i > 0 Then result = result & ","
        result = result & lstPatterns.List(i)
    Next i

    ' Update Runtime variable
    SetPatternString currentPatternCategory, result

    ' Save to INI
    WriteINISetting "Patterns", GetPatternINIKey(currentPatternCategory), result

    UpdateStatus "Patterns saved for: " & patternCategories(currentPatternCategory)
    MsgBox "Patterns saved!", vbInformation
End Sub

'-------------------------------------------------------------------------------
' TAB 2: SETTINGS
'-------------------------------------------------------------------------------

Private Sub cmdSaveSettings_Click()
    ' General
    RuntimeEnableLogging = chkLogging.value
    RuntimeLogLevel = cboLogLevel.text
    RuntimeEnableSelfImproving = chkSelfImproving.value

    WriteINISetting "General", "EnableLogging", IIf(RuntimeEnableLogging, "True", "False")
    WriteINISetting "General", "LogLevel", RuntimeLogLevel
    WriteINISetting "General", "EnableSelfImproving", IIf(RuntimeEnableSelfImproving, "True", "False")

    ' LLM
    RuntimeUseLLM = chkUseLLM.value
    RuntimeLLMEndpoint = txtEndpoint.text
    RuntimeAPIKeyMethod = cboKeyMethod.text
    RuntimeAPIKeyEnvVar = txtKeyEnvVar.text

    If IsNumeric(txtMaxTokens.text) Then RuntimeLLMMaxTokens = CInt(txtMaxTokens.text)
    If IsNumeric(txtTemperature.text) Then RuntimeLLMTemperature = CDbl(txtTemperature.text)

    WriteINISetting "LLM", "UseLLMAPI", IIf(RuntimeUseLLM, "True", "False")
    WriteINISetting "LLM", "Endpoint", RuntimeLLMEndpoint
    WriteINISetting "LLM", "APIKeyMethod", RuntimeAPIKeyMethod
    WriteINISetting "LLM", "APIKeyEnvVar", RuntimeAPIKeyEnvVar
    WriteINISetting "LLM", "MaxTokens", CStr(RuntimeLLMMaxTokens)
    WriteINISetting "LLM", "Temperature", CStr(RuntimeLLMTemperature)

    ' Folders
    RuntimeFolderProtected = txtFolderProtected.text
    RuntimeFolderReview = txtFolderReview.text
    RuntimeFolderLearnKeep = txtFolderLearnKeep.text
    RuntimeFolderLearnDelete = txtFolderLearnDelete.text
    RuntimeFolderLearnSubject = txtFolderLearnSubject.text

    WriteINISetting "Folders", "Protected", RuntimeFolderProtected
    WriteINISetting "Folders", "Review", RuntimeFolderReview
    WriteINISetting "Folders", "LearnKeep", RuntimeFolderLearnKeep
    WriteINISetting "Folders", "LearnDelete", RuntimeFolderLearnDelete
    WriteINISetting "Folders", "LearnSubject", RuntimeFolderLearnSubject

    UpdateStatus "Settings saved."
    MsgBox "Settings saved to settings.ini!" & vbCrLf & vbCrLf & _
           "Note: Folder name changes take effect on next Outlook restart.", _
           vbInformation, "Settings Saved"
End Sub

Private Sub cmdResetDefaults_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Reset ALL settings to defaults?" & vbCrLf & vbCrLf & _
                      "This will overwrite your settings.ini file.", _
                      vbYesNo + vbExclamation, "Reset to Defaults")

    If response <> vbYes Then Exit Sub

    CreateDefaultSettingsFile
    LoadAllSettings
    LoadSettingsToControls
    LoadPatternsForCategory currentPatternCategory

    UpdateStatus "Settings reset to defaults."
    MsgBox "Settings reset to defaults.", vbInformation
End Sub

'-------------------------------------------------------------------------------
' TAB 3: LEARNED RULES
'-------------------------------------------------------------------------------

Private Sub optSenders_Click()
    RefreshLearnedRules
End Sub

Private Sub optSubjects_Click()
    RefreshLearnedRules
End Sub

Private Sub txtSearch_Change()
    RefreshLearnedRules
End Sub

Private Sub RefreshLearnedRules()
    lstRules.Clear

    Dim searchText As String
    searchText = LCase(Trim(txtSearch.text))

    Dim totalCount As Long
    totalCount = 0

    If optSenders.value Then
        ' Show sender rules
        Dim senderCache As Object
        Set senderCache = GetLearnedSendersCacheCopy()

        Dim sk As Variant
        For Each sk In senderCache.keys
            Dim senderEntry As String
            senderEntry = CStr(sk) & " | " & senderCache(sk)

            If Len(searchText) = 0 Or InStr(1, LCase(senderEntry), searchText, vbTextCompare) > 0 Then
                lstRules.AddItem senderEntry
                totalCount = totalCount + 1
            End If
        Next sk

        lblRuleCount.Caption = "Total: " & senderCache.Count & " sender rules" & _
                               IIf(Len(searchText) > 0, " (showing " & totalCount & ")", "")
    Else
        ' Show subject rules
        Dim subjectCache As Object
        Set subjectCache = GetLearnedSubjectsCacheCopy()

        Dim jk As Variant
        For Each jk In subjectCache.keys
            Dim subjectEntry As String
            subjectEntry = CStr(jk) & " | " & subjectCache(jk)

            If Len(searchText) = 0 Or InStr(1, LCase(subjectEntry), searchText, vbTextCompare) > 0 Then
                lstRules.AddItem subjectEntry
                totalCount = totalCount + 1
            End If
        Next jk

        lblRuleCount.Caption = "Total: " & subjectCache.Count & " subject rules" & _
                               IIf(Len(searchText) > 0, " (showing " & totalCount & ")", "")
    End If
End Sub

Private Sub cmdDeleteRule_Click()
    If lstRules.ListIndex < 0 Then
        MsgBox "Select a rule to delete.", vbExclamation
        Exit Sub
    End If

    Dim selected As String
    selected = lstRules.List(lstRules.ListIndex)

    ' Extract the key (everything before " | ")
    Dim pipePos As Long
    pipePos = InStr(1, selected, " | ")
    If pipePos = 0 Then Exit Sub

    Dim ruleKey As String
    ruleKey = Left(selected, pipePos - 1)

    Dim response As VbMsgBoxResult
    response = MsgBox("Delete this rule?" & vbCrLf & vbCrLf & ruleKey, _
                      vbYesNo + vbQuestion, "Delete Rule")

    If response <> vbYes Then Exit Sub

    If optSenders.value Then
        DeleteLearnedSenderRule ruleKey
    Else
        DeleteLearnedSubjectRule ruleKey
    End If

    RefreshLearnedRules
    UpdateStatus "Deleted rule: " & Left(ruleKey, 40)
End Sub

Private Sub cmdCleanDuplicates_Click()
    If optSenders.value Then
        DeduplicateLearnedSenders
    Else
        DeduplicateLearnedSubjects
    End If
    RefreshLearnedRules
    UpdateStatus "Duplicates cleaned."
    MsgBox "Duplicates cleaned!", vbInformation
End Sub

Private Sub cmdRefreshRules_Click()
    ' Force reload from file
    LoadLearnedSenders True
    LoadLearnedSubjects True
    RefreshLearnedRules
    UpdateStatus "Rules refreshed from file."
End Sub

Private Sub cmdImportFolders_Click()
    Me.Hide
    If optSenders.value Then
        ImportExistingLearnedFolders
    Else
        ImportExistingLearnedSubjectFolder
    End If
    RefreshLearnedRules
    UpdateStatus "Import from folders complete."
    Me.Show
End Sub
