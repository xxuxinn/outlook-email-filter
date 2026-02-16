Attribute VB_Name = "Installer"
'===============================================================================
' Installer.bas - Automated Installation Helper for Email Filter v2.0
'===============================================================================
' This module provides automated installation and import capabilities.
'
' PREREQUISITES:
' 1. Enable macros: File -> Options -> Trust Center -> Macro Settings
' 2. Enable VBA project access: Trust Center -> Trust access to VBA project object model
'
' USAGE:
' 1. Import this Installer.bas module only
' 2. Run InstallEmailFilter in the Immediate Window
' 3. Select the folder containing the other .bas files
' 4. The installer will import all modules and set up the filter
'
' For migration/reinstall, use: ImportFromFolder "C:\path\to\modules"
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' Main Installation Routine
'-------------------------------------------------------------------------------
Public Sub InstallEmailFilter()
    Dim folderPath As String
    Dim result As VbMsgBoxResult

    On Error GoTo ErrorHandler

    ' Show welcome message
    result = MsgBox( _
        "Outlook Email Filter v2.0 - Automated Installer" & vbCrLf & vbCrLf & _
        "This will:" & vbCrLf & _
        "  1. Import all VBA modules from a folder" & vbCrLf & _
        "  2. Set up ThisOutlookSession event handlers" & vbCrLf & _
        "  3. Create necessary folders" & vbCrLf & _
        "  4. Initialize settings.ini" & vbCrLf & vbCrLf & _
        "REQUIREMENTS:" & vbCrLf & _
        "  - Macros enabled" & vbCrLf & _
        "  - 'Trust access to VBA project object model' enabled" & vbCrLf & vbCrLf & _
        "Continue?", _
        vbQuestion + vbYesNo, _
        "Email Filter Installer")

    If result <> vbYes Then
        MsgBox "Installation cancelled.", vbInformation
        Exit Sub
    End If

    ' Prompt for source folder
    folderPath = SelectFolder("Select the folder containing the Email Filter .bas files")

    If folderPath = "" Then
        MsgBox "No folder selected. Installation cancelled.", vbInformation
        Exit Sub
    End If

    ' Run the installation
    ImportFromFolder folderPath

    Exit Sub

ErrorHandler:
    MsgBox "Installation error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure you've enabled 'Trust access to VBA project object model' in:" & vbCrLf & _
           "File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings", _
           vbCritical, "Installation Failed"
End Sub

'-------------------------------------------------------------------------------
' Import All Modules from a Folder
'-------------------------------------------------------------------------------
Public Sub ImportFromFolder(ByVal folderPath As String)
    Dim vbProj As Object
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim importedCount As Long
    Dim thisSessionCode As String
    Dim moduleName As String
    Dim ext As String
    Dim successMsg As String

    On Error GoTo ErrorHandler

    Set vbProj = Application.VBE.ActiveVBProject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbCritical
        Exit Sub
    End If

    Set folder = fso.GetFolder(folderPath)
    importedCount = 0

    ' First pass: Import all .bas modules except ThisOutlookSession
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        moduleName = fso.GetBaseName(file.Name)

        If ext = "bas" And moduleName <> "ThisOutlookSession" And moduleName <> "Installer" Then
            ' Remove existing module if present
            On Error Resume Next
            vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
            On Error GoTo ErrorHandler

            ' Import the module
            vbProj.VBComponents.Import file.Path
            importedCount = importedCount + 1
        End If
    Next file

    ' Second pass: Handle ThisOutlookSession specially
    Dim thisSessionPath As String
    thisSessionPath = folderPath & "\ThisOutlookSession.bas"

    If fso.FileExists(thisSessionPath) Then
        thisSessionCode = ReadTextFile(thisSessionPath)

        ' Extract code after "Option Explicit" (skip Attribute lines)
        Dim startPos As Long
        startPos = InStr(1, thisSessionCode, "Option Explicit", vbTextCompare)

        If startPos > 0 Then
            thisSessionCode = Mid(thisSessionCode, startPos)

            ' Get or create ThisOutlookSession
            Dim thisSession As Object
            Set thisSession = vbProj.VBComponents("ThisOutlookSession")

            ' Clear existing code
            With thisSession.CodeModule
                If .CountOfLines > 0 Then
                    .DeleteLines 1, .CountOfLines
                End If

                ' Insert new code
                .AddFromString thisSessionCode
            End With

            importedCount = importedCount + 1
        End If
    End If

    ' Import UserForms if present
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        moduleName = fso.GetBaseName(file.Name)

        If ext = "frm" Then
            ' Remove existing form if present
            On Error Resume Next
            vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
            On Error GoTo ErrorHandler

            ' Try to import (may fail if .frx binary is missing)
            On Error Resume Next
            vbProj.VBComponents.Import file.Path
            If Err.Number = 0 Then
                importedCount = importedCount + 1
            End If
            On Error GoTo ErrorHandler
        End If
    Next file

    Set fso = Nothing

    ' Create folders and initialize settings
    Call CreateFilterFolders
    Call InitializeSettings

    ' Build success message
    successMsg = "Installation Complete!" & vbCrLf & vbCrLf & _
                 "Imported " & importedCount & " module(s)" & vbCrLf & vbCrLf & _
                 "Next steps:" & vbCrLf & _
                 "1. Debug -> Compile Project (check for errors)" & vbCrLf & _
                 "2. Press Ctrl+S to save" & vbCrLf & _
                 "3. Close VBA Editor" & vbCrLf & _
                 "4. Restart Outlook" & vbCrLf & vbCrLf & _
                 "After restart, test with: FilterExistingDryRun"

    MsgBox successMsg, vbInformation, "Installation Successful"

    Exit Sub

ErrorHandler:
    MsgBox "Import error: " & Err.Description & vbCrLf & vbCrLf & _
           "Source: " & Err.Source, _
           vbCritical, "Import Failed"
End Sub

'-------------------------------------------------------------------------------
' Create Filter Folders
'-------------------------------------------------------------------------------
Private Sub CreateFilterFolders()
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.MAPIFolder
    Dim folderNames As Variant
    Dim folderName As Variant
    Dim createdCount As Long

    On Error Resume Next

    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Default folder names (will be overridden by settings.ini on restart)
    folderNames = Array("Review", "Protected", "LearnKeep", "LearnDelete", "LearnSubjectDelete")
    createdCount = 0

    For Each folderName In folderNames
        ' Try to get folder; create if it doesn't exist
        Dim targetFolder As Outlook.MAPIFolder
        Set targetFolder = Nothing
        Set targetFolder = inbox.Folders(CStr(folderName))

        If targetFolder Is Nothing Then
            inbox.Folders.Add CStr(folderName), olFolderInbox
            createdCount = createdCount + 1
        End If
    Next folderName

    Set inbox = Nothing
    Set ns = Nothing

    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Initialize Settings File
'-------------------------------------------------------------------------------
Private Sub InitializeSettings()
    Dim settingsPath As String
    Dim dataDir As String
    Dim fso As Object

    On Error Resume Next

    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create data directory
    dataDir = Environ("APPDATA") & "\OutlookEmailFilter"
    If Not fso.FolderExists(dataDir) Then
        fso.CreateFolder dataDir
    End If

    settingsPath = dataDir & "\settings.ini"

    ' Only create if doesn't exist (preserve existing settings)
    If Not fso.FileExists(settingsPath) Then
        Call WriteDefaultSettings(settingsPath)
    End If

    Set fso = Nothing

    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' Write Default Settings File
'-------------------------------------------------------------------------------
Private Sub WriteDefaultSettings(ByVal filePath As String)
    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, False)

    ts.WriteLine "[General]"
    ts.WriteLine "EnableLogging=true"
    ts.WriteLine "LogLevel=INFO"
    ts.WriteLine "EnableSelfImproving=true"
    ts.WriteLine "ProgressInterval=50"
    ts.WriteLine "DryRunLimit=100"
    ts.WriteLine "LLMBatchSize=10"
    ts.WriteLine ""
    ts.WriteLine "[Folders]"
    ts.WriteLine "Protected=Protected"
    ts.WriteLine "Review=Review"
    ts.WriteLine "LearnKeep=LearnKeep"
    ts.WriteLine "LearnDelete=LearnDelete"
    ts.WriteLine "LearnSubject=LearnSubjectDelete"
    ts.WriteLine ""
    ts.WriteLine "[Patterns]"
    ts.WriteLine "ProtectedDomains=substack.com,medium.com"
    ts.WriteLine "NamePatterns=Xu Xin,Professor Xu,Dr. Xu"
    ts.WriteLine "GreetingPatterns=Dear Xu,Dear Professor,Dear Dr. Xu"
    ts.WriteLine "PolyUTags=[MM],[HRO],[CPCE],[SPEED]"
    ts.WriteLine "VIPKeywords=thesis,dissertation,deadline,urgent,grade"
    ts.WriteLine "DeleteSenderPatterns=noreply,no-reply,do-not-reply,notifications,marketing,newsletter"
    ts.WriteLine "DeleteKnownSenders=LinkedIn,Facebook,Twitter"
    ts.WriteLine "DeleteSubjectPatterns=unsubscribe,newsletter,promotion,advertisement"
    ts.WriteLine ""
    ts.WriteLine "[LLM]"
    ts.WriteLine "UseLLMAPI=false"
    ts.WriteLine "APIKeyMethod=ENV"
    ts.WriteLine "APIKeyHardcoded="
    ts.WriteLine "APIEndpoint=https://your-resource.openai.azure.com/openai/deployments/your-deployment/chat/completions?api-version=2024-02-15-preview"
    ts.WriteLine "MaxTokens=500"
    ts.WriteLine "Temperature=0.3"
    ts.WriteLine "SummarizeMaxTokens=300"
    ts.WriteLine "ReplyMaxTokens=500"
    ts.WriteLine "ReplyTemperature=0.7"

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' Helper: Read Text File Contents
'-------------------------------------------------------------------------------
Private Function ReadTextFile(ByVal filePath As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim content As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(filePath) Then
        Set ts = fso.OpenTextFile(filePath, 1, False) ' ForReading
        content = ts.ReadAll
        ts.Close
    End If

    Set ts = Nothing
    Set fso = Nothing

    ReadTextFile = content
End Function

'-------------------------------------------------------------------------------
' Helper: Folder Browser Dialog
'-------------------------------------------------------------------------------
Private Function SelectFolder(ByVal dialogTitle As String) As String
    Dim shell As Object
    Dim folder As Object
    Dim selectedPath As String

    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0, dialogTitle, 0, 0)

    If Not folder Is Nothing Then
        selectedPath = folder.Self.Path
    Else
        selectedPath = ""
    End If

    Set folder = Nothing
    Set shell = Nothing

    SelectFolder = selectedPath
End Function

'-------------------------------------------------------------------------------
' Verify Installation
'-------------------------------------------------------------------------------
Public Sub VerifyInstallation()
    Dim vbProj As Object
    Dim requiredModules As Variant
    Dim moduleName As Variant
    Dim missingModules As String
    Dim foundCount As Long
    Dim settingsPath As String
    Dim fso As Object
    Dim msg As String

    Set vbProj = Application.VBE.ActiveVBProject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check for required modules
    requiredModules = Array("Config", "Utilities", "EmailFilter", "BatchFilter", "ThisOutlookSession")
    missingModules = ""
    foundCount = 0

    For Each moduleName In requiredModules
        On Error Resume Next
        Dim comp As Object
        Set comp = vbProj.VBComponents(CStr(moduleName))

        If comp Is Nothing Then
            missingModules = missingModules & "  - " & moduleName & vbCrLf
        Else
            foundCount = foundCount + 1
        End If
        On Error GoTo 0
    Next moduleName

    ' Check settings.ini
    settingsPath = Environ("APPDATA") & "\OutlookEmailFilter\settings.ini"
    Dim settingsExists As Boolean
    settingsExists = fso.FileExists(settingsPath)

    ' Build report
    msg = "Installation Verification" & vbCrLf & vbCrLf & _
          "VBA Modules: " & foundCount & " of " & UBound(requiredModules) + 1 & " found" & vbCrLf

    If Len(missingModules) > 0 Then
        msg = msg & vbCrLf & "Missing modules:" & vbCrLf & missingModules
    End If

    msg = msg & vbCrLf & "Settings file: "
    If settingsExists Then
        msg = msg & "OK (" & settingsPath & ")"
    Else
        msg = msg & "NOT FOUND (will be created on Outlook restart)"
    End If

    If foundCount = UBound(requiredModules) + 1 And settingsExists Then
        MsgBox msg & vbCrLf & vbCrLf & "Installation appears complete!", vbInformation
    Else
        MsgBox msg & vbCrLf & vbCrLf & "Installation incomplete. Run InstallEmailFilter to fix.", vbExclamation
    End If

    Set fso = Nothing
End Sub

'-------------------------------------------------------------------------------
' Uninstall / Remove All Modules
'-------------------------------------------------------------------------------
Public Sub UninstallEmailFilter()
    Dim result As VbMsgBoxResult
    Dim vbProj As Object
    Dim modulesToRemove As Variant
    Dim moduleName As Variant
    Dim removedCount As Long

    result = MsgBox( _
        "This will REMOVE all Email Filter modules from the VBA project." & vbCrLf & vbCrLf & _
        "Data files (settings.ini, learned rules) will NOT be deleted." & vbCrLf & vbCrLf & _
        "Continue?", _
        vbExclamation + vbYesNo, _
        "Uninstall Email Filter")

    If result <> vbYes Then Exit Sub

    Set vbProj = Application.VBE.ActiveVBProject
    modulesToRemove = Array("Config", "Utilities", "EmailFilter", "BatchFilter", _
                            "frmFilterDashboard", "frmDraftReply")
    removedCount = 0

    On Error Resume Next
    For Each moduleName In modulesToRemove
        vbProj.VBComponents.Remove vbProj.VBComponents(CStr(moduleName))
        If Err.Number = 0 Then removedCount = removedCount + 1
        Err.Clear
    Next moduleName

    ' Clear ThisOutlookSession
    Dim thisSession As Object
    Set thisSession = vbProj.VBComponents("ThisOutlookSession")
    With thisSession.CodeModule
        If .CountOfLines > 0 Then
            .DeleteLines 1, .CountOfLines
        End If
    End With

    On Error GoTo 0

    MsgBox "Uninstall complete. Removed " & removedCount & " module(s)." & vbCrLf & vbCrLf & _
           "Press Ctrl+S to save, then restart Outlook.", _
           vbInformation, "Uninstall Complete"
End Sub
