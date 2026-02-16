<#
.SYNOPSIS
    Automated installer for Outlook Email Filter v2.0

.DESCRIPTION
    This PowerShell script automates the installation of the VBA-based email filter
    for Microsoft Outlook. It handles:
    - Prerequisite checks (Outlook version, admin rights)
    - Trust Center settings configuration via registry
    - VBA module import via COM automation
    - Folder creation via Outlook COM API
    - Settings initialization

.PARAMETER SourcePath
    Path to the folder containing the VBA .bas files. Defaults to .\src

.PARAMETER SkipTrustSettings
    Skip modifying Trust Center registry settings (use if already configured)

.PARAMETER Uninstall
    Remove all Email Filter modules from Outlook VBA project

.EXAMPLE
    .\Install-OutlookFilter.ps1
    Standard installation using default source path (.\src)

.EXAMPLE
    .\Install-OutlookFilter.ps1 -SourcePath "C:\Downloads\outlook-filter\src"
    Install from a custom source path

.EXAMPLE
    .\Install-OutlookFilter.ps1 -SkipTrustSettings
    Install without modifying Trust Center settings (if already enabled)

.EXAMPLE
    .\Install-OutlookFilter.ps1 -Uninstall
    Remove all Email Filter modules

.NOTES
    Version: 2.0.0
    Requires: Windows, Outlook 2016/2019/2021/365, PowerShell 5.1+
    Must be run with Outlook CLOSED
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$SourcePath = ".\src",

    [Parameter()]
    [switch]$SkipTrustSettings,

    [Parameter()]
    [switch]$Uninstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper Functions

function Write-Step {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "  ✓ $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "  ⚠ $Message" -ForegroundColor Yellow
}

function Write-Failure {
    param([string]$Message)
    Write-Host "  ✗ $Message" -ForegroundColor Red
}

function Test-OutlookInstalled {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $version = $outlook.Version
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return $version
    }
    catch {
        return $null
    }
}

function Test-OutlookRunning {
    return (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue) -ne $null
}

function Get-OutlookVersion {
    $version = Test-OutlookInstalled
    if ($null -eq $version) {
        return $null
    }

    $majorVersion = [int]($version.Split('.')[0])

    switch ($majorVersion) {
        16 { return "2016/2019/2021/365" }
        15 { return "2013" }
        14 { return "2010" }
        default { return "Unknown ($version)" }
    }
}

function Enable-TrustCenterSettings {
    Write-Step "Configuring Trust Center settings..."

    # Detect Outlook version for registry path
    $outlookVersion = Test-OutlookInstalled
    if ($null -eq $outlookVersion) {
        throw "Cannot detect Outlook version"
    }

    $majorVersion = $outlookVersion.Split('.')[0]
    $regPath = "HKCU:\Software\Microsoft\Office\$majorVersion.0\Outlook\Security"

    # Create registry path if it doesn't exist
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    # Enable macros (Level 1 = Enable all, Level 2 = Signed only, Level 3 = Notifications, Level 4 = Disable all)
    # Setting to 3 (with notifications) for security, can be changed to 1 manually
    Set-ItemProperty -Path $regPath -Name "Level" -Value 3 -Type DWord

    # Enable "Trust access to VBA project object model"
    Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord

    Write-Success "Trust Center settings configured (AccessVBOM enabled, Macro level set to notifications)"
    Write-Warning "You may need to change 'Macro Settings' to 'Enable all macros' manually if you want no prompts"
}

function Import-VBAModules {
    param(
        [string]$SourceFolder,
        [object]$OutlookApp
    )

    Write-Step "Importing VBA modules..."

    $vbProject = $OutlookApp.VBE.ActiveVBProject
    $modulesPath = Resolve-Path $SourceFolder

    # Get all .bas files except Installer and ThisOutlookSession
    $moduleFiles = Get-ChildItem -Path $modulesPath -Filter "*.bas" |
                   Where-Object { $_.BaseName -notin @("Installer", "ThisOutlookSession") }

    $importedCount = 0

    foreach ($file in $moduleFiles) {
        Write-Step "  Importing $($file.Name)..."

        # Remove existing module if present
        try {
            $existingModule = $vbProject.VBComponents.Item($file.BaseName)
            $vbProject.VBComponents.Remove($existingModule)
        }
        catch {
            # Module doesn't exist, that's fine
        }

        # Import the module
        $vbProject.VBComponents.Import($file.FullName) | Out-Null
        $importedCount++
    }

    # Handle ThisOutlookSession specially
    $thisSessionFile = Join-Path $modulesPath "ThisOutlookSession.bas"
    if (Test-Path $thisSessionFile) {
        Write-Step "  Configuring ThisOutlookSession..."

        $content = Get-Content $thisSessionFile -Raw

        # Extract code after "Option Explicit" (skip Attribute lines)
        $pattern = '(?s)Option\s+Explicit.*$'
        if ($content -match $pattern) {
            $code = $matches[0]

            $thisSession = $vbProject.VBComponents.Item("ThisOutlookSession")
            $codeModule = $thisSession.CodeModule

            # Clear existing code
            if ($codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, $codeModule.CountOfLines)
            }

            # Add new code
            $codeModule.AddFromString($code)
            $importedCount++
        }
    }

    # Try to import UserForms (may fail if .frx binary is missing)
    $formFiles = Get-ChildItem -Path $modulesPath -Filter "*.frm"
    foreach ($file in $formFiles) {
        Write-Step "  Attempting to import $($file.Name)..."

        # Remove existing form if present
        try {
            $existingForm = $vbProject.VBComponents.Item($file.BaseName)
            $vbProject.VBComponents.Remove($existingForm)
        }
        catch { }

        # Try to import (may fail without .frx)
        try {
            $vbProject.VBComponents.Import($file.FullName) | Out-Null
            $importedCount++
            Write-Success "    Imported $($file.Name)"
        }
        catch {
            Write-Warning "    Could not import $($file.Name) (binary .frx file may be missing)"
        }
    }

    Write-Success "Imported $importedCount module(s)"
}

function New-FilterFolders {
    param([object]$OutlookApp)

    Write-Step "Creating filter folders..."

    $namespace = $OutlookApp.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

    $folderNames = @("Review", "Protected", "LearnKeep", "LearnDelete", "LearnSubjectDelete")
    $createdCount = 0

    foreach ($folderName in $folderNames) {
        try {
            # Try to get folder
            $folder = $inbox.Folders.Item($folderName)
        }
        catch {
            # Folder doesn't exist, create it
            $inbox.Folders.Add($folderName, [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox) | Out-Null
            $createdCount++
        }
    }

    if ($createdCount -gt 0) {
        Write-Success "Created $createdCount folder(s)"
    }
    else {
        Write-Success "All folders already exist"
    }

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($inbox) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
}

function Initialize-SettingsFile {
    Write-Step "Initializing settings file..."

    $dataDir = Join-Path $env:APPDATA "OutlookEmailFilter"
    $settingsPath = Join-Path $dataDir "settings.ini"

    # Create directory if needed
    if (-not (Test-Path $dataDir)) {
        New-Item -Path $dataDir -ItemType Directory -Force | Out-Null
    }

    # Only create if doesn't exist (preserve existing settings)
    if (-not (Test-Path $settingsPath)) {
        $defaultSettings = @"
[General]
EnableLogging=true
LogLevel=INFO
EnableSelfImproving=true
ProgressInterval=50
DryRunLimit=100
LLMBatchSize=10

[Folders]
Protected=Protected
Review=Review
LearnKeep=LearnKeep
LearnDelete=LearnDelete
LearnSubject=LearnSubjectDelete

[Patterns]
ProtectedDomains=substack.com,medium.com
NamePatterns=Xu Xin,Professor Xu,Dr. Xu
GreetingPatterns=Dear Xu,Dear Professor,Dear Dr. Xu
PolyUTags=[MM],[HRO],[CPCE],[SPEED]
VIPKeywords=thesis,dissertation,deadline,urgent,grade
DeleteSenderPatterns=noreply,no-reply,do-not-reply,notifications,marketing,newsletter
DeleteKnownSenders=LinkedIn,Facebook,Twitter
DeleteSubjectPatterns=unsubscribe,newsletter,promotion,advertisement

[LLM]
UseLLMAPI=false
APIKeyMethod=ENV
APIKeyHardcoded=
APIEndpoint=https://your-resource.openai.azure.com/openai/deployments/your-deployment/chat/completions?api-version=2024-02-15-preview
MaxTokens=500
Temperature=0.3
SummarizeMaxTokens=300
ReplyMaxTokens=500
ReplyTemperature=0.7
"@

        Set-Content -Path $settingsPath -Value $defaultSettings -Encoding UTF8
        Write-Success "Created settings file: $settingsPath"
    }
    else {
        Write-Success "Settings file already exists (preserved): $settingsPath"
    }
}

function Remove-FilterModules {
    param([object]$OutlookApp)

    Write-Step "Removing Email Filter modules..."

    $vbProject = $OutlookApp.VBE.ActiveVBProject
    $modulesToRemove = @("Config", "Utilities", "EmailFilter", "BatchFilter", "Installer",
                         "frmFilterDashboard", "frmDraftReply")
    $removedCount = 0

    foreach ($moduleName in $modulesToRemove) {
        try {
            $module = $vbProject.VBComponents.Item($moduleName)
            $vbProject.VBComponents.Remove($module)
            $removedCount++
        }
        catch {
            # Module doesn't exist, skip
        }
    }

    # Clear ThisOutlookSession
    try {
        $thisSession = $vbProject.VBComponents.Item("ThisOutlookSession")
        $codeModule = $thisSession.CodeModule
        if ($codeModule.CountOfLines -gt 0) {
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
        }
    }
    catch { }

    Write-Success "Removed $removedCount module(s)"
}

#endregion

#region Main Script

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host "  Outlook Email Filter v2.0 - Automated Installer" -ForegroundColor Magenta
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
Write-Host ""

# Check prerequisites
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

# Check Outlook installation
$outlookVersion = Get-OutlookVersion
if ($null -eq $outlookVersion) {
    Write-Failure "Microsoft Outlook is not installed or not accessible"
    exit 1
}
Write-Success "Outlook version: $outlookVersion"

# Check if Outlook is running
if (Test-OutlookRunning) {
    Write-Failure "Outlook is currently running. Please close Outlook and run this script again."
    exit 1
}
Write-Success "Outlook is not running"

# Check source path (if not uninstalling)
if (-not $Uninstall) {
    if (-not (Test-Path $SourcePath)) {
        Write-Failure "Source path not found: $SourcePath"
        Write-Host ""
        Write-Host "Usage: .\Install-OutlookFilter.ps1 [-SourcePath <path>]" -ForegroundColor Yellow
        exit 1
    }
    Write-Success "Source path: $SourcePath"
}

Write-Host ""

# Handle uninstall
if ($Uninstall) {
    Write-Host "Starting uninstallation..." -ForegroundColor Yellow
    Write-Host ""

    try {
        $outlook = New-Object -ComObject Outlook.Application
        Remove-FilterModules -OutlookApp $outlook
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "  Uninstallation Complete!" -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host ""
        Write-Host "Data files (settings.ini, learned rules) were NOT deleted." -ForegroundColor Cyan
        Write-Host "You can now restart Outlook." -ForegroundColor Cyan
        Write-Host ""
    }
    catch {
        Write-Failure "Uninstallation failed: $_"
        Write-Host ""
        Write-Host $_.Exception.Message -ForegroundColor Red
        exit 1
    }

    exit 0
}

# Proceed with installation
Write-Host "Starting installation..." -ForegroundColor Yellow
Write-Host ""

try {
    # Step 1: Configure Trust Center (if not skipped)
    if (-not $SkipTrustSettings) {
        Enable-TrustCenterSettings
        Write-Host ""
    }

    # Step 2: Launch Outlook and import modules
    Write-Step "Launching Outlook COM automation..."
    $outlook = New-Object -ComObject Outlook.Application
    Write-Success "Outlook COM object created"
    Write-Host ""

    # Step 3: Import VBA modules
    Import-VBAModules -SourceFolder $SourcePath -OutlookApp $outlook
    Write-Host ""

    # Step 4: Create folders
    New-FilterFolders -OutlookApp $outlook
    Write-Host ""

    # Step 5: Initialize settings
    Initialize-SettingsFile
    Write-Host ""

    # Cleanup
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null

    # Success message
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Installation Complete!" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Cyan
    Write-Host "  1. Open Outlook" -ForegroundColor White
    Write-Host "  2. Press Alt+F11 to open VBA Editor" -ForegroundColor White
    Write-Host "  3. Go to Debug -> Compile Project (check for errors)" -ForegroundColor White
    Write-Host "  4. Press Ctrl+S to save" -ForegroundColor White
    Write-Host "  5. Close VBA Editor and restart Outlook" -ForegroundColor White
    Write-Host ""
    Write-Host "After restart, test with:" -ForegroundColor Cyan
    Write-Host "  Alt+F11 -> Ctrl+G -> FilterExistingDryRun" -ForegroundColor White
    Write-Host ""
    Write-Host "Configuration file:" -ForegroundColor Cyan
    Write-Host "  $env:APPDATA\OutlookEmailFilter\settings.ini" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Failure "Installation failed: $_"
    Write-Host ""
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""

    if ($_.Exception.Message -match "0x800A9C68") {
        Write-Warning "This error usually means 'Trust access to VBA project object model' is not enabled."
        Write-Host ""
        Write-Host "To fix manually:" -ForegroundColor Yellow
        Write-Host "  1. Open Outlook" -ForegroundColor White
        Write-Host "  2. File -> Options -> Trust Center -> Trust Center Settings" -ForegroundColor White
        Write-Host "  3. Macro Settings -> Check 'Trust access to VBA project object model'" -ForegroundColor White
        Write-Host "  4. Restart Outlook and run this script again" -ForegroundColor White
        Write-Host ""
    }

    exit 1
}

#endregion
