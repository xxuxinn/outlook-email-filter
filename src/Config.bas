'===============================================================================
' Config.bas - Configuration Module for Email Filter v2.0
'===============================================================================
' This module contains:
'   1. DEFAULT_* constants: compile-time fallback values
'   2. Runtime* variables: loaded from settings.ini at startup
'   3. Version constants
'   4. Infrastructure constants (data file paths, not user-configurable)
'
' Users customize via settings.ini (or the Dashboard UI), NOT by editing this file.
' To reset to defaults, delete settings.ini and restart Outlook.
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' VERSION
'-------------------------------------------------------------------------------
Public Const FILTER_VERSION As String = "2.0.0"
Public Const FILTER_VERSION_DATE As String = "2026-02-15"

'-------------------------------------------------------------------------------
' RUNTIME VARIABLES (populated by LoadAllSettings at startup)
'-------------------------------------------------------------------------------

' General
Public RuntimeEnableLogging As Boolean
Public RuntimeLogLevel As String
Public RuntimeEnableSelfImproving As Boolean
Public RuntimeProgressInterval As Integer
Public RuntimeDryRunLimit As Integer
Public RuntimeLLMBatchSize As Integer

' Folder names
Public RuntimeFolderProtected As String
Public RuntimeFolderReview As String
Public RuntimeFolderLearnKeep As String
Public RuntimeFolderLearnDelete As String
Public RuntimeFolderLearnSubject As String

' Patterns
Public RuntimeProtectedDomains As String
Public RuntimeNamePatterns As String
Public RuntimeGreetingPatterns As String
Public RuntimePolyUTags As String
Public RuntimeVIPKeywords As String
Public RuntimeDeleteSenderPatterns As String
Public RuntimeDeleteKnownSenders As String
Public RuntimeDeleteSubjectPatterns As String

' LLM
Public RuntimeUseLLM As Boolean
Public RuntimeLLMEndpoint As String
Public RuntimeAPIKeyMethod As String
Public RuntimeAPIKeyEnvVar As String
Public RuntimeAPIKeyHardcoded As String
Public RuntimeLLMSystemPrompt As String
Public RuntimeLLMMaxTokens As Integer
Public RuntimeLLMTemperature As Double

' Flag: have settings been loaded yet?
Public RuntimeSettingsLoaded As Boolean

'-------------------------------------------------------------------------------
' DEFAULT CONSTANTS (fallbacks if settings.ini is missing or incomplete)
'-------------------------------------------------------------------------------

' LLM defaults
Public Const DEFAULT_USE_LLM_API As Boolean = False
Public Const DEFAULT_AZURE_OPENAI_ENDPOINT As String = "https://YOUR-RESOURCE.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview"
Public Const DEFAULT_API_KEY_METHOD As String = "ENV"
Public Const DEFAULT_API_KEY_ENV_VAR As String = "AZURE_OPENAI_KEY"
Public Const DEFAULT_API_KEY_HARDCODED As String = ""
Public Const DEFAULT_LLM_SYSTEM_PROMPT As String = "You are filtering emails for Professor Xu Xin at PolyU Hong Kong. " & _
    "Respond with ONLY 'DELETE' or 'KEEP' followed by a brief reason (max 10 words). " & _
    "DELETE: Generic broadcasts, announcements, FYI-only, mass CC, admin notices, promotional, newsletters. " & _
    "KEEP: Personally addressed, requires action/response, from students/collaborators, important deadlines."
Public Const DEFAULT_LLM_MAX_TOKENS As Integer = 100
Public Const DEFAULT_LLM_TEMPERATURE As Double = 0.3

' Folder name defaults (v2.0 readable names)
Public Const DEFAULT_FOLDER_PROTECTED As String = "Protected"
Public Const DEFAULT_FOLDER_REVIEW As String = "Review"
Public Const DEFAULT_FOLDER_LEARN_KEEP As String = "LearnKeep"
Public Const DEFAULT_FOLDER_LEARN_DELETE As String = "LearnDelete"
Public Const DEFAULT_FOLDER_LEARN_SUBJECT_DELETE As String = "LearnSubjectDelete"

' Pattern defaults
Public Const DEFAULT_PROTECTED_DOMAINS As String = "substack.com,reddit.com,redditmail.com"
Public Const DEFAULT_DELETE_SENDER_PATTERNS As String = "notice,noreply,notification,no-reply,marketing,promo,newsletter,digest,campaign,bulk,mailer,broadcast"
Public Const DEFAULT_DELETE_KNOWN_SENDERS As String = "LinkedIn Job Alerts,edX,Cathay Pacific,HKBN,MyLink,WIRED Daily,Coursera,Udemy,Medium Daily Digest,Twitter,Facebook,Instagram,TikTok"
Public Const DEFAULT_DELETE_SUBJECT_PATTERNS As String = "優惠,offer,digest,newsletter,unsubscribe,job alert,weekly roundup,daily digest,promotional,special offer,limited time,act now,don't miss"
Public Const DEFAULT_NAME_PATTERNS As String = "Xu Xin,Xin Xu,XuXin,Xuxin,Professor Xu,Prof. Xu,Prof Xu,Dr. Xu,Dr Xu,Mr. Xu,Mr Xu,Head Xu,Professor Xin,Prof. Xin,Prof Xin,Dr. Xin,Dr Xin,Mr. Xin,Mr Xin,Head Xin,Professor Xu Xin,Prof. Xu Xin,Prof Xu Xin,Dr. Xu Xin,Dr Xu Xin,Mr. Xu Xin,Mr Xu Xin,Head Xu Xin,Professor Xin Xu,Prof. Xin Xu,Prof Xin Xu,Dr. Xin Xu,Dr Xin Xu,Mr. Xin Xu,Mr Xin Xu,Head Xin Xu"
Public Const DEFAULT_GREETING_PATTERNS As String = "Dear Professor Xu,Dear Prof. Xu,Dear Prof Xu,Dear Dr. Xu,Dear Dr Xu,Dear Mr. Xu,Dear Mr Xu,Dear Head Xu,Dear Professor Xin,Dear Prof. Xin,Dear Prof Xin,Dear Dr. Xin,Dear Dr Xin,Dear Mr. Xin,Dear Mr Xin,Dear Head Xin,Dear Xin,Dear Xu,Hi Professor Xu,Hi Prof. Xu,Hi Prof Xu,Hi Dr. Xu,Hi Dr Xu,Hi Mr. Xu,Hi Mr Xu,Hi Head Xu,Hi Professor Xin,Hi Prof. Xin,Hi Prof Xin,Hi Dr. Xin,Hi Dr Xin,Hi Mr. Xin,Hi Mr Xin,Hi Head Xin,Hi Xin,Hi Xu,Hello Professor Xu,Hello Prof. Xu,Hello Prof Xu,Hello Dr. Xu,Hello Dr Xu,Hello Mr. Xu,Hello Mr Xu,Hello Head Xu,Hello Professor Xin,Hello Prof. Xin,Hello Prof Xin,Hello Dr. Xin,Hello Dr Xin,Hello Mr. Xin,Hello Mr Xin,Hello Head Xin,Hello Xin,Hello Xu,Dear Head,Dear Director"
Public Const DEFAULT_POLYU_TAGS As String = "[MM],[HRO],[CUS],ToXX"
Public Const DEFAULT_VIP_SUBJECT_KEYWORDS As String = "thesis,dissertation,supervision,urgent,deadline,review request,paper submission,grant,conference,publication,meeting request,appointment,interview"

' Logging defaults
Public Const DEFAULT_ENABLE_LOGGING As Boolean = True
Public Const DEFAULT_LOG_LEVEL As String = "INFO"

' Batch processing defaults
Public Const DEFAULT_LLM_BATCH_SIZE As Integer = 10
Public Const DEFAULT_PROGRESS_INTERVAL As Integer = 100
Public Const DEFAULT_DRY_RUN_LIMIT As Integer = 50

' Self-improving defaults
Public Const DEFAULT_ENABLE_SELF_IMPROVING As Boolean = True

'-------------------------------------------------------------------------------
' INFRASTRUCTURE CONSTANTS (not user-configurable via settings.ini)
'-------------------------------------------------------------------------------

' Subfolder under %APPDATA% for learned data storage
Public Const LEARNED_DATA_FOLDER As String = "OutlookEmailFilter"

' Filenames for learned rules (pipe-delimited, append-only)
Public Const LEARNED_SENDERS_FILE As String = "learned_senders.txt"
Public Const LEARNED_SUBJECTS_FILE As String = "learned_subjects.txt"

' Settings file name
Public Const SETTINGS_FILE_NAME As String = "settings.ini"
