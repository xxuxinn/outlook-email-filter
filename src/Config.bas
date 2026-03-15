'===============================================================================
' Config.bas - Configuration Module for Email Agent v3.0
'===============================================================================
' This module contains:
'   1. DEFAULT_* constants: compile-time fallback values
'   2. Runtime* variables: loaded from settings.ini at startup
'   3. Version constants
'   4. Infrastructure constants (data file paths, not user-configurable)
'
' Users customize via settings.ini, NOT by editing this file.
' To reset to defaults, delete settings.ini and restart Outlook.
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' VERSION
'-------------------------------------------------------------------------------
Public Const FILTER_VERSION As String = "3.0.0"
Public Const FILTER_VERSION_DATE As String = "2026-02-19"

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

' LLM (multi-provider)
Public RuntimeUseLLM As Boolean
Public RuntimeLLMProvider As String          ' "local" | "azure" | "claude" | "openai"
Public RuntimeLLMEndpoint As String          ' Azure endpoint (legacy / provider=azure)
Public RuntimeLocalEndpoint As String        ' Local LLM endpoint (Ollama, LM Studio, Inferencer, etc.)
Public RuntimeLocalModel As String           ' Model name for local server (e.g. "qwen3:8b")
Public RuntimeClaudeEndpoint As String       ' Anthropic API endpoint
Public RuntimeClaudeModel As String          ' Claude model ID
Public RuntimeOpenAIEndpoint As String       ' OpenAI-compatible endpoint (OpenRouter, Groq, etc.)
Public RuntimeOpenAIModel As String          ' Model ID for OpenAI-compatible provider
Public RuntimeAPIKeyMethod As String
Public RuntimeAPIKeyEnvVar As String
Public RuntimeAPIKeyHardcoded As String
Public RuntimeLLMSystemPrompt As String
Public RuntimeClassifyBodyChars As Integer
Public RuntimeClassifyMaxTokens As Integer
Public RuntimeSummarizeMaxTokens As Integer
Public RuntimeReplyMaxTokens As Integer
Public RuntimeLLMTemperature As Double       ' Legacy / classify temperature
Public RuntimeReplyTemperature As Double

' Agent
Public RuntimeEnableAutoReply As Boolean
Public RuntimeAutoReplyOnArrival As Boolean
Public RuntimeFolderLearnReply As String
Public RuntimeMaxReplyExamples As Integer
Public RuntimeReplyPersona As String
Public RuntimeScanSentItems As Boolean
Public RuntimeScanSentDays As Integer
Public RuntimeAutoReplyForSenders As String

' Cloud sync
Public RuntimeEnableCloudSync As Boolean
Public RuntimeCloudSyncPath As String

' Error handling / debug
Public RuntimeDebugMode As Boolean           ' Show MsgBox on errors when True
Public RuntimeErrorLogFile As String         ' Full path to error.log (set at runtime)

' Flag: have settings been loaded yet?
Public RuntimeSettingsLoaded As Boolean

'-------------------------------------------------------------------------------
' DEFAULT CONSTANTS (fallbacks if settings.ini is missing or incomplete)
'-------------------------------------------------------------------------------

' LLM defaults
Public Const DEFAULT_USE_LLM_API As Boolean = False
Public Const DEFAULT_LLM_PROVIDER As String = "azure"
Public Const DEFAULT_AZURE_OPENAI_ENDPOINT As String = "https://YOUR-RESOURCE.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview"
Public Const DEFAULT_LOCAL_ENDPOINT As String = "http://localhost:11434/v1/chat/completions"
Public Const DEFAULT_LOCAL_MODEL As String = "qwen3:8b"
Public Const DEFAULT_CLAUDE_ENDPOINT As String = "https://api.anthropic.com/v1/messages"
Public Const DEFAULT_CLAUDE_MODEL As String = "claude-opus-4-20250115"
Public Const DEFAULT_OPENAI_COMPAT_ENDPOINT As String = "https://openrouter.ai/api/v1/chat/completions"
Public Const DEFAULT_OPENAI_COMPAT_MODEL As String = "qwen/qwen3-8b"
Public Const DEFAULT_API_KEY_METHOD As String = "ENV"
Public Const DEFAULT_API_KEY_ENV_VAR As String = "LLM_API_KEY"
Public Const DEFAULT_API_KEY_HARDCODED As String = ""
Public Const DEFAULT_LLM_SYSTEM_PROMPT As String = "You are filtering emails for Professor Xu Xin at PolyU Hong Kong. " & _
    "Respond with ONLY 'DELETE' or 'KEEP' followed by a brief reason (max 10 words). " & _
    "DELETE: Generic broadcasts, announcements, FYI-only, mass CC, admin notices, promotional, newsletters. " & _
    "KEEP: Personally addressed, requires action/response, from students/collaborators, important deadlines."
Public Const DEFAULT_CLASSIFY_BODY_CHARS As Integer = 800
Public Const DEFAULT_CLASSIFY_MAX_TOKENS As Integer = 100
Public Const DEFAULT_SUMMARIZE_MAX_TOKENS As Integer = 300
Public Const DEFAULT_REPLY_MAX_TOKENS As Integer = 800
Public Const DEFAULT_LLM_TEMPERATURE As Double = 0.3
Public Const DEFAULT_REPLY_TEMPERATURE As Double = 0.7

' Agent defaults
Public Const DEFAULT_ENABLE_AUTO_REPLY As Boolean = False
Public Const DEFAULT_AUTO_REPLY_ON_ARRIVAL As Boolean = False
Public Const DEFAULT_FOLDER_LEARN_REPLY As String = "LearnReply"
Public Const DEFAULT_MAX_REPLY_EXAMPLES As Integer = 5
Public Const DEFAULT_REPLY_PERSONA As String = ""
Public Const DEFAULT_SCAN_SENT_ITEMS As Boolean = False
Public Const DEFAULT_SCAN_SENT_DAYS As Integer = 30
Public Const DEFAULT_AUTO_REPLY_FOR_SENDERS As String = ""

' Cloud sync defaults
Public Const DEFAULT_ENABLE_CLOUD_SYNC As Boolean = False
Public Const DEFAULT_CLOUD_SYNC_PATH As String = ""

' Debug / error handling defaults
Public Const DEFAULT_DEBUG_MODE As Boolean = False

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
Public Const DEFAULT_NAME_PATTERNS As String = "Xu Xin,XuXin,Xuxin,Xin Xu,Professor Xu,Prof. Xu,Prof Xu,Dr. Xu,Dr Xu,Mr. Xu,Mr Xu"
Public Const DEFAULT_GREETING_PATTERNS As String = "Dear Professor Xu,Dear Prof. Xu,Dear Prof Xu,Dear Dr. Xu,Dear Dr Xu,Dear Xin,Hi Xin,Hello Xin,Dear Head,Dear Director"
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
Public Const LEARNED_REPLIES_FILE As String = "learned_replies.txt"

' Settings and log file names
Public Const SETTINGS_FILE_NAME As String = "settings.ini"
Public Const ERROR_LOG_FILE_NAME As String = "error.log"

' Call stack max depth
Public Const CALL_STACK_MAX_DEPTH As Integer = 20
