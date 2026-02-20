/**
 * Application Configuration Constants
 * Centralized configuration to avoid hardcoded values throughout the codebase
 */

// ============================================================================
// API CONFIGURATION
// ============================================================================

export const API_DEFAULTS = {
  /** Default API URL (OpenRouter) */
  URL: 'https://openrouter.ai/api/v1',
  /** Default model */
  MODEL: 'openai/gpt-4',
  /** Default temperature */
  TEMPERATURE: 0.7,
  /** Default max tokens */
  MAX_TOKENS: 4000,
  /** Request timeout in milliseconds */
  TIMEOUT: 30000,
  /** Maximum retries for failed requests */
  MAX_RETRIES: 3,
  /** Base delay for exponential backoff (ms) */
  RETRY_DELAY_BASE: 1000,
} as const;

export const API_PROVIDERS = {
  OPENAI: {
    name: 'OpenAI',
    defaultUrl: 'https://api.openai.com/v1',
    helpUrl: 'https://platform.openai.com/api-keys',
  },
  AZURE: {
    name: 'Azure OpenAI',
    defaultUrl: '', // Must be configured by user: https://your-resource.openai.azure.com/openai/deployments/your-deployment
    helpUrl: 'https://portal.azure.com',
  },
  OPENROUTER: {
    name: 'OpenRouter',
    defaultUrl: 'https://openrouter.ai/api/v1',
    helpUrl: 'https://openrouter.ai/keys',
  },
  LOCAL: {
    name: 'Local Model',
    defaultUrl: 'http://localhost:1234/v1',
    helpUrl: 'https://lmstudio.ai',
  },
} as const;

// ============================================================================
// RATE LIMITING
// ============================================================================

export const RATE_LIMITS = {
  /** Default requests per minute */
  DEFAULT_RPM: 60,
  /** Default tokens per minute */
  DEFAULT_TPM: 60000,
  /** Maximum queue size for pending requests */
  MAX_QUEUE_SIZE: 50,
  /** Time window for rate limiting (ms) */
  WINDOW_MS: 60000,
} as const;

// ============================================================================
// COST TRACKING
// ============================================================================

export const COST_LIMITS = {
  /** Default daily budget (USD) */
  DEFAULT_DAILY: 10.00,
  /** Default weekly budget (USD) */
  DEFAULT_WEEKLY: 50.00,
  /** Default monthly budget (USD) */
  DEFAULT_MONTHLY: 200.00,
  /** Default per-request limit (USD) */
  DEFAULT_PER_REQUEST: 1.00,
  /** Alert thresholds (percentages) */
  ALERT_THRESHOLDS: [0.5, 0.75, 0.9, 1.0],
} as const;

// ============================================================================
// STORAGE KEYS
// ============================================================================

export const STORAGE_KEYS = {
  /** Secure API key storage */
  API_KEY: 'api_key_secure',
  /** User settings */
  SETTINGS: 'excel_ai_settings',
  /** Conversation history */
  CONVERSATIONS: 'excel_ai_conversations',
  /** Cost tracking data */
  COST_TRACKER: 'ai_cost_tracker_data',
  /** Rate limiter state */
  RATE_LIMITER: 'ai_rate_limiter_state',
  /** User preferences */
  PREFERENCES: 'excel_ai_preferences',
  /** Workbook-specific settings */
  WORKBOOK_SETTINGS: 'excel_ai_workbook_settings',
  /** Analytics data */
  ANALYTICS: 'excel_ai_analytics',
} as const;

// ============================================================================
// UI CONFIGURATION
// ============================================================================

export const UI_CONFIG = {
  /** Maximum messages displayed before virtualization */
  MAX_VISIBLE_MESSAGES: 100,
  /** Debounce delay for input (ms) */
  INPUT_DEBOUNCE: 300,
  /** Toast notification duration (ms) */
  TOAST_DURATION: 5000,
  /** Maximum suggestion prompts */
  MAX_SUGGESTIONS: 3,
  /** Chat message max height (px) */
  MAX_MESSAGE_HEIGHT: 400,
  /** Sidebar width (px) */
  SIDEBAR_WIDTH: 320,
  /** Animation duration (ms) */
  ANIMATION_DURATION: 200,
} as const;

// ============================================================================
// EXCEL-SPECIFIC
// ============================================================================

export const EXCEL_CONFIG = {
  /** Maximum rows for auto-analysis */
  MAX_ANALYSIS_ROWS: 10000,
  /** Maximum columns for context */
  MAX_CONTEXT_COLUMNS: 50,
  /** Default table style */
  DEFAULT_TABLE_STYLE: 'TableStyleMedium2',
  /** Maximum nested formula depth */
  MAX_FORMULA_DEPTH: 64,
  /** Maximum chart data points */
  MAX_CHART_POINTS: 32000,
  /** Default number format */
  DEFAULT_NUMBER_FORMAT: 'General',
  /** Currency format */
  CURRENCY_FORMAT: '$#,##0.00',
  /** Percentage format */
  PERCENTAGE_FORMAT: '0.00%',
  /** Date format */
  DATE_FORMAT: 'yyyy-mm-dd',
} as const;

// ============================================================================
// ERROR MESSAGES
// ============================================================================

export const ERROR_MESSAGES = {
  /** API key missing */
  API_KEY_MISSING: 'Please configure your API key in the settings first.',
  /** API URL missing */
  API_URL_MISSING: 'API URL is required. Please configure your settings.',
  /** Network error */
  NETWORK_ERROR: 'Network error. Please check your internet connection.',
  /** Rate limit exceeded */
  RATE_LIMIT_EXCEEDED: 'Rate limit exceeded. Please wait a moment and try again.',
  /** Budget exceeded */
  BUDGET_EXCEEDED: 'Daily budget limit exceeded. Please try again tomorrow.',
  /** Excel API error */
  EXCEL_API_ERROR: 'Excel operation failed. Please try again.',
  /** Invalid range */
  INVALID_RANGE: 'Invalid cell range specified.',
  /** Permission denied */
  PERMISSION_DENIED: 'Permission denied. Please check your access rights.',
  /** Generic error */
  GENERIC_ERROR: 'An error occurred. Please try again.',
} as const;

// ============================================================================
// KEYBOARD SHORTCUTS
// ============================================================================

export const KEYBOARD_SHORTCUTS = {
  SEND_MESSAGE: 'Ctrl+Enter',
  OPEN_SETTINGS: 'Ctrl+Shift+S',
  TOGGLE_THEME: 'Ctrl+Shift+L',
  SHOW_HELP: 'Ctrl+Shift+H',
  RUN_RECIPE: 'Ctrl+Shift+R',
  NEW_CONVERSATION: 'Ctrl+Shift+N',
  CLEAR_CHAT: 'Ctrl+Shift+C',
} as const;

// ============================================================================
// THEME
// ============================================================================

export const THEME = {
  LIGHT: 'light',
  DARK: 'dark',
  EXCEL: 'excel',
  SYSTEM: 'system',
} as const;

// ============================================================================
// LOCALIZATION
// ============================================================================

export const LOCALE = {
  DEFAULT: 'en',
  SUPPORTED: ['en', 'ru', 'es', 'de', 'fr'] as const,
  FALLBACK: 'en',
} as const;

// ============================================================================
// ANIMATION
// ============================================================================

export const ANIMATION = {
  DURATION_FAST: 150,
  DURATION_NORMAL: 300,
  DURATION_SLOW: 500,
  EASING_DEFAULT: 'ease-in-out',
} as const;

// ============================================================================
// VALIDATION
// ============================================================================

export const VALIDATION = {
  /** Minimum API key length */
  MIN_API_KEY_LENGTH: 10,
  /** Maximum conversation history items */
  MAX_CONVERSATION_HISTORY: 50,
  /** Maximum message length */
  MAX_MESSAGE_LENGTH: 10000,
  /** Maximum file size for imports (bytes) */
  MAX_IMPORT_SIZE: 5 * 1024 * 1024, // 5MB
  /** Valid model name pattern */
  MODEL_NAME_PATTERN: /^[a-zA-Z0-9\-_\/]+$/,
} as const;

// ============================================================================
// FEATURE FLAGS
// ============================================================================

export const FEATURES = {
  /** Enable voice commands */
  VOICE_COMMANDS: true,
  /** Enable offline mode */
  OFFLINE_MODE: false,
  /** Enable analytics */
  ANALYTICS: true,
  /** Enable auto-save */
  AUTO_SAVE: true,
  /** Enable spell check */
  SPELL_CHECK: false,
  /** Enable experimental features */
  EXPERIMENTAL: false,
} as const;