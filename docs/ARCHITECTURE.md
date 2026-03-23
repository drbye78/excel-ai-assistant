# Excel AI Assistant - Architecture Documentation

## Overview

The Excel AI Assistant is an Office Add-in that provides natural language interaction with Excel workbooks through AI. This document describes the system architecture, design decisions, and component interactions.

## Table of Contents

1. [System Architecture](#system-architecture)
2. [Technology Stack](#technology-stack)
3. [Directory Structure](#directory-structure)
4. [Core Components](#core-components)
5. [Data Flow](#data-flow)
6. [State Management](#state-management)
7. [Security Architecture](#security-architecture)
8. [Error Handling](#error-handling)
9. [Performance Considerations](#performance-considerations)

---

## System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Excel Application                         │
│  ┌─────────────────────────────────────────────────────────────┐│
│  │                    Office.js Runtime                        ││
│  │  ┌──────────────────────────────────────────────────────┐  ││
│  │  │                   Taskpane                            │  ││
│  │  │  ┌────────────────────────────────────────────────┐  │  ││
│  │  │  │              React Application                 │  │  ││
│  │  │  │  ┌──────────────────────────────────────────┐  │  │  ││
│  │  │  │  │            AppContext                    │  │  │  ││
│  │  │  │  │  ┌────────────────────────────────────┐  │  │  │  ││
│  │  │  │  │  │         UI Components              │  │  │  │  ││
│  │  │  │  │  │  • Chat   • Settings   • History  │  │  │  │  ││
│  │  │  │  │  └────────────────────────────────────┘  │  │  │  ││
│  │  │  │  │  ┌────────────────────────────────────┐  │  │  │  ││
│  │  │  │  │  │         Services Layer             │  │  │  │  ││
│  │  │  │  │  │  • AIService  • ExcelService       │  │  │  │  ││
│  │  │  │  │  │  • RateLimiter  • CostTracker      │  │  │  │  ││
│  │  │  │  │  └────────────────────────────────────┘  │  │  │  ││
│  │  │  │  └──────────────────────────────────────────┘  │  │  ││
│  │  │  └────────────────────────────────────────────────┘  │  ││
│  │  └──────────────────────────────────────────────────────┘  ││
│  └─────────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
                    ┌───────────────────────────┐
                    │     AI API Providers      │
                    │  • OpenRouter            │
                    │  • OpenAI                │
                    │  • Azure OpenAI          │
                    │  • Local Models          │
                    └───────────────────────────┘
```

---

## Technology Stack

### Frontend
- **React 18** - UI framework with hooks
- **Fluent UI React** - Microsoft's design system components
- **TypeScript** - Type-safe JavaScript

### Office Integration
- **Office.js** - Official Microsoft Office Add-in API
- **Office.js Helpers** - Utility functions for Office Add-ins

### Build Tools
- **Webpack 5** - Module bundler
- **Babel** - JavaScript/TypeScript transpilation
- **ts-loader** - TypeScript compilation

### Testing
- **Jest** - Testing framework
- **React Testing Library** - React component testing
- **ts-jest** - TypeScript support for Jest

### AI Integration
- **OpenAI SDK** - Official OpenAI API client
- **Axios** - HTTP client for API requests

---

## Directory Structure

```
excel-ai-assistant/
├── src/
│   ├── components/           # React UI components
│   │   ├── Chat.tsx         # Chat interface
│   │   ├── Settings.tsx     # Settings configuration
│   │   ├── ErrorBoundary.tsx # Error handling
│   │   └── ...
│   │
│   ├── services/            # Business logic services
│   │   ├── aiService.ts     # AI API integration
│   │   ├── excelService.ts  # Excel operations
│   │   ├── actionHandler.ts # Action execution
│   │   ├── rateLimiter.ts   # API rate limiting
│   │   ├── costTracker.ts   # Token cost tracking
│   │   ├── undoRedoService.ts # Undo/redo functionality
│   │   └── __tests__/       # Service tests
│   │
│   ├── context/             # React Context providers
│   │   └── AppContext.tsx   # Global state management
│   │
│   ├── types/               # TypeScript definitions
│   │   └── index.ts         # All type definitions
│   │
│   ├── utils/               # Utility functions
│   │   ├── encryption.ts    # Secure storage
│   │   ├── errors.ts        # Error classes
│   │   └── ...
│   │
│   ├── config/              # Configuration
│   │   └── constants.ts     # App constants
│   │
│   ├── i18n/                # Internationalization
│   │   ├── types.ts         # i18n types
│   │   └── translations.ts  # Language translations
│   │
│   ├── hooks/               # Custom React hooks
│   │
│   ├── taskpane/            # Taskpane entry point
│   │   ├── index.tsx        # React root
│   │   ├── App.tsx          # Main app component
│   │   └── taskpane.html    # HTML template
│   │
│   └── commands/            # Office command handlers
│       └── commands.ts      # Ribbon commands
│
├── docs/                    # Documentation
├── __mocks__/              # Jest mocks
├── server/                 # Development server
├── assets/                 # Static assets
├── manifest.xml            # Office Add-in manifest
├── package.json            # Dependencies
├── tsconfig.json           # TypeScript config
├── webpack.config.js       # Webpack config
└── jest.config.js          # Jest config
```

---

## Core Components

### 1. Services Layer

#### AIService
Primary service for AI interactions. Handles:
- API communication with multiple providers
- Conversation context management
- Tool-calling for Excel operations
- Rate limiting integration
- Cost tracking

```typescript
// Usage example
const response = await AIService.sendMessage(userMessage, context);
```

#### ExcelService
Encapsulates all Excel operations. Provides:
- Cell/range manipulation
- Table/chart creation
- Pivot table operations
- Format and style operations
- Validation and protection

```typescript
// Usage example
await ExcelService.setValues(address, values, worksheet);
await ExcelService.createTable(range, options);
```

#### ActionHandler
Validates and executes AI-generated actions:
- Action whitelisting for security
- Destructive operation confirmation
- Action execution with error handling
- Undo/redo support

```typescript
// Usage example
const result = await ActionHandler.executeAction(action);
```

#### RateLimiter
Implements token bucket algorithm:
- Request rate limiting
- Queue management for overflow
- Model-specific limits
- Tier-based configurations

#### CostTracker
Monitors API usage costs:
- Token usage tracking per model
- Budget management (daily/weekly/monthly)
- Alert notifications
- Usage reporting

### 2. State Management

#### AppContext
React Context-based state management:

```typescript
// State structure
interface AppState {
  settings: AISettings;
  messages: Message[];
  isLoading: boolean;
  excelContext: ExcelContext | null;
  error: AppError | null;
  locale: string;
}

// Hooks
const { state, actions } = useApp();
const { settings, updateSettings } = useSettings();
const { messages, addMessage } = useConversation();
```

### 3. Type System

Strongly typed action payloads:

```typescript
type AIActionType = 
  | 'insert_formula'
  | 'set_values'
  | 'create_table'
  // ... 26 action types

type ActionPayloadMap = {
  insert_formula: InsertFormulaPayload;
  set_values: SetValuesPayload;
  // ... mapped payloads
}

// Discriminated union
type TypedAIAction = {
  [K in AIActionType]: {
    type: K;
    payload: ActionPayloadMap[K];
  }
}[AIActionType];
```

---

## Data Flow

### User Query Flow

```
User Input
    │
    ▼
┌─────────────┐
│   Chat.tsx  │
└─────┬───────┘
      │
      ▼
┌─────────────────────────────────────┐
│  AppContext: addMessage()            │
│  Updates messages state              │
└─────┬───────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│  AIService.sendMessage()             │
│  1. Check rate limits                 │
│  2. Check budget                      │
│  3. Build context                     │
│  4. Send to API                       │
└─────┬───────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│  AI API (OpenRouter/OpenAI/etc)      │
│  Processes query with tools          │
└─────┬───────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│  AIService: Parse response           │
│  Extract actions from tool calls     │
└─────┬───────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│  ActionHandler.executeAction()       │
│  1. Validate action                  │
│  2. Check permissions                │
│  3. Execute in Excel                 │
│  4. Register for undo                │
└─────┬───────────────────────────────┘
      │
      ▼
┌─────────────────────────────────────┐
│  AppContext: addMessage()            │
│  Add assistant response              │
└─────────────────────────────────────┘
```

---

## Security Architecture

### API Key Storage

API keys are encrypted before storage using AES-GCM encryption:

```
┌─────────────┐    ┌──────────────┐    ┌─────────────────┐
│  API Key    │───▶│   encrypt()  │───▶│  localStorage   │
│  (plaintext)│    │  AES-GCM     │    │  (encrypted)    │
└─────────────┘    └──────────────┘    └─────────────────┘
                           │
                           ▼
                    ┌──────────────┐
                    │  Device      │
                    │  Fingerprint │
                    │  (key deriv) │
                    └──────────────┘
```

### Action Whitelisting

All actions are validated against a whitelist:

```typescript
const ALLOWED_ACTIONS = new Set([
  'insert_formula',
  'set_values',
  'create_table',
  // ... approved actions
]);

// Destructive actions require confirmation
const DESTRUCTIVE_ACTIONS = [
  'delete_table',
  'delete_worksheet',
  'clear_range',
  // ...
];
```

---

## Error Handling

### Error Class Hierarchy

```
AppError (base)
├── ValidationError
├── APIError
│   ├── APIKeyMissingError
│   ├── APIUrlMissingError
│   ├── APITimeoutError
│   └── APIRateLimitError
├── AIError
│   ├── AIModelNotAvailableError
│   ├── AIQuotaExceededError
│   └── AICostLimitError
├── ExcelAPIError
│   ├── ExcelRangeInvalidError
│   ├── ExcelSheetNotFoundError
│   └── ExcelWorkbookProtectedError
├── NetworkError
│   └── NetworkTimeoutError
├── PermissionError
├── AuthenticationError
└── StorageError
    └── StorageQuotaExceededError
```

### Error Handling Flow

```typescript
try {
  await actionHandler.executeAction(action);
} catch (error) {
  const appError = handleError(error);
  
  if (appError.retryable) {
    // Show retry option
  } else {
    // Show error message
  }
  
  logError(error, 'ActionExecution');
}
```

---

## Performance Considerations

### Rate Limiting

- Token bucket algorithm with refill rate
- Queue for overflow requests
- Model-specific TPM limits

### Cost Management

- Per-request cost estimation
- Budget alerts at 50%, 75%, 90%, 100%
- Automatic request blocking when budget exceeded

### Caching

- Excel context cached per operation
- Model list cached for 24 hours
- Settings cached in memory

### Code Splitting (Planned)

```typescript
// Lazy load heavy components
const AnalyticsDashboard = React.lazy(
  () => import('@components/AnalyticsDashboard')
);
```

---

## Testing Strategy

### Unit Tests
- Service layer tests with mocked dependencies
- Type checking tests
- Error handling tests

### Integration Tests
- Chat-to-Excel flow
- Settings persistence
- Rate limiter integration

### Test Configuration

```javascript
// jest.config.js
coverageThreshold: {
  global: { statements: 50, branches: 40 }
}
```

---

## Deployment

### Build Process

```bash
npm run build    # Production build
npm run test     # Run tests
npm run validate # Validate manifest
```

### Manifest Configuration

The `manifest.xml` defines:
- Add-in endpoints
- Ribbon buttons and commands
- Permissions required
- Supported hosts

---

## Future Enhancements

1. **Offline Support** - Queue operations when offline
2. **Sync Service** - Sync data across devices
3. **Custom Themes** - User-selectable themes
4. **Collaboration** - Multi-user features
5. **Extended AI Models** - Support for more models

---

## Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for development guidelines.