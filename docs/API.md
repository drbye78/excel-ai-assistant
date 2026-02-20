# Excel AI Assistant - API Documentation

## Table of Contents

1. [AIService](#aiservice)
2. [ExcelService](#excelservice)
3. [ActionHandler](#actionhandler)
4. [RateLimiter](#ratelimiter)
5. [CostTracker](#costtracker)
6. [UndoRedoService](#undoredoservice)
7. [Encryption Utilities](#encryption-utilities)
8. [Error Classes](#error-classes)
9. [React Hooks](#react-hooks)
10. [Type Definitions](#type-definitions)

---

## AIService

Primary service for AI interactions.

### Methods

#### `initialize(): Promise<void>`

Initializes the AI service, loading settings and configuring the client.

```typescript
await AIService.initialize();
```

---

#### `sendMessage(message: string, context?: ExcelContext): Promise<AIResponse>`

Sends a message to the AI and returns the response.

**Parameters:**
- `message` - User's message/query
- `context` - Optional Excel context data

**Returns:**
```typescript
interface AIResponse {
  content: string;
  actions?: AIAction[];
  usage?: {
    promptTokens: number;
    completionTokens: number;
    totalTokens: number;
  };
}
```

**Example:**
```typescript
const response = await AIService.sendMessage(
  "Create a table from A1:D10",
  { selectedRange: "A1:D10", worksheetName: "Sheet1" }
);

console.log(response.content);
// Execute actions if any
if (response.actions) {
  for (const action of response.actions) {
    await ActionHandler.executeAction(action);
  }
}
```

---

#### `getSettings(): AISettings`

Returns current AI settings.

```typescript
const settings = AIService.getSettings();
console.log(settings.model, settings.temperature);
```

---

#### `updateSettings(settings: Partial<AISettings>, persist?: boolean, scope?: 'global' | 'workbook'): Promise<void>`

Updates AI settings.

**Parameters:**
- `settings` - Partial settings to update
- `persist` - Whether to persist to storage (default: true)
- `scope` - Settings scope (default: 'global')

```typescript
await AIService.updateSettings({
  model: 'gpt-4',
  temperature: 0.7
});
```

---

#### `getSettingsScope(): { type: string; source: string }`

Returns the current settings scope.

```typescript
const scope = AIService.getSettingsScope();
console.log(scope.type, scope.source);
```

---

#### `getAvailableModels(): Promise<ModelInfo[]>`

Returns list of available models.

```typescript
const models = await AIService.getAvailableModels();
models.forEach(model => {
  console.log(model.id, model.name, model.pricing);
});
```

---

#### `setApiKey(apiKey: string, persist?: boolean): Promise<void>`

Sets the API key securely.

```typescript
await AIService.setApiKey('sk-xxx...', true);
```

---

## ExcelService

Encapsulates all Excel operations.

### Methods

#### `getSelectedRange(): Promise<string>`

Returns the currently selected range address.

```typescript
const range = await ExcelService.getSelectedRange();
console.log(range); // "Sheet1!A1:B5"
```

---

#### `getRangeValues(address: string, worksheetName?: string): Promise<any[][]>`

Gets values from a range.

```typescript
const values = await ExcelService.getRangeValues("A1:B2");
console.log(values); // [["value1", "value2"], ["value3", "value4"]]
```

---

#### `setValues(address: string, values: any[][], worksheetName?: string): Promise<void>`

Sets values in a range.

```typescript
await ExcelService.setValues("A1:B2", [
  ["Header1", "Header2"],
  ["Value1", "Value2"]
]);
```

---

#### `insertFormula(address: string, formula: string, worksheetName?: string): Promise<void>`

Inserts a formula in a cell.

```typescript
await ExcelService.insertFormula("C1", "=SUM(A:A)");
```

---

#### `createTable(range: string, options?: TableOptions): Promise<Excel.Table>`

Creates a table from a range.

**Parameters:**
- `range` - Range address
- `options` - Table options

```typescript
interface TableOptions {
  name?: string;
  hasHeaders?: boolean;
  style?: string;
  worksheetName?: string;
}
```

**Example:**
```typescript
const table = await ExcelService.createTable("A1:D10", {
  name: "SalesData",
  hasHeaders: true,
  style: "TableStyleMedium2"
});
```

---

#### `deleteTable(name: string, keepData?: boolean, worksheetName?: string): Promise<void>`

Deletes a table.

```typescript
await ExcelService.deleteTable("SalesData", true);
```

---

#### `createChart(type: ChartType, dataRange: string, options?: ChartOptions): Promise<Excel.Chart>`

Creates a chart.

```typescript
type ChartType = 
  | 'ColumnClustered'
  | 'ColumnStacked'
  | 'Line'
  | 'Pie'
  | 'BarClustered'
  | 'Area'
  // ... more types

interface ChartOptions {
  name?: string;
  title?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
  worksheetName?: string;
}
```

**Example:**
```typescript
const chart = await ExcelService.createChart('ColumnClustered', 'A1:B10', {
  title: 'Sales by Month',
  xAxisTitle: 'Month',
  yAxisTitle: 'Sales'
});
```

---

#### `formatCells(address: string, format: CellFormat, worksheetName?: string): Promise<void>`

Formats cells in a range.

```typescript
interface CellFormat {
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
  };
  fill?: {
    color?: string;
    pattern?: string;
  };
  borders?: {
    style?: string;
    color?: string;
  };
  numberFormat?: string;
  horizontalAlignment?: string;
  verticalAlignment?: string;
}
```

**Example:**
```typescript
await ExcelService.formatCells("A1:A10", {
  font: { bold: true, color: "#FFFFFF" },
  fill: { color: "#4472C4" },
  horizontalAlignment: "center"
});
```

---

#### `addWorksheet(name?: string): Promise<Excel.Worksheet>`

Adds a new worksheet.

```typescript
const sheet = await ExcelService.addWorksheet("Analysis");
```

---

#### `deleteWorksheet(name: string): Promise<void>`

Deletes a worksheet.

```typescript
await ExcelService.deleteWorksheet("OldSheet");
```

---

#### `createNamedRange(name: string, address: string, worksheetName?: string): Promise<void>`

Creates a named range.

```typescript
await ExcelService.createNamedRange("SalesTotal", "Sheet1!$D$10");
```

---

#### `addValidation(address: string, validation: ValidationOptions, worksheetName?: string): Promise<void>`

Adds data validation to a range.

```typescript
interface ValidationOptions {
  type: 'list' | 'wholeNumber' | 'decimal' | 'textLength' | 'custom';
  operator?: string;
  formula1?: string;
  formula2?: string;
  showErrorMessage?: boolean;
  errorTitle?: string;
  error?: string;
}

await ExcelService.addValidation("B2:B100", {
  type: 'list',
  formula1: '"Option1,Option2,Option3"',
  showErrorMessage: true,
  error: 'Please select a valid option'
});
```

---

#### `addComment(address: string, comment: string, worksheetName?: string): Promise<void>`

Adds a comment to a cell.

```typescript
await ExcelService.addComment("A1", "This is an important value");
```

---

#### `addHyperlink(address: string, link: HyperlinkOptions, worksheetName?: string): Promise<void>`

Adds a hyperlink to a cell.

```typescript
interface HyperlinkOptions {
  address: string;
  textToDisplay?: string;
  screenTip?: string;
}

await ExcelService.addHyperlink("A1", {
  address: "https://example.com",
  textToDisplay: "Click here"
});
```

---

#### `clearRange(address: string, worksheetName?: string): Promise<void>`

Clears a range (values and formatting).

```typescript
await ExcelService.clearRange("A1:Z100");
```

---

#### `getWorkbookContext(): Promise<ExcelContext>`

Gets the current workbook context.

```typescript
interface ExcelContext {
  workbookName?: string;
  worksheetName?: string;
  selectedRange?: string;
  usedRange?: string;
  tables?: string[];
  charts?: string[];
  namedRanges?: string[];
  hasData?: boolean;
}

const context = await ExcelService.getWorkbookContext();
```

---

## ActionHandler

Validates and executes AI-generated actions.

### Methods

#### `executeAction(action: AIAction): Promise<ActionResult>`

Executes an action.

**Parameters:**
```typescript
interface AIAction {
  type: AIActionType;
  label: string;
  payload: ActionPayload;
  confidence?: number;
}
```

**Returns:**
```typescript
interface ActionResult {
  success: boolean;
  message?: string;
  error?: Error;
  data?: any;
}
```

**Example:**
```typescript
const result = await ActionHandler.executeAction({
  type: 'insert_formula',
  label: 'Insert SUM formula',
  payload: {
    address: 'D10',
    formula: '=SUM(D1:D9)',
    worksheetName: 'Sheet1'
  }
});

if (result.success) {
  console.log('Formula inserted successfully');
}
```

---

#### `validateAction(action: AIAction): boolean`

Validates an action without executing it.

```typescript
const isValid = ActionHandler.validateAction(action);
if (!isValid) {
  console.log('Action not allowed');
}
```

---

#### `isDestructive(action: AIAction): boolean`

Checks if an action is destructive.

```typescript
if (ActionHandler.isDestructive(action)) {
  // Ask user for confirmation
}
```

---

#### `getAllowedActions(): string[]`

Returns list of allowed action types.

```typescript
const allowed = ActionHandler.getAllowedActions();
console.log(allowed); // ['insert_formula', 'set_values', ...]
```

---

## RateLimiter

Implements rate limiting for API calls.

### Methods

#### `getInstance(): RateLimiter`

Returns the singleton instance.

```typescript
const limiter = RateLimiter.getInstance();
```

---

#### `configure(config: Partial<RateLimitConfig>): void`

Configures the rate limiter.

```typescript
interface RateLimitConfig {
  maxRequests: number;
  windowMs: number;
  maxBurst?: number;
  queueEnabled?: boolean;
  maxQueueSize?: number;
}

limiter.configure({
  maxRequests: 100,
  windowMs: 60000,
  maxBurst: 20,
  queueEnabled: true
});
```

---

#### `setTier(tier: RateLimitTier): void`

Sets rate limit tier.

```typescript
type RateLimitTier = 'free' | 'basic' | 'pro' | 'enterprise';

limiter.setTier('pro');
```

---

#### `tryAcquire(): boolean`

Attempts to acquire a token. Returns true if successful.

```typescript
if (limiter.tryAcquire()) {
  // Make API call
} else {
  // Wait or queue
}
```

---

#### `acquire(): Promise<void>`

Acquires a token, waiting if necessary (when queue enabled).

```typescript
try {
  await limiter.acquire();
  // Make API call
} catch (error) {
  if (error instanceof RateLimitError) {
    console.log(`Rate limited. Retry after ${error.retryAfter}ms`);
  }
}
```

---

#### `trackModelUsage(model: string, tokens: number): void`

Tracks token usage for a specific model.

```typescript
limiter.trackModelUsage('gpt-4', 1500);
```

---

#### `hasModelCapacity(model: string, estimatedTokens: number): boolean`

Checks if there's capacity for a model request.

```typescript
if (limiter.hasModelCapacity('gpt-4', 2000)) {
  // Safe to make request
}
```

---

#### `getState(): RateLimitState`

Returns current state.

```typescript
interface RateLimitState {
  tokens: number;
  lastRefill: Date;
  queueLength: number;
  totalRequests: number;
  totalThrottled: number;
}

const state = limiter.getState();
```

---

#### `reset(): void`

Resets the rate limiter state.

```typescript
limiter.reset();
```

---

## CostTracker

Monitors API usage costs.

### Methods

#### `getInstance(): CostTracker`

Returns the singleton instance.

```typescript
const tracker = CostTracker.getInstance();
```

---

#### `trackUsage(model: string, usage: TokenUsage): CostEntry`

Tracks token usage.

```typescript
interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
}

const entry = tracker.trackUsage('gpt-4', {
  promptTokens: 1000,
  completionTokens: 500,
  totalTokens: 1500
});

console.log(`Cost: $${entry.cost.toFixed(6)}`);
```

---

#### `setBudget(budget: BudgetConfig): void`

Sets budget limits.

```typescript
interface BudgetConfig {
  daily?: number;
  weekly?: number;
  monthly?: number;
  alertThresholds?: number[]; // e.g., [0.5, 0.75, 0.9, 1.0]
}

tracker.setBudget({
  daily: 5,
  monthly: 100,
  alertThresholds: [0.5, 0.75, 0.9]
});
```

---

#### `getUsage(period: 'day' | 'week' | 'month'): UsageReport`

Gets usage report for a period.

```typescript
interface UsageReport {
  totalTokens: number;
  totalCost: number;
  requestCount: number;
  byModel: Record<string, { tokens: number; cost: number; requests: number }>;
  budgetRemaining?: number;
  budgetPercentUsed?: number;
}

const report = tracker.getUsage('day');
console.log(`Today: $${report.totalCost.toFixed(4)} spent`);
```

---

#### `isWithinBudget(): boolean`

Checks if within budget.

```typescript
if (!tracker.isWithinBudget()) {
  console.log('Budget exceeded!');
}
```

---

#### `getRemainingBudget(): number`

Returns remaining budget.

```typescript
const remaining = tracker.getRemainingBudget();
```

---

#### `setAlertCallback(callback: AlertCallback): void`

Sets callback for budget alerts.

```typescript
tracker.setAlertCallback((alert) => {
  console.log(`Alert: ${alert.message}`);
  // Show notification to user
});
```

---

#### `reset(): void`

Resets all tracking data.

```typescript
tracker.reset();
```

---

## UndoRedoService

Provides undo/redo functionality for actions.

### Methods

#### `getInstance(): UndoRedoService`

Returns the singleton instance.

```typescript
const undoRedo = UndoRedoService.getInstance();
```

---

#### `registerAction(action: AIAction, description: string, previousState?: object): UndoableOperation | null`

Registers an action for potential undo.

```typescript
const operation = undoRedo.registerAction(
  { type: 'set_values', label: 'Set values', payload: { ... } },
  'Set values in range A1:B10',
  { previousValues: [['old', 'values']] }
);
```

---

#### `peekUndo(): UndoableOperation | null`

Gets the next undo operation without removing it.

```typescript
const nextUndo = undoRedo.peekUndo();
if (nextUndo) {
  console.log(`Can undo: ${nextUndo.description}`);
}
```

---

#### `peekRedo(): UndoableOperation | null`

Gets the next redo operation without removing it.

```typescript
const nextRedo = undoRedo.peekRedo();
```

---

#### `popUndo(): UndoableOperation | null`

Gets and removes the next undo operation.

```typescript
const operation = undoRedo.popUndo();
if (operation) {
  await ActionHandler.executeAction(operation.reverseAction);
}
```

---

#### `popRedo(): UndoableOperation | null`

Gets and removes the next redo operation.

```typescript
const operation = undoRedo.popRedo();
if (operation) {
  await ActionHandler.executeAction(operation.action);
}
```

---

#### `getState(): UndoRedoState`

Returns current state.

```typescript
interface UndoRedoState {
  canUndo: boolean;
  canRedo: boolean;
  undoCount: number;
  redoCount: number;
  historySize: number;
}

const state = undoRedo.getState();
```

---

#### `clearHistory(): void`

Clears all history.

```typescript
undoRedo.clearHistory();
```

---

#### `isReversible(actionType: string): boolean`

Checks if an action type is reversible.

```typescript
if (undoRedo.isReversible('delete_worksheet')) {
  // false - cannot undo worksheet deletion
}
```

---

## Encryption Utilities

Secure storage utilities using AES-GCM encryption.

### Methods

#### `encrypt(data: string): Promise<string>`

Encrypts data.

```typescript
import { encrypt } from '@/utils/encryption';

const encrypted = await encrypt('my-secret-api-key');
```

---

#### `decrypt(encryptedData: string): Promise<string>`

Decrypts data.

```typescript
import { decrypt } from '@/utils/encryption';

const decrypted = await decrypt(encrypted);
```

---

#### `secureStorage.store(key: string, value: string): Promise<void>`

Securely stores a value.

```typescript
import { secureStorage } from '@/utils/encryption';

await secureStorage.store('API_KEY', 'sk-xxx...');
```

---

#### `secureStorage.retrieve(key: string): Promise<string | null>`

Retrieves a securely stored value.

```typescript
const apiKey = await secureStorage.retrieve('API_KEY');
```

---

#### `secureStorage.remove(key: string): void`

Removes a securely stored value.

```typescript
secureStorage.remove('API_KEY');
```

---

#### `getDeviceFingerprint(): Promise<string>`

Gets a device fingerprint for key derivation.

```typescript
import { getDeviceFingerprint } from '@/utils/encryption';

const fingerprint = await getDeviceFingerprint();
```

---

## Error Classes

### Base Error

```typescript
class AppError extends Error {
  code: string;
  retryable: boolean;
  details?: Record<string, unknown>;
  
  constructor(message: string, code: string, retryable?: boolean);
}
```

### API Errors

```typescript
// API key missing
throw new APIKeyMissingError();

// API URL missing
throw new APIUrlMissingError();

// API timeout
throw new APITimeoutError(30000);

// Rate limit exceeded
throw new APIRateLimitError(5000); // retryAfter in ms
```

### AI Errors

```typescript
// Model not available
throw new AIModelNotAvailableError('gpt-5');

// Quota exceeded
throw new AIQuotaExceededError();

// Cost limit reached
throw new AICostLimitError(10.00, 'daily');
```

### Excel Errors

```typescript
// Invalid range
throw new ExcelRangeInvalidError('ZZZ999');

// Sheet not found
throw new ExcelSheetNotFoundError('MissingSheet');

// Workbook protected
throw new ExcelWorkbookProtectedError();
```

### Error Handling Utility

```typescript
import { handleError, logError } from '@/utils/errors';

try {
  // ... operation
} catch (error) {
  const appError = handleError(error);
  
  console.error(appError.message);
  console.error(appError.code);
  
  if (appError.retryable) {
    // Show retry button
  }
  
  logError(error, 'ContextName');
}
```

---

## React Hooks

### useApp

Main hook for app state.

```typescript
import { useApp } from '@/context/AppContext';

function MyComponent() {
  const { state, dispatch, actions } = useApp();
  
  return (
    <div>
      <p>Active tab: {state.activeTab}</p>
      <button onClick={() => actions.setActiveTab('settings')}>
        Settings
      </button>
    </div>
  );
}
```

---

### useSettings

Hook for settings state.

```typescript
import { useSettings } from '@/context/AppContext';

function SettingsComponent() {
  const { settings, settingsScope, hasSettings, updateSettings } = useSettings();
  
  const handleSave = async (newSettings: AISettings) => {
    await updateSettings(newSettings);
  };
  
  return (
    <div>
      <p>Current model: {settings.model}</p>
      <p>Settings source: {settingsScope.source}</p>
    </div>
  );
}
```

---

### useConversation

Hook for conversation state.

```typescript
import { useConversation } from '@/context/AppContext';

function ChatComponent() {
  const { 
    messages, 
    isLoading, 
    suggestedPrompts,
    addMessage, 
    clearMessages,
    setLoading 
  } = useConversation();
  
  return (
    <div>
      {messages.map(msg => (
        <div key={msg.id}>{msg.content}</div>
      ))}
      {isLoading && <p>Loading...</p>}
    </div>
  );
}
```

---

### useError

Hook for error state.

```typescript
import { useError } from '@/context/AppContext';

function ErrorDisplay() {
  const { error, setError } = useError();
  
  if (!error) return null;
  
  return (
    <div className="error">
      <p>{error.message}</p>
      <button onClick={() => setError(null)}>Dismiss</button>
    </div>
  );
}
```

---

## Type Definitions

### AI Settings

```typescript
interface AISettings {
  apiUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxTokens: number;
  systemPrompt?: string;
  provider?: 'openai' | 'openrouter' | 'azure' | 'local';
}
```

### Message

```typescript
interface Message {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: Date;
  actions?: AIAction[];
  error?: boolean;
}
```

### AI Action Types

```typescript
type AIActionType =
  | 'insert_formula'
  | 'set_values'
  | 'clear_range'
  | 'format_cells'
  | 'create_table'
  | 'delete_table'
  | 'create_chart'
  | 'delete_chart'
  | 'add_worksheet'
  | 'delete_worksheet'
  | 'create_named_range'
  | 'delete_named_range'
  | 'add_validation'
  | 'add_comment'
  | 'delete_comment'
  | 'add_hyperlink'
  | 'remove_hyperlink'
  | 'create_pivot_table'
  | 'apply_conditional_formatting'
  | 'sort_range'
  | 'filter_range'
  | 'explain';
```

### Excel Context

```typescript
interface ExcelContext {
  workbookName?: string;
  worksheetName?: string;
  selectedRange?: string;
  usedRange?: string;
  tables?: string[];
  charts?: string[];
  namedRanges?: string[];
  hasData?: boolean;
  columnCount?: number;
  rowCount?: number;
  dataPreview?: any[][];
}
```

---

## Constants

```typescript
// API Defaults
API_DEFAULTS = {
  URL: 'https://openrouter.ai/api/v1',
  MODEL: 'openai/gpt-4o-mini',
  TEMPERATURE: 0.7,
  MAX_TOKENS: 4096
}

// Rate Limits
RATE_LIMITS = {
  TIERS: {
    free: { maxRequests: 20, windowMs: 60000 },
    basic: { maxRequests: 60, windowMs: 60000 },
    pro: { maxRequests: 300, windowMs: 60000 },
    enterprise: { maxRequests: 1000, windowMs: 60000 }
  }
}

// Cost Limits
COST_LIMITS = {
  DEFAULT_DAILY_BUDGET: 10,
  DEFAULT_MONTHLY_BUDGET: 100,
  MODEL_PRICING: {
    'openai/gpt-4o': { input: 0.005, output: 0.015 },
    'openai/gpt-4o-mini': { input: 0.00015, output: 0.0006 },
    // ...
  }
}

// Storage Keys
STORAGE_KEYS = {
  API_KEY: 'excel_ai_api_key',
  SETTINGS: 'excel_ai_settings',
  CONVERSATIONS: 'excel_ai_conversations'
}
```

---

## Error Codes

| Code | Description |
|------|-------------|
| `API_KEY_MISSING` | API key not configured |
| `API_URL_MISSING` | API URL not configured |
| `API_TIMEOUT` | API request timed out |
| `API_RATE_LIMIT` | Rate limit exceeded |
| `AI_MODEL_NOT_AVAILABLE` | Requested model not available |
| `AI_QUOTA_EXCEEDED` | Provider quota exceeded |
| `AI_COST_LIMIT` | Cost budget exceeded |
| `EXCEL_RANGE_INVALID` | Invalid cell range |
| `EXCEL_SHEET_NOT_FOUND` | Worksheet not found |
| `EXCEL_WORKBOOK_PROTECTED` | Workbook is protected |
| `VALIDATION_FAILED` | Input validation failed |
| `NETWORK_ERROR` | Network connectivity issue |
| `STORAGE_ERROR` | Storage operation failed |
| `PERMISSION_DENIED` | Permission not granted |