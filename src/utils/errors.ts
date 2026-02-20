/**
 * Production Error Handling
 * Standardized error classes with user-friendly messages
 * Phase 5: Enhanced Error Handling
 */

import { ERROR_MESSAGES } from '@/config/constants';

// ============================================================================
// ERROR CODES
// ============================================================================

export enum ErrorCode {
  // General errors
  UNKNOWN_ERROR = 'UNKNOWN_ERROR',
  VALIDATION_ERROR = 'VALIDATION_ERROR',
  
  // API errors
  API_KEY_MISSING = 'API_KEY_MISSING',
  API_URL_MISSING = 'API_URL_MISSING',
  API_ERROR = 'API_ERROR',
  API_TIMEOUT = 'API_TIMEOUT',
  API_RATE_LIMIT = 'API_RATE_LIMIT',
  
  // AI errors
  AI_ERROR = 'AI_ERROR',
  AI_MODEL_NOT_AVAILABLE = 'AI_MODEL_NOT_AVAILABLE',
  AI_QUOTA_EXCEEDED = 'AI_QUOTA_EXCEEDED',
  AI_COST_LIMIT = 'AI_COST_LIMIT',
  
  // Excel errors
  EXCEL_API_ERROR = 'EXCEL_API_ERROR',
  EXCEL_RANGE_INVALID = 'EXCEL_RANGE_INVALID',
  EXCEL_SHEET_NOT_FOUND = 'EXCEL_SHEET_NOT_FOUND',
  EXCEL_TABLE_NOT_FOUND = 'EXCEL_TABLE_NOT_FOUND',
  EXCEL_WORKBOOK_PROTECTED = 'EXCEL_WORKBOOK_PROTECTED',
  
  // Network errors
  NETWORK_ERROR = 'NETWORK_ERROR',
  NETWORK_TIMEOUT = 'NETWORK_TIMEOUT',
  
  // Permission errors
  PERMISSION_ERROR = 'PERMISSION_ERROR',
  AUTHENTICATION_ERROR = 'AUTHENTICATION_ERROR',
  
  // Storage errors
  STORAGE_ERROR = 'STORAGE_ERROR',
  STORAGE_QUOTA_EXCEEDED = 'STORAGE_QUOTA_EXCEEDED',
  
  // Feature errors
  FEATURE_NOT_AVAILABLE = 'FEATURE_NOT_AVAILABLE',
  OPERATION_CANCELLED = 'OPERATION_CANCELLED',
}

// ============================================================================
// BASE ERROR CLASS
// ============================================================================

export class AppError extends Error {
  public readonly timestamp: Date;
  
  constructor(
    message: string,
    public readonly code: ErrorCode | string,
    public readonly userMessage?: string,
    public readonly retryable: boolean = false,
    public readonly details?: Record<string, unknown>
  ) {
    super(message);
    this.name = 'AppError';
    this.timestamp = new Date();
    
    // Maintain proper stack trace in V8 environments
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, this.constructor);
    }
  }

  /**
   * Convert error to a plain object for serialization
   */
  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      message: this.message,
      code: this.code,
      userMessage: this.userMessage,
      retryable: this.retryable,
      timestamp: this.timestamp.toISOString(),
      details: this.details,
      stack: this.stack
    };
  }
}

// ============================================================================
// VALIDATION ERRORS
// ============================================================================

export class ValidationError extends AppError {
  constructor(
    message: string,
    public readonly field?: string,
    public readonly value?: unknown
  ) {
    super(
      message,
      ErrorCode.VALIDATION_ERROR,
      `Invalid input${field ? ` for ${field}` : ''}: ${message}`,
      false,
      { field, value }
    );
    this.name = 'ValidationError';
  }
}

// ============================================================================
// API ERRORS
// ============================================================================

export class APIKeyMissingError extends AppError {
  constructor() {
    super(
      'API key is required',
      ErrorCode.API_KEY_MISSING,
      ERROR_MESSAGES.API_KEY_MISSING,
      false
    );
    this.name = 'APIKeyMissingError';
  }
}

export class APIUrlMissingError extends AppError {
  constructor() {
    super(
      'API URL is required',
      ErrorCode.API_URL_MISSING,
      ERROR_MESSAGES.API_URL_MISSING,
      false
    );
    this.name = 'APIUrlMissingError';
  }
}

export class APIError extends AppError {
  constructor(
    message: string,
    public readonly statusCode?: number,
    public readonly apiError?: unknown
  ) {
    super(
      message,
      ErrorCode.API_ERROR,
      `API error: ${message}`,
      statusCode ? statusCode >= 500 : true,
      { statusCode, apiError }
    );
    this.name = 'APIError';
  }
}

export class APITimeoutError extends AppError {
  constructor(timeout: number) {
    super(
      `Request timed out after ${timeout}ms`,
      ErrorCode.API_TIMEOUT,
      'Request timed out. Please try again.',
      true,
      { timeout }
    );
    this.name = 'APITimeoutError';
  }
}

export class APIRateLimitError extends AppError {
  constructor(
    public readonly retryAfter?: number,
    public readonly limit?: number
  ) {
    super(
      'Rate limit exceeded',
      ErrorCode.API_RATE_LIMIT,
      ERROR_MESSAGES.RATE_LIMIT_EXCEEDED,
      true,
      { retryAfter, limit }
    );
    this.name = 'APIRateLimitError';
  }
}

// ============================================================================
// AI ERRORS
// ============================================================================

export class AIError extends AppError {
  constructor(
    message: string,
    public readonly model?: string,
    public readonly provider?: string
  ) {
    super(
      message,
      ErrorCode.AI_ERROR,
      'AI service temporarily unavailable. Please try again.',
      true,
      { model, provider }
    );
    this.name = 'AIError';
  }
}

export class AIModelNotAvailableError extends AppError {
  constructor(model: string) {
    super(
      `Model '${model}' is not available`,
      ErrorCode.AI_MODEL_NOT_AVAILABLE,
      `The selected AI model '${model}' is not available. Please select a different model.`,
      false,
      { model }
    );
    this.name = 'AIModelNotAvailableError';
  }
}

export class AIQuotaExceededError extends AppError {
  constructor(
    public readonly quotaType: 'requests' | 'tokens' | 'cost'
  ) {
    super(
      `AI ${quotaType} quota exceeded`,
      ErrorCode.AI_QUOTA_EXCEEDED,
      `You have exceeded your ${quotaType} limit. Please try again later.`,
      false,
      { quotaType }
    );
    this.name = 'AIQuotaExceededError';
  }
}

export class AICostLimitError extends AppError {
  constructor(
    public readonly period: 'perRequest' | 'daily' | 'weekly' | 'monthly',
    public readonly currentCost: number,
    public readonly limit: number
  ) {
    super(
      `Cost limit exceeded for ${period}`,
      ErrorCode.AI_COST_LIMIT,
      ERROR_MESSAGES.BUDGET_EXCEEDED,
      false,
      { period, currentCost, limit }
    );
    this.name = 'AICostLimitError';
  }
}

// ============================================================================
// EXCEL ERRORS
// ============================================================================

export class ExcelAPIError extends AppError {
  constructor(
    message: string,
    public readonly operation?: string,
    public readonly excelError?: unknown
  ) {
    super(
      message,
      ErrorCode.EXCEL_API_ERROR,
      ERROR_MESSAGES.EXCEL_API_ERROR,
      true,
      { operation, excelError }
    );
    this.name = 'ExcelAPIError';
  }
}

export class ExcelRangeInvalidError extends AppError {
  constructor(range: string) {
    super(
      `Invalid range: ${range}`,
      ErrorCode.EXCEL_RANGE_INVALID,
      ERROR_MESSAGES.INVALID_RANGE,
      false,
      { range }
    );
    this.name = 'ExcelRangeInvalidError';
  }
}

export class ExcelSheetNotFoundError extends AppError {
  constructor(sheetName: string) {
    super(
      `Worksheet '${sheetName}' not found`,
      ErrorCode.EXCEL_SHEET_NOT_FOUND,
      `The worksheet '${sheetName}' does not exist.`,
      false,
      { sheetName }
    );
    this.name = 'ExcelSheetNotFoundError';
  }
}

export class ExcelTableNotFoundError extends AppError {
  constructor(tableName: string) {
    super(
      `Table '${tableName}' not found`,
      ErrorCode.EXCEL_TABLE_NOT_FOUND,
      `The table '${tableName}' does not exist.`,
      false,
      { tableName }
    );
    this.name = 'ExcelTableNotFoundError';
  }
}

export class ExcelWorkbookProtectedError extends AppError {
  constructor(operation: string) {
    super(
      `Cannot perform ${operation}: workbook is protected`,
      ErrorCode.EXCEL_WORKBOOK_PROTECTED,
      'This workbook is protected. Please unprotect it to perform this operation.',
      false,
      { operation }
    );
    this.name = 'ExcelWorkbookProtectedError';
  }
}

// ============================================================================
// NETWORK ERRORS
// ============================================================================

export class NetworkError extends AppError {
  constructor(message: string = 'Network error') {
    super(
      message,
      ErrorCode.NETWORK_ERROR,
      ERROR_MESSAGES.NETWORK_ERROR,
      true
    );
    this.name = 'NetworkError';
  }
}

export class NetworkTimeoutError extends AppError {
  constructor(timeout: number) {
    super(
      `Network request timed out after ${timeout}ms`,
      ErrorCode.NETWORK_TIMEOUT,
      'Connection timed out. Please check your internet connection.',
      true,
      { timeout }
    );
    this.name = 'NetworkTimeoutError';
  }
}

// ============================================================================
// PERMISSION ERRORS
// ============================================================================

export class PermissionError extends AppError {
  constructor(
    message: string,
    public readonly resource?: string
  ) {
    super(
      message,
      ErrorCode.PERMISSION_ERROR,
      ERROR_MESSAGES.PERMISSION_DENIED,
      false,
      { resource }
    );
    this.name = 'PermissionError';
  }
}

export class AuthenticationError extends AppError {
  constructor(message: string = 'Authentication failed') {
    super(
      message,
      ErrorCode.AUTHENTICATION_ERROR,
      'Authentication failed. Please check your credentials.',
      false
    );
    this.name = 'AuthenticationError';
  }
}

// ============================================================================
// STORAGE ERRORS
// ============================================================================

export class StorageError extends AppError {
  constructor(
    message: string,
    public readonly operation: 'read' | 'write' | 'delete'
  ) {
    super(
      message,
      ErrorCode.STORAGE_ERROR,
      'Failed to access storage. Please try again.',
      true,
      { operation }
    );
    this.name = 'StorageError';
  }
}

export class StorageQuotaExceededError extends AppError {
  constructor() {
    super(
      'Storage quota exceeded',
      ErrorCode.STORAGE_QUOTA_EXCEEDED,
      'Storage is full. Please clear some data and try again.',
      false
    );
    this.name = 'StorageQuotaExceededError';
  }
}

// ============================================================================
// FEATURE ERRORS
// ============================================================================

export class FeatureNotAvailableError extends AppError {
  constructor(feature: string) {
    super(
      `Feature '${feature}' is not available`,
      ErrorCode.FEATURE_NOT_AVAILABLE,
      `This feature is not available in your current plan or configuration.`,
      false,
      { feature }
    );
    this.name = 'FeatureNotAvailableError';
  }
}

export class OperationCancelledError extends AppError {
  constructor(operation?: string) {
    super(
      operation ? `Operation '${operation}' was cancelled` : 'Operation was cancelled',
      ErrorCode.OPERATION_CANCELLED,
      'The operation was cancelled.',
      false,
      { operation }
    );
    this.name = 'OperationCancelledError';
  }
}

// ============================================================================
// ERROR HANDLING UTILITIES
// ============================================================================

/**
 * Handle unknown errors and convert to AppError
 */
export function handleError(error: unknown): AppError {
  // Already an AppError
  if (error instanceof AppError) {
    return error;
  }
  
  // Standard Error
  if (error instanceof Error) {
    // Check for specific error types
    const message = error.message.toLowerCase();
    
    if (message.includes('network') || message.includes('fetch')) {
      return new NetworkError(error.message);
    }
    
    if (message.includes('timeout')) {
      return new NetworkTimeoutError(30000);
    }
    
    if (message.includes('quota') || message.includes('storage')) {
      return new StorageQuotaExceededError();
    }
    
    if (message.includes('permission') || message.includes('denied')) {
      return new PermissionError(error.message);
    }
    
    if (message.includes('unauthorized') || message.includes('authentication')) {
      return new AuthenticationError(error.message);
    }
    
    return new AppError(
      error.message,
      ErrorCode.UNKNOWN_ERROR,
      ERROR_MESSAGES.GENERIC_ERROR,
      false
    );
  }
  
  // String error
  if (typeof error === 'string') {
    return new AppError(
      error,
      ErrorCode.UNKNOWN_ERROR,
      ERROR_MESSAGES.GENERIC_ERROR
    );
  }
  
  // Unknown error type
  return new AppError(
    String(error),
    ErrorCode.UNKNOWN_ERROR,
    ERROR_MESSAGES.GENERIC_ERROR
  );
}

/**
 * Check if an error is retryable
 */
export function isRetryable(error: unknown): boolean {
  if (error instanceof AppError) {
    return error.retryable;
  }
  return false;
}

/**
 * Get user-friendly error message
 */
export function getUserMessage(error: unknown): string {
  if (error instanceof AppError) {
    return error.userMessage || ERROR_MESSAGES.GENERIC_ERROR;
  }
  if (error instanceof Error) {
    return error.message || ERROR_MESSAGES.GENERIC_ERROR;
  }
  return ERROR_MESSAGES.GENERIC_ERROR;
}

/**
 * Log error to console and optionally to external service
 */
export function logError(
  error: unknown,
  context?: string,
  additionalData?: Record<string, unknown>
): void {
  const appError = handleError(error);
  
  // Console logging
  console.error(`[${appError.code}] ${context || 'Error'}:`, {
    message: appError.message,
    userMessage: appError.userMessage,
    details: appError.details,
    ...additionalData
  });
  
  // Could add external error reporting here
  // e.g., Sentry, LogRocket, etc.
}

/**
 * Create an error from an HTTP response
 */
export async function createErrorFromResponse(
  response: Response,
  defaultMessage: string = 'Request failed'
): Promise<APIError> {
  let errorMessage = defaultMessage;
  let errorDetails: unknown;
  
  try {
    const body = await response.json();
    errorDetails = body;
    
    if (body.error?.message) {
      errorMessage = body.error.message;
    } else if (body.message) {
      errorMessage = body.message;
    }
  } catch {
    // Response body not JSON, use status text
    errorMessage = response.statusText || defaultMessage;
  }
  
  // Check for specific error types
  if (response.status === 401) {
    return new AuthenticationError(errorMessage);
  }
  
  if (response.status === 403) {
    return new PermissionError(errorMessage);
  }
  
  if (response.status === 429) {
    const retryAfter = response.headers.get('Retry-After');
    return new APIRateLimitError(
      retryAfter ? parseInt(retryAfter, 10) * 1000 : undefined
    );
  }
  
  return new APIError(errorMessage, response.status, errorDetails);
}
