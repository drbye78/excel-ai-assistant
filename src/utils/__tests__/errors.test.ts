/**
 * Error Handling Tests
 *
 * Unit tests for the error handling utilities
 * @module utils/__tests__/errors
 */

import {
  handleError,
  AppError,
  ValidationError,
  ExcelAPIError,
  NetworkError,
  PermissionError,
  AIError
} from '../errors';

describe('Error Handling', () => {
  describe('AppError', () => {
    test('should create AppError with all properties', () => {
      const error = new AppError(
        'Something went wrong',
        'CUSTOM_ERROR',
        'Please try again later',
        true
      );

      expect(error).toBeInstanceOf(Error);
      expect(error).toBeInstanceOf(AppError);
      expect(error.message).toBe('Something went wrong');
      expect(error.code).toBe('CUSTOM_ERROR');
      expect(error.userMessage).toBe('Please try again later');
      expect(error.retryable).toBe(true);
      expect(error.name).toBe('AppError');
    });

    test('should create AppError without optional properties', () => {
      const error = new AppError('Basic error', 'BASIC_ERROR');

      expect(error.message).toBe('Basic error');
      expect(error.code).toBe('BASIC_ERROR');
      expect(error.userMessage).toBeUndefined();
      expect(error.retryable).toBe(false);
    });
  });

  describe('ValidationError', () => {
    test('should create ValidationError with field', () => {
      const error = new ValidationError('Invalid email', 'email');

      expect(error).toBeInstanceOf(AppError);
      expect(error).toBeInstanceOf(ValidationError);
      expect(error.message).toBe('Invalid email');
      expect(error.code).toBe('VALIDATION_ERROR');
      expect(error.userMessage).toBe('Invalid input: Invalid email');
      expect(error.field).toBe('email');
      expect(error.name).toBe('ValidationError');
    });

    test('should create ValidationError without field', () => {
      const error = new ValidationError('Invalid form');

      expect(error.message).toBe('Invalid form');
      expect(error.field).toBeUndefined();
    });
  });

  describe('ExcelAPIError', () => {
    test('should create ExcelAPIError with Excel error details', () => {
      const excelError = { code: 'GenericError', message: 'Excel failed' };
      const error = new ExcelAPIError('Failed to read cell', excelError);

      expect(error).toBeInstanceOf(AppError);
      expect(error).toBeInstanceOf(ExcelAPIError);
      expect(error.message).toBe('Failed to read cell');
      expect(error.code).toBe('EXCEL_API_ERROR');
      expect(error.userMessage).toBe('Excel operation failed. Please try again.');
      expect(error.retryable).toBe(true);
      expect(error.excelError).toBe(excelError);
      expect(error.name).toBe('ExcelAPIError');
    });

    test('should create ExcelAPIError without Excel error', () => {
      const error = new ExcelAPIError('Operation failed');

      expect(error.excelError).toBeUndefined();
    });
  });

  describe('NetworkError', () => {
    test('should create NetworkError', () => {
      const error = new NetworkError('Connection timeout');

      expect(error).toBeInstanceOf(AppError);
      expect(error).toBeInstanceOf(NetworkError);
      expect(error.message).toBe('Connection timeout');
      expect(error.code).toBe('NETWORK_ERROR');
      expect(error.userMessage).toBe('Connection failed. Check your internet connection.');
      expect(error.retryable).toBe(true);
      expect(error.name).toBe('NetworkError');
    });
  });

  describe('PermissionError', () => {
    test('should create PermissionError', () => {
      const error = new PermissionError('Access denied');

      expect(error).toBeInstanceOf(AppError);
      expect(error).toBeInstanceOf(PermissionError);
      expect(error.message).toBe('Access denied');
      expect(error.code).toBe('PERMISSION_ERROR');
      expect(error.userMessage).toBe('You don\'t have permission to perform this action.');
      expect(error.retryable).toBe(false);
      expect(error.name).toBe('PermissionError');
    });
  });

  describe('AIError', () => {
    test('should create AIError with model info', () => {
      const error = new AIError('Model unavailable', 'gpt-4');

      expect(error).toBeInstanceOf(AppError);
      expect(error).toBeInstanceOf(AIError);
      expect(error.message).toBe('Model unavailable');
      expect(error.code).toBe('AI_ERROR');
      expect(error.userMessage).toBe('AI service temporarily unavailable. Please try again.');
      expect(error.retryable).toBe(true);
      expect(error.model).toBe('gpt-4');
      expect(error.name).toBe('AIError');
    });

    test('should create AIError without model info', () => {
      const error = new AIError('Service error');

      expect(error.model).toBeUndefined();
    });
  });

  describe('handleError', () => {
    test('should preserve AppError instances', () => {
      const original = new ValidationError('test error');
      const handled = handleError(original);

      expect(handled).toBe(original);
    });

    test('should wrap standard Error', () => {
      const original = new Error('test error');
      const handled = handleError(original);

      expect(handled).toBeInstanceOf(AppError);
      expect(handled).not.toBe(original);
      expect(handled.message).toBe('test error');
      expect(handled.code).toBe('UNKNOWN_ERROR');
      expect(handled.userMessage).toBe('An unexpected error occurred.');
      expect(handled.retryable).toBe(false);
    });

    test('should wrap string error', () => {
      const handled = handleError('string error');

      expect(handled).toBeInstanceOf(AppError);
      expect(handled.message).toBe('string error');
      expect(handled.code).toBe('UNKNOWN_ERROR');
    });

    test('should wrap number error', () => {
      const handled = handleError(404);

      expect(handled).toBeInstanceOf(AppError);
      expect(handled.message).toBe('404');
    });

    test('should wrap null error', () => {
      const handled = handleError(null);

      expect(handled).toBeInstanceOf(AppError);
      expect(handled.message).toBe('null');
    });

    test('should wrap undefined error', () => {
      const handled = handleError(undefined);

      expect(handled).toBeInstanceOf(AppError);
      expect(handled.message).toBe('undefined');
    });

    test('should wrap object error', () => {
      const handled = handleError({ custom: 'error' });

      expect(handled).toBeInstanceOf(AppError);
      expect(handled.message).toBe('[object Object]');
    });

    test('should preserve original error in cause', () => {
      const original = new Error('original');
      const handled = handleError(original);

      expect(handled.cause).toBe(original);
    });
  });

  describe('Error properties', () => {
    test('all error types should have correct names', () => {
      expect(new AppError('test', 'CODE').name).toBe('AppError');
      expect(new ValidationError('test').name).toBe('ValidationError');
      expect(new ExcelAPIError('test').name).toBe('ExcelAPIError');
      expect(new NetworkError('test').name).toBe('NetworkError');
      expect(new PermissionError('test').name).toBe('PermissionError');
      expect(new AIError('test').name).toBe('AIError');
    });

    test('all error types should be instances of Error', () => {
      expect(new AppError('test', 'CODE')).toBeInstanceOf(Error);
      expect(new ValidationError('test')).toBeInstanceOf(Error);
      expect(new ExcelAPIError('test')).toBeInstanceOf(Error);
      expect(new NetworkError('test')).toBeInstanceOf(Error);
      expect(new PermissionError('test')).toBeInstanceOf(Error);
      expect(new AIError('test')).toBeInstanceOf(Error);
    });

    test('all error types should be instances of AppError', () => {
      expect(new ValidationError('test')).toBeInstanceOf(AppError);
      expect(new ExcelAPIError('test')).toBeInstanceOf(AppError);
      expect(new NetworkError('test')).toBeInstanceOf(AppError);
      expect(new PermissionError('test')).toBeInstanceOf(AppError);
      expect(new AIError('test')).toBeInstanceOf(AppError);
    });
  });

  describe('Error serialization', () => {
    test('should serialize to JSON', () => {
      const error = new ValidationError('Invalid field', 'email');
      const json = JSON.parse(JSON.stringify(error));

      expect(json.message).toBe('Invalid field');
      expect(json.code).toBe('VALIDATION_ERROR');
      expect(json.field).toBe('email');
      expect(json.name).toBe('ValidationError');
    });

    test('should have stack trace', () => {
      const error = new AppError('test', 'CODE');

      expect(error.stack).toBeDefined();
      expect(error.stack).toContain('AppError');
    });
  });
});
