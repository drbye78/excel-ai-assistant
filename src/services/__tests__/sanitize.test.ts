/**
 * Unit Tests for Sanitize Utility
 * Tests input sanitization functions
 */

import {
  sanitizeCellValue,
  sanitizeVbaName,
  sanitizeQueryInput,
  sanitizeHtml,
  sanitizeFilePath,
  sanitizeEmail,
  sanitizeJson,
  isSafeContent,
  sanitizeIdentifier,
  sanitizeRangeName,
  sanitizeCellReference
} from '../../utils/sanitize';

describe('Sanitize Utility', () => {
  
  describe('sanitizeCellValue', () => {
    it('should return empty string for null/undefined', () => {
      expect(sanitizeCellValue(null)).toBe('');
      expect(sanitizeCellValue(undefined)).toBe('');
    });

    it('should return normal strings unchanged', () => {
      expect(sanitizeCellValue('hello world')).toBe('hello world');
      expect(sanitizeCellValue('12345')).toBe('12345');
    });

    it('should prefix formula-like values with single quote', () => {
      expect(sanitizeCellValue('=SUM(A1:A10)')).toBe("'=SUM(A1:A10)");
      expect(sanitizeCellValue('+A1+B1')).toBe("'+A1+B1");
      expect(sanitizeCellValue('-A1')).toBe("'-A1");
      expect(sanitizeCellValue('@SUM(A1)')).toBe("'@SUM(A1)");
    });

    it('should convert numbers to strings', () => {
      expect(sanitizeCellValue(123)).toBe('123');
      expect(sanitizeCellValue(0)).toBe('0');
    });
  });

  describe('sanitizeVbaName', () => {
    it('should remove non-letter prefix', () => {
      expect(sanitizeVbaName('123test')).toBe('test');
      expect(sanitizeVbaName('_test')).toBe('test');
    });

    it('should replace invalid characters with underscore', () => {
      expect(sanitizeVbaName('test-name')).toBe('test_name');
      expect(sanitizeVbaName('test.name')).toBe('test_name');
      expect(sanitizeVbaName('test name')).toBe('test_name');
    });

    it('should limit length to 255 characters', () => {
      const longName = 'a'.repeat(300);
      expect(sanitizeVbaName(longName).length).toBe(255);
    });

    it('should keep valid names unchanged', () => {
      expect(sanitizeVbaName('TestMacro')).toBe('TestMacro');
      expect(sanitizeVbaName('test_macro_1')).toBe('test_macro_1');
    });
  });

  describe('sanitizeQueryInput', () => {
    it('should escape single quotes', () => {
      expect(sanitizeQueryInput("it's")).toBe("it''s");
      expect(sanitizeQueryInput("O'Brien")).toBe("O''Brien");
    });

    it('should remove dangerous characters', () => {
      expect(sanitizeQueryInput('test; DROP TABLE')).toBe('test DROP TABLE');
      expect(sanitizeQueryInput('test--comment')).toBe('testcomment');
      expect(sanitizeQueryInput('test/*comment*/')).toBe('testcomment');
    });

    it('should keep safe input unchanged', () => {
      expect(sanitizeQueryInput('SELECT * FROM table')).toBe('SELECT * FROM table');
    });
  });

  describe('sanitizeHtml', () => {
    it('should encode HTML entities', () => {
      expect(sanitizeHtml('<script>alert("xss")</script>')).toBe(
        '<script>alert("xss")</script>'
      );
      expect(sanitizeHtml('Tom & Jerry')).toBe('Tom & Jerry');
      expect(sanitizeHtml("it's")).toBe("it&#x27;s");
    });

    it('should keep safe text unchanged', () => {
      expect(sanitizeHtml('Hello World')).toBe('Hello World');
    });
  });

  describe('sanitizeFilePath', () => {
    it('should remove invalid characters', () => {
      expect(sanitizeFilePath('file<name>.txt')).toBe('filename.txt');
      expect(sanitizeFilePath('file:name.txt')).toBe('filename.txt');
      expect(sanitizeFilePath('file"name".txt')).toBe('filename.txt');
    });

    it('should remove parent directory references', () => {
      expect(sanitizeFilePath('../file.txt')).toBe('/file.txt');
      expect(sanitizeFilePath('dir/../file.txt')).toBe('dir//file.txt');
    });

    it('should normalize slashes', () => {
      expect(sanitizeFilePath('dir///file.txt')).toBe('dir/file.txt');
    });

    it('should trim whitespace', () => {
      expect(sanitizeFilePath('  file.txt  ')).toBe('file.txt');
    });
  });

  describe('sanitizeEmail', () => {
    it('should validate correct emails', () => {
      const result = sanitizeEmail('user@example.com');
      expect(result.valid).toBe(true);
      expect(result.sanitized).toBe('user@example.com');
    });

    it('should lowercase and trim emails', () => {
      const result = sanitizeEmail('  USER@EXAMPLE.COM  ');
      expect(result.valid).toBe(true);
      expect(result.sanitized).toBe('user@example.com');
    });

    it('should reject invalid formats', () => {
      const result = sanitizeEmail('not-an-email');
      expect(result.valid).toBe(false);
      expect(result.error).toBeDefined();
    });

    it('should reject emails with dangerous characters', () => {
      const result = sanitizeEmail('user<script>@example.com');
      expect(result.valid).toBe(false);
    });
  });

  describe('sanitizeJson', () => {
    it('should remove control characters', () => {
      const input = 'test\x00\x1F\x7F\x9Fdata';
      expect(sanitizeJson(input)).toBe('testdata');
    });

    it('should keep valid JSON unchanged', () => {
      const json = '{"key": "value"}';
      expect(sanitizeJson(json)).toBe(json);
    });
  });

  describe('isSafeContent', () => {
    it('should detect dangerous protocols', () => {
      expect(isSafeContent('javascript:alert(1)')).toBe(false);
      expect(isSafeContent('data:text/html,<script>')).toBe(false);
      expect(isSafeContent('vbscript:msgbox(1)')).toBe(false);
    });

    it('should detect event handlers', () => {
      expect(isSafeContent('onclick=alert(1)')).toBe(false);
      expect(isSafeContent('onerror=alert(1)')).toBe(false);
    });

    it('should detect dangerous tags', () => {
      expect(isSafeContent('<script>alert(1)</script>')).toBe(false);
      expect(isSafeContent('<iframe src="evil">')).toBe(false);
      expect(isSafeContent('<object data="evil">')).toBe(false);
    });

    it('should allow safe content', () => {
      expect(isSafeContent('Hello World')).toBe(true);
      expect(isSafeContent('https://example.com')).toBe(true);
    });
  });

  describe('sanitizeIdentifier', () => {
    it('should remove invalid characters', () => {
      expect(sanitizeIdentifier('test-name')).toBe('testname');
      expect(sanitizeIdentifier('test.name')).toBe('testname');
      expect(sanitizeIdentifier('test name')).toBe('testname');
    });

    it('should prefix numeric start with col_', () => {
      expect(sanitizeIdentifier('123test')).toBe('col_123test');
    });

    it('should keep valid identifiers unchanged', () => {
      expect(sanitizeIdentifier('test_name')).toBe('test_name');
      expect(sanitizeIdentifier('TestName')).toBe('TestName');
    });
  });

  describe('sanitizeRangeName', () => {
    it('should replace invalid characters with underscore', () => {
      expect(sanitizeRangeName('test-name')).toBe('test_name');
      expect(sanitizeRangeName('test.name')).toBe('test_name');
    });

    it('should prefix numeric start with Range_', () => {
      expect(sanitizeRangeName('123test')).toBe('Range_123test');
    });

    it('should limit length to 255 characters', () => {
      const longName = 'a'.repeat(300);
      expect(sanitizeRangeName(longName).length).toBe(255);
    });
  });

  describe('sanitizeCellReference', () => {
    it('should normalize to uppercase', () => {
      expect(sanitizeCellReference('a1')).toBe('A1');
      expect(sanitizeCellReference('$a$1')).toBe('$A$1');
    });

    it('should remove invalid characters', () => {
      expect(sanitizeCellReference('A1:B10')).toBe('A1:B10');
    });

    it('should throw error for invalid references', () => {
      expect(() => sanitizeCellReference('invalid')).toThrow();
      expect(() => sanitizeCellReference('A')).toThrow();
    });

    it('should validate correct references', () => {
      expect(sanitizeCellReference('A1')).toBe('A1');
      expect(sanitizeCellReference('$A$1')).toBe('$A$1');
      expect(sanitizeCellReference('A1:B10')).toBe('A1:B10');
    });
  });
});