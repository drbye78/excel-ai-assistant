/**
 * Unit Tests for HyperlinkService
 * Tests URL validation, security, and hyperlink operations
 */

import { HyperlinkService } from '../hyperlinkService';

// Mock logger
jest.mock('../../utils/logger', () => ({
  logger: {
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn()
  }
}));

// Mock Excel.run
const mockContext = {
  workbook: {
    worksheets: {
      getActiveWorksheet: jest.fn(),
      getItem: jest.fn()
    }
  },
  sync: jest.fn()
};

const mockWorksheet = {
  getRange: jest.fn(),
  name: 'Sheet1'
};

const mockRange = {
  address: 'A1',
  hyperlink: undefined as any,
  load: jest.fn(),
  getCell: jest.fn()
};

beforeEach(() => {
  jest.clearAllMocks();
  (global as any).Excel = {
    run: jest.fn((callback) => callback(mockContext))
  };
  
  mockContext.workbook.worksheets.getActiveWorksheet.mockReturnValue(mockWorksheet);
  mockWorksheet.getRange.mockReturnValue(mockRange);
  mockRange.load.mockResolvedValue(undefined);
  mockContext.sync.mockResolvedValue(undefined);
});

describe('HyperlinkService', () => {
  let service: HyperlinkService;

  beforeEach(() => {
    service = HyperlinkService.getInstance();
  });

  describe('singleton', () => {
    it('should return the same instance', () => {
      const instance1 = HyperlinkService.getInstance();
      const instance2 = HyperlinkService.getInstance();
      expect(instance1).toBe(instance2);
    });
  });

  describe('validateUrl', () => {
    it('should reject empty URLs', () => {
      const result = service.validateUrl('');
      expect(result.valid).toBe(false);
      expect(result.error).toBe('URL cannot be empty');
    });

    it('should reject dangerous javascript: protocol', () => {
      const result = service.validateUrl('javascript:alert(1)');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Protocol not allowed');
    });

    it('should reject dangerous data: protocol', () => {
      const result = service.validateUrl('data:text/html,<script>alert(1)</script>');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Protocol not allowed');
    });

    it('should reject dangerous vbscript: protocol', () => {
      const result = service.validateUrl('vbscript:msgbox(1)');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Protocol not allowed');
    });

    it('should reject file: protocol for security', () => {
      const result = service.validateUrl('file:///C:/test.txt');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Protocol not allowed');
    });

    it('should accept valid http URLs', () => {
      const result = service.validateUrl('http://example.com');
      expect(result.valid).toBe(true);
    });

    it('should accept valid https URLs', () => {
      const result = service.validateUrl('https://example.com/path?query=1');
      expect(result.valid).toBe(true);
    });

    it('should accept valid mailto: URLs', () => {
      const result = service.validateUrl('mailto:user@example.com');
      expect(result.valid).toBe(true);
    });

    it('should accept internal cell references', () => {
      const result = service.validateUrl("#'Sheet2'!A1");
      expect(result.valid).toBe(true);
    });

    it('should accept Windows paths', () => {
      const result = service.validateUrl('C:\\Users\\test\\file.xlsx');
      expect(result.valid).toBe(true);
    });

    it('should reject private IP 10.x.x.x', () => {
      const result = service.validateUrl('http://10.0.0.1/api');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Private/internal URLs');
    });

    it('should reject private IP 192.168.x.x', () => {
      const result = service.validateUrl('http://192.168.1.1/api');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Private/internal URLs');
    });

    it('should reject private IP 172.16-31.x.x', () => {
      const result = service.validateUrl('http://172.16.0.1/api');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Private/internal URLs');
    });

    it('should reject localhost', () => {
      const result = service.validateUrl('http://localhost:3000/api');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Private/internal URLs');
    });

    it('should reject 127.x.x.x', () => {
      const result = service.validateUrl('http://127.0.0.1/api');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Private/internal URLs');
    });

    it('should reject invalid URL formats', () => {
      const result = service.validateUrl('not-a-url');
      expect(result.valid).toBe(false);
      expect(result.error).toContain('Invalid URL format');
    });
  });

  describe('extractUrls', () => {
    it('should extract HTTP URLs from text', () => {
      const text = 'Visit http://example.com for more info';
      const urls = service.extractUrls(text);
      expect(urls).toContain('http://example.com');
    });

    it('should extract HTTPS URLs from text', () => {
      const text = 'Visit https://example.com/path for more info';
      const urls = service.extractUrls(text);
      expect(urls).toContain('https://example.com/path');
    });

    it('should extract multiple URLs', () => {
      const text = 'Visit http://example.com or https://test.com';
      const urls = service.extractUrls(text);
      expect(urls.length).toBe(2);
    });

    it('should return empty array for text without URLs', () => {
      const text = 'No URLs here';
      const urls = service.extractUrls(text);
      expect(urls).toEqual([]);
    });
  });

  describe('addHyperlink', () => {
    it('should add a URL hyperlink to a cell', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addHyperlink({
        cellAddress: 'A1',
        url: 'https://example.com',
        displayText: 'Click here'
      });

      expect(result).toBeDefined();
      expect(result.cellAddress).toBe('A1');
      expect(result.url).toBe('https://example.com');
      expect(result.type).toBe('url');
    });

    it('should detect email hyperlinks', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addHyperlink({
        cellAddress: 'A1',
        url: 'mailto:test@example.com'
      });

      expect(result.type).toBe('email');
    });

    it('should detect cell reference hyperlinks', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addHyperlink({
        cellAddress: 'A1',
        url: "#'Sheet2'!B5"
      });

      expect(result.type).toBe('cell');
    });
  });

  describe('addCellLink', () => {
    it('should create an internal cell link', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addCellLink('A1', 'Sheet2', 'B5');

      expect(result.url).toContain('Sheet2');
      expect(result.url).toContain('B5');
      expect(result.type).toBe('cell');
    });
  });

  describe('addEmailLink', () => {
    it('should create a mailto link', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addEmailLink('A1', 'test@example.com');

      expect(result.url).toBe('mailto:test@example.com');
      expect(result.type).toBe('email');
    });

    it('should create a mailto link with subject', async () => {
      mockRange.hyperlink = undefined;
      
      const result = await service.addEmailLink('A1', 'test@example.com', 'Hello');

      expect(result.url).toContain('mailto:test@example.com');
      expect(result.url).toContain('subject=');
      expect(result.type).toBe('email');
    });
  });

  describe('getHyperlink', () => {
    it('should return hyperlink data', async () => {
      mockRange.hyperlink = {
        address: 'https://example.com',
        textToDisplay: 'Example',
        screenTip: 'Click to visit'
      };
      
      const result = await service.getHyperlink('A1');

      expect(result).toBeDefined();
      expect(result?.url).toBe('https://example.com');
      expect(result?.displayText).toBe('Example');
    });

    it('should return null for cell without hyperlink', async () => {
      mockRange.hyperlink = null;
      
      const result = await service.getHyperlink('A1');

      expect(result).toBeNull();
    });
  });

  describe('removeHyperlink', () => {
    it('should remove hyperlink from cell', async () => {
      mockRange.hyperlink = { address: 'https://example.com' };
      
      await service.removeHyperlink('A1');

      expect(mockRange.hyperlink).toBeUndefined();
    });
  });

  describe('error handling', () => {
    it('should handle Excel.run errors gracefully', async () => {
      (global as any).Excel.run.mockRejectedValue(new Error('Excel error'));
      
      const result = await service.getHyperlink('A1');
      expect(result).toBeNull();
    });

    it('should throw on addHyperlink errors', async () => {
      (global as any).Excel.run.mockRejectedValue(new Error('Excel error'));
      
      await expect(service.addHyperlink({
        cellAddress: 'A1',
        url: 'https://example.com'
      })).rejects.toThrow();
    });
  });
});