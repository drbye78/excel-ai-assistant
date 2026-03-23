// Hyperlink Service - Manages Hyperlinks in Excel
// Production Implementation with Office.js API

import { logger } from '../utils/logger';

/**
 * Type for accessing hyperlink property on Range
 * Office.js doesn't include hyperlink in standard types
 */
interface HyperlinkProperty {
  address: string;
  textToDisplay?: string;
  screenTip?: string;
}

/**
 * Helper type for Range with hyperlink access
 * Note: Using intersection type to avoid interface extension conflicts
 */
type RangeWithHyperlink = Excel.Range & {
  hyperlink?: HyperlinkProperty;
};

export interface Hyperlink {
  cellAddress: string;
  worksheetName: string;
  url: string;
  displayText?: string;
  screenTip?: string;
  type: 'url' | 'cell' | 'email' | 'file';
}

export interface HyperlinkCreateOptions {
  cellAddress: string;
  url: string;
  displayText?: string;
  screenTip?: string;
  worksheetName?: string;
}

export class HyperlinkService {
  private static instance: HyperlinkService;

  private constructor() {}

  static getInstance(): HyperlinkService {
    if (!HyperlinkService.instance) {
      HyperlinkService.instance = new HyperlinkService();
    }
    return HyperlinkService.instance;
  }

  /**
   * Add a URL hyperlink to a cell
   */
  async addHyperlink(options: HyperlinkCreateOptions): Promise<Hyperlink> {
    return Excel.run(async (context) => {
      const worksheet = options.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(options.cellAddress);
      
      // Set hyperlink using the range's hyperlink property
      const rangeWithLink = range as RangeWithHyperlink;
      rangeWithLink.hyperlink = {
        address: options.url,
        textToDisplay: options.displayText || options.url,
        screenTip: options.screenTip
      };
      
      await context.sync();
      
      return {
        cellAddress: options.cellAddress,
        worksheetName: worksheet.name,
        url: options.url,
        displayText: options.displayText || options.url,
        screenTip: options.screenTip,
        type: this.detectHyperlinkType(options.url)
      };
    }).catch((error) => {
      logger.error('Failed to add hyperlink', { error, options });
      throw new Error(`Failed to add hyperlink: ${error.message}`);
    });
  }

  /**
   * Add a link to another cell/range in the workbook
   */
  async addCellLink(
    cellAddress: string, 
    targetSheet: string, 
    targetCell: string, 
    displayText?: string,
    worksheetName?: string
  ): Promise<Hyperlink> {
    // Internal cell links use # prefix
    const url = `#'${targetSheet}'!${targetCell}`;
    return this.addHyperlink({
      cellAddress,
      url,
      displayText: displayText || `${targetSheet}!${targetCell}`,
      worksheetName
    });
  }

  /**
   * Add an email hyperlink (mailto)
   */
  async addEmailLink(
    cellAddress: string, 
    email: string, 
    subject?: string, 
    displayText?: string,
    worksheetName?: string
  ): Promise<Hyperlink> {
    let url = `mailto:${email}`;
    if (subject) {
      url += `?subject=${encodeURIComponent(subject)}`;
    }
    return this.addHyperlink({
      cellAddress,
      url,
      displayText: displayText || email,
      worksheetName
    });
  }

  /**
   * Add a file link
   */
  async addFileLink(
    cellAddress: string, 
    filePath: string, 
    displayText?: string,
    worksheetName?: string
  ): Promise<Hyperlink> {
    // Convert to file URL format
    const url = filePath.startsWith('file://') ? filePath : `file:///${filePath.replace(/\\/g, '/')}`;
    return this.addHyperlink({
      cellAddress,
      url,
      displayText: displayText || filePath.split(/[/\\]/).pop() || filePath,
      worksheetName
    });
  }

  /**
   * Get hyperlink from a cell
   */
  async getHyperlink(cellAddress: string, worksheetName?: string): Promise<Hyperlink | null> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(cellAddress) as RangeWithHyperlink;
      const hyperlink = range.hyperlink;
      
      await context.sync();
      
      if (hyperlink && hyperlink.address) {
        return {
          cellAddress,
          worksheetName: worksheet.name,
          url: hyperlink.address,
          displayText: hyperlink.textToDisplay,
          screenTip: hyperlink.screenTip,
          type: this.detectHyperlinkType(hyperlink.address)
        };
      }
      
      return null;
    }).catch((error) => {
      logger.error('Failed to get hyperlink', { error, cellAddress, worksheetName });
      return null;
    });
  }

  /**
   * Get all hyperlinks in a range
   * Configurable scan limit for performance
   */
  async getHyperlinksInRange(
    rangeAddress: string,
    worksheetName?: string,
    maxRows = 500,
    maxCols = 50
  ): Promise<Hyperlink[]> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(rangeAddress);
      range.load(['rowCount', 'columnCount', 'address']);

      await context.sync();

      const hyperlinks: Hyperlink[] = [];
      const rowCount = range.rowCount;
      const colCount = range.columnCount;

      // Scan cells for hyperlinks with configurable limits
      // Default: 500 rows x 50 cols = 25,000 cells max
      // Adjust based on performance requirements
      const scanRows = Math.min(rowCount, maxRows);
      const scanCols = Math.min(colCount, maxCols);

      for (let row = 0; row < scanRows; row++) {
        for (let col = 0; col < scanCols; col++) {
          try {
            const cell = range.getCell(row, col) as RangeWithHyperlink;
            const hyperlink = cell.hyperlink;
            cell.load('address');
            await context.sync();

            if (hyperlink && hyperlink.address) {
              hyperlinks.push({
                cellAddress: cell.address,
                worksheetName: worksheet.name,
                url: hyperlink.address,
                displayText: hyperlink.textToDisplay,
                screenTip: hyperlink.screenTip,
                type: this.detectHyperlinkType(hyperlink.address)
              });
            }
          } catch {
            // Skip cells that can't be accessed
          }
        }
      }

      return hyperlinks;
    }).catch((error) => {
      logger.error('Failed to get hyperlinks in range', undefined, error as Error);
      return [];
    });
  }

  /**
   * Get all hyperlinks in the workbook or worksheet
   */
  async getAllHyperlinks(scope: 'workbook' | 'worksheet' = 'worksheet'): Promise<Hyperlink[]> {
    return Excel.run(async (context) => {
      const hyperlinks: Hyperlink[] = [];
      
      if (scope === 'workbook') {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();
        
        for (const worksheet of worksheets.items) {
          const wsHyperlinks = await this.getWorksheetHyperlinks(worksheet.name, context);
          hyperlinks.push(...wsHyperlinks);
        }
      } else {
        const activeWs = context.workbook.worksheets.getActiveWorksheet();
        activeWs.load('name');
        await context.sync();
        const wsHyperlinks = await this.getWorksheetHyperlinks(activeWs.name, context);
        hyperlinks.push(...wsHyperlinks);
      }
      
      return hyperlinks;
    }).catch((error) => {
      logger.error('Failed to get all hyperlinks', { error, scope });
      return [];
    });
  }

  /**
   * Helper to get hyperlinks from a worksheet
   */
  private async getWorksheetHyperlinks(worksheetName: string, context: Excel.RequestContext): Promise<Hyperlink[]> {
    const hyperlinks: Hyperlink[] = [];
    const worksheet = context.workbook.worksheets.getItem(worksheetName);
    
    // Get used range
    const usedRange = worksheet.getUsedRange();
    usedRange.load(['rowCount', 'columnCount']);
    
    await context.sync();
    
    if (!usedRange.rowCount || usedRange.rowCount === 0) {
      return hyperlinks;
    }
    
    // Scan cells for hyperlinks - limited for performance
    const maxRows = Math.min(usedRange.rowCount, 50);
    const maxCols = Math.min(usedRange.columnCount, 10);
    
    for (let row = 0; row < maxRows; row++) {
      for (let col = 0; col < maxCols; col++) {
        const cellAddress = this.getCellAddress(row, col);
        try {
          const range = worksheet.getRange(cellAddress) as RangeWithHyperlink;
          const hyperlink = range.hyperlink;
          await context.sync();
          
          if (hyperlink && hyperlink.address) {
            hyperlinks.push({
              cellAddress,
              worksheetName,
              url: hyperlink.address,
              displayText: hyperlink.textToDisplay,
              screenTip: hyperlink.screenTip,
              type: this.detectHyperlinkType(hyperlink.address)
            });
          }
        } catch (e) {
          // Continue on error
        }
      }
    }
    
    return hyperlinks;
  }

  /**
   * Update an existing hyperlink
   */
  async updateHyperlink(
    cellAddress: string, 
    updates: Partial<HyperlinkCreateOptions>,
    worksheetName?: string
  ): Promise<Hyperlink> {
    // Delete and recreate is often easier than update
    await this.removeHyperlink(cellAddress, worksheetName);
    return this.addHyperlink({
      cellAddress,
      url: updates.url || '',
      displayText: updates.displayText,
      screenTip: updates.screenTip,
      worksheetName
    });
  }

  /**
   * Remove hyperlink from a cell
   */
  async removeHyperlink(cellAddress: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(cellAddress);
      
      // Set hyperlink to undefined to remove
      const rangeWithLink = range as RangeWithHyperlink;
      (rangeWithLink as { hyperlink: HyperlinkProperty | undefined }).hyperlink = undefined;
      
      await context.sync();
    }).catch((error) => {
      logger.error('Failed to remove hyperlink', { error, cellAddress, worksheetName });
      throw new Error(`Failed to remove hyperlink: ${error.message}`);
    });
  }

  /**
   * Remove all hyperlinks in a range
   */
  async removeAllHyperlinks(rangeAddress: string, worksheetName?: string): Promise<number> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(rangeAddress);
      range.load(['rowCount', 'columnCount']);
      
      await context.sync();
      
      // Clear hyperlink for entire range
      const rangeWithLink = range as RangeWithHyperlink;
      (rangeWithLink as { hyperlink: HyperlinkProperty | undefined }).hyperlink = undefined;
      
      await context.sync();
      
      return 1;
    }).catch((error) => {
      logger.error('Failed to remove hyperlinks', { error, rangeAddress, worksheetName });
      return 0;
    });
  }

  /**
   * Detect hyperlink type from URL
   */
  private detectHyperlinkType(url: string): Hyperlink['type'] {
    if (url.startsWith('mailto:')) return 'email';
    if (url.startsWith('#') || url.includes('!')) return 'cell';
    if (url.startsWith('http://') || url.startsWith('https://')) return 'url';
    if (url.startsWith('file://') || url.includes('\\') || /^[A-Za-z]:/.test(url)) return 'file';
    return 'url';
  }

  /**
   * Validate a URL - blocks dangerous protocols
   * @param url - URL to validate
   * @returns Validation result with error message if invalid
   */
  validateUrl(url: string): { valid: boolean; error?: string } {
    if (!url || url.length === 0) {
      return { valid: false, error: 'URL cannot be empty' };
    }

    // Block dangerous protocols
    const dangerousProtocols = [
      /^javascript:/i,
      /^data:/i,
      /^vbscript:/i,
      /^file:/i,  // Block file:// for security
    ];

    if (dangerousProtocols.some(pattern => pattern.test(url))) {
      return { valid: false, error: 'Protocol not allowed for security reasons' };
    }

    // Allow safe protocols
    const safeProtocols = [
      /^https?:\/\//i,       // http/https
      /^mailto:/i,            // email
      /^#.+!/i,               // internal cell reference
      /^[A-Za-z]:[/\\]/,      // Windows path (local file)
    ];
    
    const isValid = safeProtocols.some(pattern => pattern.test(url));

    if (!isValid) {
      return { valid: false, error: 'Invalid URL format. Use http://, https://, mailto:, or internal cell reference' };
    }

    // Additional validation for http/https URLs
    if (url.startsWith('http://') || url.startsWith('https://')) {
      try {
        const parsed = new URL(url);
        // Block private/internal IPs
        const hostname = parsed.hostname;
        const privateIpPatterns = [
          /^10\./,
          /^172\.(1[6-9]|2[0-9]|3[01])\./,
          /^192\.168\./,
          /^127\./,
          /^localhost$/i,
        ];
        
        if (privateIpPatterns.some(pattern => pattern.test(hostname))) {
          return { valid: false, error: 'Private/internal URLs are not allowed' };
        }
      } catch {
        return { valid: false, error: 'Invalid URL format' };
      }
    }

    return { valid: true };
  }

  /**
   * Extract URLs from text
   */
  extractUrls(text: string): string[] {
    const urlPattern = /(https?:\/\/[^\s<>"{}|\\^`[\]]+)/gi;
    const matches = text.match(urlPattern);
    return matches || [];
  }

  /**
   * Convert hyperlinks to display text only (remove links)
   */
  async convertToText(rangeAddress: string, worksheetName?: string): Promise<number> {
    // Get hyperlinks first
    const hyperlinks = await this.getHyperlinksInRange(rangeAddress, worksheetName);
    
    // Remove all hyperlinks
    const removed = await this.removeAllHyperlinks(rangeAddress, worksheetName);
    
    return hyperlinks.length;
  }

  /**
   * Create hyperlinks from URLs in cells
   */
  async createHyperlinksFromText(rangeAddress: string, worksheetName?: string): Promise<number> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(rangeAddress);
      range.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      let created = 0;
      const rowCount = Math.min(range.rowCount, 50); // Limit for performance
      const colCount = Math.min(range.columnCount, 10);
      
      for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < colCount; col++) {
          const cellValue = range.values[row][col];
          if (cellValue && typeof cellValue === 'string') {
            const urls = this.extractUrls(cellValue);
            if (urls.length > 0) {
              const cellAddress = this.getCellAddress(row, col);
              try {
                await this.addHyperlink({
                  cellAddress,
                  url: urls[0],
                  worksheetName: worksheet.name
                });
                created++;
              } catch (e) {
                // Continue on error
              }
            }
          }
        }
      }
      
      return created;
    }).catch((error) => {
      logger.error('Failed to create hyperlinks from text', { error, rangeAddress, worksheetName });
      return 0;
    });
  }

  /**
   * Get hyperlink statistics
   */
  async getHyperlinkStatistics(scope: 'workbook' | 'worksheet' = 'worksheet'): Promise<{
    total: number;
    byType: Record<string, number>;
    external: number;
    internal: number;
  }> {
    const hyperlinks = await this.getAllHyperlinks(scope);
    
    const byType: Record<string, number> = {
      url: 0,
      email: 0,
      cell: 0,
      file: 0
    };
    
    let external = 0;
    let internal = 0;
    
    for (const link of hyperlinks) {
      byType[link.type] = (byType[link.type] || 0) + 1;
      
      if (link.type === 'cell') {
        internal++;
      } else {
        external++;
      }
    }
    
    return {
      total: hyperlinks.length,
      byType,
      external,
      internal
    };
  }

  /**
   * Batch add hyperlinks
   */
  async batchAddHyperlinks(hyperlinks: HyperlinkCreateOptions[]): Promise<Hyperlink[]> {
    const results: Hyperlink[] = [];
    
    for (const link of hyperlinks) {
      try {
        const result = await this.addHyperlink(link);
        results.push(result);
      } catch (e) {
        // Continue on error
      }
    }
    
    return results;
  }

  /**
   * Helper to convert row/col to cell address
   */
  private getCellAddress(row: number, col: number): string {
    let colLetter = '';
    let n = col;
    while (n >= 0) {
      colLetter = String.fromCharCode(65 + (n % 26)) + colLetter;
      n = Math.floor(n / 26) - 1;
    }
    return `${colLetter}${row + 1}`;
  }
}

export default HyperlinkService.getInstance();
