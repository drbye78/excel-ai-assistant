// Named Range Service - Manages Named Ranges in Excel
// Production Implementation with Office.js API

import { logger } from '../utils/logger';

export interface NamedRange {
  name: string;
  range: string;
  scope: 'workbook' | string; // 'workbook' or worksheet name
  refersTo: string;
  comment?: string;
}

export interface NamedRangeCreateOptions {
  name: string;
  range: string;
  scope?: 'workbook' | string;
  comment?: string;
}

export interface NamedRangeUpdateOptions {
  newName?: string;
  newRange?: string;
  newComment?: string;
}

export class NamedRangeService {
  private static instance: NamedRangeService;

  private constructor() {}

  static getInstance(): NamedRangeService {
    if (!NamedRangeService.instance) {
      NamedRangeService.instance = new NamedRangeService();
    }
    return NamedRangeService.instance;
  }

  /**
   * Create a new named range
   */
  async createNamedRange(options: NamedRangeCreateOptions): Promise<NamedRange> {
    // Validate the name first
    const validation = this.validateName(options.name);
    if (!validation.valid) {
      throw new Error(`Invalid named range name: ${validation.error}`);
    }

    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const names = workbook.names;
      
      // Build the full reference with scope if worksheet-level
      let fullReference = options.range;
      if (options.scope && options.scope !== 'workbook') {
        // For worksheet-level names, we need to prefix with the worksheet name
        fullReference = `'${options.scope}'!${options.range}`;
      }

      // Add the named range
      names.add(options.name, fullReference);
      
      // Add comment if provided (Office.js doesn't directly support comments on names,
      // but we store it in our return object)
      await context.sync();
      
      return {
        name: options.name,
        range: options.range,
        scope: options.scope || 'workbook',
        refersTo: fullReference,
        comment: options.comment
      };
    }).catch((error) => {
      logger.error('Failed to create named range', { error, options });
      throw new Error(`Failed to create named range: ${error.message}`);
    });
  }

  /**
   * Get all named ranges in the workbook
   */
  async getAllNamedRanges(): Promise<NamedRange[]> {
    return Excel.run(async (context) => {
      const names = context.workbook.names;
      names.load(['name', 'scope', 'type']);
      
      await context.sync();
      
      const namedRanges: NamedRange[] = [];
      
      for (const name of names.items) {
        try {
          const namedItem = context.workbook.names.getItem(name.name);
          namedItem.load('name');
          
          // Try to get the range reference
          try {
            const range = namedItem.getRange();
            range.load('address');
            await context.sync();
            
            namedRanges.push({
              name: name.name,
              range: range.address || '',
              scope: name.scope || 'workbook',
              refersTo: range.address || ''
            });
          } catch {
            // Some named ranges might not have a valid range (formula names, etc.)
            namedRanges.push({
              name: name.name,
              range: '',
              scope: name.scope || 'workbook',
              refersTo: ''
            });
          }
        } catch (e) {
          // Continue with other names
        }
      }
      
      return namedRanges;
    }).catch((error) => {
      logger.error('Failed to get all named ranges', { error });
      return [];
    });
  }

  /**
   * Get a specific named range by name
   */
  async getNamedRange(name: string): Promise<NamedRange | null> {
    return Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(name);
      namedItem.load('name', 'scope');
      
      await context.sync();
      
      try {
        const range = namedItem.getRange();
        range.load('address');
        await context.sync();
        
        return {
          name: name,
          range: range.address || '',
          scope: namedItem.scope || 'workbook',
          refersTo: range.address || ''
        };
      } catch {
        return {
          name: name,
          range: '',
          scope: namedItem.scope || 'workbook',
          refersTo: ''
        };
      }
    }).catch((error) => {
      logger.error('Failed to get named range', { error, name });
      return null;
    });
  }

  /**
   * Update an existing named range
   */
  async updateNamedRange(name: string, options: NamedRangeUpdateOptions): Promise<NamedRange> {
    // If renaming, validate the new name
    if (options.newName) {
      const validation = this.validateName(options.newName);
      if (!validation.valid) {
        throw new Error(`Invalid named range name: ${validation.error}`);
      }
    }

    return Excel.run(async (context) => {
      // Get the old named range to find its current scope
      const oldNamedItem = context.workbook.names.getItem(name);
      oldNamedItem.load('scope');
      await context.sync();
      
      const currentScope = oldNamedItem.scope === 'Workbook' ? 'workbook' : (oldNamedItem.scope || 'workbook');
      
      // Delete the old name
      oldNamedItem.delete();
      await context.sync();
      
      // Create new name with updated properties
      const newName = options.newName || name;
      const newRange = options.newRange || '';
      
      // Build full reference
      let fullReference = newRange;
      if (currentScope !== 'workbook') {
        fullReference = `'${currentScope}'!${newRange}`;
      }
      
      // Add the new named range
      context.workbook.names.add(newName, fullReference);
      await context.sync();
      
      return {
        name: newName,
        range: newRange,
        scope: currentScope,
        refersTo: fullReference,
        comment: options.newComment
      };
    }).catch((error) => {
      logger.error('Failed to update named range', { error, name, options });
      throw new Error(`Failed to update named range: ${error.message}`);
    });
  }

  /**
   * Delete a named range
   */
  async deleteNamedRange(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(name);
      namedItem.delete();
      await context.sync();
    }).catch((error) => {
      logger.error('Failed to delete named range', { error, name });
      throw new Error(`Failed to delete named range: ${error.message}`);
    });
  }

  /**
   * Delete all named ranges (with optional filter)
   */
  async deleteAllNamedRanges(filter?: { prefix?: string; scope?: string }): Promise<number> {
    const allRanges = await this.getAllNamedRanges();
    let deleted = 0;
    
    for (const nr of allRanges) {
      // Apply filter
      if (filter?.prefix && !nr.name.startsWith(filter.prefix)) {
        continue;
      }
      if (filter?.scope && nr.scope !== filter.scope) {
        continue;
      }
      
      try {
        await this.deleteNamedRange(nr.name);
        deleted++;
      } catch (e) {
        // Continue even if one fails
      }
    }
    
    return deleted;
  }

  /**
   * Check if a named range exists
   */
  async namedRangeExists(name: string): Promise<boolean> {
    const namedRange = await this.getNamedRange(name);
    return namedRange !== null;
  }

  /**
   * Get named ranges that reference a specific range
   */
  async getNamedRangesForRange(range: string): Promise<NamedRange[]> {
    const allRanges = await this.getAllNamedRanges();
    return allRanges.filter(nr => 
      nr.range.toUpperCase() === range.toUpperCase() ||
      nr.refersTo.toUpperCase().includes(range.toUpperCase())
    );
  }

  /**
   * Navigate to a named range (select it)
   */
  async navigateToNamedRange(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(name);
      const range = namedItem.getRange();
      range.select();
      await context.sync();
    }).catch((error) => {
      logger.error('Failed to navigate to named range', { error, name });
      throw new Error(`Failed to navigate to named range: ${error.message}`);
    });
  }

  /**
   * Create named ranges from table headers (auto-create)
   */
  async createNamedRangesFromHeaders(tableRange: string, options?: {
    prefix?: string;
    suffix?: string;
    scope?: 'workbook' | string;
  }): Promise<NamedRange[]> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load('name');
      
      const range = worksheet.getRange(tableRange);
      range.load(['rowCount', 'columnCount', 'values']);
      
      await context.sync();
      
      const headerRow = range.values[0] || [];
      const columnCount = range.columnCount;
      const namedRanges: NamedRange[] = [];
      
      // Get existing names to avoid duplicates
      const existingRanges = await this.getAllNamedRanges();
      const existingNames = new Set(existingRanges.map(nr => nr.name));
      
      for (let col = 0; col < columnCount; col++) {
        const headerText = headerRow[col];
        if (!headerText || typeof headerText !== 'string') continue;
        
        // Build the name
        let name = this.sanitizeName(headerText);
        if (options?.prefix) {
          name = `${options.prefix}${name}`;
        }
        if (options?.suffix) {
          name = `${name}${options.suffix}`;
        }
        
        // Skip if name already exists
        if (existingNames.has(name)) continue;
        
        // Build the column range (e.g., A2:A100)
        const colLetter = this.getColumnLetter(col);
        const lastRow = range.rowCount;
        const columnRange = `${colLetter}2:${colLetter}${lastRow}`;
        
        try {
          const created = await this.createNamedRange({
            name,
            range: columnRange,
            scope: options?.scope
          });
          namedRanges.push(created);
          existingNames.add(name);
        } catch (e) {
          // Continue with other columns
        }
      }
      
      return namedRanges;
    }).catch((error) => {
      logger.error('Failed to create named ranges from headers', { error, tableRange, options });
      return [];
    });
  }

  /**
   * Create named ranges from selection using top row as names
   */
  async createNamedRangesFromSelection(options?: {
    includeHeaders?: boolean;
    scope?: 'workbook' | string;
  }): Promise<NamedRange[]> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load('name');
      
      // Get the selected range
      const selection = context.workbook.getSelectedRange();
      selection.load(['rowCount', 'columnCount', 'values', 'address']);
      
      await context.sync();
      
      const headerRow = options?.includeHeaders ? selection.values[0] : null;
      const rowCount = options?.includeHeaders ? selection.rowCount - 1 : selection.rowCount;
      const columnCount = selection.columnCount;
      const namedRanges: NamedRange[] = [];
      
      // Get existing names to avoid duplicates
      const existingRanges = await this.getAllNamedRanges();
      const existingNames = new Set(existingRanges.map(nr => nr.name));
      
      // Get the starting row (2 if including headers, 1 if not)
      const startRow = options?.includeHeaders ? 2 : 1;
      const selectionAddress = selection.address;
      // Parse the selection address to get the starting column
      const rangeParts = selectionAddress.split('!')[1]?.split(':');
      if (!rangeParts) return [];
      
      const startColLetter = rangeParts[0].replace(/[\d$]/g, '');
      
      for (let col = 0; col < columnCount; col++) {
        // Determine the header name
        let name: string;
        if (headerRow && headerRow[col]) {
          name = this.sanitizeName(String(headerRow[col]));
        } else {
          name = `Column_${this.getColumnLetter(col)}`;
        }
        
        // Skip if name already exists
        if (existingNames.has(name)) continue;
        
        // Build the column range
        const colLetter = this.getColumnLetterFromBase(startColLetter, col);
        const lastRow = startRow + rowCount - 1;
        const columnRange = `${colLetter}${startRow}:${colLetter}${lastRow}`;
        
        try {
          const created = await this.createNamedRange({
            name,
            range: columnRange,
            scope: options?.scope
          });
          namedRanges.push(created);
          existingNames.add(name);
        } catch (e) {
          // Continue with other columns
        }
      }
      
      return namedRanges;
    }).catch((error) => {
      logger.error('Failed to create named ranges from selection', { error, options });
      return [];
    });
  }

  /**
   * Helper to convert column index to letter (0 = A, 25 = Z, 26 = AA, etc.)
   */
  private getColumnLetter(colIndex: number): string {
    let letter = '';
    let n = colIndex;
    while (n >= 0) {
      letter = String.fromCharCode(65 + (n % 26)) + letter;
      n = Math.floor(n / 26) - 1;
    }
    return letter;
  }

  /**
   * Helper to get column letter from a base letter plus offset
   */
  private getColumnLetterFromBase(baseLetter: string, offset: number): string {
    const baseIndex = this.columnLetterToIndex(baseLetter);
    return this.getColumnLetter(baseIndex + offset);
  }

  /**
   * Convert column letter to index
   */
  private columnLetterToIndex(letter: string): number {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index - 1;
  }

  /**
   * Validate a named range name
   */
  validateName(name: string): { valid: boolean; error?: string } {
    // Excel naming rules:
    // - Must start with letter or underscore
    // - Can contain letters, numbers, underscores, periods
    // - Cannot contain spaces
    // - Cannot be a cell reference (A1, R1C1, etc.)
    // - Max 255 characters

    if (!name || name.length === 0) {
      return { valid: false, error: 'Name cannot be empty' };
    }

    if (name.length > 255) {
      return { valid: false, error: 'Name cannot exceed 255 characters' };
    }

    if (!/^[a-zA-Z_]/.test(name)) {
      return { valid: false, error: 'Name must start with a letter or underscore' };
    }

    if (!/^[a-zA-Z0-9_\.]+$/.test(name)) {
      return { valid: false, error: 'Name can only contain letters, numbers, underscores, and periods' };
    }

    // Check for cell reference patterns
    const cellRefPattern = /^[A-Z]+\d+$/i;
    const r1c1Pattern = /^R\d+C\d+$/i;
    if (cellRefPattern.test(name) || r1c1Pattern.test(name)) {
      return { valid: false, error: 'Name cannot be a cell reference' };
    }

    // Reserved words in Excel
    const reservedWords = ['TRUE', 'FALSE', 'NULL', 'DIV', 'REF', 'NAME', 'N/A', 'NUM', 'VALUE', 'NULL'];
    if (reservedWords.includes(name.toUpperCase())) {
      return { valid: false, error: 'Name cannot be a reserved word' };
    }

    return { valid: true };
  }

  /**
   * Generate a valid name from a string (sanitize)
   */
  sanitizeName(input: string): string {
    return input
      .replace(/^[^a-zA-Z_]+/, '') // Remove leading non-letter/underscore
      .replace(/[^a-zA-Z0-9_\.]/g, '_') // Replace invalid chars with underscore
      .substring(0, 255); // Truncate to max length
  }

  /**
   * Get suggestions for named range names based on context
   */
  getNameSuggestions(context: {
    headerText?: string;
    columnLetter?: string;
    tableName?: string;
    existingNames?: string[];
  }): string[] {
    const suggestions: string[] = [];

    if (context.headerText) {
      // Sanitize header text
      const sanitized = this.sanitizeName(context.headerText);
      if (sanitized) suggestions.push(sanitized);

      // Add prefix variations
      if (context.tableName) {
        suggestions.push(`${context.tableName}_${sanitized}`);
        suggestions.push(`${this.sanitizeName(context.tableName)}${sanitized}`);
      }
    }

    if (context.columnLetter) {
      suggestions.push(`Column_${context.columnLetter}`);
      suggestions.push(`Col_${context.columnLetter}`);
    }

    // Filter out existing names
    if (context.existingNames) {
      return suggestions.filter(s => !context.existingNames!.includes(s));
    }

    return suggestions;
  }
}

export default NamedRangeService.getInstance();
