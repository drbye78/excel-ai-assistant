// Comment Service - Manages Cell Comments in Excel
// Production Implementation with Office.js API

import { logger } from '../utils/logger';

export interface CellComment {
  cellAddress: string;
  worksheetName: string;
  text: string;
  author?: string;
  timestamp?: Date;
  resolved?: boolean;
}

export interface CommentCreateOptions {
  cellAddress: string;
  text: string;
  author?: string;
  worksheetName?: string;
}

export interface CommentUpdateOptions {
  newText?: string;
  resolved?: boolean;
}

export class CommentService {
  private static instance: CommentService;

  private constructor() {}

  static getInstance(): CommentService {
    if (!CommentService.instance) {
      CommentService.instance = new CommentService();
    }
    return CommentService.instance;
  }

  /**
   * Add a comment to a cell
   */
  async addComment(options: CommentCreateOptions): Promise<CellComment> {
    return Excel.run(async (context) => {
      const worksheet = options.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(options.cellAddress);
      
      // Check if comment already exists
      range.load('comment');
      await context.sync();
      
      if (range.comment) {
        // Update existing comment
        range.comment.content = options.text;
        if (options.author) {
          range.comment.author = options.author;
        }
      } else {
        // Add new comment using the modern comments API
        // Note: Office.js has limited comment support, using legacy approach
        range.comments.add(options.text);
      }
      
      await context.sync();
      
      return {
        cellAddress: options.cellAddress,
        worksheetName: worksheet.name,
        text: options.text,
        author: options.author,
        timestamp: new Date(),
        resolved: false
      };
    }).catch((error) => {
      logger.error('Failed to add comment', { error, options });
      throw new Error(`Failed to add comment: ${error.message}`);
    });
  }

  /**
   * Get comment from a specific cell
   */
  async getComment(cellAddress: string, worksheetName?: string): Promise<CellComment | null> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(cellAddress);
      range.load({
        comment: true,
        address: true
      });
      
      await context.sync();
      
      if (range.comment && range.comment.content) {
        return {
          cellAddress: range.address,
          worksheetName: worksheet.name,
          text: range.comment.content,
          author: range.comment.author,
          timestamp: range.comment.dateTime ? new Date(range.comment.dateTime) : undefined,
          resolved: false // Office.js doesn't expose resolved state directly
        };
      }
      
      return null;
    }).catch((error) => {
      logger.error('Failed to get comment', { error, cellAddress, worksheetName });
      return null;
    });
  }

  /**
   * Get all comments in the workbook or current worksheet
   */
  async getAllComments(scope: 'workbook' | 'worksheet' = 'worksheet'): Promise<CellComment[]> {
    return Excel.run(async (context) => {
      const comments: CellComment[] = [];
      
      const worksheets = scope === 'workbook' 
        ? context.workbook.worksheets 
        : context.workbook.worksheets.getActiveWorksheet();
      
      if (scope === 'workbook') {
        worksheets.load('items/name');
        await context.sync();
        
        for (const worksheet of worksheets.items) {
          const wsComments = await this.getWorksheetComments(worksheet.name, context);
          comments.push(...wsComments);
        }
      } else {
        const activeWs = context.workbook.worksheets.getActiveWorksheet();
        activeWs.load('name');
        await context.sync();
        const wsComments = await this.getWorksheetComments(activeWs.name, context);
        comments.push(...wsComments);
      }
      
      return comments;
    }).catch((error) => {
      logger.error('Failed to get all comments', { error, scope });
      return [];
    });
  }

  /**
   * Helper to get comments from a worksheet
   * Note: Office.js has limited comment iteration support
   * This implementation uses the modern comments API where available
   */
  private async getWorksheetComments(worksheetName: string, context: Excel.RequestContext): Promise<CellComment[]> {
    const comments: CellComment[] = [];
    const worksheet = context.workbook.worksheets.getItem(worksheetName);

    try {
      // Try to use the modern worksheet.comments API (ExcelApi 1.13+)
      const wsComments = worksheet.comments;
      wsComments.load('items');
      await context.sync();

      if (wsComments.items && wsComments.items.length > 0) {
        for (const comment of wsComments.items) {
          comment.load('id, content, author, resolved, ranges');
          await context.sync();

          // Get the cell address from the comment's range
          let cellAddress = 'A1'; // Default fallback
          if (comment.ranges && comment.ranges.length > 0) {
            const range = comment.ranges.getItemAt(0);
            range.load('address');
            await context.sync();
            cellAddress = range.address;
          }

          comments.push({
            cellAddress,
            worksheetName,
            text: comment.content,
            author: comment.author,
            resolved: comment.resolved,
            timestamp: new Date()
          });
        }
      }
    } catch (error) {
      // Modern comments API not available, fall back to legacy approach
      logger.warn('Modern comments API not available, using legacy approach', { error });

      // Legacy approach: scan used range for cells with comments
      const usedRange = worksheet.getUsedRange();
      if (usedRange) {
        usedRange.load('address, rowCount, columnCount');
        await context.sync();

        if (usedRange.address && usedRange.rowCount > 0 && usedRange.columnCount > 0) {
          // Limit scanning to prevent performance issues
          const maxRows = Math.min(usedRange.rowCount, 100);
          const maxCols = Math.min(usedRange.columnCount, 20);

          for (let row = 0; row < maxRows; row++) {
            for (let col = 0; col < maxCols; col++) {
              try {
                const cell = usedRange.getCell(row, col);
                cell.load('comment');
                await context.sync();

                if (cell.comment) {
                  cell.comment.load('content, author');
                  await context.sync();

                  const address = await this.getCellAddress(cell, context);
                  comments.push({
                    cellAddress: address,
                    worksheetName,
                    text: cell.comment.content,
                    author: cell.comment.author,
                    timestamp: new Date(),
                    resolved: false
                  });
                }
              } catch {
                // Skip cells that can't be accessed
              }
            }
          }
        }
      }
    }

    return comments;
  }

  /**
   * Helper to get cell address from a range object
   */
  private async getCellAddress(range: Excel.Range, context: Excel.RequestContext): Promise<string> {
    range.load('address');
    await context.sync();
    return range.address;
  }

  /**
   * Update an existing comment
   */
  async updateComment(cellAddress: string, options: CommentUpdateOptions, worksheetName?: string): Promise<CellComment> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(cellAddress);
      range.load('comment');
      
      await context.sync();
      
      if (!range.comment) {
        throw new Error(`No comment exists at ${cellAddress}`);
      }
      
      if (options.newText) {
        range.comment.content = options.newText;
      }
      
      await context.sync();
      
      return {
        cellAddress,
        worksheetName: worksheet.name,
        text: options.newText || '',
        author: range.comment.author,
        resolved: options.resolved
      };
    }).catch((error) => {
      logger.error('Failed to update comment', { error, cellAddress, options });
      throw new Error(`Failed to update comment: ${error.message}`);
    });
  }

  /**
   * Delete a comment from a cell
   */
  async deleteComment(cellAddress: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(cellAddress);
      range.load('comment');
      
      await context.sync();
      
      if (range.comment) {
        range.comment.delete();
        await context.sync();
      }
    }).catch((error) => {
      logger.error('Failed to delete comment', { error, cellAddress, worksheetName });
      throw new Error(`Failed to delete comment: ${error.message}`);
    });
  }

  /**
   * Delete all comments in a range
   * Uses batch operations for better performance
   */
  async deleteCommentsInRange(rangeAddress: string, worksheetName?: string): Promise<number> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(rangeAddress);
      range.load('address, rowCount, columnCount');

      await context.sync();

      let deletedCount = 0;
      const totalCells = range.rowCount * range.columnCount;

      // Limit iteration for performance
      const maxRows = Math.min(range.rowCount, 100);
      const maxCols = Math.min(range.columnCount, 20);

      try {
        // Try modern comments API first (ExcelApi 1.13+)
        const wsComments = worksheet.comments;
        wsComments.load('items');
        await context.sync();

        if (wsComments.items && wsComments.items.length > 0) {
          for (const comment of wsComments.items) {
            comment.load('ranges');
            await context.sync();

            // Check if comment is within our range
            if (comment.ranges && comment.ranges.length > 0) {
              const commentRange = comment.ranges.getItemAt(0);
              commentRange.load('address');
              await context.sync();

              // Simple check if comment address is within our range
              if (commentRange.address.startsWith(range.address.split('!')[1] || range.address)) {
                comment.delete();
                deletedCount++;
              }
            }
          }
        }
      } catch {
        // Modern API not available, use cell-by-cell approach
        for (let row = 0; row < maxRows; row++) {
          for (let col = 0; col < maxCols; col++) {
            try {
              const cell = range.getCell(row, col);
              cell.load('comment');
              await context.sync();

              if (cell.comment) {
                cell.comment.delete();
                deletedCount++;
              }
            } catch {
              // Skip cells that can't be accessed
            }
          }
        }
      }

      await context.sync();
      return deletedCount;
    }).catch((error) => {
      logger.error('Failed to delete comments in range', { error, rangeAddress, worksheetName });
      return 0;
    });
  }

  /**
   * Delete all comments in the workbook
   */
  async deleteAllComments(scope: 'workbook' | 'worksheet' = 'worksheet'): Promise<number> {
    // This is expensive - iterate through all worksheets
    const comments = await this.getAllComments(scope);
    let deleted = 0;
    
    for (const comment of comments) {
      try {
        await this.deleteComment(comment.cellAddress, comment.worksheetName);
        deleted++;
      } catch (e) {
        // Continue even if one fails
      }
    }
    
    return deleted;
  }

  /**
   * Show/hide comments visibility
   */
  async setCommentsVisibility(visible: boolean): Promise<void> {
    return Excel.run(async (context) => {
      // Office.js doesn't have direct comment visibility control
      // Comments are shown based on user settings
      // We can show individual comments but not globally
      logger.warn('Comment visibility control is limited in Office.js');
    });
  }

  /**
   * Show a specific comment (make it visible)
   */
  async showComment(cellAddress: string, worksheetName?: string): Promise<void> {
    // Office.js doesn't have a method to programmatically show comments
    // User must manually show them via Excel UI
    logger.warn('Programmatic comment show is not supported in Office.js');
  }

  /**
   * Hide a specific comment
   */
  async hideComment(cellAddress: string, worksheetName?: string): Promise<void> {
    // Office.js doesn't have a method to programmatically hide comments
    logger.warn('Programmatic comment hide is not supported in Office.js');
  }

  /**
   * Reply to an existing comment (threaded comments)
   */
  async replyToComment(cellAddress: string, replyText: string, author?: string, worksheetName?: string): Promise<CellComment> {
    // Office.js doesn't support threaded comments in the same way
    // We can update the existing comment content
    return this.updateComment(cellAddress, { newText: replyText }, worksheetName);
  }

  /**
   * Mark a comment as resolved/unresolved
   */
  async setCommentResolved(cellAddress: string, resolved: boolean, worksheetName?: string): Promise<void> {
    // Office.js doesn't support resolution state for comments
    logger.warn('Comment resolution is not supported in Office.js');
  }

  /**
   * Get comments by author
   */
  async getCommentsByAuthor(author: string): Promise<CellComment[]> {
    const allComments = await this.getAllComments('workbook');
    return allComments.filter(c => c.author === author);
  }

  /**
   * Search comments by text content
   */
  async searchComments(searchText: string): Promise<CellComment[]> {
    const allComments = await this.getAllComments('workbook');
    return allComments.filter(c =>
      c.text.toLowerCase().includes(searchText.toLowerCase())
    );
  }

  /**
   * Count comments in scope
   */
  async countComments(scope: 'workbook' | 'worksheet' = 'worksheet'): Promise<number> {
    const comments = await this.getAllComments(scope);
    return comments.length;
  }

  /**
   * Check if a cell has a comment
   */
  async cellHasComment(cellAddress: string, worksheetName?: string): Promise<boolean> {
    const comment = await this.getComment(cellAddress, worksheetName);
    return comment !== null;
  }

  /**
   * Copy comments from source range to target range
   */
  async copyComments(sourceRange: string, targetRange: string, worksheetName?: string): Promise<number> {
    // Get source comment and apply to target
    const sourceComment = await this.getComment(sourceRange, worksheetName);
    if (!sourceComment) return 0;
    
    try {
      await this.addComment({
        cellAddress: targetRange,
        text: sourceComment.text,
        author: sourceComment.author,
        worksheetName
      });
      return 1;
    } catch (e) {
      return 0;
    }
  }

  /**
   * Add comment to a range (same comment for all cells)
   */
  async addCommentToRange(
    rangeAddress: string, 
    text: string, 
    options?: {
      author?: string;
      worksheetName?: string;
    }
  ): Promise<CellComment[]> {
    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      const range = worksheet.getRange(rangeAddress);
      range.load(['rowCount', 'columnCount']);
      
      await context.sync();
      
      const comments: CellComment[] = [];
      const rowCount = range.rowCount;
      const colCount = range.columnCount;
      
      // Note: Adding comments to many cells is expensive
      // For large ranges, consider using a single comment
      
      for (let row = 0; row < Math.min(rowCount, 100); row++) { // Limit to 100 cells
        for (let col = 0; col < Math.min(colCount, 10); col++) {
          try {
            const cellAddress = this.getCellAddress(row, col);
            const comment = await this.addComment({
              cellAddress,
              text,
              author: options?.author,
              worksheetName: worksheet.name
            });
            comments.push(comment);
          } catch (e) {
            // Continue on error
          }
        }
      }
      
      return comments;
    });
  }

  /**
   * Helper to convert row/col to cell address
   */
  private getCellAddress(row: number, col: number): string {
    const colLetter = String.fromCharCode(65 + col); // A=0, B=1, etc.
    return `${colLetter}${row + 1}`;
  }

  /**
   * Generate comment text using templates
   */
  generateCommentText(
    template: 'instruction' | 'warning' | 'note' | 'formula', 
    data: {
      description?: string;
      formula?: string;
      expectedValue?: string;
      warningMessage?: string;
    }
  ): string {
    switch (template) {
      case 'instruction':
        return `📝 Instruction: ${data.description || 'Enter value here'}`;
      
      case 'warning':
        return `⚠️ Warning: ${data.warningMessage || 'Check this value before proceeding'}`;
      
      case 'note':
        return `💡 Note: ${data.description || 'Additional information'}`;
      
      case 'formula':
        const formulaText = data.formula ? `=${data.formula}` : 'N/A';
        return `🔢 Formula: ${formulaText}\nExpected: ${data.expectedValue || 'Calculated value'}`;
      
      default:
        return data.description || '';
    }
  }

  /**
   * Get comment statistics
   */
  async getCommentStatistics(): Promise<{
    total: number;
    byWorksheet: Record<string, number>;
    byAuthor: Record<string, number>;
    resolved: number;
    unresolved: number;
  }> {
    const allComments = await this.getAllComments('workbook');
    
    const byWorksheet: Record<string, number> = {};
    const byAuthor: Record<string, number> = {};
    let resolved = 0;
    let unresolved = 0;
    
    for (const comment of allComments) {
      byWorksheet[comment.worksheetName] = (byWorksheet[comment.worksheetName] || 0) + 1;
      
      if (comment.author) {
        byAuthor[comment.author] = (byAuthor[comment.author] || 0) + 1;
      }
      
      if (comment.resolved) {
        resolved++;
      } else {
        unresolved++;
      }
    }
    
    return {
      total: allComments.length,
      byWorksheet,
      byAuthor,
      resolved,
      unresolved
    };
  }
}

export default CommentService.getInstance();
