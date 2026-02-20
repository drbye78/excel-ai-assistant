/**
 * Cell Reference Parser Utility
 * 
 * Parses cell references from text and AI responses to enable
 * automatic highlighting of cells mentioned by the AI assistant.
 * 
 * Supports:
 * - A1 notation (A1, B2, AB100)
 * - Range notation (A1:D10, B2:B100)
 * - Sheet-qualified references (Sheet1!A1, 'Sheet Name'!B2:D10)
 * - Table column references (Table1[Column], Table1[@Column])
 * - Named ranges
 * - R1C1 notation (R1C1, R[1]C[1])
 */

import { ExcelRange } from '../types';

export interface ParsedCellReference {
  /** The original reference string found in text */
  original: string;
  /** The sheet name (null if current sheet) */
  sheetName: string | null;
  /** The parsed range */
  range: ExcelRange;
  /** Type of reference */
  type: 'cell' | 'range' | 'table' | 'named' | 'r1c1';
  /** Whether the reference is absolute ($A$1) */
  isAbsolute: boolean;
  /** Start position in the original text */
  startIndex: number;
  /** End position in the original text */
  endIndex: number;
}

export interface TableReference {
  tableName: string;
  columnName: string;
  isThisRow: boolean;
}

/**
 * Regular expression patterns for different reference types
 */
const PATTERNS = {
  // Sheet name: 'Sheet Name' or SheetName
  SHEET_NAME: /(?:'([^']+)'|([A-Za-z0-9_]+))!/g,
  
  // A1 notation cell reference
  CELL_A1: /\$?[A-Z]+\$?\d+/g,
  
  // A1 notation range (Cell:Cell)
  RANGE_A1: /\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+/g,
  
  // R1C1 notation
  CELL_R1C1: /R\[?-?\d*\]?C\[?-?\d*\]?/gi,
  
  // Table structured references
  TABLE_REF: /([A-Za-z_][A-Za-z0-9_]*)\[([^\]]+)\]/g,
  
  // Named range (simplified - word that could be a named range)
  NAMED_RANGE: /\b([A-Za-z_][A-Za-z0-9_]*)\b/g,
  
  // Full reference with sheet (Sheet!Range or 'Sheet'!Range)
  FULL_REFERENCE: /(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/g,
};

/**
 * Parse all cell references from a text string
 */
export function parseCellReferences(text: string): ParsedCellReference[] {
  const references: ParsedCellReference[] = [];
  const seen = new Set<string>();
  
  // Parse full references with sheet names first
  parseFullReferences(text, references, seen);
  
  // Parse table structured references
  parseTableReferences(text, references, seen);
  
  // Parse simple ranges
  parseSimpleRanges(text, references, seen);
  
  // Parse simple cells
  parseSimpleCells(text, references, seen);
  
  // Sort by position in text
  references.sort((a, b) => a.startIndex - b.startIndex);
  
  return references;
}

/**
 * Parse full references including sheet names
 */
function parseFullReferences(
  text: string,
  references: ParsedCellReference[],
  seen: Set<string>
): void {
  const pattern = PATTERNS.FULL_REFERENCE;
  let match: RegExpExecArray | null;
  
  // Reset lastIndex
  pattern.lastIndex = 0;
  
  while ((match = pattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const sheetName = match[1] || match[2];
    const rangeStr = match[3];
    
    const key = `${sheetName}!${rangeStr}`;
    if (seen.has(key)) continue;
    seen.add(key);
    
    const range = parseA1Range(rangeStr);
    if (!range) continue;
    
    references.push({
      original: fullMatch,
      sheetName: sheetName,
      range,
      type: rangeStr.includes(':') ? 'range' : 'cell',
      isAbsolute: rangeStr.includes('$'),
      startIndex: match.index,
      endIndex: match.index + fullMatch.length,
    });
  }
}

/**
 * Parse table structured references
 */
function parseTableReferences(
  text: string,
  references: ParsedCellReference[],
  seen: Set<string>
): void {
  const pattern = PATTERNS.TABLE_REF;
  let match: RegExpExecArray | null;
  
  pattern.lastIndex = 0;
  
  while ((match = pattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const tableName = match[1];
    const columnRef = match[2];
    
    if (seen.has(fullMatch)) continue;
    seen.add(fullMatch);
    
    const isThisRow = columnRef.startsWith('@');
    const columnName = isThisRow ? columnRef.substring(1) : columnRef;
    
    // For table references, we'll create a placeholder range
    // The actual resolution needs to happen at runtime
    references.push({
      original: fullMatch,
      sheetName: null,
      range: {
        sheetId: null,
        address: fullMatch,
        startCell: { column: 0, row: 0 },
        endCell: { column: 0, row: 0 },
      },
      type: 'table',
      isAbsolute: false,
      startIndex: match.index,
      endIndex: match.index + fullMatch.length,
    });
  }
}

/**
 * Parse simple A1 ranges (without sheet qualification)
 */
function parseSimpleRanges(
  text: string,
  references: ParsedCellReference[],
  seen: Set<string>
): void {
  const pattern = PATTERNS.RANGE_A1;
  let match: RegExpExecArray | null;
  
  pattern.lastIndex = 0;
  
  while ((match = pattern.exec(text)) !== null) {
    const rangeStr = match[0];
    
    // Skip if this is part of a full reference (already parsed)
    if (isPartOfFullReference(text, match.index)) continue;
    
    if (seen.has(rangeStr)) continue;
    seen.add(rangeStr);
    
    const range = parseA1Range(rangeStr);
    if (!range) continue;
    
    references.push({
      original: rangeStr,
      sheetName: null,
      range,
      type: 'range',
      isAbsolute: rangeStr.includes('$'),
      startIndex: match.index,
      endIndex: match.index + rangeStr.length,
    });
  }
}

/**
 * Parse simple A1 cell references (without sheet qualification)
 */
function parseSimpleCells(
  text: string,
  references: ParsedCellReference[],
  seen: Set<string>
): void {
  const pattern = PATTERNS.CELL_A1;
  let match: RegExpExecArray | null;
  
  pattern.lastIndex = 0;
  
  while ((match = pattern.exec(text)) !== null) {
    const cellStr = match[0];
    
    // Skip if this is part of a range (already parsed)
    if (isPartOfRange(text, match.index, cellStr.length)) continue;
    
    // Skip if this is part of a full reference
    if (isPartOfFullReference(text, match.index)) continue;
    
    if (seen.has(cellStr)) continue;
    seen.add(cellStr);
    
    const range = parseA1Cell(cellStr);
    if (!range) continue;
    
    references.push({
      original: cellStr,
      sheetName: null,
      range,
      type: 'cell',
      isAbsolute: cellStr.includes('$'),
      startIndex: match.index,
      endIndex: match.index + cellStr.length,
    });
  }
}

/**
 * Check if a position is part of a full reference (Sheet!Cell)
 */
function isPartOfFullReference(text: string, position: number): boolean {
  // Look backwards for ! without crossing whitespace or special chars
  for (let i = position - 1; i >= 0; i--) {
    const char = text[i];
    if (char === '!') return true;
    if (/\s|[^A-Za-z0-9_']/.test(char)) break;
  }
  return false;
}

/**
 * Check if a cell is part of a range (Cell:Cell)
 */
function isPartOfRange(text: string, position: number, length: number): boolean {
  const endPos = position + length;
  // Check if followed by :Cell or preceded by Cell:
  if (text[endPos] === ':') return true;
  
  // Check if preceded by :
  for (let i = position - 1; i >= 0; i--) {
    const char = text[i];
    if (char === ':') return true;
    if (/\s/.test(char)) break;
    if (/[A-Z]\d/i.test(char)) continue;
    break;
  }
  return false;
}

/**
 * Parse an A1 range string (e.g., "A1:D10" or "$A$1")
 */
export function parseA1Range(rangeStr: string): ExcelRange | null {
  // Remove all $ signs for parsing
  const cleanStr = rangeStr.replace(/\$/g, '');
  
  if (cleanStr.includes(':')) {
    // It's a range
    const [start, end] = cleanStr.split(':');
    const startCell = parseA1CellInternal(start);
    const endCell = parseA1CellInternal(end);
    
    if (!startCell || !endCell) return null;
    
    return {
      sheetId: null,
      address: rangeStr,
      startCell,
      endCell,
    };
  } else {
    // Single cell as range
    const cell = parseA1CellInternal(cleanStr);
    if (!cell) return null;
    
    return {
      sheetId: null,
      address: rangeStr,
      startCell: cell,
      endCell: cell,
    };
  }
}

/**
 * Parse an A1 cell string (e.g., "A1" or "$A$1")
 */
export function parseA1Cell(cellStr: string): ExcelRange | null {
  const cleanStr = cellStr.replace(/\$/g, '');
  const cell = parseA1CellInternal(cleanStr);
  
  if (!cell) return null;
  
  return {
    sheetId: null,
    address: cellStr,
    startCell: cell,
    endCell: cell,
  };
}

/**
 * Internal function to parse A1 cell coordinates
 */
function parseA1CellInternal(cellStr: string): { column: number; row: number } | null {
  const match = cellStr.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  
  const colLetters = match[1].toUpperCase();
  const rowNum = parseInt(match[2], 10);
  
  // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
  let colNum = 0;
  for (let i = 0; i < colLetters.length; i++) {
    colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
  }
  
  return {
    column: colNum,
    row: rowNum,
  };
}

/**
 * Convert column number to letters (1 -> A, 27 -> AA)
 */
export function columnNumberToLetters(colNum: number): string {
  let result = '';
  let num = colNum;
  
  while (num > 0) {
    num--;
    result = String.fromCharCode(65 + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  
  return result;
}

/**
 * Convert ExcelRange to A1 notation string
 */
export function rangeToA1(range: ExcelRange, absolute: boolean = false): string {
  const startCol = columnNumberToLetters(range.startCell.column);
  const startRow = range.startCell.row;
  
  if (range.startCell.column === range.endCell.column && 
      range.startCell.row === range.endCell.row) {
    // Single cell
    return absolute ? `$${startCol}$${startRow}` : `${startCol}${startRow}`;
  }
  
  // Range
  const endCol = columnNumberToLetters(range.endCell.column);
  const endRow = range.endCell.row;
  
  if (absolute) {
    return `$${startCol}$${startRow}:$${endCol}$${endRow}`;
  }
  return `${startCol}${startRow}:${endCol}${endRow}`;
}

/**
 * Parse R1C1 notation (e.g., "R1C1", "R[-1]C[2]")
 */
export function parseR1C1(r1c1Str: string, baseRow: number = 1, baseCol: number = 1): ExcelRange | null {
  const match = r1c1Str.match(/^R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?$/i);
  if (!match) return null;
  
  let row: number;
  let col: number;
  
  // Parse row
  if (!match[1]) {
    row = baseRow;
  } else if (match[1].startsWith('[')) {
    const offset = parseInt(match[1].slice(1, -1), 10);
    row = baseRow + offset;
  } else {
    row = parseInt(match[1], 10);
  }
  
  // Parse column
  if (!match[2]) {
    col = baseCol;
  } else if (match[2].startsWith('[')) {
    const offset = parseInt(match[2].slice(1, -1), 10);
    col = baseCol + offset;
  } else {
    col = parseInt(match[2], 10);
  }
  
  return {
    sheetId: null,
    address: r1c1Str,
    startCell: { column: col, row },
    endCell: { column: col, row },
  };
}

/**
 * Parse a table structured reference
 */
export function parseTableReference(refStr: string): TableReference | null {
  const match = refStr.match(/^([A-Za-z_][A-Za-z0-9_]*)\[(@?)([^\]]+)\]$/);
  if (!match) return null;
  
  return {
    tableName: match[1],
    columnName: match[3],
    isThisRow: match[2] === '@',
  };
}

/**
 * Extract cell references specifically from AI responses
 * This is optimized for parsing natural language responses
 */
export function extractReferencesFromAIResponse(response: string): ParsedCellReference[] {
  const references: ParsedCellReference[] = [];
  
  // Common patterns in AI responses
  const patterns = [
    // "cell A1", "cells A1 and B2"
    { regex: /(?:cell|cells)\s+([A-Z]+\d+(?::[A-Z]+\d+)?)/gi, type: 'explicit' },
    // "range A1:D10"
    { regex: /(?:range|ranges)\s+([A-Z]+\d+:[A-Z]+\d+)/gi, type: 'explicit' },
    // "in A1", "at B2"
    { regex: /(?:in|at)\s+([A-Z]+\d+)/gi, type: 'context' },
    // "column A", "row 1"
    { regex: /(?:column|col)\s+([A-Z]+)/gi, type: 'column' },
    { regex: /(?:row|rows)\s+(\d+)/gi, type: 'row' },
    // "Sheet1!A1" pattern
    { regex: /(?:sheet\s+)?['"]?([^'"\s]+)['"]?\s*!\s*([A-Z]+\d+)/gi, type: 'sheet' },
  ];
  
  for (const pattern of patterns) {
    let match: RegExpExecArray | null;
    pattern.regex.lastIndex = 0;
    
    while ((match = pattern.regex.exec(response)) !== null) {
      const refStr = match[1] + (match[2] || '');
      const parsed = parseA1Range(refStr) || parseA1Cell(refStr);
      
      if (parsed) {
        references.push({
          original: match[0],
          sheetName: null,
          range: parsed,
          type: parsed.startCell === parsed.endCell ? 'cell' : 'range',
          isAbsolute: false,
          startIndex: match.index,
          endIndex: match.index + match[0].length,
        });
      }
    }
  }
  
  // Also do general parsing
  const generalRefs = parseCellReferences(response);
  references.push(...generalRefs);
  
  // Deduplicate
  const seen = new Set<string>();
  return references.filter(ref => {
    const key = `${ref.sheetName}:${ref.range.address}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/**
 * Get surrounding context for a reference in text
 */
export function getReferenceContext(
  text: string,
  reference: ParsedCellReference,
  contextChars: number = 50
): string {
  const start = Math.max(0, reference.startIndex - contextChars);
  const end = Math.min(text.length, reference.endIndex + contextChars);
  
  let context = text.substring(start, end);
  
  // Add ellipsis if truncated
  if (start > 0) context = '...' + context;
  if (end < text.length) context = context + '...';
  
  return context;
}

/**
 * Group references by sheet
 */
export function groupReferencesBySheet(
  references: ParsedCellReference[]
): Map<string | null, ParsedCellReference[]> {
  const groups = new Map<string | null, ParsedCellReference[]>();
  
  for (const ref of references) {
    const key = ref.sheetName;
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key)!.push(ref);
  }
  
  return groups;
}

/**
 * Merge overlapping or adjacent ranges
 */
export function mergeRanges(ranges: ExcelRange[]): ExcelRange[] {
  if (ranges.length <= 1) return ranges;
  
  // Sort by start cell
  const sorted = [...ranges].sort((a, b) => {
    if (a.startCell.row !== b.startCell.row) {
      return a.startCell.row - b.startCell.row;
    }
    return a.startCell.column - b.startCell.column;
  });
  
  const merged: ExcelRange[] = [];
  let current = sorted[0];
  
  for (let i = 1; i < sorted.length; i++) {
    const next = sorted[i];
    
    // Check if ranges overlap or are adjacent
    if (next.startCell.row <= current.endCell.row + 1 &&
        next.startCell.column <= current.endCell.column + 1) {
      // Merge ranges
      current = {
        sheetId: current.sheetId,
        address: `${current.address},${next.address}`,
        startCell: current.startCell,
        endCell: {
          row: Math.max(current.endCell.row, next.endCell.row),
          column: Math.max(current.endCell.column, next.endCell.column),
        },
      };
    } else {
      merged.push(current);
      current = next;
    }
  }
  
  merged.push(current);
  return merged;
}

/**
 * Convert range to Excel API format
 */
export function toExcelApiRange(range: ExcelRange): string {
  const a1Notation = rangeToA1(range, false);
  
  if (range.startCell === range.endCell) {
    return a1Notation;
  }
  
  return a1Notation;
}

/**
 * Validate if a string is a valid cell reference
 */
export function isValidCellReference(str: string): boolean {
  // Check A1 format
  if (/^\$?[A-Z]+\$?\d+$/i.test(str)) {
    const parsed = parseA1CellInternal(str.replace(/\$/g, ''));
    if (parsed) {
      // Check Excel limits (XFD1048576 in modern Excel)
      return parsed.column <= 16384 && parsed.row <= 1048576;
    }
  }
  
  // Check A1:D10 format
  if (/^\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+$/i.test(str)) {
    const parsed = parseA1Range(str);
    return parsed !== null;
  }
  
  return false;
}

/**
 * Expand a column reference to a full range (e.g., "A" -> "A1:A1048576")
 */
export function expandColumnToRange(colLetter: string): ExcelRange | null {
  const colNum = parseA1CellInternal(`${colLetter}1`)?.column;
  if (!colNum) return null;
  
  return {
    sheetId: null,
    address: `${colLetter}:${colLetter}`,
    startCell: { column: colNum, row: 1 },
    endCell: { column: colNum, row: 1048576 },
  };
}

/**
 * Expand a row reference to a full range (e.g., "5" -> "A5:XFD5")
 */
export function expandRowToRange(rowNum: number): ExcelRange | null {
  if (rowNum < 1 || rowNum > 1048576) return null;
  
  return {
    sheetId: null,
    address: `${rowNum}:${rowNum}`,
    startCell: { column: 1, row: rowNum },
    endCell: { column: 16384, row: rowNum },
  };
}

/**
 * Calculate the size of a range (number of cells)
 */
export function getRangeSize(range: ExcelRange): number {
  const cols = range.endCell.column - range.startCell.column + 1;
  const rows = range.endCell.row - range.startCell.row + 1;
  return cols * rows;
}

/**
 * Check if a cell is within a range
 */
export function isCellInRange(
  cell: { column: number; row: number },
  range: ExcelRange
): boolean {
  return cell.column >= range.startCell.column &&
         cell.column <= range.endCell.column &&
         cell.row >= range.startCell.row &&
         cell.row <= range.endCell.row;
}

/**
 * Get the intersection of two ranges
 */
export function getRangeIntersection(a: ExcelRange, b: ExcelRange): ExcelRange | null {
  const startCol = Math.max(a.startCell.column, b.startCell.column);
  const endCol = Math.min(a.endCell.column, b.endCell.column);
  const startRow = Math.max(a.startCell.row, b.startCell.row);
  const endRow = Math.min(a.endCell.row, b.endCell.row);
  
  if (startCol > endCol || startRow > endRow) {
    return null; // No intersection
  }
  
  return {
    sheetId: a.sheetId,
    address: `${columnNumberToLetters(startCol)}${startRow}:${columnNumberToLetters(endCol)}${endRow}`,
    startCell: { column: startCol, row: startRow },
    endCell: { column: endCol, row: endRow },
  };
}

/**
 * Parse multiple references from comma-separated list
 */
export function parseMultipleReferences(refList: string): ParsedCellReference[] {
  const parts = refList.split(',').map(p => p.trim()).filter(Boolean);
  const references: ParsedCellReference[] = [];
  
  for (const part of parts) {
    const parsed = parseCellReferences(part);
    references.push(...parsed);
  }
  
  return references;
}

/**
 * Cell Reference Parser singleton instance
 */
export const cellReferenceParser = {
  parseCellReferences,
  parseA1Range,
  parseA1Cell,
  parseR1C1,
  parseTableReference,
  extractReferencesFromAIResponse,
  getReferenceContext,
  groupReferencesBySheet,
  mergeRanges,
  toExcelApiRange,
  isValidCellReference,
  expandColumnToRange,
  expandRowToRange,
  getRangeSize,
  isCellInRange,
  getRangeIntersection,
  parseMultipleReferences,
  columnNumberToLetters,
  rangeToA1,
};

export default cellReferenceParser;
