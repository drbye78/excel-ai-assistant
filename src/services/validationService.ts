// Validation Service - Manages Data Validation Rules in Excel
// Production Implementation with Office.js API

import { logger } from '../utils/logger';

export interface ValidationRule {
  range: string;
  type: 'list' | 'date' | 'number' | 'textLength' | 'custom' | 'any';
  criteria?: {
    operator?: 'between' | 'notBetween' | 'equalTo' | 'notEqualTo' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqualTo' | 'lessThanOrEqualTo';
    value1?: string | number;
    value2?: string | number;
  };
  allowedValues?: string[];
  customFormula?: string;
  ignoreBlank?: boolean;
  showInputMessage?: boolean;
  inputMessage?: string;
  inputTitle?: string;
  showErrorMessage?: boolean;
  errorMessage?: string;
  errorTitle?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
}

export interface ValidationCheckResult {
  valid: boolean;
  violations: Array<{
    cell: string;
    issue: string;
    currentValue: any;
  }>;
  summary: {
    totalChecked: number;
    violationsFound: number;
  };
}

// Validation cache for storing validation rules
const validationCache = new Map<string, ValidationRule>();

export class ValidationService {
  private static instance: ValidationService;

  private constructor() {}

  static getInstance(): ValidationService {
    if (!ValidationService.instance) {
      ValidationService.instance = new ValidationService();
    }
    return ValidationService.instance;
  }

  /**
   * Apply a validation rule to a range using Office.js API
   */
  private async applyValidationRule(
    worksheet: Excel.Worksheet,
    rangeAddress: string,
    rule: ValidationRule
  ): Promise<void> {
    const range = worksheet.getRange(rangeAddress);
    const dataValidation = range.dataValidation;
    
    // Set validation type
    switch (rule.type) {
      case 'list':
        if (rule.allowedValues) {
          dataValidation.rule = {
            type: 'List',
            formula1: rule.allowedValues.join(','),
            ignoreBlank: rule.ignoreBlank
          };
        }
        break;
      case 'number':
        if (rule.criteria) {
          dataValidation.rule = {
            type: 'Decimal',
            operator: this.mapOperator(rule.criteria.operator),
            formula1: String(rule.criteria.value1 || ''),
            formula2: String(rule.criteria.value2 || ''),
            ignoreBlank: rule.ignoreBlank
          };
        }
        break;
      case 'date':
        if (rule.criteria) {
          dataValidation.rule = {
            type: 'Date',
            operator: this.mapOperator(rule.criteria.operator),
            formula1: String(rule.criteria.value1 || ''),
            formula2: String(rule.criteria.value2 || ''),
            ignoreBlank: rule.ignoreBlank
          };
        }
        break;
      case 'textLength':
        if (rule.criteria) {
          dataValidation.rule = {
            type: 'TextLength',
            operator: this.mapOperator(rule.criteria.operator),
            formula1: String(rule.criteria.value1 || ''),
            formula2: String(rule.criteria.value2 || ''),
            ignoreBlank: rule.ignoreBlank
          };
        }
        break;
      case 'custom':
        if (rule.customFormula) {
          dataValidation.rule = {
            type: 'Custom',
            formula1: rule.customFormula,
            ignoreBlank: rule.ignoreBlank
          };
        }
        break;
    }

    // Set input message
    if (rule.showInputMessage && rule.inputMessage) {
      dataValidation.prompt = {
        message: rule.inputMessage,
        title: rule.inputTitle || 'Input'
      };
    }

    // Set error message
    if (rule.showErrorMessage && rule.errorMessage) {
      dataValidation.errorAlert = {
        message: rule.errorMessage,
        title: rule.errorTitle || 'Invalid',
        type: rule.errorStyle || 'Stop'
      };
    }
  }

  /**
   * Map operator string to Office.js format
   */
  private mapOperator(operator?: string): Excel.DataValidationOperator {
    const operatorMap: Record<string, Excel.DataValidationOperator> = {
      'between': 'Between',
      'notBetween': 'NotBetween',
      'equalTo': 'Equal',
      'notEqualTo': 'NotEqual',
      'greaterThan': 'GreaterThan',
      'lessThan': 'LessThan',
      'greaterThanOrEqualTo': 'GreaterThanOrEqual',
      'lessThanOrEqualTo': 'LessThanOrEqual'
    };
    return (operatorMap[operator || ''] || 'Between') as Excel.DataValidationOperator;
  }

  /**
   * Create and apply a list validation (dropdown)
   */
  async createListValidation(range: string, values: string[], options?: {
    ignoreBlank?: boolean;
    showInputMessage?: boolean;
    inputMessage?: string;
    inputTitle?: string;
    showErrorMessage?: boolean;
    errorMessage?: string;
    errorTitle?: string;
    errorStyle?: 'stop' | 'warning' | 'information';
    worksheetName?: string;
  }): Promise<ValidationRule> {
    const rule: ValidationRule = {
      range,
      type: 'list',
      allowedValues: values,
      ignoreBlank: options?.ignoreBlank ?? true,
      showInputMessage: options?.showInputMessage ?? true,
      inputMessage: options?.inputMessage,
      inputTitle: options?.inputTitle,
      showErrorMessage: options?.showErrorMessage ?? true,
      errorMessage: options?.errorMessage,
      errorTitle: options?.errorTitle,
      errorStyle: options?.errorStyle
    };

    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, range, rule);
      await context.sync();

      // Cache the rule
      validationCache.set(`${worksheet.name}!${range}`, rule);

      return rule;
    }).catch((error) => {
      logger.error('Failed to create list validation', { error, range, allowedValues, worksheetName });
      throw new Error(`Failed to create list validation: ${error.message}`);
    });
  }

  /**
   * Create and apply a number validation
   */
  async createNumberValidation(
    range: string, 
    criteria: {
      operator: NonNullable<ValidationRule['criteria']>['operator'];
      value1: number;
      value2?: number;
    }, 
    options?: {
      ignoreBlank?: boolean;
      showInputMessage?: boolean;
      inputMessage?: string;
      showErrorMessage?: boolean;
      errorMessage?: string;
      worksheetName?: string;
    }
  ): Promise<ValidationRule> {
    const rule: ValidationRule = {
      range,
      type: 'number',
      criteria,
      ignoreBlank: options?.ignoreBlank ?? true
    };

    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, range, rule);
      await context.sync();

      validationCache.set(`${worksheet.name}!${range}`, rule);
      return rule;
    }).catch((error) => {
      logger.error('Failed to create number validation', { error, range, criteria, worksheetName });
      throw new Error(`Failed to create number validation: ${error.message}`);
    });
  }

  /**
   * Create and apply a date validation
   */
  async createDateValidation(
    range: string, 
    criteria: {
      operator: NonNullable<ValidationRule['criteria']>['operator'];
      value1: string;
      value2?: string;
    }, 
    options?: {
      ignoreBlank?: boolean;
      worksheetName?: string;
    }
  ): Promise<ValidationRule> {
    const rule: ValidationRule = {
      range,
      type: 'date',
      criteria: {
        operator: criteria.operator,
        value1: criteria.value1,
        value2: criteria.value2
      },
      ignoreBlank: options?.ignoreBlank ?? true
    };

    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, range, rule);
      await context.sync();

      validationCache.set(`${worksheet.name}!${range}`, rule);
      return rule;
    }).catch((error) => {
      logger.error('Failed to create date validation', { error, range, criteria, worksheetName });
      throw new Error(`Failed to create date validation: ${error.message}`);
    });
  }

  /**
   * Create and apply a text length validation
   */
  async createTextLengthValidation(
    range: string, 
    criteria: {
      operator: NonNullable<ValidationRule['criteria']>['operator'];
      length1: number;
      length2?: number;
    },
    options?: {
      ignoreBlank?: boolean;
      worksheetName?: string;
    }
  ): Promise<ValidationRule> {
    const rule: ValidationRule = {
      range,
      type: 'textLength',
      criteria: {
        operator: criteria.operator,
        value1: criteria.length1,
        value2: criteria.length2
      },
      ignoreBlank: options?.ignoreBlank ?? true
    };

    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, range, rule);
      await context.sync();

      validationCache.set(`${worksheet.name}!${range}`, rule);
      return rule;
    }).catch((error) => {
      logger.error('Failed to create text length validation', { error, range, criteria, worksheetName });
      throw new Error(`Failed to create text length validation: ${error.message}`);
    });
  }

  /**
   * Create and apply a custom formula validation
   */
  async createCustomValidation(
    range: string, 
    formula: string, 
    options?: {
      ignoreBlank?: boolean;
      worksheetName?: string;
    }
  ): Promise<ValidationRule> {
    const rule: ValidationRule = {
      range,
      type: 'custom',
      customFormula: formula,
      ignoreBlank: options?.ignoreBlank ?? true
    };

    return Excel.run(async (context) => {
      const worksheet = options?.worksheetName
        ? context.workbook.worksheets.getItem(options.worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, range, rule);
      await context.sync();

      validationCache.set(`${worksheet.name}!${range}`, rule);
      return rule;
    }).catch((error) => {
      logger.error('Failed to create custom validation', { error, range, customFormula, worksheetName });
      throw new Error(`Failed to create custom validation: ${error.message}`);
    });
  }

  /**
   * Remove validation from a range
   */
  async removeValidation(range: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.dataValidation.clear();
      
      await context.sync();
      
      // Remove from cache
      validationCache.delete(`${worksheet.name}!${range}`);
    }).catch((error) => {
      logger.error('Failed to remove validation', { error, range, worksheetName });
      throw new Error(`Failed to remove validation: ${error.message}`);
    });
  }

  /**
   * Get validation rule for a range
   */
  async getValidation(range: string, worksheetName?: string): Promise<ValidationRule | null> {
    // Check cache first
    const cacheKey = `${worksheetName || 'active'}!${range}`;
    if (validationCache.has(cacheKey)) {
      return validationCache.get(cacheKey)!;
    }

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.dataValidation.load(['rule', 'prompt', 'errorAlert']);
      
      await context.sync();
      
      const dv = rangeObj.dataValidation;
      if (!dv || !dv.rule || !dv.rule.type) {
        return null;
      }

      // Convert Office.js rule back to our ValidationRule format
      const rule: ValidationRule = {
        range,
        type: this.mapValidationType(dv.rule.type),
        ignoreBlank: dv.rule.ignoreBlank
      };

      if (dv.rule.formula1) {
        rule.criteria = {
          operator: this.mapOperatorBack(dv.rule.operator)
        };
        
        if (dv.rule.type === 'List') {
          rule.allowedValues = dv.rule.formula1.split(',');
        } else {
          rule.criteria.value1 = dv.rule.formula1;
          rule.criteria.value2 = dv.rule.formula2;
        }
      }

      if (dv.prompt && dv.prompt.message) {
        rule.showInputMessage = true;
        rule.inputMessage = dv.prompt.message;
        rule.inputTitle = dv.prompt.title;
      }

      if (dv.errorAlert && dv.errorAlert.message) {
        rule.showErrorMessage = true;
        rule.errorMessage = dv.errorAlert.message;
        rule.errorTitle = dv.errorAlert.title;
        rule.errorStyle = dv.errorAlert.type as 'stop' | 'warning' | 'information';
      }

      // Cache the rule
      validationCache.set(`${worksheet.name}!${range}`, rule);
      
      return rule;
    }).catch((error) => {
      logger.error('Failed to get validation', { error, range, worksheetName });
      return null;
    });
  }

  /**
   * Map Office.js validation type to our type
   */
  private mapValidationType(type: string): ValidationRule['type'] {
    const typeMap: Record<string, ValidationRule['type']> = {
      'List': 'list',
      'Decimal': 'number',
      'Date': 'date',
      'TextLength': 'textLength',
      'Custom': 'custom',
      'Any': 'any'
    };
    return typeMap[type] || 'any';
  }

  /**
   * Map operator back to our format
   */
  private mapOperatorBack(operator?: Excel.DataValidationOperator): ValidationRule['criteria']['operator'] {
    const operatorMap: Record<string, ValidationRule['criteria']['operator']> = {
      'Between': 'between',
      'NotBetween': 'notBetween',
      'Equal': 'equalTo',
      'NotEqual': 'notEqualTo',
      'GreaterThan': 'greaterThan',
      'LessThan': 'lessThan',
      'GreaterThanOrEqual': 'greaterThanOrEqualTo',
      'LessThanOrEqual': 'lessThanOrEqualTo'
    };
    return (operatorMap[operator || ''] || 'between') as ValidationRule['criteria']['operator'];
  }

  /**
   * Check if range has validation
   */
  async hasValidation(range: string, worksheetName?: string): Promise<boolean> {
    const rule = await this.getValidation(range, worksheetName);
    return rule !== null;
  }

  /**
   * Validate data against rules and return violations
   */
  async validateData(range: string, worksheetName?: string): Promise<ValidationCheckResult> {
    const violations: ValidationCheckResult['violations'] = [];
    let totalChecked = 0;

    const rule = await this.getValidation(range, worksheetName);
    if (!rule) {
      return {
        valid: true,
        violations: [],
        summary: { totalChecked: 0, violationsFound: 0 }
      };
    }

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      const values = rangeObj.values;
      const rowCount = rangeObj.rowCount;
      const columnCount = rangeObj.columnCount;

      for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < columnCount; col++) {
          const cellValue = values[row][col];
          const cellAddress = this.getCellAddress(row, col);
          totalChecked++;

          // Skip blank cells if ignoreBlank is true
          if (rule.ignoreBlank && (cellValue === null || cellValue === undefined || cellValue === '')) {
            continue;
          }

          // Validate based on rule type
          const violation = this.validateCellValue(cellValue, rule, cellAddress);
          if (violation) {
            violations.push(violation);
          }
        }
      }

      return {
        valid: violations.length === 0,
        violations,
        summary: { totalChecked, violationsFound: violations.length }
      };
    }).catch((error) => {
      logger.error('Failed to validate data', { error, range, worksheetName });
      return {
        valid: false,
        violations: [],
        summary: { totalChecked: 0, violationsFound: 0 }
      };
    });
  }

  /**
   * Validate a single cell value
   */
  private validateCellValue(
    value: any, 
    rule: ValidationRule, 
    cellAddress: string
  ): ValidationCheckResult['violations'][0] | null {
    switch (rule.type) {
      case 'list':
        if (rule.allowedValues && !rule.allowedValues.includes(String(value))) {
          return {
            cell: cellAddress,
            issue: `Value "${value}" is not in the allowed list`,
            currentValue: value
          };
        }
        break;
      case 'number':
        if (rule.criteria) {
          const numValue = parseFloat(value);
          if (isNaN(numValue)) {
            return { cell: cellAddress, issue: 'Value is not a number', currentValue: value };
          }
          if (!this.checkNumberCriteria(numValue, rule.criteria)) {
            return { cell: cellAddress, issue: 'Number does not meet criteria', currentValue: value };
          }
        }
        break;
      case 'date':
        if (rule.criteria) {
          const dateValue = new Date(value);
          if (isNaN(dateValue.getTime())) {
            return { cell: cellAddress, issue: 'Value is not a valid date', currentValue: value };
          }
        }
        break;
      case 'textLength':
        if (rule.criteria) {
          const strValue = String(value || '');
          const length = strValue.length;
          if (!this.checkLengthCriteria(length, rule.criteria)) {
            return { cell: cellAddress, issue: 'Text length does not meet criteria', currentValue: value };
          }
        }
        break;
    }
    return null;
  }

  /**
   * Check number against criteria
   */
  private checkNumberCriteria(value: number, criteria: ValidationRule['criteria']): boolean {
    if (!criteria.operator) return true;
    
    switch (criteria.operator) {
      case 'between':
        return value >= (criteria.value1 as number) && value <= (criteria.value2 as number);
      case 'notBetween':
        return value < (criteria.value1 as number) || value > (criteria.value2 as number);
      case 'equalTo':
        return value === criteria.value1;
      case 'notEqualTo':
        return value !== criteria.value1;
      case 'greaterThan':
        return value > (criteria.value1 as number);
      case 'lessThan':
        return value < (criteria.value1 as number);
      case 'greaterThanOrEqualTo':
        return value >= (criteria.value1 as number);
      case 'lessThanOrEqualTo':
        return value <= (criteria.value1 as number);
      default:
        return true;
    }
  }

  /**
   * Check length against criteria
   */
  private checkLengthCriteria(length: number, criteria: ValidationRule['criteria']): boolean {
    if (!criteria.operator) return true;
    
    switch (criteria.operator) {
      case 'between':
        return length >= (criteria.value1 as number) && length <= (criteria.value2 as number);
      case 'notBetween':
        return length < (criteria.value1 as number) || length > (criteria.value2 as number);
      case 'equalTo':
        return length === criteria.value1;
      case 'notEqualTo':
        return length !== criteria.value1;
      case 'greaterThan':
        return length > (criteria.value1 as number);
      case 'lessThan':
        return length < (criteria.value1 as number);
      case 'greaterThanOrEqualTo':
        return length >= (criteria.value1 as number);
      case 'lessThanOrEqualTo':
        return length <= (criteria.value1 as number);
      default:
        return true;
    }
  }

  /**
   * Get cell address from row and column
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

  /**
   * Find duplicate values in a range
   */
  async findDuplicates(range: string, worksheetName?: string): Promise<Array<{ cell: string; value: any; duplicates: string[] }>> {
    const duplicates: Map<string, { cell: string; value: any; cells: string[] }> = new Map();

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      const values = rangeObj.values;
      
      for (let row = 0; row < rangeObj.rowCount; row++) {
        for (let col = 0; col < rangeObj.columnCount; col++) {
          const value = values[row][col];
          if (value === null || value === undefined || value === '') continue;
          
          const key = String(value).toLowerCase();
          const cellAddress = this.getCellAddress(row, col);
          
          if (duplicates.has(key)) {
            const existing = duplicates.get(key)!;
            existing.cells.push(cellAddress);
          } else {
            duplicates.set(key, { cell: cellAddress, value, cells: [cellAddress] });
          }
        }
      }

      // Filter to only actual duplicates (more than one cell)
      return Array.from(duplicates.values())
        .filter(d => d.cells.length > 1)
        .map(d => ({
          cell: d.cell,
          value: d.value,
          duplicates: d.cells
        }));
    }).catch((error) => {
      logger.error('Failed to find duplicates', { error, range, worksheetName });
      return [];
    });
  }

  /**
   * Find blank/empty cells in a range
   */
  async findBlanks(range: string, worksheetName?: string): Promise<string[]> {
    const blanks: string[] = [];

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      for (let row = 0; row < rangeObj.rowCount; row++) {
        for (let col = 0; col < rangeObj.columnCount; col++) {
          const value = values[row][col];
          if (value === null || value === undefined || value === '') {
            blanks.push(this.getCellAddress(row, col));
          }
        }
      }
      
      return blanks;
    }).catch((error) => {
      logger.error('Failed to find blanks', { error, range, worksheetName });
      return [];
    });
  }

  /**
   * Find cells with errors
   */
  async findErrors(range: string, worksheetName?: string): Promise<Array<{ cell: string; errorType: string }>> {
    const errors: Array<{ cell: string; errorType: string }> = [];

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(range);
      rangeObj.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      const errorPatterns = ['#N/A', '#VALUE!', '#REF!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!', '#ERROR'];
      
      for (let row = 0; row < rangeObj.rowCount; row++) {
        for (let col = 0; col < rangeObj.columnCount; col++) {
          const value = values[row][col];
          if (typeof value === 'string') {
            const matchedError = errorPatterns.find(e => value.includes(e));
            if (matchedError) {
              errors.push({
                cell: this.getCellAddress(row, col),
                errorType: matchedError
              });
            }
          }
        }
      }
      
      return errors;
    }).catch((error) => {
      logger.error('Failed to find errors', { error, range, worksheetName });
      return [];
    });
  }

  /**
   * Get validation statistics for a worksheet
   */
  async getValidationStatistics(worksheetName?: string): Promise<{
    totalRules: number;
    rulesByType: Record<string, number>;
    rangesWithValidation: string[];
  }> {
    const rulesByType: Record<string, number> = {};
    const rangesWithValidation: string[] = [];

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      worksheet.load('name');
      
      // Get used range
      const usedRange = worksheet.getUsedRange();
      usedRange.load(['rowCount', 'columnCount']);
      
      await context.sync();
      
      // Scan for validations (this is a simplified approach)
      // In production, you'd track validations as they're created
      let totalRules = 0;
      
      for (const [key, rule] of validationCache.entries()) {
        if (key.startsWith(worksheet.name)) {
          totalRules++;
          rulesByType[rule.type] = (rulesByType[rule.type] || 0) + 1;
          rangesWithValidation.push(rule.range);
        }
      }

      return {
        totalRules,
        rulesByType,
        rangesWithValidation
      };
    }).catch((error) => {
      logger.error('Failed to get validation statistics', { error, worksheetName });
      return {
        totalRules: 0,
        rulesByType: {},
        rangesWithValidation: []
      };
    });
  }

  /**
   * Copy validation from one range to another
   */
  async copyValidation(sourceRange: string, targetRange: string, worksheetName?: string): Promise<void> {
    const sourceRule = await this.getValidation(sourceRange, worksheetName);
    if (!sourceRule) {
      throw new Error(`No validation found in source range ${sourceRange}`);
    }

    // Apply the same rule to target
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      await this.applyValidationRule(worksheet, targetRange, {
        ...sourceRule,
        range: targetRange
      });
      await context.sync();

      validationCache.set(`${worksheet.name}!${targetRange}`, {
        ...sourceRule,
        range: targetRange
      });
    }).catch((error) => {
      logger.error('Failed to copy validation', { error, sourceRange, targetRange, worksheetName });
      throw new Error(`Failed to copy validation: ${error.message}`);
    });
  }

  /**
   * Create validation from existing values (auto-detect list)
   */
  async createValidationFromValues(sourceRange: string, targetRange: string, worksheetName?: string): Promise<ValidationRule> {
    // Get unique values from source range
    const uniqueValues = new Set<string>();

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const rangeObj = worksheet.getRange(sourceRange);
      rangeObj.load(['values', 'rowCount', 'columnCount']);
      
      await context.sync();
      
      // Collect unique values
      for (let row = 0; row < rangeObj.rowCount; row++) {
        for (let col = 0; col < rangeObj.columnCount; col++) {
          const value = values[row][col];
          if (value !== null && value !== undefined && value !== '') {
            uniqueValues.add(String(value));
          }
        }
      }

      // Create list validation with unique values
      return this.createListValidation(
        targetRange,
        Array.from(uniqueValues),
        { worksheetName }
      );
    }).catch((error) => {
      logger.error('Failed to create validation from values', { error, sourceRange, targetRange, worksheetName });
      throw new Error(`Failed to create validation from values: ${error.message}`);
    });
  }

  /**
   * Generate common validation presets
   */
  getValidationPresets(): Array<{ name: string; description: string; rule: Partial<ValidationRule> }> {
    return [
      {
        name: 'Yes/No',
        description: 'Dropdown with Yes/No options',
        rule: { type: 'list', allowedValues: ['Yes', 'No'] }
      },
      {
        name: 'Status',
        description: 'Common status options',
        rule: { type: 'list', allowedValues: ['Pending', 'In Progress', 'Complete', 'Blocked'] }
      },
      {
        name: 'Priority',
        description: 'Priority levels',
        rule: { type: 'list', allowedValues: ['Low', 'Medium', 'High', 'Critical'] }
      },
      {
        name: 'Positive Number',
        description: 'Only positive numbers allowed',
        rule: { type: 'number', criteria: { operator: 'greaterThan', value1: 0 } }
      },
      {
        name: 'Percentage',
        description: '0-100 percentage range',
        rule: { type: 'number', criteria: { operator: 'between', value1: 0, value2: 100 } }
      },
      {
        name: 'Future Date',
        description: 'Only future dates allowed',
        rule: { type: 'date', criteria: { operator: 'greaterThan', value1: 'TODAY()' } }
      },
      {
        name: 'Email',
        description: 'Valid email format',
        rule: { type: 'custom', customFormula: '=ISNUMBER(SEARCH("@",A1))' }
      },
      {
        name: 'US States',
        description: 'Dropdown with US state abbreviations',
        rule: { type: 'list', allowedValues: ['AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA', 'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD', 'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ', 'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY'] }
      },
      {
        name: 'Currency',
        description: 'Non-negative currency values',
        rule: { type: 'number', criteria: { operator: 'greaterThanOrEqualTo', value1: 0 } }
      },
      {
        name: 'Whole Number',
        description: 'Only whole numbers allowed',
        rule: { type: 'custom', customFormula: '=MOD(A1,1)=0' }
      }
    ];
  }

  /**
   * Suggest validation type based on data analysis
   */
  suggestValidationType(sampleData: any[]): { type: string; confidence: number; reason: string } {
    if (!sampleData || sampleData.length === 0) {
      return { type: 'any', confidence: 0, reason: 'No data to analyze' };
    }

    const validData = sampleData.filter(v => v !== null && v !== undefined && v !== '');
    if (validData.length === 0) {
      return { type: 'any', confidence: 0, reason: 'No valid data to analyze' };
    }

    const uniqueValues = new Set(validData.map(v => String(v))).size;
    const allNumbers = validData.every(v => typeof v === 'number' || !isNaN(parseFloat(String(v))));
    const allDates = validData.every(v => !isNaN(Date.parse(String(v))));
    const allEmails = validData.every(v => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v)));

    if (uniqueValues <= 15 && uniqueValues / validData.length < 0.5) {
      return {
        type: 'list',
        confidence: 0.9,
        reason: `Only ${uniqueValues} unique values found, suggesting a dropdown list`
      };
    }

    if (allEmails) {
      return {
        type: 'custom',
        confidence: 0.95,
        reason: 'All values appear to be valid email addresses'
      };
    }

    if (allDates) {
      return {
        type: 'date',
        confidence: 0.85,
        reason: 'All values appear to be dates'
      };
    }

    if (allNumbers) {
      return {
        type: 'number',
        confidence: 0.8,
        reason: 'All values are numeric'
      };
    }

    return {
      type: 'any',
      confidence: 0.3,
      reason: 'Mixed data types detected'
    };
  }
}

export default ValidationService.getInstance();
