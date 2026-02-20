/**
 * Custom Functions Service
 *
 * Enables users to create JavaScript-based custom functions (UDFs) for Excel
 * and supports Excel's LAMBDA function for user-defined calculations.
 *
 * Features:
 * - JavaScript UDF registration
 * - LAMBDA function creation and management
 * - Function validation and testing
 * - Import/export of custom functions
 * - Function library management
 *
 * @module services/customFunctions
 */

import { notificationManager } from '../utils/notificationManager';

// ============================================================================
// Type Definitions
// ============================================================================

/** Custom function definition */
export interface CustomFunction {
  id: string;
  name: string;
  description: string;
  category: string;
  parameters: FunctionParameter[];
  returnType: 'string' | 'number' | 'boolean' | 'array' | 'any';
  javascriptCode: string;
  lambdaExpression?: string;
  examples: string[];
  isLambda: boolean;
  createdAt: Date;
  updatedAt: Date;
  usageCount: number;
}

/** Function parameter */
export interface FunctionParameter {
  name: string;
  description: string;
  type: 'string' | 'number' | 'boolean' | 'array' | 'range' | 'any';
  optional?: boolean;
  defaultValue?: any;
}

/** Function test case */
export interface FunctionTestCase {
  name: string;
  inputs: any[];
  expectedOutput: any;
  description?: string;
}

/** Function validation result */
export interface FunctionValidation {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

/** Validation error */
export interface ValidationError {
  line: number;
  column: number;
  message: string;
  code: string;
}

/** Validation warning */
export interface ValidationWarning {
  line: number;
  message: string;
  suggestion: string;
}

/** Function execution result */
export interface FunctionExecutionResult {
  success: boolean;
  result?: any;
  error?: string;
  executionTime: number;
  logs: string[];
}

/** Function library */
export interface FunctionLibrary {
  id: string;
  name: string;
  description: string;
  functions: CustomFunction[];
  isBuiltIn: boolean;
  isShared: boolean;
}

// ============================================================================
// Custom Functions Service
// ============================================================================

export class CustomFunctionsService {
  private static instance: CustomFunctionsService;
  private functions: Map<string, CustomFunction> = new Map();
  private libraries: Map<string, FunctionLibrary> = new Map();
  private isInitialized: boolean = false;

  private constructor() {
    this.initializeBuiltInFunctions();
  }

  static getInstance(): CustomFunctionsService {
    if (!CustomFunctionsService.instance) {
      CustomFunctionsService.instance = new CustomFunctionsService();
    }
    return CustomFunctionsService.instance;
  }

  // ============================================================================
  // Initialization
  // ============================================================================

  private initializeBuiltInFunctions(): void {
    // Create built-in library
    const builtInLibrary: FunctionLibrary = {
      id: 'builtin',
      name: 'Built-in Functions',
      description: 'Essential custom functions for Excel',
      functions: [],
      isBuiltIn: true,
      isShared: true,
    };

    // Add some useful built-in functions
    const builtInFunctions: Omit<CustomFunction, 'id' | 'createdAt' | 'updatedAt' | 'usageCount'>[] = [
      {
        name: 'REVERSE',
        description: 'Reverses a string or array',
        category: 'Text',
        parameters: [
          { name: 'value', description: 'String or array to reverse', type: 'any' },
        ],
        returnType: 'any',
        javascriptCode: `
          if (typeof value === 'string') {
            return value.split('').reverse().join('');
          } else if (Array.isArray(value)) {
            return value.reverse();
          }
          return value;
        `,
        examples: ['=REVERSE("hello") returns "olleh"', '=REVERSE({1,2,3}) returns {3,2,1}'],
        isLambda: false,
      },
      {
        name: 'COUNTWORDS',
        description: 'Counts the number of words in a text string',
        category: 'Text',
        parameters: [
          { name: 'text', description: 'Text to count words in', type: 'string' },
        ],
        returnType: 'number',
        javascriptCode: `
          if (!text || typeof text !== 'string') return 0;
          return text.trim().split(/\\s+/).filter(word => word.length > 0).length;
        `,
        examples: ['=COUNTWORDS("Hello world") returns 2'],
        isLambda: false,
      },
      {
        name: 'ISPRIME',
        description: 'Checks if a number is prime',
        category: 'Math',
        parameters: [
          { name: 'number', description: 'Number to check', type: 'number' },
        ],
        returnType: 'boolean',
        javascriptCode: `
          if (number < 2) return false;
          if (number === 2) return true;
          if (number % 2 === 0) return false;
          for (let i = 3; i <= Math.sqrt(number); i += 2) {
            if (number % i === 0) return false;
          }
          return true;
        `,
        examples: ['=ISPRIME(7) returns TRUE', '=ISPRIME(4) returns FALSE'],
        isLambda: false,
      },
      {
        name: 'FIBONACCI',
        description: 'Returns the nth Fibonacci number',
        category: 'Math',
        parameters: [
          { name: 'n', description: 'Position in Fibonacci sequence', type: 'number' },
        ],
        returnType: 'number',
        javascriptCode: `
          if (n < 0) return undefined;
          if (n === 0) return 0;
          if (n === 1) return 1;
          let a = 0, b = 1;
          for (let i = 2; i <= n; i++) {
            const temp = a + b;
            a = b;
            b = temp;
          }
          return b;
        `,
        examples: ['=FIBONACCI(10) returns 55'],
        isLambda: false,
      },
      {
        name: 'INTERPOLATE',
        description: 'Linear interpolation between two values',
        category: 'Math',
        parameters: [
          { name: 'x', description: 'Input value', type: 'number' },
          { name: 'x1', description: 'First known x', type: 'number' },
          { name: 'y1', description: 'First known y', type: 'number' },
          { name: 'x2', description: 'Second known x', type: 'number' },
          { name: 'y2', description: 'Second known y', type: 'number' },
        ],
        returnType: 'number',
        javascriptCode: `
          return y1 + (x - x1) * (y2 - y1) / (x2 - x1);
        `,
        examples: ['=INTERPOLATE(5, 0, 0, 10, 100) returns 50'],
        isLambda: false,
      },
    ];

    for (const func of builtInFunctions) {
      const fullFunction: CustomFunction = {
        ...func,
        id: `builtin-${func.name.toLowerCase()}`,
        createdAt: new Date(),
        updatedAt: new Date(),
        usageCount: 0,
      };
      this.functions.set(fullFunction.id, fullFunction);
      builtInLibrary.functions.push(fullFunction);
    }

    this.libraries.set(builtInLibrary.id, builtInLibrary);
    this.isInitialized = true;
  }

  // ============================================================================
  // Function Management
  // ============================================================================

  /**
   * Create a new custom function
   */
  createFunction(
    func: Omit<CustomFunction, 'id' | 'createdAt' | 'updatedAt' | 'usageCount'>
  ): CustomFunction {
    const validation = this.validateFunction(func);
    if (!validation.isValid) {
      throw new Error('Function validation failed: ' + validation.errors.map(e => e.message).join(', '));
    }

    const id = `func-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    const newFunction: CustomFunction = {
      ...func,
      id,
      createdAt: new Date(),
      updatedAt: new Date(),
      usageCount: 0,
    };

    this.functions.set(id, newFunction);
    notificationManager.success(`Function ${func.name} created successfully`);
    return newFunction;
  }

  /**
   * Update an existing function
   */
  updateFunction(id: string, updates: Partial<CustomFunction>): CustomFunction | null {
    const func = this.functions.get(id);
    if (!func) return null;

    const updatedFunction: CustomFunction = {
      ...func,
      ...updates,
      id: func.id,
      updatedAt: new Date(),
    };

    this.functions.set(id, updatedFunction);
    notificationManager.success(`Function ${updatedFunction.name} updated`);
    return updatedFunction;
  }

  /**
   * Delete a function
   */
  deleteFunction(id: string): boolean {
    const func = this.functions.get(id);
    if (!func) return false;

    if (func.isLambda) {
      notificationManager.warning('Cannot delete LAMBDA functions directly');
      return false;
    }

    this.functions.delete(id);
    notificationManager.success(`Function ${func.name} deleted`);
    return true;
  }

  /**
   * Get a function by ID
   */
  getFunction(id: string): CustomFunction | undefined {
    return this.functions.get(id);
  }

  /**
   * Get all functions
   */
  getAllFunctions(): CustomFunction[] {
    return Array.from(this.functions.values());
  }

  /**
   * Get functions by category
   */
  getFunctionsByCategory(category: string): CustomFunction[] {
    return this.getAllFunctions().filter(f => f.category === category);
  }

  /**
   * Search functions
   */
  searchFunctions(query: string): CustomFunction[] {
    const lowerQuery = query.toLowerCase();
    return this.getAllFunctions().filter(
      f =>
        f.name.toLowerCase().includes(lowerQuery) ||
        f.description.toLowerCase().includes(lowerQuery) ||
        f.category.toLowerCase().includes(lowerQuery)
    );
  }

  // ============================================================================
  // Function Validation
  // ============================================================================

  /**
   * Validate a function definition
   */
  validateFunction(
    func: Omit<CustomFunction, 'id' | 'createdAt' | 'updatedAt' | 'usageCount'>
  ): FunctionValidation {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    // Validate name
    if (!func.name || func.name.length === 0) {
      errors.push({ line: 0, column: 0, message: 'Function name is required', code: 'MISSING_NAME' });
    } else if (!/^[A-Z][A-Z0-9_]*$/.test(func.name)) {
      errors.push({
        line: 0,
        column: 0,
        message: 'Function name must start with uppercase letter and contain only letters, numbers, and underscores',
        code: 'INVALID_NAME',
      });
    }

    // Check for duplicate names
    const existing = Array.from(this.functions.values()).find(f => f.name === func.name);
    if (existing) {
      errors.push({ line: 0, column: 0, message: `Function ${func.name} already exists`, code: 'DUPLICATE_NAME' });
    }

    // Validate JavaScript code
    if (!func.isLambda && func.javascriptCode) {
      try {
        // Try to parse as function
        new Function(...func.parameters.map(p => p.name), func.javascriptCode);
      } catch (e) {
        errors.push({
          line: 0,
          column: 0,
          message: 'JavaScript code has syntax errors: ' + e,
          code: 'JS_SYNTAX_ERROR',
        });
      }

      // Check for dangerous patterns
      const dangerousPatterns = ['eval(', 'Function(', 'setTimeout', 'setInterval', 'fetch(', 'XMLHttpRequest'];
      for (const pattern of dangerousPatterns) {
        if (func.javascriptCode.includes(pattern)) {
          warnings.push({
            line: 0,
            message: `Potentially dangerous pattern detected: ${pattern}`,
            suggestion: 'Avoid using eval, Function constructor, or network requests in custom functions',
          });
        }
      }
    }

    // Validate LAMBDA expression
    if (func.isLambda && func.lambdaExpression) {
      if (!func.lambdaExpression.includes('LAMBDA')) {
        warnings.push({
          line: 0,
          message: 'LAMBDA expression should use LAMBDA function',
          suggestion: 'Use =LAMBDA(parameters, calculation) syntax',
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  // ============================================================================
  // Function Execution
  // ============================================================================

  /**
   * Execute a custom function
   */
  async executeFunction(
    functionId: string,
    inputs: any[]
  ): Promise<FunctionExecutionResult> {
    const func = this.functions.get(functionId);
    if (!func) {
      return {
        success: false,
        error: 'Function not found',
        executionTime: 0,
        logs: [],
      };
    }

    const startTime = Date.now();
    const logs: string[] = [];

    try {
      let result: any;

      if (func.isLambda && func.lambdaExpression) {
        // Execute LAMBDA expression
        result = await this.executeLambda(func.lambdaExpression, inputs, func.parameters);
      } else {
        // Execute JavaScript code
        const jsFunction = new Function(
          ...func.parameters.map(p => p.name),
          `
            "use strict";
            ${func.javascriptCode}
          `
        );
        result = jsFunction(...inputs);
      }

      // Update usage count
      func.usageCount++;
      func.updatedAt = new Date();

      return {
        success: true,
        result,
        executionTime: Date.now() - startTime,
        logs,
      };
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : String(error),
        executionTime: Date.now() - startTime,
        logs,
      };
    }
  }

  /**
   * Execute a LAMBDA expression
   */
  private async executeLambda(
    expression: string,
    inputs: any[],
    parameters: FunctionParameter[]
  ): Promise<any> {
    // Parse LAMBDA expression
    // Simplified implementation - real one would use Excel's LAMBDA evaluation
    const lambdaMatch = expression.match(/LAMBDA\s*\(([^)]+)\)\s*,\s*(.+)/i);
    if (!lambdaMatch) {
      throw new Error('Invalid LAMBDA expression');
    }

    const paramNames = lambdaMatch[1].split(',').map(p => p.trim());
    const calculation = lambdaMatch[2];

    // Create evaluation context
    const context: Record<string, any> = {};
    for (let i = 0; i < paramNames.length && i < inputs.length; i++) {
      context[paramNames[i]] = inputs[i];
    }

    // Basic expression evaluation
    // Note: This is a simplified implementation. For full Excel formula support,
    // you would need to use Excel's formula evaluation engine or a formula parser library
    try {
      // Replace parameter names with their values
      let evalExpression = calculation;
      for (const [paramName, value] of Object.entries(context)) {
        const regex = new RegExp(`\\b${paramName}\\b`, 'g');
        evalExpression = evalExpression.replace(regex, String(value));
      }
      
      // Basic arithmetic evaluation (safe evaluation)
      // Only allow basic math operations for security
      const sanitized = evalExpression.replace(/[^0-9+\-*/().\s]/g, '');
      if (sanitized !== evalExpression.replace(/\s/g, '')) {
        // Contains non-numeric operations, use Excel's evaluation
        // For now, return a message indicating Excel evaluation is needed
        throw new Error('Complex formula evaluation requires Excel engine');
      }
      
      // Use Function constructor for safe evaluation (limited to arithmetic)
      const result = new Function(`return ${evalExpression}`)();
      return result;
    } catch (error) {
      // If evaluation fails, indicate that Excel's formula engine should be used
      // In a production environment, this would call Excel's LAMBDA evaluation API
      throw new Error(
        `LAMBDA evaluation requires Excel's formula engine. ` +
        `Expression: ${calculation}, Parameters: ${JSON.stringify(context)}`
      );
    }
  }

  // ============================================================================
  // LAMBDA Functions
  // ============================================================================

  /**
   * Create a LAMBDA function from Excel expression
   */
  createLambdaFunction(
    name: string,
    lambdaExpression: string,
    description: string,
    category: string = 'Custom'
  ): CustomFunction {
    // Extract parameters from LAMBDA expression
    const paramMatch = lambdaExpression.match(/LAMBDA\s*\(([^)]+)\)/i);
    const paramNames = paramMatch ? paramMatch[1].split(',').map(p => p.trim()) : [];

    const parameters: FunctionParameter[] = paramNames.map(name => ({
      name,
      description: `Parameter ${name}`,
      type: 'any',
    }));

    return this.createFunction({
      name,
      description,
      category,
      parameters,
      returnType: 'any',
      javascriptCode: '',
      lambdaExpression,
      examples: [`=${name}(${paramNames.join(', ')})`],
      isLambda: true,
    });
  }

  /**
   * Convert JavaScript function to LAMBDA
   */
  convertToLambda(func: CustomFunction): string | null {
    if (func.isLambda) return func.lambdaExpression || null;

    // This is a simplified conversion
    // Real implementation would analyze JavaScript AST and convert to Excel formula
    const params = func.parameters.map(p => p.name).join(', ');
    return `=LAMBDA(${params}, /* JavaScript: ${func.name} */)`;
  }

  // ============================================================================
  // Testing
  // ============================================================================

  /**
   * Run test cases for a function
   */
  async runTests(
    functionId: string,
    testCases: FunctionTestCase[]
  ): Promise<{ passed: number; failed: number; results: any[] }> {
    const func = this.functions.get(functionId);
    if (!func) {
      throw new Error('Function not found');
    }

    let passed = 0;
    let failed = 0;
    const results: any[] = [];

    for (const testCase of testCases) {
      const result = await this.executeFunction(functionId, testCase.inputs);

      const testResult = {
        name: testCase.name,
        passed: false,
        expected: testCase.expectedOutput,
        actual: result.result,
        error: result.error,
      };

      if (result.success && JSON.stringify(result.result) === JSON.stringify(testCase.expectedOutput)) {
        testResult.passed = true;
        passed++;
      } else {
        failed++;
      }

      results.push(testResult);
    }

    return { passed, failed, results };
  }

  // ============================================================================
  // Import/Export
  // ============================================================================

  /**
   * Export function to JSON
   */
  exportFunction(functionId: string): string | null {
    const func = this.functions.get(functionId);
    if (!func) return null;

    return JSON.stringify(func, null, 2);
  }

  /**
   * Import function from JSON
   */
  importFunction(json: string): CustomFunction | null {
    try {
      const funcData = JSON.parse(json);
      return this.createFunction({
        name: funcData.name,
        description: funcData.description,
        category: funcData.category,
        parameters: funcData.parameters,
        returnType: funcData.returnType,
        javascriptCode: funcData.javascriptCode,
        lambdaExpression: funcData.lambdaExpression,
        examples: funcData.examples,
        isLambda: funcData.isLambda,
      });
    } catch (error) {
      notificationManager.error('Failed to import function: ' + error);
      return null;
    }
  }

  /**
   * Export all functions
   */
  exportAllFunctions(): string {
    const exportData = {
      version: '1.0',
      exportedAt: new Date().toISOString(),
      functions: Array.from(this.functions.values()).filter(f => !f.id.startsWith('builtin-')),
    };
    return JSON.stringify(exportData, null, 2);
  }

  /**
   * Import multiple functions
   */
  importFunctions(json: string): number {
    try {
      const data = JSON.parse(json);
      let count = 0;

      for (const funcData of data.functions || []) {
        const func = this.createFunction({
          name: funcData.name,
          description: funcData.description,
          category: funcData.category,
          parameters: funcData.parameters,
          returnType: funcData.returnType,
          javascriptCode: funcData.javascriptCode,
          lambdaExpression: funcData.lambdaExpression,
          examples: funcData.examples,
          isLambda: funcData.isLambda,
        });
        if (func) count++;
      }

      notificationManager.success(`Imported ${count} functions`);
      return count;
    } catch (error) {
      notificationManager.error('Failed to import functions: ' + error);
      return 0;
    }
  }

  // ============================================================================
  // Excel Integration
  // ============================================================================

  /**
   * Register function with Excel (requires Script Lab or similar)
   */
  async registerWithExcel(functionId: string): Promise<boolean> {
    const func = this.functions.get(functionId);
    if (!func) return false;

    // This would require Office.js custom functions API
    // For now, just log that it would be registered
    notificationManager.info(`Function ${func.name} ready for Excel registration`);
    return true;
  }

  /**
   * Get Excel formula for function
   */
  getExcelFormula(functionId: string): string | null {
    const func = this.functions.get(functionId);
    if (!func) return null;

    if (func.isLambda && func.lambdaExpression) {
      return func.lambdaExpression;
    }

    // For JavaScript functions, return usage example
    const params = func.parameters.map(p => p.name).join(', ');
    return `=${func.name}(${params})`;
  }
}

// Export singleton instance
export const customFunctions = CustomFunctionsService.getInstance();
export default customFunctions;
