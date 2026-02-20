// Formula Parser Service
// Parses Excel formulas into AST and generates human-readable explanations

import {
  Token,
  TokenType,
  tokenizeFormula,
  getFunctionInfo,
  isExcelFunction
} from '../utils/formulaTokenizer';
import { localeService } from '../i18n';

// AST Node Types
export type ASTNodeType =
  | 'Function'
  | 'Reference'
  | 'Range'
  | 'Number'
  | 'String'
  | 'BinaryOp'
  | 'UnaryOp'
  | 'Name'
  | 'Error';

export interface ASTNode {
  type: ASTNodeType;
  value?: string;
  name?: string;
  arguments?: ASTNode[];
  left?: ASTNode;
  right?: ASTNode;
  operator?: string;
  operands?: ASTNode[];
}

export interface FormulaExplanation {
  summary: string;
  breakdown: ExplanationStep[];
  dependencies: string[];
  complexity: 'Simple' | 'Moderate' | 'Complex';
  optimizations?: string[];
}

export interface ExplanationStep {
  expression: string;
  description: string;
  result?: string;
  depth: number;
}

export interface DependencyTrace {
  cell: string;
  sheet?: string;
  type: 'direct' | 'indirect' | 'range';
}

export class FormulaParser {
  private tokens: Token[];
  private position: number;

  constructor(formula: string) {
    this.tokens = tokenizeFormula(formula);
    this.position = 0;
  }

  // Get current token
  private current(): Token {
    return this.tokens[this.position] || { type: 'EOF', value: '', position: 0 };
  }

  // Advance to next token
  private advance(): void {
    if (this.position < this.tokens.length - 1) {
      this.position++;
    }
  }

  // Check if current token matches type
  private match(...types: TokenType[]): boolean {
    return types.includes(this.current().type);
  }

  // Expect a specific token type
  private expect(type: TokenType): Token {
    const token = this.current();
    if (token.type !== type) {
      throw new Error(`Expected ${type}, got ${token.type}`);
    }
    this.advance();
    return token;
  }

  // Parse the entire formula
  parse(): ASTNode {
    return this.parseExpression();
  }

  // Parse expression (handles operators)
  private parseExpression(): ASTNode {
    return this.parseComparison();
  }

  // Parse comparison operators
  private parseComparison(): ASTNode {
    let left = this.parseConcatenation();

    while (this.match('OPERATOR') && /^(=|<>|<|>|<=|>=)$/.test(this.current().value)) {
      const operator = this.current().value;
      this.advance();
      const right = this.parseConcatenation();
      left = {
        type: 'BinaryOp',
        operator,
        left,
        right
      };
    }

    return left;
  }

  // Parse concatenation (&)
  private parseConcatenation(): ASTNode {
    let left = this.parseAddition();

    while (this.match('OPERATOR') && this.current().value === '&') {
      this.advance();
      const right = this.parseAddition();
      left = {
        type: 'BinaryOp',
        operator: '&',
        left,
        right
      };
    }

    return left;
  }

  // Parse addition/subtraction
  private parseAddition(): ASTNode {
    let left = this.parseMultiplication();

    while (this.match('OPERATOR') && /^[+-]$/.test(this.current().value)) {
      const operator = this.current().value;
      this.advance();
      const right = this.parseMultiplication();
      left = {
        type: 'BinaryOp',
        operator,
        left,
        right
      };
    }

    return left;
  }

  // Parse multiplication/division
  private parseMultiplication(): ASTNode {
    let left = this.parsePower();

    while (this.match('OPERATOR') && /^[*/]$/.test(this.current().value)) {
      const operator = this.current().value;
      this.advance();
      const right = this.parsePower();
      left = {
        type: 'BinaryOp',
        operator,
        left,
        right
      };
    }

    return left;
  }

  // Parse power (^)
  private parsePower(): ASTNode {
    let left = this.parseUnary();

    if (this.match('OPERATOR') && this.current().value === '^') {
      this.advance();
      const right = this.parsePower(); // Right associative
      left = {
        type: 'BinaryOp',
        operator: '^',
        left,
        right
      };
    }

    return left;
  }

  // Parse unary operators (+, -)
  private parseUnary(): ASTNode {
    if (this.match('OPERATOR') && /^[+-]$/.test(this.current().value)) {
      const operator = this.current().value;
      this.advance();
      const operand = this.parseUnary();
      return {
        type: 'UnaryOp',
        operator,
        operands: [operand]
      };
    }

    return this.parsePrimary();
  }

  // Parse primary expressions
  private parsePrimary(): ASTNode {
    const token = this.current();

    switch (token.type) {
      case 'NUMBER':
        this.advance();
        return { type: 'Number', value: token.value };

      case 'STRING':
        this.advance();
        return { type: 'String', value: token.value };

      case 'REFERENCE':
        this.advance();
        return {
          type: token.value.includes(':') ? 'Range' : 'Reference',
          value: token.value
        };

      case 'RANGE':
        this.advance();
        return { type: 'Range', value: token.value };

      case 'FUNCTION':
        return this.parseFunction();

      case 'NAME':
        this.advance();
        return { type: 'Name', value: token.value };

      case 'LPAREN':
        this.advance();
        const expr = this.parseExpression();
        this.expect('RPAREN');
        return expr;

      default:
        this.advance();
        return { type: 'Error', value: `Unexpected token: ${token.value}` };
    }
  }

  // Parse function calls
  private parseFunction(): ASTNode {
    const name = this.current().value;
    this.advance(); // Consume function name
    this.expect('LPAREN');

    const args: ASTNode[] = [];

    if (!this.match('RPAREN')) {
      args.push(this.parseExpression());

      while (this.match('COMMA', 'SEMICOLON')) {
        this.advance();
        args.push(this.parseExpression());
      }
    }

    this.expect('RPAREN');

    return {
      type: 'Function',
      name,
      arguments: args
    };
  }

  // Generate explanation from AST
  explainFormula(ast: ASTNode): FormulaExplanation {
    const breakdown = this.generateBreakdown(ast);
    const dependencies = this.extractDependencies(ast);
    const complexity = this.calculateComplexity(ast);
    const optimizations = this.suggestOptimizations(ast);

    return {
      summary: this.generateSummary(ast),
      breakdown,
      dependencies,
      complexity,
      optimizations
    };
  }

  // Generate step-by-step breakdown
  private generateBreakdown(node: ASTNode, depth: number = 0): ExplanationStep[] {
    const steps: ExplanationStep[] = [];

    switch (node.type) {
      case 'Function':
        steps.push(this.explainFunction(node, depth));
        if (node.arguments) {
          for (const arg of node.arguments) {
            steps.push(...this.generateBreakdown(arg, depth + 1));
          }
        }
        break;

      case 'BinaryOp':
        steps.push(this.explainBinaryOp(node, depth));
        if (node.left) {
          steps.push(...this.generateBreakdown(node.left, depth + 1));
        }
        if (node.right) {
          steps.push(...this.generateBreakdown(node.right, depth + 1));
        }
        break;

      case 'UnaryOp':
        steps.push(this.explainUnaryOp(node, depth));
        if (node.operands) {
          for (const operand of node.operands) {
            steps.push(...this.generateBreakdown(operand, depth + 1));
          }
        }
        break;

      case 'Reference':
        steps.push({
          expression: node.value!,
          description: `References the value in cell ${node.value}`,
          depth
        });
        break;

      case 'Range':
        steps.push({
          expression: node.value!,
          description: `References the range of cells from ${node.value?.split(':')[0]} to ${node.value?.split(':')[1]}`,
          depth
        });
        break;

      case 'Number':
        steps.push({
          expression: node.value!,
          description: `Uses the constant value ${node.value}`,
          depth
        });
        break;

      case 'String':
        steps.push({
          expression: `"${node.value}"`,
          description: `Uses the text "${node.value}"`,
          depth
        });
        break;
    }

    return steps;
  }

  // Explain function calls
  private explainFunction(node: ASTNode, depth: number): ExplanationStep {
    const funcName = node.name!.toUpperCase();
    const funcInfo = getFunctionInfo(funcName);
    const argCount = node.arguments?.length || 0;

    let description: string;

    switch (funcName) {
      case 'SUM':
        description = `Calculates the sum of ${argCount} value(s)`;
        break;
      case 'AVERAGE':
        description = `Calculates the average of ${argCount} value(s)`;
        break;
      case 'IF':
        description = `Evaluates a condition and returns one value if true, another if false`;
        break;
      case 'VLOOKUP':
        description = `Looks up a value in the first column of a table and returns a value from another column`;
        break;
      case 'SUMIF':
        description = `Adds values that meet a specific condition`;
        break;
      case 'COUNT':
        description = `Counts the number of values`;
        break;
      case 'CONCAT':
      case 'CONCATENATE':
        description = `Joins ${argCount} text string(s) together`;
        break;
      case 'TODAY':
        description = `Returns the current date`;
        break;
      case 'NOW':
        description = `Returns the current date and time`;
        break;
      default:
        if (funcInfo) {
          description = funcInfo.description;
        } else {
          description = `Calls the ${funcName} function with ${argCount} argument(s)`;
        }
    }

    return {
      expression: `${funcName}()`,
      description,
      depth
    };
  }

  // Explain binary operations
  private explainBinaryOp(node: ASTNode, depth: number): ExplanationStep {
    const operatorDescriptions: Record<string, string> = {
      '+': 'Adds two values together',
      '-': 'Subtracts the second value from the first',
      '*': 'Multiplies two values',
      '/': 'Divides the first value by the second',
      '^': 'Raises the first value to the power of the second',
      '&': 'Joins two text strings together',
      '=': 'Checks if two values are equal',
      '<>': 'Checks if two values are not equal',
      '<': 'Checks if the first value is less than the second',
      '>': 'Checks if the first value is greater than the second',
      '<=': 'Checks if the first value is less than or equal to the second',
      '>=': 'Checks if the first value is greater than or equal to the second'
    };

    return {
      expression: `... ${node.operator} ...`,
      description: operatorDescriptions[node.operator!] || `Performs ${node.operator} operation`,
      depth
    };
  }

  // Explain unary operations
  private explainUnaryOp(node: ASTNode, depth: number): ExplanationStep {
    const descriptions: Record<string, string> = {
      '+': 'Returns the positive value',
      '-': 'Negates the value (makes it negative)'
    };

    return {
      expression: `${node.operator}...`,
      description: descriptions[node.operator!] || `Applies ${node.operator} operator`,
      depth
    };
  }

  // Generate high-level summary
  private generateSummary(node: ASTNode): string {
    switch (node.type) {
      case 'Function':
        const funcName = node.name!.toUpperCase();
        const funcInfo = getFunctionInfo(funcName);
        return funcInfo?.description || `Uses the ${funcName} function`;

      case 'BinaryOp':
        if (node.operator === '&') {
          return 'Joins text strings together';
        } else if (/^[+-*/^]$/.test(node.operator!)) {
          return 'Performs a mathematical calculation';
        } else {
          return 'Compares values';
        }

      case 'Reference':
        return `Returns the value from cell ${node.value}`;

      case 'Range':
        return `References a range of cells`;

      case 'Number':
        return `Returns the constant value ${node.value}`;

      case 'String':
        return `Returns the text "${node.value}"`;

      default:
        return 'Performs a calculation';
    }
  }

  // Extract all cell/range dependencies
  private extractDependencies(node: ASTNode, deps: Set<string> = new Set()): string[] {
    switch (node.type) {
      case 'Reference':
      case 'Range':
        deps.add(node.value!);
        break;

      case 'Function':
        if (node.arguments) {
          for (const arg of node.arguments) {
            this.extractDependencies(arg, deps);
          }
        }
        break;

      case 'BinaryOp':
        if (node.left) this.extractDependencies(node.left, deps);
        if (node.right) this.extractDependencies(node.right, deps);
        break;

      case 'UnaryOp':
        if (node.operands) {
          for (const operand of node.operands) {
            this.extractDependencies(operand, deps);
          }
        }
        break;
    }

    return Array.from(deps);
  }

  // Calculate formula complexity
  private calculateComplexity(node: ASTNode): 'Simple' | 'Moderate' | 'Complex' {
    let score = 0;

    const countNodes = (n: ASTNode): number => {
      let count = 1;

      if (n.type === 'Function') {
        score += 2;
        if (n.arguments) {
          for (const arg of n.arguments) {
            count += countNodes(arg);
          }
        }
      } else if (n.type === 'BinaryOp') {
        score += 1;
        if (n.left) count += countNodes(n.left);
        if (n.right) count += countNodes(n.right);
      }

      return count;
    };

    const totalNodes = countNodes(node);

    if (totalNodes <= 3 && score <= 2) return 'Simple';
    if (totalNodes <= 10 && score <= 8) return 'Moderate';
    return 'Complex';
  }

  // Suggest formula optimizations
  private suggestOptimizations(node: ASTNode): string[] {
    const suggestions: string[] = [];

    // Check for nested IFs that could be simplified
    if (node.type === 'Function' && node.name?.toUpperCase() === 'IF') {
      if (node.arguments && node.arguments[2]?.type === 'Function') {
        const elseBranch = node.arguments[2];
        if (elseBranch.name?.toUpperCase() === 'IF') {
          suggestions.push('Consider using IFS() instead of nested IF statements for better readability');
        }
      }
    }

    // Check for CONCATENATE that could use CONCAT or &
    if (node.type === 'Function' && node.name?.toUpperCase() === 'CONCATENATE') {
      suggestions.push('Consider using CONCAT() or the & operator instead of CONCATENATE (legacy function)');
    }

    // Check for VLOOKUP that could use XLOOKUP
    if (node.type === 'Function' && node.name?.toUpperCase() === 'VLOOKUP') {
      suggestions.push('Consider using XLOOKUP() for more flexibility (available in Excel 365/2021+)');
    }

    // Check for SUM/COUNT with IF that could use SUMIF/COUNTIF
    if (node.type === 'Function' && ['SUM', 'COUNT'].includes(node.name?.toUpperCase() || '')) {
      if (node.arguments?.some(arg => arg.type === 'Function' && arg.name?.toUpperCase() === 'IF')) {
        suggestions.push(`Consider using ${node.name}IF() for conditional aggregation`);
      }
    }

    // Check for volatile functions used excessively
    const volatileFunctions = ['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'OFFSET', 'INDIRECT'];
    if (node.type === 'Function' && volatileFunctions.includes(node.name?.toUpperCase() || '')) {
      suggestions.push(`${node.name}() is a volatile function that recalculates on every change. Use sparingly in large workbooks.`);
    }

    return suggestions;
  }
}

// Main export functions
export function parseFormula(formula: string): ASTNode {
  const parser = new FormulaParser(formula);
  return parser.parse();
}

export function explainFormula(formula: string): FormulaExplanation {
  const parser = new FormulaParser(formula);
  const ast = parser.parse();
  return parser.explainFormula(ast);
}

export function traceDependencies(formula: string): string[] {
  const parser = new FormulaParser(formula);
  const ast = parser.parse();
  return parser['extractDependencies'](ast);
}

// Helper to get formula complexity
export function getFormulaComplexity(formula: string): 'Simple' | 'Moderate' | 'Complex' {
  const parser = new FormulaParser(formula);
  const ast = parser.parse();
  return parser['calculateComplexity'](ast);
}

// Batch explain multiple formulas
export function explainFormulas(formulas: string[]): Map<string, FormulaExplanation> {
  const results = new Map<string, FormulaExplanation>();

  for (const formula of formulas) {
    try {
      results.set(formula, explainFormula(formula));
    } catch (error: any) {
      results.set(formula, {
        summary: 'Error parsing formula',
        breakdown: [{ expression: formula, description: `Parse error: ${error?.message || 'Unknown error'}`, depth: 0 }],
        dependencies: [],
        complexity: 'Simple',
        optimizations: []
      });
    }
  }

  return results;
}

// ==================== LOCALIZED FORMULA SUPPORT ====================

/**
 * Parse a localized formula (converts to English first)
 */
export function parseLocalizedFormula(formula: string): ASTNode {
  const englishFormula = localeService.delocalizeFormula(formula);
  return parseFormula(englishFormula);
}

/**
 * Explain a localized formula
 */
export function explainLocalizedFormula(formula: string): FormulaExplanation {
  const englishFormula = localeService.delocalizeFormula(formula);
  return explainFormula(englishFormula);
}

/**
 * Get localized function name for display
 */
export function getLocalizedFunctionNameForDisplay(englishName: string): string {
  return localeService.getLocalizedFunction(englishName);
}

/**
 * Convert English formula to localized version for display
 */
export function localizeFormulaForDisplay(formula: string): string {
  return localeService.localizeFormula(formula);
}

/**
 * Convert localized formula to English for processing
 */
export function delocalizeFormulaForProcessing(formula: string): string {
  return localeService.delocalizeFormula(formula);
}
