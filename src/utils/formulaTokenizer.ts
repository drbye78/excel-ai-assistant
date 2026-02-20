// Excel Formula Tokenizer
// Parses Excel formulas into tokens for analysis and explanation

export type TokenType =
  | 'FUNCTION'
  | 'OPERATOR'
  | 'NUMBER'
  | 'STRING'
  | 'REFERENCE'
  | 'RANGE'
  | 'NAME'
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'SEMICOLON'
  | 'LBRACKET'
  | 'RBRACKET'
  | 'ERROR'
  | 'EOF';

export interface Token {
  type: TokenType;
  value: string;
  position: number;
}

export class FormulaTokenizer {
  private formula: string;
  private position: number;
  private currentChar: string | null;

  constructor(formula: string) {
    this.formula = formula.startsWith('=') ? formula.slice(1) : formula;
    this.position = 0;
    this.currentChar = this.formula[0] || null;
  }

  private advance(): void {
    this.position++;
    this.currentChar = this.position < this.formula.length
      ? this.formula[this.position]
      : null;
  }

  private peek(offset: number = 1): string | null {
    const pos = this.position + offset;
    return pos < this.formula.length ? this.formula[pos] : null;
  }

  private skipWhitespace(): void {
    while (this.currentChar && /\s/.test(this.currentChar)) {
      this.advance();
    }
  }

  private readString(): Token {
    const startPos = this.position;
    const quote = this.currentChar;
    let value = '';
    this.advance(); // Skip opening quote

    while (this.currentChar && this.currentChar !== quote) {
      if (this.currentChar === '\\' && this.peek() === quote) {
        this.advance();
      }
      value += this.currentChar;
      this.advance();
    }

    if (this.currentChar === quote) {
      this.advance(); // Skip closing quote
    }

    return {
      type: 'STRING',
      value,
      position: startPos
    };
  }

  private readNumber(): Token {
    const startPos = this.position;
    let value = '';

    while (this.currentChar && (/\d/.test(this.currentChar) || this.currentChar === '.')) {
      value += this.currentChar;
      this.advance();
    }

    // Check for scientific notation
    if (this.currentChar && (this.currentChar === 'e' || this.currentChar === 'E')) {
      value += this.currentChar;
      this.advance();
      if (this.currentChar && (this.currentChar === '+' || this.currentChar === '-')) {
        value += this.currentChar;
        this.advance();
      }
      while (this.currentChar && /\d/.test(this.currentChar)) {
        value += this.currentChar;
        this.advance();
      }
    }

    return {
      type: 'NUMBER',
      value,
      position: startPos
    };
  }

  private readNameOrFunction(): Token {
    const startPos = this.position;
    let value = '';

    while (this.currentChar && (/[a-zA-Z0-9_.]/.test(this.currentChar) || this.currentChar === '?')) {
      value += this.currentChar;
      this.advance();
    }

    // Check if it's a function (followed by opening parenthesis)
    this.skipWhitespace();
    const isFunction = this.currentChar === '(';

    return {
      type: isFunction ? 'FUNCTION' : 'NAME',
      value,
      position: startPos
    };
  }

  private readReference(): Token {
    const startPos = this.position;
    let value = '';

    // Handle $A$1 style references
    if (this.currentChar === '$') {
      value += this.currentChar;
      this.advance();
    }

    // Column letters
    while (this.currentChar && /[a-zA-Z]/.test(this.currentChar)) {
      value += this.currentChar;
      this.advance();
    }

    // Row numbers
    while (this.currentChar && /\d/.test(this.currentChar)) {
      value += this.currentChar;
      this.advance();
    }

    return {
      type: 'REFERENCE',
      value,
      position: startPos
    };
  }

  private readStructuredReference(): Token {
    const startPos = this.position;
    let value = '';

    this.advance(); // Skip opening bracket

    while (this.currentChar && this.currentChar !== ']') {
      value += this.currentChar;
      this.advance();
    }

    if (this.currentChar === ']') {
      this.advance(); // Skip closing bracket
    }

    return {
      type: 'REFERENCE',
      value: `[${value}]`,
      position: startPos
    };
  }

  tokenize(): Token[] {
    const tokens: Token[] = [];

    while (this.currentChar !== null) {
      this.skipWhitespace();

      if (this.currentChar === null) break;

      const startPos = this.position;

      // String literals
      if (this.currentChar === '"' || this.currentChar === "'") {
        tokens.push(this.readString());
        continue;
      }

      // Numbers
      if (/\d/.test(this.currentChar)) {
        tokens.push(this.readNumber());
        continue;
      }

      // Operators
      if (/[+\-*/^&=<>]/.test(this.currentChar)) {
        let value = this.currentChar;
        this.advance();

        // Handle multi-character operators
        if (this.currentChar === '=' || this.currentChar === '>') {
          value += this.currentChar;
          this.advance();
        }

        tokens.push({
          type: 'OPERATOR',
          value,
          position: startPos
        });
        continue;
      }

      // Parentheses
      if (this.currentChar === '(') {
        tokens.push({ type: 'LPAREN', value: '(', position: startPos });
        this.advance();
        continue;
      }

      if (this.currentChar === ')') {
        tokens.push({ type: 'RPAREN', value: ')', position: startPos });
        this.advance();
        continue;
      }

      // Comma
      if (this.currentChar === ',') {
        tokens.push({ type: 'COMMA', value: ',', position: startPos });
        this.advance();
        continue;
      }

      // Semicolon (used in some locales)
      if (this.currentChar === ';') {
        tokens.push({ type: 'SEMICOLON', value: ';', position: startPos });
        this.advance();
        continue;
      }

      // Brackets (structured references)
      if (this.currentChar === '[') {
        tokens.push(this.readStructuredReference());
        continue;
      }

      // Cell references (like A1, $A$1, Sheet1!A1)
      if (/[a-zA-Z$]/.test(this.currentChar)) {
        const ref = this.readReference();

        // Check for sheet reference (Sheet1!A1)
        if (this.currentChar === '!') {
          this.advance();
          const sheetRef = this.readReference();
          ref.value += '!' + sheetRef.value;
        }

        // Check for range (A1:B10)
        if (this.currentChar === ':') {
          this.advance();
          const endRef = this.readReference();
          ref.value += ':' + endRef.value;
          ref.type = 'RANGE';
        }

        tokens.push(ref);
        continue;
      }

      // Names and functions
      if (/[a-zA-Z_]/.test(this.currentChar)) {
        tokens.push(this.readNameOrFunction());
        continue;
      }

      // Unknown character
      tokens.push({
        type: 'ERROR',
        value: this.currentChar,
        position: startPos
      });
      this.advance();
    }

    tokens.push({ type: 'EOF', value: '', position: this.position });
    return tokens;
  }
}

// Helper function for easy tokenization
export function tokenizeFormula(formula: string): Token[] {
  const tokenizer = new FormulaTokenizer(formula);
  return tokenizer.tokenize();
}

// Excel functions list for validation and categorization
export const EXCEL_FUNCTIONS: Record<string, { category: string; description: string }> = {
  // Math & Trig
  SUM: { category: 'Math', description: 'Adds all numbers in a range' },
  AVERAGE: { category: 'Math', description: 'Calculates the average of numbers' },
  MAX: { category: 'Math', description: 'Returns the maximum value' },
  MIN: { category: 'Math', description: 'Returns the minimum value' },
  COUNT: { category: 'Math', description: 'Counts numbers in a range' },
  COUNTA: { category: 'Math', description: 'Counts non-empty cells' },
  PRODUCT: { category: 'Math', description: 'Multiplies all numbers' },
  POWER: { category: 'Math', description: 'Raises a number to a power' },
  SQRT: { category: 'Math', description: 'Returns the square root' },
  ROUND: { category: 'Math', description: 'Rounds a number' },
  ROUNDUP: { category: 'Math', description: 'Rounds up' },
  ROUNDDOWN: { category: 'Math', description: 'Rounds down' },
  ABS: { category: 'Math', description: 'Returns absolute value' },
  MOD: { category: 'Math', description: 'Returns remainder' },

  // Statistical
  STDEV: { category: 'Statistical', description: 'Standard deviation' },
  STDEVP: { category: 'Statistical', description: 'Population standard deviation' },
  VAR: { category: 'Statistical', description: 'Variance' },
  VARP: { category: 'Statistical', description: 'Population variance' },
  MEDIAN: { category: 'Statistical', description: 'Median value' },
  MODE: { category: 'Statistical', description: 'Most frequent value' },

  // Logical
  IF: { category: 'Logical', description: 'Conditional logic' },
  AND: { category: 'Logical', description: 'All conditions must be true' },
  OR: { category: 'Logical', description: 'At least one condition true' },
  NOT: { category: 'Logical', description: 'Negates a condition' },
  IFERROR: { category: 'Logical', description: 'Handles errors' },
  IFNA: { category: 'Logical', description: 'Handles #N/A errors' },

  // Lookup & Reference
  VLOOKUP: { category: 'Lookup', description: 'Vertical lookup' },
  HLOOKUP: { category: 'Lookup', description: 'Horizontal lookup' },
  INDEX: { category: 'Lookup', description: 'Returns value at index' },
  MATCH: { category: 'Lookup', description: 'Finds position of value' },
  XLOOKUP: { category: 'Lookup', description: 'Modern lookup function' },
  OFFSET: { category: 'Lookup', description: 'Returns offset range' },
  INDIRECT: { category: 'Lookup', description: 'Returns reference from text' },

  // Text
  CONCAT: { category: 'Text', description: 'Joins text strings' },
  CONCATENATE: { category: 'Text', description: 'Joins text (legacy)' },
  LEFT: { category: 'Text', description: 'Extracts left characters' },
  RIGHT: { category: 'Text', description: 'Extracts right characters' },
  MID: { category: 'Text', description: 'Extracts middle characters' },
  LEN: { category: 'Text', description: 'Returns text length' },
  TRIM: { category: 'Text', description: 'Removes extra spaces' },
  UPPER: { category: 'Text', description: 'Converts to uppercase' },
  LOWER: { category: 'Text', description: 'Converts to lowercase' },
  PROPER: { category: 'Text', description: 'Capitalizes words' },
  FIND: { category: 'Text', description: 'Finds text position' },
  SUBSTITUTE: { category: 'Text', description: 'Replaces text' },
  REPLACE: { category: 'Text', description: 'Replaces characters' },

  // Date & Time
  TODAY: { category: 'Date', description: 'Current date' },
  NOW: { category: 'Date', description: 'Current date and time' },
  DATE: { category: 'Date', description: 'Creates date from parts' },
  YEAR: { category: 'Date', description: 'Extracts year' },
  MONTH: { category: 'Date', description: 'Extracts month' },
  DAY: { category: 'Date', description: 'Extracts day' },
  HOUR: { category: 'Date', description: 'Extracts hour' },
  MINUTE: { category: 'Date', description: 'Extracts minute' },
  SECOND: { category: 'Date', description: 'Extracts second' },
  WEEKDAY: { category: 'Date', description: 'Day of week' },
  DATEDIF: { category: 'Date', description: 'Difference between dates' },
  EOMONTH: { category: 'Date', description: 'End of month' },
  EDATE: { category: 'Date', description: 'Adds months to date' },

  // Financial
  PMT: { category: 'Financial', description: 'Loan payment' },
  PV: { category: 'Financial', description: 'Present value' },
  FV: { category: 'Financial', description: 'Future value' },
  NPV: { category: 'Financial', description: 'Net present value' },
  IRR: { category: 'Financial', description: 'Internal rate of return' },
  RATE: { category: 'Financial', description: 'Interest rate' },

  // Information
  ISBLANK: { category: 'Information', description: 'Checks if blank' },
  ISERROR: { category: 'Information', description: 'Checks if error' },
  ISNUMBER: { category: 'Information', description: 'Checks if number' },
  ISTEXT: { category: 'Information', description: 'Checks if text' },
  ISLOGICAL: { category: 'Information', description: 'Checks if boolean' },

  // Database
  DSUM: { category: 'Database', description: 'Sum from database' },
  DCOUNT: { category: 'Database', description: 'Count from database' },
  DAVERAGE: { category: 'Database', description: 'Average from database' },

  // Array
  SUMPRODUCT: { category: 'Array', description: 'Sum of products' },
  SUMIF: { category: 'Array', description: 'Conditional sum' },
  SUMIFS: { category: 'Array', description: 'Multi-criteria sum' },
  COUNTIF: { category: 'Array', description: 'Conditional count' },
  COUNTIFS: { category: 'Array', description: 'Multi-criteria count' },
  AVERAGEIF: { category: 'Array', description: 'Conditional average' },
  AVERAGEIFS: { category: 'Array', description: 'Multi-criteria average' },
};

// Helper to get function info
export function getFunctionInfo(name: string): { category: string; description: string } | null {
  const upperName = name.toUpperCase();
  return EXCEL_FUNCTIONS[upperName] || null;
}

// Check if a name is a known Excel function
export function isExcelFunction(name: string): boolean {
  return name.toUpperCase() in EXCEL_FUNCTIONS;
}
