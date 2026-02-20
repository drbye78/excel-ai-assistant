// Conditional Command Engine - Handles conditional logic in natural language commands
// Phase 4 Implementation - Advanced Features

import {
  ParsedCommand,
  NLContext,
  CommandIntent,
  CommandTarget,
  SupportedLocale
} from './naturalLanguageCommandParser';
import { excelService } from './excelService';

export interface ConditionalCommand {
  condition: Condition;
  trueCommand: ParsedCommand;
  falseCommand?: ParsedCommand;
}

export interface Condition {
  type: 'data_exists' | 'value_comparison' | 'cell_state' | 'range_size' | 'error_check';
  parameters: Record<string, any>;
  evaluate: (context: NLContext) => boolean | Promise<boolean>;
}

export interface ConditionResult {
  conditionMet: boolean;
  evaluatedCommand: ParsedCommand;
  explanation: string;
}

export class ConditionalCommandEngine {
  private static instance: ConditionalCommandEngine;

  private constructor() {}

  static getInstance(): ConditionalCommandEngine {
    if (!ConditionalCommandEngine.instance) {
      ConditionalCommandEngine.instance = new ConditionalCommandEngine();
    }
    return ConditionalCommandEngine.instance;
  }

  /**
   * Parse a conditional command from natural language
   */
  parseConditionalCommand(text: string): ConditionalCommand | null {
    // English patterns
    const patterns = [
      {
        regex: /if\s+(.+?)\s+(?:then\s+)?(.+?)(?:\s+else\s+(.+))?$/i,
        extract: (match: RegExpMatchArray) => ({
          condition: match[1].trim(),
          trueAction: match[2].trim(),
          falseAction: match[3]?.trim()
        })
      },
      // Russian patterns
      {
        regex: /если\s+(.+?)\s+(?:то\s+)?(.+?)(?:\s+иначе\s+(.+))?$/i,
        extract: (match: RegExpMatchArray) => ({
          condition: match[1].trim(),
          trueAction: match[2].trim(),
          falseAction: match[3]?.trim()
        })
      },
      {
        regex: /если\s+(.+?)\s*,\s*(?:то\s+)?(.+?)(?:\s*,\s*иначе\s+(.+))?$/i,
        extract: (match: RegExpMatchArray) => ({
          condition: match[1].trim(),
          trueAction: match[2].trim(),
          falseAction: match[3]?.trim()
        })
      }
    ];

    for (const pattern of patterns) {
      const match = text.match(pattern.regex);
      if (match) {
        const extracted = pattern.extract(match);
        return {
          condition: this.parseCondition(extracted.condition),
          trueCommand: this.parseSubCommand(extracted.trueAction),
          falseCommand: extracted.falseAction
            ? this.parseSubCommand(extracted.falseAction)
            : undefined
        };
      }
    }

    return null;
  }

  /**
   * Parse condition from text
   */
  private parseCondition(conditionText: string): Condition {
    const lower = conditionText.toLowerCase();

    // Data existence checks
    if (this.matchesPattern(lower, ['there are', 'есть', 'имеются', 'обнаружены'])) {
      if (this.matchesPattern(lower, ['errors', 'ошибки', 'error', 'ошибка'])) {
        return this.createErrorCheckCondition('exists');
      }
      if (this.matchesPattern(lower, ['blanks', 'пустые', 'blank cells', 'пустые ячейки'])) {
        return this.createCellStateCondition('blank', 'exists');
      }
      if (this.matchesPattern(lower, ['duplicates', 'дубликаты', 'duplicate', 'повторы'])) {
        return this.createDataExistCondition('duplicates');
      }
    }

    // Value comparisons
    const comparisonMatch = lower.match(/(.+?)\s*(>|>=|<|<=|=|равно|больше|меньше|не равно)\s*(.+)/);
    if (comparisonMatch) {
      return this.createValueComparisonCondition(
        comparisonMatch[1].trim(),
        comparisonMatch[2].trim(),
        comparisonMatch[3].trim()
      );
    }

    // Range size checks
    const rangeSizeMatch = lower.match(/(?:more than|больше|more)\s*(\d+)\s*(?:rows|строк|row|строка)/);
    if (rangeSizeMatch) {
      return this.createRangeSizeCondition('>', parseInt(rangeSizeMatch[1]));
    }

    // Cell state checks
    if (this.matchesPattern(lower, ['cell', 'ячейка', 'cells', 'ячейки'])) {
      if (this.matchesPattern(lower, ['is empty', 'пустая', 'is blank', 'пусто'])) {
        return this.createCellStateCondition('blank', 'all');
      }
      if (this.matchesPattern(lower, ['has error', 'с ошибкой', 'contains error', 'содержит ошибку'])) {
        return this.createCellStateCondition('error', 'any');
      }
    }

    // Default: data exists condition
    return this.createDataExistCondition('any');
  }

  /**
   * Parse sub-command from conditional branch
   */
  private parseSubCommand(text: string): ParsedCommand {
    // This would normally call the main parser
    // For now, return a minimal command structure
    return {
      originalText: text,
      intent: this.detectIntent(text),
      target: this.detectTarget(text),
      confidence: 'medium',
      parameters: {},
      constraints: []
    };
  }

  /**
   * Evaluate a conditional command
   */
  async evaluateConditional(
    conditional: ConditionalCommand,
    context: NLContext,
    locale: SupportedLocale = 'en'
  ): Promise<ConditionResult> {
    const evaluateResult = conditional.condition.evaluate(context);
    const conditionMet = evaluateResult instanceof Promise 
      ? await evaluateResult 
      : evaluateResult;
    const t = this.getTranslations(locale);

    if (conditionMet) {
      return {
        conditionMet: true,
        evaluatedCommand: conditional.trueCommand,
        explanation: t.conditionMetTrue
      };
    } else if (conditional.falseCommand) {
      return {
        conditionMet: false,
        evaluatedCommand: conditional.falseCommand,
        explanation: t.conditionMetFalse
      };
    } else {
      return {
        conditionMet: false,
        evaluatedCommand: {
          originalText: 'No action needed',
          intent: 'explain',
          target: 'data',
          confidence: 'high',
          parameters: {},
          constraints: []
        },
        explanation: t.noActionNeeded
      };
    }
  }

  /**
   * Check if text is a conditional command
   */
  isConditionalCommand(text: string): boolean {
    const conditionalIndicators = [
      /^if\s+/i,
      /^если\s+/i,
      /\s+if\s+there\s+/i,
      /\s+если\s+есть\s+/i,
      /^only\s+if\s+/i,
      /^только\s+если\s+/i
    ];
    return conditionalIndicators.some(pattern => pattern.test(text));
  }

  /**
   * Create data existence condition
   */
  private createDataExistCondition(dataType: string): Condition {
    return {
      type: 'data_exists',
      parameters: { dataType },
      evaluate: (context: NLContext) => {
        // Check for data in selection
        if (dataType === 'any') {
          return !!(context.selectedRange && context.rowCount && context.rowCount > 0);
        }
        // For other types, would need actual data inspection
        return true;
      }
    };
  }

  /**
   * Create value comparison condition
   */
  private createValueComparisonCondition(
    column: string,
    operator: string,
    value: string
  ): Condition {
    const normalizeOperator = (op: string): string => {
      const map: Record<string, string> = {
        '>': 'gt', '>=': 'gte', '<': 'lt', '<=': 'lte', '=': 'eq',
        'больше': 'gt', 'меньше': 'lt', 'равно': 'eq', 'не равно': 'ne'
      };
      return map[op] || op;
    };

    return {
      type: 'value_comparison',
      parameters: { column, operator: normalizeOperator(operator), value },
      evaluate: async (context: NLContext): Promise<boolean> => {
        // Evaluate value comparison by checking actual data
        if (!context.selectedRange && !context.activeTable) {
          return false;
        }
        
        try {
          const range = context.selectedRange || context.activeTable || '';
          const rangeContext = await excelService.getRange(range, context.selectedWorksheet);
          const values = rangeContext.values;
          
          // Find column index if column name provided
          let columnIndex = 0;
          if (column && typeof column === 'string') {
            // Try to find column by name (check first row for headers)
            if (values.length > 0) {
              const headerRow = values[0];
              columnIndex = headerRow.findIndex((h: any) => 
                String(h).toLowerCase() === column.toLowerCase()
              );
              if (columnIndex === -1) {
                // Try parsing as column letter (A, B, C, etc.)
                const colMatch = column.match(/^([A-Z]+)/i);
                if (colMatch) {
                  const colLetter = colMatch[1].toUpperCase();
                  columnIndex = colLetter.charCodeAt(0) - 65; // A=0, B=1, etc.
                } else {
                  columnIndex = 0; // Default to first column
                }
              }
            }
          }
          
          // Compare values in the column
          const op = normalizeOperator(operator);
          for (let row = 1; row < values.length; row++) {
            const cellValue = values[row][columnIndex];
            const numValue = typeof value === 'string' ? parseFloat(value) : value;
            const numCellValue = typeof cellValue === 'string' ? parseFloat(cellValue) : cellValue;
            
            let matches = false;
            switch (op) {
              case 'gt': matches = numCellValue > numValue; break;
              case 'lt': matches = numCellValue < numValue; break;
              case 'eq': matches = numCellValue === numValue || String(cellValue).toLowerCase() === String(value).toLowerCase(); break;
              case 'ne': matches = numCellValue !== numValue && String(cellValue).toLowerCase() !== String(value).toLowerCase(); break;
              case 'gte': matches = numCellValue >= numValue; break;
              case 'lte': matches = numCellValue <= numValue; break;
            }
            
            if (matches) return true;
          }
          
          return false;
        } catch (error) {
          // If evaluation fails, return false
          return false;
        }
      }
    } as Condition & { evaluate: (context: NLContext) => Promise<boolean> };
  }

  /**
   * Create cell state condition
   */
  private createCellStateCondition(state: string, scope: string): Condition {
    return {
      type: 'cell_state',
      parameters: { state, scope },
      evaluate: async (context: NLContext): Promise<boolean> => {
        // Evaluate cell state by checking actual data
        if (!context.selectedRange) {
          return false;
        }
        
        try {
          const rangeContext = await excelService.getRange(context.selectedRange, context.selectedWorksheet);
          const values = rangeContext.values;
          
          switch (state.toLowerCase()) {
            case 'blank':
            case 'empty':
              // Check if any cells are blank
              for (const row of values) {
                for (const cell of row) {
                  if (cell === null || cell === undefined || cell === '') {
                    return true;
                  }
                }
              }
              return false;
              
            case 'filled':
            case 'has_data':
              // Check if all cells have data
              for (const row of values) {
                for (const cell of row) {
                  if (cell === null || cell === undefined || cell === '') {
                    return false;
                  }
                }
              }
              return values.length > 0;
              
            case 'has_errors':
              // Check for error values (would need formula evaluation)
              // For now, check for common error indicators
              for (const row of values) {
                for (const cell of row) {
                  const cellStr = String(cell).toUpperCase();
                  if (cellStr.includes('#ERROR') || cellStr.includes('#N/A') || 
                      cellStr.includes('#VALUE') || cellStr.includes('#REF')) {
                    return true;
                  }
                }
              }
              return false;
              
            default:
              return false;
          }
        } catch (error) {
          return false;
        }
      }
    } as Condition & { evaluate: (context: NLContext) => Promise<boolean> };
  }

  /**
   * Create range size condition
   */
  private createRangeSizeCondition(operator: string, threshold: number): Condition {
    return {
      type: 'range_size',
      parameters: { operator, threshold },
      evaluate: (context: NLContext) => {
        if (!context.rowCount) return false;
        switch (operator) {
          case '>': return context.rowCount > threshold;
          case '>=': return context.rowCount >= threshold;
          case '<': return context.rowCount < threshold;
          case '<=': return context.rowCount <= threshold;
          default: return false;
        }
      }
    };
  }

  /**
   * Create error check condition
   */
  private createErrorCheckCondition(checkType: string): Condition {
    return {
      type: 'error_check',
      parameters: { checkType },
      evaluate: async (context: NLContext): Promise<boolean> => {
        // Evaluate error check by inspecting actual data
        if (!context.selectedRange) {
          return false;
        }
        
        try {
          const rangeContext = await excelService.getRange(context.selectedRange, context.selectedWorksheet);
          const values = rangeContext.values;
          const formulas = rangeContext.formulas;
          
          switch (checkType.toLowerCase()) {
            case 'errors':
            case 'has_errors':
              // Check for error values in formulas or cells
              for (let row = 0; row < values.length; row++) {
                for (let col = 0; col < values[row].length; col++) {
                  const cellValue = String(values[row][col] || '').toUpperCase();
                  const formula = String(formulas[row][col] || '');
                  
                  if (cellValue.includes('#ERROR') || cellValue.includes('#N/A') || 
                      cellValue.includes('#VALUE') || cellValue.includes('#REF') ||
                      cellValue.includes('#DIV/0') || cellValue.includes('#NAME') ||
                      cellValue.includes('#NUM') || cellValue.includes('#NULL')) {
                    return true;
                  }
                }
              }
              return false;
              
            case 'duplicates':
              // Check for duplicate values
              const seen = new Set<string>();
              for (const row of values) {
                for (const cell of row) {
                  const cellStr = String(cell || '').trim();
                  if (cellStr && seen.has(cellStr)) {
                    return true;
                  }
                  if (cellStr) seen.add(cellStr);
                }
              }
              return false;
              
            default:
              return false;
          }
        } catch (error) {
          return false;
        }
      }
    } as Condition & { evaluate: (context: NLContext) => Promise<boolean> };
  }

  /**
   * Helper to match pattern against text
   */
  private matchesPattern(text: string, patterns: string[]): boolean {
    return patterns.some(p => text.includes(p.toLowerCase()));
  }

  /**
   * Detect intent from text (simplified)
   */
  private detectIntent(text: string): CommandIntent {
    const intentPatterns: Record<CommandIntent, string[]> = {
      'create': ['create', 'make', 'add', 'создать', 'сделать', 'добавить'],
      'modify': ['modify', 'change', 'update', 'изменить', 'поменять', 'обновить'],
      'delete': ['delete', 'remove', 'clear', 'удалить', 'убрать', 'очистить'],
      'explain': ['explain', 'describe', 'объяснить', 'описать', 'показать'],
      'format': ['format', 'style', 'форматировать', 'стиль'],
      'analyze': ['analyze', 'calculate', 'анализировать', 'вычислить', 'посчитать'],
      'filter': ['filter', 'show', 'фильтровать', 'показать', 'отфильтровать'],
      'sort': ['sort', 'order', 'сортировать', 'упорядочить'],
      'calculate': ['calculate', 'compute', 'calculate', 'вычислить', 'посчитать'],
      'refresh': ['refresh', 'update', 'обновить', 'перезагрузить'],
      'export': ['export', 'save', 'экспортировать', 'сохранить'],
      'import': ['import', 'load', 'импортировать', 'загрузить']
    };

    const lower = text.toLowerCase();
    for (const [intent, patterns] of Object.entries(intentPatterns)) {
      if (patterns.some(p => lower.includes(p))) {
        return intent as CommandIntent;
      }
    }
    return 'explain';
  }

  /**
   * Detect target from text (simplified)
   */
  private detectTarget(text: string): CommandTarget {
    const targetPatterns: Record<CommandTarget, string[]> = {
      'pivot': ['pivot', 'сводная'],
      'chart': ['chart', 'graph', 'диаграмма', 'график'],
      'table': ['table', 'таблица'],
      'query': ['query', 'запрос'],
      'measure': ['measure', 'мера'],
      'range': ['range', 'cells', 'диапазон', 'ячейки'],
      'worksheet': ['sheet', 'worksheet', 'лист'],
      'workbook': ['workbook', 'книга'],
      'shape': ['shape', 'фигура'],
      'image': ['image', 'изображение', 'картинка'],
      'formula': ['formula', 'формула'],
      'data': ['data', 'данные']
    };

    const lower = text.toLowerCase();
    for (const [target, patterns] of Object.entries(targetPatterns)) {
      if (patterns.some(p => lower.includes(p))) {
        return target as CommandTarget;
      }
    }
    return 'range';
  }

  /**
   * Get translations
   */
  private getTranslations(locale: SupportedLocale) {
    const translations = {
      en: {
        conditionMetTrue: 'Condition met, executing primary action',
        conditionMetFalse: 'Condition not met, executing alternative action',
        noActionNeeded: 'Condition not met and no alternative specified'
      },
      ru: {
        conditionMetTrue: 'Условие выполнено, выполняется основное действие',
        conditionMetFalse: 'Условие не выполнено, выполняется альтернативное действие',
        noActionNeeded: 'Условие не выполнено, альтернатива не указана'
      }
    };
    return translations[locale];
  }
}

export default ConditionalCommandEngine.getInstance();
