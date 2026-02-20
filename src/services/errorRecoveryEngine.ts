// Error Recovery Engine - Smart error handling with alternative suggestions
// Phase 3 Implementation - Intelligence & Optimization

import {
  ParsedCommand,
  NLContext,
  SupportedLocale
} from './naturalLanguageCommandParser';

export interface RecoverySuggestion {
  type: 'alternative_command' | 'use_selection' | 'create_first' | 'fix_syntax' | 'simplify';
  description: string;
  command: string;
  confidence: 'high' | 'medium' | 'low';
}

export interface ErrorRecoveryResult {
  originalCommand: string;
  errorType: string;
  explanation: string;
  suggestions: RecoverySuggestion[];
  canAutoFix: boolean;
  autoFixCommand?: string;
}

export class ErrorRecoveryEngine {
  private static instance: ErrorRecoveryEngine;

  private constructor() {}

  static getInstance(): ErrorRecoveryEngine {
    if (!ErrorRecoveryEngine.instance) {
      ErrorRecoveryEngine.instance = new ErrorRecoveryEngine();
    }
    return ErrorRecoveryEngine.instance;
  }

  /**
   * Analyze error and generate recovery suggestions
   */
  analyzeError(
    command: ParsedCommand,
    error: Error,
    context: NLContext,
    locale: SupportedLocale = 'en'
  ): ErrorRecoveryResult {
    const errorMsg = error.message.toLowerCase();
    const t = this.getTranslations(locale);

    // Range/address errors
    if (errorMsg.includes('range') || errorMsg.includes('address') || errorMsg.includes('invalid reference')) {
      return this.handleRangeError(command, context, t);
    }

    // Table not found
    if (errorMsg.includes('table') && (errorMsg.includes('not found') || errorMsg.includes('does not exist'))) {
      return this.handleTableError(command, context, t);
    }

    // Chart/pivot not found
    if (errorMsg.includes('chart') || errorMsg.includes('pivot')) {
      return this.handleChartError(command, context, t);
    }

    // Formula errors
    if (errorMsg.includes('formula') || errorMsg.includes('syntax')) {
      return this.handleFormulaError(command, t);
    }

    // Permission/protection errors
    if (errorMsg.includes('protected') || errorMsg.includes('permission') || errorMsg.includes('locked')) {
      return this.handleProtectionError(command, t);
    }

    // Data type errors
    if (errorMsg.includes('data type') || errorMsg.includes('cannot convert')) {
      return this.handleDataTypeError(command, t);
    }

    // Generic fallback
    return this.handleGenericError(command, t);
  }

  /**
   * Handle range/address related errors
   */
  private handleRangeError(
    command: ParsedCommand,
    context: NLContext,
    t: any
  ): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    // Suggest using current selection
    if (context.selectedRange) {
      suggestions.push({
        type: 'use_selection',
        description: t.useCurrentSelection.replace('{range}', context.selectedRange),
        command: command.originalText.replace(
          command.parameters?.range || '',
          context.selectedRange
        ),
        confidence: 'high'
      });
    }

    // Suggest creating the range first
    suggestions.push({
      type: 'create_first',
      description: t.createRangeFirst,
      command: `Select or create range ${command.parameters?.range || 'A1:D10'}`,
      confidence: 'medium'
    });

    // Simplify command
    suggestions.push({
      type: 'simplify',
      description: t.simplifyCommand,
      command: this.simplifyCommand(command.originalText),
      confidence: 'medium'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'range_not_found',
      explanation: t.rangeNotFound,
      suggestions,
      canAutoFix: !!context.selectedRange,
      autoFixCommand: context.selectedRange
        ? command.originalText.replace(command.parameters?.range || '', context.selectedRange)
        : undefined
    };
  }

  /**
   * Handle table not found errors
   */
  private handleTableError(
    command: ParsedCommand,
    context: NLContext,
    t: any
  ): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];
    const availableTables = context.availableTables || [];

    // Suggest using available tables
    if (availableTables.length > 0) {
      availableTables.slice(0, 3).forEach(table => {
        suggestions.push({
          type: 'alternative_command',
          description: t.useExistingTable.replace('{table}', table),
          command: command.originalText.replace(
            command.parameters?.tableName || '',
            table
          ),
          confidence: 'high'
        });
      });
    }

    // Suggest creating table from selection
    if (context.selectedRange) {
      suggestions.push({
        type: 'create_first',
        description: t.createTableFromSelection,
        command: `Create table from ${context.selectedRange}`,
        confidence: 'medium'
      });
    }

    // Suggest using range instead
    suggestions.push({
      type: 'alternative_command',
      description: t.useRangeInstead,
      command: command.originalText
        .replace(/table\s+\w+/i, `range ${context.selectedRange || 'A1:D10'}`),
      confidence: 'low'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'table_not_found',
      explanation: t.tableNotFound.replace('{table}', command.parameters?.tableName || ''),
      suggestions,
      canAutoFix: availableTables.length > 0,
      autoFixCommand: availableTables.length > 0
        ? command.originalText.replace(command.parameters?.tableName || '', availableTables[0])
        : undefined
    };
  }

  /**
   * Handle chart/pivot related errors
   */
  private handleChartError(
    command: ParsedCommand,
    context: NLContext,
    t: any
  ): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    // Suggest creating first
    if (context.selectedRange) {
      suggestions.push({
        type: 'create_first',
        description: t.createChartFirst,
        command: `Create chart from ${context.selectedRange}`,
        confidence: 'high'
      });
    }

    // Simplify
    suggestions.push({
      type: 'simplify',
      description: t.simplifyCommand,
      command: this.simplifyCommand(command.originalText),
      confidence: 'medium'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'chart_not_found',
      explanation: t.chartNotFound,
      suggestions,
      canAutoFix: false
    };
  }

  /**
   * Handle formula syntax errors
   */
  private handleFormulaError(command: ParsedCommand, t: any): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    // Suggest simpler formula
    suggestions.push({
      type: 'simplify',
      description: t.useSimplerFormula,
      command: 'Explain this formula',
      confidence: 'high'
    });

    // Suggest validation
    suggestions.push({
      type: 'fix_syntax',
      description: t.checkFormulaSyntax,
      command: 'Validate formula',
      confidence: 'medium'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'formula_error',
      explanation: t.formulaError,
      suggestions,
      canAutoFix: false
    };
  }

  /**
   * Handle protection/permission errors
   */
  private handleProtectionError(command: ParsedCommand, t: any): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    suggestions.push({
      type: 'alternative_command',
      description: t.unprotectFirst,
      command: 'Unprotect sheet',
      confidence: 'high'
    });

    suggestions.push({
      type: 'alternative_command',
      description: t.workOnDifferentSheet,
      command: command.originalText.replace(/current|this/i, 'new'),
      confidence: 'medium'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'protection_error',
      explanation: t.protectionError,
      suggestions,
      canAutoFix: false
    };
  }

  /**
   * Handle data type errors
   */
  private handleDataTypeError(command: ParsedCommand, t: any): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    suggestions.push({
      type: 'fix_syntax',
      description: t.convertDataFirst,
      command: command.originalText.replace('calculate', 'convert to numbers then calculate'),
      confidence: 'high'
    });

    suggestions.push({
      type: 'alternative_command',
      description: t.cleanDataFirst,
      command: 'Clean data before ' + command.originalText,
      confidence: 'medium'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'data_type_error',
      explanation: t.dataTypeError,
      suggestions,
      canAutoFix: false
    };
  }

  /**
   * Handle generic errors
   */
  private handleGenericError(command: ParsedCommand, t: any): ErrorRecoveryResult {
    const suggestions: RecoverySuggestion[] = [];

    suggestions.push({
      type: 'simplify',
      description: t.simplifyCommand,
      command: this.simplifyCommand(command.originalText),
      confidence: 'medium'
    });

    suggestions.push({
      type: 'alternative_command',
      description: t.rephraseCommand,
      command: 'Please ' + command.originalText,
      confidence: 'low'
    });

    return {
      originalCommand: command.originalText,
      errorType: 'unknown_error',
      explanation: t.unknownError,
      suggestions,
      canAutoFix: false
    };
  }

  /**
   * Simplify a complex command
   */
  private simplifyCommand(command: string): string {
    return command
      .replace(/and\s+then/gi, '. Then')
      .replace(/,\s*/g, '. ')
      .replace(/with\s+(.+?)\s+and/gi, '')
      .split('.')[0]
      .trim();
  }

  /**
   * Get translations for the current locale
   */
  private getTranslations(locale: SupportedLocale) {
    const translations = {
      en: {
        useCurrentSelection: 'Use current selection ({range})',
        createRangeFirst: 'Create or select the range first',
        simplifyCommand: 'Try a simpler version of this command',
        useExistingTable: 'Use existing table "{table}"',
        createTableFromSelection: 'Create table from current selection',
        useRangeInstead: 'Use a cell range instead',
        createChartFirst: 'Create a chart from selection first',
        useSimplerFormula: 'Use a simpler formula',
        checkFormulaSyntax: 'Check formula syntax',
        unprotectFirst: 'Unprotect the sheet first',
        workOnDifferentSheet: 'Work on a different sheet',
        convertDataFirst: 'Convert data to numbers first',
        cleanDataFirst: 'Clean data first',
        rephraseCommand: 'Rephrase the command',
        rangeNotFound: 'The specified range could not be found. Try using your current selection.',
        tableNotFound: 'Table "{table}" does not exist. Use an existing table or create one first.',
        chartNotFound: 'Chart not found. Create a chart from your selection first.',
        formulaError: 'The formula syntax appears to be invalid. Try a simpler formula.',
        protectionError: 'The sheet or workbook is protected. Unprotect it first or work on a different sheet.',
        dataTypeError: 'Data type mismatch. Convert data to appropriate format first.',
        unknownError: 'An unexpected error occurred. Try simplifying your command.'
      },
      ru: {
        useCurrentSelection: 'Использовать текущее выделение ({range})',
        createRangeFirst: 'Сначала создайте или выберите диапазон',
        simplifyCommand: 'Попробуйте более простую версию команды',
        useExistingTable: 'Использовать существующую таблицу "{table}"',
        createTableFromSelection: 'Создать таблицу из текущего выделения',
        useRangeInstead: 'Использовать диапазон ячеек вместо таблицы',
        createChartFirst: 'Сначала создайте диаграмму из выделения',
        useSimplerFormula: 'Использовать более простую формулу',
        checkFormulaSyntax: 'Проверить синтаксис формулы',
        unprotectFirst: 'Сначала снять защиту с листа',
        workOnDifferentSheet: 'Работать на другом листе',
        convertDataFirst: 'Сначала преобразовать данные в числа',
        cleanDataFirst: 'Сначала очистить данные',
        rephraseCommand: 'Переформулировать команду',
        rangeNotFound: 'Указанный диапазон не найден. Попробуйте использовать текущее выделение.',
        tableNotFound: 'Таблица "{table}" не существует. Используйте существующую таблицу или создайте новую.',
        chartNotFound: 'Диаграмма не найдена. Сначала создайте диаграмму из выделения.',
        formulaError: 'Синтаксис формулы некорректен. Попробуйте более простую формулу.',
        protectionError: 'Лист или книга защищены. Снимите защиту или работайте на другом листе.',
        dataTypeError: 'Несоответствие типов данных. Сначала преобразуйте данные в подходящий формат.',
        unknownError: 'Произошла непредвиденная ошибка. Попробуйте упростить команду.'
      }
    };

    return translations[locale];
  }

  /**
   * Check if a command can be auto-fixed
   */
  canAutoFix(error: Error): boolean {
    const autoFixableErrors = [
      'range',
      'table',
      'address'
    ];
    return autoFixableErrors.some(pattern =>
      error.message.toLowerCase().includes(pattern)
    );
  }
}

export default ErrorRecoveryEngine.getInstance();
