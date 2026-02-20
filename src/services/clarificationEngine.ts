// Clarification Engine - Handles low confidence parsing with user clarification
// Phase 2 Implementation - Conversational Interface

import {
  ParsedCommand,
  NLContext,
  CommandIntent,
  CommandTarget,
  SupportedLocale
} from './naturalLanguageCommandParser';

export type ClarificationType = 
  | 'ambiguous_intent'
  | 'missing_target'
  | 'missing_parameters'
  | 'low_confidence'
  | 'multiple_matches';

export interface ClarificationOption {
  id: string;
  label: string;
  value: string;
  icon?: string;
}

export interface ClarificationRequest {
  type: ClarificationType;
  message: string;
  description: string;
  options?: ClarificationOption[];
  expectedInput: 'selection' | 'text' | 'range' | 'confirmation' | 'none';
  canSkip: boolean;
  context?: Record<string, any>;
}

export interface ClarificationResponse {
  requestId: string;
  selectedOption?: string;
  textInput?: string;
  confirmed: boolean;
}

export class ClarificationEngine {
  private static instance: ClarificationEngine;
  private activeRequests: Map<string, ClarificationRequest> = new Map();
  private requestCounter: number = 0;

  private constructor() {}

  static getInstance(): ClarificationEngine {
    if (!ClarificationEngine.instance) {
      ClarificationEngine.instance = new ClarificationEngine();
    }
    return ClarificationEngine.instance;
  }

  /**
   * Check if a command needs clarification
   */
  needsClarification(
    command: ParsedCommand,
    context?: NLContext
  ): ClarificationRequest | null {
    // Low confidence detection
    if (command.confidence === 'low') {
      return this.generateLowConfidenceClarification(command);
    }

    // Missing target
    if (command.target === 'range' && !this.hasExplicitTarget(command)) {
      return this.generateMissingTargetClarification(command, context);
    }

    // Missing critical parameters
    if (this.hasMissingCriticalParameters(command)) {
      return this.generateMissingParametersClarification(command, context);
    }

    // Ambiguous intent detection
    if (this.isAmbiguousIntent(command)) {
      return this.generateAmbiguousIntentClarification(command);
    }

    return null;
  }

  /**
   * Generate clarification for low confidence parsing
   */
  private generateLowConfidenceClarification(
    command: ParsedCommand,
    locale: SupportedLocale = 'en'
  ): ClarificationRequest {
    const t = this.getTranslations(locale);
    const summary = this.summarizeCommand(command, locale);

    return {
      type: 'low_confidence',
      message: t.lowConfidenceTitle,
      description: t.lowConfidenceMessage.replace('{summary}', summary),
      options: [
        { id: 'yes', label: t.yesProceed, value: 'proceed', icon: '✓' },
        { id: 'no', label: t.noRephrase, value: 'rephrase', icon: '✗' },
        { id: 'examples', label: t.showExamples, value: 'examples', icon: '💡' }
      ],
      expectedInput: 'selection',
      canSkip: false,
      context: { command }
    };
  }

  /**
   * Generate clarification for missing target
   */
  private generateMissingTargetClarification(
    command: ParsedCommand,
    context?: NLContext,
    locale: SupportedLocale = 'en'
  ): ClarificationRequest {
    const t = this.getTranslations(locale);

    const options: ClarificationOption[] = [
      { id: 'chart', label: t.targetChart, value: 'chart', icon: '📊' },
      { id: 'table', label: t.targetTable, value: 'table', icon: '📋' },
      { id: 'pivot', label: t.targetPivot, value: 'pivot', icon: '🔄' },
      { id: 'range', label: t.targetRange, value: 'range', icon: '▭' },
      { id: 'formula', label: t.targetFormula, value: 'formula', icon: '𝑓' }
    ];

    // Add context-aware options
    if (context?.selectedChart) {
      options.unshift({
        id: 'current_chart',
        label: t.currentChart.replace('{name}', context.selectedChart),
        value: 'current_chart',
        icon: '📊'
      });
    }

    if (context?.activeTable) {
      options.unshift({
        id: 'current_table',
        label: t.currentTable.replace('{name}', context.activeTable),
        value: 'current_table',
        icon: '📋'
      });
    }

    return {
      type: 'missing_target',
      message: t.missingTargetTitle,
      description: t.missingTargetMessage,
      options,
      expectedInput: 'selection',
      canSkip: false,
      context: { command }
    };
  }

  /**
   * Generate clarification for missing parameters
   */
  private generateMissingParametersClarification(
    command: ParsedCommand,
    context?: NLContext,
    locale: SupportedLocale = 'en'
  ): ClarificationRequest {
    const t = this.getTranslations(locale);
    const missingParams = this.getMissingParameters(command);

    if (missingParams.includes('range') || missingParams.includes('dataRange')) {
      return {
        type: 'missing_parameters',
        message: t.missingRangeTitle,
        description: t.missingRangeMessage,
        options: context?.selectedRange ? [
          {
            id: 'current_selection',
            label: t.useCurrentSelection.replace('{range}', context.selectedRange),
            value: context.selectedRange,
            icon: '✓'
          },
          { id: 'entire_sheet', label: t.useEntireSheet, value: 'entire_sheet', icon: '▭' },
          { id: 'specify', label: t.specifyRange, value: 'specify', icon: '✏️' }
        ] : undefined,
        expectedInput: context?.selectedRange ? 'selection' : 'range',
        canSkip: false,
        context: { command, missingParams }
      };
    }

    if (missingParams.includes('column') || missingParams.includes('columns')) {
      const columnOptions = context?.availableColumns?.slice(0, 5).map((col, idx) => ({
        id: `col_${idx}`,
        label: col,
        value: col,
        icon: '▭'
      })) || [];

      return {
        type: 'missing_parameters',
        message: t.missingColumnTitle,
        description: t.missingColumnMessage,
        options: columnOptions.length > 0 ? columnOptions : undefined,
        expectedInput: columnOptions.length > 0 ? 'selection' : 'text',
        canSkip: false,
        context: { command, missingParams }
      };
    }

    return {
      type: 'missing_parameters',
      message: t.missingParamsTitle,
      description: t.missingParamsMessage,
      expectedInput: 'text',
      canSkip: false,
      context: { command, missingParams }
    };
  }

  /**
   * Generate clarification for ambiguous intent
   */
  private generateAmbiguousIntentClarification(
    command: ParsedCommand,
    locale: SupportedLocale = 'en'
  ): ClarificationRequest {
    const t = this.getTranslations(locale);

    return {
      type: 'ambiguous_intent',
      message: t.ambiguousIntentTitle,
      description: t.ambiguousIntentMessage,
      options: [
        { id: 'create', label: t.intentCreate, value: 'create', icon: '➕' },
        { id: 'modify', label: t.intentModify, value: 'modify', icon: '✏️' },
        { id: 'analyze', label: t.intentAnalyze, value: 'analyze', icon: '📊' },
        { id: 'format', label: t.intentFormat, value: 'format', icon: '🎨' }
      ],
      expectedInput: 'selection',
      canSkip: false,
      context: { command }
    };
  }

  /**
   * Handle clarification response and update command
   */
  handleClarificationResponse(
    originalCommand: ParsedCommand,
    clarification: ClarificationRequest,
    response: ClarificationResponse
  ): ParsedCommand {
    const updatedCommand = { ...originalCommand };

    switch (clarification.type) {
      case 'low_confidence':
        if (response.selectedOption === 'proceed') {
          updatedCommand.confidence = 'high';
        } else if (response.selectedOption === 'rephrase') {
          updatedCommand.confidence = 'low';
          updatedCommand.suggestions = ['Try a different wording'];
        }
        break;

      case 'missing_target':
        if (response.selectedOption) {
          updatedCommand.target = response.selectedOption as CommandTarget;
        }
        break;

      case 'missing_parameters':
        if (response.selectedOption) {
          updatedCommand.parameters = {
            ...updatedCommand.parameters,
            ...this.parseParameterResponse(clarification, response)
          };
        } else if (response.textInput) {
          updatedCommand.parameters = {
            ...updatedCommand.parameters,
            [this.getMissingParameters(updatedCommand)[0]]: response.textInput
          };
        }
        break;

      case 'ambiguous_intent':
        if (response.selectedOption) {
          updatedCommand.intent = response.selectedOption as CommandIntent;
        }
        break;
    }

    return updatedCommand;
  }

  /**
   * Check if command has explicit target
   */
  private hasExplicitTarget(command: ParsedCommand): boolean {
    const text = command.originalText.toLowerCase();
    const targetIndicators = [
      'chart', 'table', 'pivot', 'graph', 'диаграмма', 'таблица', 'сводная'
    ];
    return targetIndicators.some(indicator => text.includes(indicator));
  }

  /**
   * Check for missing critical parameters
   */
  private hasMissingCriticalParameters(command: ParsedCommand): boolean {
    const missing = this.getMissingParameters(command);
    return missing.length > 0;
  }

  /**
   * Get list of missing critical parameters
   */
  private getMissingParameters(command: ParsedCommand): string[] {
    const missing: string[] = [];

    switch (command.target) {
      case 'chart':
        if (!command.parameters?.range && !command.parameters?.dataRange) {
          missing.push('dataRange');
        }
        break;
      case 'pivot':
        if (!command.parameters?.sourceData) {
          missing.push('sourceData');
        }
        break;
      case 'table':
        if (!command.parameters?.range) {
          missing.push('range');
        }
        break;
    }

    // Check for missing columns based on intent
    if (command.intent === 'sort' && !command.parameters?.sortBy && !command.parameters?.columns) {
      missing.push('column');
    }

    return missing;
  }

  /**
   * Check if intent is ambiguous
   */
  private isAmbiguousIntent(command: ParsedCommand): boolean {
    const ambiguousPatterns = [
      /make\s+a\s+(chart|table|pivot)/i,
      /create\s+(chart|table|pivot)/i,
      /do\s+(analysis|calculation)/i,
      /сделай\s+(диаграмму|таблицу)/i
    ];
    return ambiguousPatterns.some(p => p.test(command.originalText));
  }

  /**
   * Summarize command for user confirmation
   */
  private summarizeCommand(command: ParsedCommand, locale: SupportedLocale): string {
    const t = this.getTranslations(locale);
    const intent = command.intent;
    const target = command.target;

    return t.commandSummary
      .replace('{intent}', intent)
      .replace('{target}', target);
  }

  /**
   * Parse parameter response from clarification
   */
  private parseParameterResponse(
    clarification: ClarificationRequest,
    response: ClarificationResponse
  ): Record<string, any> {
    const params: Record<string, any> = {};

    if (clarification.context?.missingParams) {
      const missingParam = clarification.context.missingParams[0];
      params[missingParam] = response.selectedOption || response.textInput;
    }

    return params;
  }

  /**
   * Get translations for current locale
   */
  private getTranslations(locale: SupportedLocale) {
    const translations = {
      en: {
        lowConfidenceTitle: 'Did I understand correctly?',
        lowConfidenceMessage: 'I understood: "{summary}"',
        yesProceed: 'Yes, proceed',
        noRephrase: 'No, rephrase',
        showExamples: 'Show examples',
        missingTargetTitle: 'What would you like to work with?',
        missingTargetMessage: 'Please select what you want to work with:',
        targetChart: 'Chart',
        targetTable: 'Table',
        targetPivot: 'Pivot Table',
        targetRange: 'Cell Range',
        targetFormula: 'Formula',
        currentChart: 'Current chart ({name})',
        currentTable: 'Current table ({name})',
        missingRangeTitle: 'Which data range?',
        missingRangeMessage: 'Please specify the data range to use:',
        useCurrentSelection: 'Current selection ({range})',
        useEntireSheet: 'Entire worksheet',
        specifyRange: 'Specify range manually',
        missingColumnTitle: 'Which column?',
        missingColumnMessage: 'Please select the column to use:',
        missingParamsTitle: 'Missing information',
        missingParamsMessage: 'Please provide the missing information:',
        ambiguousIntentTitle: 'What would you like to do?',
        ambiguousIntentMessage: 'Please clarify what action you want to perform:',
        intentCreate: 'Create something new',
        intentModify: 'Modify existing',
        intentAnalyze: 'Analyze data',
        intentFormat: 'Format/style',
        commandSummary: '{intent} {target}'
      },
      ru: {
        lowConfidenceTitle: 'Я правильно понял?',
        lowConfidenceMessage: 'Я понял: "{summary}"',
        yesProceed: 'Да, выполнить',
        noRephrase: 'Нет, переформулировать',
        showExamples: 'Показать примеры',
        missingTargetTitle: 'С чем вы хотите работать?',
        missingTargetMessage: 'Пожалуйста, выберите объект для работы:',
        targetChart: 'Диаграмма',
        targetTable: 'Таблица',
        targetPivot: 'Сводная таблица',
        targetRange: 'Диапазон ячеек',
        targetFormula: 'Формула',
        currentChart: 'Текущая диаграмма ({name})',
        currentTable: 'Текущая таблица ({name})',
        missingRangeTitle: 'Какой диапазон данных?',
        missingRangeMessage: 'Укажите диапазон данных:',
        useCurrentSelection: 'Текущее выделение ({range})',
        useEntireSheet: 'Весь лист',
        specifyRange: 'Указать диапазон вручную',
        missingColumnTitle: 'Какой столбец?',
        missingColumnMessage: 'Выберите столбец:',
        missingParamsTitle: 'Недостающая информация',
        missingParamsMessage: 'Пожалуйста, укажите недостающую информацию:',
        ambiguousIntentTitle: 'Что вы хотите сделать?',
        ambiguousIntentMessage: 'Уточните, какое действие выполнить:',
        intentCreate: 'Создать что-то новое',
        intentModify: 'Изменить существующее',
        intentAnalyze: 'Проанализировать данные',
        intentFormat: 'Отформатировать',
        commandSummary: '{intent} {target}'
      }
    };

    return translations[locale];
  }

  /**
   * Store active clarification request
   */
  storeRequest(request: ClarificationRequest): string {
    const id = `clarification_${++this.requestCounter}`;
    this.activeRequests.set(id, request);
    return id;
  }

  /**
   * Get stored clarification request
   */
  getRequest(id: string): ClarificationRequest | undefined {
    return this.activeRequests.get(id);
  }

  /**
   * Remove stored clarification request
   */
  removeRequest(id: string): void {
    this.activeRequests.delete(id);
  }
}

export default ClarificationEngine.getInstance();
