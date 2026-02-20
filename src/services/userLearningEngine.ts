// User Learning Engine - Learns from user preferences and command patterns
// Phase 4 Implementation - Advanced Features

import {
  ParsedCommand,
  CommandIntent,
  CommandTarget,
  SupportedLocale
} from './naturalLanguageCommandParser';

export interface UserPreference {
  key: string;
  value: any;
  frequency: number;
  lastUsed: Date;
  contexts: string[];
}

export interface LearningPattern {
  pattern: string;
  preferredOutcome: string;
  successRate: number;
  usageCount: number;
}

export interface UserLearningProfile {
  preferredChartType?: string;
  preferredNumberFormat?: string;
  preferredAggregation?: string;
  preferredTableStyle?: string;
  frequentlyUsedColumns: Map<string, number>;
  commandPatterns: LearningPattern[];
  corrections: Map<string, string>;
}

export class UserLearningEngine {
  private static instance: UserLearningEngine;
  private profile: UserLearningProfile;
  private readonly STORAGE_KEY = 'excel_ai_user_learning_profile';

  private constructor() {
    this.profile = this.loadProfile();
  }

  static getInstance(): UserLearningEngine {
    if (!UserLearningEngine.instance) {
      UserLearningEngine.instance = new UserLearningEngine();
    }
    return UserLearningEngine.instance;
  }

  /**
   * Learn from a successfully executed command
   */
  learnFromCommand(command: ParsedCommand, context?: Record<string, any>): void {
    // Learn preferred chart types
    if (command.target === 'chart' && command.parameters?.chartType) {
      this.learnPreference('preferredChartType', command.parameters.chartType);
    }

    // Learn preferred number formats
    if (command.intent === 'format' && command.parameters?.numberFormat) {
      this.learnPreference('preferredNumberFormat', command.parameters.numberFormat);
    }

    // Learn preferred aggregations for pivot/measure
    if ((command.target === 'pivot' || command.target === 'measure') && command.parameters?.aggregation) {
      this.learnPreference('preferredAggregation', command.parameters.aggregation);
    }

    // Learn preferred table styles
    if (command.target === 'table' && command.parameters?.style) {
      this.learnPreference('preferredTableStyle', command.parameters.style);
    }

    // Learn frequently used columns
    if (command.parameters?.columns) {
      command.parameters.columns.forEach((col: string) => {
        this.learnColumnUsage(col);
      });
    }

    // Learn command patterns
    this.learnCommandPattern(command);

    // Save profile
    this.saveProfile();
  }

  /**
   * Learn from explicit user correction
   */
  learnCorrection(original: string, corrected: string, wasSuccessful: boolean): void {
    if (wasSuccessful) {
      const normalizedOriginal = this.normalizeForLearning(original);
      this.profile.corrections.set(normalizedOriginal, corrected);
      
      // Also learn as a pattern
      this.profile.commandPatterns.push({
        pattern: normalizedOriginal,
        preferredOutcome: corrected,
        successRate: 1.0,
        usageCount: 1
      });
      
      this.saveProfile();
    }
  }

  /**
   * Apply learned preferences to a command
   */
  applyPreferences(command: ParsedCommand): ParsedCommand {
    const enhanced = { ...command };
    const params = { ...command.parameters };

    // Apply default chart type if not specified
    if (command.target === 'chart' && !params.chartType && this.profile.preferredChartType) {
      params.chartType = this.profile.preferredChartType;
    }

    // Apply default number format if not specified
    if (command.intent === 'format' && !params.numberFormat && this.profile.preferredNumberFormat) {
      params.numberFormat = this.profile.preferredNumberFormat;
    }

    // Apply default aggregation if not specified
    if ((command.target === 'pivot' || command.target === 'measure') && 
        !params.aggregation && this.profile.preferredAggregation) {
      params.aggregation = this.profile.preferredAggregation;
    }

    // Apply default table style if not specified
    if (command.target === 'table' && !params.style && this.profile.preferredTableStyle) {
      params.style = this.profile.preferredTableStyle;
    }

    // Apply learned column preferences
    if (params.columns && params.columns.length === 0) {
      const topColumns = this.getTopColumns(3);
      if (topColumns.length > 0) {
        params.columns = topColumns;
      }
    }

    enhanced.parameters = params;
    return enhanced;
  }

  /**
   * Suggest improvements based on learned patterns
   */
  suggestImprovements(command: ParsedCommand, locale: SupportedLocale = 'en'): string[] {
    const suggestions: string[] = [];
    const t = this.getTranslations(locale);

    // Check for known corrections
    const normalized = this.normalizeForLearning(command.originalText);
    const correction = this.profile.corrections.get(normalized);
    if (correction) {
      suggestions.push(t.didYouMean.replace('{command}', correction));
    }

    // Suggest based on frequently used patterns
    const similarPatterns = this.findSimilarPatterns(command.originalText);
    similarPatterns.forEach(pattern => {
      if (pattern.successRate > 0.8) {
        suggestions.push(t.suggestedAlternative.replace('{alternative}', pattern.preferredOutcome));
      }
    });

    // Suggest preferred options
    if (command.target === 'chart' && !command.parameters?.chartType && this.profile.preferredChartType) {
      suggestions.push(t.suggestChartType.replace('{type}', this.profile.preferredChartType));
    }

    return suggestions.slice(0, 3);
  }

  /**
   * Get personalized example commands based on user history
   */
  getPersonalizedExamples(target: CommandTarget, locale: SupportedLocale = 'en'): string[] {
    const t = this.getTranslations(locale);
    const examples: string[] = [];

    switch (target) {
      case 'chart':
        if (this.profile.preferredChartType) {
          examples.push(t.createChartOfType.replace('{type}', this.profile.preferredChartType));
        }
        break;
      case 'table':
        if (this.profile.preferredTableStyle) {
          examples.push(t.createTableWithStyle.replace('{style}', this.profile.preferredTableStyle));
        }
        break;
      case 'pivot':
        if (this.profile.preferredAggregation) {
          examples.push(t.createPivotWithAggregation.replace('{aggregation}', this.profile.preferredAggregation));
        }
        break;
    }

    return examples;
  }

  /**
   * Learn a preference value
   */
  private learnPreference(key: keyof UserLearningProfile, value: any): void {
    if (key === 'frequentlyUsedColumns' || key === 'commandPatterns' || key === 'corrections') {
      return; // These are handled separately
    }
    (this.profile as any)[key] = value;
  }

  /**
   * Learn column usage frequency
   */
  private learnColumnUsage(column: string): void {
    const normalized = column.toLowerCase().trim();
    const current = this.profile.frequentlyUsedColumns.get(normalized) || 0;
    this.profile.frequentlyUsedColumns.set(normalized, current + 1);
  }

  /**
   * Learn command patterns
   */
  private learnCommandPattern(command: ParsedCommand): void {
    const pattern = this.extractPattern(command.originalText);
    const existing = this.profile.commandPatterns.find(p => p.pattern === pattern);

    if (existing) {
      existing.usageCount++;
      existing.successRate = (existing.successRate * (existing.usageCount - 1) + 1) / existing.usageCount;
    } else {
      this.profile.commandPatterns.push({
        pattern,
        preferredOutcome: command.originalText,
        successRate: 1.0,
        usageCount: 1
      });
    }
  }

  /**
   * Extract pattern from command text
   */
  private extractPattern(text: string): string {
    return text
      .toLowerCase()
      .replace(/\b[a-z]+\d+(?::[a-z]+\d+)?\b/g, 'RANGE')
      .replace(/\b\d+\b/g, 'NUM')
      .replace(/\b\d{4}\b/g, 'YEAR')
      .trim();
  }

  /**
   * Normalize text for learning
   */
  private normalizeForLearning(text: string): string {
    return text
      .toLowerCase()
      .replace(/\b[a-z]+\d+(?::[a-z]+\d+)?\b/g, 'RANGE')
      .replace(/\b\d+\b/g, 'NUM')
      .trim();
  }

  /**
   * Find similar learned patterns
   */
  private findSimilarPatterns(text: string): LearningPattern[] {
    const normalized = this.normalizeForLearning(text);
    return this.profile.commandPatterns
      .filter(p => this.calculateSimilarity(p.pattern, normalized) > 0.7)
      .sort((a, b) => b.successRate - a.successRate);
  }

  /**
   * Calculate string similarity
   */
  private calculateSimilarity(a: string, b: string): number {
    if (a === b) return 1.0;
    const longer = a.length > b.length ? a : b;
    const shorter = a.length > b.length ? b : a;
    if (longer.length === 0) return 1.0;
    
    const distance = this.levenshteinDistance(longer, shorter);
    return (longer.length - distance) / longer.length;
  }

  /**
   * Levenshtein distance calculation
   */
  private levenshteinDistance(a: string, b: string): number {
    const matrix: number[][] = [];
    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }
    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        const cost = b[i - 1] === a[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j - 1] + cost
        );
      }
    }
    return matrix[b.length][a.length];
  }

  /**
   * Get top N most frequently used columns
   */
  private getTopColumns(n: number): string[] {
    return Array.from(this.profile.frequentlyUsedColumns.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, n)
      .map(([col]) => col);
  }

  /**
   * Load profile from storage
   */
  private loadProfile(): UserLearningProfile {
    try {
      const stored = localStorage.getItem(this.STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        return {
          ...parsed,
          frequentlyUsedColumns: new Map(Object.entries(parsed.frequentlyUsedColumns || {})),
          corrections: new Map(Object.entries(parsed.corrections || {}))
        };
      }
    } catch {
      // Ignore storage errors
    }
    return {
      frequentlyUsedColumns: new Map(),
      commandPatterns: [],
      corrections: new Map()
    };
  }

  /**
   * Save profile to storage
   */
  private saveProfile(): void {
    try {
      const toStore = {
        ...this.profile,
        frequentlyUsedColumns: Object.fromEntries(this.profile.frequentlyUsedColumns),
        corrections: Object.fromEntries(this.profile.corrections)
      };
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(toStore));
    } catch {
      // Ignore storage errors
    }
  }

  /**
   * Clear all learned preferences
   */
  clearProfile(): void {
    this.profile = {
      frequentlyUsedColumns: new Map(),
      commandPatterns: [],
      corrections: new Map()
    };
    localStorage.removeItem(this.STORAGE_KEY);
  }

  /**
   * Get translations
   */
  private getTranslations(locale: SupportedLocale) {
    const translations = {
      en: {
        didYouMean: 'Did you mean: "{command}"?',
        suggestedAlternative: 'Alternative: {alternative}',
        suggestChartType: 'Try using your preferred {type} chart',
        createChartOfType: 'Create a {type} chart',
        createTableWithStyle: 'Create table with {style} style',
        createPivotWithAggregation: 'Create pivot with {aggregation}'
      },
      ru: {
        didYouMean: 'Возможно, вы имели в виду: "{command}"?',
        suggestedAlternative: 'Альтернатива: {alternative}',
        suggestChartType: 'Попробуйте предпочитаемую диаграмму {type}',
        createChartOfType: 'Создать диаграмму {type}',
        createTableWithStyle: 'Создать таблицу со стилем {style}',
        createPivotWithAggregation: 'Создать сводную таблицу с агрегацией {aggregation}'
      }
    };
    return translations[locale];
  }
}

export default UserLearningEngine.getInstance();
