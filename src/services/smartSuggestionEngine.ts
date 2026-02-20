// Smart Suggestion Engine - Context-aware command suggestions
// Phase 3 Implementation - Intelligence & Optimization

import {
  NaturalLanguageCommandParser,
  NLContext,
  CommandIntent,
  CommandTarget,
  SupportedLocale
} from './naturalLanguageCommandParser';

export interface SuggestionCategory {
  id: string;
  title: string;
  suggestions: string[];
}

export interface ContextualSuggestion {
  text: string;
  icon?: string;
  category: string;
  priority: number;
}

export class SmartSuggestionEngine {
  private static instance: SmartSuggestionEngine;
  private parser: NaturalLanguageCommandParser;

  private constructor() {
    this.parser = NaturalLanguageCommandParser.getInstance();
  }

  static getInstance(): SmartSuggestionEngine {
    if (!SmartSuggestionEngine.instance) {
      SmartSuggestionEngine.instance = new SmartSuggestionEngine();
    }
    return SmartSuggestionEngine.instance;
  }

  /**
   * Generate contextual suggestions based on current selection and context
   */
  generateSuggestions(context: NLContext, locale: SupportedLocale = 'en'): ContextualSuggestion[] {
    const suggestions: ContextualSuggestion[] = [];

    // Selection-based suggestions
    if (context.selectedRange) {
      suggestions.push(...this.getRangeSuggestions(context, locale));
    }

    if (context.activeTable) {
      suggestions.push(...this.getTableSuggestions(context, locale));
    }

    if (context.selectedChart) {
      suggestions.push(...this.getChartSuggestions(context, locale));
    }

    if (context.selectedPivot) {
      suggestions.push(...this.getPivotSuggestions(context, locale));
    }

    // Data type based suggestions
    if (context.dataType) {
      suggestions.push(...this.getDataTypeSuggestions(context, locale));
    }

    // Recent activity based suggestions
    if (context.commandHistory && context.commandHistory.length > 0) {
      suggestions.push(...this.getRecentActivitySuggestions(context, locale));
    }

    // Sort by priority and return top 8
    return suggestions
      .sort((a, b) => b.priority - a.priority)
      .slice(0, 8);
  }

  /**
   * Get suggestions based on selected range
   */
  private getRangeSuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);
    const range = context.selectedRange;

    return [
      {
        text: t.formatAsCurrency.replace('{range}', range || ''),
        icon: '💰',
        category: 'formatting',
        priority: 90
      },
      {
        text: t.createChartFromSelection.replace('{range}', range || ''),
        icon: '📊',
        category: 'visualization',
        priority: 85
      },
      {
        text: t.applyConditionalFormatting,
        icon: '🎨',
        category: 'formatting',
        priority: 80
      },
      {
        text: t.sortAscending,
        icon: '🔼',
        category: 'sorting',
        priority: 75
      },
      {
        text: t.sortDescending,
        icon: '🔽',
        category: 'sorting',
        priority: 74
      },
      {
        text: t.copyToNewSheet,
        icon: '📋',
        category: 'organization',
        priority: 70
      }
    ];
  }

  /**
   * Get suggestions based on active table
   */
  private getTableSuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);
    const tableName = context.activeTable;

    return [
      {
        text: t.addTotalsRow.replace('{table}', tableName || ''),
        icon: '➕',
        category: 'table',
        priority: 95
      },
      {
        text: t.createPivotFromTable.replace('{table}', tableName || ''),
        icon: '🔄',
        category: 'analysis',
        priority: 90
      },
      {
        text: t.removeDuplicates,
        icon: '🧹',
        category: 'cleanup',
        priority: 85
      },
      {
        text: t.filterThisMonth,
        icon: '📅',
        category: 'filtering',
        priority: 80
      }
    ];
  }

  /**
   * Get suggestions based on selected chart
   */
  private getChartSuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);

    return [
      {
        text: t.addTrendline,
        icon: '📈',
        category: 'chart',
        priority: 90
      },
      {
        text: t.addDataLabels,
        icon: '🏷️',
        category: 'chart',
        priority: 85
      },
      {
        text: t.changeChartType,
        icon: '🔄',
        category: 'chart',
        priority: 80
      },
      {
        text: t.exportChartAsImage,
        icon: '🖼️',
        category: 'export',
        priority: 75
      }
    ];
  }

  /**
   * Get suggestions based on selected pivot
   */
  private getPivotSuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);

    return [
      {
        text: t.refreshPivot,
        icon: '🔄',
        category: 'pivot',
        priority: 90
      },
      {
        text: t.addSlicer,
        icon: '🔪',
        category: 'pivot',
        priority: 85
      },
      {
        text: t.changeAggregation,
        icon: '🔢',
        category: 'pivot',
        priority: 80
      }
    ];
  }

  /**
   * Get suggestions based on data type
   */
  private getDataTypeSuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);
    const suggestions: ContextualSuggestion[] = [];

    switch (context.dataType) {
      case 'numeric':
        suggestions.push(
          {
            text: t.calculateSummaryStats,
            icon: '📊',
            category: 'analysis',
            priority: 95
          },
          {
            text: t.createHistogram,
            icon: '📊',
            category: 'visualization',
            priority: 88
          },
          {
            text: t.addSparklines,
            icon: '✨',
            category: 'visualization',
            priority: 85
          },
          {
            text: t.findOutliers,
            icon: '🔍',
            category: 'analysis',
            priority: 80
          }
        );
        break;

      case 'date':
        suggestions.push(
          {
            text: t.groupByMonth,
            icon: '📅',
            category: 'grouping',
            priority: 95
          },
          {
            text: t.createTimelineChart,
            icon: '📈',
            category: 'visualization',
            priority: 90
          },
          {
            text: t.calculateYOYGrowth,
            icon: '📊',
            category: 'analysis',
            priority: 85
          }
        );
        break;

      case 'text':
        suggestions.push(
          {
            text: t.removeDuplicates,
            icon: '🧹',
            category: 'cleanup',
            priority: 90
          },
          {
            text: t.countUniqueValues,
            icon: '🔢',
            category: 'analysis',
            priority: 85
          },
          {
            text: t.textToColumns,
            icon: '✂️',
            category: 'transformation',
            priority: 80
          }
        );
        break;

      case 'mixed':
        suggestions.push(
          {
            text: t.splitTextNumbers,
            icon: '✂️',
            category: 'transformation',
            priority: 85
          },
          {
            text: t.cleanData,
            icon: '🧹',
            category: 'cleanup',
            priority: 80
          }
        );
        break;
    }

    // Large dataset warning
    if (context.rowCount && context.rowCount > 1000) {
      suggestions.push({
        text: t.largeDatasetWarning,
        icon: '⚠️',
        category: 'warning',
        priority: 100
      });
    }

    return suggestions;
  }

  /**
   * Get suggestions based on recent command history
   */
  private getRecentActivitySuggestions(context: NLContext, locale: SupportedLocale): ContextualSuggestion[] {
    const t = this.getTranslations(locale);
    const suggestions: ContextualSuggestion[] = [];
    const recentTargets = new Set(
      context.commandHistory?.slice(-3).map(cmd => {
        const parsed = this.parser.parseCommand(cmd);
        return parsed.target;
      }) || []
    );

    // Suggest related operations based on recent targets
    if (recentTargets.has('chart')) {
      suggestions.push(
        {
          text: t.addTrendline,
          icon: '📈',
          category: 'chart',
          priority: 82
        },
        {
          text: t.changeChartType,
          icon: '🔄',
          category: 'chart',
          priority: 78
        }
      );
    }

    if (recentTargets.has('pivot')) {
      suggestions.push(
        {
          text: t.createPivotChart,
          icon: '📊',
          category: 'pivot',
          priority: 82
        },
        {
          text: t.addCalculatedField,
          icon: '➕',
          category: 'pivot',
          priority: 78
        }
      );
    }

    if (recentTargets.has('table')) {
      suggestions.push(
        {
          text: t.formatAsTable,
          icon: '🎨',
          category: 'table',
          priority: 82
        }
      );
    }

    return suggestions;
  }

  /**
   * Get translations for the current locale
   */
  private getTranslations(locale: SupportedLocale) {
    const translations = {
      en: {
        formatAsCurrency: 'Format {range} as currency',
        createChartFromSelection: 'Create chart from {range}',
        applyConditionalFormatting: 'Apply conditional formatting',
        sortAscending: 'Sort ascending',
        sortDescending: 'Sort descending',
        copyToNewSheet: 'Copy to new worksheet',
        addTotalsRow: 'Add totals row to {table}',
        createPivotFromTable: 'Create pivot table from {table}',
        removeDuplicates: 'Remove duplicates',
        filterThisMonth: 'Filter for this month',
        addTrendline: 'Add trendline',
        addDataLabels: 'Add data labels',
        changeChartType: 'Change chart type',
        exportChartAsImage: 'Export as image',
        refreshPivot: 'Refresh pivot table',
        addSlicer: 'Add slicer',
        changeAggregation: 'Change aggregation function',
        calculateSummaryStats: 'Calculate summary statistics',
        createHistogram: 'Create histogram',
        addSparklines: 'Add sparklines',
        findOutliers: 'Find outliers',
        groupByMonth: 'Group by month',
        createTimelineChart: 'Create timeline chart',
        calculateYOYGrowth: 'Calculate year-over-year growth',
        countUniqueValues: 'Count unique values',
        textToColumns: 'Split text to columns',
        splitTextNumbers: 'Split text and numbers',
        cleanData: 'Clean data',
        largeDatasetWarning: 'Large dataset detected - consider filtering first',
        createPivotChart: 'Create pivot chart',
        addCalculatedField: 'Add calculated field',
        formatAsTable: 'Format as Excel table'
      },
      ru: {
        formatAsCurrency: 'Отформатировать {range} как валюту',
        createChartFromSelection: 'Создать диаграмму из {range}',
        applyConditionalFormatting: 'Применить условное форматирование',
        sortAscending: 'Сортировать по возрастанию',
        sortDescending: 'Сортировать по убыванию',
        copyToNewSheet: 'Копировать на новый лист',
        addTotalsRow: 'Добавить строку итогов к {table}',
        createPivotFromTable: 'Создать сводную таблицу из {table}',
        removeDuplicates: 'Удалить дубликаты',
        filterThisMonth: 'Отфильтровать за текущий месяц',
        addTrendline: 'Добавить линию тренда',
        addDataLabels: 'Добавить подписи данных',
        changeChartType: 'Изменить тип диаграммы',
        exportChartAsImage: 'Экспортировать как изображение',
        refreshPivot: 'Обновить сводную таблицу',
        addSlicer: 'Добавить срез',
        changeAggregation: 'Изменить функцию агрегации',
        calculateSummaryStats: 'Рассчитать сводную статистику',
        createHistogram: 'Создать гистограмму',
        addSparklines: 'Добавить спарклайны',
        findOutliers: 'Найти выбросы',
        groupByMonth: 'Сгруппировать по месяцам',
        createTimelineChart: 'Создать временную шкалу',
        calculateYOYGrowth: 'Рассчитать рост год к году',
        countUniqueValues: 'Посчитать уникальные значения',
        textToColumns: 'Разделить текст по столбцам',
        splitTextNumbers: 'Разделить текст и числа',
        cleanData: 'Очистить данные',
        largeDatasetWarning: 'Обнаружен большой набор данных - рассмотрите фильтрацию',
        createPivotChart: 'Создать сводную диаграмму',
        addCalculatedField: 'Добавить вычисляемое поле',
        formatAsTable: 'Форматировать как таблицу Excel'
      }
    };

    return translations[locale];
  }

  /**
   * Get quick actions based on current context
   */
  getQuickActions(context: NLContext, locale: SupportedLocale = 'en'): string[] {
    const t = this.getTranslations(locale);
    const actions: string[] = [];

    if (context.selectedRange) {
      actions.push(t.formatAsCurrency.replace('{range}', ''));
      actions.push(t.createChartFromSelection.replace('{range}', ''));
    }

    if (context.dataType === 'numeric') {
      actions.push(t.calculateSummaryStats);
    }

    if (context.activeTable) {
      actions.push(t.createPivotFromTable.replace('{table}', ''));
    }

    return actions.slice(0, 3);
  }
}

export default SmartSuggestionEngine.getInstance();
