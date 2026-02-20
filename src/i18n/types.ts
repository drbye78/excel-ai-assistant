/**
 * Internationalization Types
 * 
 * Type definitions for multi-language support including
 * Excel function localization and UI translations.
 */

export type SupportedLocale = 
  | 'en'      // English
  | 'ru'      // Russian
  | 'es'      // Spanish
  | 'de'      // German
  | 'fr'      // French
  | 'zh'      // Chinese
  | 'ja'      // Japanese
  | 'pt'      // Portuguese
  | 'it'      // Italian
  | 'pl';     // Polish

export interface LocaleInfo {
  code: SupportedLocale;
  name: string;
  nativeName: string;
  direction: 'ltr' | 'rtl';
  excelCode: string;  // Excel's internal locale code
}

export interface ExcelSeparators {
  listSeparator: string;      // Usually comma or semicolon
  decimalSeparator: string;   // Usually dot or comma
  thousandsSeparator: string; // Usually comma or space
  arraySeparator: string;     // Usually comma or backslash
}

export interface LocaleConfig {
  locale: SupportedLocale;
  separators: ExcelSeparators;
  dateFormat: string;
  timeFormat: string;
  numberFormat: {
    decimalPlaces: number;
    useThousandsSeparator: boolean;
  };
}

// Excel function mapping: English -> Localized
export type FunctionMap = Map<string, string>;
export type ReverseFunctionMap = Map<string, string>;

export interface ExcelFunctionTranslations {
  [locale: string]: {
    functions: Record<string, string>;  // englishName -> localizedName
    reverse: Record<string, string>;    // localizedName -> englishName
  };
}

// UI Translation keys
export type TranslationKey = 
  // Common
  | 'common.ok'
  | 'common.cancel'
  | 'common.save'
  | 'common.delete'
  | 'common.edit'
  | 'common.close'
  | 'common.loading'
  | 'common.error'
  | 'common.success'
  | 'common.warning'
  | 'common.info'
  | 'common.yes'
  | 'common.no'
  | 'common.apply'
  | 'common.reset'
  | 'common.search'
  | 'common.filter'
  | 'common.sort'
  | 'common.copy'
  | 'common.paste'
  | 'common.undo'
  | 'common.redo'
  
  // Navigation
  | 'nav.home'
  | 'nav.chat'
  | 'nav.history'
  | 'nav.settings'
  | 'nav.help'
  | 'nav.formulas'
  | 'nav.analysis'
  | 'nav.recipes'
  | 'nav.batch'
  | 'nav.visual'
  | 'nav.powerQuery'
  | 'nav.dax'
  | 'nav.analytics'
  
  // Chat
  | 'chat.placeholder'
  | 'chat.send'
  | 'chat.clear'
  | 'chat.export'
  | 'chat.thinking'
  | 'chat.error.message'
  | 'chat.error.retry'
  | 'chat.suggestions.formula'
  | 'chat.suggestions.analysis'
  | 'chat.suggestions.format'
  | 'chat.welcome'
  | 'chat.noApiKey'
  | 'chat.actionsCompleted'
  | 'chat.actionFailed'
  | 'chat.suggested'
  | 'chat.role.user'
  | 'chat.role.system'
  | 'chat.role.assistant'
  | 'chat.confirm.title'
  | 'chat.confirm.message'
  
  // Formula Explainer
  | 'formula.input.label'
  | 'formula.input.placeholder'
  | 'formula.explain.button'
  | 'formula.optimize.button'
  | 'formula.complexity.label'
  | 'formula.result.label'
  | 'formula.breakdown.label'
  | 'formula.optimizations.label'
  | 'formula.noOptimizations'
  | 'formula.error.parse'
  | 'formula.error.unknown'
  
  // Data Analysis
  | 'analysis.title'
  | 'analysis.range.label'
  | 'analysis.run.button'
  | 'analysis.statistics.label'
  | 'analysis.trends.label'
  | 'analysis.anomalies.label'
  | 'analysis.correlations.label'
  | 'analysis.noData'
  | 'analysis.insights.label'
  
  // Settings
  | 'settings.title'
  | 'settings.language.label'
  | 'settings.language.description'
  | 'settings.language.autoDetected'
  | 'settings.theme.label'
  | 'settings.theme.light'
  | 'settings.theme.dark'
  | 'settings.theme.system'
  | 'settings.api.label'
  | 'settings.api.key'
  | 'settings.api.url'
  | 'settings.api.model'
  | 'settings.api.testConnection'
  | 'settings.api.testing'
  | 'settings.api.connectionSuccess'
  | 'settings.api.connectionError'
  | 'settings.api.loadedModels'
  | 'settings.api.enterKeyToLoad'
  | 'settings.scope.label'
  | 'settings.scope.enableWorkbook'
  | 'settings.scope.workbookDescription'
  | 'settings.scope.globalDescription'
  | 'settings.scope.whereToSave'
  | 'settings.scope.global'
  | 'settings.scope.workbook'
  | 'settings.advanced.label'
  | 'settings.advanced.show'
  | 'settings.temperature.label'
  | 'settings.temperature.description'
  | 'settings.maxTokens.label'
  | 'settings.maxTokens.description'
  | 'settings.save.global'
  | 'settings.save.workbook'
  | 'settings.privacy.label'
  | 'settings.privacy.analytics'
  | 'settings.privacy.storage'
  
  // Recipes
  | 'recipes.title'
  | 'recipes.create'
  | 'recipes.import'
  | 'recipes.export'
  | 'recipes.search'
  | 'recipes.categories.all'
  | 'recipes.categories.dataCleaning'
  | 'recipes.categories.formatting'
  | 'recipes.categories.analysis'
  | 'recipes.categories.charts'
  | 'recipes.apply'
  | 'recipes.preview'
  | 'recipes.share'
  
  // Batch Operations
  | 'batch.title'
  | 'batch.addOperation'
  | 'batch.run'
  | 'batch.clear'
  | 'batch.operations.count'
  | 'batch.operation.remove'
  | 'batch.settings.stopOnError'
  | 'batch.settings.enableUndo'
  
  // Visual Highlighting
  | 'visual.title'
  | 'visual.highlight.formulas'
  | 'visual.highlight.errors'
  | 'visual.highlight.duplicates'
  | 'visual.highlight.constants'
  | 'visual.highlight.precedents'
  | 'visual.highlight.dependents'
  | 'visual.clearAll'
  | 'visual.colors.label'
  
  // Power Query
  | 'pq.title'
  | 'pq.operations.label'
  | 'pq.generate.button'
  | 'pq.explain.button'
  | 'pq.apply.button'
  | 'pq.preview.label'
  | 'pq.operations.filter'
  | 'pq.operations.sort'
  | 'pq.operations.group'
  | 'pq.operations.merge'
  | 'pq.operations.pivot'
  
  // DAX
  | 'dax.title'
  | 'dax.measures.label'
  | 'dax.columns.label'
  | 'dax.explain.button'
  | 'dax.complexity.label'
  | 'dax.performance.label'
  | 'dax.model.label'
  | 'dax.noMeasures'
  
  // Analytics
  | 'analytics.title'
  | 'analytics.usage.label'
  | 'analytics.performance.label'
  | 'analytics.ai.label'
  | 'analytics.export.button'
  | 'analytics.timeRange'
  | 'analytics.timeRange.24h'
  | 'analytics.timeRange.7d'
  | 'analytics.timeRange.30d'
  | 'analytics.timeRange.90d'
  
  // Errors
  | 'error.generic'
  | 'error.network'
  | 'error.excel'
  | 'error.permission'
  | 'error.validation'
  | 'error.notFound'
  | 'error.timeout'
  | 'error.rateLimit'
  | 'error.ai'
  
  // Notifications
  | 'notification.operationSuccess'
  | 'notification.operationFailed'
  | 'notification.saved'
  | 'notification.deleted'
  | 'notification.copied'
  | 'notification.applied'
  
  // Confirmation dialogs
  | 'confirm.delete.title'
  | 'confirm.delete.message'
  | 'confirm.unsaved.title'
  | 'confirm.unsaved.message'
  
  // App
  | 'app.title'
  | 'app.skipToContent'
  | 'app.welcome'
  | 'app.configurePrompt'
  | 'app.configureSettings'
  | 'app.using'
  | 'app.notConfigured'
  | 'app.version';

export interface UITranslations {
  en: Record<TranslationKey, string>;
  [locale: string]: Partial<Record<TranslationKey, string>>;
}

// Language-specific number/date formatting
export interface FormatConfig {
  shortDate: string;
  longDate: string;
  shortTime: string;
  longTime: string;
  currency: string;
  number: string;
  percentage: string;
}

// AI context for different locales
export interface AIContextConfig {
  systemPrompt: string;
  formulaPrefix: string;
  explanationStyle: 'detailed' | 'concise' | 'technical';
  useLocalizedExamples: boolean;
}
