/**
 * Internationalization (i18n) Module
 * 
 * Central export for all internationalization functionality.
 * Provides support for multiple languages with Russian as primary focus.
 */

// Types
export type {
  SupportedLocale,
  LocaleInfo,
  ExcelSeparators,
  LocaleConfig,
  TranslationKey,
  UITranslations,
  FormatConfig,
  AIContextConfig
} from './types';

// Excel Function Translations
export {
  excelFunctionTranslations,
  localeConfigs,
  getLocalizedFunctionName,
  getEnglishFunctionName,
  isLocalizedFunctionName,
  localizeFormula,
  delocalizeFormula
} from './excelFunctions';

// UI Translations
export {
  translations,
  t,
  tFormat,
  isSupportedLocale,
  getSupportedLocales
} from './translations';

// Locale Service
export {
  localeService,
  getLocale,
  setLocale,
  tCurrent,
  tFormatCurrent
} from './localeService';

// Re-export default
export { default } from './localeService';
