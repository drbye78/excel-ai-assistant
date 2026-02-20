/**
 * Locale Service
 *
 * Manages internationalization settings including:
 * - Current locale detection and switching
 * - Excel locale integration
 * - Locale persistence
 * - Formula localization/delocalization
 */

import {
  SupportedLocale,
  LocaleConfig,
  ExcelSeparators,
  LocaleInfo
} from './types';
import {
  excelFunctionTranslations,
  localeConfigs,
  getLocalizedFunctionName,
  getEnglishFunctionName,
  localizeFormula,
  delocalizeFormula
} from './excelFunctions';
import { t, tFormat, isSupportedLocale, getSupportedLocales } from './translations';
import { logger } from '../utils/logger';

const LOCALE_STORAGE_KEY = 'excel_ai_locale';
const DEFAULT_LOCALE: SupportedLocale = 'en';

class LocaleService {
  private static instance: LocaleService;
  private currentLocale: SupportedLocale = DEFAULT_LOCALE;
  private listeners: Set<(locale: SupportedLocale) => void> = new Set();

  private constructor() {
    this.loadSavedLocale();
  }

  static getInstance(): LocaleService {
    if (!LocaleService.instance) {
      LocaleService.instance = new LocaleService();
    }
    return LocaleService.instance;
  }

  // ==================== LOCALE MANAGEMENT ====================

  /**
   * Get the current locale
   */
  getCurrentLocale(): SupportedLocale {
    return this.currentLocale;
  }

  /**
   * Set the current locale
   */
  setLocale(locale: SupportedLocale): void {
    if (!isSupportedLocale(locale)) {
      logger.warn(`Unsupported locale: ${locale}. Falling back to ${DEFAULT_LOCALE}`);
      locale = DEFAULT_LOCALE;
    }

    if (this.currentLocale !== locale) {
      this.currentLocale = locale;
      this.saveLocale(locale);
      this.notifyListeners(locale);
      logger.info(`Locale changed to: ${locale}`);
    }
  }

  /**
   * Detect locale from Excel
   * Note: In Office.js, we can try to detect from Excel's settings
   */
  async detectLocaleFromExcel(): Promise<SupportedLocale | null> {
    try {
      // Office.js doesn't directly expose the Excel locale in a simple way
      // We can try to detect from the user's environment or use a heuristic
      // based on the decimal separator

      // For now, we'll check if there's any Office context available
      if (typeof Office !== 'undefined' && Office.context) {
        // Check display language
        const displayLanguage = Office.context.displayLanguage;
        if (displayLanguage) {
          const locale = this.mapOfficeLanguageToLocale(displayLanguage);
          if (locale) {
            return locale;
          }
        }
      }

      return null;
    } catch (error) {
      logger.error('Failed to detect Excel locale:', error);
      return null;
    }
  }

  /**
   * Map Office.js language code to our locale codes
   */
  private mapOfficeLanguageToLocale(displayLanguage: string): SupportedLocale | null {
    const languageMap: Record<string, SupportedLocale> = {
      'en-US': 'en',
      'en-GB': 'en',
      'en': 'en',
      'ru-RU': 'ru',
      'ru': 'ru',
      'es-ES': 'es',
      'es-MX': 'es',
      'es': 'es',
      'de-DE': 'de',
      'de': 'de',
      'fr-FR': 'fr',
      'fr-CA': 'fr',
      'fr': 'fr',
      'zh-CN': 'zh',
      'zh-TW': 'zh',
      'zh': 'zh',
      'ja-JP': 'ja',
      'ja': 'ja',
      'pt-BR': 'pt',
      'pt-PT': 'pt',
      'pt': 'pt',
      'it-IT': 'it',
      'it': 'it',
      'pl-PL': 'pl',
      'pl': 'pl'
    };

    return languageMap[displayLanguage] || null;
  }

  /**
   * Initialize locale detection and setup
   */
  async initialize(): Promise<void> {
    // First try to load saved locale
    const savedLocale = this.loadSavedLocale();

    if (savedLocale) {
      this.currentLocale = savedLocale;
      logger.info(`Loaded saved locale: ${savedLocale}`);
    } else {
      // Try to detect from Excel
      const detectedLocale = await this.detectLocaleFromExcel();
      if (detectedLocale) {
        this.currentLocale = detectedLocale;
        logger.info(`Detected Excel locale: ${detectedLocale}`);
        this.saveLocale(detectedLocale);
      } else {
        // Try browser language
        const browserLocale = this.detectBrowserLocale();
        if (browserLocale) {
          this.currentLocale = browserLocale;
          logger.info(`Using browser locale: ${browserLocale}`);
          this.saveLocale(browserLocale);
        }
      }
    }
  }

  /**
   * Detect locale from browser settings
   */
  private detectBrowserLocale(): SupportedLocale | null {
    const browserLang = navigator.language || (navigator as any).userLanguage;
    if (!browserLang) return null;

    // Check exact match first
    if (isSupportedLocale(browserLang)) {
      return browserLang;
    }

    // Check language code only (e.g., 'en' from 'en-US')
    const langCode = browserLang.split('-')[0].toLowerCase();
    if (isSupportedLocale(langCode)) {
      return langCode;
    }

    return null;
  }

  /**
   * Load saved locale from storage
   */
  private loadSavedLocale(): SupportedLocale | null {
    try {
      const saved = localStorage.getItem(LOCALE_STORAGE_KEY);
      if (saved && isSupportedLocale(saved)) {
        return saved;
      }
    } catch (error) {
      logger.warn('Failed to load saved locale:', error);
    }
    return null;
  }

  /**
   * Save locale to storage
   */
  private saveLocale(locale: SupportedLocale): void {
    try {
      localStorage.setItem(LOCALE_STORAGE_KEY, locale);
    } catch (error) {
      logger.warn('Failed to save locale:', error);
    }
  }

  // ==================== EVENT LISTENERS ====================

  /**
   * Subscribe to locale changes
   */
  onLocaleChange(callback: (locale: SupportedLocale) => void): () => void {
    this.listeners.add(callback);
    return () => {
      this.listeners.delete(callback);
    };
  }

  private notifyListeners(locale: SupportedLocale): void {
    this.listeners.forEach(callback => {
      try {
        callback(locale);
      } catch (error) {
        logger.error('Locale change listener error:', error);
      }
    });
  }

  // ==================== TRANSLATION HELPERS ====================

  /**
   * Translate a key to the current locale
   */
  translate(key: Parameters<typeof t>[0]): string {
    return t(key, this.currentLocale);
  }

  /**
   * Translate with variable interpolation
   */
  translateFormat(
    key: Parameters<typeof tFormat>[0],
    values: Parameters<typeof tFormat>[1]
  ): string {
    return tFormat(key, values, this.currentLocale);
  }

  // ==================== EXCEL FUNCTION LOCALIZATION ====================

  /**
   * Get localized Excel function name
   */
  getLocalizedFunction(englishName: string): string {
    return getLocalizedFunctionName(englishName, this.currentLocale);
  }

  /**
   * Get English Excel function name from localized
   */
  getEnglishFunction(localizedName: string): string {
    return getEnglishFunctionName(localizedName, this.currentLocale);
  }

  /**
   * Localize a formula (English -> Current locale)
   */
  localizeFormula(formula: string): string {
    return localizeFormula(formula, this.currentLocale);
  }

  /**
   * Delocalize a formula (Current locale -> English)
   */
  delocalizeFormula(formula: string): string {
    return delocalizeFormula(formula, this.currentLocale);
  }

  // ==================== LOCALE CONFIG ====================

  /**
   * Get locale configuration
   */
  getLocaleConfig(): LocaleConfig {
    const config = localeConfigs[this.currentLocale];

    return {
      locale: this.currentLocale,
      separators: config?.separators || localeConfigs.en.separators,
      dateFormat: config?.dateFormat || 'MM/DD/YYYY',
      timeFormat: 'HH:mm:ss',
      numberFormat: {
        decimalPlaces: 2,
        useThousandsSeparator: true
      }
    };
  }

  /**
   * Get Excel separators for current locale
   */
  getSeparators(): ExcelSeparators {
    return this.getLocaleConfig().separators;
  }

  /**
   * Get Excel internal locale code
   */
  getExcelLocaleCode(): string {
    return localeConfigs[this.currentLocale]?.excelCode || '1033';
  }

  // ==================== LOCALE INFO ====================

  /**
   * Get information about all supported locales
   */
  getSupportedLocalesInfo(): LocaleInfo[] {
    return getSupportedLocales().map(({ code, name, nativeName }) => ({
      code,
      name,
      nativeName,
      direction: 'ltr' as const,
      excelCode: localeConfigs[code]?.excelCode || '1033'
    }));
  }

  /**
   * Get current locale info
   */
  getCurrentLocaleInfo(): LocaleInfo {
    const info = this.getSupportedLocalesInfo().find(l => l.code === this.currentLocale);
    return info || {
      code: 'en',
      name: 'English',
      nativeName: 'English',
      direction: 'ltr',
      excelCode: '1033'
    };
  }

  // ==================== UTILITY METHODS ====================

  /**
   * Check if current locale uses comma as decimal separator
   */
  usesCommaDecimal(): boolean {
    return this.getSeparators().decimalSeparator === ',';
  }

  /**
   * Check if current locale uses semicolon as list separator
   */
  usesSemicolonListSeparator(): boolean {
    return this.getSeparators().listSeparator === ';';
  }

  /**
   * Format a number according to current locale
   */
  formatNumber(value: number, decimalPlaces: number = 2): string {
    const config = this.getLocaleConfig();
    const sep = config.separators;

    let formatted = value.toFixed(decimalPlaces);

    // Replace decimal separator
    formatted = formatted.replace('.', sep.decimalSeparator);

    // Add thousands separator
    if (config.numberFormat.useThousandsSeparator) {
      const parts = formatted.split(sep.decimalSeparator);
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, sep.thousandsSeparator);
      formatted = parts.join(sep.decimalSeparator);
    }

    return formatted;
  }

  /**
   * Parse a localized number string to a number
   */
  parseNumber(value: string): number {
    const sep = this.getSeparators();

    // Remove thousands separators
    let normalized = value.replace(new RegExp(`\\${sep.thousandsSeparator}`, 'g'), '');

    // Replace decimal separator with dot
    normalized = normalized.replace(sep.decimalSeparator, '.');

    return parseFloat(normalized);
  }

  /**
   * Format a date according to current locale
   */
  formatDate(date: Date): string {
    const config = this.getLocaleConfig();

    // Simple formatting - could be enhanced with a proper date library
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();

    return config.dateFormat
      .replace('DD', day)
      .replace('MM', month)
      .replace('YYYY', year.toString())
      .replace('M', String(date.getMonth() + 1))
      .replace('D', String(date.getDate()));
  }
}

// Export singleton instance
export const localeService = LocaleService.getInstance();

// Export convenient accessor functions
export const getLocale = () => localeService.getCurrentLocale();
export const setLocale = (locale: SupportedLocale) => localeService.setLocale(locale);
export const tCurrent = (key: Parameters<typeof t>[0]) => localeService.translate(key);
export const tFormatCurrent = (
  key: Parameters<typeof tFormat>[0],
  values: Parameters<typeof tFormat>[1]
) => localeService.translateFormat(key, values);

export default localeService;
