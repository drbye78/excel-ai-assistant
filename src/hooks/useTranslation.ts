// Translation Hook - Easy access to translations in components
import { useState, useEffect, useCallback } from 'react';
import { t, tFormat, translations } from '@/i18n/translations';
import { SupportedLocale } from '@/i18n/types';
import SettingsService from '@/services/settingsService';

// Current locale state
let currentLocale: SupportedLocale = 'en';
let listeners: Set<(locale: SupportedLocale) => void> = new Set();

/**
 * Get current locale
 */
export function getCurrentLocale(): SupportedLocale {
  return currentLocale;
}

/**
 * Set current locale and notify listeners
 */
export function setCurrentLocale(locale: SupportedLocale): void {
  currentLocale = locale;
  SettingsService.setLocale(locale);
  listeners.forEach(callback => callback(locale));
}

/**
 * Subscribe to locale changes
 */
export function onLocaleChange(callback: (locale: SupportedLocale) => void): () => void {
  listeners.add(callback);
  return () => {
    listeners.delete(callback);
  };
}

/**
 * Hook to use translations in components
 */
export function useTranslation() {
  const [locale, setLocale] = useState<SupportedLocale>(currentLocale);

  useEffect(() => {
    // Try to get locale from settings service
    const savedLocale = SettingsService.getLocale();
    if (savedLocale) {
      currentLocale = savedLocale;
      setLocale(savedLocale);
    }

    // Subscribe to locale changes
    const unsubscribe = onLocaleChange(setLocale);
    return unsubscribe;
  }, []);

  const translate = useCallback((key: string): string => {
    return t(key as any, locale);
  }, [locale]);

  const translateFormat = useCallback((key: string, values: Record<string, string | number>): string => {
    return tFormat(key as any, values, locale);
  }, [locale]);

  const changeLocale = useCallback((newLocale: SupportedLocale) => {
    setCurrentLocale(newLocale);
    setLocale(newLocale);
  }, []);

  return {
    locale,
    t: translate,
    tFormat: translateFormat,
    setLocale: changeLocale
  };
}

// Convenience function for quick translations
export function translate(key: string, locale?: SupportedLocale): string {
  return t(key as any, locale || currentLocale);
}

export default useTranslation;
