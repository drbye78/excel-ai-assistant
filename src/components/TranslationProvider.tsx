// Translation Provider - Wraps app with translation context
import * as React from "react";
import { useState, useEffect, createContext, useContext } from "react";
import { SupportedLocale } from "@/i18n/types";
import { t, tFormat } from "@/i18n/translations";
import SettingsService from "@/services/settingsService";

// Context type
interface TranslationContextType {
  locale: SupportedLocale;
  setLocale: (locale: SupportedLocale) => void;
  t: (key: string) => string;
  tFormat: (key: string, values: Record<string, string | number>) => string;
}

// Create context
const TranslationContext = createContext<TranslationContextType | null>(null);

// Provider component
export const TranslationProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [locale, setLocaleState] = useState<SupportedLocale>("en");

  useEffect(() => {
    // Initialize locale from settings
    const initLocale = async () => {
      try {
        // Try to detect from Excel first
        const detectedLocale = await SettingsService.detectAndSetLocaleFromExcel();
        setLocaleState(detectedLocale);
      } catch (error) {
        // Fall back to browser locale
        const browserLocale = SettingsService.detectAndSetBrowserLocale();
        setLocaleState(browserLocale);
      }
    };

    initLocale();
  }, []);

  const setLocale = (newLocale: SupportedLocale) => {
    setLocaleState(newLocale);
    SettingsService.setLocale(newLocale);
  };

  const translate = (key: string): string => {
    return t(key as any, locale);
  };

  const translateFormat = (key: string, values: Record<string, string | number>): string => {
    return tFormat(key as any, values, locale);
  };

  return (
    <TranslationContext.Provider value={{ locale, setLocale, t: translate, tFormat: translateFormat }}>
      {children}
    </TranslationContext.Provider>
  );
};

// Hook to use translations
export const useTranslation = (): TranslationContextType => {
  const context = useContext(TranslationContext);
  if (!context) {
    throw new Error("useTranslation must be used within a TranslationProvider");
  }
  return context;
};

export default TranslationProvider;
