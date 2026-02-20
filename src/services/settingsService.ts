// Settings Service - Persistent Configuration Management
// Supports both global (localStorage) and per-workbook (Office Settings) configurations

// Type references for Office.js
/// <reference types="@types/office-js" />
/// <reference types="@types/office-runtime" />

import { AISettings } from "@/types";
import { SupportedLocale } from "@/i18n/types";
import { logger } from "../utils/logger";

export interface AppSettings extends AISettings {
  // Global settings stored in localStorage
  isGlobalSettings: boolean;
  // Per-workbook settings stored in Office document
  isPerWorkbookSettings: boolean;
  // UI locale
  locale?: SupportedLocale;
}

export interface SettingsScope {
  type: 'global' | 'workbook' | 'default';
  source: string; // Description of where settings came from
}

// Default settings when nothing is configured
const DEFAULT_SETTINGS: AISettings = {
  apiUrl: "https://openrouter.ai/api/v1",
  apiKey: "",
  model: "openai/gpt-4",
  temperature: 0.7,
  maxTokens: 4000,
  systemPrompt: undefined
};

// Keys for localStorage
const GLOBAL_STORAGE_KEY = 'excel_ai_global_settings';
const WORKBOOK_OVERRIDE_KEY = 'excel_ai_workbook_override';
const SETTINGS_SCOPE_KEY = 'excel_ai_settings_scope';

// Office Settings keys (stored in workbook)
const WORKBOOK_KEYS = {
  API_URL: 'ai_apiUrl',
  API_KEY: 'ai_apiKey',
  MODEL: 'ai_model',
  TEMPERATURE: 'ai_temperature',
  MAX_TOKENS: 'ai_maxTokens',
  SYSTEM_PROMPT: 'ai_systemPrompt'
};

export class SettingsService {
  private static instance: SettingsService;
  private currentSettings: AppSettings | null = null;
  private currentScope: SettingsScope = { type: 'default', source: 'Built-in defaults' };
  private isInitialized: boolean = false;
  private initPromise: Promise<void> | null = null;

  private constructor() {}

  static getInstance(): SettingsService {
    if (!SettingsService.instance) {
      SettingsService.instance = new SettingsService();
    }
    return SettingsService.instance;
  }

  /**
   * Initialize settings - load from storage
   * Must be called when Office.js is ready
   */
  async initialize(): Promise<void> {
    if (this.isInitialized) return;
    if (this.initPromise) return this.initPromise;

    this.initPromise = this.doInitialize();
    await this.initPromise;
    this.isInitialized = true;
  }

  private async doInitialize(): Promise<void> {
    // First load global settings from localStorage
    const globalSettings = this.loadGlobalSettings();

    // Check if user has workbook override enabled
    const useWorkbookOverride = this.getWorkbookOverrideEnabled();

    if (useWorkbookOverride) {
      // Try to load workbook-specific settings
      const workbookSettings = await this.loadWorkbookSettings();
      
      if (workbookSettings && workbookSettings.apiKey) {
        // Workbook settings take priority when override is enabled
        this.currentSettings = {
          ...workbookSettings,
          isGlobalSettings: false,
          isPerWorkbookSettings: true
        };
        this.currentScope = { 
          type: 'workbook', 
          source: 'Current workbook settings (override enabled)' 
        };
        return;
      }
    }

    // Fall back to global settings
    if (globalSettings) {
      this.currentSettings = {
        ...globalSettings,
        isGlobalSettings: true,
        isPerWorkbookSettings: false
      };
      this.currentScope = { 
        type: 'global', 
        source: 'Global settings (localStorage)' 
      };
      return;
    }

    // Use defaults
    this.currentSettings = {
      ...DEFAULT_SETTINGS,
      isGlobalSettings: false,
      isPerWorkbookSettings: false
    };
    this.currentScope = { 
      type: 'default', 
      source: 'Built-in defaults' 
    };
  }

  /**
   * Get current settings (with fallback to defaults)
   */
  getSettings(): AppSettings {
    if (!this.currentSettings) {
      // Synchronous fallback - should rarely happen
      return {
        ...DEFAULT_SETTINGS,
        isGlobalSettings: false,
        isPerWorkbookSettings: false
      };
    }
    return this.currentSettings;
  }

  /**
   * Get current settings scope info
   */
  getScope(): SettingsScope {
    return this.currentScope;
  }

  /**
   * Update settings with automatic persistence
   * @param settings - New settings to apply
   * @param scope - Where to save: 'global', 'workbook', or 'auto' (auto = current scope)
   */
  async updateSettings(settings: Partial<AISettings>, scope: 'global' | 'workbook' | 'auto' = 'auto'): Promise<void> {
    const targetScope = scope === 'auto' ? this.currentScope.type : scope;

    if (targetScope === 'workbook') {
      await this.saveWorkbookSettings(settings);
      // Reload to get merged settings
      await this.reloadSettings();
    } else {
      this.saveGlobalSettings(settings);
      // Reload to get merged settings
      await this.reloadSettings();
    }
  }

  /**
   * Enable/disable workbook-specific settings override
   */
  setWorkbookOverrideEnabled(enabled: boolean): void {
    localStorage.setItem(WORKBOOK_OVERRIDE_KEY, enabled ? 'true' : 'false');
  }

  /**
   * Check if workbook override is enabled
   */
  getWorkbookOverrideEnabled(): boolean {
    return localStorage.getItem(WORKBOOK_OVERRIDE_KEY) === 'true';
  }

  /**
   * Reload settings from storage
   */
  async reloadSettings(): Promise<void> {
    this.isInitialized = false;
    this.initPromise = null;
    await this.initialize();
  }

  // ==================== Global Settings (localStorage) ====================

  /**
   * Load global settings from localStorage
   */
  private loadGlobalSettings(): AISettings | null {
    try {
      const stored = localStorage.getItem(GLOBAL_STORAGE_KEY);
      if (stored) {
        const parsed = JSON.parse(stored);
        // Validate required fields
        if (parsed.apiUrl && parsed.apiKey) {
          return parsed;
        }
      }
    } catch (error) {
      logger.warn('Failed to load global settings', { error });
    }
    return null;
  }

  /**
   * Save global settings to localStorage
   */
  private saveGlobalSettings(settings: Partial<AISettings>): void {
    try {
      const current = this.loadGlobalSettings() || { ...DEFAULT_SETTINGS };
      const updated = { ...current, ...settings };
      localStorage.setItem(GLOBAL_STORAGE_KEY, JSON.stringify(updated));
    } catch (error) {
      logger.error('Failed to save global settings', { error, settings });
      throw new Error('Failed to save settings to localStorage');
    }
  }

  /**
   * Clear global settings
   */
  clearGlobalSettings(): void {
    localStorage.removeItem(GLOBAL_STORAGE_KEY);
  }

  // ==================== Workbook Settings (Office API) ====================

  /**
   * Load settings from current workbook
   */
  private async loadWorkbookSettings(): Promise<AISettings | null> {
    if (!this.isOfficeApiAvailable()) {
      logger.warn('Office API not available');
      return null;
    }

    try {
      const settings = Office.context.document.settings;
      const result: Partial<AISettings> = {};

      // Load each setting
      const apiUrl = settings.get(WORKBOOK_KEYS.API_URL);
      const apiKey = settings.get(WORKBOOK_KEYS.API_KEY);
      const model = settings.get(WORKBOOK_KEYS.MODEL);
      const temperature = settings.get(WORKBOOK_KEYS.TEMPERATURE);
      const maxTokens = settings.get(WORKBOOK_KEYS.MAX_TOKENS);
      const systemPrompt = settings.get(WORKBOOK_KEYS.SYSTEM_PROMPT);

      // Only return if we have at least API key
      if (!apiKey) return null;

      if (apiUrl) result.apiUrl = apiUrl;
      if (apiKey) result.apiKey = apiKey;
      if (model) result.model = model;
      if (temperature !== undefined) result.temperature = temperature;
      if (maxTokens !== undefined) result.maxTokens = maxTokens;
      if (systemPrompt) result.systemPrompt = systemPrompt;

      return result as AISettings;
    } catch (error) {
      logger.warn('Failed to load workbook settings', { error });
      return null;
    }
  }

  /**
   * Save settings to current workbook
   */
  private async saveWorkbookSettings(settings: Partial<AISettings>): Promise<void> {
    if (!this.isOfficeApiAvailable()) {
      throw new Error('Office API not available');
    }

    return new Promise((resolve, reject) => {
      try {
        const docSettings = Office.context.document.settings;

        // Save each setting
        if (settings.apiUrl !== undefined) {
          docSettings.set(WORKBOOK_KEYS.API_URL, settings.apiUrl);
        }
        if (settings.apiKey !== undefined) {
          docSettings.set(WORKBOOK_KEYS.API_KEY, settings.apiKey);
        }
        if (settings.model !== undefined) {
          docSettings.set(WORKBOOK_KEYS.MODEL, settings.model);
        }
        if (settings.temperature !== undefined) {
          docSettings.set(WORKBOOK_KEYS.TEMPERATURE, settings.temperature);
        }
        if (settings.maxTokens !== undefined) {
          docSettings.set(WORKBOOK_KEYS.MAX_TOKENS, settings.maxTokens);
        }
        if (settings.systemPrompt !== undefined) {
          docSettings.set(WORKBOOK_KEYS.SYSTEM_PROMPT, settings.systemPrompt);
        }

        // Save to document
        docSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error('Failed to save workbook settings: ' + result.error.message));
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Clear workbook settings
   */
  async clearWorkbookSettings(): Promise<void> {
    if (!this.isOfficeApiAvailable()) return;

    return new Promise((resolve, reject) => {
      try {
        const docSettings = Office.context.document.settings;

        // Remove each setting
        Object.values(WORKBOOK_KEYS).forEach(key => {
          docSettings.remove(key);
        });

        docSettings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error('Failed to clear workbook settings: ' + result.error.message));
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Check if workbook has settings saved
   */
  async hasWorkbookSettings(): Promise<boolean> {
    if (!this.isOfficeApiAvailable()) return false;
    
    try {
      const settings = Office.context.document.settings;
      const apiKey = settings.get(WORKBOOK_KEYS.API_KEY) as string | false;
      return !!apiKey && typeof apiKey === 'string';
    } catch {
      return false;
    }
  }

  // ==================== Utilities ====================

  /**
   * Check if Office API is available
   */
  private isOfficeApiAvailable(): boolean {
    return typeof Office !== 'undefined' && 
           Office.context && 
           Office.context.document && 
           Office.context.document.settings;
  }

  /**
   * Export settings (without API key for sharing)
   */
  exportSettings(): Partial<AISettings> & { hasApiKey: boolean } {
    const settings = this.getSettings();
    return {
      apiUrl: settings.apiUrl,
      model: settings.model,
      temperature: settings.temperature,
      maxTokens: settings.maxTokens,
      systemPrompt: settings.systemPrompt,
      hasApiKey: !!settings.apiKey
    };
  }

  /**
   * Get all available model suggestions
   */
  getModelSuggestions(): string[] {
    return [
      // OpenAI
      'openai/gpt-4',
      'openai/gpt-4-turbo',
      'openai/gpt-3.5-turbo',
      // Anthropic
      'anthropic/claude-3-opus-20240229',
      'anthropic/claude-3-sonnet-20240229',
      'anthropic/claude-3-haiku-20240307',
      // Meta
      'meta-llama/llama-3-70b-instruct',
      'meta-llama/llama-3-8b-instruct',
      // Mistral
      'mistralai/mixtral-8x7b-instruct',
      'mistralai/mistral-7b-instruct',
      // Google
      'google/gemini-pro-1.5',
      // Cohere
      'cohere/command-r-plus',
      'cohere/command-r'
    ];
  }

  // ==================== LOCALE ====================

  /**
   * Detect locale from Excel and set as default if not already configured
   */
  async detectAndSetLocaleFromExcel(): Promise<SupportedLocale> {
    try {
      if (typeof Office !== 'undefined' && Office.context) {
        const displayLanguage = Office.context.displayLanguage;
        if (displayLanguage) {
          // Map Office language to our supported locales
          const locale = this.mapOfficeLanguageToLocale(displayLanguage);
          if (locale) {
            // Only set if no locale is saved yet
            const currentLocale = this.getLocale();
            if (!currentLocale) {
              this.setLocale(locale);
              return locale;
            }
            return currentLocale;
          }
        }
      }
    } catch (error) {
      logger.warn('Failed to detect Excel locale', { error });
    }
    
    // Fall back to browser locale
    return this.detectAndSetBrowserLocale();
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
   * Detect locale from browser and set as default
   */
  detectAndSetBrowserLocale(): SupportedLocale {
    const browserLang = navigator.language || (navigator as any).userLanguage || 'en';
    
    // Try exact match first
    let locale = browserLang as SupportedLocale;
    if (this.isValidLocale(locale)) {
      this.setLocale(locale);
      return locale;
    }
    
    // Try language code only (e.g., 'en' from 'en-US')
    const langCode = browserLang.split('-')[0].toLowerCase();
    if (this.isValidLocale(langCode)) {
      locale = langCode as SupportedLocale;
      this.setLocale(locale);
      return locale;
    }
    
    // Default to English
    this.setLocale('en');
    return 'en';
  }

  /**
   * Check if locale is valid
   */
  private isValidLocale(locale: string): boolean {
    const validLocales: SupportedLocale[] = ['en', 'ru', 'es', 'de', 'fr', 'zh', 'ja', 'pt', 'it', 'pl'];
    return validLocales.includes(locale as SupportedLocale);
  }

  /**
   * Get current locale
   */
  getLocale(): SupportedLocale | undefined {
    return this.currentSettings?.locale;
  }

  /**
   * Set locale (saves to localStorage)
   */
  setLocale(locale: SupportedLocale): void {
    // Save to localStorage
    try {
      const current = this.loadGlobalSettings() || { apiUrl: '', apiKey: '', model: '', temperature: 0.7, maxTokens: 0 };
      (current as any).locale = locale;
      localStorage.setItem(GLOBAL_STORAGE_KEY, JSON.stringify(current));
    } catch (error) {
      logger.warn('Failed to save locale', { error, locale });
    }
    
    // Update current settings if loaded
    if (this.currentSettings) {
      this.currentSettings.locale = locale;
    }
  }

  /**
   * Get available locales for UI
   */
  getAvailableLocales(): Array<{ code: SupportedLocale; name: string; nativeName: string }> {
    return [
      { code: 'en', name: 'English', nativeName: 'English' },
      { code: 'ru', name: 'Russian', nativeName: 'Русский' },
      { code: 'es', name: 'Spanish', nativeName: 'Español' },
      { code: 'de', name: 'German', nativeName: 'Deutsch' },
      { code: 'fr', name: 'French', nativeName: 'Français' },
      { code: 'zh', name: 'Chinese', nativeName: '中文' },
      { code: 'ja', name: 'Japanese', nativeName: '日本語' },
      { code: 'pt', name: 'Portuguese', nativeName: 'Português' },
      { code: 'it', name: 'Italian', nativeName: 'Italiano' },
      { code: 'pl', name: 'Polish', nativeName: 'Polski' }
    ];
  }
}

export default SettingsService.getInstance();
