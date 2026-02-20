import * as React from "react";
import { useState, useEffect } from "react";
import { Chat } from "@components/Chat";
import { Settings } from "@components/Settings";
import { ErrorBoundary } from "@components/ErrorBoundary";
import { TranslationProvider, useTranslation } from "@components/TranslationProvider";
import { AISettings } from "@/types";
import { SupportedLocale } from "@/i18n/types";
import { lightTheme } from "@/theme";
import aiService from "@/services/aiService";
import settingsService from "@/services/settingsService";
import logger, { LogLevel } from "@/utils/logger";
import {
  Stack,
  DefaultButton,
  Text,
  Spinner
} from "@fluentui/react";

// Destructure theme tokens for cleaner usage
const { colors, spacing } = lightTheme;

// Inner component that uses translation hook
const AppContent: React.FC = () => {
  const { t, locale, setLocale } = useTranslation();
  const [activeTab, setActiveTab] = useState("chat");
  const [settings, setSettings] = useState<AISettings>({
    apiUrl: "https://openrouter.ai/api/v1",
    apiKey: "",
    model: "openai/gpt-4",
    temperature: 0.7,
    maxTokens: 4000
  });
  const [hasSettings, setHasSettings] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [settingsScope, setSettingsScope] = useState<{ type: string; source: string }>({
    type: 'default',
    source: 'Built-in defaults'
  });

  // Initialize services and load persisted settings on mount
  useEffect(() => {
    const initializeApp = async () => {
      try {
        // Configure logger with server URL
        logger.configure({
          serverUrl: window.location.origin.includes('localhost') 
            ? 'http://localhost:3001' 
            : window.location.origin,
          batchSize: 10,
          flushInterval: 5000,
          enableServerLogging: true,
          minServerLogLevel: LogLevel.INFO
        });
        logger.setLevel(LogLevel.INFO);
        
        logger.info('Application initializing');
        
        // Initialize AI service (loads persisted settings)
        await aiService.initialize();
        
        // Detect and set locale from Excel first time
        const detectedLocale = await settingsService.detectAndSetLocaleFromExcel();
        setLocale(detectedLocale);
        
        // Get loaded settings
        const loadedSettings = aiService.getSettings();
        setSettings(loadedSettings);
        
        // Get settings scope info
        const scope = aiService.getSettingsScope();
        setSettingsScope(scope);
        
        // Check if we have valid API key
        if (loadedSettings.apiKey) {
          setHasSettings(true);
        }
        
        logger.info('Application initialized successfully', { 
          locale: detectedLocale,
          hasApiKey: !!loadedSettings.apiKey,
          model: loadedSettings.model
        });
      } catch (error) {
        logger.error('Failed to initialize app', undefined, error instanceof Error ? error : new Error(String(error)));
        // Use browser locale as fallback
        const browserLocale = settingsService.detectAndSetBrowserLocale();
        setLocale(browserLocale);
      } finally {
        setIsLoading(false);
      }
    };

    initializeApp();
  }, []);

  const handleSaveSettings = async (
    newSettings: AISettings, 
    scope: 'global' | 'workbook' = 'global',
    newLocale?: SupportedLocale
  ) => {
    try {
      // Update settings with persistence
      await aiService.updateSettings(newSettings, true, scope);
      
      // Update locale if changed
      if (newLocale && newLocale !== locale) {
        settingsService.setLocale(newLocale);
        setLocale(newLocale);
      }
      
      // Reload settings to get merged result
      const updatedSettings = aiService.getSettings();
      setSettings(updatedSettings);
      
      // Update scope info
      const scopeInfo = aiService.getSettingsScope();
      setSettingsScope(scopeInfo);
      
      setHasSettings(true);
      setActiveTab("chat");
    } catch (error) {
      logger.error('Failed to save settings', { scope }, error instanceof Error ? error : new Error(String(error)));
      // Still update local state
      setSettings(newSettings);
      setHasSettings(true);
      setActiveTab("chat");
    }
  };

  const handleCancelSettings = () => {
    if (hasSettings) {
      setActiveTab("chat");
    }
  };

  const handleSwitchToSettings = () => {
    // Refresh settings scope when opening settings
    const scope = aiService.getSettingsScope();
    setSettingsScope(scope);
    setActiveTab("settings");
  };

  // Show loading spinner while initializing
  if (isLoading) {
    return (
      <Stack 
        horizontalAlign="center" 
        verticalAlign="center" 
        styles={{ root: { height: "100vh" } }}
        aria-live="polite"
        aria-busy="true"
      >
        <Spinner label={t('common.loading')} labelPosition="bottom" />
      </Stack>
    );
  }

  return (
    <Stack 
        styles={{ root: { height: "100vh", display: "flex", flexDirection: "column" } }}
        role="application"
        aria-label="Excel AI Assistant"
      >
      {/* Skip link for accessibility */}
      <a 
        href="#main-content" 
        className="skip-link"
        style={{
          position: 'absolute',
          left: '-9999px',
          top: 'auto',
          width: '1px',
          height: '1px',
          overflow: 'hidden',
          zIndex: 9999
        }}
        onFocus={(e) => {
          e.currentTarget.style.position = 'fixed';
          e.currentTarget.style.left = '10px';
          e.currentTarget.style.top = '10px';
          e.currentTarget.style.width = 'auto';
          e.currentTarget.style.height = 'auto';
          e.currentTarget.style.padding = '10px 20px';
          e.currentTarget.style.background = colors.brand.primary;
          e.currentTarget.style.color = colors.text.onBrand;
          e.currentTarget.style.textDecoration = 'none';
          e.currentTarget.style.borderRadius = '4px';
        }}
        onBlur={(e) => {
          e.currentTarget.style.position = 'absolute';
          e.currentTarget.style.left = '-9999px';
          e.currentTarget.style.width = '1px';
          e.currentTarget.style.height = '1px';
        }}
      >
        {t('app.skipToContent')}
      </a>

      {/* Header */}
      <Stack
        as="header"
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        aria-label="Application header"
        styles={{
          root: {
            padding: `${spacing.sm} ${spacing.md}`,
            backgroundColor: colors.brand.primary,
            color: colors.text.onBrand,
            borderBottom: `1px solid ${colors.brand.primaryDark}`
          }
        }}
      >
        <Text 
          variant="large" 
          styles={{ root: { color: colors.text.onBrand, fontWeight: 600 } }}
        >
          🤖 {t('app.title')}
        </Text>
        <DefaultButton
          text={t('nav.settings')}
          aria-label={t('nav.settings')}
          onClick={handleSwitchToSettings}
          styles={{
            root: {
              backgroundColor: 'transparent',
              color: colors.text.onBrand,
              borderColor: colors.text.onBrand
            },
            rootHovered: {
              backgroundColor: 'rgba(255,255,255,0.1)',
              color: colors.text.onBrand
            },
            rootPressed: {
              backgroundColor: 'rgba(255,255,255,0.2)',
              color: colors.text.onBrand
            }
          }}
        />
      </Stack>

      {/* Main Content */}
      <main 
        id="main-content" 
        style={{ flex: 1, overflow: "hidden" }}
        role="main"
        aria-label="Main content area"
      >
        <ErrorBoundary section="main content">
          {activeTab === "chat" && (
            <div style={{ height: "100%" }} role="region" aria-label="Chat interface">
              {!hasSettings ? (
                <Stack
                  horizontalAlign="center"
                  verticalAlign="center"
                  styles={{ root: { height: "100%", padding: spacing.lg } }}
                  role="region"
                  aria-label="Welcome screen"
                >
                  <Text variant="large" role="heading" aria-level={1}>
                    {t('app.welcome')}
                  </Text>
                  <Text styles={{ root: { margin: `${spacing.md} 0 ${spacing.lg} 0` } }}>
                    {t('app.configurePrompt')}
                  </Text>
                  <DefaultButton
                    text={t('app.configureSettings')}
                    primary
                    onClick={() => setActiveTab("settings")}
                    aria-label={t('app.configureSettings')}
                  />
                </Stack>
              ) : (
                <ErrorBoundary section="chat">
                  <Chat settings={settings} />
                </ErrorBoundary>
              )}
            </div>
          )}

          {activeTab === "settings" && (
            <ErrorBoundary section="settings">
              <Settings
                currentSettings={settings}
                settingsScope={settingsScope}
                currentLocale={locale}
                onSave={handleSaveSettings}
                onCancel={handleCancelSettings}
              />
            </ErrorBoundary>
          )}
        </ErrorBoundary>
      </main>

      {/* Status Bar */}
      <Stack
        as="footer"
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        role="contentinfo"
        aria-label="Application status"
        styles={{
          root: {
            padding: `${spacing.xs} ${spacing.md}`,
            backgroundColor: colors.neutral.gray20,
            borderTop: `1px solid ${colors.neutral.gray40}`,
            fontSize: '11px',
            color: colors.neutral.gray80
          }
        }}
      >
        <Text aria-live="polite">
          {hasSettings ? (
            <span>
              ✓ {t('app.using')}: {settings.model}
              <span style={{ opacity: 0.7, marginLeft: spacing.sm }}>
                ({settingsScope.source})
              </span>
            </span>
          ) : (
            <span>⚠ {t('app.notConfigured')}</span>
          )}
        </Text>
        <Text>{t('app.version')}</Text>
      </Stack>
      </Stack>
  );
};

// Main App component that provides translation context
export const App: React.FC = () => {
  return (
    <TranslationProvider>
      <AppContent />
    </TranslationProvider>
  );
};

export default App;
