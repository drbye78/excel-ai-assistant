import * as React from "react";
import { useState, useEffect, useMemo, useCallback } from "react";
import { AISettings } from "@/types";
import { SupportedLocale } from "@/i18n/types";
import { lightTheme } from "@/theme";
import { useTranslation } from "./TranslationProvider";
import AIService from "@/services/aiService";
import SettingsService from "@/services/settingsService";
import OpenRouterService, { OpenRouterModel } from "@/services/openRouterService";
import { logger } from "@/utils/logger";
import {
  Stack,
  TextField,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Toggle,
  Separator,
  Text,
  MessageBar,
  MessageBarType,
  Link,
  Spinner,
  ComboBox,
  IComboBox,
  IComboBoxOption,
  Label
} from "@fluentui/react";

// Theme tokens
const { colors, spacing } = lightTheme;

interface SettingsProps {
  currentSettings?: AISettings;
  settingsScope?: { type: string; source: string };
  currentLocale?: SupportedLocale;
  onSave: (settings: AISettings, scope?: 'global' | 'workbook', locale?: SupportedLocale) => void;
  onCancel: () => void;
}

type ApiProvider = "openai" | "azure" | "openrouter" | "local" | "custom";

interface ProviderConfig {
  name: string;
  defaultUrl: string;
  requiresKey: boolean;
  keyPlaceholder: string;
  keyDescription: string;
  models: IDropdownOption[];
  supportsCustomModel: boolean;
  helpUrl?: string;
}

const providers: Record<ApiProvider, ProviderConfig> = {
  openai: {
    name: "OpenAI",
    defaultUrl: "https://api.openai.com/v1",
    requiresKey: true,
    keyPlaceholder: "sk-...",
    keyDescription: "Your OpenAI API key from platform.openai.com",
    models: [
      { key: "gpt-4o", text: "GPT-4o (Recommended)" },
      { key: "gpt-4o-mini", text: "GPT-4o Mini (Faster)" },
      { key: "gpt-4-turbo", text: "GPT-4 Turbo" },
      { key: "gpt-4", text: "GPT-4" },
      { key: "gpt-3.5-turbo", text: "GPT-3.5 Turbo" }
    ],
    supportsCustomModel: false,
    helpUrl: "https://platform.openai.com/api-keys"
  },
  azure: {
    name: "Azure OpenAI",
    defaultUrl: "https://your-resource.openai.azure.com/openai/deployments/your-deployment",
    requiresKey: true,
    keyPlaceholder: "your-azure-api-key",
    keyDescription: "Your Azure OpenAI API key",
    models: [
      { key: "gpt-4", text: "GPT-4" },
      { key: "gpt-4-32k", text: "GPT-4 32K" },
      { key: "gpt-35-turbo", text: "GPT-3.5 Turbo" }
    ],
    supportsCustomModel: true,
    helpUrl: "https://portal.azure.com"
  },
  openrouter: {
    name: "OpenRouter",
    defaultUrl: "https://openrouter.ai/api/v1",
    requiresKey: true,
    keyPlaceholder: "sk-or-...",
    keyDescription: "Your OpenRouter API key from openrouter.ai/keys",
    models: [],
    supportsCustomModel: true,
    helpUrl: "https://openrouter.ai/keys"
  },
  local: {
    name: "Local Model (LM Studio/Ollama)",
    defaultUrl: "http://localhost:1234/v1",
    requiresKey: false,
    keyPlaceholder: "not-required",
    keyDescription: "API key is not required for local models",
    models: [
      { key: "local-model", text: "Local Model (Auto-detected)" },
      { key: "llama2", text: "Llama 2" },
      { key: "mistral", text: "Mistral" },
      { key: "codellama", text: "CodeLlama" }
    ],
    supportsCustomModel: true,
    helpUrl: "https://lmstudio.ai"
  },
  custom: {
    name: "Custom/OpenAI-Compatible",
    defaultUrl: "", // User must provide their own URL
    requiresKey: true,
    keyPlaceholder: "your-api-key",
    keyDescription: "API key for your custom endpoint",
    models: [],
    supportsCustomModel: true
  }
};

const providerOptions: IDropdownOption[] = [
  { key: "openai", text: "OpenAI" },
  { key: "azure", text: "Azure OpenAI" },
  { key: "openrouter", text: "OpenRouter (Multi-provider)" },
  { key: "local", text: "Local Model (LM Studio/Ollama)" },
  { key: "custom", text: "Custom/OpenAI-Compatible" }
];

export const Settings: React.FC<SettingsProps> = ({ 
  currentSettings, 
  settingsScope,
  currentLocale,
  onSave, 
  onCancel 
}) => {
  const { t, tFormat } = useTranslation();
  const [provider, setProvider] = useState<ApiProvider>("openrouter");
  const [settings, setSettings] = useState<AISettings>({
    apiUrl: providers.openrouter.defaultUrl,
    apiKey: "",
    model: "",
    temperature: 0.7,
    maxTokens: 4000
  });
  const [locale, setLocale] = useState<SupportedLocale>(currentLocale || 'en');
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [testStatus, setTestStatus] = useState<"none" | "testing" | "success" | "error">("none");
  const [errorMessage, setErrorMessage] = useState("");
  
  // Custom model input
  const [useCustomModel, setUseCustomModel] = useState(false);
  const [customModelName, setCustomModelName] = useState("");
  
  // Model search/filter
  const [modelSearchText, setModelSearchText] = useState("");
  
  // Dynamic model loading for OpenRouter
  const [openRouterModels, setOpenRouterModels] = useState<OpenRouterModel[]>([]);
  const [isLoadingModels, setIsLoadingModels] = useState(false);
  const [modelsLoaded, setModelsLoaded] = useState(false);
  
  // Scope and override settings
  const [saveScope, setSaveScope] = useState<'global' | 'workbook'>('global');
  const [workbookOverrideEnabled, setWorkbookOverrideEnabled] = useState(false);

  // Get available locales
  const availableLocales = SettingsService.getAvailableLocales();
  const localeOptions: IDropdownOption[] = availableLocales.map(loc => ({
    key: loc.code,
    text: `${loc.nativeName} (${loc.name})`
  }));

  // Filter models based on search text
  const filteredModelOptions = useMemo((): IComboBoxOption[] => {
    if (provider === "openrouter" && openRouterModels.length > 0) {
      let models = openRouterModels;
      
      if (modelSearchText) {
        const searchLower = modelSearchText.toLowerCase();
        models = models.filter(m => 
          m.id.toLowerCase().includes(searchLower) ||
          m.name.toLowerCase().includes(searchLower)
        );
      }
      
      return models.map(model => ({
        key: model.id,
        text: `${model.name} - ${OpenRouterService.formatPrice(model.pricing.prompt)}/M`
      }));
    }
    
    // For other providers, use static models
    const staticModels = providers[provider].models;
    if (modelSearchText) {
      const searchLower = modelSearchText.toLowerCase();
      return staticModels
        .filter(m => 
          String(m.key).toLowerCase().includes(searchLower) ||
          m.text.toLowerCase().includes(searchLower)
        )
        .map(m => ({
          key: String(m.key),
          text: m.text
        }));
    }
    
    return staticModels.map(m => ({
      key: String(m.key),
      text: m.text
    }));
  }, [provider, openRouterModels, modelSearchText]);

  // Load OpenRouter models when API key is available
  const loadOpenRouterModels = useCallback(async (apiKey: string) => {
    if (!apiKey || apiKey.length < 10) return;
    
    setIsLoadingModels(true);
    try {
      const openRouterUrl = "https://openrouter.ai/api/v1";
      const models = await OpenRouterService.getAvailableModels(openRouterUrl, apiKey);
      setOpenRouterModels(models);
      setModelsLoaded(true);
    } catch (error) {
      logger.error('Failed to load OpenRouter models', { error });
      // Fall back to default models
      setOpenRouterModels(OpenRouterService.getDefaultModels());
      setModelsLoaded(true);
    } finally {
      setIsLoadingModels(false);
    }
  }, []);

  useEffect(() => {
    // Set initial locale
    if (currentLocale) {
      setLocale(currentLocale);
    }
  }, [currentLocale]);

  useEffect(() => {
    // Use passed settings or load from service
    if (currentSettings) {
      setSettings(currentSettings);
      const detectedProvider = detectProvider(currentSettings.apiUrl);
      setProvider(detectedProvider);
      
      // Check if model is custom (not in default list)
      if (detectedProvider === "openrouter" && currentSettings.apiKey) {
        loadOpenRouterModels(currentSettings.apiKey);
      }
    } else {
      const savedSettings = AIService.getSettings();
      if (savedSettings.apiKey) {
        setSettings(savedSettings);
        const detectedProvider = detectProvider(savedSettings.apiUrl);
        setProvider(detectedProvider);
        
        if (detectedProvider === "openrouter" && savedSettings.apiKey) {
          loadOpenRouterModels(savedSettings.apiKey);
        }
      }
    }
    
    setWorkbookOverrideEnabled(SettingsService.getWorkbookOverrideEnabled());
  }, [currentSettings, loadOpenRouterModels]);

  // Auto-load models when API key changes for OpenRouter
  useEffect(() => {
    if (provider === "openrouter" && settings.apiKey && settings.apiKey.length > 10 && !modelsLoaded) {
      loadOpenRouterModels(settings.apiKey);
    }
  }, [provider, settings.apiKey, modelsLoaded, loadOpenRouterModels]);

  const detectProvider = (url: string): ApiProvider => {
    if (url.includes("openai.com") && !url.includes("openrouter")) return "openai";
    if (url.includes("azure.com") || url.includes("openai.azure")) return "azure";
    if (url.includes("openrouter.ai")) return "openrouter";
    if (url.includes("localhost")) return "local";
    return "custom";
  };

  const handleProviderChange = (newProvider: ApiProvider) => {
    setProvider(newProvider);
    const config = providers[newProvider];
    setSettings({
      ...settings,
      apiUrl: config.defaultUrl,
      model: config.models[0]?.key as string || "",
      apiKey: config.requiresKey ? settings.apiKey : ""
    });
    setTestStatus("none");
    setErrorMessage("");
    setModelsLoaded(false);
    setOpenRouterModels([]);
    setUseCustomModel(false);
    setCustomModelName("");
    setModelSearchText("");
    
    // Auto-load models for OpenRouter if API key exists
    if (newProvider === "openrouter" && settings.apiKey && settings.apiKey.length > 10) {
      loadOpenRouterModels(settings.apiKey);
    }
  };

  const handleSave = () => {
    const finalModel = useCustomModel ? customModelName : settings.model;
    const finalSettings = {
      ...settings,
      model: finalModel
    };
    onSave(finalSettings, saveScope, locale);
  };

  const handleWorkbookOverrideToggle = (enabled: boolean) => {
    SettingsService.setWorkbookOverrideEnabled(enabled);
    setWorkbookOverrideEnabled(enabled);
  };

  const testConnection = async () => {
    setTestStatus("testing");
    setErrorMessage("");

    try {
      const headers: Record<string, string> = {
        "Content-Type": "application/json"
      };
      
      if (settings.apiKey) {
        headers["Authorization"] = `Bearer ${settings.apiKey}`;
      }
      
      if (provider === "openrouter") {
        headers["HTTP-Referer"] = "https://excel-ai-assistant.com";
        headers["X-Title"] = "Excel AI Assistant";
      }

      const response = await fetch(`${settings.apiUrl}/models`, {
        method: "GET",
        headers
      });

      if (response.ok) {
        setTestStatus("success");
        
        // Also refresh models after successful connection for OpenRouter
        if (provider === "openrouter") {
          setModelsLoaded(false);
          await loadOpenRouterModels(settings.apiKey);
        }
      } else {
        const error = await response.json().catch(() => ({ error: { message: `HTTP ${response.status}` } }));
        setTestStatus("error");
        setErrorMessage(error.error?.message || `Error: ${response.status}`);
      }
    } catch (error: any) {
      setTestStatus("error");
      setErrorMessage(error.message || "Failed to connect to API");
    }
  };

  const currentProvider = providers[provider];

  return (
    <div 
      style={{ 
        display: "flex", 
        flexDirection: "column",
        height: "100%",
        minHeight: 0
      }}
      role="form" 
      aria-label="Settings form"
    >
      {/* Header */}
      <div
        style={{
          padding: `${spacing.md} ${spacing.lg}`,
          backgroundColor: colors.brand.primary,
          borderBottom: `1px solid ${colors.brand.primaryDark}`,
          flexShrink: 0
        }}
      >
        <Text variant="xLarge" styles={{ root: { color: colors.text.onBrand, fontWeight: 600 } }}>
          {t('settings.title')}
        </Text>
      </div>

      {/* Scrollable Content */}
      <div 
        style={{ 
          flex: 1, 
          overflow: "auto",
          minHeight: 0,
          padding: spacing.lg
        }}
        role="region" 
        aria-label="Settings content"
      >
        <div style={{ maxWidth: "600px" }}>
          {/* Current Settings Scope Info */}
          {settingsScope && (
            <MessageBar messageBarType={MessageBarType.info}>
              Currently using: {settingsScope.source}
            </MessageBar>
          )}

          {/* Language Selection */}
          <Label style={{ fontWeight: 600, marginTop: spacing.sm }}>{t('settings.language.label')}</Label>
          
          <Dropdown
            selectedKey={locale}
            options={localeOptions}
            onChange={(_, option) => {
              if (option?.key) {
                setLocale(option.key as SupportedLocale);
              }
            }}
            aria-label={t('settings.language.label')}
            styles={{ root: { width: "100%", marginBottom: spacing.xs } }}
          />
          
          <Text variant="small" styles={{ root: { color: colors.text.secondary, marginBottom: spacing.md } }}>
            {t('settings.language.autoDetected')}
          </Text>

          <Separator />

          {/* API Provider Selection */}
          <Label style={{ fontWeight: 600 }}>{t('settings.api.label')}</Label>
          
          <Dropdown
            selectedKey={provider}
            options={providerOptions}
            onChange={(_, option) => handleProviderChange(option?.key as ApiProvider)}
            aria-label="Select API provider"
            styles={{ root: { width: "100%", marginBottom: spacing.xs } }}
          />

          {currentProvider.helpUrl && (
            <Text variant="small" styles={{ root: { color: colors.text.secondary, marginBottom: spacing.md } }}>
              Need an API key?{" "}
              <Link href={currentProvider.helpUrl} target="_blank" rel="noopener noreferrer">
                Get one here
              </Link>
            </Text>
          )}

          <Separator />

          {/* API Configuration */}
          <Label style={{ fontWeight: 600 }}>{t('settings.api.label')}</Label>

          <TextField
            label={t('settings.api.url')}
            value={settings.apiUrl}
            onChange={(_, value) => setSettings({ ...settings, apiUrl: value || "" })}
            placeholder={currentProvider.defaultUrl}
            styles={{ root: { marginBottom: spacing.sm } }}
          />

          {currentProvider.requiresKey && (
            <TextField
              label={t('settings.api.key')}
              type="password"
              value={settings.apiKey}
              onChange={(_, value) => {
                setSettings({ ...settings, apiKey: value || "" });
                if (provider === "openrouter" && value && value.length > 10) {
                  setModelsLoaded(false);
                }
              }}
              placeholder={currentProvider.keyPlaceholder}
              description={currentProvider.keyDescription}
              canRevealPassword
              required
              styles={{ root: { marginBottom: spacing.sm } }}
            />
          )}

          <Separator />

          {/* Model Selection */}
          <Label style={{ fontWeight: 600 }}>{t('settings.api.model')}</Label>

          {/* Custom Model Toggle */}
          {currentProvider.supportsCustomModel && (
            <Toggle
              label="Use custom model name"
              checked={useCustomModel}
              onChange={(_, checked) => setUseCustomModel(checked || false)}
              onText={t('common.yes')}
              offText={t('common.no')}
              styles={{ root: { marginBottom: spacing.sm } }}
            />
          )}

          {useCustomModel ? (
            <TextField
              label="Custom Model Name"
              value={customModelName}
              onChange={(_, value) => setCustomModelName(value || "")}
              placeholder="e.g., openai/gpt-4-turbo"
              description="Enter the exact model identifier"
              styles={{ root: { marginBottom: spacing.sm } }}
            />
          ) : (
            <>
              {provider === "openrouter" && isLoadingModels ? (
                <Stack horizontal tokens={{ childrenGap: spacing.sm }} verticalAlign="center" style={{ marginBottom: spacing.sm }}>
                  <Spinner size={1} />
                  <Text>{t('common.loading')}</Text>
                </Stack>
              ) : (
                <>
                  {provider === "openrouter" && (
                    <TextField
                      placeholder={t('common.search')}
                      value={modelSearchText}
                      onChange={(_, value) => setModelSearchText(value || "")}
                      styles={{ root: { marginBottom: spacing.xs } }}
                    />
                  )}
                  
                  <ComboBox
                    label={t('settings.api.model')}
                    selectedKey={settings.model}
                    text={settings.model || "Select a model..."}
                    options={filteredModelOptions}
                    onChange={(_, option) => {
                      if (option) {
                        setSettings({ ...settings, model: String(option.key) });
                      }
                    }}
                    onInputValueChange={(value) => {
                      if (provider === "openrouter") {
                        setModelSearchText(value || "");
                      }
                    }}
                    styles={{ root: { width: "100%", marginBottom: spacing.sm } }}
                    disabled={provider === "openrouter" && openRouterModels.length === 0 && !isLoadingModels}
                    allowFreeform
                    autoComplete="on"
                  />
                </>
              )}

              {provider === "openrouter" && openRouterModels.length > 0 && (
                <MessageBar messageBarType={MessageBarType.success} styles={{ root: { marginBottom: spacing.sm } }}>
                  ✓ {tFormat('settings.api.loadedModels', { count: openRouterModels.length })}
                </MessageBar>
              )}

              {provider === "openrouter" && !isLoadingModels && openRouterModels.length === 0 && (
                <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: spacing.sm } }}>
                  {t('settings.api.enterKeyToLoad')}
                </MessageBar>
              )}
            </>
          )}

          {provider === "local" && (
            <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { marginBottom: spacing.sm } }}>
              Make sure your local server (LM Studio, Ollama, etc.) is running on the specified port.
            </MessageBar>
          )}

          <Separator />

          {/* Test Connection */}
          <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="center" style={{ marginBottom: spacing.sm }}>
            <DefaultButton
              text={testStatus === "testing" ? t('settings.api.testing') : t('settings.api.testConnection')}
              onClick={testConnection}
              disabled={testStatus === "testing" || (currentProvider.requiresKey && !settings.apiKey)}
            />
          </Stack>

          {testStatus === "success" && (
            <MessageBar messageBarType={MessageBarType.success} styles={{ root: { marginBottom: spacing.sm } }}>
              ✓ {t('settings.api.connectionSuccess')}
            </MessageBar>
          )}

          {testStatus === "error" && (
            <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: spacing.sm } }}>
              {errorMessage}
            </MessageBar>
          )}

          <Separator />

          {/* Storage Scope Selection */}
          <Label style={{ fontWeight: 600 }}>{t('settings.scope.label')}</Label>

          <Toggle
            label={t('settings.scope.enableWorkbook')}
            checked={workbookOverrideEnabled}
            onChange={(_, checked) => handleWorkbookOverrideToggle(checked || false)}
            onText={t('common.yes')}
            offText={t('common.no')}
            styles={{ root: { marginBottom: spacing.sm } }}
          />
          
          {workbookOverrideEnabled ? (
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: spacing.sm } }}>
              {t('settings.scope.workbookDescription')}
            </MessageBar>
          ) : (
            <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: spacing.sm } }}>
              {t('settings.scope.globalDescription')}
            </MessageBar>
          )}

          <Dropdown
            label={t('settings.scope.whereToSave')}
            selectedKey={saveScope}
            options={[
              { key: 'global', text: t('settings.scope.global') },
              { key: 'workbook', text: t('settings.scope.workbook') }
            ]}
            onChange={(_, option) => {
              if (option?.key === 'global' || option?.key === 'workbook') {
                setSaveScope(option.key);
              }
            }}
            disabled={!workbookOverrideEnabled && saveScope === 'workbook'}
            styles={{ root: { width: '100%', marginBottom: spacing.sm } }}
          />

          <Separator />

          {/* Advanced Settings */}
          <Toggle
            label={t('settings.advanced.show')}
            checked={showAdvanced}
            onChange={(_, checked) => setShowAdvanced(checked || false)}
            onText={t('common.yes')}
            offText={t('common.no')}
            styles={{ root: { marginBottom: spacing.sm } }}
          />

          {showAdvanced && (
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField
                label={t('settings.temperature.label')}
                type="number"
                min={0}
                max={2}
                step={0.1}
                value={settings.temperature.toString()}
                onChange={(_, value) => setSettings({ ...settings, temperature: parseFloat(value || "0.7") })}
                description={t('settings.temperature.description')}
              />

              <TextField
                label={t('settings.maxTokens.label')}
                type="number"
                min={100}
                max={16000}
                step={100}
                value={settings.maxTokens.toString()}
                onChange={(_, value) => setSettings({ ...settings, maxTokens: parseInt(value || "4000") })}
                description={t('settings.maxTokens.description')}
              />
            </Stack>
          )}

          {/* Bottom padding for scroll */}
          <div style={{ height: "60px" }} />
        </div>
      </div>

      {/* Footer */}
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          alignItems: "center",
          gap: spacing.sm,
          padding: `${spacing.md} ${spacing.lg}`,
          backgroundColor: colors.neutral.gray20,
          borderTop: `1px solid ${colors.neutral.gray40}`,
          flexShrink: 0
        }}
      >
        <DefaultButton 
          text={t('common.cancel')} 
          onClick={onCancel} 
        />
        <PrimaryButton 
          text={saveScope === 'workbook' ? t('settings.save.workbook') : t('settings.save.global')} 
          onClick={handleSave} 
          disabled={currentProvider.requiresKey && !settings.apiKey}
        />
      </div>
    </div>
  );
};