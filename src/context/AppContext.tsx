/**
 * Global Application Context
 * Phase 4: State Management
 * 
 * Provides centralized state management for the application
 */

import React, { createContext, useContext, useReducer, useEffect, ReactNode } from 'react';
import { AISettings, Message, AIAction, ExcelContext } from '@/types';
import AIService from '@/services/aiService';
import { secureStorage } from '@/utils/encryption';
import { STORAGE_KEYS, API_DEFAULTS } from '@/config/constants';
import { logger } from '@/utils/logger';

// ============================================================================
// STATE TYPES
// ============================================================================

export interface AppState {
  // Settings
  settings: AISettings;
  settingsScope: { type: string; source: string };
  hasSettings: boolean;
  
  // Conversation
  messages: Message[];
  isLoading: boolean;
  suggestedPrompts: string[];
  pendingActions: AIAction[] | null;
  
  // Excel Context
  excelContext: ExcelContext | null;
  
  // UI State
  activeTab: 'chat' | 'settings' | 'history' | 'analytics';
  error: AppError | null;
  
  // User
  locale: string;
}

export interface AppError {
  message: string;
  code: string;
  timestamp: Date;
}

// ============================================================================
// ACTION TYPES
// ============================================================================

type AppAction =
  | { type: 'SET_SETTINGS'; payload: AISettings }
  | { type: 'SET_SETTINGS_SCOPE'; payload: { type: string; source: string } }
  | { type: 'SET_HAS_SETTINGS'; payload: boolean }
  | { type: 'ADD_MESSAGE'; payload: Message }
  | { type: 'SET_MESSAGES'; payload: Message[] }
  | { type: 'CLEAR_MESSAGES' }
  | { type: 'SET_LOADING'; payload: boolean }
  | { type: 'SET_SUGGESTED_PROMPTS'; payload: string[] }
  | { type: 'SET_PENDING_ACTIONS'; payload: AIAction[] | null }
  | { type: 'SET_EXCEL_CONTEXT'; payload: ExcelContext | null }
  | { type: 'SET_ACTIVE_TAB'; payload: 'chat' | 'settings' | 'history' | 'analytics' }
  | { type: 'SET_ERROR'; payload: AppError | null }
  | { type: 'SET_LOCALE'; payload: string }
  | { type: 'RESET_STATE' };

// ============================================================================
// INITIAL STATE
// ============================================================================

const initialState: AppState = {
  settings: {
    apiUrl: API_DEFAULTS.URL,
    apiKey: '',
    model: API_DEFAULTS.MODEL,
    temperature: API_DEFAULTS.TEMPERATURE,
    maxTokens: API_DEFAULTS.MAX_TOKENS
  },
  settingsScope: { type: 'default', source: 'Built-in defaults' },
  hasSettings: false,
  messages: [],
  isLoading: false,
  suggestedPrompts: [],
  pendingActions: null,
  excelContext: null,
  activeTab: 'chat',
  error: null,
  locale: 'en'
};

// ============================================================================
// REDUCER
// ============================================================================

function appReducer(state: AppState, action: AppAction): AppState {
  switch (action.type) {
    case 'SET_SETTINGS':
      return { ...state, settings: action.payload };
    
    case 'SET_SETTINGS_SCOPE':
      return { ...state, settingsScope: action.payload };
    
    case 'SET_HAS_SETTINGS':
      return { ...state, hasSettings: action.payload };
    
    case 'ADD_MESSAGE':
      return { ...state, messages: [...state.messages, action.payload] };
    
    case 'SET_MESSAGES':
      return { ...state, messages: action.payload };
    
    case 'CLEAR_MESSAGES':
      return { ...state, messages: [] };
    
    case 'SET_LOADING':
      return { ...state, isLoading: action.payload };
    
    case 'SET_SUGGESTED_PROMPTS':
      return { ...state, suggestedPrompts: action.payload };
    
    case 'SET_PENDING_ACTIONS':
      return { ...state, pendingActions: action.payload };
    
    case 'SET_EXCEL_CONTEXT':
      return { ...state, excelContext: action.payload };
    
    case 'SET_ACTIVE_TAB':
      return { ...state, activeTab: action.payload };
    
    case 'SET_ERROR':
      return { ...state, error: action.payload };
    
    case 'SET_LOCALE':
      return { ...state, locale: action.payload };
    
    case 'RESET_STATE':
      return initialState;
    
    default:
      return state;
  }
}

// ============================================================================
// CONTEXT
// ============================================================================

interface AppContextValue {
  state: AppState;
  dispatch: React.Dispatch<AppAction>;
  // Convenience actions
  actions: {
    updateSettings: (settings: AISettings, scope?: 'global' | 'workbook') => Promise<void>;
    addMessage: (message: Message) => void;
    clearMessages: () => void;
    setLoading: (loading: boolean) => void;
    setError: (error: AppError | null) => void;
    setActiveTab: (tab: 'chat' | 'settings' | 'history' | 'analytics') => void;
    setLocale: (locale: string) => void;
    resetState: () => void;
  };
}

const AppContext = createContext<AppContextValue | undefined>(undefined);

// ============================================================================
// PROVIDER
// ============================================================================

interface AppProviderProps {
  children: ReactNode;
}

export function AppProvider({ children }: AppProviderProps) {
  const [state, dispatch] = useReducer(appReducer, initialState);

  // Initialize on mount
  useEffect(() => {
    const initializeApp = async () => {
      try {
        // Initialize AI service
        await AIService.initialize();
        
        // Load settings
        const settings = AIService.getSettings();
        dispatch({ type: 'SET_SETTINGS', payload: settings });
        
        // Get scope
        const scope = AIService.getSettingsScope();
        dispatch({ type: 'SET_SETTINGS_SCOPE', payload: scope });
        
        // Check if we have settings
        dispatch({ type: 'SET_HAS_SETTINGS', payload: !!settings.apiKey });
        
      } catch (error) {
        logger.error('Failed to initialize app', { error });
        dispatch({
          type: 'SET_ERROR',
          payload: {
            message: 'Failed to initialize application',
            code: 'INIT_ERROR',
            timestamp: new Date()
          }
        });
      }
    };

    initializeApp();
  }, []);

  // Persist messages when they change
  useEffect(() => {
    if (state.messages.length > 0) {
      try {
        localStorage.setItem(
          STORAGE_KEYS.CONVERSATIONS,
          JSON.stringify(state.messages.slice(-50)) // Keep last 50 messages
        );
      } catch (error) {
        logger.warn('Failed to save conversation', { error });
      }
    }
  }, [state.messages]);

  // Convenience actions
  const actions = {
    updateSettings: async (settings: AISettings, scope: 'global' | 'workbook' = 'global') => {
      try {
        // Securely store API key
        if (settings.apiKey) {
          await secureStorage.store(STORAGE_KEYS.API_KEY, settings.apiKey);
        }
        
        // Update AI service
        await AIService.updateSettings(settings, true, scope);
        
        dispatch({ type: 'SET_SETTINGS', payload: settings });
        dispatch({ type: 'SET_HAS_SETTINGS', payload: !!settings.apiKey });
        
        // Update scope info
        const scopeInfo = AIService.getSettingsScope();
        dispatch({ type: 'SET_SETTINGS_SCOPE', payload: scopeInfo });
        
      } catch (error) {
        logger.error('Failed to update settings', { error, settings });
        dispatch({
          type: 'SET_ERROR',
          payload: {
            message: 'Failed to save settings',
            code: 'SETTINGS_ERROR',
            timestamp: new Date()
          }
        });
      }
    },

    addMessage: (message: Message) => {
      dispatch({ type: 'ADD_MESSAGE', payload: message });
    },

    clearMessages: () => {
      dispatch({ type: 'CLEAR_MESSAGES' });
      localStorage.removeItem(STORAGE_KEYS.CONVERSATIONS);
    },

    setLoading: (loading: boolean) => {
      dispatch({ type: 'SET_LOADING', payload: loading });
    },

    setError: (error: AppError | null) => {
      dispatch({ type: 'SET_ERROR', payload: error });
    },

    setActiveTab: (tab: 'chat' | 'settings' | 'history' | 'analytics') => {
      dispatch({ type: 'SET_ACTIVE_TAB', payload: tab });
    },

    setLocale: (locale: string) => {
      dispatch({ type: 'SET_LOCALE', payload: locale });
    },

    resetState: () => {
      dispatch({ type: 'RESET_STATE' });
    }
  };

  return (
    <AppContext.Provider value={{ state, dispatch, actions }}>
      {children}
    </AppContext.Provider>
  );
}

// ============================================================================
// HOOKS
// ============================================================================

export function useApp() {
  const context = useContext(AppContext);
  if (context === undefined) {
    throw new Error('useApp must be used within an AppProvider');
  }
  return context;
}

export function useSettings() {
  const { state, actions } = useApp();
  return {
    settings: state.settings,
    settingsScope: state.settingsScope,
    hasSettings: state.hasSettings,
    updateSettings: actions.updateSettings
  };
}

export function useConversation() {
  const { state, actions } = useApp();
  return {
    messages: state.messages,
    isLoading: state.isLoading,
    suggestedPrompts: state.suggestedPrompts,
    pendingActions: state.pendingActions,
    addMessage: actions.addMessage,
    clearMessages: actions.clearMessages,
    setLoading: actions.setLoading
  };
}

export function useError() {
  const { state, actions } = useApp();
  return {
    error: state.error,
    setError: actions.setError
  };
}

export default AppContext;