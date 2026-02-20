// Theme Provider - Context for theme management
// Supports dark mode, theme switching, and persistence

import React, { createContext, useContext, useState, useEffect, useCallback } from 'react';
import { Theme, ThemeName, allThemes, lightTheme } from './tokens';

interface ThemeContextType {
  theme: Theme;
  themeName: ThemeName;
  setTheme: (name: ThemeName) => void;
  toggleDarkMode: () => void;
  isDark: boolean;
}

const ThemeContext = createContext<ThemeContextType | undefined>(undefined);

const STORAGE_KEY = 'excel-ai-assistant-theme';

export interface ThemeProviderProps {
  children: React.ReactNode;
  defaultTheme?: ThemeName;
}

/**
 * Theme Provider Component
 * 
 * Manages theme state, persistence, and CSS custom properties.
 * 
 * Usage:
 * ```tsx
 * <ThemeProvider defaultTheme="light">
 *   <App />
 * </ThemeProvider>
 * ```
 */
export const ThemeProvider: React.FC<ThemeProviderProps> = ({
  children,
  defaultTheme = 'light',
}) => {
  const [themeName, setThemeName] = useState<ThemeName>(() => {
    // Load from localStorage or use default
    if (typeof window !== 'undefined') {
      const saved = localStorage.getItem(STORAGE_KEY) as ThemeName;
      return saved && allThemes[saved] ? saved : defaultTheme;
    }
    return defaultTheme;
  });

  const theme = allThemes[themeName] || lightTheme;

  // Persist theme changes
  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, themeName);
  }, [themeName]);

  // Apply CSS custom properties
  useEffect(() => {
    const root = document.documentElement;
    
    // Colors
    root.style.setProperty('--color-brand-primary', theme.colors.brand.primary);
    root.style.setProperty('--color-brand-primary-dark', theme.colors.brand.primaryDark);
    root.style.setProperty('--color-brand-primary-light', theme.colors.brand.primaryLight);
    root.style.setProperty('--color-brand-primary-lighter', theme.colors.brand.primaryLighter);
    
    root.style.setProperty('--color-semantic-success', theme.colors.semantic.success);
    root.style.setProperty('--color-semantic-warning', theme.colors.semantic.warning);
    root.style.setProperty('--color-semantic-error', theme.colors.semantic.error);
    root.style.setProperty('--color-semantic-info', theme.colors.semantic.info);
    
    root.style.setProperty('--color-text-primary', theme.colors.text.primary);
    root.style.setProperty('--color-text-secondary', theme.colors.text.secondary);
    root.style.setProperty('--color-text-disabled', theme.colors.text.disabled);
    
    root.style.setProperty('--color-background-page', theme.colors.background.page);
    root.style.setProperty('--color-background-card', theme.colors.background.card);
    root.style.setProperty('--color-background-overlay', theme.colors.background.overlay);
    
    root.style.setProperty('--color-border-default', theme.colors.border.default);
    root.style.setProperty('--color-border-hover', theme.colors.border.hover);
    root.style.setProperty('--color-border-active', theme.colors.border.active);

    // Apply data attribute for CSS selectors
    root.setAttribute('data-theme', themeName);
    
    // Optional: Sync with system preference
    const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
    const handleChange = (e: MediaQueryListEvent) => {
      // Only auto-switch if user hasn't manually set preference
      const hasUserPreference = localStorage.getItem(STORAGE_KEY + '-user-set');
      if (!hasUserPreference) {
        setThemeName(e.matches ? 'dark' : 'light');
      }
    };
    
    mediaQuery.addEventListener('change', handleChange);
    return () => mediaQuery.removeEventListener('change', handleChange);
  }, [theme, themeName]);

  const setTheme = useCallback((name: ThemeName) => {
    localStorage.setItem(STORAGE_KEY + '-user-set', 'true');
    setThemeName(name);
  }, []);

  const toggleDarkMode = useCallback(() => {
    setThemeName((current) => {
      const currentTheme = allThemes[current];
      return currentTheme.isDark ? 'light' : 'dark';
    });
    localStorage.setItem(STORAGE_KEY + '-user-set', 'true');
  }, []);

  const value: ThemeContextType = {
    theme,
    themeName,
    setTheme,
    toggleDarkMode,
    isDark: theme.isDark,
  };

  return (
    <ThemeContext.Provider value={value}>
      {children}
    </ThemeContext.Provider>
  );
};

/**
 * Hook to access theme context
 */
export const useTheme = (): ThemeContextType => {
  const context = useContext(ThemeContext);
  if (!context) {
    throw new Error('useTheme must be used within a ThemeProvider');
  }
  return context;
};

/**
 * Hook to get CSS custom property value
 */
export const useCssVar = (name: string): string => {
  const [value, setValue] = useState('');
  
  useEffect(() => {
    const root = document.documentElement;
    const computed = getComputedStyle(root).getPropertyValue(name).trim();
    setValue(computed);
  }, [name]);
  
  return value;
};

export default ThemeProvider;
