// Design Tokens - Single source of truth for all design values
// Supports theming, dark mode, and white-label customization

// ==========================================
// COLOR TOKENS
// ==========================================

export interface ColorTokens {
  // Brand Colors
  brand: {
    primary: string;
    primaryDark: string;
    primaryLight: string;
    primaryLighter: string;
  };
  
  // Semantic Colors
  semantic: {
    success: string;
    successBackground: string;
    warning: string;
    warningBackground: string;
    error: string;
    errorBackground: string;
    info: string;
    infoBackground: string;
  };
  
  // Neutral Colors (Light Mode)
  neutral: {
    white: string;
    gray10: string;
    gray20: string;
    gray30: string;
    gray40: string;
    gray50: string;
    gray60: string;
    gray70: string;
    gray80: string;
    gray90: string;
    gray100: string;
    black: string;
  };
  
  // Text Colors
  text: {
    primary: string;
    secondary: string;
    disabled: string;
    onBrand: string;
  };
  
  // Background Colors
  background: {
    page: string;
    card: string;
    overlay: string;
    elevated: string;
  };
  
  // Border Colors
  border: {
    default: string;
    hover: string;
    active: string;
    disabled: string;
  };
}

// Light mode color tokens
export const lightColors: ColorTokens = {
  brand: {
    primary: '#0078d4',
    primaryDark: '#005a9e',
    primaryLight: '#2b88d8',
    primaryLighter: '#c7e0f4',
  },
  semantic: {
    success: '#107c10',
    successBackground: '#dff6dd',
    warning: '#ffc107',
    warningBackground: '#fff4ce',
    error: '#d13438',
    errorBackground: '#fed9cc',
    info: '#0078d4',
    infoBackground: '#deecf9',
  },
  neutral: {
    white: '#ffffff',
    gray10: '#faf9f8',
    gray20: '#f3f2f1',
    gray30: '#edebe9',
    gray40: '#e1dfdd',
    gray50: '#d2d0ce',
    gray60: '#c8c6c4',
    gray70: '#a19f9d',
    gray80: '#605e5c',
    gray90: '#3b3a39',
    gray100: '#323130',
    black: '#000000',
  },
  text: {
    primary: '#323130',
    secondary: '#605e5c',
    disabled: '#a19f9d',
    onBrand: '#ffffff',
  },
  background: {
    page: '#faf9f8',
    card: '#ffffff',
    overlay: 'rgba(0, 0, 0, 0.4)',
    elevated: '#ffffff',
  },
  border: {
    default: '#e1dfdd',
    hover: '#c8c6c4',
    active: '#0078d4',
    disabled: '#f3f2f1',
  },
};

// Dark mode color tokens
export const darkColors: ColorTokens = {
  brand: {
    primary: '#2899f5',
    primaryDark: '#0078d4',
    primaryLight: '#6cb8f6',
    primaryLighter: '#106ebe',
  },
  semantic: {
    success: '#54b054',
    successBackground: '#393d1b',
    warning: '#fce100',
    warningBackground: '#4c4400',
    error: '#f1707b',
    errorBackground: '#442726',
    info: '#2899f5',
    infoBackground: '#012b4d',
  },
  neutral: {
    white: '#ffffff',
    gray10: '#1b1a19',
    gray20: '#201f1e',
    gray30: '#252423',
    gray40: '#292827',
    gray50: '#323130',
    gray60: '#3b3a39',
    gray70: '#605e5c',
    gray80: '#a19f9d',
    gray90: '#c8c6c4',
    gray100: '#f3f2f1',
    black: '#000000',
  },
  text: {
    primary: '#f3f2f1',
    secondary: '#a19f9d',
    disabled: '#605e5c',
    onBrand: '#000000',
  },
  background: {
    page: '#1b1a19',
    card: '#201f1e',
    overlay: 'rgba(0, 0, 0, 0.6)',
    elevated: '#252423',
  },
  border: {
    default: '#3b3a39',
    hover: '#605e5c',
    active: '#2899f5',
    disabled: '#292827',
  },
};

// ==========================================
// TYPOGRAPHY TOKENS
// ==========================================

export interface TypographyToken {
  fontSize: string;
  fontWeight: number;
  lineHeight: string;
  fontFamily?: string;
}

export interface TypographyTokens {
  hero: TypographyToken;
  title: TypographyToken;
  subtitle: TypographyToken;
  heading: TypographyToken;
  subheading: TypographyToken;
  bodyLarge: TypographyToken;
  body: TypographyToken;
  bodySmall: TypographyToken;
  caption: TypographyToken;
  button: TypographyToken;
  label: TypographyToken;
}

export const typography: TypographyTokens = {
  hero: { fontSize: '28px', fontWeight: 600, lineHeight: '36px' },
  title: { fontSize: '24px', fontWeight: 600, lineHeight: '32px' },
  subtitle: { fontSize: '20px', fontWeight: 600, lineHeight: '28px' },
  heading: { fontSize: '18px', fontWeight: 600, lineHeight: '24px' },
  subheading: { fontSize: '16px', fontWeight: 600, lineHeight: '22px' },
  bodyLarge: { fontSize: '16px', fontWeight: 400, lineHeight: '24px' },
  body: { fontSize: '14px', fontWeight: 400, lineHeight: '20px' },
  bodySmall: { fontSize: '12px', fontWeight: 400, lineHeight: '16px' },
  caption: { fontSize: '11px', fontWeight: 400, lineHeight: '14px' },
  button: { fontSize: '14px', fontWeight: 600, lineHeight: '20px' },
  label: { fontSize: '12px', fontWeight: 600, lineHeight: '16px' },
};

// ==========================================
// SPACING TOKENS
// ==========================================

export interface SpacingTokens {
  none: string;
  xs: string;
  sm: string;
  md: string;
  lg: string;
  xl: string;
  xxl: string;
  xxxl: string;
}

export const spacing: SpacingTokens = {
  none: '0',
  xs: '4px',
  sm: '8px',
  md: '16px',
  lg: '24px',
  xl: '32px',
  xxl: '48px',
  xxxl: '64px',
};

// ==========================================
// SHADOW TOKENS
// ==========================================

export interface ShadowTokens {
  none: string;
  sm: string;
  md: string;
  lg: string;
  xl: string;
}

export const shadows: ShadowTokens = {
  none: 'none',
  sm: '0 1px 2px rgba(0, 0, 0, 0.1)',
  md: '0 4px 8px rgba(0, 0, 0, 0.12)',
  lg: '0 8px 16px rgba(0, 0, 0, 0.14)',
  xl: '0 16px 32px rgba(0, 0, 0, 0.16)',
};

// ==========================================
// BORDER RADIUS TOKENS
// ==========================================

export interface BorderRadiusTokens {
  none: string;
  sm: string;
  md: string;
  lg: string;
  xl: string;
  full: string;
}

export const borderRadius: BorderRadiusTokens = {
  none: '0',
  sm: '2px',
  md: '4px',
  lg: '8px',
  xl: '16px',
  full: '9999px',
};

// ==========================================
// Z-INDEX TOKENS
// ==========================================

export interface ZIndexTokens {
  base: number;
  dropdown: number;
  sticky: number;
  fixed: number;
  modalBackdrop: number;
  modal: number;
  popover: number;
  tooltip: number;
  toast: number;
}

export const zIndex: ZIndexTokens = {
  base: 0,
  dropdown: 100,
  sticky: 200,
  fixed: 300,
  modalBackdrop: 400,
  modal: 500,
  popover: 600,
  tooltip: 700,
  toast: 1000,
};

// ==========================================
// COMPLETE THEME INTERFACE
// ==========================================

export interface Theme {
  name: string;
  isDark: boolean;
  colors: ColorTokens;
  typography: TypographyTokens;
  spacing: SpacingTokens;
  shadows: ShadowTokens;
  borderRadius: BorderRadiusTokens;
  zIndex: ZIndexTokens;
}

// Pre-defined themes
export const lightTheme: Theme = {
  name: 'light',
  isDark: false,
  colors: lightColors,
  typography,
  spacing,
  shadows,
  borderRadius,
  zIndex,
};

export const darkTheme: Theme = {
  name: 'dark',
  isDark: true,
  colors: darkColors,
  typography,
  spacing,
  shadows,
  borderRadius,
  zIndex,
};

export const excelTheme: Theme = {
  ...lightTheme,
  name: 'excel',
  colors: {
    ...lightColors,
    brand: {
      primary: '#217346',
      primaryDark: '#106ebe',
      primaryLight: '#2e8c57',
      primaryLighter: '#d4edda',
    },
  },
};

export const allThemes = {
  light: lightTheme,
  dark: darkTheme,
  excel: excelTheme,
};

export type ThemeName = keyof typeof allThemes;
