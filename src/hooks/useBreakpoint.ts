// Breakpoint Hook - Responsive design utilities
// Detect screen size for adaptive layouts

import { useState, useEffect, useCallback } from 'react';

// Breakpoint definitions (in pixels)
export const breakpoints = {
  xs: 0,
  sm: 576,
  md: 768,
  lg: 992,
  xl: 1200,
  xxl: 1400,
};

export type Breakpoint = keyof typeof breakpoints;

export interface BreakpointState {
  /** Current width in pixels */
  width: number;
  /** Current height in pixels */
  height: number;
  /** Current breakpoint name */
  breakpoint: Breakpoint;
  /** True if mobile (xs, sm) */
  isMobile: boolean;
  /** True if tablet (md) */
  isTablet: boolean;
  /** True if desktop (lg, xl, xxl) */
  isDesktop: boolean;
  /** True if >= sm breakpoint */
  isSm: boolean;
  /** True if >= md breakpoint */
  isMd: boolean;
  /** True if >= lg breakpoint */
  isLg: boolean;
  /** True if >= xl breakpoint */
  isXl: boolean;
}

/**
 * Get breakpoint name from width
 */
function getBreakpoint(width: number): Breakpoint {
  if (width >= breakpoints.xxl) return 'xxl';
  if (width >= breakpoints.xl) return 'xl';
  if (width >= breakpoints.lg) return 'lg';
  if (width >= breakpoints.md) return 'md';
  if (width >= breakpoints.sm) return 'sm';
  return 'xs';
}

/**
 * Hook for responsive breakpoint detection
 * 
 * Usage:
 * ```tsx
 * const { isMobile, isDesktop, breakpoint } = useBreakpoint();
 * 
 * return (
 *   <Stack horizontal={!isMobile} vertical={isMobile}>
 *     {isMobile ? <MobileView /> : <DesktopView />}
 *   </Stack>
 * );
 * ```
 */
export function useBreakpoint(): BreakpointState {
  const [state, setState] = useState<BreakpointState>(() => {
    const width = typeof window !== 'undefined' ? window.innerWidth : 0;
    const height = typeof window !== 'undefined' ? window.innerHeight : 0;
    const breakpoint = getBreakpoint(width);
    
    return {
      width,
      height,
      breakpoint,
      isMobile: breakpoint === 'xs' || breakpoint === 'sm',
      isTablet: breakpoint === 'md',
      isDesktop: breakpoint === 'lg' || breakpoint === 'xl' || breakpoint === 'xxl',
      isSm: width >= breakpoints.sm,
      isMd: width >= breakpoints.md,
      isLg: width >= breakpoints.lg,
      isXl: width >= breakpoints.xl,
    };
  });

  useEffect(() => {
    const handleResize = () => {
      const width = window.innerWidth;
      const height = window.innerHeight;
      const breakpoint = getBreakpoint(width);
      
      setState({
        width,
        height,
        breakpoint,
        isMobile: breakpoint === 'xs' || breakpoint === 'sm',
        isTablet: breakpoint === 'md',
        isDesktop: breakpoint === 'lg' || breakpoint === 'xl' || breakpoint === 'xxl',
        isSm: width >= breakpoints.sm,
        isMd: width >= breakpoints.md,
        isLg: width >= breakpoints.lg,
        isXl: width >= breakpoints.xl,
      });
    };

    window.addEventListener('resize', handleResize);
    return () => window.removeEventListener('resize', handleResize);
  }, []);

  return state;
}

/**
 * Hook that runs callback when breakpoint changes
 */
export function useBreakpointEffect(
  callback: (state: BreakpointState) => void,
  deps: React.DependencyList = []
) {
  const breakpoint = useBreakpoint();
  const prevBreakpoint = usePrevious(breakpoint.breakpoint);

  useEffect(() => {
    if (prevBreakpoint && prevBreakpoint !== breakpoint.breakpoint) {
      callback(breakpoint);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [breakpoint.breakpoint, ...deps]);
}

/**
 * Hook to get previous value
 */
function usePrevious<T>(value: T): T | undefined {
  const [prev, setPrev] = useState<T | undefined>(undefined);
  const [curr, setCurr] = useState<T>(value);

  if (value !== curr) {
    setPrev(curr);
    setCurr(value);
  }

  return prev;
}

/**
 * Media query hook for specific breakpoint
 */
export function useMediaQuery(query: string): boolean {
  const [matches, setMatches] = useState(() => {
    if (typeof window === 'undefined') return false;
    return window.matchMedia(query).matches;
  });

  useEffect(() => {
    const media = window.matchMedia(query);
    const listener = (e: MediaQueryListEvent) => setMatches(e.matches);
    
    media.addEventListener('change', listener);
    return () => media.removeEventListener('change', listener);
  }, [query]);

  return matches;
}

/**
 * Hook for responsive visibility
 * Shows/hides based on breakpoint
 */
export function useResponsiveVisibility(
  options: {
    showOn?: Breakpoint[];
    hideOn?: Breakpoint[];
  }
): boolean {
  const { breakpoint } = useBreakpoint();
  const { showOn, hideOn } = options;

  if (hideOn?.includes(breakpoint)) return false;
  if (showOn && !showOn.includes(breakpoint)) return false;
  return true;
}

// Utility for responsive values
export function responsiveValue<T>(
  value: T | Partial<Record<Breakpoint, T>>,
  breakpoint: Breakpoint
): T {
  if (typeof value !== 'object' || value === null) {
    return value as T;
  }

  const map = value as Partial<Record<Breakpoint, T>>;
  const breakpointsOrder: Breakpoint[] = ['xxl', 'xl', 'lg', 'md', 'sm', 'xs'];
  const currentIndex = breakpointsOrder.indexOf(breakpoint);

  // Find the best match (equal or next smaller)
  for (let i = currentIndex; i < breakpointsOrder.length; i++) {
    const bp = breakpointsOrder[i];
    if (map[bp] !== undefined) {
      return map[bp] as T;
    }
  }

  // Fallback to xs
  return map.xs as T;
}

export default useBreakpoint;
