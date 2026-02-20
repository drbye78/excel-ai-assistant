/**
 * Excel AI Assistant - Visual Highlighter Service
 * Real-time cell highlighting and visual feedback in Excel
 * 
 * @module services/visualHighlighter
 */

// ============================================================================
// Type Definitions
// ============================================================================

/** A highlighted range with metadata */
export interface HighlightedRange {
  id: string;
  address: string;
  /** Alias for address for component compatibility */
  rangeAddress: string;
  worksheetName: string;
  color: string;
  originalColor?: string;
  pulse?: boolean;
  tooltip?: string;
  category: HighlightCategory;
  createdAt: Date;
  /** Alias for createdAt for component compatibility */
  timestamp: Date;
  autoClear?: boolean;
  clearAfter?: number; // milliseconds
}

/** Alias for HighlightedRange for component compatibility */
export type HighlightInfo = HighlightedRange;

/** Extended highlight options with pulse speed */
export interface ExtendedHighlightOptions extends HighlightOptions {
  pulseSpeed?: 'slow' | 'normal' | 'fast';
}

/** Categories of highlights for different purposes */
export type HighlightCategory =
  | "ai-reference"
  | "error"
  | "warning"
  | "success"
  | "info"
  | "anomaly"
  | "trend"
  | "selection";

/** Highlight configuration options */
export interface HighlightOptions {
  color?: string;
  pulse?: boolean;
  tooltip?: string;
  category?: HighlightCategory;
  autoClear?: boolean;
  clearAfter?: number; // milliseconds, default 30000 (30 seconds)
  preserveOriginal?: boolean;
}

/** Predefined color schemes */
export interface ColorScheme {
  name: string;
  background: string;
  border?: string;
  text?: string;
}

/** Pulse animation configuration */
export interface PulseConfig {
  duration: number;
  iterations: number;
  easing: string;
}

// ============================================================================
// Predefined Color Schemes
// ============================================================================

export const HIGHLIGHT_COLORS: Record<HighlightCategory, ColorScheme> = {
  "ai-reference": {
    name: "AI Reference",
    background: "#FFF4CE", // Yellow
    border: "#FFC107"
  },
  error: {
    name: "Error",
    background: "#FDE7E9", // Red
    border: "#F44336",
    text: "#B71C1C"
  },
  warning: {
    name: "Warning",
    background: "#FFF4CE", // Amber
    border: "#FFC107"
  },
  success: {
    name: "Success",
    background: "#DFF6DD", // Green
    border: "#107C10"
  },
  info: {
    name: "Info",
    background: "#E8F4FD", // Blue
    border: "#0078D4"
  },
  anomaly: {
    name: "Anomaly",
    background: "#FFE5E5", // Light Red
    border: "#D13438"
  },
  trend: {
    name: "Trend",
    background: "#E6F4EA", // Light Green
    border: "#34A853"
  },
  selection: {
    name: "Selection",
    background: "#F3F2F1", // Gray
    border: "#8A8886"
  }
};

// ============================================================================
// Visual Highlighter Service
// ============================================================================

import { logger } from '../utils/logger';

export class VisualHighlighter {
  private static instance: VisualHighlighter;
  private highlightedRanges: Map<string, HighlightedRange> = new Map();
  private autoClearTimers: Map<string, number> = new Map();
  private isInitialized: boolean = false;

  private constructor() {}

  static getInstance(): VisualHighlighter {
    if (!VisualHighlighter.instance) {
      VisualHighlighter.instance = new VisualHighlighter();
    }
    return VisualHighlighter.instance;
  }

  // ============================================================================
  // Initialization
  // ============================================================================

  initialize(): void {
    if (this.isInitialized) return;
    this.isInitialized = true;
    logger.info('VisualHighlighter initialized');
  }

  // ============================================================================
  // Highlight Operations
  // ============================================================================

  /**
   * Highlight a range in Excel
   */
  async highlightRange(
    address: string,
    worksheetName?: string,
    options: HighlightOptions = {}
  ): Promise<string | null> {
    const {
      color,
      pulse = false,
      tooltip,
      category = "ai-reference",
      autoClear = false,
      clearAfter = 30000,
      preserveOriginal = true
    } = options;

    const id = `highlight-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

    try {
      // @ts-ignore - Office.js types
      await Excel.run(async (context: Excel.RequestContext) => {
        const worksheet = worksheetName
          ? context.workbook.worksheets.getItem(worksheetName)
          : context.workbook.worksheets.getActiveWorksheet();

        const range = worksheet.getRange(address);

        // Load current format if preserving
        if (preserveOriginal) {
          range.load("format/fill/color");
          await context.sync();
        }

        // Apply highlight color
        const highlightColor = color || HIGHLIGHT_COLORS[category].background;
        range.format.fill.color = highlightColor;

        // Add border if specified
        const borderColor = HIGHLIGHT_COLORS[category].border;
        if (borderColor) {
          range.format.borders.getItem("EdgeTop").color = borderColor;
          range.format.borders.getItem("EdgeBottom").color = borderColor;
          range.format.borders.getItem("EdgeLeft").color = borderColor;
          range.format.borders.getItem("EdgeRight").color = borderColor;
          range.format.borders.getItem("EdgeTop").style = "Continuous";
          range.format.borders.getItem("EdgeBottom").style = "Continuous";
          range.format.borders.getItem("EdgeLeft").style = "Continuous";
          range.format.borders.getItem("EdgeRight").style = "Continuous";
        }

        await context.sync();

        // Store highlight info
        const now = new Date();
        const highlightedRange: HighlightedRange = {
          id,
          address,
          rangeAddress: address,
          worksheetName: worksheetName || (await this.getActiveWorksheetName()),
          color: highlightColor,
          originalColor: preserveOriginal ? range.format.fill.color : undefined,
          pulse,
          tooltip,
          category,
          createdAt: now,
          timestamp: now,
          autoClear,
          clearAfter
        };

        this.highlightedRanges.set(id, highlightedRange);

        // Setup auto-clear if enabled
        if (autoClear && clearAfter > 0) {
          this.setupAutoClear(id, clearAfter);
        }

        // Apply pulse animation if requested
        if (pulse) {
          this.applyPulseAnimation(id);
        }
      });

      return id;
    } catch (error) {
      logger.error('Failed to highlight range', { address, worksheetName }, error as Error);
      return null;
    }
  }

  /**
   * Highlight multiple ranges at once
   */
  async highlightRanges(
    ranges: { address: string; worksheetName?: string; options?: HighlightOptions }[]
  ): Promise<string[]> {
    const ids: string[] = [];

    for (const range of ranges) {
      const id = await this.highlightRange(
        range.address,
        range.worksheetName,
        range.options
      );
      if (id) ids.push(id);
    }

    return ids;
  }

  /**
   * Clear a specific highlight
   */
  async clearHighlight(id: string): Promise<boolean> {
    const highlight = this.highlightedRanges.get(id);
    if (!highlight) return false;

    try {
      // @ts-ignore - Office.js types
      await Excel.run(async (context: Excel.RequestContext) => {
        const worksheet = context.workbook.worksheets.getItem(highlight.worksheetName);
        const range = worksheet.getRange(highlight.address);

        // Restore original color if available
        if (highlight.originalColor) {
          range.format.fill.color = highlight.originalColor;
        } else {
          range.format.fill.clear();
        }

        // Clear borders
        range.format.borders.getItem("EdgeTop").style = "None";
        range.format.borders.getItem("EdgeBottom").style = "None";
        range.format.borders.getItem("EdgeLeft").style = "None";
        range.format.borders.getItem("EdgeRight").style = "None";

        await context.sync();
      });

      // Clear timer if exists
      this.clearAutoClearTimer(id);
      this.highlightedRanges.delete(id);

      return true;
    } catch (error) {
      logger.error('Failed to clear highlight', { highlightId: id }, error as Error);
      return false;
    }
  }

  /**
   * Clear all highlights
   */
  async clearAllHighlights(): Promise<void> {
    const ids = Array.from(this.highlightedRanges.keys());
    
    for (const id of ids) {
      await this.clearHighlight(id);
    }

    this.highlightedRanges.clear();
    this.autoClearTimers.clear();
  }

  /**
   * Clear highlights by category
   */
  async clearHighlightsByCategory(category: HighlightCategory): Promise<void> {
    const toClear: string[] = [];

    this.highlightedRanges.forEach((highlight, id) => {
      if (highlight.category === category) {
        toClear.push(id);
      }
    });

    for (const id of toClear) {
      await this.clearHighlight(id);
    }
  }

  // ============================================================================
  // Auto-Clear Functionality
  // ============================================================================

  private setupAutoClear(id: string, delay: number): void {
    // Clear existing timer if any
    this.clearAutoClearTimer(id);

    // Set new timer
    const timerId = window.setTimeout(() => {
      this.clearHighlight(id);
    }, delay);

    this.autoClearTimers.set(id, timerId);
  }

  private clearAutoClearTimer(id: string): void {
    const timerId = this.autoClearTimers.get(id);
    if (timerId) {
      window.clearTimeout(timerId);
      this.autoClearTimers.delete(id);
    }
  }

  // ============================================================================
  // Pulse Animation
  // ============================================================================

  private async applyPulseAnimation(id: string): Promise<void> {
    const highlight = this.highlightedRanges.get(id);
    if (!highlight) return;

    // Simple pulse by alternating opacity
    let pulseCount = 0;
    const maxPulses = 3;
    const interval = 500; // ms

    const pulseInterval = window.setInterval(async () => {
      if (pulseCount >= maxPulses * 2) {
        window.clearInterval(pulseInterval);
        return;
      }

      try {
        // @ts-ignore - Office.js types
        await Excel.run(async (context: Excel.RequestContext) => {
          const worksheet = context.workbook.worksheets.getItem(highlight.worksheetName);
          const range = worksheet.getRange(highlight.address);

          // Toggle between highlight color and slightly darker
          if (pulseCount % 2 === 0) {
            range.format.fill.color = this.darkenColor(highlight.color, 20);
          } else {
            range.format.fill.color = highlight.color;
          }

          await context.sync();
        });

        pulseCount++;
      } catch (error) {
        logger.error('Pulse animation error', { highlightId: id }, error as Error);
        window.clearInterval(pulseInterval);
      }
    }, interval);
  }

  private darkenColor(color: string, percent: number): string {
    // Simple hex color darkening
    const num = parseInt(color.replace("#", ""), 16);
    const amt = Math.round(2.55 * percent);
    const R = Math.max((num >> 16) - amt, 0);
    const G = Math.max((num >> 8 & 0x00FF) - amt, 0);
    const B = Math.max((num & 0x0000FF) - amt, 0);
    return "#" + (0x1000000 + R * 0x10000 + G * 0x100 + B).toString(16).slice(1);
  }

  // ============================================================================
  // Navigation
  // ============================================================================

  /**
   * Navigate to a highlighted range and select it
   */
  async navigateToHighlight(id: string): Promise<boolean> {
    const highlight = this.highlightedRanges.get(id);
    if (!highlight) return false;

    try {
      // @ts-ignore - Office.js types
      await Excel.run(async (context: Excel.RequestContext) => {
        const worksheet = context.workbook.worksheets.getItem(highlight.worksheetName);
        const range = worksheet.getRange(highlight.address);

        // Activate the worksheet
        worksheet.activate();

        // Select the range
        range.select();

        await context.sync();
      });

      return true;
    } catch (error) {
      logger.error('Failed to navigate to highlight', { highlightId: id }, error as Error);
      return false;
    }
  }

  // ============================================================================
  // Utility Methods
  // ============================================================================

  private async getActiveWorksheetName(): Promise<string> {
    try {
      // @ts-ignore - Office.js types
      return await Excel.run(async (context: Excel.RequestContext) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load("name");
        await context.sync();
        return worksheet.name;
      });
    } catch {
      return "Sheet1";
    }
  }

  /**
   * Get all active highlights (alias for getActiveHighlights)
   */
  getAllHighlights(): HighlightedRange[] {
    return this.getActiveHighlights();
  }

  /**
   * Get all active highlights
   */
  getActiveHighlights(): HighlightedRange[] {
    return Array.from(this.highlightedRanges.values());
  }

  /**
   * Get a specific highlight by ID
   */
  getHighlight(id: string): HighlightedRange | undefined {
    return this.highlightedRanges.get(id);
  }

  /**
   * Get highlights by category
   */
  getHighlightsByCategory(category: HighlightCategory): HighlightedRange[] {
    return Array.from(this.highlightedRanges.values()).filter(
      h => h.category === category
    );
  }

  /**
   * Clear highlights by category (alias for clearHighlightsByCategory)
   */
  async clearByCategory(category: HighlightCategory): Promise<void> {
    return this.clearHighlightsByCategory(category);
  }

  /**
   * Check if a range is currently highlighted
   */
  isHighlighted(address: string, worksheetName?: string): boolean {
    const normalizedAddress = address.toUpperCase();
    const normalizedWorksheet = worksheetName?.toUpperCase();

    return Array.from(this.highlightedRanges.values()).some(
      h => 
        h.address.toUpperCase() === normalizedAddress &&
        (!normalizedWorksheet || h.worksheetName.toUpperCase() === normalizedWorksheet)
    );
  }
}

// Export singleton instance
export const visualHighlighter = VisualHighlighter.getInstance();
export default visualHighlighter;
