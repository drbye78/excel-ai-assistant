// Conditional Formatting Service
// Applies various conditional formatting rules to Excel ranges via natural language

export type ConditionalFormatType =
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'topBottom'
  | 'aboveBelowAverage'
  | 'uniqueDuplicate'
  | 'customFormula';

export interface ColorScaleRule {
  type: 'colorScale';
  minimum: { type: 'min' | 'num' | 'percent' | 'percentile' | 'formula'; value?: string; color: string };
  midpoint?: { type: 'num' | 'percent' | 'percentile' | 'formula'; value: string; color: string };
  maximum: { type: 'max' | 'num' | 'percent' | 'percentile' | 'formula'; value?: string; color: string };
}

export interface DataBarRule {
  type: 'dataBar';
  axisPosition?: 'automatic' | 'middle' | 'none';
  barDirection?: 'leftToRight' | 'rightToLeft';
  color: string;
  showDataBarOnly?: boolean;
  minLength?: number;
  maxLength?: number;
}

export interface IconSetRule {
  type: 'iconSet';
  iconSet: string; // '3TrafficLights', '3Arrows', '5Rating', etc.
  reverseIconOrder?: boolean;
  showIconOnly?: boolean;
}

export interface TopBottomRule {
  type: 'topBottom';
  rank: number;
  isTop: boolean;
  isPercent?: boolean;
  format: CellFormat;
}

export interface AboveBelowAverageRule {
  type: 'aboveBelowAverage';
  isAbove: boolean;
  isEqualToAverage?: boolean;
  standardDeviations?: number;
  format: CellFormat;
}

export interface UniqueDuplicateRule {
  type: 'uniqueDuplicate';
  isDuplicate: boolean;
  format: CellFormat;
}

export interface CustomFormulaRule {
  type: 'customFormula';
  formula: string;
  format: CellFormat;
  stopIfTrue?: boolean;
}

export interface CellFormat {
  font?: {
    bold?: boolean;
    italic?: boolean;
    color?: string;
  };
  fill?: {
    color?: string;
  };
  borders?: {
    color?: string;
    style?: 'continuous' | 'dash' | 'dashDot' | 'dashDotDot' | 'dot' | 'double' | 'slantDashDot';
  };
}

export type ConditionalFormatRule =
  | ColorScaleRule
  | DataBarRule
  | IconSetRule
  | TopBottomRule
  | AboveBelowAverageRule
  | UniqueDuplicateRule
  | CustomFormulaRule;

// Predefined color scales
export const COLOR_SCALES = {
  greenYellowRed: {
    name: 'Green - Yellow - Red',
    minimum: { type: 'min' as const, color: '#63BE7B' },
    midpoint: { type: 'percentile' as const, value: '50', color: '#FFEB84' },
    maximum: { type: 'max' as const, color: '#F8696B' }
  },
  redYellowGreen: {
    name: 'Red - Yellow - Green',
    minimum: { type: 'min' as const, color: '#F8696B' },
    midpoint: { type: 'percentile' as const, value: '50', color: '#FFEB84' },
    maximum: { type: 'max' as const, color: '#63BE7B' }
  },
  greenWhiteRed: {
    name: 'Green - White - Red',
    minimum: { type: 'min' as const, color: '#63BE7B' },
    midpoint: { type: 'percentile' as const, value: '50', color: '#FFFFFF' },
    maximum: { type: 'max' as const, color: '#F8696B' }
  },
  redWhiteGreen: {
    name: 'Red - White - Green',
    minimum: { type: 'min' as const, color: '#F8696B' },
    midpoint: { type: 'percentile' as const, value: '50', color: '#FFFFFF' },
    maximum: { type: 'max' as const, color: '#63BE7B' }
  },
  blueWhiteRed: {
    name: 'Blue - White - Red',
    minimum: { type: 'min' as const, color: '#5B9BD5' },
    midpoint: { type: 'percentile' as const, value: '50', color: '#FFFFFF' },
    maximum: { type: 'max' as const, color: '#FF0000' }
  },
  whiteRed: {
    name: 'White - Red',
    minimum: { type: 'min' as const, color: '#FFFFFF' },
    maximum: { type: 'max' as const, color: '#F8696B' }
  },
  greenWhite: {
    name: 'Green - White',
    minimum: { type: 'min' as const, color: '#63BE7B' },
    maximum: { type: 'max' as const, color: '#FFFFFF' }
  }
};

// Predefined icon sets
export const ICON_SETS = {
  '3Arrows': { name: '3 Arrows', icons: ['↑', '→', '↓'] },
  '3ArrowsGray': { name: '3 Arrows (Gray)', icons: ['↑', '→', '↓'] },
  '3TrafficLights': { name: '3 Traffic Lights', icons: ['🔴', '🟡', '🟢'] },
  '3TrafficLightsRimless': { name: '3 Traffic Lights (Rimless)', icons: ['🔴', '🟡', '🟢'] },
  '3Signs': { name: '3 Signs', icons: ['✓', '!', '✗'] },
  '3Symbols': { name: '3 Symbols', icons: ['●', '◐', '○'] },
  '3SymbolsCircle': { name: '3 Symbols (Circle)', icons: ['●', '◐', '○'] },
  '4Arrows': { name: '4 Arrows', icons: ['↑', '↗', '↘', '↓'] },
  '4ArrowsGray': { name: '4 Arrows (Gray)', icons: ['↑', '↗', '↘', '↓'] },
  '4Rating': { name: '4 Rating', icons: ['★★★★', '★★★', '★★', '★'] },
  '4TrafficLights': { name: '4 Traffic Lights', icons: ['🔴', '🟡', '🟢', '⚫'] },
  '5Arrows': { name: '5 Arrows', icons: ['↑', '↗', '→', '↘', '↓'] },
  '5ArrowsGray': { name: '5 Arrows (Gray)', icons: ['↑', '↗', '→', '↘', '↓'] },
  '5Rating': { name: '5 Rating', icons: ['★★★★★', '★★★★', '★★★', '★★', '★'] },
  '5Quarters': { name: '5 Quarters', icons: ['●●●●', '●●●○', '●●○○', '●○○○', '○○○○'] }
};

export class ConditionalFormattingService {
  // Apply color scale
  async applyColorScale(
    range: string,
    rule: ColorScaleRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);

      // Build color scale criteria
      const criteria: Excel.ConditionalColorScaleCriteria = {
        minimum: this.buildColorScaleCriterion(rule.minimum),
        maximum: this.buildColorScaleCriterion(rule.maximum)
      };

      if (rule.midpoint) {
        criteria.midpoint = this.buildColorScaleCriterion(rule.midpoint);
      }

      targetRange.conditionalFormat.addColorScale(criteria);

      await context.sync();
    });
  }

  // Apply data bar
  async applyDataBar(
    range: string,
    rule: DataBarRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);

      const criteria: Excel.ConditionalDataBarRule = {
        barDirection: rule.barDirection || 'leftToRight',
        showDataBarOnly: rule.showDataBarOnly || false,
        axisPosition: rule.axisPosition || 'automatic',
        fillColor: rule.color,
        borderColor: rule.color
      };

      if (rule.minLength !== undefined) {
        criteria.minLength = rule.minLength;
      }
      if (rule.maxLength !== undefined) {
        criteria.maxLength = rule.maxLength;
      }

      targetRange.conditionalFormat.addDataBar(criteria);

      await context.sync();
    });
  }

  // Apply icon set
  async applyIconSet(
    range: string,
    rule: IconSetRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);

      const criteria: Excel.ConditionalIconSetRule = {
        iconSet: rule.iconSet as any, // Type assertion needed for icon set names
        reverseIconOrder: rule.reverseIconOrder || false,
        showIconOnly: rule.showIconOnly || false
      };

      targetRange.conditionalFormat.addIconSet(criteria);

      await context.sync();
    });
  }

  // Apply top/bottom rule
  async applyTopBottom(
    range: string,
    rule: TopBottomRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      const format = this.buildCellFormat(rule.format);

      if (rule.isTop) {
        if (rule.isPercent) {
          targetRange.conditionalFormat.addTopPercent(rule.rank, format);
        } else {
          targetRange.conditionalFormat.addTopN(rule.rank, format);
        }
      } else {
        if (rule.isPercent) {
          targetRange.conditionalFormat.addBottomPercent(rule.rank, format);
        } else {
          targetRange.conditionalFormat.addBottomN(rule.rank, format);
        }
      }

      await context.sync();
    });
  }

  // Apply above/below average
  async applyAboveBelowAverage(
    range: string,
    rule: AboveBelowAverageRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      const format = this.buildCellFormat(rule.format);

      if (rule.standardDeviations) {
        // Use standard deviation based rule
        if (rule.isAbove) {
          targetRange.conditionalFormat.addAboveAverage(format);
        } else {
          targetRange.conditionalFormat.addBelowAverage(format);
        }
      } else {
        // Standard above/below average
        if (rule.isAbove) {
          if (rule.isEqualToAverage) {
            targetRange.conditionalFormat.addAboveOrEqualAverage(format);
          } else {
            targetRange.conditionalFormat.addAboveAverage(format);
          }
        } else {
          if (rule.isEqualToAverage) {
            targetRange.conditionalFormat.addBelowOrEqualAverage(format);
          } else {
            targetRange.conditionalFormat.addBelowAverage(format);
          }
        }
      }

      await context.sync();
    });
  }

  // Apply unique/duplicate rule
  async applyUniqueDuplicate(
    range: string,
    rule: UniqueDuplicateRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      const format = this.buildCellFormat(rule.format);

      if (rule.isDuplicate) {
        targetRange.conditionalFormat.addDuplicateValues(format);
      } else {
        targetRange.conditionalFormat.addUniqueValues(format);
      }

      await context.sync();
    });
  }

  // Apply custom formula rule
  async applyCustomFormula(
    range: string,
    rule: CustomFormulaRule,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      const format = this.buildCellFormat(rule.format);

      targetRange.conditionalFormat.addCustom(
        rule.formula,
        format,
        rule.stopIfTrue
      );

      await context.sync();
    });
  }

  // Clear all conditional formatting
  async clearConditionalFormatting(
    range: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      targetRange.conditionalFormat.clearAll();

      await context.sync();
    });
  }

  // Get all conditional formatting rules for a range
  async getConditionalFormattingRules(
    range: string,
    worksheetName?: string
  ): Promise<any[]> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const targetRange = worksheet.getRange(range);
      const formats = targetRange.conditionalFormat;
      formats.load('items');

      await context.sync();

      return formats.items.map((format, index) => ({
        index,
        priority: format.priority,
        stopIfTrue: format.stopIfTrue,
        type: format.type
      }));
    });
  }

  // Apply formatting from natural language description
  async applyFromDescription(
    range: string,
    description: string,
    worksheetName?: string
  ): Promise<void> {
    const rule = this.parseNaturalLanguage(description);
    await this.applyRule(range, rule, worksheetName);
  }

  // Parse natural language to formatting rule
  parseNaturalLanguage(description: string): ConditionalFormatRule {
    const lowerDesc = description.toLowerCase();

    // Color scales
    if (lowerDesc.includes('green') && lowerDesc.includes('yellow') && lowerDesc.includes('red')) {
      if (lowerDesc.includes('reverse') || lowerDesc.includes('red') < lowerDesc.indexOf('green')) {
        return { type: 'colorScale', ...COLOR_SCALES.redYellowGreen };
      }
      return { type: 'colorScale', ...COLOR_SCALES.greenYellowRed };
    }

    if (lowerDesc.includes('green') && lowerDesc.includes('white') && lowerDesc.includes('red')) {
      if (lowerDesc.includes('reverse')) {
        return { type: 'colorScale', ...COLOR_SCALES.redWhiteGreen };
      }
      return { type: 'colorScale', ...COLOR_SCALES.greenWhiteRed };
    }

    if (lowerDesc.includes('blue') && lowerDesc.includes('red')) {
      return { type: 'colorScale', ...COLOR_SCALES.blueWhiteRed };
    }

    if (lowerDesc.includes('white') && lowerDesc.includes('red')) {
      return { type: 'colorScale', ...COLOR_SCALES.whiteRed };
    }

    if (lowerDesc.includes('green') && lowerDesc.includes('white')) {
      return { type: 'colorScale', ...COLOR_SCALES.greenWhite };
    }

    // Data bars
    if (lowerDesc.includes('data bar') || lowerDesc.includes('databar')) {
      const color = this.extractColor(description) || '#5B9BD5';
      return {
        type: 'dataBar',
        color,
        showDataBarOnly: lowerDesc.includes('only') || lowerDesc.includes('without numbers')
      };
    }

    // Icon sets
    if (lowerDesc.includes('traffic light')) {
      return {
        type: 'iconSet',
        iconSet: '3TrafficLights'
      };
    }

    if (lowerDesc.includes('arrow')) {
      if (lowerDesc.includes('4')) {
        return { type: 'iconSet', iconSet: '4Arrows' };
      }
      if (lowerDesc.includes('5')) {
        return { type: 'iconSet', iconSet: '5Arrows' };
      }
      return { type: 'iconSet', iconSet: '3Arrows' };
    }

    if (lowerDesc.includes('rating') || lowerDesc.includes('star')) {
      if (lowerDesc.includes('5')) {
        return { type: 'iconSet', iconSet: '5Rating' };
      }
      return { type: 'iconSet', iconSet: '4Rating' };
    }

    // Top/bottom
    const topMatch = description.match(/top\s+(\d+)/i);
    const bottomMatch = description.match(/bottom\s+(\d+)/i);

    if (topMatch) {
      return {
        type: 'topBottom',
        rank: parseInt(topMatch[1]),
        isTop: true,
        isPercent: lowerDesc.includes('percent'),
        format: { fill: { color: '#90EE90' } }
      };
    }

    if (bottomMatch) {
      return {
        type: 'topBottom',
        rank: parseInt(bottomMatch[1]),
        isTop: false,
        isPercent: lowerDesc.includes('percent'),
        format: { fill: { color: '#FFB6C1' } }
      };
    }

    // Above/below average
    if (lowerDesc.includes('above average')) {
      return {
        type: 'aboveBelowAverage',
        isAbove: true,
        format: { fill: { color: '#90EE90' } }
      };
    }

    if (lowerDesc.includes('below average')) {
      return {
        type: 'aboveBelowAverage',
        isAbove: false,
        format: { fill: { color: '#FFB6C1' } }
      };
    }

    // Duplicates
    if (lowerDesc.includes('duplicate')) {
      return {
        type: 'uniqueDuplicate',
        isDuplicate: true,
        format: { fill: { color: '#FFB6C1' } }
      };
    }

    if (lowerDesc.includes('unique')) {
      return {
        type: 'uniqueDuplicate',
        isDuplicate: false,
        format: { fill: { color: '#90EE90' } }
      };
    }

    // Default: highlight cells containing specific text/value
    return {
      type: 'customFormula',
      formula: 'TRUE', // Default formula, should be customized based on description
      format: { fill: { color: '#FFEB3B' } }
    };
  }

  // Apply a rule based on its type
  private async applyRule(
    range: string,
    rule: ConditionalFormatRule,
    worksheetName?: string
  ): Promise<void> {
    switch (rule.type) {
      case 'colorScale':
        await this.applyColorScale(range, rule, worksheetName);
        break;
      case 'dataBar':
        await this.applyDataBar(range, rule, worksheetName);
        break;
      case 'iconSet':
        await this.applyIconSet(range, rule, worksheetName);
        break;
      case 'topBottom':
        await this.applyTopBottom(range, rule, worksheetName);
        break;
      case 'aboveBelowAverage':
        await this.applyAboveBelowAverage(range, rule, worksheetName);
        break;
      case 'uniqueDuplicate':
        await this.applyUniqueDuplicate(range, rule, worksheetName);
        break;
      case 'customFormula':
        await this.applyCustomFormula(range, rule, worksheetName);
        break;
    }
  }

  // Helper: Build color scale criterion
  private buildColorScaleCriterion(
    criterion: ColorScaleRule['minimum'] | ColorScaleRule['midpoint'] | ColorScaleRule['maximum']
  ): Excel.ConditionalColorScaleCriterion {
    const result: Excel.ConditionalColorScaleCriterion = {
      type: criterion.type as any,
      color: criterion.color
    };

    if (criterion.value !== undefined) {
      result.formula = criterion.value;
    }

    return result;
  }

  // Helper: Build cell format
  private buildCellFormat(format: CellFormat): Excel.ConditionalCellValueRule {
    const result: Excel.ConditionalCellValueRule = {};

    if (format.font) {
      result.font = {
        bold: format.font.bold,
        italic: format.font.italic,
        color: format.font.color
      };
    }

    if (format.fill) {
      result.fill = {
        color: format.fill.color
      };
    }

    if (format.borders) {
      result.borders = {
        color: format.borders.color,
        style: format.borders.style as any
      };
    }

    return result;
  }

  // Helper: Extract color from description
  private extractColor(description: string): string | null {
    const colorMap: Record<string, string> = {
      'red': '#FF0000',
      'green': '#00FF00',
      'blue': '#0000FF',
      'yellow': '#FFFF00',
      'orange': '#FFA500',
      'purple': '#800080',
      'pink': '#FFC0CB',
      'cyan': '#00FFFF',
      'magenta': '#FF00FF',
      'white': '#FFFFFF',
      'black': '#000000',
      'gray': '#808080',
      'grey': '#808080'
    };

    const lowerDesc = description.toLowerCase();
    for (const [name, hex] of Object.entries(colorMap)) {
      if (lowerDesc.includes(name)) {
        return hex;
      }
    }

    return null;
  }
}

// Singleton instance
let serviceInstance: ConditionalFormattingService | null = null;

export function getConditionalFormattingService(): ConditionalFormattingService {
  if (!serviceInstance) {
    serviceInstance = new ConditionalFormattingService();
  }
  return serviceInstance;
}

export default getConditionalFormattingService;
