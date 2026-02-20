// Pivot Table Explainer - Natural Language Explanations for Pivot Tables
// Phase 2 Implementation - Natural Language Interface

import { PivotTableService, PivotTableInfo } from './pivotTableService';

export interface PivotExplanation {
  summary: string;
  structure: {
    rows: string[];
    columns: string[];
    values: Array<{ name: string; aggregation: string; format?: string }>;
    filters: string[];
  };
  insights: string[];
  recommendations: string[];
}

export interface PivotValueInsight {
  label: string;
  value: number;
  context: string;
  significance: 'high' | 'medium' | 'low';
}

export class PivotTableExplainer {
  private static instance: PivotTableExplainer;
  private pivotTableService: typeof PivotTableService;

  private constructor() {
    this.pivotTableService = PivotTableService;
  }

  static getInstance(): PivotTableExplainer {
    if (!PivotTableExplainer.instance) {
      PivotTableExplainer.instance = new PivotTableExplainer();
    }
    return PivotTableExplainer.instance;
  }

  /**
   * Generate a comprehensive natural language explanation of a pivot table
   */
  async explainPivotTable(pivotName: string, worksheetName?: string): Promise<PivotExplanation> {
    const info = await this.pivotTableService.getPivotTableInfo(pivotName, worksheetName);
    const data = await this.pivotTableService.getPivotData(pivotName, worksheetName);

    return {
      summary: this.generateSummary(info, data),
      structure: this.explainStructure(info),
      insights: this.generateInsights(info, data),
      recommendations: this.generateRecommendations(info, data)
    };
  }

  /**
   * Generate a concise summary of the pivot table
   */
  private generateSummary(info: PivotTableInfo, data: any[][]): string {
    const totalRows = data.length;
    const totalCols = data[0]?.length || 0;

    let summary = `This pivot table "${info.name}" analyzes data with `;

    // Describe structure
    if (info.rowFields.length > 0) {
      summary += `${info.rowFields.join(', ')} in rows`;
      if (info.columnFields.length > 0) {
        summary += ` and ${info.columnFields.join(', ')} in columns`;
      }
    } else if (info.columnFields.length > 0) {
      summary += `${info.columnFields.join(', ')} in columns`;
    }

    // Describe values
    if (info.dataFields.length > 0) {
      const valueDesc = info.dataFields.map(f => {
        const agg = f.aggregation === 'sum' ? 'total' : f.aggregation;
        return `${agg} ${f.name}`;
      }).join(', ');
      summary += `, showing ${valueDesc}`;
    }

    // Describe filters
    if (info.filterFields.length > 0) {
      summary += `. It's filtered by ${info.filterFields.join(', ')}`;
    }

    summary += `. The table contains ${totalRows} rows and ${totalCols} columns of data.`;

    return summary;
  }

  /**
   * Explain the structure of the pivot table
   */
  private explainStructure(info: PivotTableInfo): PivotExplanation['structure'] {
    return {
      rows: info.rowFields,
      columns: info.columnFields,
      values: info.dataFields.map(f => ({
        name: f.name,
        aggregation: f.aggregation,
        format: this.describeAggregation(f.aggregation)
      })),
      filters: info.filterFields
    };
  }

  /**
   * Generate insights from the pivot table data
   */
  private generateInsights(info: PivotInfo, data: any[][]): string[] {
    const insights: string[] = [];

    // Skip if no data
    if (!data || data.length < 2) {
      return ['Insufficient data to generate insights.'];
    }

    // Calculate totals if we have numeric data
    const numericData = this.extractNumericValues(data);
    if (numericData.length > 0) {
      const total = numericData.reduce((a, b) => a + b, 0);
      const avg = total / numericData.length;
      const max = Math.max(...numericData);
      const min = Math.min(...numericData);

      insights.push(`The overall total is ${this.formatNumber(total)}.`);
      insights.push(`Values range from ${this.formatNumber(min)} to ${this.formatNumber(max)}, with an average of ${this.formatNumber(avg)}.`);

      // Find highest/lowest by row if applicable
      if (info.rowFields.length > 0) {
        const rowExtremes = this.findRowExtremes(data);
        if (rowExtremes.highest) {
          insights.push(`The highest value is in "${rowExtremes.highest.label}" with ${this.formatNumber(rowExtremes.highest.value)}.`);
        }
        if (rowExtremes.lowest && rowExtremes.lowest.label !== rowExtremes.highest?.label) {
          insights.push(`The lowest value is in "${rowExtremes.lowest.label}" with ${this.formatNumber(rowExtremes.lowest.value)}.`);
        }
      }
    }

    // Analyze distribution
    const distribution = this.analyzeDistribution(data);
    if (distribution.dominantCategory) {
      insights.push(`"${distribution.dominantCategory}" is the largest category, representing ${distribution.dominantPercentage?.toFixed(1)}% of the total.`);
    }

    return insights;
  }

  /**
   * Generate recommendations based on the pivot table
   */
  private generateRecommendations(info: PivotInfo, data: any[][]): string[] {
    const recommendations: string[] = [];

    // Suggest adding time-based grouping if date field present
    if (info.rowFields.some(f => f.toLowerCase().includes('date')) ||
        info.columnFields.some(f => f.toLowerCase().includes('date'))) {
      recommendations.push('Consider grouping the date field by months or quarters for better trend analysis.');
    }

    // Suggest adding grand totals if not present
    if (!info.layout.showGrandTotalsForRows && !info.layout.showGrandTotalsForColumns) {
      recommendations.push('Adding grand totals would help you see the overall summary at a glance.');
    }

    // Suggest calculated field
    if (info.dataFields.length >= 2) {
      recommendations.push('You could add a calculated field to show ratios or differences between value fields.');
    }

    // Suggest chart creation
    if (info.dataFields.length > 0) {
      recommendations.push('Creating a pivot chart would visualize these trends more effectively.');
    }

    // Suggest filtering if many rows
    if (data.length > 50) {
      recommendations.push('With many rows, consider adding a filter field to focus on specific segments.');
    }

    return recommendations;
  }

  /**
   * Generate a detailed breakdown of a specific data point
   */
  async explainDataPoint(
    pivotName: string,
    rowIndex: number,
    colIndex: number,
    worksheetName?: string
  ): Promise<string> {
    const info = await this.pivotTableService.getPivotTableInfo(pivotName, worksheetName);
    const data = await this.pivotTableService.getPivotData(pivotName, worksheetName);

    if (!data[rowIndex] || !data[rowIndex][colIndex]) {
      return 'The specified data point was not found.';
    }

    const value = data[rowIndex][colIndex];
    const rowLabel = data[rowIndex][0] || `Row ${rowIndex}`;
    const colLabel = data[0]?.[colIndex] || `Column ${colIndex}`;

    let explanation = `The value at "${rowLabel}" × "${colLabel}" is ${value}.\n\n`;

    // Add context
    explanation += `This represents `;
    if (info.dataFields.length > 0) {
      explanation += `the ${info.dataFields[0].aggregation} of ${info.dataFields[0].name} `;
    }
    explanation += `for ${rowLabel}`;
    if (info.columnFields.length > 0) {
      explanation += ` in the ${colLabel} category`;
    }
    explanation += '.';

    // Compare to total
    const numericValues = this.extractNumericValues(data);
    if (numericValues.length > 0 && !isNaN(Number(value))) {
      const total = numericValues.reduce((a, b) => a + b, 0);
      const percentage = (Number(value) / total) * 100;
      explanation += `\n\nThis represents ${percentage.toFixed(2)}% of the total.`;
    }

    return explanation;
  }

  /**
   * Compare two data points in the pivot table
   */
  async compareDataPoints(
    pivotName: string,
    point1: { row: number; col: number },
    point2: { row: number; col: number },
    worksheetName?: string
  ): Promise<string> {
    const data = await this.pivotTableService.getPivotData(pivotName, worksheetName);

    const value1 = data[point1.row]?.[point1.col];
    const value2 = data[point2.row]?.[point2.col];

    if (value1 === undefined || value2 === undefined) {
      return 'One or both data points were not found.';
    }

    const num1 = Number(value1);
    const num2 = Number(value2);

    if (isNaN(num1) || isNaN(num2)) {
      return `Comparison: "${value1}" vs "${value2}" (non-numeric values)`;
    }

    const diff = num1 - num2;
    const percentDiff = ((diff / num2) * 100);
    const row1 = data[point1.row]?.[0] || `Row ${point1.row}`;
    const row2 = data[point2.row]?.[0] || `Row ${point2.row}`;

    let comparison = `Comparing "${row1}" (${this.formatNumber(num1)}) with "${row2}" (${this.formatNumber(num2)}):\n\n`;

    if (diff > 0) {
      comparison += `"${row1}" is ${this.formatNumber(Math.abs(diff))} higher (${Math.abs(percentDiff).toFixed(1)}% more).`;
    } else if (diff < 0) {
      comparison += `"${row1}" is ${this.formatNumber(Math.abs(diff))} lower (${Math.abs(percentDiff).toFixed(1)}% less).`;
    } else {
      comparison += 'Both values are equal.';
    }

    return comparison;
  }

  /**
   * Generate a trend analysis for time-based pivot tables
   */
  async analyzeTrend(
    pivotName: string,
    worksheetName?: string
  ): Promise<string> {
    const info = await this.pivotTableService.getPivotTableInfo(pivotName, worksheetName);
    const data = await this.pivotTableService.getPivotData(pivotName, worksheetName);

    const isTimeBased = info.rowFields.some(f =>
      f.toLowerCase().includes('date') ||
      f.toLowerCase().includes('month') ||
      f.toLowerCase().includes('year') ||
      f.toLowerCase().includes('quarter')
    ) || info.columnFields.some(f =>
      f.toLowerCase().includes('date') ||
      f.toLowerCase().includes('month') ||
      f.toLowerCase().includes('year')
    );

    if (!isTimeBased) {
      return 'Trend analysis requires a date, month, quarter, or year field in the pivot table.';
    }

    const numericValues = this.extractNumericValues(data);
    if (numericValues.length < 2) {
      return 'Insufficient numeric data for trend analysis.';
    }

    // Calculate trend
    const firstHalf = numericValues.slice(0, Math.floor(numericValues.length / 2));
    const secondHalf = numericValues.slice(Math.floor(numericValues.length / 2));

    const firstAvg = firstHalf.reduce((a, b) => a + b, 0) / firstHalf.length;
    const secondAvg = secondHalf.reduce((a, b) => a + b, 0) / secondHalf.length;

    const trend = secondAvg > firstAvg ? 'increasing' : 'decreasing';
    const percentChange = ((secondAvg - firstAvg) / firstAvg) * 100;

    let analysis = `Trend Analysis:\n\n`;
    analysis += `The data shows an overall ${trend} trend of ${Math.abs(percentChange).toFixed(1)}%.\n`;

    // Find peak and trough
    const max = Math.max(...numericValues);
    const min = Math.min(...numericValues);
    const maxIndex = numericValues.indexOf(max);
    const minIndex = numericValues.indexOf(min);

    analysis += `\nPeak value: ${this.formatNumber(max)} at position ${maxIndex + 1}`;
    analysis += `\nLowest value: ${this.formatNumber(min)} at position ${minIndex + 1}`;

    // Volatility
    const volatility = this.calculateVolatility(numericValues);
    analysis += `\n\nVolatility: ${volatility > 0.3 ? 'High' : volatility > 0.15 ? 'Medium' : 'Low'} (${(volatility * 100).toFixed(1)}%)`;

    return analysis;
  }

  /**
   * Explain what-if scenarios
   */
  generateWhatIfExplanation(
    currentValue: number,
    scenario: 'increase' | 'decrease',
    percentage: number
  ): string {
    const multiplier = scenario === 'increase' ? 1 + (percentage / 100) : 1 - (percentage / 100);
    const newValue = currentValue * multiplier;
    const change = newValue - currentValue;

    let explanation = `What-if Scenario Analysis:\n\n`;
    explanation += `Current value: ${this.formatNumber(currentValue)}\n`;
    explanation += `Scenario: ${percentage}% ${scenario}\n`;
    explanation += `New value: ${this.formatNumber(newValue)}\n`;
    explanation += `Change: ${change > 0 ? '+' : ''}${this.formatNumber(change)}`;

    return explanation;
  }

  // ============================================================================
  // UTILITY METHODS
  // ============================================================================

  private extractNumericValues(data: any[][]): number[] {
    const values: number[] = [];
    for (let i = 1; i < data.length; i++) { // Skip header
      for (let j = 1; j < data[i].length; j++) { // Skip row labels
        const val = Number(data[i][j]);
        if (!isNaN(val) && data[i][j] !== '') {
          values.push(val);
        }
      }
    }
    return values;
  }

  private findRowExtremes(data: any[][]): { highest?: { label: string; value: number }; lowest?: { label: string; value: number } } {
    let highest: { label: string; value: number } | undefined;
    let lowest: { label: string; value: number } | undefined;

    for (let i = 1; i < data.length; i++) {
      const label = data[i][0];
      for (let j = 1; j < data[i].length; j++) {
        const val = Number(data[i][j]);
        if (!isNaN(val)) {
          if (!highest || val > highest.value) {
            highest = { label, value: val };
          }
          if (!lowest || val < lowest.value) {
            lowest = { label, value: val };
          }
        }
      }
    }

    return { highest, lowest };
  }

  private analyzeDistribution(data: any[][]): { dominantCategory?: string; dominantPercentage?: number } {
    const rowTotals: Map<string, number> = new Map();

    for (let i = 1; i < data.length; i++) {
      const label = data[i][0];
      let total = 0;
      for (let j = 1; j < data[i].length; j++) {
        const val = Number(data[i][j]);
        if (!isNaN(val)) {
          total += val;
        }
      }
      rowTotals.set(label, total);
    }

    let grandTotal = 0;
    let dominantCategory = '';
    let maxTotal = 0;

    for (const [category, total] of rowTotals) {
      grandTotal += total;
      if (total > maxTotal) {
        maxTotal = total;
        dominantCategory = category;
      }
    }

    return {
      dominantCategory,
      dominantPercentage: grandTotal > 0 ? (maxTotal / grandTotal) * 100 : 0
    };
  }

  private calculateVolatility(values: number[]): number {
    if (values.length < 2) return 0;

    const avg = values.reduce((a, b) => a + b, 0) / values.length;
    const squaredDiffs = values.map(v => Math.pow(v - avg, 2));
    const variance = squaredDiffs.reduce((a, b) => a + b, 0) / values.length;
    const stdDev = Math.sqrt(variance);

    return stdDev / avg; // Coefficient of variation
  }

  private formatNumber(num: number): string {
    if (Math.abs(num) >= 1000000) {
      return (num / 1000000).toFixed(2) + 'M';
    } else if (Math.abs(num) >= 1000) {
      return (num / 1000).toFixed(2) + 'K';
    }
    return num.toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  private describeAggregation(agg: string): string {
    const descriptions: Record<string, string> = {
      'sum': 'Total',
      'count': 'Count',
      'average': 'Average',
      'max': 'Maximum',
      'min': 'Minimum',
      'product': 'Product',
      'countNumbers': 'Count of Numbers',
      'stdDev': 'Standard Deviation',
      'stdDevp': 'Standard Deviation (Population)',
      'var': 'Variance',
      'varp': 'Variance (Population)'
    };
    return descriptions[agg] || agg;
  }

  // ============================================================================
  // NATURAL LANGUAGE GENERATION
  // ============================================================================

  /**
   * Generate a human-readable description from pivot info
   */
  generateNaturalLanguageDescription(info: PivotTableInfo): string {
    let description = '';

    // Opening
    description += `This pivot table `;

    // Purpose
    if (info.dataFields.length > 0) {
      const dataDesc = info.dataFields.map(f => {
        return `${this.describeAggregation(f.aggregation)} of ${f.name}`;
      }).join(' and ');
      description += `shows ${dataDesc}`;
    }

    // Grouping
    if (info.rowFields.length > 0) {
      description += `, organized by ${info.rowFields.join(' and ')} in rows`;
    }

    if (info.columnFields.length > 0) {
      if (info.rowFields.length > 0) {
        description += ` and `;
      } else {
        description += `, organized by `;
      }
      description += `${info.columnFields.join(' and ')} in columns`;
    }

    // Filters
    if (info.filterFields.length > 0) {
      description += `. The data is filtered by ${info.filterFields.join(' and ')}`;
    }

    description += '.';

    return description;
  }

  /**
   * Explain a specific configuration change
   */
  explainConfigurationChange(
    action: 'add' | 'remove' | 'move' | 'change',
    target: string,
    from?: string,
    to?: string
  ): string {
    switch (action) {
      case 'add':
        return `Added "${target}" to the pivot table.`;
      case 'remove':
        return `Removed "${target}" from the pivot table.`;
      case 'move':
        return `Moved "${target}" from ${from} to ${to}.`;
      case 'change':
        return `Changed the configuration of "${target}".`;
      default:
        return `Modified "${target}".`;
    }
  }
}

// Type alias for internal use
type PivotInfo = PivotTableInfo;

export default PivotTableExplainer.getInstance();
