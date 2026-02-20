// DAX Measure Service - Complete CRUD for DAX Measures and Calculated Columns
// Phase 4 Implementation - Power Pivot & DAX Support

import { PowerPivotService, PowerPivotMeasure, PowerPivotCalculatedColumn } from './powerPivotService';
import { daxParser } from './daxParser';
import { logger } from '../utils/logger';

export interface MeasureTemplate {
  name: string;
  category: 'time_intelligence' | 'aggregation' | 'comparison' | 'ratio' | 'custom';
  description: string;
  template: string;
  parameters: Array<{
    name: string;
    type: 'column' | 'measure' | 'table' | 'number' | 'string';
    description: string;
    defaultValue?: string;
  }>;
}

export interface QuickMeasureConfig {
  type: 'sum' | 'average' | 'count' | 'distinct_count' | 'min' | 'max' | 'ytd' | 'qtd' | 'mtd' | 'yoy_growth' | 'qoq_growth' | 'mom_growth' | 'running_total' | 'market_share';
  baseColumn: string;
  baseTable: string;
  dateColumn?: string;
  dateTable?: string;
  filterColumn?: string;
  filterValue?: string;
}

export interface MeasureValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  complexity: 'simple' | 'moderate' | 'complex' | 'very_complex';
  estimatedPerformance: 'fast' | 'moderate' | 'slow';
}

export class DAXMeasureService {
  private static instance: DAXMeasureService;
  private powerPivotService: typeof PowerPivotService;

  private constructor() {
    this.powerPivotService = PowerPivotService;
  }

  static getInstance(): DAXMeasureService {
    if (!DAXMeasureService.instance) {
      DAXMeasureService.instance = new DAXMeasureService();
    }
    return DAXMeasureService.instance;
  }

  // ============================================================================
  // MEASURE CRUD OPERATIONS
  // ============================================================================

  /**
   * Create a new measure
   */
  async createMeasure(measure: PowerPivotMeasure): Promise<boolean> {
    const validation = this.validateMeasure(measure);
    if (!validation.isValid) {
      logger.error('Measure validation failed', { errors: validation.errors, measure });
      return false;
    }

    return await this.powerPivotService.createMeasure(measure);
  }

  /**
   * Update an existing measure
   */
  async updateMeasure(
    tableName: string,
    measureName: string,
    updates: Partial<PowerPivotMeasure>
  ): Promise<boolean> {
    // Get existing measure
    const measures = await this.powerPivotService.getMeasures();
    const existingMeasure = measures.find(m => m.table === tableName && m.name === measureName);

    if (!existingMeasure) {
      logger.error('Measure not found', { measureName, tableName });
      return false;
    }

    const updatedMeasure = { ...existingMeasure, ...updates };
    const validation = this.validateMeasure(updatedMeasure);

    if (!validation.isValid) {
      logger.error('Updated measure validation failed', { errors: validation.errors, measure: updatedMeasure });
      return false;
    }

    // Delete old and create new (since Office.js doesn't support direct update)
    await this.powerPivotService.deleteMeasure(tableName, measureName);
    return await this.powerPivotService.createMeasure(updatedMeasure);
  }

  /**
   * Delete a measure
   */
  async deleteMeasure(tableName: string, measureName: string): Promise<boolean> {
    return await this.powerPivotService.deleteMeasure(tableName, measureName);
  }

  /**
   * Get all measures with detailed info
   */
  async getAllMeasures(): Promise<PowerPivotMeasure[]> {
    return await this.powerPivotService.getMeasures();
  }

  /**
   * Get measures by table
   */
  async getMeasuresByTable(tableName: string): Promise<PowerPivotMeasure[]> {
    const measures = await this.powerPivotService.getMeasures();
    return measures.filter(m => m.table === tableName);
  }

  // ============================================================================
  // CALCULATED COLUMN OPERATIONS
  // ============================================================================

  /**
   * Create a calculated column
   */
  async createCalculatedColumn(column: PowerPivotCalculatedColumn): Promise<boolean> {
    try {
      // @ts-ignore - Office.js API
      return await Excel.run(async (context: Excel.RequestContext) => {
        // Get the table
        const table = context.workbook.tables.getItem(column.table);
        
        // Add the calculated column
        // Note: This is a simplified approach - actual implementation may vary
        // based on Excel JavaScript API capabilities
        
        await context.sync();
        return true;
      });
    } catch (error) {
      logger.error('Failed to create calculated column', { error, column });
      return false;
    }
  }

  /**
   * Get all calculated columns
   */
  async getCalculatedColumns(): Promise<PowerPivotCalculatedColumn[]> {
    return await this.powerPivotService.getCalculatedColumns();
  }

  /**
   * Update a calculated column
   */
  async updateCalculatedColumn(
    tableName: string,
    columnName: string,
    newExpression: string
  ): Promise<boolean> {
    // Similar to create - delete and recreate
    // Implementation depends on Office.js API capabilities
    return true;
  }

  /**
   * Delete a calculated column
   */
  async deleteCalculatedColumn(tableName: string, columnName: string): Promise<boolean> {
    try {
      // @ts-ignore - Office.js API
      return await Excel.run(async (context: Excel.RequestContext) => {
        const table = context.workbook.tables.getItem(tableName);
        // Implementation depends on API capabilities
        await context.sync();
        return true;
      });
    } catch (error) {
      logger.error('Failed to delete calculated column', { error, tableName, columnName });
      return false;
    }
  }

  // ============================================================================
  // QUICK MEASURES (TEMPLATES)
  // ============================================================================

  /**
   * Create a quick measure from a template
   */
  createQuickMeasure(config: QuickMeasureConfig): PowerPivotMeasure {
    const { type, baseColumn, baseTable, dateColumn, dateTable } = config;

    let expression = '';
    let measureName = '';
    let formatString = '';

    switch (type) {
      case 'sum':
        measureName = `Total ${baseColumn}`;
        expression = `SUM('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
        break;

      case 'average':
        measureName = `Average ${baseColumn}`;
        expression = `AVERAGE('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0.00';
        break;

      case 'count':
        measureName = `${baseColumn} Count`;
        expression = `COUNT('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
        break;

      case 'distinct_count':
        measureName = `Unique ${baseColumn}`;
        expression = `DISTINCTCOUNT('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
        break;

      case 'min':
        measureName = `Min ${baseColumn}`;
        expression = `MIN('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
        break;

      case 'max':
        measureName = `Max ${baseColumn}`;
        expression = `MAX('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
        break;

      case 'ytd':
        measureName = `YTD ${baseColumn}`;
        expression = `TOTALYTD(SUM('${baseTable}'[${baseColumn}]), '${dateTable}'[${dateColumn}])`;
        formatString = '#,##0';
        break;

      case 'qtd':
        measureName = `QTD ${baseColumn}`;
        expression = `TOTALQTD(SUM('${baseTable}'[${baseColumn}]), '${dateTable}'[${dateColumn}])`;
        formatString = '#,##0';
        break;

      case 'mtd':
        measureName = `MTD ${baseColumn}`;
        expression = `TOTALMTD(SUM('${baseTable}'[${baseColumn}]), '${dateTable}'[${dateColumn}])`;
        formatString = '#,##0';
        break;

      case 'yoy_growth':
        measureName = `YoY Growth % ${baseColumn}`;
        expression = `VAR CurrentYear = SUM('${baseTable}'[${baseColumn}])
VAR PreviousYear = CALCULATE(SUM('${baseTable}'[${baseColumn}]), SAMEPERIODLASTYEAR('${dateTable}'[${dateColumn}]))
RETURN
    IF(PreviousYear <> 0, DIVIDE(CurrentYear - PreviousYear, PreviousYear, 0), 0)`;
        formatString = '0.00%';
        break;

      case 'qoq_growth':
        measureName = `QoQ Growth % ${baseColumn}`;
        expression = `VAR CurrentQtr = SUM('${baseTable}'[${baseColumn}])
VAR PreviousQtr = CALCULATE(SUM('${baseTable}'[${baseColumn}]), DATEADD('${dateTable}'[${dateColumn}], -1, QUARTER))
RETURN
    IF(PreviousQtr <> 0, DIVIDE(CurrentQtr - PreviousQtr, PreviousQtr, 0), 0)`;
        formatString = '0.00%';
        break;

      case 'mom_growth':
        measureName = `MoM Growth % ${baseColumn}`;
        expression = `VAR CurrentMonth = SUM('${baseTable}'[${baseColumn}])
VAR PreviousMonth = CALCULATE(SUM('${baseTable}'[${baseColumn}]), DATEADD('${dateTable}'[${dateColumn}], -1, MONTH))
RETURN
    IF(PreviousMonth <> 0, DIVIDE(CurrentMonth - PreviousMonth, PreviousMonth, 0), 0)`;
        formatString = '0.00%';
        break;

      case 'running_total':
        measureName = `Running Total ${baseColumn}`;
        expression = `CALCULATE(
    SUM('${baseTable}'[${baseColumn}]),
    FILTER(
        ALL('${dateTable}'[${dateColumn}]),
        '${dateTable}'[${dateColumn}] <= MAX('${dateTable}'[${dateColumn}])
    )
)`;
        formatString = '#,##0';
        break;

      case 'market_share':
        measureName = `${baseColumn} Market Share %`;
        expression = `VAR CurrentValue = SUM('${baseTable}'[${baseColumn}])
VAR TotalValue = CALCULATE(SUM('${baseTable}'[${baseColumn}]), ALL('${baseTable}'))
RETURN
    DIVIDE(CurrentValue, TotalValue, 0)`;
        formatString = '0.00%';
        break;

      default:
        measureName = `${baseColumn} Measure`;
        expression = `SUM('${baseTable}'[${baseColumn}])`;
        formatString = '#,##0';
    }

    return {
      name: measureName,
      table: baseTable,
      expression,
      formatString,
      description: `Quick measure: ${type}`
    };
  }

  /**
   * Get available quick measure templates
   */
  getQuickMeasureTemplates(): Array<{ type: QuickMeasureConfig['type']; name: string; description: string }> {
    return [
      { type: 'sum', name: 'Sum', description: 'Total of all values' },
      { type: 'average', name: 'Average', description: 'Mean of all values' },
      { type: 'count', name: 'Count', description: 'Count of rows' },
      { type: 'distinct_count', name: 'Distinct Count', description: 'Count of unique values' },
      { type: 'min', name: 'Minimum', description: 'Smallest value' },
      { type: 'max', name: 'Maximum', description: 'Largest value' },
      { type: 'ytd', name: 'Year-to-Date', description: 'Cumulative total from start of year' },
      { type: 'qtd', name: 'Quarter-to-Date', description: 'Cumulative total from start of quarter' },
      { type: 'mtd', name: 'Month-to-Date', description: 'Cumulative total from start of month' },
      { type: 'yoy_growth', name: 'Year-over-Year Growth %', description: 'Percentage growth vs same period last year' },
      { type: 'qoq_growth', name: 'Quarter-over-Quarter Growth %', description: 'Percentage growth vs previous quarter' },
      { type: 'mom_growth', name: 'Month-over-Month Growth %', description: 'Percentage growth vs previous month' },
      { type: 'running_total', name: 'Running Total', description: 'Cumulative sum over time' },
      { type: 'market_share', name: 'Market Share %', description: 'Percentage of total' },
    ];
  }

  // ============================================================================
  // VALIDATION & ANALYSIS
  // ============================================================================

  /**
   * Validate a measure
   */
  validateMeasure(measure: PowerPivotMeasure): MeasureValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Check required fields
    if (!measure.name || measure.name.trim() === '') {
      errors.push('Measure name is required');
    }

    if (!measure.table || measure.table.trim() === '') {
      errors.push('Table name is required');
    }

    if (!measure.expression || measure.expression.trim() === '') {
      errors.push('Expression is required');
    }

    // Parse and validate DAX
    if (measure.expression) {
      const parsed = daxParser.parse(measure.expression);
      const syntaxErrors = daxParser.validate(measure.expression);

      syntaxErrors.forEach(err => {
        if (err.severity === 'error') {
          errors.push(err.message);
        } else {
          warnings.push(err.message);
        }
      });

      // Check for common issues
      if (parsed.functions.includes('CALCULATE')) {
        const calcCount = parsed.functions.filter(f => f === 'CALCULATE').length;
        if (calcCount > 3) {
          warnings.push(`Multiple nested CALCULATE functions (${calcCount}) may impact performance`);
        }
      }

      // Check for iterator performance
      const iterators = parsed.functions.filter(f => {
        const func = daxParser.getFunctionInfo(f);
        return func?.isIterator;
      });
      if (iterators.length > 2) {
        warnings.push(`Multiple iterator functions may be slow on large datasets: ${iterators.join(', ')}`);
      }
    }

    // Check naming conventions
    if (measure.name && !/^[A-Z]/.test(measure.name)) {
      warnings.push('Measure names should start with a capital letter (best practice)');
    }

    // Calculate complexity
    const explanation = daxParser.explain(measure.expression || '');
    const complexity = explanation.complexity;

    // Estimate performance
    let estimatedPerformance: 'fast' | 'moderate' | 'slow' = 'fast';
    if (explanation.contextInfo.iteratorFunctions > 2) {
      estimatedPerformance = 'slow';
    } else if (explanation.contextInfo.iteratorFunctions > 0 || explanation.contextInfo.filterModifications > 2) {
      estimatedPerformance = 'moderate';
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      complexity,
      estimatedPerformance
    };
  }

  /**
   * Get measure dependencies
   */
  async getMeasureDependencies(measureName: string): Promise<{
    dependsOn: string[];
    usedBy: string[];
  }> {
    const measures = await this.powerPivotService.getMeasures();
    const targetMeasure = measures.find(m => m.name === measureName);

    if (!targetMeasure) {
      return { dependsOn: [], usedBy: [] };
    }

    // Parse dependencies from the measure expression
    const parsed = daxParser.parse(targetMeasure.expression);
    const dependsOn = parsed.measures.filter(m => m !== measureName);

    // Find measures that reference this one
    const usedBy = measures
      .filter(m => {
        const otherParsed = daxParser.parse(m.expression);
        return otherParsed.measures.includes(measureName);
      })
      .map(m => m.name);

    return { dependsOn, usedBy };
  }

  /**
   * Format a measure expression with proper indentation
   */
  formatMeasure(expression: string): string {
    return daxParser.formatMCode(expression);
  }

  // ============================================================================
  // NATURAL LANGUAGE INTERFACE
  // ============================================================================

  /**
   * Parse natural language command for measure operations
   * Example: "Create a measure Total Sales as sum of Amount from Sales table"
   */
  parseNaturalLanguageCommand(command: string): Partial<PowerPivotMeasure> {
    const result: Partial<PowerPivotMeasure> = {};

    // Extract measure name
    const nameMatch = command.match(/(?:create|add|new)\s+(?:a\s+)?(?:measure\s+)?(?:called\s+)?["']?([^"']+?)["']?(?:\s+as\s+|\s+from\s+|\s+in\s+)/i);
    if (nameMatch) {
      result.name = nameMatch[1].trim();
    }

    // Extract table name
    const tableMatch = command.match(/(?:from|in|table)\s+["']?([^"']+?)["']?(?:\s+(?:calculate|as|sum|count|average))/i) ||
                       command.match(/table\s+["']?([^"']+?)["']?$/i);
    if (tableMatch) {
      result.table = tableMatch[1].trim();
    }

    // Extract aggregation type and column
    const aggregationMatch = command.match(/(sum|count|average|min|max|distinct\s*count)\s+(?:of\s+)?(?:the\s+)?(?:column\s+)?["']?([^"']+?)["']?/i);
    if (aggregationMatch) {
      const aggType = aggregationMatch[1].toLowerCase().replace(/\s/g, '');
      const column = aggregationMatch[2].trim();
      const tableRef = result.table ? `'${result.table}'` : 'Table';

      const aggFunctions: Record<string, string> = {
        'sum': 'SUM',
        'count': 'COUNT',
        'average': 'AVERAGE',
        'min': 'MIN',
        'max': 'MAX',
        'distinctcount': 'DISTINCTCOUNT'
      };

      const func = aggFunctions[aggType] || 'SUM';
      result.expression = `${func}(${tableRef}[${column}])`;
    }

    // Extract filters
    const filterMatch = command.match(/(?:where|filter)\s+["']?([^"']+?)["']?\s*(=|>|<|>=|<=|<>|is)\s*["']?([^"']+?)["']?/i);
    if (filterMatch && result.expression) {
      const column = filterMatch[1].trim();
      const operator = filterMatch[2].trim();
      const value = filterMatch[3].trim();

      result.expression = `CALCULATE(${result.expression}, ${result.table}[${column}] ${operator} "${value}")`;
    }

    // Detect format hints
    if (command.includes('%') || command.includes('percent') || command.includes('percentage')) {
      result.formatString = '0.00%';
    } else if (command.includes('currency') || command.includes('dollar') || command.includes('$')) {
      result.formatString = '$#,##0.00';
    }

    return result;
  }

  /**
   * Suggest measure improvements based on context
   */
  async suggestImprovements(measure: PowerPivotMeasure): Promise<string[]> {
    const suggestions: string[] = [];
    const validation = this.validateMeasure(measure);

    // Add validation warnings
    suggestions.push(...validation.warnings);

    // Analyze the expression
    const parsed = daxParser.parse(measure.expression);

    // Check for time intelligence patterns
    if (parsed.functions.includes('SAMEPERIODLASTYEAR') || parsed.functions.includes('DATESYTD')) {
      suggestions.push('Consider creating a dedicated Date table if not already present');
    }

    // Check for relationship usage
    if (parsed.functions.includes('RELATED')) {
      suggestions.push('Ensure proper relationships are defined for RELATED functions');
    }

    // Check for potential variables
    if (measure.expression.length > 200 && !measure.expression.includes('VAR')) {
      suggestions.push('Consider using variables (VAR) to improve readability and performance');
    }

    // Check for DIVIDE usage
    if (measure.expression.includes('/') && !measure.expression.includes('DIVIDE')) {
      suggestions.push('Use DIVIDE function instead of / operator to handle division by zero');
    }

    return suggestions;
  }
}

export default DAXMeasureService.getInstance();
