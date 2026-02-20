// Pivot Table Service - Professional-grade Pivot Table operations with Natural Language support
// Phase 2 Implementation - Complete Pivot Table Operations

import { ExcelService } from './excelService';

export interface PivotTableField {
  name: string;
  orientation: 'row' | 'column' | 'data' | 'filter';
  aggregation?: 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'countNumbers' | 'stdDev' | 'stdDevp' | 'var' | 'varp';
  showValuesAs?: 'none' | 'percentOfGrandTotal' | 'percentOfColumnTotal' | 'percentOfRowTotal' | 'percentOfParentRowTotal' | 'percentOfParentColumnTotal' | 'percentOfParentTotal' | 'differenceFrom' | 'percentDifferenceFrom' | 'runningTotal' | 'rankAscending' | 'rankDescending';
  baseField?: string;
  baseItem?: string;
  numberFormat?: string;
  caption?: string;
}

export interface PivotTableLayout {
  form: 'compact' | 'outline' | 'tabular';
  showSubtotals: boolean;
  subtotalPosition: 'bottom' | 'top';
  showGrandTotalsForRows: boolean;
  showGrandTotalsForColumns: boolean;
  blankRowsAfterItems: boolean;
  repeatItemLabels: boolean;
}

export interface PivotTableConfig {
  name: string;
  sourceData: string;
  destination?: string;
  fields: PivotTableField[];
  layout?: PivotTableLayout;
}

export interface CalculatedField {
  name: string;
  formula: string;
  numberFormat?: string;
}

export interface CalculatedItem {
  name: string;
  fieldName: string;
  formula: string;
}

export interface PivotTableGroupConfig {
  fieldName: string;
  by?: 'days' | 'months' | 'quarters' | 'years' | number[];
  start?: number | Date;
  end?: number | Date;
}

export interface PivotTableInfo {
  name: string;
  worksheetName: string;
  sourceData: string;
  rowFields: string[];
  columnFields: string[];
  dataFields: Array<{ name: string; aggregation: string }>;
  filterFields: string[];
  layout: PivotTableLayout;
}

export class PivotTableService {
  private static instance: PivotTableService;
  private excelService: typeof ExcelService;

  private constructor() {
    this.excelService = ExcelService;
  }

  static getInstance(): PivotTableService {
    if (!PivotTableService.instance) {
      PivotTableService.instance = new PivotTableService();
    }
    return PivotTableService.instance;
  }

  // ============================================================================
  // CREATION & STRUCTURE
  // ============================================================================

  /**
   * Create a new Pivot Table with full field configuration
   */
  async createPivotTable(config: PivotTableConfig, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Get source range
      const sourceRange = worksheet.getRange(config.sourceData);

      // Create pivot table
      const pivotTable = worksheet.pivotTables.add(
        config.name,
        sourceRange,
        config.destination ? worksheet.getRange(config.destination) : undefined
      );

      await context.sync();

      // Configure fields
      for (const field of config.fields) {
        await this.addField(context, pivotTable, field);
      }

      // Apply layout options
      if (config.layout) {
        await this.applyLayout(context, pivotTable, config.layout);
      }

      await context.sync();
    });
  }

  /**
   * Add a field to pivot table
   */
  private async addField(context: Excel.RequestContext, pivotTable: Excel.PivotTable, field: PivotTableField): Promise<void> {
    const pivotField = pivotTable.hierarchies.getItemOrNullObject(field.name);
    await context.sync();

    if (pivotField.isNullObject) {
      throw new Error(`Field "${field.name}" not found in source data`);
    }

    switch (field.orientation) {
      case 'row':
        pivotTable.rowHierarchies.add(pivotField);
        break;
      case 'column':
        pivotTable.columnHierarchies.add(pivotField);
        break;
      case 'data':
        const dataHierarchy = pivotTable.dataHierarchies.add(pivotField);
        if (field.aggregation) {
          dataHierarchy.summarizeBy = this.mapAggregation(field.aggregation);
        }
        if (field.numberFormat) {
          dataHierarchy.numberFormat = field.numberFormat;
        }
        if (field.caption) {
          dataHierarchy.caption = field.caption;
        }
        // Note: showValuesAs requires more complex implementation with PivotField APIs
        break;
      case 'filter':
        pivotTable.filterHierarchies.add(pivotField);
        break;
    }

    await context.sync();
  }

  /**
   * Map aggregation string to Excel.AggregationFunction
   */
  private mapAggregation(agg: string): Excel.AggregationFunction {
    const mapping: Record<string, Excel.AggregationFunction> = {
      'sum': Excel.AggregationFunction.sum,
      'count': Excel.AggregationFunction.count,
      'average': Excel.AggregationFunction.average,
      'max': Excel.AggregationFunction.max,
      'min': Excel.AggregationFunction.min,
      'product': Excel.AggregationFunction.product,
      'countNumbers': Excel.AggregationFunction.countNumbers,
      'stdDev': Excel.AggregationFunction.standardDeviation,
      'stdDevp': Excel.AggregationFunction.standardDeviationP,
      'var': Excel.AggregationFunction.variance,
      'varp': Excel.AggregationFunction.varianceP
    };
    return mapping[agg] || Excel.AggregationFunction.sum;
  }

  /**
   * Apply layout configuration to pivot table
   */
  private async applyLayout(context: Excel.RequestContext, pivotTable: Excel.PivotTable, layout: PivotTableLayout): Promise<void> {
    // Layout form (compact/outline/tabular)
    if (layout.form) {
      pivotTable.layout.load();
      await context.sync();
      // Note: Office.js has limited layout API, some options may require UI
    }

    // Grand totals
    pivotTable.rowHierarchies.load('items');
    pivotTable.columnHierarchies.load('items');
    await context.sync();

    // Configure row field subtotals
    for (const rowHierarchy of pivotTable.rowHierarchies.items) {
      rowHierarchy.load('fields');
      await context.sync();
      for (const field of rowHierarchy.fields.items) {
        field.showSubtotals = layout.showSubtotals;
      }
    }

    await context.sync();
  }

  // ============================================================================
  // FIELD MANAGEMENT
  // ============================================================================

  /**
   * Add a field to an existing pivot table
   */
  async addFieldToPivotTable(
    pivotName: string,
    field: PivotTableField,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      await this.addField(context, pivotTable, field);
      await context.sync();
    });
  }

  /**
   * Remove a field from pivot table
   */
  async removeField(
    pivotName: string,
    fieldName: string,
    orientation: 'row' | 'column' | 'data' | 'filter',
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      switch (orientation) {
        case 'row':
          const rowHierarchy = pivotTable.rowHierarchies.getItem(fieldName);
          rowHierarchy.remove();
          break;
        case 'column':
          const colHierarchy = pivotTable.columnHierarchies.getItem(fieldName);
          colHierarchy.remove();
          break;
        case 'data':
          const dataHierarchy = pivotTable.dataHierarchies.getItem(fieldName);
          dataHierarchy.remove();
          break;
        case 'filter':
          const filterHierarchy = pivotTable.filterHierarchies.getItem(fieldName);
          filterHierarchy.remove();
          break;
      }

      await context.sync();
    });
  }

  /**
   * Move a field from one area to another
   */
  async moveField(
    pivotName: string,
    fieldName: string,
    fromOrientation: 'row' | 'column' | 'data' | 'filter',
    toOrientation: 'row' | 'column' | 'data' | 'filter',
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      // Get the hierarchy
      let hierarchy: Excel.PivotHierarchy;
      switch (fromOrientation) {
        case 'row':
          hierarchy = pivotTable.rowHierarchies.getItem(fieldName);
          break;
        case 'column':
          hierarchy = pivotTable.columnHierarchies.getItem(fieldName);
          break;
        case 'data':
          hierarchy = pivotTable.dataHierarchies.getItem(fieldName);
          break;
        case 'filter':
          hierarchy = pivotTable.filterHierarchies.getItem(fieldName);
          break;
      }

      // Remove from current location
      hierarchy.remove();
      await context.sync();

      // Add to new location
      const newField: PivotTableField = {
        name: fieldName,
        orientation: toOrientation
      };
      await this.addField(context, pivotTable, newField);

      await context.sync();
    });
  }

  /**
   * Change aggregation function for a data field
   */
  async changeAggregation(
    pivotName: string,
    fieldName: string,
    aggregation: PivotTableField['aggregation'],
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const dataHierarchy = pivotTable.dataHierarchies.getItem(fieldName);

      dataHierarchy.summarizeBy = this.mapAggregation(aggregation || 'sum');
      await context.sync();
    });
  }

  /**
   * Set number format for a data field
   */
  async setNumberFormat(
    pivotName: string,
    fieldName: string,
    format: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const dataHierarchy = pivotTable.dataHierarchies.getItem(fieldName);

      dataHierarchy.numberFormat = format;
      await context.sync();
    });
  }

  // ============================================================================
  // CALCULATED FIELDS & ITEMS
  // ============================================================================

  /**
   * Add a calculated field to pivot table
   */
  async addCalculatedField(
    pivotName: string,
    calculatedField: CalculatedField,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      // Office.js API for calculated fields is limited
      // This uses the data hierarchy API which may have limitations
      // For full support, may need to use Office Scripts or VBA fallback

      const dataHierarchy = pivotTable.dataHierarchies.addCalculated(
        calculatedField.name,
        calculatedField.formula
      );

      if (calculatedField.numberFormat) {
        dataHierarchy.numberFormat = calculatedField.numberFormat;
      }

      await context.sync();
    });
  }

  /**
   * Remove a calculated field
   */
  async removeCalculatedField(
    pivotName: string,
    fieldName: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const dataHierarchy = pivotTable.dataHierarchies.getItem(fieldName);

      dataHierarchy.remove();
      await context.sync();
    });
  }

  // ============================================================================
  // GROUPING
  // ============================================================================

  /**
   * Group date/number field in pivot table
   */
  async groupField(
    pivotName: string,
    groupConfig: PivotTableGroupConfig,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      // Find the field in either rows or columns
      let field: Excel.PivotField | null = null;

      try {
        const rowHierarchy = pivotTable.rowHierarchies.getItem(groupConfig.fieldName);
        rowHierarchy.load('fields');
        await context.sync();
        field = rowHierarchy.fields.items[0];
      } catch {
        try {
          const colHierarchy = pivotTable.columnHierarchies.getItem(groupConfig.fieldName);
          colHierarchy.load('fields');
          await context.sync();
          field = colHierarchy.fields.items[0];
        } catch {
          throw new Error(`Field "${groupConfig.fieldName}" not found in pivot table`);
        }
      }

      if (!field) {
        throw new Error(`Field "${groupConfig.fieldName}" not found`);
      }

      // Apply grouping
      if (groupConfig.by) {
        if (Array.isArray(groupConfig.by)) {
          // Number grouping with intervals
          field.groupBy({
            startingAt: groupConfig.start as number,
            endingAt: groupConfig.end as number,
            by: groupConfig.by[0]
          });
        } else {
          // Date grouping
          field.groupBy({
            by: groupConfig.by as any
          });
        }
      }

      await context.sync();
    });
  }

  /**
   * Ungroup a field
   */
  async ungroupField(
    pivotName: string,
    fieldName: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      let field: Excel.PivotField | null = null;

      try {
        const rowHierarchy = pivotTable.rowHierarchies.getItem(fieldName);
        rowHierarchy.load('fields');
        await context.sync();
        field = rowHierarchy.fields.items[0];
      } catch {
        try {
          const colHierarchy = pivotTable.columnHierarchies.getItem(fieldName);
          colHierarchy.load('fields');
          await context.sync();
          field = colHierarchy.fields.items[0];
        } catch {
          throw new Error(`Field "${fieldName}" not found in pivot table`);
        }
      }

      if (field) {
        field.ungroup();
      }

      await context.sync();
    });
  }

  // ============================================================================
  // FILTERING
  // ============================================================================

  /**
   * Apply filter to a pivot field
   */
  async applyFilter(
    pivotName: string,
    fieldName: string,
    items: string[],
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      // Try to find field in filter hierarchies first
      let field: Excel.PivotField | null = null;

      try {
        const filterHierarchy = pivotTable.filterHierarchies.getItem(fieldName);
        filterHierarchy.load('fields');
        await context.sync();
        field = filterHierarchy.fields.items[0];
      } catch {
        // Try rows
        try {
          const rowHierarchy = pivotTable.rowHierarchies.getItem(fieldName);
          rowHierarchy.load('fields');
          await context.sync();
          field = rowHierarchy.fields.items[0];
        } catch {
          // Try columns
          try {
            const colHierarchy = pivotTable.columnHierarchies.getItem(fieldName);
            colHierarchy.load('fields');
            await context.sync();
            field = colHierarchy.fields.items[0];
          } catch {
            throw new Error(`Field "${fieldName}" not found in pivot table`);
          }
        }
      }

      if (field) {
        field.applyFilter({
          manualFilter: {
            selectedItems: items
          }
        });
      }

      await context.sync();
    });
  }

  /**
   * Clear filter from a pivot field
   */
  async clearFilter(
    pivotName: string,
    fieldName: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      let field: Excel.PivotField | null = null;

      try {
        const filterHierarchy = pivotTable.filterHierarchies.getItem(fieldName);
        filterHierarchy.load('fields');
        await context.sync();
        field = filterHierarchy.fields.items[0];
      } catch {
        try {
          const rowHierarchy = pivotTable.rowHierarchies.getItem(fieldName);
          rowHierarchy.load('fields');
          await context.sync();
          field = rowHierarchy.fields.items[0];
        } catch {
          try {
            const colHierarchy = pivotTable.columnHierarchies.getItem(fieldName);
            colHierarchy.load('fields');
            await context.sync();
            field = colHierarchy.fields.items[0];
          } catch {
            throw new Error(`Field "${fieldName}" not found in pivot table`);
          }
        }
      }

      if (field) {
        field.clearFilter();
      }

      await context.sync();
    });
  }

  // ============================================================================
  // REFRESH & DATA
  // ============================================================================

  /**
   * Refresh a specific pivot table
   */
  async refreshPivotTable(pivotName: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      pivotTable.refresh();
      await context.sync();
    });
  }

  /**
   * Refresh all pivot tables in the workbook
   */
  async refreshAllPivotTables(): Promise<void> {
    return Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load('items');
      await context.sync();

      for (const worksheet of worksheets.items) {
        const pivotTables = worksheet.pivotTables;
        pivotTables.load('items');
        await context.sync();

        for (const pivotTable of pivotTables.items) {
          pivotTable.refresh();
        }
      }

      await context.sync();
    });
  }

  /**
   * Change the source data range for a pivot table
   */
  async changeSourceData(
    pivotName: string,
    newSourceData: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const newRange = worksheet.getRange(newSourceData);

      pivotTable.changeDataSource(newRange);
      await context.sync();
    });
  }

  // ============================================================================
  // PIVOT CHART
  // ============================================================================

  /**
   * Create a pivot chart from a pivot table
   */
  async createPivotChart(
    pivotName: string,
    chartType: Excel.ChartType,
    title?: string,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const chart = worksheet.charts.add(chartType, pivotTable.getRange(), 'Auto');

      if (title) {
        chart.title.text = title;
      }

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  // ============================================================================
  // INFORMATION & ANALYSIS
  // ============================================================================

  /**
   * Get detailed information about a pivot table
   */
  async getPivotTableInfo(pivotName: string, worksheetName?: string): Promise<PivotTableInfo> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);

      pivotTable.load('name');
      pivotTable.rowHierarchies.load('items/name');
      pivotTable.columnHierarchies.load('items/name');
      pivotTable.dataHierarchies.load('items/name, items/summarizeBy');
      pivotTable.filterHierarchies.load('items/name');

      await context.sync();

      return {
        name: pivotTable.name,
        worksheetName: worksheet.name,
        sourceData: '', // API limitation
        rowFields: pivotTable.rowHierarchies.items.map(h => h.name),
        columnFields: pivotTable.columnHierarchies.items.map(h => h.name),
        dataFields: pivotTable.dataHierarchies.items.map(h => ({
          name: h.name,
          aggregation: Excel.AggregationFunction[h.summarizeBy] || 'sum'
        })),
        filterFields: pivotTable.filterHierarchies.items.map(h => h.name),
        layout: {
          form: 'compact',
          showSubtotals: true,
          subtotalPosition: 'bottom',
          showGrandTotalsForRows: true,
          showGrandTotalsForColumns: true,
          blankRowsAfterItems: false,
          repeatItemLabels: false
        }
      };
    });
  }

  /**
   * Get pivot table data as array
   */
  async getPivotData(pivotName: string, worksheetName?: string): Promise<any[][]> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const range = pivotTable.getRange();
      range.load('values');
      await context.sync();

      return range.values;
    });
  }

  /**
   * Get a specific value from pivot table
   */
  async getPivotValue(
    pivotName: string,
    rowCriteria: Record<string, string>,
    columnCriteria?: Record<string, string>,
    dataField?: string,
    worksheetName?: string
  ): Promise<any> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      const range = pivotTable.getRange();
      range.load('values, rowCount, columnCount');
      await context.sync();

      // This is a simplified implementation
      // Full implementation would need to match criteria against row/column labels
      const values = range.values;

      // Find the value based on criteria (simplified)
      // Real implementation would parse the pivot table structure
      return values;
    });
  }

  // ============================================================================
  // DELETE
  // ============================================================================

  /**
   * Delete a pivot table
   */
  async deletePivotTable(pivotName: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(pivotName);
      pivotTable.delete();
      await context.sync();
    });
  }

  // ============================================================================
  // NATURAL LANGUAGE PARSING
  // ============================================================================

  /**
   * Parse natural language command for pivot table creation
   * Example: "Create pivot table from Sales data with Region in rows and Sum of Amount in values"
   */
  parseNaturalLanguageCommand(command: string): Partial<PivotTableConfig> {
    const config: Partial<PivotTableConfig> = {
      fields: []
    };

    // Extract source data
    const sourceMatch = command.match(/from\s+([\w\s]+?)(?:\s+with|\s+showing|\s+having|$)/i);
    if (sourceMatch) {
      config.sourceData = sourceMatch[1].trim();
    }

    // Extract row fields
    const rowMatch = command.match(/(\w+)\s+in\s+rows?/gi);
    if (rowMatch) {
      rowMatch.forEach(match => {
        const fieldName = match.replace(/\s+in\s+rows?/i, '').trim();
        config.fields!.push({
          name: fieldName,
          orientation: 'row'
        });
      });
    }

    // Extract column fields
    const colMatch = command.match(/(\w+)\s+in\s+columns?/gi);
    if (colMatch) {
      colMatch.forEach(match => {
        const fieldName = match.replace(/\s+in\s+columns?/i, '').trim();
        config.fields!.push({
          name: fieldName,
          orientation: 'column'
        });
      });
    }

    // Extract data/values fields with aggregation
    const valueMatch = command.match(/(sum|count|average|max|min)?\s*of?\s*(\w+)\s+in\s+(?:values?|data)/gi);
    if (valueMatch) {
      valueMatch.forEach(match => {
        const aggMatch = match.match(/(sum|count|average|max|min)/i);
        const fieldMatch = match.match(/of\s+(\w+)/i);
        config.fields!.push({
          name: fieldMatch ? fieldMatch[1] : match.replace(/\s+in\s+(?:values?|data)/i, '').trim(),
          orientation: 'data',
          aggregation: (aggMatch ? aggMatch[1].toLowerCase() : 'sum') as any
        });
      });
    }

    // Extract filter fields
    const filterMatch = command.match(/(?:filter|show)\s+(\w+)\s*=\s*['"]([^'"]+)['"]/gi);
    if (filterMatch) {
      filterMatch.forEach(match => {
        const parts = match.match(/(?:filter|show)\s+(\w+)\s*=/i);
        if (parts) {
          config.fields!.push({
            name: parts[1].trim(),
            orientation: 'filter'
          });
        }
      });
    }

    return config;
  }
}

export default PivotTableService.getInstance();
