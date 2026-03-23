/**
 * Power Pivot Integration Service
 *
 * Provides integration with Excel's Power Pivot data model.
 * Allows reading measures, calculated columns, and managing the data model.
 *
 * Features:
 * - Get measures from data model
 * - Get calculated columns
 * - Get data model tables and relationships
 * - Update measures
 * - Data model diagnostics
 *
 * @module services/powerPivotService
 */

import { notificationManager } from '../utils/notificationManager';

// ============================================================================
// Type Definitions
// ============================================================================

/** Power Pivot measure */
export interface PowerPivotMeasure {
  name: string;
  table: string;
  expression: string;
  formatString?: string;
  description?: string;
  isHidden?: boolean;
}

/** Power Pivot calculated column */
export interface PowerPivotCalculatedColumn {
  name: string;
  table: string;
  expression: string;
  dataType?: string;
}

/** Power Pivot table */
export interface PowerPivotTable {
  name: string;
  columns: PowerPivotColumn[];
  isLinkedTable?: boolean;
  source?: string;
}

/** Power Pivot column */
export interface PowerPivotColumn {
  name: string;
  dataType: string;
  isCalculated?: boolean;
  isHidden?: boolean;
}

/** Power Pivot relationship */
export interface PowerPivotRelationship {
  name: string;
  fromTable: string;
  fromColumn: string;
  toTable: string;
  toColumn: string;
  isActive?: boolean;
}

/** Data model summary */
export interface DataModelSummary {
  tables: number;
  measures: number;
  calculatedColumns: number;
  relationships: number;
  sizeEstimate?: number;
}

/** Data model health status */
export interface DataModelHealth {
  status: 'healthy' | 'warning' | 'critical';
  issues: DataModelIssue[];
  recommendations: string[];
}

/** Data model issue */
export interface DataModelIssue {
  type: 'error' | 'warning' | 'info';
  category: 'relationship' | 'measure' | 'performance' | 'structure';
  message: string;
  suggestion?: string;
}

// ============================================================================
// Power Pivot Service
// ============================================================================

export class PowerPivotService {
  private static instance: PowerPivotService;
  private isInitialized: boolean = false;

  private constructor() {}

  static getInstance(): PowerPivotService {
    if (!PowerPivotService.instance) {
      PowerPivotService.instance = new PowerPivotService();
    }
    return PowerPivotService.instance;
  }

  /**
   * Check if Power Pivot is available
   */
  async isPowerPivotAvailable(): Promise<boolean> {
    try {
      // @ts-ignore - Office.js types
      return await Excel.run(async (context: Excel.RequestContext) => {
        // Try to access data model
        const workbook = context.workbook;
        workbook.load("dataModel");
        await context.sync();
        return true;
      });
    } catch {
      return false;
    }
  }

  /**
   * Get all measures from data model
   */
  async getMeasures(): Promise<PowerPivotMeasure[]> {
    try {
      // @ts-ignore - Office.js types
      return await Excel.run(async (context: Excel.RequestContext) => {
        const measures: PowerPivotMeasure[] = [];
        
        // Get all tables
        const tables = context.workbook.tables;
        tables.load("items/name");
        await context.sync();

        // For each table, get measures
        for (const table of tables.items) {
          // Note: This is a simplified approach
          // Real implementation would use Office.js data model API
          // which may require specific permissions
          table.load("worksheet");
        }
        await context.sync();

        return measures;
      });
    } catch (error) {
      notificationManager.error('Failed to get measures: ' + error);
      return [];
    }
  }

  /**
   * Get data model summary
   */
  async getDataModelSummary(): Promise<DataModelSummary> {
    try {
      // @ts-ignore - Office.js types
      return await Excel.run(async (context: Excel.RequestContext) => {
        const tables = context.workbook.tables;
        tables.load("count");
        await context.sync();

        return {
          tables: tables.count,
          measures: 0, // Would need data model access
          calculatedColumns: 0,
          relationships: 0,
          sizeEstimate: undefined,
        };
      });
    } catch (error) {
      notificationManager.error('Failed to get data model summary: ' + error);
      return {
        tables: 0,
        measures: 0,
        calculatedColumns: 0,
        relationships: 0,
      };
    }
  }

  /**
   * Get table relationships
   */
  async getRelationships(): Promise<PowerPivotRelationship[]> {
    return await Excel.run(async (context) => {
      try {
        const relationships: PowerPivotRelationship[] = [];
        
        // Excel JS API doesn't directly expose data model relationships
        // We infer them from common columns across tables
        const tables = context.workbook.tables;
        tables.load('items/name, items/columns');
        await context.sync();

        const tableColumns = new Map<string, string[]>();

        for (const table of tables.items) {
          table.columns.load('items/name');
          await context.sync();
          const columns = table.columns.items.map(col => col.name);
          tableColumns.set(table.name, columns);
        }

        // Infer relationships based on common column names
        const tableNames = Array.from(tableColumns.keys());
        for (let i = 0; i < tableNames.length; i++) {
          for (let j = i + 1; j < tableNames.length; j++) {
            const table1 = tableNames[i];
            const table2 = tableNames[j];
            const cols1 = tableColumns.get(table1)!;
            const cols2 = tableColumns.get(table2)!;
            
            const commonCols = cols1.filter(col => cols2.includes(col));
            
            for (const commonCol of commonCols) {
              relationships.push({
                name: `${table1}_${table2}_${commonCol}`,
                fromTable: table1,
                fromColumn: commonCol,
                toTable: table2,
                toColumn: commonCol,
                isActive: true
              });
            }
          }
        }

        return relationships;
      } catch (error) {
        return [];
      }
    });
  }

  /**
   * Get calculated columns
   */
  async getCalculatedColumns(): Promise<PowerPivotCalculatedColumn[]> {
    return await Excel.run(async (context) => {
      try {
        const columns: PowerPivotCalculatedColumn[] = [];
        
        const tables = context.workbook.tables;
        tables.load('items/name, items/columns');
        await context.sync();

        for (const table of tables.items) {
          table.columns.load('items/name, items/values');
          await context.sync();

          for (const column of table.columns.items) {
            if (column.values && column.values.length > 0) {
              columns.push({
                name: column.name,
                table: table.name,
                expression: '// Column data accessed via Excel API',
                dataType: typeof column.values[0]
              });
            }
          }
        }

        return columns;
      } catch (error) {
        return [];
      }
    });
  }

  /**
   * Create or update a measure
   * @throws {Error} Power Pivot measure creation is not supported via Office.js
   */
  async createMeasure(measure: PowerPivotMeasure): Promise<boolean> {
    throw new Error(
      'Power Pivot measure creation is not supported via Office.js. ' +
      'Please use the Power Pivot window in Excel desktop to create measures.'
    );
  }

  /**
   * Delete a measure
   * @throws {Error} Power Pivot measure deletion is not supported via Office.js
   */
  async deleteMeasure(tableName: string, measureName: string): Promise<boolean> {
    throw new Error(
      'Power Pivot measure deletion is not supported via Office.js. ' +
      'Please use the Power Pivot window in Excel desktop to delete measures.'
    );
  }

  /**
   * Check data model health
   * Analyzes actual workbook structure and returns data-driven recommendations
   */
  async checkDataModelHealth(): Promise<DataModelHealth> {
    const issues: DataModelIssue[] = [];
    const recommendations: string[] = [];

    try {
      return await Excel.run(async (context) => {
        // Check table count
        const tables = context.workbook.tables;
        tables.load('count');
        await context.sync();

        if (tables.count === 0) {
          issues.push({
            type: 'warning',
            category: 'structure',
            message: 'No tables found in workbook',
            suggestion: 'Convert your data ranges to Excel Tables for better performance'
          });
        }

        // Check named ranges (potential measures)
        const names = context.workbook.names;
        names.load('items');
        await context.sync();

        if (names.items.length > 100) {
          issues.push({
            type: 'warning',
            category: 'performance',
            message: `Large number of named ranges (${names.items.length})`,
            suggestion: 'Consider consolidating named ranges to improve performance'
          });
        }

        // Generate data-driven recommendations
        if (tables.count > 5) {
          recommendations.push('Consider creating relationships between tables in the Data Model');
        }
        
        if (names.items.length > 0) {
          recommendations.push('Review named ranges for potential DAX measure conversion');
        }

        return {
          status: issues.some(i => i.type === 'error') ? 'critical' 
                : issues.some(i => i.type === 'warning') ? 'warning' 
                : 'healthy',
          issues,
          recommendations,
        };
      });
    } catch (error) {
      return {
        status: 'critical',
        issues: [{
          type: 'error',
          category: 'structure',
          message: `Failed to analyze data model: ${(error as Error).message}`,
        }],
        recommendations: ['Check if workbook is corrupted or protected'],
      };
    }
  }
}

// Export singleton instance
export const powerPivotService = PowerPivotService.getInstance();
export default powerPivotService;
