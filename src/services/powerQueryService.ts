/**
 * Excel AI Assistant - Power Query Service
 * Excel integration for Power Query operations
 * 
 * @module services/powerQueryService
 */

import { QueryMetadata, QueryStep, mCodeGenerator } from "./mCodeGenerator";

// ============================================================================
// Type Definitions
// ============================================================================

/** Power Query information from Excel */
export interface PowerQueryInfo {
  name: string;
  formula: string;
  connection?: string;
  description?: string;
  resultType?: string;
}

/** Query execution result */
export interface QueryExecutionResult {
  success: boolean;
  data?: any[][];
  columnNames?: string[];
  rowCount?: number;
  error?: string;
  executionTime?: number;
}

/** Query modification result */
export interface QueryModificationResult {
  success: boolean;
  queryName?: string;
  newFormula?: string;
  error?: string;
  warnings?: string[];
}

/** Query refresh options */
export interface RefreshOptions {
  background?: boolean;
  timeout?: number;
  onProgress?: (percent: number) => void;
}

/** Dependency graph for queries */
export interface QueryDependencyGraph {
  queries: string[];
  dependencies: Map<string, string[]>;
  rootQueries: string[];
}

/** Data source information */
export interface DataSourceInfo {
  type: "excel" | "csv" | "sql" | "odata" | "web" | "folder" | "other";
  location: string;
  authentication?: string;
  isRefreshable: boolean;
}

/** Query performance metrics */
export interface QueryPerformanceMetrics {
  queryName: string;
  loadTime: number;
  rowCount: number;
  columnCount: number;
  memoryUsage?: number;
  lastRefresh?: Date;
}

// ============================================================================
// Power Query Service
// ============================================================================

export class PowerQueryService {
  private static instance: PowerQueryService;
  private queryCache: Map<string, QueryMetadata> = new Map();

  private constructor() {}

  static getInstance(): PowerQueryService {
    if (!PowerQueryService.instance) {
      PowerQueryService.instance = new PowerQueryService();
    }
    return PowerQueryService.instance;
  }

  // ============================================================================
  // Query Discovery
  // ============================================================================

  /**
   * Get all Power Queries in the current workbook
   */
  async getAllQueries(): Promise<PowerQueryInfo[]> {
    return new Promise((resolve, reject) => {
      try {
        // @ts-ignore - Office.js types not available during dev
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const queries = context.workbook.queries;
          queries.load("items");
          await context.sync();

          const queryInfos: PowerQueryInfo[] = queries.items.map((q: any) => ({
            name: q.name,
            formula: q.formula,
            connection: q.connection,
            description: q.description,
            resultType: q.resultType
          }));

          resolve(queryInfos);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Get a specific query by name
   */
  async getQuery(name: string): Promise<PowerQueryInfo | null> {
    return new Promise((resolve, reject) => {
      try {
        // @ts-ignore
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const queries = context.workbook.queries;
          const query = queries.getItemOrNullObject(name);
          query.load("name, formula, connection, description, resultType");
          await context.sync();

          if (query.isNullObject) {
            resolve(null);
          } else {
            resolve({
              name: query.name,
              formula: query.formula,
              connection: query.connection,
              description: query.description,
              resultType: query.resultType
            });
          }
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Get query metadata including parsed steps
   */
  async getQueryMetadata(name: string): Promise<QueryMetadata | null> {
    // Check cache first
    if (this.queryCache.has(name)) {
      return this.queryCache.get(name)!;
    }

    const query = await this.getQuery(name);
    if (!query) return null;

    const steps = mCodeGenerator.parseQuery(query.formula);
    const columns = await this.getQueryColumns(name);

    const metadata: QueryMetadata = {
      name: query.name,
      source: this.detectSourceType(query.formula),
      steps: steps.map((s, i) => ({
        name: s.name,
        mCode: s.expression,
        lineNumber: i + 2 // Approximate line number
      })),
      columns,
      isLoaded: true
    };

    this.queryCache.set(name, metadata);
    return metadata;
  }

  /**
   * Detect the source type from M code
   */
  private detectSourceType(mCode: string): string {
    const lower = mCode.toLowerCase();
    if (lower.includes("excel.workbook")) return "Excel Workbook";
    if (lower.includes("csv.document")) return "CSV File";
    if (lower.includes("sql.database")) return "SQL Database";
    if (lower.includes("odata.feed")) return "OData Feed";
    if (lower.includes("web.contents")) return "Web Page";
    if (lower.includes("folder.files")) return "Folder";
    if (lower.includes("sharepoint.contents")) return "SharePoint";
    return "Unknown";
  }

  /**
   * Get columns from a query result
   */
  private async getQueryColumns(queryName: string): Promise<any[]> {
    // This would need to be implemented based on actual Excel API capabilities
    // For now, return empty array
    return [];
  }

  // ============================================================================
  // Query Creation & Modification
  // ============================================================================

  /**
   * Create a new Power Query
   */
  async createQuery(
    name: string,
    mCode: string,
    description?: string
  ): Promise<QueryModificationResult> {
    return new Promise((resolve, reject) => {
      try {
        // Validate M code first
        const validation = mCodeGenerator.validateMCode(mCode);
        if (!validation.isValid) {
          resolve({
            success: false,
            error: `Invalid M code: ${validation.errors[0].message}`,
            warnings: validation.warnings.map(w => w.message)
          });
          return;
        }

        // @ts-ignore
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const queries = context.workbook.queries;

          // Check if query already exists
          const existingQuery = queries.getItemOrNullObject(name);
          existingQuery.load();
          await context.sync();

          if (!existingQuery.isNullObject) {
            resolve({
              success: false,
              error: `Query "${name}" already exists. Use updateQuery to modify it.`
            });
            return;
          }

          // Create new query
          queries.add(name, mCode, description || "");
          await context.sync();

          // Clear cache
          this.queryCache.delete(name);

          resolve({
            success: true,
            queryName: name,
            newFormula: mCode,
            warnings: validation.warnings.map(w => w.message)
          });
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Update an existing Power Query
   */
  async updateQuery(
    name: string,
    newMCode: string,
    description?: string
  ): Promise<QueryModificationResult> {
    return new Promise((resolve, reject) => {
      try {
        // Validate M code
        const validation = mCodeGenerator.validateMCode(newMCode);
        if (!validation.isValid) {
          resolve({
            success: false,
            error: `Invalid M code: ${validation.errors[0].message}`,
            warnings: validation.warnings.map(w => w.message)
          });
          return;
        }

        // @ts-ignore
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const queries = context.workbook.queries;

          // Get existing query
          const query = queries.getItemOrNullObject(name);
          query.load();
          await context.sync();

          if (query.isNullObject) {
            resolve({
              success: false,
              error: `Query "${name}" not found. Use createQuery to create it.`
            });
            return;
          }

          // Update query formula
          query.formula = newMCode;
          if (description) {
            query.description = description;
          }

          await context.sync();

          // Clear cache
          this.queryCache.delete(name);

          resolve({
            success: true,
            queryName: name,
            newFormula: newMCode,
            warnings: validation.warnings.map(w => w.message)
          });
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Delete a Power Query
   */
  async deleteQuery(name: string): Promise<boolean> {
    return new Promise((resolve, reject) => {
      try {
        // @ts-ignore
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const queries = context.workbook.queries;
          const query = queries.getItemOrNullObject(name);
          query.load();
          await context.sync();

          if (query.isNullObject) {
            resolve(false);
            return;
          }

          query.delete();
          await context.sync();

          // Clear cache
          this.queryCache.delete(name);

          resolve(true);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  /**
   * Add a step to an existing query
   */
  async addStepToQuery(
    queryName: string,
    stepName: string,
    stepExpression: string
  ): Promise<QueryModificationResult> {
    try {
      const query = await this.getQuery(queryName);
      if (!query) {
        return {
          success: false,
          error: `Query "${queryName}" not found`
        };
      }

      const steps = mCodeGenerator.parseQuery(query.formula);
      steps.push({ name: stepName, expression: stepExpression });

      const newMCode = mCodeGenerator.buildQuery(steps);
      return this.updateQuery(queryName, newMCode);
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : "Unknown error"
      };
    }
  }

  /**
   * Remove a step from a query
   */
  async removeStepFromQuery(
    queryName: string,
    stepName: string
  ): Promise<QueryModificationResult> {
    try {
      const query = await this.getQuery(queryName);
      if (!query) {
        return {
          success: false,
          error: `Query "${queryName}" not found`
        };
      }

      const steps = mCodeGenerator.parseQuery(query.formula);
      const filteredSteps = steps.filter(s => s.name !== stepName);

      if (filteredSteps.length === steps.length) {
        return {
          success: false,
          error: `Step "${stepName}" not found in query`
        };
      }

      const newMCode = mCodeGenerator.buildQuery(filteredSteps);
      return this.updateQuery(queryName, newMCode);
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : "Unknown error"
      };
    }
  }

  // ============================================================================
  // Query Execution & Refresh
  // ============================================================================

  /**
   * Refresh a specific query
   */
  async refreshQuery(
    queryName: string,
    options?: RefreshOptions
  ): Promise<QueryExecutionResult> {
    return new Promise((resolve, reject) => {
      const startTime = Date.now();

      try {
        // @ts-ignore
        Excel.run(async (context: Excel.RequestContext) => {
          // @ts-ignore
          const workbook = context.workbook;
          const connections = workbook.connections;
          connections.load("items");
          await context.sync();

          // Find connection associated with the query
          const connection = connections.items.find((c: any) =>
            c.name === queryName || c.name.includes(queryName)
          );

          if (!connection) {
            resolve({
              success: false,
              error: `No connection found for query "${queryName}"`
            });
            return;
          }

          // Refresh the connection
          connection.refresh();
          await context.sync();

          const executionTime = Date.now() - startTime;

          resolve({
            success: true,
            executionTime
          });
        });
      } catch (error) {
        resolve({
          success: false,
          error: error instanceof Error ? error.message : "Refresh failed"
        });
      }
    });
  }

  /**
   * Refresh all queries in the workbook
   */
  async refreshAllQueries(options?: RefreshOptions): Promise<{
    success: boolean;
    results: Map<string, QueryExecutionResult>;
    totalTime: number;
  }> {
    const startTime = Date.now();
    const results = new Map<string, QueryExecutionResult>();

    try {
      const queries = await this.getAllQueries();

      for (const query of queries) {
        if (options?.onProgress) {
          options.onProgress(Math.round((results.size / queries.length) * 100));
        }

        const result = await this.refreshQuery(query.name, options);
        results.set(query.name, result);
      }

      return {
        success: true,
        results,
        totalTime: Date.now() - startTime
      };
    } catch (error) {
      return {
        success: false,
        results,
        totalTime: Date.now() - startTime
      };
    }
  }

  // ============================================================================
  // Dependency Analysis
  // ============================================================================

  /**
   * Build dependency graph for all queries
   */
  async buildDependencyGraph(): Promise<QueryDependencyGraph> {
    const queries = await this.getAllQueries();
    const dependencies = new Map<string, string[]>();
    const rootQueries: string[] = [];

    for (const query of queries) {
      // Find references to other queries in the M code
      const referencedQueries: string[] = [];
      const queryReferencePattern = /#"([^"]+)"/g;
      let match;

      while ((match = queryReferencePattern.exec(query.formula)) !== null) {
        const refName = match[1];
        if (queries.some(q => q.name === refName) && refName !== query.name) {
          referencedQueries.push(refName);
        }
      }

      dependencies.set(query.name, referencedQueries);

      // Check if this is a root query (not referenced by others)
      const isReferenced = queries.some(q => {
        if (q.name === query.name) return false;
        return q.formula.includes(`#"${query.name}"`);
      });

      if (!isReferenced) {
        rootQueries.push(query.name);
      }
    }

    return {
      queries: queries.map(q => q.name),
      dependencies,
      rootQueries
    };
  }

  /**
   * Get queries that depend on a specific query
   */
  async getDependentQueries(queryName: string): Promise<string[]> {
    const graph = await this.buildDependencyGraph();
    const dependents: string[] = [];

    for (const [name, deps] of graph.dependencies) {
      if (deps.includes(queryName)) {
        dependents.push(name);
      }
    }

    return dependents;
  }

  // ============================================================================
  // Data Source Analysis
  // ============================================================================

  /**
   * Analyze data sources used by queries
   */
  async analyzeDataSources(): Promise<Map<string, DataSourceInfo>> {
    const queries = await this.getAllQueries();
    const sources = new Map<string, DataSourceInfo>();

    for (const query of queries) {
      const lower = query.formula.toLowerCase();
      let sourceInfo: DataSourceInfo;

      if (lower.includes("file.contents")) {
        const pathMatch = query.formula.match(/File\.Contents\("([^"]+)"\)/);
        const isCsv = lower.includes("csv.document");
        const isExcel = lower.includes("excel.workbook");

        sourceInfo = {
          type: isExcel ? "excel" : isCsv ? "csv" : "other",
          location: pathMatch ? pathMatch[1] : "Unknown",
          isRefreshable: true
        };
      } else if (lower.includes("sql.database")) {
        const serverMatch = query.formula.match(/Sql\.Database\("([^"]+)"/);
        const dbMatch = query.formula.match(/Sql\.Database\([^,]+,\s*"([^"]+)"/);

        sourceInfo = {
          type: "sql",
          location: `${serverMatch?.[1] || "Unknown"}.${dbMatch?.[1] || "Unknown"}`,
          authentication: "Windows/Database",
          isRefreshable: true
        };
      } else if (lower.includes("odata.feed")) {
        const urlMatch = query.formula.match(/OData\.Feed\("([^"]+)"/);
        sourceInfo = {
          type: "odata",
          location: urlMatch?.[1] || "Unknown",
          isRefreshable: true
        };
      } else if (lower.includes("web.contents")) {
        const urlMatch = query.formula.match(/Web\.Contents\("([^"]+)"/);
        sourceInfo = {
          type: "web",
          location: urlMatch?.[1] || "Unknown",
          isRefreshable: true
        };
      } else {
        sourceInfo = {
          type: "other",
          location: "Internal/Calculated",
          isRefreshable: false
        };
      }

      sources.set(query.name, sourceInfo);
    }

    return sources;
  }

  // ============================================================================
  // Performance Monitoring
  // ============================================================================

  /**
   * Get performance metrics for all queries
   */
  async getPerformanceMetrics(): Promise<QueryPerformanceMetrics[]> {
    const queries = await this.getAllQueries();
    const metrics: QueryPerformanceMetrics[] = [];

    for (const query of queries) {
      // In a real implementation, this would track actual load times
      // For now, return estimated metrics
      metrics.push({
        queryName: query.name,
        loadTime: 0,
        rowCount: 0,
        columnCount: 0,
        lastRefresh: new Date()
      });
    }

    return metrics;
  }

  // ============================================================================
  // Utility Methods
  // ============================================================================

  /**
   * Clear the query cache
   */
  clearCache(): void {
    this.queryCache.clear();
  }

  /**
   * Check if a query name is available
   */
  async isQueryNameAvailable(name: string): Promise<boolean> {
    const query = await this.getQuery(name);
    return query === null;
  }

  /**
   * Generate a unique query name
   */
  async generateUniqueName(baseName: string): Promise<string> {
    let counter = 1;
    let name = baseName;

    while (!(await this.isQueryNameAvailable(name))) {
      name = `${baseName}_${counter}`;
      counter++;
    }

    return name;
  }

  /**
   * Duplicate an existing query
   */
  async duplicateQuery(sourceName: string, newName?: string): Promise<QueryModificationResult> {
    try {
      const sourceQuery = await this.getQuery(sourceName);
      if (!sourceQuery) {
        return {
          success: false,
          error: `Source query "${sourceName}" not found`
        };
      }

      const targetName = newName || await this.generateUniqueName(`${sourceName}_Copy`);

      return this.createQuery(targetName, sourceQuery.formula, sourceQuery.description);
    } catch (error) {
      return {
        success: false,
        error: error instanceof Error ? error.message : "Unknown error"
      };
    }
  }

  /**
   * Export query to M code file
   */
  exportQueryToFile(queryName: string): { filename: string; content: string } | null {
    const query = this.queryCache.get(queryName);
    if (!query) return null;

    const formattedMCode = mCodeGenerator.formatMCode(
      this.queryCache.get(queryName)?.steps.map(s => s.mCode).join("\n") || ""
    );

    return {
      filename: `${queryName}.pq`,
      content: `// Query: ${query.name}\n// Source: ${query.source}\n// Exported: ${new Date().toISOString()}\n\n${formattedMCode}`
    };
  }
}

// Export singleton instance
export const powerQueryService = PowerQueryService.getInstance();
export default powerQueryService;
