// M Code Explainer - Natural Language Explanations for Power Query M Code
// Phase 3 Implementation - Natural Language Interface for Power Query

import { QueryMetadata, QueryStep } from './mCodeGenerator';

export interface MCodeExplanationDetail {
  stepNumber: number;
  stepName: string;
  plainEnglish: string;
  technicalDetails: string;
  affectedColumns: string[];
  estimatedImpact: 'low' | 'medium' | 'high';
}

export interface QueryFlowExplanation {
  summary: string;
  totalSteps: number;
  dataSource: string;
  transformations: MCodeExplanationDetail[];
  dataLineage: {
    sourceColumns: string[];
    intermediateColumns: string[];
    finalColumns: string[];
  };
  recommendations: string[];
}

export interface DataQualityInsight {
  columnName: string;
  issue: string;
  severity: 'info' | 'warning' | 'error';
  suggestion: string;
}

export class MCodeExplainer {
  private static instance: MCodeExplainer;

  private constructor() {}

  static getInstance(): MCodeExplainer {
    if (!MCodeExplainer.instance) {
      MCodeExplainer.instance = new MCodeExplainer();
    }
    return MCodeExplainer.instance;
  }

  /**
   * Generate a comprehensive natural language explanation of a Power Query
   */
  explainQuery(metadata: QueryMetadata): QueryFlowExplanation {
    const transformations = this.explainSteps(metadata.steps);
    const dataLineage = this.analyzeDataLineage(metadata);

    return {
      summary: this.generateSummary(metadata),
      totalSteps: metadata.steps.length,
      dataSource: metadata.source,
      transformations,
      dataLineage,
      recommendations: this.generateRecommendations(metadata, transformations)
    };
  }

  /**
   * Explain individual steps in plain English
   */
  private explainSteps(steps: QueryStep[]): MCodeExplanationDetail[] {
    return steps.map((step, index) => ({
      stepNumber: index + 1,
      stepName: step.name,
      plainEnglish: this.explainStepInPlainEnglish(step),
      technicalDetails: this.explainStepTechnically(step),
      affectedColumns: this.extractAffectedColumns(step),
      estimatedImpact: this.estimateStepImpact(step)
    }));
  }

  /**
   * Convert a single M step to plain English
   */
  private explainStepInPlainEnglish(step: QueryStep): string {
    const mCode = step.mCode.toLowerCase();

    // Source operations
    if (mCode.includes('excel.workbook')) {
      const path = this.extractStringLiteral(step.mCode) || 'specified file';
      return `Imports data from Excel workbook located at "${path}"`;
    }

    if (mCode.includes('csv.document')) {
      const path = this.extractStringLiteral(step.mCode) || 'specified file';
      return `Reads data from CSV file at "${path}"`;
    }

    if (mCode.includes('sql.database')) {
      return 'Connects to SQL Server database and retrieves data';
    }

    if (mCode.includes('web.contents')) {
      const url = this.extractStringLiteral(step.mCode) || 'specified URL';
      return `Fetches data from web page/API at ${url}`;
    }

    if (mCode.includes('folder.files')) {
      return 'Combines all files from a specified folder';
    }

    if (mCode.includes('sharepoint.contents')) {
      return 'Retrieves data from SharePoint list or library';
    }

    // Filter operations
    if (mCode.includes('table.selectrows')) {
      const condition = this.extractFilterCondition(step.mCode);
      return `Keeps only rows where ${condition}`;
    }

    // Column operations
    if (mCode.includes('table.selectcolumns')) {
      const columns = this.extractColumnList(step.mCode);
      return `Selects only the following columns: ${columns.join(', ')}`;
    }

    if (mCode.includes('table.removecolumns')) {
      const columns = this.extractColumnList(step.mCode);
      return `Removes the columns: ${columns.join(', ')}`;
    }

    if (mCode.includes('table.renamecolumns')) {
      const renames = this.extractRenames(step.mCode);
      return `Renames columns: ${renames.map(r => `"${r.old}" to "${r.new}"`).join(', ')}`;
    }

    if (mCode.includes('table.reordercolumns')) {
      return 'Changes the order of columns in the table';
    }

    // Transformation operations
    if (mCode.includes('table.transformcolumntypes')) {
      return 'Converts column data types (e.g., text to numbers, strings to dates)';
    }

    if (mCode.includes('table.addcolumn')) {
      const colName = this.extractNewColumnName(step.mCode);
      return `Adds a new calculated column named "${colName}"`;
    }

    if (mCode.includes('table.duplicatecolumn')) {
      return 'Creates a copy of an existing column';
    }

    if (mCode.includes('table.splitcolumn')) {
      return 'Splits a column into multiple columns based on delimiter or position';
    }

    if (mCode.includes('table.mergecolumns')) {
      return 'Combines multiple columns into a single column';
    }

    // Sort operations
    if (mCode.includes('table.sort')) {
      const sortCols = this.extractSortColumns(step.mCode);
      return `Sorts the data by ${sortCols.join(', ')}`;
    }

    // Group/Aggregate operations
    if (mCode.includes('table.group')) {
      const groupCols = this.extractGroupColumns(step.mCode);
      return `Groups rows by ${groupCols.join(', ')} and calculates aggregated values`;
    }

    // Merge/Join operations
    if (mCode.includes('table.nestedjoin') || mCode.includes('table.join')) {
      return 'Combines data with another table based on matching keys (like SQL JOIN)';
    }

    if (mCode.includes('table.combine') || mCode.includes('table.append')) {
      return 'Stacks multiple tables vertically (like SQL UNION)';
    }

    // Pivot/Unpivot operations
    if (mCode.includes('table.pivot')) {
      return 'Transforms column values into separate columns (cross-tabulation)';
    }

    if (mCode.includes('table.unpivot')) {
      return 'Converts multiple columns into rows (normalization)';
    }

    // Deduplication
    if (mCode.includes('table.distinct')) {
      return 'Removes duplicate rows keeping only unique records';
    }

    // Text transformations
    if (mCode.includes('text.trim')) {
      return 'Removes leading and trailing whitespace from text';
    }

    if (mCode.includes('text.clean')) {
      return 'Removes non-printable characters from text';
    }

    if (mCode.includes('text.upper')) {
      return 'Converts text to UPPERCASE';
    }

    if (mCode.includes('text.lower')) {
      return 'Converts text to lowercase';
    }

    if (mCode.includes('text.proper')) {
      return 'Converts text to Proper Case (First Letter Capitalized)';
    }

    // Fill operations
    if (mCode.includes('table.filldown')) {
      return 'Fills empty cells with the value from the cell above';
    }

    if (mCode.includes('table.fillup')) {
      return 'Fills empty cells with the value from the cell below';
    }

    // Replace operations
    if (mCode.includes('table.replacevalue')) {
      return 'Replaces specific values with new values';
    }

    if (mCode.includes('table.replaceerrors')) {
      return 'Replaces error values with specified default values';
    }

    // Conditional operations
    if (mCode.includes('table.addconditionalcolumn')) {
      return 'Adds a column based on conditional rules (IF/THEN logic)';
    }

    if (mCode.includes('table.selectrowswitherrors')) {
      return 'Filters to show only rows containing errors';
    }

    if (mCode.includes('table.removeerrors')) {
      return 'Removes rows that contain errors';
    }

    // Index operations
    if (mCode.includes('table.addindexcolumn')) {
      return 'Adds a row number/index column';
    }

    // Custom operations
    if (mCode.includes('table.transformcolumns')) {
      return 'Applies custom transformations to specific columns';
    }

    if (mCode.includes('table.buffer')) {
      return 'Loads the entire table into memory for better performance';
    }

    // Default explanation
    return `Performs ${step.name} operation on the data`;
  }

  /**
   * Provide technical details about a step
   */
  private explainStepTechnically(step: QueryStep): string {
    const mCode = step.mCode;

    // Extract the function call
    const funcMatch = mCode.match(/^(\w+\.)*(\w+)\s*\(/);
    if (funcMatch) {
      const funcName = funcMatch[0].replace('(', '').trim();

      // Map common functions to technical explanations
      const explanations: Record<string, string> = {
        'Table.SelectRows': 'Filters dataset using row-wise predicate evaluation',
        'Table.SelectColumns': 'Projects specific columns, reducing dataset width',
        'Table.RemoveColumns': 'Removes specified columns from projection',
        'Table.TransformColumnTypes': 'Applies type casting to columns',
        'Table.Sort': 'Applies ordering algorithm to rows',
        'Table.Group': 'Performs aggregation using grouping keys',
        'Table.AddColumn': 'Creates derived column using expression evaluation',
        'Table.NestedJoin': 'Performs relational join with nested table expansion',
        'Table.Pivot': 'Performs matrix transpose on column values',
        'Table.Unpivot': 'Normalizes wide format to long format',
        'Table.Distinct': 'Applies set-based deduplication',
        'Table.ReplaceValue': 'Performs value substitution using replacer function',
        'Text.Trim': 'Applies Unicode whitespace trimming',
        'Text.Clean': 'Removes control characters (0x00-0x1F)',
        'Excel.Workbook': 'Uses COM interop to read Excel binary format',
        'Csv.Document': 'Parses RFC 4180 compliant CSV format',
        'Sql.Database': 'Executes T-SQL via ADO.NET connection'
      };

      return explanations[funcName] || `Executes ${funcName} function`;
    }

    return 'Custom transformation step';
  }

  /**
   * Extract columns affected by a step
   */
  private extractAffectedColumns(step: QueryStep): string[] {
    const mCode = step.mCode;
    const columns: string[] = [];

    // Extract quoted column names
    const columnMatches = mCode.match(/\[([^\]]+)\]/g);
    if (columnMatches) {
      columnMatches.forEach(match => {
        const colName = match.replace(/\[|\]/g, '');
        if (!columns.includes(colName)) {
          columns.push(colName);
        }
      });
    }

    // Extract string literals that might be column names in certain contexts
    if (mCode.includes('Table.SelectColumns') || mCode.includes('Table.RemoveColumns')) {
      const stringMatches = mCode.match(/"([^"]+)"/g);
      if (stringMatches) {
        stringMatches.forEach(match => {
          const colName = match.replace(/"/g, '');
          if (!columns.includes(colName) && !colName.includes('\\') && !colName.includes('http')) {
            columns.push(colName);
          }
        });
      }
    }

    return columns;
  }

  /**
   * Estimate the performance impact of a step
   */
  private estimateStepImpact(step: QueryStep): 'low' | 'medium' | 'high' {
    const mCode = step.mCode.toLowerCase();

    // High impact operations
    if (mCode.includes('nestedjoin') ||
        mCode.includes('table.join') ||
        mCode.includes('table.combine') ||
        mCode.includes('table.group') ||
        mCode.includes('table.sort') ||
        mCode.includes('table.distinct') ||
        mCode.includes('table.pivot') ||
        mCode.includes('table.buffer')) {
      return 'high';
    }

    // Medium impact operations
    if (mCode.includes('table.selectrows') ||
        mCode.includes('table.replacevalue') ||
        mCode.includes('table.filldown') ||
        mCode.includes('table.fillup') ||
        mCode.includes('table.splitcolumn') ||
        mCode.includes('list.')) {
      return 'medium';
    }

    // Low impact (default)
    return 'low';
  }

  /**
   * Analyze data lineage through the query
   */
  private analyzeDataLineage(metadata: QueryMetadata): QueryFlowExplanation['dataLineage'] {
    // This is a simplified analysis
    // A full implementation would track column lineage through each step

    const allColumns = new Set<string>();
    metadata.steps.forEach(step => {
      this.extractAffectedColumns(step).forEach(col => allColumns.add(col));
    });

    return {
      sourceColumns: Array.from(allColumns).slice(0, 5), // First 5 as source
      intermediateColumns: Array.from(allColumns).slice(2, 8), // Middle as intermediate
      finalColumns: metadata.columns.map(c => c.name)
    };
  }

  /**
   * Generate overall query summary
   */
  private generateSummary(metadata: QueryMetadata): string {
    const stepCount = metadata.steps.length;
    const sourceType = metadata.source;

    let summary = `This Power Query imports data from ${sourceType} and `;

    if (stepCount === 0) {
      summary += 'passes it through without any transformations.';
    } else if (stepCount === 1) {
      summary += 'applies a single transformation.';
    } else {
      summary += `applies ${stepCount} transformation steps to clean, shape, and prepare the data.`;
    }

    // Add purpose hint based on transformations
    const hasGrouping = metadata.steps.some(s => s.mCode.includes('Table.Group'));
    const hasJoin = metadata.steps.some(s => s.mCode.includes('Join'));
    const hasPivot = metadata.steps.some(s => s.mCode.includes('Pivot'));

    if (hasGrouping) {
      summary += ' The query performs aggregation analysis.';
    }
    if (hasJoin) {
      summary += ' It combines data from multiple sources.';
    }
    if (hasPivot) {
      summary += ' It reshapes the data structure for reporting.';
    }

    return summary;
  }

  /**
   * Generate optimization recommendations
   */
  private generateRecommendations(
    metadata: QueryMetadata,
    transformations: MCodeExplanationDetail[]
  ): string[] {
    const recommendations: string[] = [];

    // Check for multiple filters
    const filterSteps = transformations.filter(t =>
      t.technicalDetails.includes('Filters dataset')
    );
    if (filterSteps.length > 1) {
      recommendations.push(`Consider combining ${filterSteps.length} filter steps into a single step with AND conditions for better performance.`);
    }

    // Check for column selection timing
    const selectColumnsStep = metadata.steps.findIndex(s =>
      s.mCode.includes('Table.SelectColumns')
    );
    const filterStep = metadata.steps.findIndex(s =>
      s.mCode.includes('Table.SelectRows')
    );
    if (selectColumnsStep > filterStep && filterStep !== -1) {
      recommendations.push('Move column selection before row filtering to reduce memory usage.');
    }

    // Check for type changes
    const typeChangeStep = metadata.steps.findIndex(s =>
      s.mCode.includes('TransformColumnTypes')
    );
    const calculationsAfterType = metadata.steps.slice(typeChangeStep + 1).some(s =>
      s.mCode.includes('AddColumn') || s.mCode.includes('Group')
    );
    if (typeChangeStep > 2 && calculationsAfterType) {
      recommendations.push('Set data types earlier in the query to enable better query folding.');
    }

    // Check for distinct before sort
    const distinctStep = metadata.steps.findIndex(s =>
      s.mCode.includes('Table.Distinct')
    );
    const sortStep = metadata.steps.findIndex(s =>
      s.mCode.includes('Table.Sort')
    );
    if (sortStep < distinctStep && sortStep !== -1) {
      recommendations.push('Consider moving distinct operation before sort for better performance.');
    }

    // Check for merge without expansion
    const mergeSteps = metadata.steps.filter(s =>
      s.mCode.includes('NestedJoin')
    );
    mergeSteps.forEach((step, idx) => {
      const stepIndex = metadata.steps.indexOf(step);
      const nextStep = metadata.steps[stepIndex + 1];
      if (!nextStep || !nextStep.mCode.includes('ExpandTableColumn')) {
        recommendations.push(`Step "${step.name}" performs a merge but doesn't expand columns. Add an expand step to access joined data.`);
      }
    });

    // High impact steps warning
    const highImpactSteps = transformations.filter(t => t.estimatedImpact === 'high');
    if (highImpactSteps.length > 2) {
      recommendations.push(`Query has ${highImpactSteps.length} high-impact operations. Monitor performance with large datasets.`);
    }

    return recommendations;
  }

  /**
   * Explain what a query does at a high level
   */
  explainQueryPurpose(metadata: QueryMetadata): string {
    const steps = metadata.steps;

    // Detect common patterns
    const hasSource = steps.some(s =>
      s.mCode.includes('Excel') || s.mCode.includes('Csv') || s.mCode.includes('Sql')
    );
    const hasClean = steps.some(s =>
      s.mCode.includes('Trim') || s.mCode.includes('Clean') || s.mCode.includes('Replace')
    );
    const hasFilter = steps.some(s => s.mCode.includes('SelectRows'));
    const hasJoin = steps.some(s => s.mCode.includes('Join'));
    const hasGroup = steps.some(s => s.mCode.includes('Group'));

    let purpose = '';

    if (hasSource) {
      purpose += 'Data Import: Retrieves data from external source. ';
    }
    if (hasClean) {
      purpose += 'Data Cleansing: Standardizes and cleans data values. ';
    }
    if (hasFilter) {
      purpose += 'Data Filtering: Applies business rules to select relevant records. ';
    }
    if (hasJoin) {
      purpose += 'Data Integration: Combines data from multiple sources. ';
    }
    if (hasGroup) {
      purpose += 'Data Aggregation: Summarizes data for analysis. ';
    }

    return purpose || 'Custom data transformation pipeline';
  }

  /**
   * Compare two versions of a query
   */
  compareQueries(oldMetadata: QueryMetadata, newMetadata: QueryMetadata): string[] {
    const changes: string[] = [];

    if (oldMetadata.steps.length !== newMetadata.steps.length) {
      const diff = newMetadata.steps.length - oldMetadata.steps.length;
      changes.push(`Step count changed by ${diff > 0 ? '+' : ''}${diff} (${oldMetadata.steps.length} → ${newMetadata.steps.length})`);
    }

    // Compare step by step
    const maxSteps = Math.max(oldMetadata.steps.length, newMetadata.steps.length);
    for (let i = 0; i < maxSteps; i++) {
      const oldStep = oldMetadata.steps[i];
      const newStep = newMetadata.steps[i];

      if (!oldStep && newStep) {
        changes.push(`Added step ${i + 1}: "${newStep.name}" - ${this.explainStepInPlainEnglish(newStep)}`);
      } else if (oldStep && !newStep) {
        changes.push(`Removed step ${i + 1}: "${oldStep.name}"`);
      } else if (oldStep && newStep && oldStep.mCode !== newStep.mCode) {
        changes.push(`Modified step ${i + 1}: "${newStep.name}"`);
      }
    }

    return changes;
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  private extractStringLiteral(mCode: string): string | null {
    const match = mCode.match(/"([^"]+)"/);
    return match ? match[1] : null;
  }

  private extractFilterCondition(mCode: string): string {
    // Extract condition from Table.SelectRows
    const eachMatch = mCode.match(/each\s+(.+?)(?:,|\))/);
    if (eachMatch) {
      return eachMatch[1].trim();
    }
    return 'specified condition';
  }

  private extractColumnList(mCode: string): string[] {
    const columns: string[] = [];
    const matches = mCode.match(/"([^"]+)"/g);
    if (matches) {
      matches.forEach(m => columns.push(m.replace(/"/g, '')));
    }
    return columns.length > 0 ? columns : ['specified columns'];
  }

  private extractRenames(mCode: string): Array<{ old: string; new: string }> {
    const renames: Array<{ old: string; new: string }> = [];
    // Match {"OldName", "NewName"} patterns
    const matches = mCode.match(/\{\"([^\"]+)\",\s*\"([^\"]+)\"\}/g);
    if (matches) {
      matches.forEach(match => {
        const parts = match.match(/\"([^\"]+)\"/g);
        if (parts && parts.length >= 2) {
          renames.push({
            old: parts[0].replace(/"/g, ''),
            new: parts[1].replace(/"/g, '')
          });
        }
      });
    }
    return renames;
  }

  private extractSortColumns(mCode: string): string[] {
    const columns: string[] = [];
    // Match {"Column", Order.Ascending} patterns
    const matches = mCode.match(/\{\"([^\"]+)\"/g);
    if (matches) {
      matches.forEach(m => {
        const col = m.match(/\"([^\"]+)\"/);
        if (col) columns.push(col[1]);
      });
    }
    return columns.length > 0 ? columns : ['specified columns'];
  }

  private extractGroupColumns(mCode: string): string[] {
    const match = mCode.match(/\{\"([^\"]+)\"\}/);
    if (match) {
      return [match[1]];
    }
    return ['specified columns'];
  }

  private extractNewColumnName(mCode: string): string {
    const match = mCode.match(/"([^"]+)"\s*,\s*each/);
    return match ? match[1] : 'New Column';
  }
}

export default MCodeExplainer.getInstance();
