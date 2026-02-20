/**
 * Excel AI Assistant - M Code Generator Service
 * Power Query M language code generation, validation, and explanation
 * 
 * @module services/mCodeGenerator
 */

import { logger } from '../utils/logger';
import aiService from './aiService';

// ============================================================================
// Type Definitions
// ============================================================================

/** Represents a Power Query operation with M code */
export interface PowerQueryOperation {
  name: string;
  mCode: string;
  description: string;
  category: OperationCategory;
  inputs: OperationInput[];
  preview?: string;
}

/** Categories of Power Query operations */
export type OperationCategory =
  | "source"
  | "transform"
  | "merge"
  | "aggregate"
  | "pivot"
  | "filter"
  | "sort"
  | "addColumn"
  | "group";

/** Input parameter for an operation */
export interface OperationInput {
  name: string;
  type: "string" | "number" | "boolean" | "column" | "columns" | "expression" | "table";
  required: boolean;
  defaultValue?: any;
  description: string;
  options?: string[]; // For dropdown selections
}

/** Validation result for M code */
export interface ValidationResult {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

/** Validation error details */
export interface ValidationError {
  line: number;
  column: number;
  message: string;
  severity: "error";
}

/** Validation warning details */
export interface ValidationWarning {
  line: number;
  message: string;
  severity: "warning";
}

/** Explanation of M code in plain English */
export interface MCodeExplanation {
  summary: string;
  steps: StepExplanation[];
  dataTransformations: DataTransformation[];
  estimatedRowCount?: string;
  performanceNotes: string[];
}

/** Explanation of a single step */
export interface StepExplanation {
  stepName: string;
  lineNumber: number;
  description: string;
  inputTable?: string;
  outputTable?: string;
  transformations: string[];
}

/** Data transformation description */
export interface DataTransformation {
  type: string;
  description: string;
  affectedColumns: string[];
}

/** Query metadata */
export interface QueryMetadata {
  name: string;
  source: string;
  steps: QueryStep[];
  columns: QueryColumn[];
  rowCount?: number;
  isLoaded?: boolean;
}

/** Single step in a query */
export interface QueryStep {
  name: string;
  mCode: string;
  lineNumber: number;
  description?: string;
}

/** Column metadata */
export interface QueryColumn {
  name: string;
  type: string;
  isNullable: boolean;
}

// ============================================================================
// M Code Templates
// ============================================================================

const commonOperations: Record<string, PowerQueryOperation> = {
  sourceExcel: {
    name: "Excel Workbook",
    description: "Import data from an Excel workbook",
    category: "source",
    mCode: `Excel.Workbook(File.Contents("${"C:\\\\Path\\\\To\\\\File.xlsx"}"), true, true)`,
    inputs: [
      { name: "filePath", type: "string", required: true, description: "Full path to the Excel file" },
      { name: "useHeaders", type: "boolean", required: false, defaultValue: true, description: "Use first row as headers" },
      { name: "delayTypes", type: "boolean", required: false, defaultValue: true, description: "Delay type inference" }
    ]
  },

  sourceCsv: {
    name: "CSV File",
    description: "Import data from a CSV file",
    category: "source",
    mCode: `Csv.Document(File.Contents("${"C:\\\\Path\\\\To\\\\File.csv"}"),[Delimiter=",", Columns=10, Encoding=1252, QuoteStyle=QuoteStyle.None])`,
    inputs: [
      { name: "filePath", type: "string", required: true, description: "Full path to the CSV file" },
      { name: "delimiter", type: "string", required: false, defaultValue: ",", description: "Field delimiter" },
      { name: "encoding", type: "number", required: false, defaultValue: 1252, description: "File encoding (1252 = ANSI)" },
      { name: "hasHeaders", type: "boolean", required: false, defaultValue: true, description: "First row contains headers" }
    ]
  },

  sourceSql: {
    name: "SQL Server",
    description: "Import data from SQL Server",
    category: "source",
    mCode: `Sql.Database("serverName", "databaseName", [Query="SELECT * FROM TableName"])`,
    inputs: [
      { name: "server", type: "string", required: true, description: "SQL Server name" },
      { name: "database", type: "string", required: true, description: "Database name" },
      { name: "query", type: "expression", required: true, description: "SQL query or table name" }
    ]
  },

  removeDuplicates: {
    name: "Remove Duplicates",
    description: "Remove duplicate rows from the table",
    category: "transform",
    mCode: `Table.Distinct(#"Previous Step", {"Column1", "Column2"})`,
    inputs: [
      { name: "columns", type: "columns", required: false, description: "Columns to check for duplicates (empty = all columns)" }
    ]
  },

  filterRows: {
    name: "Filter Rows",
    description: "Filter rows based on conditions",
    category: "filter",
    mCode: `Table.SelectRows(#"Previous Step", each [ColumnName] > 100)`,
    inputs: [
      { name: "column", type: "column", required: true, description: "Column to filter on" },
      { name: "operator", type: "string", required: true, description: "Comparison operator", options: ["=", "<>", ">", ">=", "<", "<=", "contains", "starts with", "ends with"] },
      { name: "value", type: "expression", required: true, description: "Value to compare against" }
    ]
  },

  sortRows: {
    name: "Sort Rows",
    description: "Sort rows by one or more columns",
    category: "sort",
    mCode: `Table.Sort(#"Previous Step",{{"ColumnName", Order.Ascending}})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Columns to sort by" },
      { name: "order", type: "string", required: false, defaultValue: "Ascending", description: "Sort order", options: ["Ascending", "Descending"] }
    ]
  },

  groupBy: {
    name: "Group By",
    description: "Group rows and aggregate values",
    category: "group",
    mCode: `Table.Group(#"Previous Step", {"GroupColumn"}, {{"Sum", each List.Sum([ValueColumn]), type number}})`,
    inputs: [
      { name: "groupColumns", type: "columns", required: true, description: "Columns to group by" },
      { name: "aggregation", type: "expression", required: true, description: "Aggregation expression", options: ["Sum", "Count", "Average", "Min", "Max", "All Rows"] }
    ]
  },

  pivotColumn: {
    name: "Pivot Column",
    description: "Pivot a column into multiple columns",
    category: "pivot",
    mCode: `Table.Pivot(#"Previous Step", List.Distinct(#"Previous Step"[AttributeColumn]), "AttributeColumn", "ValueColumn", List.Sum)`,
    inputs: [
      { name: "pivotColumn", type: "column", required: true, description: "Column to pivot" },
      { name: "valueColumn", type: "column", required: true, description: "Column containing values" },
      { name: "aggregation", type: "string", required: false, defaultValue: "Sum", description: "Aggregation function", options: ["Sum", "Count", "Average", "Min", "Max", "Don't Aggregate"] }
    ]
  },

  unpivotColumns: {
    name: "Unpivot Columns",
    description: "Unpivot multiple columns into rows",
    category: "pivot",
    mCode: `Table.UnpivotOtherColumns(#"Previous Step", {"IDColumn"}, "Attribute", "Value")`,
    inputs: [
      { name: "keepColumns", type: "columns", required: true, description: "Columns to keep (not unpivot)" },
      { name: "attributeColumn", type: "string", required: false, defaultValue: "Attribute", description: "Name for attribute column" },
      { name: "valueColumn", type: "string", required: false, defaultValue: "Value", description: "Name for value column" }
    ]
  },

  mergeQueries: {
    name: "Merge Queries",
    description: "Join two tables together",
    category: "merge",
    mCode: `Table.NestedJoin(#"First Table", {"KeyColumn"}, #"Second Table", {"KeyColumn"}, "NewColumn", JoinKind.LeftOuter)`,
    inputs: [
      { name: "primaryTable", type: "table", required: true, description: "Primary table step name" },
      { name: "primaryKey", type: "column", required: true, description: "Primary table key column" },
      { name: "secondaryTable", type: "table", required: true, description: "Secondary table step name" },
      { name: "secondaryKey", type: "column", required: true, description: "Secondary table key column" },
      { name: "joinKind", type: "string", required: false, defaultValue: "LeftOuter", description: "Type of join", options: ["LeftOuter", "RightOuter", "FullOuter", "Inner", "LeftAnti", "RightAnti"] }
    ]
  },

  appendQueries: {
    name: "Append Queries",
    description: "Stack multiple tables vertically",
    category: "merge",
    mCode: `Table.Combine({#"First Table", #"Second Table"})`,
    inputs: [
      { name: "tables", type: "columns", required: true, description: "Table step names to append" }
    ]
  },

  addCustomColumn: {
    name: "Add Custom Column",
    description: "Add a new column with a custom formula",
    category: "addColumn",
    mCode: `Table.AddColumn(#"Previous Step", "NewColumn", each [Column1] + [Column2], type number)`,
    inputs: [
      { name: "columnName", type: "string", required: true, description: "Name for the new column" },
      { name: "formula", type: "expression", required: true, description: "M formula for the column value" },
      { name: "dataType", type: "string", required: false, defaultValue: "any", description: "Data type", options: ["any", "text", "number", "date", "datetime", "logical", "currency"] }
    ]
  },

  changeType: {
    name: "Change Column Type",
    description: "Change the data type of columns",
    category: "transform",
    mCode: `Table.TransformColumnTypes(#"Previous Step",{{"ColumnName", Int64.Type}})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Columns to change" },
      { name: "newType", type: "string", required: true, description: "New data type", options: ["Text", "Number", "Whole Number", "Decimal", "Currency", "Date", "DateTime", "Time", "Percentage", "Logical", "Binary"] }
    ]
  },

  renameColumns: {
    name: "Rename Columns",
    description: "Rename one or more columns",
    category: "transform",
    mCode: `Table.RenameColumns(#"Previous Step",{{"OldName", "NewName"}})`,
    inputs: [
      { name: "renames", type: "expression", required: true, description: "Object mapping old names to new names" }
    ]
  },

  removeColumns: {
    name: "Remove Columns",
    description: "Remove columns from the table",
    category: "transform",
    mCode: `Table.RemoveColumns(#"Previous Step",{"Column1", "Column2"})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Columns to remove" }
    ]
  },

  replaceValues: {
    name: "Replace Values",
    description: "Replace specific values in a column",
    category: "transform",
    mCode: `Table.ReplaceValue(#"Previous Step","oldValue","newValue",Replacer.ReplaceText,{"ColumnName"})`,
    inputs: [
      { name: "column", type: "column", required: true, description: "Column to modify" },
      { name: "oldValue", type: "expression", required: true, description: "Value to find" },
      { name: "newValue", type: "expression", required: true, description: "Value to replace with" }
    ]
  },

  fillDown: {
    name: "Fill Down",
    description: "Fill null values with the value above",
    category: "transform",
    mCode: `Table.FillDown(#"Previous Step", {"ColumnName"})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Columns to fill down" }
    ]
  },

  splitColumn: {
    name: "Split Column",
    description: "Split a column by delimiter or positions",
    category: "transform",
    mCode: `Table.SplitColumn(#"Previous Step", "ColumnName", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"Part1", "Part2"})`,
    inputs: [
      { name: "column", type: "column", required: true, description: "Column to split" },
      { name: "by", type: "string", required: true, description: "Split method", options: ["Delimiter", "Positions", "Number of Characters"] },
      { name: "delimiter", type: "string", required: false, description: "Delimiter character (if using delimiter)" }
    ]
  },

  trimText: {
    name: "Trim Text",
    description: "Remove leading and trailing whitespace",
    category: "transform",
    mCode: `Table.TransformColumns(#"Previous Step", {{"TextColumn", Text.Trim, type text}})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Text columns to trim" }
    ]
  },

  cleanText: {
    name: "Clean Text",
    description: "Remove non-printable characters",
    category: "transform",
    mCode: `Table.TransformColumns(#"Previous Step", {{"TextColumn", Text.Clean, type text}})`,
    inputs: [
      { name: "columns", type: "columns", required: true, description: "Text columns to clean" }
    ]
  }
};

// ============================================================================
// M Code Generator Service
// ============================================================================

export class MCodeGenerator {
  /**
   * Get all available operations grouped by category
   */
  getOperationsByCategory(): Record<OperationCategory, PowerQueryOperation[]> {
    const grouped: Record<OperationCategory, PowerQueryOperation[]> = {
      source: [],
      transform: [],
      merge: [],
      aggregate: [],
      pivot: [],
      filter: [],
      sort: [],
      addColumn: [],
      group: []
    };

    Object.values(commonOperations).forEach(op => {
      grouped[op.category].push(op);
    });

    return grouped;
  }

  /**
   * Get a specific operation by name
   */
  getOperation(name: string): PowerQueryOperation | undefined {
    return commonOperations[name];
  }

  /**
   * Generate M code from natural language description using AI
   * Falls back to pattern matching if AI is unavailable
   */
  async generateFromNaturalLanguage(description: string, context?: QueryMetadata): Promise<string> {
    logger.info('Generating M code from natural language', { description: description.substring(0, 100) });

    try {
      // Try AI-powered generation first
      const aiGenerated = await this.generateWithAI(description, context);
      if (aiGenerated) {
        logger.info('Successfully generated M code using AI');
        return aiGenerated;
      }
    } catch (error) {
      logger.warn('AI generation failed, falling back to pattern matching', { error });
    }

    // Fallback to pattern-based generation
    return this.generateWithPatterns(description, context);
  }

  /**
   * Generate M code using AI service
   */
  private async generateWithAI(description: string, context?: QueryMetadata): Promise<string | null> {
    const settings = aiService.getSettings();
    
    if (!settings.apiKey) {
      logger.debug('No API key configured, skipping AI generation');
      return null;
    }

    const systemPrompt = `You are an expert Power Query M code generator. Convert natural language descriptions into valid M code.

Rules:
1. Return ONLY valid M code without markdown formatting or explanations
2. Use proper M syntax with 'let' and 'in' blocks
3. Reference previous steps using #"Step Name" syntax
4. Use descriptive step names in PascalCase
5. Handle errors gracefully with try/otherwise when appropriate

Common M patterns:
- Source: Excel.Workbook(File.Contents("path"), true, true)
- Filter: Table.SelectRows(#"Previous Step", each [Column] > 0)
- Sort: Table.Sort(#"Previous Step", {{"Column", Order.Ascending}})
- Group: Table.Group(#"Previous Step", {"Column"}, {{"Sum", each List.Sum([Value]), type number}})
- Merge: Table.NestedJoin(#"Table1", {"Key"}, #"Table2", {"Key"}, "JoinResult", JoinKind.LeftOuter)
- Pivot: Table.Pivot(#"Previous Step", List.Distinct(#"Previous Step"[Attribute]), "Attribute", "Value", List.Sum)
- Add Column: Table.AddColumn(#"Previous Step", "NewCol", each [Col1] + [Col2], type number)
- Change Type: Table.TransformColumnTypes(#"Previous Step", {{"Col", Int64.Type}})`;

    const contextPrompt = context ? `
Query Context:
- Query Name: ${context.name}
- Source: ${context.source}
- Available Columns: ${context.columns.map(c => c.name).join(', ')}
- Existing Steps: ${context.steps.map(s => s.name).join(' → ')}
` : '';

    try {
      const response = await fetch(`${settings.apiUrl}/chat/completions`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`
        },
        body: JSON.stringify({
          model: settings.model,
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: `Generate M code for: ${description}${contextPrompt}` }
          ],
          temperature: 0.3,
          max_tokens: 1000
        })
      });

      if (!response.ok) {
        throw new Error(`API Error: ${response.status}`);
      }

      const data = await response.json();
      let mCode = data.choices[0]?.message?.content?.trim() || '';

      // Clean up any markdown formatting
      mCode = mCode.replace(/^```m?\n?/i, '').replace(/```$/m, '').trim();

      // Validate the generated code
      const validation = this.validateMCode(mCode);
      if (!validation.isValid) {
        logger.warn('AI generated invalid M code', { errors: validation.errors });
        return null;
      }

      return mCode;
    } catch (error) {
      logger.error('AI generation error', { error });
      return null;
    }
  }

  /**
   * Generate M code using pattern matching (fallback)
   */
  private generateWithPatterns(description: string, context?: QueryMetadata): string {
    logger.debug('Using pattern-based M code generation');
    
    // Parse common patterns in natural language
    const lowerDesc = description.toLowerCase();

    // Source patterns
    if (lowerDesc.includes("import") || lowerDesc.includes("load") || lowerDesc.includes("connect")) {
      if (lowerDesc.includes("excel") || lowerDesc.includes("xlsx")) {
        return this.generateSourceExcel(description);
      }
      if (lowerDesc.includes("csv") || lowerDesc.includes("text file")) {
        return this.generateSourceCsv(description);
      }
      if (lowerDesc.includes("sql") || lowerDesc.includes("database")) {
        return this.generateSourceSql(description);
      }
    }

    // Transformation patterns
    if (lowerDesc.includes("remove duplicates") || lowerDesc.includes("unique rows")) {
      return commonOperations.removeDuplicates.mCode;
    }

    if (lowerDesc.includes("filter") || lowerDesc.includes("where")) {
      return this.generateFilter(description);
    }

    if (lowerDesc.includes("sort") || lowerDesc.includes("order")) {
      return this.generateSort(description);
    }

    if (lowerDesc.includes("group") || lowerDesc.includes("aggregate")) {
      return this.generateGroupBy(description);
    }

    if (lowerDesc.includes("pivot") || lowerDesc.includes("cross-tab")) {
      return commonOperations.pivotColumn.mCode;
    }

    if (lowerDesc.includes("unpivot")) {
      return commonOperations.unpivotColumns.mCode;
    }

    if (lowerDesc.includes("merge") || lowerDesc.includes("join")) {
      return this.generateMerge(description);
    }

    if (lowerDesc.includes("append") || lowerDesc.includes("combine tables")) {
      return commonOperations.appendQueries.mCode;
    }

    if (lowerDesc.includes("add column") || lowerDesc.includes("new column")) {
      return this.generateAddColumn(description);
    }

    // Return error message instead of TODO comment
    throw new Error(
      `Unable to generate M code for: "${description}". ` +
      `Supported operations: source, filter, sort, group, pivot, merge, append, add column. ` +
      `Please provide a more specific description or use the Power Query Builder interface.`
    );
  }

  /**
   * Generate Excel source M code
   */
  private generateSourceExcel(description: string): string {
    // Extract potential file path from description
    const pathMatch = description.match(/"([^"]+)"/);
    const filePath = pathMatch ? pathMatch[1] : "C:\\\\Path\\\\To\\\\File.xlsx";

    return `let
    Source = Excel.Workbook(File.Contents("${filePath}"), true, true),
    Sheet1 = Source{[Item="Sheet1",Kind="Sheet"]}[Data]
in
    Sheet1`;
  }

  /**
   * Generate CSV source M code
   */
  private generateSourceCsv(description: string): string {
    const pathMatch = description.match(/"([^"]+)"/);
    const filePath = pathMatch ? pathMatch[1] : "C:\\\\Path\\\\To\\\\File.csv";

    return `let
    Source = Csv.Document(File.Contents("${filePath}"),[Delimiter=",", Columns=10, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders`;
  }

  /**
   * Generate SQL source M code
   */
  private generateSourceSql(description: string): string {
    // Extract potential table name
    const tableMatch = description.match(/from\s+(\w+)/i);
    const tableName = tableMatch ? tableMatch[1] : "TableName";

    return `let
    Source = Sql.Database("ServerName", "DatabaseName"),
    ${tableName} = Source{[Schema="dbo",Item="${tableName}"]}[Data]
in
    ${tableName}`;
  }

  /**
   * Generate filter M code
   */
  private generateFilter(description: string): string {
    // Try to extract column and condition
    const columnMatch = description.match(/column\s+(\w+)/i) || description.match(/(\w+)\s+(?:is|>|>=|<|<=|=)/i);
    const column = columnMatch ? columnMatch[1] : "ColumnName";

    return `Table.SelectRows(#"Previous Step", each [${column}] <> null)`;
  }

  /**
   * Generate sort M code
   */
  private generateSort(description: string): string {
    const columnMatch = description.match(/by\s+(\w+)/i) || description.match(/sort\s+(\w+)/i);
    const column = columnMatch ? columnMatch[1] : "ColumnName";
    const order = description.toLowerCase().includes("desc") ? "Order.Descending" : "Order.Ascending";

    return `Table.Sort(#"Previous Step", {{"${column}", ${order}}})`;
  }

  /**
   * Generate group by M code
   */
  private generateGroupBy(description: string): string {
    const columnMatch = description.match(/by\s+(\w+)/i);
    const groupColumn = columnMatch ? columnMatch[1] : "GroupColumn";

    return `Table.Group(#"Previous Step", {"${groupColumn}"}, {{"Count", each Table.RowCount(_), Int64.Type}})`;
  }

  /**
   * Generate merge/join M code
   */
  private generateMerge(description: string): string {
    return `Table.NestedJoin(#"Primary Table", {"KeyColumn"}, #"Secondary Table", {"KeyColumn"}, "JoinResult", JoinKind.LeftOuter)`;
  }

  /**
   * Generate add column M code
   */
  private generateAddColumn(description: string): string {
    const nameMatch = description.match(/called\s+"([^"]+)"/i) || description.match(/named\s+(\w+)/i);
    const columnName = nameMatch ? nameMatch[1] : "NewColumn";

    return `Table.AddColumn(#"Previous Step", "${columnName}", each null, type any)`;
  }

  /**
   * Validate M code syntax
   */
  validateMCode(mCode: string): ValidationResult {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    const lines = mCode.split("\n");

    // Check for balanced parentheses
    let parenCount = 0;
    let bracketCount = 0;
    let braceCount = 0;

    lines.forEach((line, index) => {
      for (const char of line) {
        if (char === "(") parenCount++;
        if (char === ")") parenCount--;
        if (char === "[") bracketCount++;
        if (char === "]") bracketCount--;
        if (char === "{") braceCount++;
        if (char === "}") braceCount--;
      }

      // Check for common issues
      if (line.includes("\"")) {
        const quoteCount = (line.match(/"/g) || []).length;
        if (quoteCount % 2 !== 0) {
          errors.push({
            line: index + 1,
            column: 0,
            message: "Unclosed string literal",
            severity: "error"
          });
        }
      }
    });

    // Final balance check
    if (parenCount !== 0) {
      errors.push({
        line: lines.length,
        column: 0,
        message: `Unbalanced parentheses: ${parenCount > 0 ? "missing" : "extra"} ${Math.abs(parenCount)}`,
        severity: "error"
      });
    }

    if (bracketCount !== 0) {
      errors.push({
        line: lines.length,
        column: 0,
        message: `Unbalanced brackets: ${bracketCount > 0 ? "missing" : "extra"} ${Math.abs(bracketCount)}`,
        severity: "error"
      });
    }

    if (braceCount !== 0) {
      errors.push({
        line: lines.length,
        column: 0,
        message: `Unbalanced braces: ${braceCount > 0 ? "missing" : "extra"} ${Math.abs(braceCount)}`,
        severity: "error"
      });
    }

    // Check for required keywords in let expressions
    if (mCode.toLowerCase().includes("let")) {
      if (!mCode.toLowerCase().includes("in")) {
        errors.push({
          line: lines.length,
          column: 0,
          message: "Missing 'in' keyword for let expression",
          severity: "error"
        });
      }
    }

    // Warnings
    if (mCode.toLowerCase().includes("table.column")) {
      warnings.push({
        line: 0,
        message: "Consider using field access syntax [ColumnName] instead of Table.Column()",
        severity: "warning"
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Explain M code in plain English
   */
  explainMCode(mCode: string): MCodeExplanation {
    const steps: StepExplanation[] = [];
    const dataTransformations: DataTransformation[] = [];
    const performanceNotes: string[] = [];

    // Parse the M code into steps
    const lines = mCode.split("\n");
    let inLetBlock = false;
    let currentStepName = "";
    let stepLineNumber = 0;

    lines.forEach((line, index) => {
      const trimmedLine = line.trim();

      // Detect let block
      if (trimmedLine.toLowerCase() === "let") {
        inLetBlock = true;
        return;
      }

      // Detect in clause
      if (trimmedLine.toLowerCase().startsWith("in")) {
        inLetBlock = false;
        return;
      }

      // Parse steps within let block
      if (inLetBlock && trimmedLine.includes("=")) {
        const stepMatch = trimmedLine.match(/^(\w+)\s*=/);
        if (stepMatch) {
          currentStepName = stepMatch[1];
          stepLineNumber = index + 1;

          const stepDescription = this.explainStep(trimmedLine);
          steps.push({
            stepName: currentStepName,
            lineNumber: stepLineNumber,
            description: stepDescription.summary,
            inputTable: stepDescription.input,
            outputTable: currentStepName,
            transformations: stepDescription.transformations
          });

          dataTransformations.push(...stepDescription.dataTransformations);
        }
      }
    });

    // Generate performance notes
    if (mCode.toLowerCase().includes("nestedjoin")) {
      performanceNotes.push("Merge operations can be memory-intensive with large tables");
    }

    if (mCode.toLowerCase().includes("table.distinct")) {
      performanceNotes.push("Remove duplicates sorts the data, which can be slow on large datasets");
    }

    if ((mCode.match(/Table\.SelectRows/g) || []).length > 2) {
      performanceNotes.push("Multiple filter steps can be combined for better performance");
    }

    return {
      summary: this.generateSummary(steps),
      steps,
      dataTransformations,
      performanceNotes,
      estimatedRowCount: "Unknown - depends on source data"
    };
  }

  /**
   * Explain a single M code step
   */
  private explainStep(line: string): {
    summary: string;
    input: string;
    transformations: string[];
    dataTransformations: DataTransformation[];
  } {
    const transformations: string[] = [];
    const dataTransformations: DataTransformation[] = [];
    let input = "";
    let summary = "";

    // Detect common operations
    if (line.includes("Table.SelectRows")) {
      summary = "Filters rows based on a condition";
      transformations.push("Row filtering applied");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Filter",
        description: "Rows removed based on condition",
        affectedColumns: []
      });
    } else if (line.includes("Table.SelectColumns")) {
      summary = "Selects specific columns from the table";
      transformations.push("Column selection applied");
      input = this.extractInputTable(line);
    } else if (line.includes("Table.RemoveColumns")) {
      summary = "Removes specified columns from the table";
      transformations.push("Columns removed");
      input = this.extractInputTable(line);
    } else if (line.includes("Table.RenameColumns")) {
      summary = "Renames one or more columns";
      transformations.push("Column renaming applied");
      input = this.extractInputTable(line);
    } else if (line.includes("Table.TransformColumnTypes")) {
      summary = "Changes the data type of columns";
      transformations.push("Data type conversion applied");
      input = this.extractInputTable(line);
    } else if (line.includes("Table.Sort")) {
      summary = "Sorts the table by specified columns";
      transformations.push("Row ordering changed");
      input = this.extractInputTable(line);
    } else if (line.includes("Table.Group")) {
      summary = "Groups rows and aggregates values";
      transformations.push("Data grouped and aggregated");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Group",
        description: "Rows grouped by key columns",
        affectedColumns: []
      });
    } else if (line.includes("Table.Distinct")) {
      summary = "Removes duplicate rows";
      transformations.push("Duplicate rows removed");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Filter",
        description: "Duplicate rows eliminated",
        affectedColumns: []
      });
    } else if (line.includes("Table.AddColumn")) {
      summary = "Adds a new calculated column";
      transformations.push("New column added");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "AddColumn",
        description: "Custom column added with formula",
        affectedColumns: []
      });
    } else if (line.includes("Table.NestedJoin")) {
      summary = "Merges data from another table";
      transformations.push("Table join performed");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Join",
        description: "Data merged from secondary table",
        affectedColumns: []
      });
    } else if (line.includes("Table.Combine")) {
      summary = "Appends multiple tables together";
      transformations.push("Tables stacked vertically");
      input = "Multiple tables";
    } else if (line.includes("Table.Pivot")) {
      summary = "Pivots column values into new columns";
      transformations.push("Data pivoted");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Pivot",
        description: "Column values transformed into separate columns",
        affectedColumns: []
      });
    } else if (line.includes("Table.Unpivot")) {
      summary = "Unpivots columns into rows";
      transformations.push("Data unpivoted");
      input = this.extractInputTable(line);
      dataTransformations.push({
        type: "Unpivot",
        description: "Columns converted to rows",
        affectedColumns: []
      });
    } else if (line.includes("Excel.Workbook")) {
      summary = "Connects to an Excel workbook data source";
      transformations.push("Excel data imported");
      input = "External file";
    } else if (line.includes("Csv.Document")) {
      summary = "Connects to a CSV file data source";
      transformations.push("CSV data imported");
      input = "External file";
    } else if (line.includes("Sql.Database")) {
      summary = "Connects to a SQL Server database";
      transformations.push("SQL data imported");
      input = "Database";
    } else {
      summary = "Custom transformation step";
      transformations.push("Transformation applied");
      input = this.extractInputTable(line) || "Previous step";
    }

    return { summary, input, transformations, dataTransformations };
  }

  /**
   * Extract input table name from M code line
   */
  private extractInputTable(line: string): string {
    // Match #"Step Name" or StepName patterns
    const match = line.match(/#"([^"]+)"/);
    if (match) return match[1];

    // Try simple word pattern
    const wordMatch = line.match(/\b([A-Z][a-zA-Z0-9]*)\s*$/);
    if (wordMatch) return wordMatch[1];

    return "Previous step";
  }

  /**
   * Generate overall summary from steps
   */
  private generateSummary(steps: StepExplanation[]): string {
    if (steps.length === 0) {
      return "Empty query with no transformation steps.";
    }

    const stepCount = steps.length;
    const sourceSteps = steps.filter(s =>
      s.description.includes("Connects") || s.description.includes("imported")
    ).length;
    const transformSteps = stepCount - sourceSteps;

    if (sourceSteps === 0) {
      return `Query with ${stepCount} transformation step${stepCount !== 1 ? "s" : ""} applied to existing data.`;
    }

    return `Query imports data and applies ${transformSteps} transformation${transformSteps !== 1 ? "s" : ""} in ${stepCount} total step${stepCount !== 1 ? "s" : ""}.`;
  }

  /**
   * Suggest optimizations for M code
   */
  suggestOptimizations(mCode: string): string[] {
    const suggestions: string[] = [];
    const lowerCode = mCode.toLowerCase();

    // Check for multiple filter steps that could be combined
    const selectRowsCount = (mCode.match(/Table\.SelectRows/g) || []).length;
    if (selectRowsCount > 1) {
      suggestions.push(`Combine ${selectRowsCount} filter steps into a single Table.SelectRows with AND conditions`);
    }

    // Check for early column removal
    if (lowerCode.includes("table.removecolumns") && lowerCode.indexOf("table.removecolumns") < lowerCode.indexOf("table.selectrows")) {
      suggestions.push("Consider filtering rows before removing columns to reduce memory usage");
    }

    // Check for sort before distinct
    if (lowerCode.indexOf("table.sort") < lowerCode.indexOf("table.distinct")) {
      suggestions.push("Remove duplicates before sorting for better performance");
    }

    // Check for nested joins without column selection
    if (lowerCode.includes("nestedjoin") && !lowerCode.includes("table.expandtablecolumn")) {
      suggestions.push("Expand the joined table columns to access merged data");
    }

    // Check for transformation on all columns
    if (lowerCode.includes("table.transformcolumns") && lowerCode.includes("table.columnnames")) {
      suggestions.push("Consider specifying exact columns instead of using Table.ColumnNames for better performance");
    }

    // Check for data type changes
    if ((mCode.match(/Int64\.Type|Number\.Type|Text\.Type/g) || []).length > 5) {
      suggestions.push("Set data types early in the query to improve performance of subsequent operations");
    }

    // Check for list operations
    if (lowerCode.includes("list.distinct") && lowerCode.includes("list.select")) {
      suggestions.push("Use List.Distinct before List.Select for better performance on large lists");
    }

    return suggestions;
  }

  /**
   * Format M code with proper indentation
   */
  formatMCode(mCode: string): string {
    const lines = mCode.split("\n");
    const formatted: string[] = [];
    let indentLevel = 0;

    lines.forEach(line => {
      const trimmedLine = line.trim();

      // Decrease indent for closing braces/parens at start of line
      if (trimmedLine.startsWith(")") || trimmedLine.startsWith("]") || trimmedLine.startsWith("}")) {
        indentLevel = Math.max(0, indentLevel - 1);
      }

      // Add line with proper indentation
      formatted.push("    ".repeat(indentLevel) + trimmedLine);

      // Increase indent for opening braces/parens at end of line
      if (trimmedLine.endsWith("(") || trimmedLine.endsWith("[") || trimmedLine.endsWith("{")) {
        indentLevel++;
      }
    });

    return formatted.join("\n");
  }

  /**
   * Create a complete M query from a series of steps
   */
  buildQuery(steps: { name: string; expression: string }[]): string {
    if (steps.length === 0) {
      return "let\n    Source = #table({}, {})\nin\n    Source";
    }

    const stepLines = steps.map((step, index) => {
      const isLast = index === steps.length - 1;
      const comma = isLast ? "" : ",";
      return `    ${step.name} = ${step.expression}${comma}`;
    });

    const lastStepName = steps[steps.length - 1].name;

    return `let\n${stepLines.join("\n")}\nin\n    ${lastStepName}`;
  }

  /**
   * Parse an existing M query into steps
   */
  parseQuery(mCode: string): { name: string; expression: string }[] {
    const steps: { name: string; expression: string }[] = [];
    const lines = mCode.split("\n");

    let inLetBlock = false;

    for (const line of lines) {
      const trimmedLine = line.trim();

      if (trimmedLine.toLowerCase() === "let") {
        inLetBlock = true;
        continue;
      }

      if (trimmedLine.toLowerCase().startsWith("in")) {
        inLetBlock = false;
        continue;
      }

      if (inLetBlock) {
        // Remove trailing comma
        const cleanLine = trimmedLine.replace(/,$/, "");

        // Match step name and expression
        const match = cleanLine.match(/^(\w+)\s*=\s*(.+)$/);
        if (match) {
          steps.push({
            name: match[1],
            expression: match[2]
          });
        }
      }
    }

    return steps;
  }
}

// Export singleton instance
export const mCodeGenerator = new MCodeGenerator();
export default mCodeGenerator;
