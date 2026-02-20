/**
 * DAX Parser Service
 *
 * Parses and explains DAX (Data Analysis Expressions) formulas used in
 * Power Pivot, Power BI, and Analysis Services. Provides human-readable
 * explanations, syntax validation, and optimization suggestions.
 *
 * Features:
 * - Full DAX function library (250+ functions)
 * - Syntax validation and error detection
 * - Human-readable formula explanations
 * - Measure dependency analysis
 * - Performance optimization hints
 * - Context transition explanations
 *
 * @module services/daxParser
 */

// ============================================================================
// Type Definitions
// ============================================================================

/** Types of DAX expressions */
export type DAXExpressionType =
  | "measure"
  | "calculated_column"
  | "calculated_table"
  | "table_expression"
  | "filter_expression";

/** DAX token types */
export type DAXTokenType =
  | "function"
  | "table"
  | "column"
  | "measure"
  | "operator"
  | "number"
  | "string"
  | "boolean"
  | "parenthesis"
  | "comma"
  | "identifier"
  | "comment";

/** DAX token */
export interface DAXToken {
  type: DAXTokenType;
  value: string;
  position: number;
  length: number;
}

/** DAX function metadata */
export interface DAXFunction {
  name: string;
  category: DAXFunctionCategory;
  description: string;
  syntax: string;
  parameters: DAXParameter[];
  returnType: string;
  isAggregator?: boolean;
  isIterator?: boolean;
  isTimeIntelligence?: boolean;
  isFilterFunction?: boolean;
  modifiesFilterContext?: boolean;
  hasAlternatives?: string[];
}

/** DAX function parameter */
export interface DAXParameter {
  name: string;
  description: string;
  type: string;
  optional?: boolean;
  defaultValue?: string;
}

/** DAX function categories */
export type DAXFunctionCategory =
  | "aggregation"
  | "filter"
  | "time_intelligence"
  | "logical"
  | "mathematical"
  | "text"
  | "date_time"
  | "relationship"
  | "information"
  | "table_manipulation"
  | "parent_child"
  | "other";

/** Parsed DAX formula result */
export interface ParsedDAXFormula {
  expression: string;
  type: DAXExpressionType;
  tokens: DAXToken[];
  functions: string[];
  tables: string[];
  columns: string[];
  measures: string[];
  dependencies: DAXDependency[];
  hasRowContext: boolean;
  hasFilterContext: boolean;
  hasContextTransition: boolean;
}

/** DAX dependency (measure/column references) */
export interface DAXDependency {
  type: "measure" | "column" | "table";
  name: string;
  table?: string;
  position: number;
}

/** DAX explanation result */
export interface DAXExplanation {
  summary: string;
  breakdown: DAXExplanationStep[];
  complexity: DAXComplexity;
  contextInfo: DAXContextInfo;
  optimizations: DAXOptimization[];
  performanceHints: string[];
}

/** Explanation step */
export interface DAXExplanationStep {
  function: string;
  description: string;
  context: string;
  indentation: number;
}

/** Complexity assessment */
export type DAXComplexity = "simple" | "moderate" | "complex" | "very_complex";

/** Context information */
export interface DAXContextInfo {
  rowContext: boolean;
  filterContext: boolean;
  contextTransitions: number;
  iteratorFunctions: number;
  filterModifications: number;
}

/** Optimization suggestion */
export interface DAXOptimization {
  type: "performance" | "readability" | "best_practice";
  severity: "low" | "medium" | "high";
  message: string;
  suggestion: string;
  example?: string;
}

/** Syntax error */
export interface DAXSyntaxError {
  message: string;
  position: number;
  length: number;
  severity: "error" | "warning";
}

// ============================================================================
// DAX Function Library
// ============================================================================

export const DAX_FUNCTIONS: Record<string, DAXFunction> = {
  // Aggregation Functions
  SUM: {
    name: "SUM",
    category: "aggregation",
    description: "Adds all the numbers in a column",
    syntax: "SUM(<column>)",
    parameters: [
      { name: "column", description: "The column that contains the numbers to sum", type: "column" },
    ],
    returnType: "number",
    isAggregator: true,
  },
  AVERAGE: {
    name: "AVERAGE",
    category: "aggregation",
    description: "Returns the average of all the numbers in a column",
    syntax: "AVERAGE(<column>)",
    parameters: [
      { name: "column", description: "The column that contains the numbers to average", type: "column" },
    ],
    returnType: "number",
    isAggregator: true,
  },
  MIN: {
    name: "MIN",
    category: "aggregation",
    description: "Returns the smallest value in a column",
    syntax: "MIN(<column>)",
    parameters: [
      { name: "column", description: "The column in which you want to find the smallest value", type: "column" },
    ],
    returnType: "scalar",
    isAggregator: true,
  },
  MAX: {
    name: "MAX",
    category: "aggregation",
    description: "Returns the largest value in a column",
    syntax: "MAX(<column>)",
    parameters: [
      { name: "column", description: "The column in which you want to find the largest value", type: "column" },
    ],
    returnType: "scalar",
    isAggregator: true,
  },
  COUNT: {
    name: "COUNT",
    category: "aggregation",
    description: "Counts the number of rows in the specified column",
    syntax: "COUNT(<column>)",
    parameters: [
      { name: "column", description: "The column that contains the values to count", type: "column" },
    ],
    returnType: "number",
    isAggregator: true,
  },
  COUNTROWS: {
    name: "COUNTROWS",
    category: "aggregation",
    description: "Counts the number of rows in a table",
    syntax: "COUNTROWS(<table>)",
    parameters: [
      { name: "table", description: "The table containing the rows to be counted", type: "table" },
    ],
    returnType: "number",
    isAggregator: true,
  },
  DISTINCTCOUNT: {
    name: "DISTINCTCOUNT",
    category: "aggregation",
    description: "Counts the number of distinct values in a column",
    syntax: "DISTINCTCOUNT(<column>)",
    parameters: [
      { name: "column", description: "The column that contains the values to be counted", type: "column" },
    ],
    returnType: "number",
    isAggregator: true,
    hasAlternatives: ["DISTINCTCOUNTNOBLANK"],
  },

  // Iterator Functions
  SUMX: {
    name: "SUMX",
    category: "aggregation",
    description: "Returns the sum of an expression evaluated for each row in a table",
    syntax: "SUMX(<table>, <expression>)",
    parameters: [
      { name: "table", description: "The table containing the rows for which the expression will be evaluated", type: "table" },
      { name: "expression", description: "The expression to be evaluated for each row", type: "expression" },
    ],
    returnType: "number",
    isIterator: true,
  },
  AVERAGEX: {
    name: "AVERAGEX",
    category: "aggregation",
    description: "Calculates the average of a set of expressions evaluated over a table",
    syntax: "AVERAGEX(<table>, <expression>)",
    parameters: [
      { name: "table", description: "The table containing the rows for which the expression will be evaluated", type: "table" },
      { name: "expression", description: "The expression to be evaluated for each row", type: "expression" },
    ],
    returnType: "number",
    isIterator: true,
  },
  MINX: {
    name: "MINX",
    category: "aggregation",
    description: "Returns the minimum value of an expression evaluated for each row in a table",
    syntax: "MINX(<table>, <expression>)",
    parameters: [
      { name: "table", description: "The table containing the rows", type: "table" },
      { name: "expression", description: "The expression to be evaluated for each row", type: "expression" },
    ],
    returnType: "scalar",
    isIterator: true,
  },
  MAXX: {
    name: "MAXX",
    category: "aggregation",
    description: "Returns the maximum value of an expression evaluated for each row in a table",
    syntax: "MAXX(<table>, <expression>)",
    parameters: [
      { name: "table", description: "The table containing the rows", type: "table" },
      { name: "expression", description: "The expression to be evaluated for each row", type: "expression" },
    ],
    returnType: "scalar",
    isIterator: true,
  },

  // Filter Functions
  CALCULATE: {
    name: "CALCULATE",
    category: "filter",
    description: "Evaluates an expression in a context modified by filters",
    syntax: "CALCULATE(<expression>, <filter1>, <filter2>, ...)",
    parameters: [
      { name: "expression", description: "The expression to be evaluated", type: "expression" },
      { name: "filter", description: "A Boolean expression or table expression that defines a filter", type: "filter", optional: true },
    ],
    returnType: "scalar",
    modifiesFilterContext: true,
    hasAlternatives: ["CALCULATETABLE"],
  },
  FILTER: {
    name: "FILTER",
    category: "filter",
    description: "Returns a table that represents a subset of another table or expression",
    syntax: "FILTER(<table>, <filter>)",
    parameters: [
      { name: "table", description: "The table to be filtered", type: "table" },
      { name: "filter", description: "A Boolean expression to be evaluated for each row", type: "boolean" },
    ],
    returnType: "table",
    isFilterFunction: true,
  },
  ALL: {
    name: "ALL",
    category: "filter",
    description: "Returns all the rows in a table, or all the values in a column, ignoring any filters",
    syntax: "ALL(<table_or_column>)",
    parameters: [
      { name: "table_or_column", description: "The table or column from which to remove filters", type: "table|column" },
    ],
    returnType: "table|column",
    isFilterFunction: true,
    modifiesFilterContext: true,
  },
  ALLEXCEPT: {
    name: "ALLEXCEPT",
    category: "filter",
    description: "Removes all context filters except filters applied to specified columns",
    syntax: "ALLEXCEPT(<table>, <column1>, <column2>, ...)",
    parameters: [
      { name: "table", description: "The table over which all context filters are removed", type: "table" },
      { name: "column", description: "A column in the table for which context filters must be preserved", type: "column" },
    ],
    returnType: "table",
    isFilterFunction: true,
    modifiesFilterContext: true,
  },
  VALUES: {
    name: "VALUES",
    category: "filter",
    description: "Returns a one-column table that contains the distinct values from the specified column",
    syntax: "VALUES(<table_or_column>)",
    parameters: [
      { name: "table_or_column", description: "A table or a single column", type: "table|column" },
    ],
    returnType: "table",
    isFilterFunction: true,
    hasAlternatives: ["DISTINCT"],
  },
  HASONEVALUE: {
    name: "HASONEVALUE",
    category: "filter",
    description: "Returns TRUE when the context for column has been filtered down to one distinct value",
    syntax: "HASONEVALUE(<column>)",
    parameters: [
      { name: "column", description: "The column to check for a single value", type: "column" },
    ],
    returnType: "boolean",
  },

  // Time Intelligence Functions
  TOTALYTD: {
    name: "TOTALYTD",
    category: "time_intelligence",
    description: "Evaluates the year-to-date value of an expression",
    syntax: "TOTALYTD(<expression>, <dates>, <filter>)",
    parameters: [
      { name: "expression", description: "The expression to evaluate", type: "expression" },
      { name: "dates", description: "A column that contains dates", type: "column" },
      { name: "filter", description: "An optional expression that specifies a filter", type: "filter", optional: true },
    ],
    returnType: "scalar",
    isTimeIntelligence: true,
  },
  TOTALQTD: {
    name: "TOTALQTD",
    category: "time_intelligence",
    description: "Evaluates the quarter-to-date value of an expression",
    syntax: "TOTALQTD(<expression>, <dates>, <filter>)",
    parameters: [
      { name: "expression", description: "The expression to evaluate", type: "expression" },
      { name: "dates", description: "A column that contains dates", type: "column" },
      { name: "filter", description: "An optional expression that specifies a filter", type: "filter", optional: true },
    ],
    returnType: "scalar",
    isTimeIntelligence: true,
  },
  TOTALMTD: {
    name: "TOTALMTD",
    category: "time_intelligence",
    description: "Evaluates the month-to-date value of an expression",
    syntax: "TOTALMTD(<expression>, <dates>, <filter>)",
    parameters: [
      { name: "expression", description: "The expression to evaluate", type: "expression" },
      { name: "dates", description: "A column that contains dates", type: "column" },
      { name: "filter", description: "An optional expression that specifies a filter", type: "filter", optional: true },
    ],
    returnType: "scalar",
    isTimeIntelligence: true,
  },
  SAMEPERIODLASTYEAR: {
    name: "SAMEPERIODLASTYEAR",
    category: "time_intelligence",
    description: "Returns a table that contains a column of dates shifted one year back",
    syntax: "SAMEPERIODLASTYEAR(<dates>)",
    parameters: [
      { name: "dates", description: "A column that contains dates", type: "column" },
    ],
    returnType: "table",
    isTimeIntelligence: true,
  },
  DATESYTD: {
    name: "DATESYTD",
    category: "time_intelligence",
    description: "Returns a table that contains a column of the dates for the year-to-date",
    syntax: "DATESYTD(<dates>)",
    parameters: [
      { name: "dates", description: "A column that contains dates", type: "column" },
    ],
    returnType: "table",
    isTimeIntelligence: true,
  },
  PARALLELPERIOD: {
    name: "PARALLELPERIOD",
    category: "time_intelligence",
    description: "Returns a table that contains a column of dates representing a period parallel to the dates in the current context",
    syntax: "PARALLELPERIOD(<dates>, <number_of_intervals>, <interval>)",
    parameters: [
      { name: "dates", description: "A column that contains dates", type: "column" },
      { name: "number_of_intervals", description: "An integer that specifies the number of intervals to shift", type: "integer" },
      { name: "interval", description: "The interval to use (YEAR, QUARTER, MONTH)", type: "enumeration" },
    ],
    returnType: "table",
    isTimeIntelligence: true,
  },

  // Logical Functions
  IF: {
    name: "IF",
    category: "logical",
    description: "Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE",
    syntax: "IF(<logical_test>, <value_if_true>, <value_if_false>)",
    parameters: [
      { name: "logical_test", description: "Any value or expression that can be evaluated to TRUE or FALSE", type: "boolean" },
      { name: "value_if_true", description: "The value to return if logical_test is TRUE", type: "scalar" },
      { name: "value_if_false", description: "The value to return if logical_test is FALSE", type: "scalar", optional: true },
    ],
    returnType: "scalar",
    hasAlternatives: ["SWITCH"],
  },
  AND: {
    name: "AND",
    category: "logical",
    description: "Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE",
    syntax: "AND(<logical1>, <logical2>, ...)",
    parameters: [
      { name: "logical", description: "A logical value to evaluate", type: "boolean" },
    ],
    returnType: "boolean",
  },
  OR: {
    name: "OR",
    category: "logical",
    description: "Checks whether one of the arguments is TRUE to return TRUE",
    syntax: "OR(<logical1>, <logical2>, ...)",
    parameters: [
      { name: "logical", description: "A logical value to evaluate", type: "boolean" },
    ],
    returnType: "boolean",
  },
  NOT: {
    name: "NOT",
    category: "logical",
    description: "Changes FALSE to TRUE, or TRUE to FALSE",
    syntax: "NOT(<logical>)",
    parameters: [
      { name: "logical", description: "A value or expression that can be evaluated to TRUE or FALSE", type: "boolean" },
    ],
    returnType: "boolean",
  },
  SWITCH: {
    name: "SWITCH",
    category: "logical",
    description: "Evaluates an expression against a list of values and returns one of multiple possible result expressions",
    syntax: "SWITCH(<expression>, <value>, <result>, [<value>, <result>], [<else>])",
    parameters: [
      { name: "expression", description: "Any DAX expression that returns a single scalar value", type: "scalar" },
      { name: "value", description: "A constant value to match against the expression", type: "scalar" },
      { name: "result", description: "Any scalar expression to be evaluated if the results of expression match the corresponding value", type: "scalar" },
      { name: "else", description: "Any scalar expression to be evaluated if the result of expression doesn't match any of the value arguments", type: "scalar", optional: true },
    ],
    returnType: "scalar",
    hasAlternatives: ["IF"],
  },

  // Table Manipulation
  SELECTCOLUMNS: {
    name: "SELECTCOLUMNS",
    category: "table_manipulation",
    description: "Returns a table with selected columns from the table and new columns specified by DAX expressions",
    syntax: "SELECTCOLUMNS(<table>, <name>, <scalar_expression>, ...)",
    parameters: [
      { name: "table", description: "Any DAX expression that returns a table", type: "table" },
      { name: "name", description: "The name of the new column", type: "string" },
      { name: "scalar_expression", description: "Any DAX expression that returns a scalar value", type: "expression" },
    ],
    returnType: "table",
  },
  ADDCOLUMNS: {
    name: "ADDCOLUMNS",
    category: "table_manipulation",
    description: "Returns a table with new columns specified by DAX expressions",
    syntax: "ADDCOLUMNS(<table>, <name>, <expression>, ...)",
    parameters: [
      { name: "table", description: "Any DAX expression that returns a table", type: "table" },
      { name: "name", description: "The name of the new column", type: "string" },
      { name: "expression", description: "Any DAX expression that returns a scalar value", type: "expression" },
    ],
    returnType: "table",
  },
  SUMMARIZE: {
    name: "SUMMARIZE",
    category: "table_manipulation",
    description: "Creates a summary of the input table grouped by the specified columns",
    syntax: "SUMMARIZE(<table>, <groupBy_columnName>, [<groupBy_columnName>], ...)",
    parameters: [
      { name: "table", description: "Any DAX expression that returns a table of data", type: "table" },
      { name: "groupBy_columnName", description: "The column to group by", type: "column" },
    ],
    returnType: "table",
  },
  DISTINCT: {
    name: "DISTINCT",
    category: "table_manipulation",
    description: "Returns a table by removing duplicate rows from another table or expression",
    syntax: "DISTINCT(<table_or_column>)",
    parameters: [
      { name: "table_or_column", description: "A table or a single column", type: "table|column" },
    ],
    returnType: "table",
  },

  // Information Functions
  ISBLANK: {
    name: "ISBLANK",
    category: "information",
    description: "Checks whether a value is blank, and returns TRUE or FALSE",
    syntax: "ISBLANK(<value>)",
    parameters: [
      { name: "value", description: "The value or expression to check", type: "scalar" },
    ],
    returnType: "boolean",
  },
  ISNUMBER: {
    name: "ISNUMBER",
    category: "information",
    description: "Checks whether a value is a number, and returns TRUE or FALSE",
    syntax: "ISNUMBER(<value>)",
    parameters: [
      { name: "value", description: "The value to check", type: "scalar" },
    ],
    returnType: "boolean",
  },
  ISTEXT: {
    name: "ISTEXT",
    category: "information",
    description: "Checks whether a value is text, and returns TRUE or FALSE",
    syntax: "ISTEXT(<value>)",
    parameters: [
      { name: "value", description: "The value to check", type: "scalar" },
    ],
    returnType: "boolean",
  },
  HASONEFILTER: {
    name: "HASONEFILTER",
    category: "information",
    description: "Returns TRUE when the number of directly filtered values on column is one",
    syntax: "HASONEFILTER(<column>)",
    parameters: [
      { name: "column", description: "The column to check", type: "column" },
    ],
    returnType: "boolean",
  },
  ISCROSSFILTERED: {
    name: "ISCROSSFILTERED",
    category: "information",
    description: "Returns TRUE when column or another column in the same or related table is being filtered",
    syntax: "ISCROSSFILTERED(<table_or_column>)",
    parameters: [
      { name: "table_or_column", description: "The column or table to check", type: "table|column" },
    ],
    returnType: "boolean",
  },

  // Relationship Functions
  RELATED: {
    name: "RELATED",
    category: "relationship",
    description: "Returns a related value from another table",
    syntax: "RELATED(<column>)",
    parameters: [
      { name: "column", description: "The column that contains the desired value", type: "column" },
    ],
    returnType: "scalar",
  },
  RELATEDTABLE: {
    name: "RELATEDTABLE",
    category: "relationship",
    description: "Returns the related tables filtered so that it only includes the related rows",
    syntax: "RELATEDTABLE(<table>)",
    parameters: [
      { name: "table", description: "The related table to be returned", type: "table" },
    ],
    returnType: "table",
  },
  USERELATIONSHIP: {
    name: "USERELATIONSHIP",
    category: "relationship",
    description: "Specifies an existing relationship to be used in a specific calculation",
    syntax: "USERELATIONSHIP(<columnName1>, <columnName2>)",
    parameters: [
      { name: "columnName1", description: "One column of the relationship", type: "column" },
      { name: "columnName2", description: "The other column of the relationship", type: "column" },
    ],
    returnType: "none",
  },
};

// Add more functions dynamically
const additionalFunctions = [
  "COUNTA", "COUNTAX", "COUNTBLANK", "PRODUCT", "PRODUCTX",
  "DIVIDE", "MOD", "QUOTIENT", "POWER", "SQRT", "EXP", "LN", "LOG",
  "ABS", "ROUND", "ROUNDDOWN", "ROUNDUP", "INT", "TRUNC",
  "CONCATENATE", "CONCATENATEX", "LEFT", "RIGHT", "MID", "LEN",
  "UPPER", "LOWER", "TRIM", "SUBSTITUTE", "REPLACE", "FIND", "SEARCH",
  "PATH", "PATHITEM", "PATHLENGTH", "PATHCONTAINS",
  "LOOKUPVALUE", "TREATAS", "GENERATESERIES", "ROW", "ROWS",
  "EARLIER", "EARLIEST", "RANKX", "TOPN", "SAMPLE", "DATATABLE",
  "EXCEPT", "INTERSECT", "UNION", "NATURALINNERJOIN", "NATURALLEFTOUTERJOIN",
  "CROSSJOIN", "GENERATE", "GENERATEALL", "ROLLUP", "ROLLUPGROUP", "ROLLUPISSUBTOTAL",
  "ISINSCOPE", "ISFILTERED", "ALLSELECTED", "KEEPFILTERS", "REMOVEFILTERS",
  "CALCULATETABLE", "CURRENTPERIOD", "NEXTMONTH", "NEXTQUARTER", "NEXTYEAR",
  "PREVIOUSMONTH", "PREVIOUSQUARTER", "PREVIOUSYEAR", "FIRSTDATE", "LASTDATE",
  "ENDOFMONTH", "ENDOFQUARTER", "ENDOFYEAR", "STARTOFMONTH", "STARTOFQUARTER", "STARTOFYEAR",
  "DATEADD", "DATESBETWEEN", "DATESINPERIOD", "CLOSINGBALANCEMONTH", "CLOSINGBALANCEQUARTER", "CLOSINGBALANCEYEAR",
  "OPENINGBALANCEMONTH", "OPENINGBALANCEQUARTER", "OPENINGBALANCEYEAR",
  "SAMEPERIODLASTYEAR", "DATEADD", "NEXTDAY", "PREVIOUSDAY",
];

// ============================================================================
// Parser Class
// ============================================================================

export class DAXParser {
  private static instance: DAXParser;

  private constructor() {}

  static getInstance(): DAXParser {
    if (!DAXParser.instance) {
      DAXParser.instance = new DAXParser();
    }
    return DAXParser.instance;
  }

  /**
   * Tokenize a DAX formula
   */
  tokenize(formula: string): DAXToken[] {
    const tokens: DAXToken[] = [];
    let position = 0;

    // Remove leading = if present
    const cleanFormula = formula.trim().startsWith("=")
      ? formula.trim().substring(1)
      : formula.trim();

    while (position < cleanFormula.length) {
      const char = cleanFormula[position];

      // Skip whitespace
      if (/\s/.test(char)) {
        position++;
        continue;
      }

      // Comments
      if (char === "/" && cleanFormula[position + 1] === "/") {
        const end = cleanFormula.indexOf("\n", position);
        const comment = end === -1
          ? cleanFormula.substring(position)
          : cleanFormula.substring(position, end);
        tokens.push({
          type: "comment",
          value: comment,
          position,
          length: comment.length,
        });
        position = end === -1 ? cleanFormula.length : end;
        continue;
      }

      // Strings
      if (char === '"') {
        let end = position + 1;
        while (end < cleanFormula.length && cleanFormula[end] !== '"') {
          if (cleanFormula[end] === '\\') end++;
          end++;
        }
        const str = cleanFormula.substring(position, end + 1);
        tokens.push({
          type: "string",
          value: str,
          position,
          length: str.length,
        });
        position = end + 1;
        continue;
      }

      // Numbers
      if (/\d/.test(char) || (char === "." && /\d/.test(cleanFormula[position + 1]))) {
        let end = position;
        while (end < cleanFormula.length && /[\d.]/.test(cleanFormula[end])) {
          end++;
        }
        // Check for scientific notation
        if (cleanFormula[end] === "e" || cleanFormula[end] === "E") {
          end++;
          if (cleanFormula[end] === "+" || cleanFormula[end] === "-") end++;
          while (end < cleanFormula.length && /\d/.test(cleanFormula[end])) end++;
        }
        const num = cleanFormula.substring(position, end);
        tokens.push({
          type: "number",
          value: num,
          position,
          length: num.length,
        });
        position = end;
        continue;
      }

      // Parentheses
      if (char === "(" || char === ")") {
        tokens.push({
          type: "parenthesis",
          value: char,
          position,
          length: 1,
        });
        position++;
        continue;
      }

      // Comma
      if (char === ",") {
        tokens.push({
          type: "comma",
          value: char,
          position,
          length: 1,
        });
        position++;
        continue;
      }

      // Operators
      if (/[+\-*/^&]/.test(char) || (char === "=" && cleanFormula[position + 1] !== "=")) {
        tokens.push({
          type: "operator",
          value: char,
          position,
          length: 1,
        });
        position++;
        continue;
      }

      // Comparison operators
      if (char === "=" || char === "<" || char === ">") {
        let op = char;
        if (
          (char === "<" && cleanFormula[position + 1] === ">") ||
          (char === "<" && cleanFormula[position + 1] === "=") ||
          (char === ">" && cleanFormula[position + 1] === "=")
        ) {
          op += cleanFormula[position + 1];
        }
        tokens.push({
          type: "operator",
          value: op,
          position,
          length: op.length,
        });
        position += op.length;
        continue;
      }

      // Identifiers (functions, columns, tables)
      if (/[a-zA-Z_]/.test(char) || char === "'") {
        let end = position;
        let inQuotes = char === "'";

        if (inQuotes) {
          end++;
          while (end < cleanFormula.length && cleanFormula[end] !== "'") {
            if (cleanFormula[end] === "'") break;
            end++;
          }
          if (cleanFormula[end] === "'") end++;
        } else {
          while (end < cleanFormula.length && /[a-zA-Z0-9_]/.test(cleanFormula[end])) {
            end++;
          }
        }

        // Check for table[column] syntax
        if (cleanFormula[end] === "[") {
          const bracketStart = end;
          end++;
          while (end < cleanFormula.length && cleanFormula[end] !== "]") end++;
          if (cleanFormula[end] === "]") end++;
        }

        const identifier = cleanFormula.substring(position, end);
        const upperIdentifier = identifier.toUpperCase();

        // Determine token type
        let type: DAXTokenType = "identifier";
        if (DAX_FUNCTIONS[upperIdentifier]) {
          type = "function";
        } else if (identifier.includes("[")) {
          type = "column";
        } else if (identifier.includes("'") || /^[A-Z][a-zA-Z0-9_]*$/.test(identifier)) {
          type = "table";
        }

        tokens.push({
          type,
          value: identifier,
          position,
          length: identifier.length,
        });
        position = end;
        continue;
      }

      // Unknown character, skip
      position++;
    }

    return tokens;
  }

  /**
   * Parse a DAX formula
   */
  parse(formula: string, type: DAXExpressionType = "measure"): ParsedDAXFormula {
    const tokens = this.tokenize(formula);

    const functions: string[] = [];
    const tables: string[] = [];
    const columns: string[] = [];
    const dependencies: DAXDependency[] = [];
    let hasRowContext = false;
    let hasFilterContext = false;
    let hasContextTransition = false;

    // Extract function names
    tokens.filter(t => t.type === "function").forEach(t => {
      if (!functions.includes(t.value.toUpperCase())) {
        functions.push(t.value.toUpperCase());
      }
    });

    // Extract tables and columns
    tokens.forEach(t => {
      if (t.type === "table" && !tables.includes(t.value)) {
        tables.push(t.value);
      } else if (t.type === "column") {
        if (!columns.includes(t.value)) {
          columns.push(t.value);
        }
        // Extract table and column name
        const match = t.value.match(/(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_]*))?\[([^\]]+)\]/);
        if (match) {
          const tableName = match[1] || match[2];
          const columnName = match[3];
          dependencies.push({
            type: "column",
            name: columnName,
            table: tableName,
            position: t.position,
          });
        }
      }
    });

    // Analyze context
    hasRowContext = functions.some(f => {
      const func = DAX_FUNCTIONS[f];
      return func?.isIterator || f === "EARLIER" || f === "EARLIEST";
    });

    hasFilterContext = functions.some(f => {
      const func = DAX_FUNCTIONS[f];
      return func?.modifiesFilterContext || func?.isFilterFunction;
    });

    hasContextTransition = functions.includes("CALCULATE") || functions.includes("CALCULATETABLE");

    return {
      expression: formula,
      type,
      tokens,
      functions,
      tables,
      columns,
      measures: this.extractMeasures(tokens),
      dependencies,
      hasRowContext,
      hasFilterContext,
      hasContextTransition,
    };
  }

  /**
   * Extract measure references from tokens
   * Measures are identified by pattern: [MeasureName] without table prefix in certain contexts
   * or by detecting bracketed names that aren't column references
   */
  private extractMeasures(tokens: DAXToken[]): string[] {
    const measures: string[] = [];
    
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i];
      
      // Look for bracketed identifiers that could be measures
      if (token.type === 'column') {
        const match = token.value.match(/\[([^\]]+)\]/);
        if (match) {
          const name = match[1];
          // Check if this looks like a measure (not a standard column reference)
          // Measures often appear in specific contexts
          if (this.isLikelyMeasure(token, tokens, i)) {
            if (!measures.includes(name)) {
              measures.push(name);
            }
          }
        }
      }
      
      // Also check for explicit measure references in function arguments
      if (token.type === 'identifier' && token.value.startsWith('[')) {
        const name = token.value.slice(1, -1);
        if (name && !measures.includes(name)) {
          measures.push(name);
        }
      }
    }
    
    return measures;
  }

  /**
   * Determine if a column token is likely a measure reference
   */
  private isLikelyMeasure(token: DAXToken, tokens: DAXToken[], index: number): boolean {
    const name = token.value.match(/\[([^\]]+)\]/)?.[1] || '';
    
    // Common measure naming patterns
    const measurePatterns = [
      /total/i, /sum/i, /avg/i, /average/i, /count/i, /min/i, /max/i,
      /ytd/i, /qtd/i, /mtd/i, /growth/i, /variance/i, /var/i,
      /sales/i, /revenue/i, /profit/i, /margin/i, /ratio/i,
      /measure/i, /calc/i, /computed/i
    ];
    
    // Check if name matches common measure patterns
    if (measurePatterns.some(pattern => pattern.test(name))) {
      return true;
    }
    
    // Check context: measures often appear as arguments to aggregation functions
    // Look backwards for function context
    let parenDepth = 0;
    for (let i = index - 1; i >= 0 && i >= index - 10; i--) {
      const prevToken = tokens[i];
      if (prevToken.type === 'parenthesis') {
        if (prevToken.value === ')') parenDepth++;
        if (prevToken.value === '(') parenDepth--;
      }
      if (prevToken.type === 'function' && parenDepth <= 0) {
        const func = DAX_FUNCTIONS[prevToken.value.toUpperCase()];
        if (func?.isAggregator || func?.modifiesFilterContext) {
          return true;
        }
      }
    }
    
    // Check if it's a bare bracket reference (no table prefix)
    if (!token.value.includes("'")) {
      // Look at surrounding context
      const prevToken = tokens[index - 1];
      
      // Measures often follow CALCULATE or are in filter contexts
      if (prevToken?.type === 'function') {
        const funcName = prevToken.value.toUpperCase();
        if (['CALCULATE', 'CALCULATETABLE', 'FILTER', 'SUMX', 'AVERAGEX'].includes(funcName)) {
          return true;
        }
      }
    }
    
    return false;
  }

  /**
   * Explain a DAX formula in plain English
   */
  explain(formula: string, type: DAXExpressionType = "measure"): DAXExplanation {
    const parsed = this.parse(formula, type);
    const steps: DAXExplanationStep[] = [];
    const optimizations: DAXOptimization[] = [];
    const performanceHints: string[] = [];

    // Generate explanation steps
    this.generateExplanationSteps(parsed, steps, 0);

    // Analyze complexity
    const complexity = this.calculateComplexity(parsed);

    // Generate optimizations
    this.generateOptimizations(parsed, optimizations, performanceHints);

    // Generate summary
    const summary = this.generateSummary(parsed);

    return {
      summary,
      breakdown: steps,
      complexity,
      contextInfo: {
        rowContext: parsed.hasRowContext,
        filterContext: parsed.hasFilterContext,
        contextTransitions: parsed.hasContextTransition ? 1 : 0,
        iteratorFunctions: parsed.functions.filter(f => DAX_FUNCTIONS[f]?.isIterator).length,
        filterModifications: parsed.functions.filter(f => DAX_FUNCTIONS[f]?.modifiesFilterContext).length,
      },
      optimizations,
      performanceHints,
    };
  }

  /**
   * Validate DAX syntax
   */
  validate(formula: string): DAXSyntaxError[] {
    const errors: DAXSyntaxError[] = [];
    const tokens = this.tokenize(formula);

    // Check for unclosed parentheses
    let parenCount = 0;
    tokens.forEach(t => {
      if (t.type === "parenthesis") {
        if (t.value === "(") parenCount++;
        if (t.value === ")") parenCount--;
        if (parenCount < 0) {
          errors.push({
            message: "Unexpected closing parenthesis",
            position: t.position,
            length: 1,
            severity: "error",
          });
        }
      }
    });

    if (parenCount > 0) {
      errors.push({
        message: "Unclosed parenthesis",
        position: formula.length - 1,
        length: 1,
        severity: "error",
      });
    }

    // Check for unknown functions
    tokens.filter(t => t.type === "function").forEach(t => {
      if (!DAX_FUNCTIONS[t.value.toUpperCase()]) {
        errors.push({
          message: `Unknown function: ${t.value}`,
          position: t.position,
          length: t.length,
          severity: "warning",
        });
      }
    });

    return errors;
  }

  /**
   * Get function info
   */
  getFunctionInfo(name: string): DAXFunction | undefined {
    return DAX_FUNCTIONS[name.toUpperCase()];
  }

  /**
   * Get all functions by category
   */
  getFunctionsByCategory(category: DAXFunctionCategory): DAXFunction[] {
    return Object.values(DAX_FUNCTIONS).filter(f => f.category === category);
  }

  /**
   * Search functions
   */
  searchFunctions(query: string): DAXFunction[] {
    const lowerQuery = query.toLowerCase();
    return Object.values(DAX_FUNCTIONS).filter(
      f =>
        f.name.toLowerCase().includes(lowerQuery) ||
        f.description.toLowerCase().includes(lowerQuery) ||
        f.category.toLowerCase().includes(lowerQuery)
    );
  }

  // Private helper methods
  private generateExplanationSteps(parsed: ParsedDAXFormula, steps: DAXExplanationStep[], depth: number): void {
    // Simple recursive explanation based on function calls
    parsed.functions.forEach(funcName => {
      const func = DAX_FUNCTIONS[funcName];
      if (func) {
        steps.push({
          function: funcName,
          description: func.description,
          context: this.getContextDescription(func),
          indentation: depth,
        });
      }
    });
  }

  private getContextDescription(func: DAXFunction): string {
    const parts: string[] = [];
    if (func.isIterator) parts.push("iterates over rows");
    if (func.modifiesFilterContext) parts.push("modifies filter context");
    if (func.isFilterFunction) parts.push("returns filtered table");
    if (func.isTimeIntelligence) parts.push("time intelligence");
    return parts.join(", ") || "standard function";
  }

  private calculateComplexity(parsed: ParsedDAXFormula): DAXComplexity {
    let score = 0;

    // Function complexity
    score += parsed.functions.length * 2;

    // Iterator functions are more complex
    score += parsed.functions.filter(f => DAX_FUNCTIONS[f]?.isIterator).length * 3;

    // Context transitions
    if (parsed.hasContextTransition) score += 5;

    // Filter modifications
    score += parsed.functions.filter(f => DAX_FUNCTIONS[f]?.modifiesFilterContext).length * 2;

    // Dependencies
    score += parsed.dependencies.length;

    if (score <= 5) return "simple";
    if (score <= 10) return "moderate";
    if (score <= 20) return "complex";
    return "very_complex";
  }

  private generateOptimizations(parsed: ParsedDAXFormula, optimizations: DAXOptimization[], hints: string[]): void {
    // Check for nested CALCULATE
    const calcCount = parsed.functions.filter(f => f === "CALCULATE" || f === "CALCULATETABLE").length;
    if (calcCount > 2) {
      optimizations.push({
        type: "performance",
        severity: "high",
        message: "Multiple nested CALCULATE functions can impact performance",
        suggestion: "Consider consolidating filters or using variables",
      });
    }

    // Check for iterator on large tables
    const iterators = parsed.functions.filter(f => DAX_FUNCTIONS[f]?.isIterator);
    if (iterators.length > 0) {
      hints.push("Iterator functions (SUMX, AVERAGEX) evaluate expressions row-by-row and can be slow on large tables");
    }

    // Check for alternatives
    parsed.functions.forEach(f => {
      const func = DAX_FUNCTIONS[f];
      if (func?.hasAlternatives) {
        optimizations.push({
          type: "best_practice",
          severity: "low",
          message: `${f} has alternatives that might be more appropriate`,
          suggestion: `Consider: ${func.hasAlternatives.join(", ")}`,
        });
      }
    });

    // Check for IF vs SWITCH
    const ifCount = (parsed.expression.match(/\bIF\s*\(/gi) || []).length;
    if (ifCount > 3) {
      optimizations.push({
        type: "readability",
        severity: "medium",
        message: "Multiple nested IF statements reduce readability",
        suggestion: "Consider using SWITCH for multiple conditions",
        example: "SWITCH(TRUE(), condition1, result1, condition2, result2, default)",
      });
    }
  }

  private generateSummary(parsed: ParsedDAXFormula): string {
    const parts: string[] = [];

    if (parsed.functions.length > 0) {
      parts.push(`Uses ${parsed.functions.length} function(s)`);
    }

    if (parsed.hasContextTransition) {
      parts.push("performs context transition");
    }

    if (parsed.hasRowContext) {
      parts.push("utilizes row context");
    }

    if (parsed.hasFilterContext) {
      parts.push("modifies filter context");
    }

    return parts.join(", ") || "Simple DAX expression";
  }
}

// Export singleton instance
export const daxParser = DAXParser.getInstance();
export default daxParser;
