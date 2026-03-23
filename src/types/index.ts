// Excel AI Assistant - Type Definitions

export interface Message {
  id: string;
  role: "user" | "assistant" | "system";
  content: string;
  timestamp: Date;
  actions?: AIAction[];
}

// ============================================================================
// STRONGLY TYPED ACTION PAYLOADS (Phase 1 - Type Safety)
// ============================================================================

/**
 * All possible action types
 */
export type AIActionType =
  | "insert_formula"
  | "set_values"
  | "create_chart"
  | "delete_chart"
  | "create_table"
  | "delete_table"
  | "format_cells"
  | "add_validation"
  | "remove_validation"
  | "clear_range"
  | "create_pivot_table"
  | "refresh_pivot_table"
  | "add_worksheet"
  | "delete_worksheet"
  | "rename_worksheet"
  | "auto_fit_columns"
  | "create_named_range"
  | "delete_named_range"
  | "add_comment"
  | "delete_comment"
  | "add_hyperlink"
  | "remove_hyperlink"
  | "create_vba_macro"
  | "explain_vba_code"
  | "refactor_vba_code"
  | "explain"
  | "apply_suggestion"
  // New action types
  | "duplicate_worksheet"
  | "move_worksheet"
  | "hide_worksheet"
  | "show_worksheet"
  | "hide_rows"
  | "show_rows"
  | "hide_columns"
  | "show_columns"
  | "freeze_panes"
  | "unfreeze_panes"
  | "insert_rows"
  | "insert_columns"
  | "delete_rows"
  | "delete_columns"
  | "set_row_height"
  | "set_column_width"
  | "group_rows"
  | "ungroup_rows"
  | "group_columns"
  | "ungroup_columns"
  | "convert_table_to_range"
  | "convert_formulas_to_values"
  | "set_print_area"
  | "set_page_orientation"
  | "find_and_replace"
  | "duplicate_chart"
  | "move_column"
  | "move_row";

// ============================================================================
// INDIVIDUAL ACTION PAYLOAD TYPES
// ============================================================================

/** Base payload with common properties */
interface BasePayload {
  worksheetName?: string;
}

/** Insert a formula into a cell or range */
export interface InsertFormulaPayload extends BasePayload {
  formula: string;
  address: string;
}

/** Set values in a range */
export interface SetValuesPayload extends BasePayload {
  values: any[][];
  address: string;
}

/** Create a chart */
export interface CreateChartPayload extends BasePayload {
  dataRange: string;
  type: 'columnClustered' | 'columnStacked' | 'barClustered' | 'line' | 'lineMarkers' | 'pie' | 'doughnut' | 'scatter' | 'area' | 'radar';
  title?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
  legendPosition?: 'bottom' | 'top' | 'left' | 'right';
  dataLabels?: boolean;
  trendline?: boolean;
}

/** Delete a chart */
export interface DeleteChartPayload extends BasePayload {
  name: string;
}

/** Create a table */
export interface CreateTablePayload extends BasePayload {
  range: string;
  name?: string;
  hasHeaders: boolean;
  style?: string;
  showTotals?: boolean;
  totalColumns?: string[];
}

/** Delete a table */
export interface DeleteTablePayload extends BasePayload {
  name: string;
  keepData?: boolean;
}

/** Format cells */
export interface FormatCellsPayload extends BasePayload {
  address: string;
  numberFormat?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fillColor?: string;
  fontSize?: number;
  fontName?: string;
  horizontalAlignment?: 'left' | 'center' | 'right' | 'justify' | 'distributed';
  verticalAlignment?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
}

/** Add data validation */
export interface AddValidationPayload extends BasePayload {
  address: string;
  type: 'list' | 'whole' | 'decimal' | 'date' | 'textLength' | 'custom';
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  inputMessage?: string;
  inputTitle?: string;
  showErrorMessage?: boolean;
  errorMessage?: string;
  errorTitle?: string;
  errorStyle?: 'stop' | 'warning' | 'information';
}

/** Remove data validation */
export interface RemoveValidationPayload extends BasePayload {
  address: string;
}

/** Clear a range */
export interface ClearRangePayload extends BasePayload {
  address: string;
  clearContents?: boolean;
  clearFormat?: boolean;
}

/** Create a pivot table */
export interface CreatePivotTablePayload extends BasePayload {
  name: string;
  sourceData: string;
  destination?: string;
  rowFields?: string[];
  columnFields?: string[];
  dataFields?: Array<{
    name: string;
    function: 'sum' | 'count' | 'average' | 'max' | 'min' | 'product';
  }>;
  filterFields?: string[];
}

/** Refresh a pivot table */
export interface RefreshPivotTablePayload extends BasePayload {
  name: string;
}

/** Add a worksheet */
export interface AddWorksheetPayload {
  name?: string;
}

/** Delete a worksheet */
export interface DeleteWorksheetPayload {
  name: string;
}

/** Rename a worksheet */
export interface RenameWorksheetPayload {
  name: string;
  newName: string;
}

/** Auto-fit columns */
export interface AutoFitColumnsPayload extends BasePayload {
  address: string;
}

/** Create a named range */
export interface CreateNamedRangePayload extends BasePayload {
  name: string;
  address: string;
  comment?: string;
}

/** Delete a named range */
export interface DeleteNamedRangePayload {
  name: string;
}

/** Add a comment */
export interface AddCommentPayload extends BasePayload {
  address: string;
  text: string;
  author?: string;
}

/** Delete a comment */
export interface DeleteCommentPayload extends BasePayload {
  address: string;
}

/** Add a hyperlink */
export interface AddHyperlinkPayload extends BasePayload {
  address: string;
  link: string;
  displayText?: string;
  screenTip?: string;
}

/** Remove a hyperlink */
export interface RemoveHyperlinkPayload extends BasePayload {
  address: string;
}

/** Create a VBA macro */
export interface CreateVBAMacroPayload {
  macroName: string;
  description: string;
  worksheetName?: string;
}

/** Explain VBA code */
export interface ExplainVBACodePayload {
  vbaCode: string;
  detailLevel?: 'brief' | 'detailed' | 'comprehensive';
}

/** Refactor VBA code */
export interface RefactorVBACodePayload {
  vbaCode: string;
  goals?: Array<'performance' | 'readability' | 'error_handling' | 'modern_vba'>;
}

/** Explain something */
export interface ExplainPayload {
  target: string;
  context?: string;
}

/** Apply a suggestion */
export interface ApplySuggestionPayload {
  suggestionId: string;
  suggestion: string;
}

// ============================================================================
// DISCRIMINATED UNION FOR ALL ACTION PAYLOADS
// ============================================================================

/**
 * Maps action types to their corresponding payload types
 */
export type ActionPayloadMap = {
  insert_formula: InsertFormulaPayload;
  set_values: SetValuesPayload;
  create_chart: CreateChartPayload;
  delete_chart: DeleteChartPayload;
  create_table: CreateTablePayload;
  delete_table: DeleteTablePayload;
  format_cells: FormatCellsPayload;
  add_validation: AddValidationPayload;
  remove_validation: RemoveValidationPayload;
  clear_range: ClearRangePayload;
  create_pivot_table: CreatePivotTablePayload;
  refresh_pivot_table: RefreshPivotTablePayload;
  add_worksheet: AddWorksheetPayload;
  delete_worksheet: DeleteWorksheetPayload;
  rename_worksheet: RenameWorksheetPayload;
  auto_fit_columns: AutoFitColumnsPayload;
  create_named_range: CreateNamedRangePayload;
  delete_named_range: DeleteNamedRangePayload;
  add_comment: AddCommentPayload;
  delete_comment: DeleteCommentPayload;
  add_hyperlink: AddHyperlinkPayload;
  remove_hyperlink: RemoveHyperlinkPayload;
  create_vba_macro: CreateVBAMacroPayload;
  explain_vba_code: ExplainVBACodePayload;
  refactor_vba_code: RefactorVBACodePayload;
  explain: ExplainPayload;
  apply_suggestion: ApplySuggestionPayload;
};

/**
 * Discriminated union for type-safe actions
 */
export type TypedAIAction = {
  [K in AIActionType]: {
    type: K;
    label: string;
    payload: ActionPayloadMap[K];
  };
}[AIActionType];

/**
 * AI Action interface - supports both typed and legacy formats
 * Use TypedAIAction for new code, AIAction for backward compatibility
 */
export interface AIAction {
  type: AIActionType;
  label: string;
  payload?: ActionPayloadMap[AIActionType] | Record<string, unknown>;
}

/**
 * Type guard to check if an action has a specific type
 */
export function isActionType<T extends AIActionType>(
  action: AIAction,
  type: T
): action is AIAction & { type: T; payload: ActionPayloadMap[T] } {
  return action.type === type;
}

/**
 * Type guard for destructive actions that require confirmation
 */
export function isDestructiveAction(actionType: AIActionType): boolean {
  const destructiveTypes: AIActionType[] = [
    'delete_chart',
    'delete_table',
    'delete_worksheet',
    'delete_named_range',
    'delete_comment',
    'clear_range'
  ];
  return destructiveTypes.includes(actionType);
}

// ============================================================================
// LEGACY TYPE EXPORTS (for backward compatibility)
// ============================================================================

export interface AIRequest {
  message: string;
  context: ExcelContext;
  conversationHistory: Message[];
  settings: AISettings;
}

export interface AIResponse {
  message: string;
  actions?: AIAction[];
  suggestedPrompts?: string[];
  requiresConfirmation?: boolean;
}

export interface AISettings {
  apiUrl: string;
  apiKey: string;
  model: string;
  temperature: number;
  maxTokens: number;
  systemPrompt?: string;
}

export interface ExcelContext {
  workbook: WorkbookContext;
  selection?: RangeContext;
  activeWorksheet: WorksheetContext;
  tables: TableContext[];
  charts: ChartContext[];
  pivotTables: PivotTableContext[];
  dataModel?: DataModelContext;
  powerQuery?: PowerQueryContext;
}

export interface WorkbookContext {
  name: string;
  worksheets: string[];
  namedRanges: string[];
}

export interface WorksheetContext {
  name: string;
  index?: number;
  usedRange?: RangeContext;
  tables: string[];
  charts: string[];
}

export interface RangeContext {
  address: string;
  worksheetName: string;
  values: any[][];
  formulas?: string[][];
  numberFormat?: string[][];
  rowCount: number;
  columnCount: number;
}

export interface TableContext {
  name: string;
  worksheetName: string;
  range: string;
  headerRowRange: string;
  dataBodyRange: string;
  headers: string[];
  rowCount: number;
  columnCount: number;
  style?: string;
}

export interface ChartContext {
  name: string;
  type: string;
  worksheetName: string;
}

export interface PivotTableContext {
  name: string;
  worksheetName: string;
  sourceData?: string;
}

export interface DataModelContext {
  tables: DataModelTable[];
  relationships: DataModelRelationship[];
}

export interface DataModelTable {
  name: string;
  source: string;
  columns: string[];
}

export interface DataModelRelationship {
  name: string;
  fromTable: string;
  fromColumn: string;
  toTable: string;
  toColumn: string;
}

export interface PowerQueryContext {
  queries: PowerQueryInfo[];
}

export interface PowerQueryInfo {
  name: string;
  formula: string;
  connection?: string;
}

export interface CellValidation {
  type: "list" | "whole" | "decimal" | "date" | "textLength" | "custom";
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  inputMessage?: string;
  inputTitle?: string;
  showErrorMessage?: boolean;
  errorMessage?: string;
  errorTitle?: string;
  errorStyle?: "stop" | "warning" | "information";
}

export interface FormatOptions {
  numberFormat?: string;
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
  };
  fill?: {
    color?: string;
    pattern?: string;
  };
  borders?: {
    style?: string;
    color?: string;
    weight?: string;
  };
  alignment?: {
    horizontal?: "left" | "center" | "right" | "justify" | "distributed";
    vertical?: "top" | "center" | "bottom" | "justify" | "distributed";
    wrapText?: boolean;
  };
}

export interface ChartOptions {
  type: string;
  title?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
  legendPosition?: "bottom" | "top" | "left" | "right";
  dataLabels?: boolean;
  trendline?: boolean;
}

export interface TableOptions {
  name: string;
  hasHeaders: boolean;
  style?: string;
  showTotals?: boolean;
  totalColumns?: string[];
}

export interface PivotTableOptions {
  name: string;
  sourceData: string;
  destination?: string;
  rowFields?: string[];
  columnFields?: string[];
  dataFields?: PivotDataField[];
  filterFields?: string[];
}

export interface PivotDataField {
  name: string;
  function: "sum" | "count" | "average" | "max" | "min" | "product";
}

export interface PowerQueryOperation {
  type: "source" | "transform" | "merge" | "append" | "group" | "pivot";
  parameters: Record<string, any>;
}

// Cell reference types for visual highlighting and parsing
export interface ExcelRange {
  sheetId: string | null;
  address: string;
  startCell: {
    column: number;  // 1-based column index (A=1, B=2, etc.)
    row: number;     // 1-based row index
  };
  endCell: {
    column: number;
    row: number;
  };
}

export interface ParsedCellReference {
  /** The original reference string found in text */
  original: string;
  /** The sheet name (null if current sheet) */
  sheetName: string | null;
  /** The parsed range */
  range: ExcelRange;
  /** Type of reference */
  type: 'cell' | 'range' | 'table' | 'named' | 'r1c1';
  /** Whether the reference is absolute ($A$1) */
  isAbsolute: boolean;
  /** Start position in the original text */
  startIndex: number;
  /** End position in the original text */
  endIndex: number;
}
