// Action Handler - Executes AI actions in Excel
import ExcelService from "./excelService";
import CommentService from "./commentService";
import HyperlinkService from "./hyperlinkService";
import MacroService from "./macroService";
import { AIAction, CellValidation } from "@/types";
import logger from "@/utils/logger";

// Whitelist of allowed action types for security
const ALLOWED_ACTIONS = new Set([
  "insert_formula",
  "set_values",
  "clear_range",
  "create_table",
  "delete_table",
  "create_chart",
  "delete_chart",
  "create_pivot_table",
  "refresh_pivot_table",
  "format_cells",
  "auto_fit_columns",
  "add_validation",
  "remove_validation",
  "add_worksheet",
  "delete_worksheet",
  "rename_worksheet",
  "create_named_range",
  "delete_named_range",
  "add_comment",
  "delete_comment",
  "add_hyperlink",
  "remove_hyperlink",
  // VBA Macro actions
  "create_vba_macro",
  "explain_vba_code",
  "refactor_vba_code",
  // New action types
  "duplicate_worksheet",
  "move_worksheet",
  "hide_worksheet",
  "show_worksheet",
  "hide_rows",
  "show_rows",
  "hide_columns",
  "show_columns",
  "freeze_panes",
  "unfreeze_panes",
  "insert_rows",
  "insert_columns",
  "delete_rows",
  "delete_columns",
  "set_row_height",
  "set_column_width",
  "group_rows",
  "ungroup_rows",
  "group_columns",
  "ungroup_columns",
  "convert_table_to_range",
  "convert_formulas_to_values",
  "set_print_area",
  "set_page_orientation",
  "find_and_replace",
  "duplicate_chart",
  "move_column",
  "move_row"
]);

// Destructive actions that require confirmation
const DESTRUCTIVE_ACTIONS = new Set([
  "delete_table",
  "delete_worksheet",
  "clear_range",
  "delete_chart",
  "delete_named_range",
  "delete_comment",
  "delete_rows",
  "delete_columns",
  "convert_formulas_to_values"
]);

export class ActionHandler {
  private static instance: ActionHandler;

  private constructor() {}

  static getInstance(): ActionHandler {
    if (!ActionHandler.instance) {
      ActionHandler.instance = new ActionHandler();
    }
    return ActionHandler.instance;
  }

  /**
   * Validate action before execution
   */
  private validateAction(action: AIAction): void {
    if (!action) {
      throw new Error("Action is required");
    }
    if (!action.type) {
      throw new Error("Action type is required");
    }
    if (!ALLOWED_ACTIONS.has(action.type)) {
      throw new Error(`Action type "${action.type}" is not allowed`);
    }
    if (!action.payload) {
      throw new Error("Action payload is required");
    }
  }

  /**
   * Check if action is destructive
   */
  isDestructive(actionType: string): boolean {
    return DESTRUCTIVE_ACTIONS.has(actionType);
  }

  async executeAction(action: AIAction): Promise<string> {
    // Validate action before execution
    this.validateAction(action);

    // Ensure payload exists (validation should guarantee this, but TypeScript doesn't know)
    if (!action.payload) {
      throw new Error("Action payload is required");
    }
    
    const { type, payload } = action;
    
    // Cast payload to a typed record for safer property access
    const p = payload as Record<string, unknown>;

    try {
      switch (type) {
        // Cell & Formula Operations
        case "insert_formula":
          if (!p.address || !p.formula) {
            throw new Error("Missing required fields: address and formula");
          }
          await ExcelService.setFormulas(
            String(p.address),
            [[String(p.formula)]],
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Formula inserted at ${p.address}`;

        case "set_values":
          if (!p.address || !p.values) {
            throw new Error("Missing required fields: address and values");
          }
          await ExcelService.setValues(
            String(p.address),
            p.values as any[][],
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Values set at ${p.address}`;

        case "clear_range":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await ExcelService.clearRange(
            String(p.address), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Range ${p.address} cleared`;

        // Table Operations
        case "create_table":
          if (!p.range) {
            throw new Error("Missing required field: range");
          }
          const table = await ExcelService.createTable(
            {
              name: p.name ? String(p.name) : undefined,
              hasHeaders: p.hasHeaders !== undefined ? Boolean(p.hasHeaders) : true,
              style: p.style ? String(p.style) : undefined,
              showTotals: p.showTotals !== undefined ? Boolean(p.showTotals) : undefined,
              totalColumns: p.totalColumns as string[] | undefined
            },
            String(p.range),
            p.worksheetName ? String(p.worksheetName) : undefined
          ) as { name: string; range: string };
          return `Table "${table.name}" created at ${table.range}`;

        case "delete_table":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.deleteTable(
            String(p.name), 
            p.keepData !== undefined ? Boolean(p.keepData) : true, 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Table "${p.name}" deleted`;

        // Chart Operations
        case "create_chart":
          if (!p.dataRange || !p.type) {
            throw new Error("Missing required fields: dataRange and type");
          }
          const chart = await ExcelService.createChart(
            String(p.dataRange),
            {
              type: String(p.type),
              title: p.title ? String(p.title) : undefined
            },
            p.worksheetName ? String(p.worksheetName) : undefined
          ) as { name: string };
          return `Chart "${chart.name}" created`;

        case "delete_chart":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.deleteChart(
            String(p.name), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Chart "${p.name}" deleted`;

        // Pivot Table Operations
        case "create_pivot_table":
          if (!p.name || !p.sourceData) {
            throw new Error("Missing required fields: name and sourceData");
          }
          const pivot = await ExcelService.createPivotTable(
            {
              name: String(p.name),
              sourceData: String(p.sourceData),
              destination: p.destination ? String(p.destination) : undefined,
              rowFields: p.rowFields as string[] | undefined,
              columnFields: p.columnFields as string[] | undefined,
              dataFields: p.dataFields as any[] | undefined,
              filterFields: p.filterFields as string[] | undefined
            },
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Pivot table "${pivot.name}" created`;

        case "refresh_pivot_table":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.refreshPivotTable(
            String(p.name), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Pivot table "${p.name}" refreshed`;

        // Formatting
        case "format_cells":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await ExcelService.formatRange(
            String(p.address),
            {
              numberFormat: p.numberFormat ? String(p.numberFormat) : undefined,
              font: {
                bold: p.bold !== undefined ? Boolean(p.bold) : undefined,
                italic: p.italic !== undefined ? Boolean(p.italic) : undefined,
                color: p.color ? String(p.color) : undefined
              },
              fill: {
                color: p.fillColor ? String(p.fillColor) : undefined
              }
            },
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Cells at ${p.address} formatted`;

        case "auto_fit_columns":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await ExcelService.autoFitColumns(
            String(p.address), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Columns auto-fitted`;

        // Data Validation
        case "add_validation":
          if (!p.address || !p.type) {
            throw new Error("Missing required fields: address and type");
          }
          await ExcelService.addValidation(
            String(p.address),
            {
              type: String(p.type) as CellValidation['type'],
              formula1: p.formula1 ? String(p.formula1) : undefined,
              formula2: p.formula2 ? String(p.formula2) : undefined,
              allowBlank: p.allowBlank !== undefined ? Boolean(p.allowBlank) : undefined,
              showInputMessage: p.showInputMessage !== undefined ? Boolean(p.showInputMessage) : (p.inputMessage ? true : false),
              inputMessage: p.inputMessage ? String(p.inputMessage) : undefined,
              inputTitle: p.inputTitle ? String(p.inputTitle) : undefined,
              showErrorMessage: p.showErrorMessage !== undefined ? Boolean(p.showErrorMessage) : (p.errorMessage ? true : false),
              errorMessage: p.errorMessage ? String(p.errorMessage) : undefined,
              errorTitle: p.errorTitle ? String(p.errorTitle) : undefined,
              errorStyle: p.errorStyle ? String(p.errorStyle) as CellValidation['errorStyle'] : undefined
            },
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Data validation added to ${p.address}`;

        case "remove_validation":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await ExcelService.clearValidation(
            String(p.address), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Data validation removed from ${p.address}`;

        // Worksheet Operations
        case "add_worksheet":
          const newSheetName = await ExcelService.addWorksheet(p.name ? String(p.name) : undefined);
          return `Worksheet "${newSheetName}" created`;

        case "delete_worksheet":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.deleteWorksheet(String(p.name));
          return `Worksheet "${p.name}" deleted`;

        case "rename_worksheet":
          if (!p.name || !p.newName) {
            throw new Error("Missing required fields: name and newName");
          }
          await ExcelService.renameWorksheet(String(p.name), String(p.newName));
          return `Worksheet "${p.name}" renamed to "${p.newName}"`;

        // Named Ranges
        case "create_named_range":
          if (!p.name || !p.address) {
            throw new Error("Missing required fields: name and address");
          }
          await ExcelService.createNamedRange(
            String(p.name),
            String(p.address),
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Named range "${p.name}" created`;

        case "delete_named_range":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.deleteNamedRange(String(p.name));
          return `Named range "${p.name}" deleted`;

        // Comments
        case "add_comment":
          if (!p.address || !p.text) {
            throw new Error("Missing required fields: address and text");
          }
          const comment = await CommentService.addComment({
            cellAddress: String(p.address),
            text: String(p.text),
            author: p.author ? String(p.author) : undefined,
            worksheetName: p.worksheetName ? String(p.worksheetName) : undefined
          });
          return `Comment added to ${p.address}`;

        case "delete_comment":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await CommentService.deleteComment(
            String(p.address), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Comment deleted from ${p.address}`;

        // Hyperlinks
        case "add_hyperlink":
          if (!p.address || !p.link) {
            throw new Error("Missing required fields: address and link");
          }
          await HyperlinkService.addHyperlink({
            cellAddress: String(p.address),
            url: String(p.link),
            displayText: p.displayText ? String(p.displayText) : undefined,
            screenTip: p.screenTip ? String(p.screenTip) : undefined,
            worksheetName: p.worksheetName ? String(p.worksheetName) : undefined
          });
          return `Hyperlink added to ${p.address}`;

        case "remove_hyperlink":
          if (!p.address) {
            throw new Error("Missing required field: address");
          }
          await HyperlinkService.removeHyperlink(
            String(p.address), 
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Hyperlink removed from ${p.address}`;

        // VBA Macro Operations
        case "create_vba_macro":
          if (!p.macroName || !p.description) {
            throw new Error("Missing required fields: macroName and description");
          }
          // VBA macros are generated by the AI and returned as text
          // The actual execution happens outside Office.js context
          logger.info("VBA Macro generation requested", { macroName: p.macroName, description: p.description });
          return `VBA macro "${p.macroName}" generated. The code is provided in the response.`;

        case "explain_vba_code":
          if (!p.vbaCode) {
            throw new Error("Missing required field: vbaCode");
          }
          // VBA explanation is generated by the AI and returned as text
          logger.info("VBA code explanation requested", { codeLength: String(p.vbaCode).length });
          return `VBA code explanation generated. The explanation is provided in the response.`;

        case "refactor_vba_code":
          if (!p.vbaCode) {
            throw new Error("Missing required field: vbaCode");
          }
          // VBA refactoring is generated by the AI and returned as text
          logger.info("VBA code refactoring requested", { codeLength: String(p.vbaCode).length, goals: p.goals });
          return `VBA code refactored. The improved code is provided in the response.`;

        // New worksheet operations
        case "duplicate_worksheet":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          const dupSheetName = await ExcelService.duplicateWorksheet(
            String(p.name),
            p.newName ? String(p.newName) : undefined
          );
          return `Worksheet "${p.name}" duplicated as "${dupSheetName}"`;

        case "move_worksheet":
          if (!p.name || p.position === undefined) {
            throw new Error("Missing required fields: name and position");
          }
          await ExcelService.moveWorksheet(String(p.name), Number(p.position));
          return `Worksheet "${p.name}" moved to position ${p.position}`;

        case "hide_worksheet":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.hideWorksheet(String(p.name));
          return `Worksheet "${p.name}" hidden`;

        case "show_worksheet":
          if (!p.name) {
            throw new Error("Missing required field: name");
          }
          await ExcelService.showWorksheet(String(p.name));
          return `Worksheet "${p.name}" shown`;

        // Row/Column operations
        case "hide_rows":
          if (p.startRow === undefined || p.endRow === undefined) {
            throw new Error("Missing required fields: startRow and endRow");
          }
          await ExcelService.hideRows(Number(p.startRow), Number(p.endRow), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Rows ${p.startRow} to ${p.endRow} hidden`;

        case "show_rows":
          if (p.startRow === undefined || p.endRow === undefined) {
            throw new Error("Missing required fields: startRow and endRow");
          }
          await ExcelService.showRows(Number(p.startRow), Number(p.endRow), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Rows ${p.startRow} to ${p.endRow} shown`;

        case "hide_columns":
          if (!p.startCol || !p.endCol) {
            throw new Error("Missing required fields: startCol and endCol");
          }
          await ExcelService.hideColumns(String(p.startCol), String(p.endCol), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Columns ${p.startCol} to ${p.endCol} hidden`;

        case "show_columns":
          if (!p.startCol || !p.endCol) {
            throw new Error("Missing required fields: startCol and endCol");
          }
          await ExcelService.showColumns(String(p.startCol), String(p.endCol), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Columns ${p.startCol} to ${p.endCol} shown`;

        case "freeze_panes":
          await ExcelService.freezePanes(
            p.freezeCell ? String(p.freezeCell) : undefined,
            p.freezeTopRow !== undefined ? Boolean(p.freezeTopRow) : undefined,
            p.freezeFirstColumn !== undefined ? Boolean(p.freezeFirstColumn) : undefined,
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return "Panes frozen";

        case "unfreeze_panes":
          await ExcelService.unfreezePanes(p.worksheetName ? String(p.worksheetName) : undefined);
          return "Panes unfrozen";

        case "insert_rows":
          if (p.rowIndex === undefined || p.count === undefined) {
            throw new Error("Missing required fields: rowIndex and count");
          }
          await ExcelService.insertRows(Number(p.rowIndex), Number(p.count), p.worksheetName ? String(p.worksheetName) : undefined);
          return `${p.count} rows inserted at row ${p.rowIndex}`;

        case "insert_columns":
          if (!p.columnIndex || p.count === undefined) {
            throw new Error("Missing required fields: columnIndex and count");
          }
          await ExcelService.insertColumns(String(p.columnIndex), Number(p.count), p.worksheetName ? String(p.worksheetName) : undefined);
          return `${p.count} columns inserted at column ${p.columnIndex}`;

        case "delete_rows":
          if (p.startRow === undefined || p.endRow === undefined) {
            throw new Error("Missing required fields: startRow and endRow");
          }
          await ExcelService.deleteRows(Number(p.startRow), Number(p.endRow), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Rows ${p.startRow} to ${p.endRow} deleted`;

        case "delete_columns":
          if (!p.startCol || !p.endCol) {
            throw new Error("Missing required fields: startCol and endCol");
          }
          await ExcelService.deleteColumns(String(p.startCol), String(p.endCol), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Columns ${p.startCol} to ${p.endCol} deleted`;

        case "set_row_height":
          if (p.rowIndex === undefined || p.height === undefined) {
            throw new Error("Missing required fields: rowIndex and height");
          }
          await ExcelService.setRowHeight(Number(p.rowIndex), Number(p.height), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Row ${p.rowIndex} height set to ${p.height}`;

        case "set_column_width":
          if (!p.columnIndex || p.width === undefined) {
            throw new Error("Missing required fields: columnIndex and width");
          }
          await ExcelService.setColumnWidth(String(p.columnIndex), Number(p.width), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Column ${p.columnIndex} width set to ${p.width}`;

        case "group_rows":
          if (p.startRow === undefined || p.endRow === undefined) {
            throw new Error("Missing required fields: startRow and endRow");
          }
          await ExcelService.groupRows(Number(p.startRow), Number(p.endRow), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Rows ${p.startRow} to ${p.endRow} grouped`;

        case "ungroup_rows":
          if (p.startRow === undefined || p.endRow === undefined) {
            throw new Error("Missing required fields: startRow and endRow");
          }
          await ExcelService.ungroupRows(Number(p.startRow), Number(p.endRow), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Rows ${p.startRow} to ${p.endRow} ungrouped`;

        case "group_columns":
          if (!p.startCol || !p.endCol) {
            throw new Error("Missing required fields: startCol and endCol");
          }
          await ExcelService.groupColumns(String(p.startCol), String(p.endCol), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Columns ${p.startCol} to ${p.endCol} grouped`;

        case "ungroup_columns":
          if (!p.startCol || !p.endCol) {
            throw new Error("Missing required fields: startCol and endCol");
          }
          await ExcelService.ungroupColumns(String(p.startCol), String(p.endCol), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Columns ${p.startCol} to ${p.endCol} ungrouped`;

        case "convert_table_to_range":
          if (!p.tableName) {
            throw new Error("Missing required field: tableName");
          }
          await ExcelService.convertTableToRange(String(p.tableName), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Table "${p.tableName}" converted to range`;

        case "convert_formulas_to_values":
          if (!p.range) {
            throw new Error("Missing required field: range");
          }
          await ExcelService.convertFormulasToValues(String(p.range), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Formulas in ${p.range} converted to values`;

        case "set_print_area":
          if (!p.range) {
            throw new Error("Missing required field: range");
          }
          await ExcelService.setPrintArea(String(p.range), p.worksheetName ? String(p.worksheetName) : undefined);
          return `Print area set to ${p.range}`;

        case "set_page_orientation":
          if (!p.orientation) {
            throw new Error("Missing required field: orientation");
          }
          await ExcelService.setPageOrientation(
            String(p.orientation) as 'portrait' | 'landscape',
            p.worksheetName ? String(p.worksheetName) : undefined
          );
          return `Page orientation set to ${p.orientation}`;

        default:
          throw new Error(`Unknown action type: ${type}`);
      }
    } catch (error) {
      logger.error(`Action execution failed: ${type}`, { payload: action.payload }, error instanceof Error ? error : new Error(String(error)));
      const errorMessage = error instanceof Error ? error.message : String(error);
      throw new Error(`Failed to execute ${type}: ${errorMessage}`);
    }
  }

  async executeActions(actions: AIAction[]): Promise<string[]> {
    const results: string[] = [];

    for (const action of actions) {
      try {
        const result = await this.executeAction(action);
        results.push(result);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        results.push(`Error: ${errorMessage}`);
      }
    }

    return results;
  }
}

/** Singleton instance of ActionHandler */
export const actionHandler = ActionHandler.getInstance();
export default actionHandler;
