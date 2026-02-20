// Excel Service - Handles all Excel workbook operations
// Office.js API references for TypeScript
/// <reference types="@types/office-js" />
/// <reference types="@types/office-runtime" />

import {
  ExcelContext,
  WorkbookContext,
  WorksheetContext,
  RangeContext,
  TableContext,
  ChartContext,
  PivotTableContext,
  CellValidation,
  FormatOptions,
  ChartOptions,
  TableOptions,
  PivotTableOptions
} from "@/types";

export class ExcelService {
  private static instance: ExcelService;

  private constructor() {}

  static getInstance(): ExcelService {
    if (!ExcelService.instance) {
      ExcelService.instance = new ExcelService();
    }
    return ExcelService.instance;
  }

  // ============================================================================
  // CONTEXT GATHERING
  // ============================================================================

  async getFullContext(): Promise<ExcelContext> {
    return Excel.run(async (context) => {
      const workbook = context.workbook;
      const worksheets = workbook.worksheets;
      const activeWorksheet = workbook.worksheets.getActiveWorksheet();

      // Load all necessary properties
      worksheets.load("items/name, items/index");
      activeWorksheet.load("name, index");

      await context.sync();

      // Handle selection safely - it might not exist
      let selection: RangeContext | undefined;
      try {
        const selectedRange = workbook.getSelectedRange();
        if (selectedRange) {
          selectedRange.load([
            "address",
            "values",
            "formulas",
            "numberFormat",
            "rowCount",
            "columnCount",
            "worksheet/name"
          ]);

          await context.sync();

          // Safely access properties after sync - they could be undefined
          const address = selectedRange.address;
          const worksheetName = selectedRange.worksheet?.name;
          const values = selectedRange.values;
          const formulas = selectedRange.formulas;
          const numberFormat = selectedRange.numberFormat;
          const rowCount = selectedRange.rowCount;
          const columnCount = selectedRange.columnCount;

          // Only create selection context if we have valid data
          if (address && worksheetName && values) {
            selection = {
              address,
              worksheetName,
              values,
              formulas,
              numberFormat,
              rowCount,
              columnCount
            };
          }
        }
      } catch (selectionError) {
        // No selection or cannot get selection - this is fine
        console.debug("No range selected or cannot get selection:", selectionError);
      }

      // Get all tables across all worksheets
      const tables = await this.getAllTables(context);

      // Get all charts
      const charts = await this.getAllCharts(context);

      // Get all pivot tables
      const pivotTables = await this.getAllPivotTables(context);

      return {
        workbook: {
          name: "Workbook", // Excel API doesn't expose workbook name directly
          worksheets: worksheets.items.map((ws) => ws.name),
          namedRanges: await this.getNamedRanges(context)
        },
        selection, // Will be undefined if no selection
        activeWorksheet: {
          name: activeWorksheet.name,
          tables: tables
            .filter((t) => t.worksheetName === activeWorksheet.name)
            .map((t) => t.name),
          charts: charts
            .filter((c) => c.worksheetName === activeWorksheet.name)
            .map((c) => c.name)
        },
        tables,
        charts,
        pivotTables
      };
    });
  }

  // ============================================================================
  // WORKBOOK OPERATIONS
  // ============================================================================

  async getNamedRanges(context: Excel.RequestContext): Promise<string[]> {
    const names = context.workbook.names;
    names.load("items/name");
    await context.sync();
    return names.items.map((n) => n.name);
  }

  async addWorksheet(name?: string): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.add(name);
      worksheet.load("name");
      await context.sync();
      return worksheet.name;
    });
  }

  async deleteWorksheet(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(name);
      worksheet.delete();
      await context.sync();
    });
  }

  async renameWorksheet(oldName: string, newName: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(oldName);
      worksheet.name = newName;
      await context.sync();
    });
  }

  async duplicateWorksheet(sourceName: string, newName?: string): Promise<string> {
    return Excel.run(async (context) => {
      const sourceSheet = context.workbook.worksheets.getItem(sourceName);
      const newSheet = sourceSheet.copy(Excel.WorksheetPositionType.after, sourceSheet);
      if (newName) {
        newSheet.name = newName;
      } else {
        newSheet.name = `${sourceName} (Copy)`;
      }
      await context.sync();
      return newSheet.name;
    });
  }

  async moveWorksheet(worksheetName: string, position: number): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetName);
      const targetSheet = context.workbook.worksheets.getItemAt(position);
      worksheet.position = targetSheet.position;
      await context.sync();
    });
  }

  async hideWorksheet(worksheetName: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetName);
      worksheet.visibility = Excel.SheetVisibility.hidden;
      await context.sync();
    });
  }

  async showWorksheet(worksheetName: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(worksheetName);
      worksheet.visibility = Excel.SheetVisibility.visible;
      await context.sync();
    });
  }

  async hideRows(startRow: number, endRow: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const rows = worksheet.getRange(`${startRow}:${endRow}`);
      rows.format.visibility = false;
      await context.sync();
    });
  }

  async showRows(startRow: number, endRow: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const rows = worksheet.getRange(`${startRow}:${endRow}`);
      rows.format.visibility = true;
      await context.sync();
    });
  }

  async hideColumns(startCol: string, endCol: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const columns = worksheet.getRange(`${startCol}:${endCol}`);
      columns.format.visibility = false;
      await context.sync();
    });
  }

  async showColumns(startCol: string, endCol: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const columns = worksheet.getRange(`${startCol}:${endCol}`);
      columns.format.visibility = true;
      await context.sync();
    });
  }

  async freezePanes(freezeCell?: string, freezeTopRow?: boolean, freezeFirstColumn?: boolean, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      
      if (freezeCell) {
        const range = worksheet.getRange(freezeCell);
        worksheet.freezePanes.freezeAt(range);
      } else if (freezeTopRow) {
        const topRow = worksheet.getRange('1:1');
        worksheet.freezePanes.freezeAt(topRow);
      } else if (freezeFirstColumn) {
        const firstCol = worksheet.getRange('A:A');
        worksheet.freezePanes.freezeAt(firstCol);
      }
      await context.sync();
    });
  }

  async unfreezePanes(worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      worksheet.freezePanes.unfreeze();
      await context.sync();
    });
  }

  async insertRows(rowIndex: number, count: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${rowIndex}:${rowIndex + count - 1}`);
      range.insert(Excel.InsertShiftDirection.down);
      await context.sync();
    });
  }

  async insertColumns(columnIndex: string, count: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${columnIndex}:${this.getColumnAfter(columnIndex, count - 1)}`);
      range.insert(Excel.InsertShiftDirection.right);
      await context.sync();
    });
  }

  async deleteRows(startRow: number, endRow: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startRow}:${endRow}`);
      range.delete(Excel.DeleteShiftDirection.up);
      await context.sync();
    });
  }

  async deleteColumns(startCol: string, endCol: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startCol}:${endCol}`);
      range.delete(Excel.DeleteShiftDirection.left);
      await context.sync();
    });
  }

  async setRowHeight(rowIndex: number, height: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const row = worksheet.getRange(`${rowIndex}:${rowIndex}`);
      row.format.rowHeight = height;
      await context.sync();
    });
  }

  async setColumnWidth(columnIndex: string, width: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const column = worksheet.getRange(`${columnIndex}:${columnIndex}`);
      column.format.columnWidth = width;
      await context.sync();
    });
  }

  async groupRows(startRow: number, endRow: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startRow}:${endRow}`);
      range.group();
      await context.sync();
    });
  }

  async ungroupRows(startRow: number, endRow: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startRow}:${endRow}`);
      range.ungroup();
      await context.sync();
    });
  }

  async groupColumns(startCol: string, endCol: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startCol}:${endCol}`);
      range.group();
      await context.sync();
    });
  }

  async ungroupColumns(startCol: string, endCol: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(`${startCol}:${endCol}`);
      range.ungroup();
      await context.sync();
    });
  }

  async convertTableToRange(tableName: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const table = worksheet.tables.getItem(tableName);
      table.convertToRange();
      await context.sync();
    });
  }

  async convertFormulasToValues(range: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const rangeObj = worksheet.getRange(range);
      rangeObj.load(['values', 'formulas']);
      await context.sync();
      
      // Copy values to replace formulas
      const values = rangeObj.values;
      rangeObj.values = values;
      await context.sync();
    });
  }

  async setPrintArea(range: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      worksheet.pageLayout.printArea = range;
      await context.sync();
    });
  }

  async setPageOrientation(orientation: 'portrait' | 'landscape', worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      worksheet.pageLayout.orientation = orientation === 'landscape' 
        ? Excel.PageOrientation.landscape 
        : Excel.PageOrientation.portrait;
      await context.sync();
    });
  }

  private getColumnAfter(column: string, offset: number): string {
    let colNum = 0;
    for (let i = 0; i < column.length; i++) {
      colNum = colNum * 26 + (column.charCodeAt(i) - 64);
    }
    colNum += offset;
    let result = '';
    while (colNum > 0) {
      colNum--;
      result = String.fromCharCode(65 + (colNum % 26)) + result;
      colNum = Math.floor(colNum / 26);
    }
    return result;
  }

  // ============================================================================
  // CELL & RANGE OPERATIONS
  // ============================================================================

  async getRange(address: string, worksheetName?: string): Promise<RangeContext> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.load([
        "address",
        "values",
        "formulas",
        "numberFormat",
        "rowCount",
        "columnCount",
        "worksheet/name"
      ]);

      await context.sync();

      return {
        address: range.address,
        worksheetName: range.worksheet.name,
        values: range.values,
        formulas: range.formulas,
        numberFormat: range.numberFormat,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };
    });
  }

  async setValues(address: string, values: any[][], worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.values = values;
      await context.sync();
    });
  }

  async setFormulas(address: string, formulas: string[][], worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.formulas = formulas;
      await context.sync();
    });
  }

  async clearRange(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.clear();
      await context.sync();
    });
  }

  async copyRange(sourceAddress: string, destAddress: string, sourceWorksheet?: string, destWorksheet?: string): Promise<void> {
    return Excel.run(async (context) => {
      const sourceWs = sourceWorksheet
        ? context.workbook.worksheets.getItem(sourceWorksheet)
        : context.workbook.worksheets.getActiveWorksheet();

      const destWs = destWorksheet
        ? context.workbook.worksheets.getItem(destWorksheet)
        : context.workbook.worksheets.getActiveWorksheet();

      const sourceRange = sourceWs.getRange(sourceAddress);
      const destRange = destWs.getRange(destAddress);

      destRange.copyFrom(sourceRange, Excel.RangeCopyType.all);
      await context.sync();
    });
  }

  // ============================================================================
  // TABLE OPERATIONS
  // ============================================================================

  private async getAllTables(context: Excel.RequestContext): Promise<TableContext[]> {
    const tables: TableContext[] = [];
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    for (const worksheet of worksheets.items) {
      const worksheetTables = worksheet.tables;
      worksheetTables.load("items/name");
      await context.sync();

      for (const table of worksheetTables.items) {
        // Load table properties properly - note: Table doesn't have rowCount in all Office.js versions
        // Use the rows collection count instead
        table.load("name");
        table.rows.load("count");
        table.columns.load("count");
        
        const tableRange = table.getRange();
        tableRange.load("address");
        
        const headerRange = table.getHeaderRowRange();
        headerRange.load("values, address");
        
        const dataRange = table.getDataBodyRange();
        dataRange.load("address");
        
        await context.sync();

        tables.push({
          name: table.name,
          worksheetName: worksheet.name,
          range: tableRange.address,
          headerRowRange: headerRange.address,
          dataBodyRange: dataRange.address,
          headers: headerRange.values[0] || [],
          rowCount: table.rows.count, // Use rows.count instead of rowCount
          columnCount: table.columns.count
        });
      }
    }

    return tables;
  }

  async createTable(options: TableOptions, rangeAddress: string, worksheetName?: string): Promise<TableContext> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Load worksheet name first
      worksheet.load("name");

      const range = worksheet.getRange(rangeAddress);
      // Use the correct Excel JS API: worksheet.tables.add()
      const table = worksheet.tables.add(range, options.hasHeaders);

      if (options.name) {
        table.name = options.name;
      }

      if (options.style) {
        table.style = options.style;
      }

      if (options.showTotals) {
        table.showTotals = true;
        if (options.totalColumns) {
          for (const colName of options.totalColumns) {
            const column = table.columns.getItem(colName);
            column.getTotalRowRange().values = [["=SUBTOTAL(109,[" + colName + "])"]];
          }
        }
      }

      // Load table primitive properties - Table object doesn't have rowCount in all Office.js versions
      table.load("name");
      
      // Load the rows and columns count separately
      table.rows.load("count");
      table.columns.load("count");

      // Get and load the range object separately (it's an object, not a primitive)
      const tableRange = table.getRange();
      tableRange.load("address, rowCount");
      
      const headerRange = table.getHeaderRowRange();
      headerRange.load("values, address");
      
      const dataRange = table.getDataBodyRange();
      dataRange.load("address");

      await context.sync();

      return {
        name: table.name,
        worksheetName: worksheet.name,
        range: tableRange.address,
        headerRowRange: headerRange.address,
        dataBodyRange: dataRange.address,
        headers: headerRange.values[0] || [],
        rowCount: tableRange.rowCount, // Use tableRange.rowCount (loaded above)
        columnCount: table.columns.count
      };
    });
  }

  async addTableRow(tableName: string, values: any[], worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const table = worksheet.tables.getItem(tableName);
      // Use -1 to add at the end (null is not valid, use undefined or -1)
      table.rows.add(undefined, [values]);
      await context.sync();
    });
  }

  async deleteTable(tableName: string, keepData: boolean = true, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const table = worksheet.tables.getItem(tableName);
      table.delete();
      await context.sync();
    });
  }

  async getTableData(tableName: string, worksheetName?: string): Promise<any[][]> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const table = worksheet.tables.getItem(tableName);
      const dataRange = table.getDataBodyRange();
      dataRange.load("values");
      await context.sync();

      return dataRange.values;
    });
  }

  // ============================================================================
  // CHART OPERATIONS
  // ============================================================================

  private async getAllCharts(context: Excel.RequestContext): Promise<ChartContext[]> {
    const charts: ChartContext[] = [];
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    for (const worksheet of worksheets.items) {
      const worksheetCharts = worksheet.charts;
      worksheetCharts.load("items/name");
      await context.sync();

      for (const chart of worksheetCharts.items) {
        // chartType is the correct property name in Office.js
        chart.load("name, chartType");
        await context.sync();
        
        charts.push({
          name: chart.name,
          type: String(chart.chartType),
          worksheetName: worksheet.name
        });
      }
    }

    return charts;
  }

  async createChart(dataRange: string, options: ChartOptions, worksheetName?: string): Promise<ChartContext> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(dataRange);

      // Map string type to Excel.ChartType
      const chartTypeMap: { [key: string]: Excel.ChartType } = {
        columnClustered: Excel.ChartType.columnClustered,
        columnStacked: Excel.ChartType.columnStacked,
        barClustered: Excel.ChartType.barClustered,
        line: Excel.ChartType.line,
        lineMarkers: Excel.ChartType.lineMarkers,
        pie: Excel.ChartType.pie,
        doughnut: Excel.ChartType.doughnut,
        scatter: Excel.ChartType.xyscatter,
        area: Excel.ChartType.area,
        radar: Excel.ChartType.radar
      };

      const chartType = chartTypeMap[options.type] || Excel.ChartType.columnClustered;
      // Use Excel.ChartSeriesBy.auto enum instead of string literal
      // See: https://learn.microsoft.com/en-us/javascript/api/excel/excel.chartseriesby
      const chart = worksheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);

      if (options.title) {
        chart.title.text = options.title;
      }

      if (options.legendPosition) {
        chart.legend.position = options.legendPosition as Excel.ChartLegendPosition;
      }

      // Note: chart.dataLabels.visible is not available in all Office.js versions
      // Use showValue or other properties based on the specific Office.js version
      if (options.dataLabels) {
        try {
          chart.dataLabels.showValue = true;
        } catch {
          // Data labels may not be available in all chart types
        }
      }

      // Load chart properties - use chartType instead of type
      chart.load("name, chartType");
      await context.sync();

      return {
        name: chart.name,
        type: String(chart.chartType),
        worksheetName: worksheet.name
      };
    });
  }

  async deleteChart(chartName: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      chart.delete();
      await context.sync();
    });
  }

  // ============================================================================
  // PIVOT TABLE OPERATIONS
  // ============================================================================

  private async getAllPivotTables(context: Excel.RequestContext): Promise<PivotTableContext[]> {
    const pivotTables: PivotTableContext[] = [];
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    for (const worksheet of worksheets.items) {
      const worksheetPivots = worksheet.pivotTables;
      worksheetPivots.load("items/name");
      await context.sync();

      for (const pivot of worksheetPivots.items) {
        pivotTables.push({
          name: pivot.name,
          worksheetName: worksheet.name
        });
      }
    }

    return pivotTables;
  }

  async createPivotTable(options: PivotTableOptions, worksheetName?: string): Promise<PivotTableContext> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const sourceRange = worksheet.getRange(options.sourceData);

      // Create pivot table - destination is required, default to cell next to source
      const destRange = options.destination
        ? worksheet.getRange(options.destination)
        : sourceRange.getCell(0, sourceRange.columnCount + 2);

      const pivotTable = worksheet.pivotTables.add(
        options.name,
        sourceRange,
        destRange
      );

      // Add fields
      if (options.rowFields) {
        for (const field of options.rowFields) {
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
        }
      }

      if (options.columnFields) {
        for (const field of options.columnFields) {
          pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem(field));
        }
      }

      if (options.dataFields) {
        for (const field of options.dataFields) {
          const dataHierarchy = pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(field.name));

          // Set aggregation function using summarizeBy property
          // Office.js uses specific string values (case-sensitive)
          // See: https://learn.microsoft.com/en-us/javascript/api/excel/excel.datapivothierarchy
          const functionMap: Record<string, string> = {
            sum: 'Sum',
            count: 'Count',
            average: 'Average',
            max: 'Max',
            min: 'Min',
            product: 'Product',
            countNumbers: 'CountNumbers',
            stdDev: 'StandardDeviation',
            stdDevP: 'StandardDeviationP',
            var: 'Variance',
            varP: 'VarianceP'
          };

          const dataFunction = functionMap[field.function] || 'Sum';
          dataHierarchy.summarizeBy = dataFunction as any; // Type assertion due to Office.js type limitations
        }
      }

      if (options.filterFields) {
        for (const field of options.filterFields) {
          pivotTable.filterHierarchies.add(pivotTable.hierarchies.getItem(field));
        }
      }

      pivotTable.load("name");
      await context.sync();

      return {
        name: pivotTable.name,
        worksheetName: worksheet.name,
        sourceData: options.sourceData
      };
    });
  }

  async refreshPivotTable(name: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const pivotTable = worksheet.pivotTables.getItem(name);
      pivotTable.refresh();
      await context.sync();
    });
  }

  // ============================================================================
  // CELL VALIDATION
  // ============================================================================

  /**
   * Build a DataValidationRule object based on validation type
   * Note: Excel.DataValidationRule is a discriminated union type, so we build it dynamically
   * See: https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationrule
   */
  private buildValidationRule(validation: CellValidation): Excel.DataValidationRule {
    // Start with empty rule - will be populated based on type
    const rule: Partial<Excel.DataValidationRule> = {};

    // Set the validation type - this determines which rule type to use
    switch (validation.type) {
      case 'list':
        // For list, source can be a comma-separated string or a range formula
        // If formula1 looks like a range reference (e.g., "=$A$1:$A$10"), use it directly
        // Otherwise, treat it as comma-separated values and convert to formula
        const listSource = validation.formula1 || '';
        (rule as Excel.DataValidationRule).list = {
          inCellDropDown: true,
          source: listSource
        };
        break;

      case 'whole':
        // wholeNumber requires numeric formula1/formula2 and operator
        (rule as Excel.DataValidationRule).wholeNumber = {
          formula1: Number(validation.formula1 || 0),
          operator: validation.formula2
            ? Excel.DataValidationOperator.between
            : Excel.DataValidationOperator.greaterThanOrEqualTo
        };
        // Add formula2 only if provided (for between operator)
        if (validation.formula2) {
          (rule as Excel.DataValidationRule).wholeNumber!.formula2 = Number(validation.formula2);
        }
        break;

      case 'decimal':
        (rule as Excel.DataValidationRule).decimal = {
          formula1: Number(validation.formula1 || 0),
          operator: validation.formula2
            ? Excel.DataValidationOperator.between
            : Excel.DataValidationOperator.greaterThanOrEqualTo
        };
        if (validation.formula2) {
          (rule as Excel.DataValidationRule).decimal!.formula2 = Number(validation.formula2);
        }
        break;

      case 'date':
        // Date validation uses Date objects or ISO date strings
        (rule as Excel.DataValidationRule).date = {
          formula1: validation.formula1 || '',
          operator: validation.formula2
            ? Excel.DataValidationOperator.between
            : Excel.DataValidationOperator.greaterThanOrEqualTo
        };
        if (validation.formula2) {
          (rule as Excel.DataValidationRule).date!.formula2 = validation.formula2;
        }
        break;

      case 'textLength':
        (rule as Excel.DataValidationRule).textLength = {
          formula1: Number(validation.formula1 || 0),
          operator: validation.formula2
            ? Excel.DataValidationOperator.between
            : Excel.DataValidationOperator.greaterThanOrEqualTo
        };
        if (validation.formula2) {
          (rule as Excel.DataValidationRule).textLength!.formula2 = Number(validation.formula2);
        }
        break;

      case 'custom':
        // Custom uses a formula string
        (rule as Excel.DataValidationRule).custom = {
          formula: validation.formula1 || ''
        };
        break;
    }

    return rule as Excel.DataValidationRule;
  }

  async addValidation(address: string, validation: CellValidation, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      const dataValidation = range.dataValidation;

      // Clear existing validation first
      dataValidation.clear();

      // Build the validation rule using the helper method
      const rule = this.buildValidationRule(validation);

      // Assign the rule to data validation
      dataValidation.rule = rule;

      // Set ignore blanks option
      if (validation.allowBlank !== undefined) {
        dataValidation.ignoreBlanks = validation.allowBlank;
      }

      // Set input prompt (shown when cell is selected)
      if (validation.showInputMessage && validation.inputMessage) {
        dataValidation.prompt = {
          message: validation.inputMessage,
          showPrompt: true,
          title: validation.inputTitle || "Input"
        };
      }

      // Set error alert (shown when invalid data is entered)
      if (validation.showErrorMessage && validation.errorMessage) {
        dataValidation.errorAlert = {
          message: validation.errorMessage,
          showAlert: true,
          style: (validation.errorStyle as Excel.DataValidationAlertStyle) || Excel.DataValidationAlertStyle.stop,
          title: validation.errorTitle || "Error"
        };
      }

      await context.sync();
    });
  }

  async clearValidation(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.dataValidation.clear();
      await context.sync();
    });
  }

  // ============================================================================
  // FORMATTING
  // ============================================================================

  async formatRange(address: string, options: FormatOptions, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);

      if (options.numberFormat) {
        range.numberFormat = [[options.numberFormat]];
      }

      if (options.font) {
        if (options.font.name) range.format.font.name = options.font.name;
        if (options.font.size) range.format.font.size = options.font.size;
        if (options.font.bold !== undefined) range.format.font.bold = options.font.bold;
        if (options.font.italic !== undefined) range.format.font.italic = options.font.italic;
        if (options.font.color) range.format.font.color = options.font.color;
      }

      if (options.fill) {
        if (options.fill.color) {
          range.format.fill.pattern = Excel.FillPattern.solid;
          range.format.fill.color = options.fill.color;
        }
      }

      if (options.alignment) {
        if (options.alignment.horizontal) range.format.horizontalAlignment = options.alignment.horizontal as Excel.HorizontalAlignment;
        if (options.alignment.vertical) range.format.verticalAlignment = options.alignment.vertical as Excel.VerticalAlignment;
        if (options.alignment.wrapText !== undefined) range.format.wrapText = options.alignment.wrapText;
      }

      await context.sync();
    });
  }

  async autoFitColumns(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.autofitColumns();
      await context.sync();
    });
  }

  async autoFitRows(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.autofitRows();
      await context.sync();
    });
  }

  // ============================================================================
  // NAMED RANGES
  // ============================================================================

  async createNamedRange(name: string, address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      context.workbook.names.add(name, range);
      await context.sync();
    });
  }

  async deleteNamedRange(name: string): Promise<void> {
    return Excel.run(async (context) => {
      const namedItem = context.workbook.names.getItem(name);
      namedItem.delete();
      await context.sync();
    });
  }

  // ============================================================================
  // UTILITY
  // ============================================================================

  async calculateWorksheet(worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Calculate the used range of the worksheet
      const usedRange = worksheet.getUsedRange();
      usedRange.calculate();
      await context.sync();
    });
  }

  async calculateFull(): Promise<void> {
    return Excel.run(async (context) => {
      context.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
  }

  async getUsedRange(worksheetName?: string): Promise<RangeContext> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getUsedRange();
      range.load([
        "address",
        "values",
        "formulas",
        "rowCount",
        "columnCount",
        "worksheet/name"
      ]);

      await context.sync();

      return {
        address: range.address,
        worksheetName: range.worksheet.name,
        values: range.values,
        formulas: range.formulas,
        rowCount: range.rowCount,
        columnCount: range.columnCount
      };
    });
  }

  // ============================================================================
  // CELL MERGING & ALIGNMENT OPERATIONS (Phase 1.2)
  // ============================================================================

  async mergeCells(address: string, mergeType: 'cells' | 'across' = 'cells', worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);

      if (mergeType === 'across') {
        range.merge(true); // merge across (row by row)
      } else {
        range.merge(false); // merge cells (full range)
      }

      await context.sync();
    });
  }

  async unmergeCells(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.unmerge();
      await context.sync();
    });
  }

  async setTextWrap(address: string, wrap: boolean, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.wrapText = wrap;
      await context.sync();
    });
  }

  async setTextOrientation(address: string, degrees: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.textOrientation = degrees;
      await context.sync();
    });
  }

  async setIndentation(address: string, level: number, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.indentLevel = level;
      await context.sync();
    });
  }

  async setShrinkToFit(address: string, shrink: boolean, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.shrinkToFit = shrink;
      await context.sync();
    });
  }

  // ============================================================================
  // BORDER OPERATIONS (Phase 1.3)
  // ============================================================================

  async setBorders(
    address: string,
    borders: {
      edge?: 'edgeTop' | 'edgeBottom' | 'edgeLeft' | 'edgeRight' | 'edgeHorizontal' | 'edgeVertical' | 'insideVertical' | 'insideHorizontal' | 'outline' | 'all';
      style?: 'none' | 'continuous' | 'dash' | 'dashDot' | 'dashDotDot' | 'dot' | 'double' | 'slantDashDot';
      color?: string;
      weight?: 'hairline' | 'thin' | 'medium' | 'thick';
    }[],
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);

      for (const border of borders) {
        const edge = border.edge || 'all';
        const borderObj = range.format.borders.getItem(edge as Excel.BorderIndex);

        if (border.style) {
          borderObj.style = border.style as Excel.BorderLineStyle;
        }
        if (border.color) {
          borderObj.color = border.color;
        }
        if (border.weight) {
          borderObj.weight = border.weight as Excel.BorderWeight;
        }
      }

      await context.sync();
    });
  }

  async clearBorders(address: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.borders.load('items');
      await context.sync();

      for (const border of range.format.borders.items) {
        border.style = Excel.BorderLineStyle.none;
      }

      await context.sync();
    });
  }

  // ============================================================================
  // CELL & WORKSHEET PROTECTION (Phase 1.3)
  // ============================================================================

  async setCellLocked(address: string, locked: boolean, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      range.format.protection.locked = locked;
      await context.sync();
    });
  }

  async setCellHidden(address: string, hidden: boolean, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const range = worksheet.getRange(address);
      // Use the 'formulaHidden' property to hide formulas (shows blank in formula bar when sheet is protected)
      // Office.js API: RangeFormat.protection.formulaHidden
      range.format.protection.formulaHidden = hidden;
      await context.sync();
    });
  }

  async protectWorksheet(
    options: {
      password?: string;
      allowFormatCells?: boolean;
      allowFormatColumns?: boolean;
      allowFormatRows?: boolean;
      allowInsertColumns?: boolean;
      allowInsertRows?: boolean;
      allowInsertHyperlinks?: boolean;
      allowDeleteColumns?: boolean;
      allowDeleteRows?: boolean;
      allowSort?: boolean;
      allowAutoFilter?: boolean;
      allowPivotTables?: boolean;
    },
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Build protection options with supported Office.js properties
      // See: https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheetprotectionoptions
      const protectionOptions: Excel.WorksheetProtectionOptions = {
        allowFormatCells: options.allowFormatCells ?? false,
        allowFormatColumns: options.allowFormatColumns ?? false,
        allowFormatRows: options.allowFormatRows ?? false,
        allowInsertColumns: options.allowInsertColumns ?? false,
        allowInsertRows: options.allowInsertRows ?? false,
        allowInsertHyperlinks: options.allowInsertHyperlinks ?? false,
        allowDeleteColumns: options.allowDeleteColumns ?? false,
        allowDeleteRows: options.allowDeleteRows ?? false,
        allowSort: options.allowSort ?? false,
        allowAutoFilter: options.allowAutoFilter ?? false,
        allowPivotTables: options.allowPivotTables ?? false
      };

      // Use the options argument for protection
      worksheet.protection.protect(protectionOptions, options.password || undefined);
      await context.sync();
    });
  }

  async unprotectWorksheet(password?: string, worksheetName?: string): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      worksheet.protection.unprotect(password || undefined);
      await context.sync();
    });
  }

  async isWorksheetProtected(worksheetName?: string): Promise<boolean> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      worksheet.protection.load('protected');
      await context.sync();

      return worksheet.protection.protected;
    });
  }
}

export default ExcelService.getInstance();
