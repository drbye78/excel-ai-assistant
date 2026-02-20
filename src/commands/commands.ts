// Commands for Ribbon Buttons
// Office.js API references for TypeScript
/// <reference types="@types/office-js" />
/// <reference types="@types/office-runtime" />

import { logger } from '../utils/logger';

// Declare global functions object for Office Add-in commands
declare const globals: {
  generateFormula: typeof generateFormula;
  createQuickChart: typeof createQuickChart;
  formatAsTable: typeof formatAsTable;
};

// Initialize Office Add-in commands
Office.onReady(() => {
  // Commands are ready
  logger.info('Office commands initialized');
});

/**
 * Shows a notification when the add-in command is executed
 */
function showNotification(text: string) {
  // This function can be used to show notifications to the user
  logger.info(text, { source: 'showNotification' });
}

/**
 * Open AI Assistant Task Pane
 */
function openTaskPane(event: Office.AddinCommands.Event) {
  // The task pane is opened automatically by the Office UI
  event.completed();
}

/**
 * Quick action: Generate formula from selection
 */
async function generateFormula(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      selection.load("address, values");
      await context.sync();

      // This would typically call your AI service
      showNotification(`Analyzing range ${selection.address}...`);
    });
  } catch (error) {
    logger.error('Failed to generate formula', {}, error as Error);
  } finally {
    event.completed();
  }
}

/**
 * Quick action: Create chart from selection
 */
async function createQuickChart(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      const worksheet = context.workbook.getActiveWorksheet();

      // Use Excel.ChartSeriesBy.auto enum instead of string
      const chart = worksheet.charts.add(
        Excel.ChartType.columnClustered,
        selection,
        Excel.ChartSeriesBy.auto
      );

      chart.title.text = "Quick Chart";
      await context.sync();

      showNotification("Chart created successfully!");
    });
  } catch (error) {
    logger.error('Failed to create quick chart', {}, error as Error);
  } finally {
    event.completed();
  }
}

/**
 * Quick action: Format as table
 */
async function formatAsTable(event: Office.AddinCommands.Event) {
  try {
    await Excel.run(async (context) => {
      const selection = context.workbook.getSelectedRange();
      const worksheet = context.workbook.getActiveWorksheet();

      // Use worksheet.tables.add() instead of worksheet.addTable()
      const table = worksheet.tables.add(selection, true);
      table.style = "TableStyleMedium2";
      await context.sync();

      showNotification("Table formatted successfully!");
    });
  } catch (error) {
    logger.error('Failed to format as table', {}, error as Error);
  } finally {
    event.completed();
  }
}

// Export for global access - assign after function declarations
globals.generateFormula = generateFormula;
globals.createQuickChart = createQuickChart;
globals.formatAsTable = formatAsTable;
