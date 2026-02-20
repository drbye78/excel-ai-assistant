// Chart Service - Enhanced Chart Operations with Advanced Chart Types
// Phase 5 Implementation - Diagrams & Visualization Support

export type AdvancedChartType =
  | 'combo'           // Line + Column combination
  | 'waterfall'       // Bridge/flying bricks
  | 'funnel'          // Sales pipeline
  | 'map'             // Geographic
  | 'treemap'         // Hierarchical rectangles
  | 'sunburst'        // Multi-level pie
  | 'histogram'       // Frequency distribution
  | 'boxWhisker'      // Statistical distribution
  | 'sparkline';      // Mini charts in cells

export interface ComboChartConfig {
  primarySeries: {
    name: string;
    range: string;
    type: 'column' | 'line';
    axis: 'primary';
  }[];
  secondarySeries?: {
    name: string;
    range: string;
    type: 'line';
    axis: 'secondary';
  }[];
}

export interface WaterfallConfig {
  categories: string[];
  values: number[];
  totalLabel?: string;
  increaseColor?: string;
  decreaseColor?: string;
  totalColor?: string;
}

export interface FunnelConfig {
  stages: string[];
  values: number[];
  showPercentages?: boolean;
  gapWidth?: number;
}

export interface MapConfig {
  geographicData: string; // Column with country/region names
  valueData: string;      // Column with values
  mapType: 'world' | 'europe' | 'asia' | 'americas' | 'africa';
}

export interface TreemapConfig {
  categoryColumn: string;
  valueColumn: string;
  parentColumn?: string;  // For hierarchical
}

export interface SunburstConfig {
  levels: string[];       // Column names for each level
  valueColumn: string;
}

export interface HistogramConfig {
  dataColumn: string;
  binSize?: number;
  autoBin?: boolean;
}

export interface BoxWhiskerConfig {
  categories: string[];
  dataRanges: string[];
  showMean?: boolean;
  showOutliers?: boolean;
}

export interface SparklineConfig {
  dataRange: string;
  locationRange: string;
  type: 'line' | 'column' | 'winLoss';
  color?: string;
  negativeColor?: string;
  showMarkers?: boolean;
  lineWeight?: number;
}

export interface ChartFormatting {
  style?: number;              // Pre-built style index
  colorScheme?: string;        // Theme color or custom
  dataLabels?: {
    show?: boolean;
    position?: 'center' | 'insideEnd' | 'outsideEnd' | 'bestFit';
    showPercentages?: boolean;
    showValues?: boolean;
    separator?: string;
  };
  axis?: {
    primaryY?: {
      min?: number;
      max?: number;
      unit?: number;
      title?: string;
      displayUnits?: 'none' | 'hundreds' | 'thousands' | 'tenThousands' | 'hundredThousands' | 'millions';
    };
    secondaryY?: {
      min?: number;
      max?: number;
      unit?: number;
      title?: string;
    };
    x?: {
      title?: string;
      reverse?: boolean;
    };
  };
  trendline?: {
    type: 'linear' | 'exponential' | 'logarithmic' | 'polynomial' | 'movingAverage';
    order?: number;            // For polynomial
    period?: number;           // For moving average
    displayEquation?: boolean;
    displayR2?: boolean;
  };
  errorBars?: {
    type: 'fixedValue' | 'percentage' | 'standardDeviation' | 'standardError' | 'custom';
    value?: number;            // For fixed/percentage
    direction?: 'both' | 'minus' | 'plus';
    endStyle?: 'cap' | 'noCap';
  };
  legend?: {
    position?: 'bottom' | 'top' | 'left' | 'right' | 'corner';
    show?: boolean;
  };
  dataTable?: {
    show?: boolean;
    showLegendKeys?: boolean;
  };
}

export class ChartService {
  private static instance: ChartService;

  private constructor() {}

  static getInstance(): ChartService {
    if (!ChartService.instance) {
      ChartService.instance = new ChartService();
    }
    return ChartService.instance;
  }

  // ============================================================================
  // ADVANCED CHART CREATION
  // ============================================================================

  /**
   * Create a combo chart (line + column)
   */
  async createComboChart(
    config: ComboChartConfig,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Create a column chart as base
      const chart = worksheet.charts.add(
        Excel.ChartType.columnClustered,
        worksheet.getRange(config.primarySeries[0].range),
        'Auto'
      );

      // Configure primary series
      for (let i = 0; i < config.primarySeries.length; i++) {
        const series = chart.series.getItemAt(i);
        series.name = config.primarySeries[i].name;

        if (config.primarySeries[i].type === 'line') {
          series.chartType = Excel.ChartType.line;
        }
      }

      // Add secondary series if specified
      if (config.secondarySeries && config.secondarySeries.length > 0) {
        for (const secSeries of config.secondarySeries) {
          const series = chart.series.add(secSeries.name);
          series.setValues(worksheet.getRange(secSeries.range));
          series.chartType = Excel.ChartType.line;
          series.axisGroup = Excel.ChartAxisGroup.secondary;
        }
      }

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a waterfall chart
   */
  async createWaterfallChart(
    config: WaterfallConfig,
    location: string,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Waterfall chart type
      const chart = worksheet.charts.add(
        Excel.ChartType.waterfall,
        worksheet.getRange(location),
        'Auto'
      );

      // Apply colors if specified
      if (config.increaseColor) {
        // Note: Office.js may have limited waterfall customization
      }

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a funnel chart
   */
  async createFunnelChart(
    config: FunnelConfig,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      // Prepare data range from config
      const stages = config.stages || [];
      const values = config.values || [];
      const maxLength = Math.max(stages.length, values.length);
      
      // Create data range: stages in column A, values in column B
      const startRow = 1;
      const endRow = maxLength;
      const dataRange = worksheet.getRange(`A${startRow}:B${endRow}`);
      
      // Set stage labels and values
      const dataValues: any[][] = [];
      for (let i = 0; i < maxLength; i++) {
        dataValues.push([
          stages[i] || `Stage ${i + 1}`,
          values[i] || 0
        ]);
      }
      dataRange.values = dataValues;
      
      // Create funnel chart with the data range
      const chart = worksheet.charts.add(
        Excel.ChartType.funnel,
        dataRange,
        'Auto'
      );

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a treemap chart
   */
  async createTreemapChart(
    config: TreemapConfig,
    dataRange: string,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.add(
        Excel.ChartType.treemap,
        worksheet.getRange(dataRange),
        'Auto'
      );

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a sunburst chart
   */
  async createSunburstChart(
    config: SunburstConfig,
    dataRange: string,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.add(
        Excel.ChartType.sunburst,
        worksheet.getRange(dataRange),
        'Auto'
      );

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a histogram chart
   */
  async createHistogramChart(
    config: HistogramConfig,
    dataRange: string,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.add(
        Excel.ChartType.histogram,
        worksheet.getRange(dataRange),
        'Auto'
      );

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Create a box & whisker chart
   */
  async createBoxWhiskerChart(
    config: BoxWhiskerConfig,
    worksheetName?: string
  ): Promise<string> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.add(
        Excel.ChartType.boxwhisker,
        worksheet.getRange('A1'), // Placeholder
        'Auto'
      );

      chart.load('name');
      await context.sync();

      return chart.name;
    });
  }

  /**
   * Add sparklines to cells
   */
  async addSparklines(
    configs: SparklineConfig[],
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      for (const config of configs) {
        const dataRange = worksheet.getRange(config.dataRange);
        const locationRange = worksheet.getRange(config.locationRange);

        // Map sparkline type
        let sparklineType: Excel.SparklineType;
        switch (config.type) {
          case 'column':
            sparklineType = Excel.SparklineType.column;
            break;
          case 'winLoss':
            sparklineType = Excel.SparklineType.winLoss;
            break;
          case 'line':
          default:
            sparklineType = Excel.SparklineType.line;
        }

        // Add sparklines
        locationRange.insertSparklines(
          [{
            type: sparklineType,
            sourceData: dataRange,
            seriesColor: config.color || 'blue'
          }]
        );
      }

      await context.sync();
    });
  }

  // ============================================================================
  // CHART FORMATTING
  // ============================================================================

  /**
   * Apply formatting to a chart
   */
  async formatChart(
    chartName: string,
    formatting: ChartFormatting,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);

      // Apply style
      if (formatting.style) {
        // chart.style = formatting.style; // May not be available in all Office.js versions
      }

      // Configure data labels
      if (formatting.dataLabels) {
        chart.dataLabels.visible = formatting.dataLabels.show ?? false;
        if (formatting.dataLabels.showPercentages) {
          chart.dataLabels.showPercentage = true;
        }
        if (formatting.dataLabels.showValues) {
          chart.dataLabels.showValue = true;
        }
      }

      // Configure axes
      if (formatting.axis) {
        if (formatting.axis.primaryY) {
          const axis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.primary);
          if (formatting.axis.primaryY.min !== undefined) axis.minimum = formatting.axis.primaryY.min;
          if (formatting.axis.primaryY.max !== undefined) axis.maximum = formatting.axis.primaryY.max;
          if (formatting.axis.primaryY.title) axis.title.text = formatting.axis.primaryY.title;
        }

        if (formatting.axis.secondaryY) {
          const axis = chart.axes.getItem(Excel.ChartAxisType.value, Excel.ChartAxisGroup.secondary);
          if (formatting.axis.secondaryY.min !== undefined) axis.minimum = formatting.axis.secondaryY.min;
          if (formatting.axis.secondaryY.max !== undefined) axis.maximum = formatting.axis.secondaryY.max;
          if (formatting.axis.secondaryY.title) axis.title.text = formatting.axis.secondaryY.title;
        }

        if (formatting.axis.x?.title) {
          const axis = chart.axes.getItem(Excel.ChartAxisType.category, Excel.ChartAxisGroup.primary);
          axis.title.text = formatting.axis.x.title;
        }
      }

      // Configure legend
      if (formatting.legend) {
        chart.legend.visible = formatting.legend.show ?? true;
        if (formatting.legend.position) {
          const positionMap: Record<string, Excel.ChartLegendPosition> = {
            'bottom': Excel.ChartLegendPosition.bottom,
            'top': Excel.ChartLegendPosition.top,
            'left': Excel.ChartLegendPosition.left,
            'right': Excel.ChartLegendPosition.right,
            'corner': Excel.ChartLegendPosition.corner
          };
          if (positionMap[formatting.legend.position]) {
            chart.legend.position = positionMap[formatting.legend.position];
          }
        }
      }

      await context.sync();
    });
  }

  /**
   * Add a trendline to a chart series
   */
  async addTrendline(
    chartName: string,
    seriesIndex: number,
    trendlineConfig: ChartFormatting['trendline'],
    worksheetName?: string
  ): Promise<void> {
    if (!trendlineConfig) return;

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      const series = chart.series.getItemAt(seriesIndex);

      // Add trendline
      const trendline = series.trendlines.add(Excel.TrendlineType.linear);

      if (trendlineConfig.displayEquation) {
        trendline.displayEquation = true;
      }
      if (trendlineConfig.displayR2) {
        trendline.displayRSquared = true;
      }

      await context.sync();
    });
  }

  /**
   * Add error bars to a chart series
   */
  async addErrorBars(
    chartName: string,
    seriesIndex: number,
    errorBarConfig: ChartFormatting['errorBars'],
    worksheetName?: string
  ): Promise<void> {
    if (!errorBarConfig) return;

    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      chart.load('series');
      await context.sync();
      
      const series = chart.series.getItemAt(seriesIndex);
      
      // Office.js has limited error bar support
      // Attempt to set error bars if API is available
      try {
        if (errorBarConfig.type === 'fixed') {
          // Fixed value error bars
          if (series.format && 'errorBars' in series.format) {
            // Note: Error bars API may not be available in all Office.js versions
            // This is a best-effort implementation
            (series.format as any).errorBars = {
              include: errorBarConfig.direction === 'both' ? 'Both' : 
                      errorBarConfig.direction === 'plus' ? 'Plus' : 'Minus',
              type: 'FixedValue',
              amount: errorBarConfig.value || 0,
            };
          }
        } else if (errorBarConfig.type === 'percentage') {
          // Percentage error bars
          if (series.format && 'errorBars' in series.format) {
            (series.format as any).errorBars = {
              include: errorBarConfig.direction === 'both' ? 'Both' : 
                      errorBarConfig.direction === 'plus' ? 'Plus' : 'Minus',
              type: 'Percentage',
              amount: errorBarConfig.percentage || 5,
            };
          }
        }
      } catch (error) {
        // Error bars API not available in this Office.js version
        // Log warning but don't fail the operation
        logger.warn('Error bars not supported in this Office.js version', { error });
      }

      await context.sync();
    });
  }

  // ============================================================================
  // CHART MODIFICATION
  // ============================================================================

  /**
   * Add a data series to an existing chart
   */
  async addSeries(
    chartName: string,
    seriesName: string,
    dataRange: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      const series = chart.series.add(seriesName);
      series.setValues(worksheet.getRange(dataRange));

      await context.sync();
    });
  }

  /**
   * Remove a data series from a chart
   */
  async removeSeries(
    chartName: string,
    seriesName: string,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      const series = chart.series.getItem(seriesName);
      series.delete();

      await context.sync();
    });
  }

  /**
   * Change chart type
   */
  async changeChartType(
    chartName: string,
    newType: Excel.ChartType,
    worksheetName?: string
  ): Promise<void> {
    return Excel.run(async (context) => {
      const worksheet = worksheetName
        ? context.workbook.worksheets.getItem(worksheetName)
        : context.workbook.worksheets.getActiveWorksheet();

      const chart = worksheet.charts.getItem(chartName);
      chart.chartType = newType;

      await context.sync();
    });
  }

  /**
   * Delete a chart
   */
  async deleteChart(
    chartName: string,
    worksheetName?: string
  ): Promise<void> {
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
  // NATURAL LANGUAGE PARSING
  // ============================================================================

  /**
   * Parse natural language command for chart creation
   */
  parseNaturalLanguageCommand(command: string): {
    chartType?: string;
    dataRange?: string;
    title?: string;
    options?: Partial<ChartFormatting>;
  } {
    const result: {
      chartType?: string;
      dataRange?: string;
      title?: string;
      options?: Partial<ChartFormatting>;
    } = {};

    const lowerCmd = command.toLowerCase();

    // Detect chart type
    if (lowerCmd.includes('combo') || (lowerCmd.includes('column') && lowerCmd.includes('line'))) {
      result.chartType = 'combo';
    } else if (lowerCmd.includes('waterfall')) {
      result.chartType = 'waterfall';
    } else if (lowerCmd.includes('funnel')) {
      result.chartType = 'funnel';
    } else if (lowerCmd.includes('treemap')) {
      result.chartType = 'treemap';
    } else if (lowerCmd.includes('sunburst')) {
      result.chartType = 'sunburst';
    } else if (lowerCmd.includes('histogram')) {
      result.chartType = 'histogram';
    } else if (lowerCmd.includes('box') && lowerCmd.includes('whisker')) {
      result.chartType = 'boxWhisker';
    } else if (lowerCmd.includes('sparkline')) {
      result.chartType = 'sparkline';
    } else if (lowerCmd.includes('column')) {
      result.chartType = 'columnClustered';
    } else if (lowerCmd.includes('bar')) {
      result.chartType = 'barClustered';
    } else if (lowerCmd.includes('line')) {
      result.chartType = 'line';
    } else if (lowerCmd.includes('pie')) {
      result.chartType = 'pie';
    } else if (lowerCmd.includes('scatter')) {
      result.chartType = 'xyscatter';
    }

    // Extract data range
    const rangeMatch = command.match(/(?:from|of|using)\s+([A-Z]+\d+(?::[A-Z]+\d+)?)/i);
    if (rangeMatch) {
      result.dataRange = rangeMatch[1];
    }

    // Extract title
    const titleMatch = command.match(/(?:titled?|called|named)\s+["']([^"']+)["']/i) ||
                       command.match(/(?:titled?|called|named)\s+(\w+(?:\s+\w+)*)/i);
    if (titleMatch) {
      result.title = titleMatch[1];
    }

    // Detect formatting hints
    result.options = {};

    if (lowerCmd.includes('percentage') || lowerCmd.includes('%')) {
      result.options.dataLabels = { show: true, showPercentages: true };
    }

    if (lowerCmd.includes('legend')) {
      const posMatch = lowerCmd.match(/legend\s+(on\s+)?(bottom|top|left|right)/);
      result.options.legend = {
        show: true,
        position: (posMatch?.[2] as any) || 'bottom'
      };
    }

    if (lowerCmd.includes('trendline') || lowerCmd.includes('trend')) {
      result.options.trendline = { type: 'linear' };
    }

    return result;
  }

  /**
   * Recommend chart type based on data characteristics
   */
  recommendChartType(dataDescription: {
    hasCategories: boolean;
    hasTimeSeries: boolean;
    hasHierarchical: boolean;
    hasGeographic: boolean;
    valueCount: number;
    dataType: 'numeric' | 'percentage' | 'mixed';
  }): string {
    if (dataDescription.hasGeographic) {
      return 'map';
    }

    if (dataDescription.hasHierarchical) {
      return dataDescription.valueCount > 3 ? 'sunburst' : 'treemap';
    }

    if (dataDescription.hasTimeSeries) {
      if (dataDescription.dataType === 'percentage') {
        return 'line'; // Line chart good for percentage trends
      }
      return dataDescription.valueCount > 10 ? 'line' : 'columnClustered';
    }

    if (dataDescription.dataType === 'percentage') {
      return 'pie';
    }

    if (!dataDescription.hasCategories) {
      return 'histogram';
    }

    return 'columnClustered';
  }
}

export default ChartService.getInstance();
