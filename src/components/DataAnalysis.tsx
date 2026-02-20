/**
 * Excel AI Assistant - Data Analysis Component
 * Interactive UI for statistical data analysis
 * 
 * @module components/DataAnalysis
 */

import React, { useState, useEffect, useCallback } from "react";
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Pivot,
  PivotItem,
  Icon,
  TooltipHost,
  Separator,
  Toggle,
  Slider
} from "@fluentui/react";
import { RangeContext } from "../types";
import {
  analyzeRange,
  calculateDescriptiveStats,
  detectAnomalies,
  analyzeTrend,
  calculateCorrelationMatrix,
  generateDataQualityReport,
  formatAnalysisResults,
  exportStatsToTable,
  exportCorrelationToTable,
  CompleteAnalysis,
  DescriptiveStats,
  Anomaly,
  TrendAnalysis,
  CorrelationMatrix,
  DataColumn
} from "../services/dataAnalysis";

// ============================================================================
// Types
// ============================================================================

interface DataAnalysisProps {
  selectedRange?: RangeContext;
  onApplyHighlighting?: (anomalies: { row: number; col: number; severity: string }[]) => void;
  onExportResults?: (data: (string | number)[][], sheetName: string) => void;
}

type AnalysisTab = "overview" | "statistics" | "anomalies" | "trends" | "correlations" | "quality";

// ============================================================================
// Component
// ============================================================================

export const DataAnalysis: React.FC<DataAnalysisProps> = ({
  selectedRange,
  onApplyHighlighting,
  onExportResults
}) => {
  // State
  const [activeTab, setActiveTab] = useState<AnalysisTab>("overview");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analysis, setAnalysis] = useState<CompleteAnalysis | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [zScoreThreshold, setZScoreThreshold] = useState(3);
  const [useIQR, setUseIQR] = useState(true);

  // Perform analysis when range changes
  const performAnalysis = useCallback(async () => {
    if (!selectedRange) return;

    setIsAnalyzing(true);
    setError(null);

    try {
      // Add a small delay to show the progress indicator
      await new Promise(resolve => setTimeout(resolve, 300));
      
      const result = analyzeRange(selectedRange);
      setAnalysis(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : "An error occurred during analysis");
    } finally {
      setIsAnalyzing(false);
    }
  }, [selectedRange]);

  // Auto-analyze when range is provided
  useEffect(() => {
    if (selectedRange) {
      performAnalysis();
    }
  }, [selectedRange, performAnalysis]);

  // ============================================================================
  // Render Helpers
  // ============================================================================

  const renderOverview = () => {
    if (!analysis) return null;

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
          <StatCard
            icon="AnalyticsView"
            title="Columns Analyzed"
            value={analysis.columnCount.toString()}
          />
          <StatCard
            icon="RowsChild"
            title="Data Rows"
            value={(analysis.rowCount - 1).toString()}
          />
          <StatCard
            icon="AlertSettings"
            title="Anomalies Found"
            value={analysis.anomalies.length.toString()}
            color={analysis.anomalies.length > 0 ? "#d13438" : undefined}
          />
          <StatCard
            icon="Trending12"
            title="Trends Detected"
            value={analysis.trends?.size.toString() || "0"}
          />
        </Stack>

        <Separator />

        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          Summary
        </Text>
        <Text>{analysis.summary}</Text>

        {analysis.insights.length > 0 && (
          <>
            <Separator />
            <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
              Key Insights
            </Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {analysis.insights.map((insight, index) => (
                <MessageBar
                  key={index}
                  messageBarType={MessageBarType.info}
                  styles={{ root: { borderRadius: 4 } }}
                >
                  {insight}
                </MessageBar>
              ))}
            </Stack>
          </>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            iconProps={{ iconName: "Highlight" }}
            text="Highlight Anomalies"
            disabled={analysis.anomalies.length === 0 || !onApplyHighlighting}
            onClick={() => {
              const highlights = analysis.anomalies.map(a => ({
                row: a.rowIndex,
                col: a.columnIndex,
                severity: a.severity
              }));
              onApplyHighlighting?.(highlights);
            }}
          />
          <DefaultButton
            iconProps={{ iconName: "Download" }}
            text="Export Statistics"
            disabled={analysis.descriptiveStats.size === 0 || !onExportResults}
            onClick={() => {
              const data = exportStatsToTable(analysis.descriptiveStats);
              onExportResults?.(data, "Descriptive Statistics");
            }}
          />
          {analysis.correlations && (
            <DefaultButton
              iconProps={{ iconName: "Download" }}
              text="Export Correlations"
              onClick={() => {
                const data = exportCorrelationToTable(analysis.correlations!);
                onExportResults?.(data, "Correlation Matrix");
              }}
            />
          )}
        </Stack>
      </Stack>
    );
  };

  const renderStatistics = () => {
    if (!analysis || analysis.descriptiveStats.size === 0) {
      return <Text>No numeric data available for statistical analysis.</Text>;
    }

    const columns: IColumn[] = [
      { key: "column", name: "Column", fieldName: "column", minWidth: 100, maxWidth: 150 },
      { key: "count", name: "Count", fieldName: "count", minWidth: 60, maxWidth: 80 },
      { key: "mean", name: "Mean", fieldName: "mean", minWidth: 80, maxWidth: 100 },
      { key: "median", name: "Median", fieldName: "median", minWidth: 80, maxWidth: 100 },
      { key: "stdDev", name: "Std Dev", fieldName: "stdDev", minWidth: 80, maxWidth: 100 },
      { key: "min", name: "Min", fieldName: "min", minWidth: 80, maxWidth: 100 },
      { key: "max", name: "Max", fieldName: "max", minWidth: 80, maxWidth: 100 },
      { key: "cv", name: "CV %", fieldName: "cv", minWidth: 70, maxWidth: 90 },
      { key: "skewness", name: "Skewness", fieldName: "skewness", minWidth: 80, maxWidth: 100 }
    ];

    const items = Array.from(analysis.descriptiveStats.entries()).map(([name, stats]) => ({
      key: name,
      column: name,
      count: stats.count.toLocaleString(),
      mean: stats.mean.toFixed(2),
      median: stats.median.toFixed(2),
      stdDev: stats.stdDev.toFixed(2),
      min: stats.min.toFixed(2),
      max: stats.max.toFixed(2),
      cv: stats.cv.toFixed(1),
      skewness: stats.skewness.toFixed(2)
    }));

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <DetailsList
          items={items}
          columns={columns}
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          compact
        />

        <Separator />

        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          Detailed Statistics
        </Text>

        {Array.from(analysis.descriptiveStats.entries()).map(([name, stats]) => (
          <Stack key={name} tokens={{ childrenGap: 8 }} styles={{ root: { padding: 12, backgroundColor: "#f3f2f1", borderRadius: 4 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: "#0078d4" } }}>
              {name}
            </Text>
            <Stack horizontal tokens={{ childrenGap: 32 }} wrap>
              <StatItem label="Sum" value={stats.sum.toLocaleString()} />
              <StatItem label="Range" value={stats.range.toFixed(2)} />
              <StatItem label="IQR" value={stats.iqr.toFixed(2)} />
              <StatItem label="Q1" value={stats.quartiles.q1.toFixed(2)} />
              <StatItem label="Q3" value={stats.quartiles.q3.toFixed(2)} />
              <StatItem label="Kurtosis" value={stats.kurtosis.toFixed(2)} />
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                95% Confidence Interval: [{stats.confidenceInterval.lower.toFixed(2)}, {stats.confidenceInterval.upper.toFixed(2)}]
              </Text>
            </Stack>
          </Stack>
        ))}
      </Stack>
    );
  };

  const renderAnomalies = () => {
    if (!analysis) return null;

    const filteredAnomalies = analysis.anomalies.filter(a => Math.abs(a.zScore) >= zScoreThreshold);

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 16 }} verticalAlign="center">
          <Stack styles={{ root: { flex: 1 } }}>
            <Slider
              label="Z-Score Threshold"
              min={1}
              max={5}
              step={0.5}
              value={zScoreThreshold}
              onChange={setZScoreThreshold}
              showValue
              valueFormat={(v) => `${v}σ`}
            />
          </Stack>
          <Toggle
            label="Use IQR Method"
            checked={useIQR}
            onChange={(_, checked) => setUseIQR(checked ?? true)}
            styles={{ root: { minWidth: 120 } }}
          />
        </Stack>

        <Text>
          Showing {filteredAnomalies.length} of {analysis.anomalies.length} anomalies
        </Text>

        {filteredAnomalies.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.success}>
            No anomalies detected with current threshold settings.
          </MessageBar>
        ) : (
          <>
            <DetailsList
              items={filteredAnomalies.map((a, index) => ({
                key: index,
                severity: a.severity,
                column: selectedRange?.values[0]?.[a.columnIndex] || `Column ${a.columnIndex + 1}`,
                row: a.rowIndex + 1,
                value: a.value,
                zScore: a.zScore.toFixed(2),
                description: a.description
              }))}
              columns={[
                {
                  key: "severity",
                  name: "Severity",
                  fieldName: "severity",
                  minWidth: 80,
                  onRender: (item) => (
                    <SeverityBadge severity={item.severity} />
                  )
                },
                { key: "column", name: "Column", fieldName: "column", minWidth: 100 },
                { key: "row", name: "Row", fieldName: "row", minWidth: 60 },
                { key: "value", name: "Value", fieldName: "value", minWidth: 100 },
                { key: "zScore", name: "Z-Score", fieldName: "zScore", minWidth: 80 },
                { key: "description", name: "Description", fieldName: "description", minWidth: 300 }
              ]}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              compact
            />

            <PrimaryButton
              iconProps={{ iconName: "Highlight" }}
              text="Apply Highlighting to Worksheet"
              onClick={() => {
                const highlights = filteredAnomalies.map(a => ({
                  row: a.rowIndex,
                  col: a.columnIndex,
                  severity: a.severity
                }));
                onApplyHighlighting?.(highlights);
              }}
            />
          </>
        )}
      </Stack>
    );
  };

  const renderTrends = () => {
    if (!analysis?.trends || analysis.trends.size === 0) {
      return <Text>No trend data available. Select a range with time-series data.</Text>;
    }

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        {Array.from(analysis.trends.entries()).map(([columnName, trend]) => (
          <Stack
            key={columnName}
            tokens={{ childrenGap: 12 }}
            styles={{ root: { padding: 16, backgroundColor: "#f3f2f1", borderRadius: 4 } }}
          >
            <Stack horizontal horizontalAlign="space-between">
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: "#0078d4" } }}>
                {columnName}
              </Text>
              <TrendBadge trendType={trend.trendType} />
            </Stack>

            <Text>{trend.summary}</Text>

            <Stack horizontal tokens={{ childrenGap: 32 }} wrap>
              <StatItem
                label="R²"
                value={trend.linearRegression.rSquared.toFixed(3)}
                tooltip="Coefficient of determination - how well the trend line fits the data"
              />
              <StatItem
                label="Slope"
                value={trend.linearRegression.slope.toFixed(4)}
                tooltip="Rate of change per period"
              />
              <StatItem
                label="Correlation"
                value={trend.linearRegression.correlation.toFixed(3)}
                tooltip="Pearson correlation coefficient"
              />
            </Stack>

            {trend.forecast.next3.length > 0 && (
              <>
                <Separator />
                <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                  Forecast (Next 3 Periods)
                </Text>
                <Stack horizontal tokens={{ childrenGap: 16 }}>
                  {trend.forecast.next3.map((value, i) => (
                    <Stack key={i} horizontalAlign="center">
                      <Text variant="xxLarge" styles={{ root: { color: "#0078d4", fontWeight: 600 } }}>
                        {value.toFixed(2)}
                      </Text>
                      <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                        T+{i + 1}
                      </Text>
                    </Stack>
                  ))}
                </Stack>
              </>
            )}
          </Stack>
        ))}
      </Stack>
    );
  };

  const renderCorrelations = () => {
    if (!analysis?.correlations) {
      return <Text>Need at least 2 numeric columns to calculate correlations.</Text>;
    }

    const { correlations } = analysis;
    const significantPairs = correlations.pairs
      .filter(p => p.isSignificant && p.strength !== "none")
      .sort((a, b) => Math.abs(b.correlation) - Math.abs(a.correlation));

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Text>{correlations.summary}</Text>

        <Separator />

        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          Significant Correlations
        </Text>

        {significantPairs.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No statistically significant correlations found.
          </MessageBar>
        ) : (
          <DetailsList
            items={significantPairs.map((pair, index) => ({
              key: index,
              columnA: pair.columnA,
              columnB: pair.columnB,
              correlation: pair.correlation.toFixed(3),
              rSquared: pair.rSquared.toFixed(3),
              pValue: pair.pValue < 0.001 ? "<0.001" : pair.pValue.toFixed(3),
              strength: pair.strength,
              direction: pair.direction
            }))}
            columns={[
              { key: "columnA", name: "Column A", fieldName: "columnA", minWidth: 100 },
              { key: "columnB", name: "Column B", fieldName: "columnB", minWidth: 100 },
              {
                key: "correlation",
                name: "Correlation",
                fieldName: "correlation",
                minWidth: 90,
                onRender: (item) => (
                  <Text
                    styles={{
                      root: {
                        color: item.direction === "positive" ? "#107c10" : "#d13438",
                        fontWeight: 600
                      }
                    }}
                  >
                    {item.direction === "positive" ? "+" : ""}{item.correlation}
                  </Text>
                )
              },
              { key: "rSquared", name: "R²", fieldName: "rSquared", minWidth: 70 },
              { key: "pValue", name: "p-value", fieldName: "pValue", minWidth: 80 },
              {
                key: "strength",
                name: "Strength",
                fieldName: "strength",
                minWidth: 90,
                onRender: (item) => <CorrelationStrengthBadge strength={item.strength} />
              }
            ]}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
            compact
          />
        )}

        <Separator />

        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          Correlation Matrix
        </Text>

        <div style={{ overflowX: "auto" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ padding: 8, border: "1px solid #e1dfdd" }}></th>
                {correlations.columns.map(col => (
                  <th key={col} style={{ padding: 8, border: "1px solid #e1dfdd", fontWeight: 600 }}>
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {correlations.matrix.map((row, i) => (
                <tr key={i}>
                  <th style={{ padding: 8, border: "1px solid #e1dfdd", fontWeight: 600 }}>
                    {correlations.columns[i]}
                  </th>
                  {row.map((val, j) => (
                    <td
                      key={j}
                      style={{
                        padding: 8,
                        border: "1px solid #e1dfdd",
                        textAlign: "center",
                        backgroundColor: i === j ? "#f3f2f1" : getCorrelationColor(val),
                        color: Math.abs(val) > 0.5 ? "white" : "#323130"
                      }}
                    >
                      {val.toFixed(2)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Stack>
    );
  };

  const renderQuality = () => {
    if (!selectedRange) return null;

    const quality = generateDataQualityReport(selectedRange);

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
          <StatCard
            icon="CellPhone"
            title="Total Cells"
            value={quality.totalCells.toLocaleString()}
          />
          <StatCard
            icon="CheckMark"
            title="Completeness"
            value={`${quality.completeness.toFixed(1)}%`}
            color={quality.completeness >= 90 ? "#107c10" : quality.completeness >= 70 ? "#ffc107" : "#d13438"}
          />
          <StatCard
            icon="Number"
            title="Numeric %"
            value={`${quality.numericPercentage.toFixed(1)}%`}
          />
          <StatCard
            icon="Error"
            title="Errors"
            value={quality.errorCells.toString()}
            color={quality.errorCells > 0 ? "#d13438" : undefined}
          />
        </Stack>

        <Separator />

        <Stack horizontal tokens={{ childrenGap: 32 }}>
          <Stack styles={{ root: { flex: 1 } }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, marginBottom: 8 } }}>
              Cell Breakdown
            </Text>
            <Stack tokens={{ childrenGap: 4 }}>
              <QualityBar label="Numeric" value={quality.numericCells} total={quality.totalCells} color="#0078d4" />
              <QualityBar label="Text" value={quality.textCells} total={quality.totalCells} color="#8764b8" />
              <QualityBar label="Empty" value={quality.emptyCells} total={quality.totalCells} color="#c8c6c4" />
              <QualityBar label="Errors" value={quality.errorCells} total={quality.totalCells} color="#d13438" />
            </Stack>
          </Stack>
        </Stack>

        <Separator />

        {quality.issues.length > 0 && (
          <>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Issues Detected
            </Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {quality.issues.map((issue, i) => (
                <MessageBar
                  key={i}
                  messageBarType={issue.includes("error") ? MessageBarType.error : MessageBarType.warning}
                >
                  {issue}
                </MessageBar>
              ))}
            </Stack>
          </>
        )}

        {quality.recommendations.length > 0 && (
          <>
            <Separator />
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Recommendations
            </Text>
            <Stack tokens={{ childrenGap: 4 }}>
              {quality.recommendations.map((rec, i) => (
                <Text key={i} styles={{ root: { color: "#605e5c" } }}>
                  • {rec}
                </Text>
              ))}
            </Stack>
          </>
        )}
      </Stack>
    );
  };

  // ============================================================================
  // Main Render
  // ============================================================================

  if (!selectedRange) {
    return (
      <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
        <Icon iconName="AnalyticsLogo" styles={{ root: { fontSize: 64, color: "#c8c6c4" } }} />
        <Text variant="large" styles={{ root: { color: "#605e5c", marginTop: 16 } }}>
          Select a data range in Excel to begin analysis
        </Text>
      </Stack>
    );
  }

  if (isAnalyzing) {
    return (
      <Stack tokens={{ padding: 32, childrenGap: 16 }} horizontalAlign="center">
        <ProgressIndicator label="Analyzing data..." description="Calculating statistics, detecting anomalies, and finding trends" />
      </Stack>
    );
  }

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error}>
        {error}
      </MessageBar>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          Data Analysis: {selectedRange.address}
        </Text>
        <DefaultButton iconProps={{ iconName: "Refresh" }} text="Reanalyze" onClick={performAnalysis} />
      </Stack>

      <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey as AnalysisTab)}>
        <PivotItem headerText="Overview" itemKey="overview" itemIcon="ViewDashboard">
          {renderOverview()}
        </PivotItem>
        <PivotItem headerText="Statistics" itemKey="statistics" itemIcon="Calculator">
          {renderStatistics()}
        </PivotItem>
        <PivotItem headerText="Anomalies" itemKey="anomalies" itemIcon="AlertSettings">
          {renderAnomalies()}
        </PivotItem>
        <PivotItem headerText="Trends" itemKey="trends" itemIcon="Trending12">
          {renderTrends()}
        </PivotItem>
        <PivotItem headerText="Correlations" itemKey="correlations" itemIcon="Relationship">
          {renderCorrelations()}
        </PivotItem>
        <PivotItem headerText="Data Quality" itemKey="quality" itemIcon="CheckList">
          {renderQuality()}
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

// ============================================================================
// Sub-components
// ============================================================================

interface StatCardProps {
  icon: string;
  title: string;
  value: string;
  color?: string;
}

const StatCard: React.FC<StatCardProps> = ({ icon, title, value, color }) => (
  <Stack
    horizontalAlign="center"
    tokens={{ padding: 16, childrenGap: 8 }}
    styles={{
      root: {
        backgroundColor: "#f3f2f1",
        borderRadius: 8,
        minWidth: 120,
        borderLeft: `4px solid ${color || "#0078d4"}`
      }
    }}
  >
    <Icon iconName={icon} styles={{ root: { fontSize: 24, color: color || "#0078d4" } }} />
    <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: color || "#323130" } }}>
      {value}
    </Text>
    <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
      {title}
    </Text>
  </Stack>
);

interface StatItemProps {
  label: string;
  value: string | number;
  tooltip?: string;
}

const StatItem: React.FC<StatItemProps> = ({ label, value, tooltip }) => (
  <TooltipHost content={tooltip}>
    <Stack>
      <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
        {label}
      </Text>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        {value}
      </Text>
    </Stack>
  </TooltipHost>
);

interface SeverityBadgeProps {
  severity: string;
}

const SeverityBadge: React.FC<SeverityBadgeProps> = ({ severity }) => {
  const colors: Record<string, string> = {
    low: "#ffc107",
    medium: "#ff9800",
    high: "#f44336",
    extreme: "#b71c1c"
  };

  return (
    <span
      style={{
        backgroundColor: colors[severity] || "#c8c6c4",
        color: severity === "low" ? "#323130" : "white",
        padding: "2px 8px",
        borderRadius: 12,
        fontSize: 12,
        fontWeight: 600,
        textTransform: "uppercase"
      }}
    >
      {severity}
    </span>
  );
};

interface TrendBadgeProps {
  trendType: string;
}

const TrendBadge: React.FC<TrendBadgeProps> = ({ trendType }) => {
  const config: Record<string, { icon: string; color: string; label: string }> = {
    increasing: { icon: "Trending12", color: "#107c10", label: "Increasing" },
    decreasing: { icon: "TrendingDown", color: "#d13438", label: "Decreasing" },
    stable: { icon: "Remove", color: "#605e5c", label: "Stable" },
    volatile: { icon: "LightningBolt", color: "#ffc107", label: "Volatile" },
    seasonal: { icon: "Calendar", color: "#8764b8", label: "Seasonal" }
  };

  const { icon, color, label } = config[trendType] || config.stable;

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
      <Icon iconName={icon} styles={{ root: { color, fontSize: 16 } }} />
      <Text styles={{ root: { color, fontWeight: 600 } }}>{label}</Text>
    </Stack>
  );
};

interface CorrelationStrengthBadgeProps {
  strength: string;
}

const CorrelationStrengthBadge: React.FC<CorrelationStrengthBadgeProps> = ({ strength }) => {
  const colors: Record<string, string> = {
    none: "#c8c6c4",
    weak: "#ffc107",
    moderate: "#ff9800",
    strong: "#107c10",
    perfect: "#0078d4"
  };

  return (
    <span
      style={{
        backgroundColor: colors[strength] || "#c8c6c4",
        color: strength === "weak" || strength === "none" ? "#323130" : "white",
        padding: "2px 8px",
        borderRadius: 12,
        fontSize: 11,
        fontWeight: 600,
        textTransform: "capitalize"
      }}
    >
      {strength}
    </span>
  );
};

interface QualityBarProps {
  label: string;
  value: number;
  total: number;
  color: string;
}

const QualityBar: React.FC<QualityBarProps> = ({ label, value, total, color }) => {
  const percentage = total > 0 ? (value / total) * 100 : 0;

  return (
    <Stack tokens={{ childrenGap: 4 }}>
      <Stack horizontal horizontalAlign="space-between">
        <Text variant="small">{label}</Text>
        <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
          {value.toLocaleString()} ({percentage.toFixed(1)}%)
        </Text>
      </Stack>
      <div
        style={{
          width: "100%",
          height: 8,
          backgroundColor: "#e1dfdd",
          borderRadius: 4,
          overflow: "hidden"
        }}
      >
        <div
          style={{
            width: `${percentage}%`,
            height: "100%",
            backgroundColor: color,
            borderRadius: 4,
            transition: "width 0.3s ease"
          }}
        />
      </div>
    </Stack>
  );
};

const getCorrelationColor = (value: number): string => {
  const abs = Math.abs(value);
  if (abs < 0.3) return "#ffffff";
  if (abs < 0.5) return value > 0 ? "#e6f4ea" : "#fce8e6";
  if (abs < 0.7) return value > 0 ? "#b7e1cd" : "#f5b7b1";
  if (abs < 0.9) return value > 0 ? "#4caf50" : "#f44336";
  return value > 0 ? "#2e7d32" : "#c62828";
};

export default DataAnalysis;
