/**
 * Excel AI Assistant - Data Analysis Service
 * Statistical analysis engine for Excel data
 * Features: Anomaly detection, Trend analysis, Correlation matrix, Descriptive statistics
 * 
 * @module services/dataAnalysis
 */

import { RangeContext } from "../types";

// ============================================================================
// Type Definitions
// ============================================================================

/** Statistical summary for a single numeric column */
export interface DescriptiveStats {
  count: number;
  mean: number;
  median: number;
  mode: number[];
  stdDev: number;
  variance: number;
  min: number;
  max: number;
  range: number;
  quartiles: {
    q1: number;
    q2: number;
    q3: number;
  };
  iqr: number;
  skewness: number;
  kurtosis: number;
  /** Coefficient of variation (stdDev / mean * 100) */
  cv: number;
  /** Confidence interval for the mean at 95% */
  confidenceInterval: {
    lower: number;
    upper: number;
    level: number;
  };
  /** Sum of all values */
  sum: number;
}

/** Anomaly detection result */
export interface Anomaly {
  rowIndex: number;
  columnIndex: number;
  value: number;
  /** Z-score of the value */
  zScore: number;
  /** Whether detected by Z-score method */
  detectedByZScore: boolean;
  /** Whether detected by IQR method */
  detectedByIQR: boolean;
  /** Severity level based on deviation */
  severity: "low" | "medium" | "high" | "extreme";
  /** Human-readable description */
  description: string;
}

/** Trend analysis result */
export interface TrendAnalysis {
  /** Type of trend detected */
  trendType: "increasing" | "decreasing" | "stable" | "volatile" | "seasonal";
  /** Linear regression coefficients: y = slope * x + intercept */
  linearRegression: {
    slope: number;
    intercept: number;
    rSquared: number;
    correlation: number;
  };
  /** Moving averages */
  movingAverages: {
    window3: number[];
    window5: number[];
    window10: number[];
  };
  /** Rate of change between consecutive periods */
  rateOfChange: number[];
  /** Percentage change between consecutive periods */
  percentageChange: number[];
  /** Statistical significance of trend */
  significance: {
    pValue: number;
    isSignificant: boolean;
  };
  /** Trend strength (0-1) */
  strength: number;
  /** Forecast for next N periods */
  forecast: {
    next3: number[];
    next5: number[];
    confidence: number;
  };
  /** Human-readable summary */
  summary: string;
}

/** Correlation between two variables */
export interface CorrelationPair {
  columnA: string;
  columnB: string;
  /** Pearson correlation coefficient (-1 to 1) */
  correlation: number;
  /** Coefficient of determination (R²) */
  rSquared: number;
  /** P-value for significance test */
  pValue: number;
  /** Whether correlation is statistically significant */
  isSignificant: boolean;
  /** Interpretation strength */
  strength: "none" | "weak" | "moderate" | "strong" | "perfect";
  /** Direction of correlation */
  direction: "positive" | "negative" | "none";
  /** Sample size used for calculation */
  sampleSize: number;
}

/** Complete correlation matrix for multiple columns */
export interface CorrelationMatrix {
  columns: string[];
  /** 2D array of correlations (correlationMatrix[i][j] = correlation between columns[i] and columns[j]) */
  matrix: number[][];
  pairs: CorrelationPair[];
  /** Strongest positive correlation */
  strongestPositive?: CorrelationPair;
  /** Strongest negative correlation */
  strongestNegative?: CorrelationPair;
  /** Summary of findings */
  summary: string;
}

/** Data column metadata */
export interface DataColumn {
  name: string;
  index: number;
  values: number[];
  hasHeaders: boolean;
}

/** Options for anomaly detection */
export interface AnomalyDetectionOptions {
  /** Z-score threshold (default: 3) */
  zScoreThreshold?: number;
  /** Whether to use IQR method (default: true) */
  useIQR?: boolean;
  /** IQR multiplier (default: 1.5) */
  iqrMultiplier?: number;
  /** Minimum number of data points required (default: 10) */
  minDataPoints?: number;
}

/** Options for trend analysis */
export interface TrendAnalysisOptions {
  /** Number of periods to forecast (default: 5) */
  forecastPeriods?: number;
  /** Confidence level for forecast (default: 0.95) */
  confidenceLevel?: number;
  /** Period labels (e.g., dates, month names) */
  periodLabels?: string[];
}

/** Options for correlation analysis */
export interface CorrelationOptions {
  /** Significance level for p-value test (default: 0.05) */
  significanceLevel?: number;
  /** Minimum sample size required (default: 3) */
  minSampleSize?: number;
  /** Method for handling missing values: 'pairwise' | 'listwise' (default: 'pairwise') */
  missingDataHandling?: "pairwise" | "listwise";
}

/** Complete analysis result for a data range */
export interface CompleteAnalysis {
  range: string;
  columnCount: number;
  rowCount: number;
  descriptiveStats: Map<string, DescriptiveStats>;
  anomalies: Anomaly[];
  correlations?: CorrelationMatrix;
  trends?: Map<string, TrendAnalysis>;
  insights: string[];
  summary: string;
}

// ============================================================================
// Utility Functions
// ============================================================================

/**
 * Extract numeric data columns from range context
 * Handles headers and converts strings to numbers where possible
 */
function extractNumericColumns(range: RangeContext): DataColumn[] {
  const columns: DataColumn[] = [];
  const { values, columnCount, rowCount } = range;
  
  // Detect if first row contains headers (heuristic: first row has strings, rest have numbers)
  const firstRow = values[0] || [];
  const secondRow = values[1] || [];
  let hasHeaders = false;
  
  if (rowCount > 1) {
    const firstRowStrings = firstRow.filter(v => typeof v === "string").length;
    const secondRowNumbers = secondRow.filter(v => typeof v === "number").length;
    hasHeaders = firstRowStrings > secondRowNumbers;
  }
  
  const dataStartRow = hasHeaders ? 1 : 0;
  
  for (let col = 0; col < columnCount; col++) {
    const columnName = hasHeaders ? String(firstRow[col] || `Column ${col + 1}`) : `Column ${col + 1}`;
    const numericValues: number[] = [];
    
    for (let row = dataStartRow; row < rowCount; row++) {
      const value = values[row]?.[col];
      if (value !== null && value !== undefined && value !== "") {
        const num = typeof value === "number" ? value : parseFloat(String(value));
        if (!isNaN(num) && isFinite(num)) {
          numericValues.push(num);
        }
      }
    }
    
    if (numericValues.length > 0) {
      columns.push({
        name: columnName,
        index: col,
        values: numericValues,
        hasHeaders
      });
    }
  }
  
  return columns;
}

/**
 * Sort numbers in ascending order
 */
function sortNumbers(arr: number[]): number[] {
  return [...arr].sort((a, b) => a - b);
}

/**
 * Calculate percentile
 */
function percentile(sortedArr: number[], p: number): number {
  if (sortedArr.length === 0) return 0;
  if (sortedArr.length === 1) return sortedArr[0];
  
  const index = (p / 100) * (sortedArr.length - 1);
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  const weight = index - lower;
  
  return sortedArr[lower] * (1 - weight) + sortedArr[upper] * weight;
}

/**
 * Calculate mean
 */
function mean(arr: number[]): number {
  if (arr.length === 0) return 0;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

/**
 * Calculate standard deviation
 */
function standardDeviation(arr: number[], sample = true): number {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  const variance = arr.reduce((sum, val) => sum + Math.pow(val - m, 2), 0) / (arr.length - (sample ? 1 : 0));
  return Math.sqrt(variance);
}

/**
 * Calculate variance
 */
function variance(arr: number[], sample = true): number {
  if (arr.length < 2) return 0;
  const m = mean(arr);
  return arr.reduce((sum, val) => sum + Math.pow(val - m, 2), 0) / (arr.length - (sample ? 1 : 0));
}

/**
 * Calculate median
 */
function median(arr: number[]): number {
  if (arr.length === 0) return 0;
  const sorted = sortNumbers(arr);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 
    ? (sorted[mid - 1] + sorted[mid]) / 2 
    : sorted[mid];
}

/**
 * Calculate mode(s)
 */
function mode(arr: number[]): number[] {
  if (arr.length === 0) return [];
  
  const frequency: Map<number, number> = new Map();
  let maxFreq = 0;
  
  for (const val of arr) {
    const freq = (frequency.get(val) || 0) + 1;
    frequency.set(val, freq);
    maxFreq = Math.max(maxFreq, freq);
  }
  
  // Return all values with max frequency (if more than one, it's multimodal)
  const modes: number[] = [];
  for (const [val, freq] of frequency) {
    if (freq === maxFreq && freq > 1) {
      modes.push(val);
    }
  }
  
  return modes.length > 0 ? modes.sort((a, b) => a - b) : [];
}

/**
 * Calculate skewness
 */
function skewness(arr: number[]): number {
  if (arr.length < 3) return 0;
  const m = mean(arr);
  const s = standardDeviation(arr);
  if (s === 0) return 0;
  
  const n = arr.length;
  const sumCubedDeviations = arr.reduce((sum, val) => sum + Math.pow(val - m, 3), 0);
  return (n / ((n - 1) * (n - 2))) * (sumCubedDeviations / Math.pow(s, 3));
}

/**
 * Calculate kurtosis
 */
function kurtosis(arr: number[]): number {
  if (arr.length < 4) return 0;
  const m = mean(arr);
  const s = standardDeviation(arr);
  if (s === 0) return 0;
  
  const n = arr.length;
  const sumFourthPowers = arr.reduce((sum, val) => sum + Math.pow(val - m, 4), 0);
  const numerator = n * (n + 1) * sumFourthPowers;
  const denominator = (n - 1) * (n - 2) * (n - 3) * Math.pow(s, 4);
  const adjustment = (3 * Math.pow(n - 1, 2)) / ((n - 2) * (n - 3));
  
  return (numerator / denominator) - adjustment;
}

/**
 * Calculate confidence interval for the mean
 */
function confidenceInterval(arr: number[], level = 0.95): { lower: number; upper: number; level: number } {
  if (arr.length < 2) return { lower: 0, upper: 0, level };
  
  const m = mean(arr);
  const s = standardDeviation(arr);
  const se = s / Math.sqrt(arr.length);
  
  // Z-score for given confidence level (approximation)
  const zScores: Record<number, number> = {
    0.9: 1.645,
    0.95: 1.96,
    0.99: 2.576
  };
  const z = zScores[level] || 1.96;
  
  const margin = z * se;
  return {
    lower: m - margin,
    upper: m + margin,
    level
  };
}

/**
 * Calculate Z-score
 */
function zScore(value: number, mean: number, stdDev: number): number {
  if (stdDev === 0) return 0;
  return (value - mean) / stdDev;
}

/**
 * Calculate t-statistic for correlation significance
 */
function correlationSignificance(r: number, n: number): { pValue: number; isSignificant: boolean } {
  if (n < 3 || Math.abs(r) >= 1) return { pValue: 1, isSignificant: false };
  
  const t = r * Math.sqrt((n - 2) / (1 - r * r));
  // Simplified p-value calculation (two-tailed test)
  // For production, use a proper t-distribution CDF
  const pValue = Math.min(1, 2 * (1 - tDistributionCDF(Math.abs(t), n - 2)));
  
  return { pValue, isSignificant: pValue < 0.05 };
}

/**
 * Simplified t-distribution CDF (for correlation significance)
 */
function tDistributionCDF(t: number, df: number): number {
  // Approximation of t-distribution CDF
  // Using approximation: P(T <= t) ≈ Φ(t * sqrt((df - 2/3) / df)) for df > 2
  if (df <= 0) return 0.5;
  if (df === 1) return 0.5 + Math.atan(t) / Math.PI;
  
  const adjustedT = t * Math.sqrt((df - 2/3) / df);
  return normalCDF(adjustedT);
}

/**
 * Standard normal CDF
 */
function normalCDF(x: number): number {
  // Approximation of standard normal CDF
  const a1 = 0.254829592;
  const a2 = -0.284496736;
  const a3 = 1.421413741;
  const a4 = -1.453152027;
  const a5 = 1.061405429;
  const p = 0.3275911;
  
  const sign = x < 0 ? -1 : 1;
  x = Math.abs(x) / Math.sqrt(2);
  
  const t = 1 / (1 + p * x);
  const y = 1 - (((((a5 * t + a4) * t) + a3) * t + a2) * t + a1) * t * Math.exp(-x * x);
  
  return 0.5 * (1 + sign * y);
}

/**
 * Get correlation strength description
 */
function getCorrelationStrength(r: number): "none" | "weak" | "moderate" | "strong" | "perfect" {
  const abs = Math.abs(r);
  if (abs === 1) return "perfect";
  if (abs >= 0.7) return "strong";
  if (abs >= 0.4) return "moderate";
  if (abs >= 0.1) return "weak";
  return "none";
}

/**
 * Get correlation direction
 */
function getCorrelationDirection(r: number): "positive" | "negative" | "none" {
  if (r > 0.1) return "positive";
  if (r < -0.1) return "negative";
  return "none";
}

/**
 * Interpret skewness
 */
function interpretSkewness(skew: number): string {
  if (Math.abs(skew) < 0.5) return "approximately symmetric";
  if (skew >= 0.5 && skew < 1) return "moderately right-skewed";
  if (skew >= 1) return "highly right-skewed";
  if (skew <= -0.5 && skew > -1) return "moderately left-skewed";
  return "highly left-skewed";
}

// ============================================================================
// Main Analysis Functions
// ============================================================================

/**
 * Calculate descriptive statistics for a numeric array
 */
export function calculateDescriptiveStats(values: number[]): DescriptiveStats | null {
  if (values.length === 0) return null;
  
  const sorted = sortNumbers(values);
  const m = mean(values);
  const s = standardDeviation(values);
  const v = variance(values);
  const min = sorted[0];
  const max = sorted[sorted.length - 1];
  
  return {
    count: values.length,
    mean: m,
    median: median(values),
    mode: mode(values),
    stdDev: s,
    variance: v,
    min,
    max,
    range: max - min,
    quartiles: {
      q1: percentile(sorted, 25),
      q2: percentile(sorted, 50),
      q3: percentile(sorted, 75)
    },
    iqr: percentile(sorted, 75) - percentile(sorted, 25),
    skewness: skewness(values),
    kurtosis: kurtosis(values),
    cv: m !== 0 ? (s / Math.abs(m)) * 100 : 0,
    confidenceInterval: confidenceInterval(values, 0.95),
    sum: values.reduce((a, b) => a + b, 0)
  };
}

/**
 * Detect anomalies in numeric data using Z-score and IQR methods
 */
export function detectAnomalies(
  values: number[],
  columnName: string,
  options: AnomalyDetectionOptions = {}
): Anomaly[] {
  const {
    zScoreThreshold = 3,
    useIQR = true,
    iqrMultiplier = 1.5,
    minDataPoints = 10
  } = options;
  
  if (values.length < minDataPoints) return [];
  
  const anomalies: Anomaly[] = [];
  const m = mean(values);
  const s = standardDeviation(values);
  const sorted = sortNumbers(values);
  const q1 = percentile(sorted, 25);
  const q3 = percentile(sorted, 75);
  const iqr = q3 - q1;
  
  const iqrLower = q1 - iqrMultiplier * iqr;
  const iqrUpper = q3 + iqrMultiplier * iqr;
  
  values.forEach((value, index) => {
    const z = zScore(value, m, s);
    const isZScoreAnomaly = Math.abs(z) > zScoreThreshold;
    const isIQRAnomaly = useIQR && (value < iqrLower || value > iqrUpper);
    
    if (isZScoreAnomaly || isIQRAnomaly) {
      let severity: Anomaly["severity"] = "low";
      const absZ = Math.abs(z);
      if (absZ > 5) severity = "extreme";
      else if (absZ > 4) severity = "high";
      else if (absZ > 3) severity = "medium";
      
      const deviation = ((value - m) / m) * 100;
      const description = `${value} in ${columnName} is ${absZ.toFixed(2)} standard deviations ${z > 0 ? "above" : "below"} the mean (${deviation > 0 ? "+" : ""}${deviation.toFixed(1)}%)`;
      
      anomalies.push({
        rowIndex: index,
        columnIndex: 0, // Will be set by caller if available
        value,
        zScore: z,
        detectedByZScore: isZScoreAnomaly,
        detectedByIQR: isIQRAnomaly,
        severity,
        description
      });
    }
  });
  
  return anomalies.sort((a, b) => Math.abs(b.zScore) - Math.abs(a.zScore));
}

/**
 * Perform trend analysis on time-series data
 */
export function analyzeTrend(
  values: number[],
  options: TrendAnalysisOptions = {}
): TrendAnalysis | null {
  const { forecastPeriods = 5, confidenceLevel = 0.95 } = options;
  
  if (values.length < 3) return null;
  
  const n = values.length;
  const x = Array.from({ length: n }, (_, i) => i);
  
  // Linear regression
  const xMean = mean(x);
  const yMean = mean(values);
  
  let numerator = 0;
  let denominator = 0;
  
  for (let i = 0; i < n; i++) {
    numerator += (x[i] - xMean) * (values[i] - yMean);
    denominator += Math.pow(x[i] - xMean, 2);
  }
  
  const slope = denominator !== 0 ? numerator / denominator : 0;
  const intercept = yMean - slope * xMean;
  
  // Calculate R-squared
  const ssRes = values.reduce((sum, y, i) => sum + Math.pow(y - (slope * x[i] + intercept), 2), 0);
  const ssTot = values.reduce((sum, y) => sum + Math.pow(y - yMean, 2), 0);
  const rSquared = ssTot !== 0 ? 1 - (ssRes / ssTot) : 0;
  
  // Calculate correlation
  const correlation = Math.sqrt(rSquared) * (slope >= 0 ? 1 : -1);
  
  // Moving averages
  const movingAverage = (window: number): number[] => {
    const result: number[] = [];
    for (let i = window - 1; i < values.length; i++) {
      const windowValues = values.slice(i - window + 1, i + 1);
      result.push(mean(windowValues));
    }
    return result;
  };
  
  // Rate of change
  const rateOfChange: number[] = [];
  const percentageChange: number[] = [];
  
  for (let i = 1; i < values.length; i++) {
    const change = values[i] - values[i - 1];
    rateOfChange.push(change);
    percentageChange.push(values[i - 1] !== 0 ? (change / values[i - 1]) * 100 : 0);
  }
  
  // Determine trend type
  let trendType: TrendAnalysis["trendType"] = "stable";
  
  if (rSquared > 0.5) {
    if (slope > 0.01) trendType = "increasing";
    else if (slope < -0.01) trendType = "decreasing";
  } else {
    const volatility = standardDeviation(values) / Math.abs(yMean);
    if (volatility > 0.3) trendType = "volatile";
  }
  
  // Check for seasonality (simplified)
  if (values.length >= 12) {
    const autocorrelation = calculateAutocorrelation(values, Math.floor(values.length / 4));
    if (autocorrelation > 0.5) trendType = "seasonal";
  }
  
  // Trend strength
  const strength = rSquared;
  
  // Forecast
  const standardError = Math.sqrt(ssRes / (n - 2));
  const tValue = 1.96; // Approximate for 95% confidence
  
  const forecastNext = (periods: number): number[] => {
    const result: number[] = [];
    for (let i = 1; i <= periods; i++) {
      const xNew = n - 1 + i;
      result.push(slope * xNew + intercept);
    }
    return result;
  };
  
  // Significance test
  const seSlope = Math.sqrt(ssRes / ((n - 2) * denominator));
  const tStat = slope / seSlope;
  const sigResult = correlationSignificance(correlation, n);
  
  // Generate summary
  const trendWord = slope > 0 ? "increasing" : slope < 0 ? "decreasing" : "stable";
  const strengthWord = strength > 0.7 ? "strong" : strength > 0.4 ? "moderate" : "weak";
  const avgChange = mean(percentageChange);
  
  const summary = `Data shows a ${strengthWord} ${trendWord} trend (R² = ${rSquared.toFixed(3)}). ` +
    `Average change per period: ${avgChange > 0 ? "+" : ""}${avgChange.toFixed(2)}%. ` +
    `Trend is ${sigResult.isSignificant ? "statistically significant" : "not statistically significant"} (p = ${sigResult.pValue.toFixed(4)}).`;
  
  return {
    trendType,
    linearRegression: {
      slope,
      intercept,
      rSquared,
      correlation
    },
    movingAverages: {
      window3: movingAverage(3),
      window5: movingAverage(5),
      window10: movingAverage(10)
    },
    rateOfChange,
    percentageChange,
    significance: sigResult,
    strength,
    forecast: {
      next3: forecastNext(3),
      next5: forecastNext(5),
      confidence: confidenceLevel
    },
    summary
  };
}

/**
 * Calculate autocorrelation for detecting seasonality
 */
function calculateAutocorrelation(values: number[], lag: number): number {
  if (values.length <= lag) return 0;
  
  const n = values.length - lag;
  const mean1 = mean(values.slice(0, n));
  const mean2 = mean(values.slice(lag));
  
  let numerator = 0;
  let denom1 = 0;
  let denom2 = 0;
  
  for (let i = 0; i < n; i++) {
    const diff1 = values[i] - mean1;
    const diff2 = values[i + lag] - mean2;
    numerator += diff1 * diff2;
    denom1 += diff1 * diff1;
    denom2 += diff2 * diff2;
  }
  
  return denom1 > 0 && denom2 > 0 ? numerator / Math.sqrt(denom1 * denom2) : 0;
}

/**
 * Calculate correlation matrix for multiple columns
 */
export function calculateCorrelationMatrix(
  columns: DataColumn[],
  options: CorrelationOptions = {}
): CorrelationMatrix | null {
  const { significanceLevel = 0.05, minSampleSize = 3 } = options;
  
  if (columns.length < 2) return null;
  
  const n = columns.length;
  const matrix: number[][] = Array(n).fill(null).map(() => Array(n).fill(0));
  const pairs: CorrelationPair[] = [];
  
  // Calculate correlations
  for (let i = 0; i < n; i++) {
    matrix[i][i] = 1; // Diagonal is always 1
    
    for (let j = i + 1; j < n; j++) {
      const colA = columns[i];
      const colB = columns[j];
      
      // Pairwise complete observations
      const pairs_: [number, number][] = [];
      const maxLen = Math.min(colA.values.length, colB.values.length);
      
      for (let k = 0; k < maxLen; k++) {
        if (colA.values[k] !== undefined && colB.values[k] !== undefined) {
          pairs_.push([colA.values[k], colB.values[k]]);
        }
      }
      
      if (pairs_.length < minSampleSize) {
        matrix[i][j] = matrix[j][i] = 0;
        continue;
      }
      
      const xValues = pairs_.map(p => p[0]);
      const yValues = pairs_.map(p => p[1]);
      
      const xMean = mean(xValues);
      const yMean = mean(yValues);
      
      let numerator = 0;
      let denomX = 0;
      let denomY = 0;
      
      for (const [x, y] of pairs_) {
        const diffX = x - xMean;
        const diffY = y - yMean;
        numerator += diffX * diffY;
        denomX += diffX * diffX;
        denomY += diffY * diffY;
      }
      
      const correlation = denomX > 0 && denomY > 0 ? numerator / Math.sqrt(denomX * denomY) : 0;
      
      matrix[i][j] = matrix[j][i] = correlation;
      
      const sig = correlationSignificance(correlation, pairs_.length);
      
      pairs.push({
        columnA: colA.name,
        columnB: colB.name,
        correlation,
        rSquared: correlation * correlation,
        pValue: sig.pValue,
        isSignificant: sig.isSignificant,
        strength: getCorrelationStrength(correlation),
        direction: getCorrelationDirection(correlation),
        sampleSize: pairs_.length
      });
    }
  }
  
  // Find strongest correlations
  const sortedPairs = [...pairs].sort((a, b) => Math.abs(b.correlation) - Math.abs(a.correlation));
  const strongestPositive = sortedPairs.find(p => p.correlation > 0 && p.correlation < 1);
  const strongestNegative = sortedPairs.find(p => p.correlation < 0);
  
  // Generate summary
  const significantCount = pairs.filter(p => p.isSignificant && p.strength !== "none").length;
  const strongCount = pairs.filter(p => p.strength === "strong" || p.strength === "perfect").length;
  
  let summary = `Analyzed ${n} columns with ${pairs.length} correlation pairs. `;
  summary += `Found ${significantCount} statistically significant correlations, `;
  summary += `including ${strongCount} strong correlations. `;
  
  if (strongestPositive) {
    summary += `Strongest positive: ${strongestPositive.columnA} ↔ ${strongestPositive.columnB} (r=${strongestPositive.correlation.toFixed(3)}). `;
  }
  if (strongestNegative) {
    summary += `Strongest negative: ${strongestNegative.columnA} ↔ ${strongestNegative.columnB} (r=${strongestNegative.correlation.toFixed(3)}).`;
  }
  
  return {
    columns: columns.map(c => c.name),
    matrix,
    pairs,
    strongestPositive,
    strongestNegative,
    summary
  };
}

// ============================================================================
// High-Level Analysis Functions
// ============================================================================

/**
 * Perform complete analysis on a range
 */
export function analyzeRange(range: RangeContext): CompleteAnalysis {
  const columns = extractNumericColumns(range);
  const descriptiveStats = new Map<string, DescriptiveStats>();
  const allAnomalies: Anomaly[] = [];
  const trends = new Map<string, TrendAnalysis>();
  const insights: string[] = [];
  
  // Analyze each column
  columns.forEach(column => {
    // Descriptive stats
    const stats = calculateDescriptiveStats(column.values);
    if (stats) {
      descriptiveStats.set(column.name, stats);
      
      // Generate insights
      const cv = stats.cv;
      if (cv < 10) {
        insights.push(`${column.name} has very consistent values (CV: ${cv.toFixed(1)}%).`);
      } else if (cv > 50) {
        insights.push(`${column.name} shows high variability (CV: ${cv.toFixed(1)}%).`);
      }
      
      const skewInterpretation = interpretSkewness(stats.skewness);
      if (Math.abs(stats.skewness) > 1) {
        insights.push(`${column.name} is ${skewInterpretation} (skewness: ${stats.skewness.toFixed(2)}).`);
      }
      
      // Anomaly detection
      const anomalies = detectAnomalies(column.values, column.name);
      anomalies.forEach(a => {
        a.columnIndex = column.index;
        allAnomalies.push(a);
      });
      
      if (anomalies.length > 0) {
        const extremeCount = anomalies.filter(a => a.severity === "extreme" || a.severity === "high").length;
        if (extremeCount > 0) {
          insights.push(`${column.name} contains ${extremeCount} significant outlier(s).`);
        }
      }
      
      // Trend analysis (if enough data points)
      if (column.values.length >= 5) {
        const trend = analyzeTrend(column.values);
        if (trend) {
          trends.set(column.name, trend);
          if (trend.significance.isSignificant) {
            insights.push(`${column.name}: ${trend.summary}`);
          }
        }
      }
    }
  });
  
  // Correlation analysis (if multiple columns)
  let correlations: CorrelationMatrix | undefined = undefined;
  if (columns.length >= 2) {
    const correlationResult = calculateCorrelationMatrix(columns);
    if (correlationResult) {
      correlations = correlationResult;
    }
    if (correlations && correlations.strongestPositive) {
      insights.push(correlations.summary);
    }
  }
  
  // Generate overall summary
  const summaryParts: string[] = [];
  summaryParts.push(`Analyzed ${columns.length} numeric column(s) with ${range.rowCount - (columns[0]?.hasHeaders ? 1 : 0)} data rows.`);
  
  if (allAnomalies.length > 0) {
    const extremeCount = allAnomalies.filter(a => a.severity === "extreme" || a.severity === "high").length;
    summaryParts.push(`Detected ${allAnomalies.length} anomaly/anomalies${extremeCount > 0 ? ` (${extremeCount} high/extreme)` : ""}.`);
  }
  
  if (correlations) {
    const strongCorrelations = correlations.pairs.filter(p => p.strength === "strong").length;
    if (strongCorrelations > 0) {
      summaryParts.push(`Found ${strongCorrelations} strong correlation(s).`);
    }
  }
  
  const significantTrends = Array.from(trends.values()).filter(t => t.significance.isSignificant).length;
  if (significantTrends > 0) {
    summaryParts.push(`${significantTrends} column(s) show significant trends.`);
  }
  
  return {
    range: range.address,
    columnCount: columns.length,
    rowCount: range.rowCount,
    descriptiveStats,
    anomalies: allAnomalies,
    correlations,
    trends,
    insights,
    summary: summaryParts.join(" ")
  };
}

/**
 * Compare two datasets and generate insights
 */
export function compareDatasets(
  datasetA: { name: string; values: number[] },
  datasetB: { name: string; values: number[] }
): {
  comparison: string;
  statsA: DescriptiveStats | null;
  statsB: DescriptiveStats | null;
  difference: {
    meanDiff: number;
    meanDiffPercent: number;
    medianDiff: number;
    stdDevDiff: number;
    varianceRatio: number;
  };
  tTest: {
    tStatistic: number;
    degreesOfFreedom: number;
    pValue: number;
    isSignificant: boolean;
    interpretation: string;
  } | null;
} {
  const statsA = calculateDescriptiveStats(datasetA.values);
  const statsB = calculateDescriptiveStats(datasetB.values);
  
  if (!statsA || !statsB) {
    return {
      comparison: "Insufficient data for comparison",
      statsA,
      statsB,
      difference: {
        meanDiff: 0,
        meanDiffPercent: 0,
        medianDiff: 0,
        stdDevDiff: 0,
        varianceRatio: 0
      },
      tTest: null
    };
  }
  
  const meanDiff = statsB.mean - statsA.mean;
  const meanDiffPercent = statsA.mean !== 0 ? (meanDiff / Math.abs(statsA.mean)) * 100 : 0;
  const medianDiff = statsB.median - statsA.median;
  const stdDevDiff = statsB.stdDev - statsA.stdDev;
  const varianceRatio = statsA.variance > 0 ? statsB.variance / statsA.variance : 0;
  
  // Define tTest type
  type TTestResult = {
    tStatistic: number;
    degreesOfFreedom: number;
    pValue: number;
    isSignificant: boolean;
    interpretation: string;
  } | null;
  
  // Independent samples t-test (Welch's t-test for unequal variances)
  let tTest: TTestResult = null;
  
  if (datasetA.values.length > 1 && datasetB.values.length > 1) {
    const seA = statsA.stdDev / Math.sqrt(statsA.count);
    const seB = statsB.stdDev / Math.sqrt(statsB.count);
    const seDiff = Math.sqrt(seA * seA + seB * seB);
    
    if (seDiff > 0) {
      const tStatistic = meanDiff / seDiff;
      
      // Welch-Satterthwaite equation for degrees of freedom
      const numerator = Math.pow(seA * seA + seB * seB, 2);
      const denominator = Math.pow(seA * seA, 2) / (statsA.count - 1) + Math.pow(seB * seB, 2) / (statsB.count - 1);
      const df = denominator > 0 ? numerator / denominator : statsA.count + statsB.count - 2;
      
      // Simplified p-value (two-tailed)
      const pValue = 2 * (1 - tDistributionCDF(Math.abs(tStatistic), df));
      const isSignificant = pValue < 0.05;
      
      let interpretation = "No significant difference between datasets";
      if (isSignificant) {
        interpretation = meanDiff > 0 
          ? `${datasetB.name} is significantly higher than ${datasetA.name}`
          : `${datasetB.name} is significantly lower than ${datasetA.name}`;
      }
      
      tTest = {
        tStatistic,
        degreesOfFreedom: Math.floor(df),
        pValue,
        isSignificant,
        interpretation
      };
    }
  }
  
  // Build comparison string after tTest is computed
  let comparison = "";
  if (tTest) {
    comparison = `Comparing ${datasetA.name} (n=${statsA.count}, μ=${statsA.mean.toFixed(2)}) ` +
      `vs ${datasetB.name} (n=${statsB.count}, μ=${statsB.mean.toFixed(2)}): ` +
      `Mean difference is ${meanDiff > 0 ? "+" : ""}${meanDiff.toFixed(2)} (${meanDiffPercent > 0 ? "+" : ""}${meanDiffPercent.toFixed(1)}%). ` +
      `${tTest.isSignificant ? "Statistically significant" : "Not statistically significant"} difference.`;
  } else {
    comparison = `Comparing ${datasetA.name} vs ${datasetB.name}: Insufficient data for statistical testing.`;
  }
  
  const result: {
    comparison: string;
    statsA: DescriptiveStats | null;
    statsB: DescriptiveStats | null;
    difference: {
      meanDiff: number;
      meanDiffPercent: number;
      medianDiff: number;
      stdDevDiff: number;
      varianceRatio: number;
    };
    tTest: TTestResult;
  } = {
    comparison,
    statsA,
    statsB,
    difference: {
      meanDiff,
      meanDiffPercent,
      medianDiff,
      stdDevDiff,
      varianceRatio
    },
    tTest
  };
  
  return result;
}

/**
 * Generate data quality report
 */
export function generateDataQualityReport(range: RangeContext): {
  totalCells: number;
  emptyCells: number;
  numericCells: number;
  textCells: number;
  errorCells: number;
  completeness: number;
  numericPercentage: number;
  issues: string[];
  recommendations: string[];
} {
  const { values, rowCount, columnCount } = range;
  const totalCells = rowCount * columnCount;
  
  let emptyCells = 0;
  let numericCells = 0;
  let textCells = 0;
  let errorCells = 0;
  
  const issues: string[] = [];
  
  for (let row = 0; row < rowCount; row++) {
    for (let col = 0; col < columnCount; col++) {
      const value = values[row]?.[col];
      
      if (value === null || value === undefined || value === "") {
        emptyCells++;
      } else if (typeof value === "number") {
        numericCells++;
      } else if (typeof value === "string") {
        if (value.startsWith("#") && value.includes("!")) {
          errorCells++;
        } else {
          textCells++;
        }
      }
    }
  }
  
  const completeness = ((totalCells - emptyCells) / totalCells) * 100;
  const numericPercentage = totalCells > 0 ? (numericCells / totalCells) * 100 : 0;
  
  if (emptyCells > totalCells * 0.1) {
    issues.push(`${emptyCells} empty cells (${((emptyCells/totalCells)*100).toFixed(1)}% missing data)`);
  }
  
  if (errorCells > 0) {
    issues.push(`${errorCells} cells contain errors`);
  }
  
  const recommendations: string[] = [];
  
  if (completeness < 90) {
    recommendations.push("Consider data imputation for missing values");
  }
  
  if (errorCells > 0) {
    recommendations.push("Review and fix formula errors before analysis");
  }
  
  if (numericPercentage < 50 && numericCells > 0) {
    recommendations.push("Dataset is primarily text-based; limited statistical analysis possible");
  }
  
  return {
    totalCells,
    emptyCells,
    numericCells,
    textCells,
    errorCells,
    completeness,
    numericPercentage,
    issues: issues.length > 0 ? issues : ["No major data quality issues detected"],
    recommendations: recommendations.length > 0 ? recommendations : ["Data quality is good"]
  };
}

// ============================================================================
// Natural Language Interface
// ============================================================================

/**
 * Parse natural language analysis request
 */
export function parseAnalysisRequest(request: string): {
  type: "descriptive" | "anomaly" | "trend" | "correlation" | "quality" | "complete";
  columns?: string[];
  options?: Record<string, any>;
} {
  const lowerRequest = request.toLowerCase();
  
  // Detect analysis type
  let type: ReturnType<typeof parseAnalysisRequest>["type"] = "complete";
  
  if (lowerRequest.includes("describe") || lowerRequest.includes("statistics") || lowerRequest.includes("summary")) {
    type = "descriptive";
  } else if (lowerRequest.includes("anomal") || lowerRequest.includes("outlier") || lowerRequest.includes("detect")) {
    type = "anomaly";
  } else if (lowerRequest.includes("trend") || lowerRequest.includes("forecast") || lowerRequest.includes("prediction")) {
    type = "trend";
  } else if (lowerRequest.includes("correlation") || lowerRequest.includes("relationship") || lowerRequest.includes("associate")) {
    type = "correlation";
  } else if (lowerRequest.includes("quality") || lowerRequest.includes("clean") || lowerRequest.includes("missing")) {
    type = "quality";
  }
  
  // Extract column names (simplified - looks for quoted strings or capitalized words)
  const columns: string[] = [];
  const quotedMatch = request.match(/"([^"]+)"/g);
  if (quotedMatch) {
    quotedMatch.forEach(match => columns.push(match.replace(/"/g, "")));
  }
  
  return { type, columns, options: {} };
}

/**
 * Format analysis results for display
 */
export function formatAnalysisResults(analysis: CompleteAnalysis): string {
  const lines: string[] = [];
  
  lines.push("# Data Analysis Report");
  lines.push("");
  lines.push(`**Range:** ${analysis.range}`);
  lines.push(`**Columns analyzed:** ${analysis.columnCount}`);
  lines.push(`**Rows:** ${analysis.rowCount}`);
  lines.push("");
  
  // Summary
  lines.push("## Summary");
  lines.push(analysis.summary);
  lines.push("");
  
  // Insights
  if (analysis.insights.length > 0) {
    lines.push("## Key Insights");
    analysis.insights.forEach(insight => {
      lines.push(`- ${insight}`);
    });
    lines.push("");
  }
  
  // Descriptive Statistics
  lines.push("## Descriptive Statistics");
  analysis.descriptiveStats.forEach((stats, name) => {
    lines.push(`### ${name}`);
    lines.push(`- Count: ${stats.count.toLocaleString()}`);
    lines.push(`- Mean: ${stats.mean.toFixed(2)}`);
    lines.push(`- Median: ${stats.median.toFixed(2)}`);
    lines.push(`- Std Dev: ${stats.stdDev.toFixed(2)}`);
    lines.push(`- Min: ${stats.min.toFixed(2)}`);
    lines.push(`- Max: ${stats.max.toFixed(2)}`);
    lines.push(`- Range: ${stats.range.toFixed(2)}`);
    lines.push("");
  });
  
  // Anomalies
  if (analysis.anomalies.length > 0) {
    lines.push("## Anomalies Detected");
    const topAnomalies = analysis.anomalies.slice(0, 10);
    topAnomalies.forEach(anomaly => {
      lines.push(`- **${anomaly.severity.toUpperCase()}** | ${anomaly.description}`);
    });
    if (analysis.anomalies.length > 10) {
      lines.push(`- *... and ${analysis.anomalies.length - 10} more*`);
    }
    lines.push("");
  }
  
  // Correlations
  if (analysis.correlations && analysis.correlations.pairs.length > 0) {
    lines.push("## Correlations");
    const significantPairs = analysis.correlations.pairs
      .filter(p => p.isSignificant && p.strength !== "none")
      .slice(0, 5);
    
    significantPairs.forEach(pair => {
      const emoji = pair.direction === "positive" ? "📈" : "📉";
      lines.push(`${emoji} **${pair.columnA}** ↔ **${pair.columnB}**: r=${pair.correlation.toFixed(3)} (${pair.strength} ${pair.direction})`);
    });
    lines.push("");
  }
  
  return lines.join("\n");
}

// ============================================================================
// Export for Excel Integration
// ============================================================================

/**
 * Export descriptive statistics to Excel-friendly format
 */
export function exportStatsToTable(stats: Map<string, DescriptiveStats>): (string | number)[][] {
  const headers = ["Column", "Count", "Mean", "Median", "Std Dev", "Min", "Max", "Range", "CV%", "Skewness"];
  const rows: (string | number)[][] = [headers];
  
  stats.forEach((stat, name) => {
    rows.push([
      name,
      stat.count,
      Number(stat.mean.toFixed(2)),
      Number(stat.median.toFixed(2)),
      Number(stat.stdDev.toFixed(2)),
      Number(stat.min.toFixed(2)),
      Number(stat.max.toFixed(2)),
      Number(stat.range.toFixed(2)),
      Number(stat.cv.toFixed(1)),
      Number(stat.skewness.toFixed(2))
    ]);
  });
  
  return rows;
}

/**
 * Export correlation matrix to Excel-friendly format
 */
export function exportCorrelationToTable(matrix: CorrelationMatrix): (string | number)[][] {
  const headers = ["", ...matrix.columns];
  const rows: (string | number)[][] = [headers];
  
  matrix.columns.forEach((col, i) => {
    const row: (string | number)[] = [col];
    matrix.matrix[i].forEach(corr => {
      row.push(Number(corr.toFixed(3)));
    });
    rows.push(row);
  });
  
  return rows;
}

/**
 * Export anomalies to Excel-friendly format for highlighting
 */
export function exportAnomaliesForHighlighting(anomalies: Anomaly[]): { row: number; col: number; severity: string }[] {
  return anomalies.map(a => ({
    row: a.rowIndex,
    col: a.columnIndex,
    severity: a.severity
  }));
}
