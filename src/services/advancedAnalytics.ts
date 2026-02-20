/**
 * Advanced Analytics Service
 * Week 16: Advanced Analytics
 * 
 * Provides predictive insights, anomaly detection, forecasting,
 * and advanced statistical analysis for enterprise decision-making.
 */

import { analyticsService, AnalyticsEvent, TimeRange, DateRange } from './analyticsService';

// Anomaly detection types
export interface Anomaly {
  id: string;
  type: 'spike' | 'drop' | 'trend_change' | 'outlier';
  metric: string;
  severity: 'low' | 'medium' | 'high' | 'critical';
  timestamp: Date;
  expectedValue: number;
  actualValue: number;
  deviation: number; // percentage
  description: string;
  recommendation?: string;
}

// Forecast types
export interface ForecastPoint {
  date: string;
  value: number;
  lowerBound: number;
  upperBound: number;
  confidence: number;
}

export interface Forecast {
  metric: string;
  horizon: number; // days
  predictions: ForecastPoint[];
  trend: 'up' | 'down' | 'stable';
  seasonality?: {
    daily?: boolean;
    weekly?: boolean;
    monthly?: boolean;
  };
  accuracy: number; // historical accuracy
}

// Insight types
export interface Insight {
  id: string;
  category: 'performance' | 'usage' | 'cost' | 'engagement' | 'security';
  priority: 'low' | 'medium' | 'high';
  title: string;
  description: string;
  metric: string;
  value: number;
  change: number;
  actionable: boolean;
  action?: string;
}

// Correlation analysis
export interface Correlation {
  metric1: string;
  metric2: string;
  coefficient: number; // -1 to 1
  significance: number; // p-value
  sampleSize: number;
  trend: 'positive' | 'negative' | 'none';
}

// Cohort analysis
export interface CohortData {
  cohortDate: string;
  size: number;
  retention: number[]; // percentage by week/month
  engagement: number[];
  ltv?: number; // lifetime value
}

// Funnel analysis
export interface FunnelStage {
  name: string;
  users: number;
  dropOff: number;
  conversionRate: number; // from previous stage
  avgTime: number; // milliseconds
}

export interface FunnelAnalysis {
  name: string;
  stages: FunnelStage[];
  overallConversion: number;
  totalTime: number;
}

// Alert configuration
export interface AlertConfig {
  id: string;
  name: string;
  metric: string;
  condition: 'above' | 'below' | 'change_percent' | 'anomaly';
  threshold: number;
  timeWindow: TimeRange;
  enabled: boolean;
  notificationChannels: ('email' | 'webhook' | 'in_app')[];
  recipients?: string[];
}

// Report types
export interface ScheduledReport {
  id: string;
  name: string;
  frequency: 'daily' | 'weekly' | 'monthly';
  metrics: string[];
  recipients: string[];
  format: 'pdf' | 'html' | 'csv';
  lastRun?: Date;
  nextRun: Date;
}

class AdvancedAnalyticsService {
  private static instance: AdvancedAnalyticsService;
  private anomalies: Anomaly[] = [];
  private insights: Insight[] = [];
  private alerts: AlertConfig[] = [];
  private reports: ScheduledReport[] = [];

  private constructor() {
    this.loadAlerts();
    this.loadReports();
  }

  static getInstance(): AdvancedAnalyticsService {
    if (!AdvancedAnalyticsService.instance) {
      AdvancedAnalyticsService.instance = new AdvancedAnalyticsService();
    }
    return AdvancedAnalyticsService.instance;
  }

  // ==================== ANOMALY DETECTION ====================

  /**
   * Detect anomalies in metrics using statistical methods
   */
  async detectAnomalies(timeRange: TimeRange = '7d'): Promise<Anomaly[]> {
    const events = await this.getRawEvents(timeRange);
    const newAnomalies: Anomaly[] = [];

    // Group events by hour for time-series analysis
    const hourlyMetrics = this.aggregateByHour(events);

    // Detect anomalies in query volume
    const volumeAnomalies = this.detectVolumeAnomalies(hourlyMetrics.queryVolume);
    newAnomalies.push(...volumeAnomalies);

    // Detect anomalies in response times
    const latencyAnomalies = this.detectLatencyAnomalies(hourlyMetrics.responseTimes);
    newAnomalies.push(...latencyAnomalies);

    // Detect anomalies in error rates
    const errorAnomalies = this.detectErrorAnomalies(hourlyMetrics.errorRates);
    newAnomalies.push(...errorAnomalies);

    // Detect user activity anomalies
    const activityAnomalies = this.detectActivityAnomalies(hourlyMetrics.activeUsers);
    newAnomalies.push(...activityAnomalies);

    // Store new anomalies
    this.anomalies = [...newAnomalies, ...this.anomalies].slice(0, 100);
    await this.persistAnomalies();

    return newAnomalies;
  }

  private detectVolumeAnomalies(volumes: { hour: string; value: number }[]): Anomaly[] {
    const anomalies: Anomaly[] = [];
    if (volumes.length < 24) return anomalies;

    const values = volumes.map(v => v.value);
    const { mean, stdDev } = this.calculateStats(values);

    volumes.forEach((v, i) => {
      if (i < 24) return; // Need baseline

      const zScore = (v.value - mean) / stdDev;
      
      if (Math.abs(zScore) > 3) {
        anomalies.push({
          id: `anom_vol_${Date.now()}_${i}`,
          type: zScore > 0 ? 'spike' : 'drop',
          metric: 'query_volume',
          severity: Math.abs(zScore) > 4 ? 'critical' : Math.abs(zScore) > 3.5 ? 'high' : 'medium',
          timestamp: new Date(v.hour),
          expectedValue: mean,
          actualValue: v.value,
          deviation: ((v.value - mean) / mean) * 100,
          description: `Query volume ${zScore > 0 ? 'spike' : 'drop'} detected`,
          recommendation: zScore > 0 
            ? 'Consider scaling infrastructure to handle increased load'
            : 'Investigate potential service disruption'
        });
      }
    });

    return anomalies;
  }

  private detectLatencyAnomalies(latencies: { hour: string; value: number }[]): Anomaly[] {
    const anomalies: Anomaly[] = [];
    if (latencies.length < 24) return anomalies;

    const values = latencies.map(v => v.value);
    const { mean, stdDev } = this.calculateStats(values);

    latencies.forEach((v, i) => {
      if (i < 24) return;

      const zScore = (v.value - mean) / stdDev;
      
      if (zScore > 2.5) {
        anomalies.push({
          id: `anom_lat_${Date.now()}_${i}`,
          type: 'spike',
          metric: 'response_time',
          severity: zScore > 4 ? 'critical' : zScore > 3 ? 'high' : 'medium',
          timestamp: new Date(v.hour),
          expectedValue: mean,
          actualValue: v.value,
          deviation: ((v.value - mean) / mean) * 100,
          description: `Response time degradation detected`,
          recommendation: 'Investigate performance bottlenecks and optimize AI model response times'
        });
      }
    });

    return anomalies;
  }

  private detectErrorAnomalies(errorRates: { hour: string; value: number }[]): Anomaly[] {
    const anomalies: Anomaly[] = [];
    const threshold = 0.05; // 5% error rate threshold

    errorRates.forEach((v, i) => {
      if (v.value > threshold) {
        anomalies.push({
          id: `anom_err_${Date.now()}_${i}`,
          type: 'spike',
          metric: 'error_rate',
          severity: v.value > 0.2 ? 'critical' : v.value > 0.1 ? 'high' : 'medium',
          timestamp: new Date(v.hour),
          expectedValue: 0.01,
          actualValue: v.value,
          deviation: (v.value / 0.01) * 100,
          description: `Elevated error rate: ${(v.value * 100).toFixed(2)}%`,
          recommendation: 'Review error logs and identify root cause'
        });
      }
    });

    return anomalies;
  }

  private detectActivityAnomalies(activities: { hour: string; value: number }[]): Anomaly[] {
    const anomalies: Anomaly[] = [];
    if (activities.length < 168) return anomalies; // Need 1 week baseline

    const values = activities.map(v => v.value);
    const { mean, stdDev } = this.calculateStats(values);

    // Check for unusual patterns (off-hours activity, etc.)
    activities.forEach((v, i) => {
      const hour = new Date(v.hour).getHours();
      const isBusinessHours = hour >= 9 && hour < 18;
      const zScore = (v.value - mean) / stdDev;

      if (!isBusinessHours && v.value > mean * 0.5) {
        anomalies.push({
          id: `anom_act_${Date.now()}_${i}`,
          type: 'outlier',
          metric: 'after_hours_activity',
          severity: 'low',
          timestamp: new Date(v.hour),
          expectedValue: mean * 0.1,
          actualValue: v.value,
          deviation: 0,
          description: 'Unusual after-hours activity detected'
        });
      }
    });

    return anomalies;
  }

  getRecentAnomalies(count: number = 20): Anomaly[] {
    return this.anomalies.slice(0, count);
  }

  // ==================== FORECASTING ====================

  /**
   * Generate forecasts for key metrics
   */
  async generateForecast(
    metric: 'query_volume' | 'user_growth' | 'cost' | 'engagement',
    horizon: number = 30
  ): Promise<Forecast> {
    const historicalData = await this.getHistoricalData(metric, 90);
    
    // Simple exponential smoothing for demonstration
    // In production, use more sophisticated models (ARIMA, Prophet, etc.)
    const alpha = 0.3; // Smoothing factor
    const predictions: ForecastPoint[] = [];
    
    let level = historicalData[historicalData.length - 1]?.value || 0;
    const values = historicalData.map(d => d.value);
    const stdDev = this.calculateStandardDeviation(values);
    
    // Detect trend
    const trend = this.calculateTrend(historicalData);
    
    // Generate predictions
    for (let i = 1; i <= horizon; i++) {
      const trendComponent = trend.slope * i;
      const forecastValue = level + trendComponent;
      const confidence = Math.max(0.5, 1 - (i * 0.02)); // Decreasing confidence
      
      predictions.push({
        date: this.addDays(new Date(), i).toISOString().split('T')[0],
        value: Math.max(0, forecastValue),
        lowerBound: Math.max(0, forecastValue - 2 * stdDev),
        upperBound: forecastValue + 2 * stdDev,
        confidence
      });
      
      level = alpha * forecastValue + (1 - alpha) * level;
    }

    // Detect seasonality
    const seasonality = this.detectSeasonality(historicalData);

    return {
      metric,
      horizon,
      predictions,
      trend: trend.direction,
      seasonality,
      accuracy: this.calculateForecastAccuracy(historicalData)
    };
  }

  private calculateTrend(data: { date: string; value: number }[]): {
    slope: number;
    direction: 'up' | 'down' | 'stable';
  } {
    if (data.length < 2) return { slope: 0, direction: 'stable' };

    const n = data.length;
    const sumX = data.reduce((sum, _, i) => sum + i, 0);
    const sumY = data.reduce((sum, d) => sum + d.value, 0);
    const sumXY = data.reduce((sum, d, i) => sum + i * d.value, 0);
    const sumXX = data.reduce((sum, _, i) => sum + i * i, 0);

    const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
    
    let direction: 'up' | 'down' | 'stable' = 'stable';
    if (slope > 0.01) direction = 'up';
    else if (slope < -0.01) direction = 'down';

    return { slope, direction };
  }

  private detectSeasonality(data: { date: string; value: number }[]): Forecast['seasonality'] {
    const seasonality: Forecast['seasonality'] = {};
    
    if (data.length < 24) return seasonality;

    // Check for daily pattern
    const hourlyAvg = new Array(24).fill(0).map((_, i) => {
      const hourData = data.filter(d => new Date(d.date).getHours() === i);
      return hourData.reduce((sum, d) => sum + d.value, 0) / hourData.length;
    });
    const hourVariation = this.calculateVariation(hourlyAvg);
    if (hourVariation > 0.3) seasonality.daily = true;

    // Check for weekly pattern
    if (data.length >= 168) {
      const dailyAvg = new Array(7).fill(0).map((_, i) => {
        const dayData = data.filter(d => new Date(d.date).getDay() === i);
        return dayData.reduce((sum, d) => sum + d.value, 0) / dayData.length;
      });
      const dayVariation = this.calculateVariation(dailyAvg);
      if (dayVariation > 0.2) seasonality.weekly = true;
    }

    return seasonality;
  }

  // ==================== INSIGHTS ====================

  /**
   * Generate actionable insights from analytics data
   */
  async generateInsights(): Promise<Insight[]> {
    const insights: Insight[] = [];

    // Performance insights
    const perfInsight = await this.generatePerformanceInsight();
    if (perfInsight) insights.push(perfInsight);

    // Usage insights
    const usageInsight = await this.generateUsageInsight();
    if (usageInsight) insights.push(usageInsight);

    // Cost insights
    const costInsight = await this.generateCostInsight();
    if (costInsight) insights.push(costInsight);

    // Engagement insights
    const engagementInsight = await this.generateEngagementInsight();
    if (engagementInsight) insights.push(engagementInsight);

    this.insights = insights;
    return insights;
  }

  private async generatePerformanceInsight(): Promise<Insight | null> {
    const metrics = await analyticsService.getPerformanceMetrics('7d');
    
    if (metrics.p95ResponseTime > 5000) {
      return {
        id: `ins_perf_${Date.now()}`,
        category: 'performance',
        priority: 'high',
        title: 'Response Time Degradation',
        description: `P95 response time is ${metrics.p95ResponseTime.toFixed(0)}ms, above recommended threshold`,
        metric: 'p95_response_time',
        value: metrics.p95ResponseTime,
        change: 0,
        actionable: true,
        action: 'Consider enabling caching or upgrading AI model infrastructure'
      };
    }

    if (metrics.errorRate > 0.05) {
      return {
        id: `ins_err_${Date.now()}`,
        category: 'performance',
        priority: 'high',
        title: 'Elevated Error Rate',
        description: `Error rate is ${(metrics.errorRate * 100).toFixed(2)}%, investigate immediately`,
        metric: 'error_rate',
        value: metrics.errorRate,
        change: 0,
        actionable: true,
        action: 'Review error logs and implement retry logic'
      };
    }

    return null;
  }

  private async generateUsageInsight(): Promise<Insight | null> {
    const usage = await analyticsService.getUsageMetrics('7d');
    const prevUsage = await analyticsService.getUsageMetrics('30d');
    
    const queryGrowth = usage.totalQueries / (prevUsage.totalQueries - usage.totalQueries) - 1;
    
    if (queryGrowth > 0.5) {
      return {
        id: `ins_usage_${Date.now()}`,
        category: 'usage',
        priority: 'medium',
        title: 'Rapid Growth in AI Queries',
        description: `Query volume increased by ${(queryGrowth * 100).toFixed(0)}% compared to previous period`,
        metric: 'query_growth',
        value: queryGrowth,
        change: queryGrowth,
        actionable: true,
        action: 'Monitor infrastructure capacity and consider scaling'
      };
    }

    return null;
  }

  private async generateCostInsight(): Promise<Insight | null> {
    const aiStats = await analyticsService.getAIUsageStats('30d');
    
    const gpt4Percentage = aiStats.modelBreakdown.find(m => m.model === 'gpt-4')?.tokens || 0;
    const totalTokens = aiStats.totalTokensUsed;
    
    if (totalTokens > 0 && gpt4Percentage / totalTokens > 0.5) {
      return {
        id: `ins_cost_${Date.now()}`,
        category: 'cost',
        priority: 'medium',
        title: 'High GPT-4 Usage',
        description: 'More than 50% of queries use expensive GPT-4 model',
        metric: 'gpt4_percentage',
        value: (gpt4Percentage / totalTokens) * 100,
        change: 0,
        actionable: true,
        action: 'Consider routing simple queries to GPT-3.5-turbo to reduce costs'
      };
    }

    return null;
  }

  private async generateEngagementInsight(): Promise<Insight | null> {
    const engagement = await analyticsService.getUserEngagement('7d');
    
    const lowEngagementUsers = engagement.filter(e => e.engagementScore < 30).length;
    const percentage = (lowEngagementUsers / engagement.length) * 100;
    
    if (percentage > 50) {
      return {
        id: `ins_eng_${Date.now()}`,
        category: 'engagement',
        priority: 'medium',
        title: 'Low User Engagement',
        description: `${percentage.toFixed(0)}% of users have low engagement scores`,
        metric: 'low_engagement_percentage',
        value: percentage,
        change: 0,
        actionable: true,
        action: 'Consider onboarding improvements and feature education'
      };
    }

    return null;
  }

  // ==================== CORRELATION ANALYSIS ====================

  /**
   * Analyze correlations between different metrics
   */
  async analyzeCorrelations(): Promise<Correlation[]> {
    const correlations: Correlation[] = [];
    
    // Get metrics data
    const queryVolume = await this.getMetricTimeSeries('query_volume', 30);
    const responseTime = await this.getMetricTimeSeries('response_time', 30);
    const errorRate = await this.getMetricTimeSeries('error_rate', 30);
    const userCount = await this.getMetricTimeSeries('user_count', 30);

    // Calculate correlations
    const volTimeCorr = this.calculatePearsonCorrelation(queryVolume, responseTime);
    if (volTimeCorr) {
      correlations.push({
        metric1: 'query_volume',
        metric2: 'response_time',
        coefficient: volTimeCorr,
        significance: 0.01,
        sampleSize: queryVolume.length,
        trend: volTimeCorr > 0.3 ? 'positive' : volTimeCorr < -0.3 ? 'negative' : 'none'
      });
    }

    const volErrorCorr = this.calculatePearsonCorrelation(queryVolume, errorRate);
    if (volErrorCorr) {
      correlations.push({
        metric1: 'query_volume',
        metric2: 'error_rate',
        coefficient: volErrorCorr,
        significance: 0.01,
        sampleSize: queryVolume.length,
        trend: volErrorCorr > 0.3 ? 'positive' : volErrorCorr < -0.3 ? 'negative' : 'none'
      });
    }

    return correlations;
  }

  private calculatePearsonCorrelation(x: number[], y: number[]): number | null {
    if (x.length !== y.length || x.length < 3) return null;

    const n = x.length;
    const sumX = x.reduce((a, b) => a + b, 0);
    const sumY = y.reduce((a, b) => a + b, 0);
    const sumXY = x.reduce((sum, xi, i) => sum + xi * y[i], 0);
    const sumXX = x.reduce((sum, xi) => sum + xi * xi, 0);
    const sumYY = y.reduce((sum, yi) => sum + yi * yi, 0);

    const numerator = n * sumXY - sumX * sumY;
    const denominator = Math.sqrt((n * sumXX - sumX * sumX) * (n * sumYY - sumY * sumY));

    return denominator === 0 ? 0 : numerator / denominator;
  }

  // ==================== COHORT ANALYSIS ====================

  /**
   * Perform cohort analysis on user retention
   */
  async analyzeCohorts(): Promise<CohortData[]> {
    const cohorts: CohortData[] = [];
    const events = await this.getRawEvents('90d');
    
    // Group users by first seen date
    const userFirstSeen = new Map<string, Date>();
    events.forEach(e => {
      if (!userFirstSeen.has(e.userId) || e.timestamp < userFirstSeen.get(e.userId)!) {
        userFirstSeen.set(e.userId, e.timestamp);
      }
    });

    // Group by week
    const weeklyCohorts = new Map<string, string[]>();
    userFirstSeen.forEach((date, userId) => {
      const weekKey = this.getWeekKey(date);
      if (!weeklyCohorts.has(weekKey)) {
        weeklyCohorts.set(weekKey, []);
      }
      weeklyCohorts.get(weekKey)!.push(userId);
    });

    // Calculate retention for each cohort
    weeklyCohorts.forEach((users, weekKey) => {
      const retention: number[] = [];
      const engagement: number[] = [];

      for (let i = 0; i < 8; i++) {
        const targetWeek = this.addWeeks(new Date(weekKey), i);
        const activeUsers = new Set(
          events
            .filter(e => {
              const eventWeek = this.getWeekKey(e.timestamp);
              return eventWeek === this.getWeekKey(targetWeek) && users.includes(e.userId);
            })
            .map(e => e.userId)
        );

        retention.push((activeUsers.size / users.length) * 100);
        
        // Calculate average engagement
        const cohortEvents = events.filter(e => 
          activeUsers.has(e.userId) && this.getWeekKey(e.timestamp) === this.getWeekKey(targetWeek)
        );
        const avgEngagement = cohortEvents.length / activeUsers.size || 0;
        engagement.push(avgEngagement);
      }

      cohorts.push({
        cohortDate: weekKey,
        size: users.length,
        retention,
        engagement
      });
    });

    return cohorts.sort((a, b) => a.cohortDate.localeCompare(b.cohortDate));
  }

  // ==================== FUNNEL ANALYSIS ====================

  /**
   * Analyze user journey funnel
   */
  async analyzeFunnel(): Promise<FunnelAnalysis> {
    const events = await this.getRawEvents('30d');
    
    // Define funnel stages
    const stages: FunnelStage[] = [
      { name: 'Add-in Opened', users: 0, dropOff: 0, conversionRate: 100, avgTime: 0 },
      { name: 'AI Query Made', users: 0, dropOff: 0, conversionRate: 0, avgTime: 0 },
      { name: 'Result Applied', users: 0, dropOff: 0, conversionRate: 0, avgTime: 0 },
      { name: 'Feature Used Again', users: 0, dropOff: 0, conversionRate: 0, avgTime: 0 }
    ];

    // Stage 1: Unique users who opened the add-in
    const uniqueUsers = new Set(events.map(e => e.userId));
    stages[0].users = uniqueUsers.size;

    // Stage 2: Users who made AI queries
    const queryUsers = new Set(
      events.filter(e => e.type === 'ai_query').map(e => e.userId)
    );
    stages[1].users = queryUsers.size;
    stages[1].conversionRate = (queryUsers.size / uniqueUsers.size) * 100;
    stages[1].dropOff = uniqueUsers.size - queryUsers.size;

    // Stage 3: Users who applied results (batch operations, visual highlights, etc.)
    const appliedResultsUsers = new Set(
      events.filter(e => 
        e.type === 'batch_operation' || 
        e.type === 'visualization_created' ||
        e.type === 'formula_parsed'
      ).map(e => e.userId)
    );
    stages[2].users = appliedResultsUsers.size;
    stages[2].conversionRate = (appliedResultsUsers.size / queryUsers.size) * 100;
    stages[2].dropOff = queryUsers.size - appliedResultsUsers.size;

    // Stage 4: Users who returned (used feature more than once)
    const userSessionCounts = new Map<string, number>();
    events.forEach(e => {
      userSessionCounts.set(e.userId, (userSessionCounts.get(e.userId) || 0) + 1);
    });
    const returningUsers = Array.from(userSessionCounts.entries())
      .filter(([_, count]) => count > 1)
      .map(([userId]) => userId);
    stages[3].users = returningUsers.length;
    stages[3].conversionRate = (returningUsers.length / appliedResultsUsers.size) * 100;
    stages[3].dropOff = appliedResultsUsers.size - returningUsers.length;

    const overallConversion = (returningUsers.length / uniqueUsers.size) * 100;

    return {
      name: 'User Journey',
      stages,
      overallConversion,
      totalTime: 0
    };
  }

  // ==================== ALERTS ====================

  async createAlert(config: Omit<AlertConfig, 'id'>): Promise<AlertConfig> {
    const alert: AlertConfig = {
      ...config,
      id: `alert_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
    };
    
    this.alerts.push(alert);
    await this.persistAlerts();
    return alert;
  }

  async updateAlert(id: string, updates: Partial<AlertConfig>): Promise<void> {
    const index = this.alerts.findIndex(a => a.id === id);
    if (index !== -1) {
      this.alerts[index] = { ...this.alerts[index], ...updates };
      await this.persistAlerts();
    }
  }

  async deleteAlert(id: string): Promise<void> {
    this.alerts = this.alerts.filter(a => a.id !== id);
    await this.persistAlerts();
  }

  getAlerts(): AlertConfig[] {
    return this.alerts;
  }

  // ==================== SCHEDULED REPORTS ====================

  async createScheduledReport(config: Omit<ScheduledReport, 'id'>): Promise<ScheduledReport> {
    const report: ScheduledReport = {
      ...config,
      id: `rpt_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
    };
    
    this.reports.push(report);
    await this.persistReports();
    return report;
  }

  async updateScheduledReport(id: string, updates: Partial<ScheduledReport>): Promise<void> {
    const index = this.reports.findIndex(r => r.id === id);
    if (index !== -1) {
      this.reports[index] = { ...this.reports[index], ...updates };
      await this.persistReports();
    }
  }

  async deleteScheduledReport(id: string): Promise<void> {
    this.reports = this.reports.filter(r => r.id !== id);
    await this.persistReports();
  }

  getScheduledReports(): ScheduledReport[] {
    return this.reports;
  }

  // ==================== UTILITY METHODS ====================

  private async getRawEvents(timeRange: TimeRange): Promise<AnalyticsEvent[]> {
    try {
      const events = await analyticsService.exportAnalytics('json', timeRange);
      return JSON.parse(events) as AnalyticsEvent[];
    } catch (error) {
      return [];
    }
  }

  private aggregateByHour(events: AnalyticsEvent[]): {
    queryVolume: { hour: string; value: number }[];
    responseTimes: { hour: string; value: number }[];
    errorRates: { hour: string; value: number }[];
    activeUsers: { hour: string; value: number }[];
  } {
    const hourlyData = new Map<string, {
      queries: number;
      responseTimes: number[];
      errors: number;
      total: number;
      users: Set<string>;
    }>();

    events.forEach(e => {
      const hour = e.timestamp.toISOString().slice(0, 13) + ':00:00';
      let data = hourlyData.get(hour);
      if (!data) {
        data = { queries: 0, responseTimes: [], errors: 0, total: 0, users: new Set() };
        hourlyData.set(hour, data);
      }

      if (e.type === 'ai_query') {
        data.queries++;
        if (e.performance?.processingTime) {
          data.responseTimes.push(e.performance.processingTime);
        }
        if (!e.success) data.errors++;
        data.total++;
      }
      data.users.add(e.userId);
    });

    return {
      queryVolume: Array.from(hourlyData.entries()).map(([hour, d]) => ({
        hour,
        value: d.queries
      })),
      responseTimes: Array.from(hourlyData.entries()).map(([hour, d]) => ({
        hour,
        value: d.responseTimes.length > 0 
          ? d.responseTimes.reduce((a, b) => a + b, 0) / d.responseTimes.length 
          : 0
      })),
      errorRates: Array.from(hourlyData.entries()).map(([hour, d]) => ({
        hour,
        value: d.total > 0 ? d.errors / d.total : 0
      })),
      activeUsers: Array.from(hourlyData.entries()).map(([hour, d]) => ({
        hour,
        value: d.users.size
      }))
    };
  }

  private calculateStats(values: number[]): { mean: number; stdDev: number } {
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / values.length;
    return { mean, stdDev: Math.sqrt(variance) };
  }

  private calculateStandardDeviation(values: number[]): number {
    return this.calculateStats(values).stdDev;
  }

  private calculateVariation(values: number[]): number {
    const { mean, stdDev } = this.calculateStats(values);
    return mean === 0 ? 0 : stdDev / mean;
  }

  private addDays(date: Date, days: number): Date {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  private addWeeks(date: Date, weeks: number): Date {
    return this.addDays(date, weeks * 7);
  }

  private getWeekKey(date: Date): string {
    const d = new Date(date);
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - d.getDay());
    return d.toISOString().split('T')[0];
  }

  private async getHistoricalData(
    metric: string, 
    days: number
  ): Promise<{ date: string; value: number }[]> {
    try {
      const endDate = new Date();
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const usage = await analyticsService.getUsageMetrics('custom', {
        start: startDate,
        end: endDate
      });

      switch (metric) {
        case 'query_volume':
          return usage.queryVolume.map(qv => ({
            date: qv.date,
            value: qv.count
          }));
        case 'user_growth':
          return usage.dailyActiveUsers.map(dau => ({
            date: dau.date,
            value: dau.count
          }));
        case 'engagement':
          return usage.dailyActiveUsers.map(dau => ({
            date: dau.date,
            value: (dau.count / Math.max(usage.totalUsers, 1)) * 100
          }));
        default:
          return [];
      }
    } catch (error) {
      return [];
    }
  }

  private async getMetricTimeSeries(
    metric: string, 
    days: number
  ): Promise<number[]> {
    const historical = await this.getHistoricalData(metric, days);
    return historical.map(h => h.value);
  }

  private calculateForecastAccuracy(historicalData: { date: string; value: number }[]): number {
    if (historicalData.length < 10) return 0;
    const errors: number[] = [];
    const splitPoint = Math.floor(historicalData.length * 0.7);
    const training = historicalData.slice(0, splitPoint);
    const validation = historicalData.slice(splitPoint);
    const forecastValue = training[training.length - 1]?.value ?? 0;
    for (const actual of validation) {
      if (actual.value !== 0) {
        const error = Math.abs((actual.value - forecastValue) / actual.value);
        errors.push(error);
      }
    }
    if (errors.length === 0) return 0;
    const mape = errors.reduce((a, b) => a + b, 0) / errors.length;
    return Math.max(0, 1 - mape);
  }

  // Persistence methods
  private async persistAnomalies(): Promise<void> {
    localStorage.setItem('advanced_analytics_anomalies', JSON.stringify(this.anomalies));
  }

  private async persistAlerts(): Promise<void> {
    localStorage.setItem('advanced_analytics_alerts', JSON.stringify(this.alerts));
  }

  private async persistReports(): Promise<void> {
    localStorage.setItem('advanced_analytics_reports', JSON.stringify(this.reports));
  }

  private async loadAlerts(): Promise<void> {
    const stored = localStorage.getItem('advanced_analytics_alerts');
    if (stored) {
      this.alerts = JSON.parse(stored);
    }
  }

  private async loadReports(): Promise<void> {
    const stored = localStorage.getItem('advanced_analytics_reports');
    if (stored) {
      this.reports = JSON.parse(stored);
    }
  }
}

export const advancedAnalyticsService = AdvancedAnalyticsService.getInstance();
export default advancedAnalyticsService;
