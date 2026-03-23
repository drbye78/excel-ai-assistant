/**
 * Advanced Analytics Dashboard Component
 * Week 16: Advanced Analytics
 * 
 * Enterprise-grade analytics with anomaly detection,
 * forecasting, insights, and advanced visualizations.
 */

import React, { useState, useEffect, useCallback } from 'react';
import { logger } from '../utils/logger';
import {
  advancedAnalyticsService,
  Anomaly,
  Forecast,
  Insight,
  Correlation,
  CohortData,
  FunnelAnalysis,
  AlertConfig,
  ScheduledReport
} from '../services/advancedAnalytics';
import { TimeRange } from '../services/analyticsService';

interface AnomalyCardProps {
  anomaly: Anomaly;
}

const AnomalyCard: React.FC<AnomalyCardProps> = ({ anomaly }) => {
  const severityColors = {
    low: '#107c10',
    medium: '#ff8c00',
    high: '#d13438',
    critical: '#8b0000'
  };

  const typeIcons = {
    spike: '📈',
    drop: '📉',
    trend_change: '📊',
    outlier: '⚠️'
  };

  return (
    <div className="anomaly-card" style={{ borderLeftColor: severityColors[anomaly.severity] }}>
      <div className="anomaly-header">
        <span className="anomaly-icon">{typeIcons[anomaly.type]}</span>
        <span className={`severity-badge ${anomaly.severity}`}>{anomaly.severity}</span>
        <span className="anomaly-time">{new Date(anomaly.timestamp).toLocaleString()}</span>
      </div>
      <h4 className="anomaly-title">{anomaly.description}</h4>
        <div className="anomaly-metrics">
          <div className="metric">
            <label>Expected:</label>
            <span className="metric-value">{anomaly.expectedValue.toFixed(2)}</span>
          </div>
          <div className="metric">
            <label>Actual:</label>
            <span className="metric-value">{anomaly.actualValue.toFixed(2)}</span>
          </div>
          <div className={`metric deviation ${anomaly.deviation > 0 ? 'positive' : 'negative'}`}>
            <label>Deviation:</label>
            <span className="metric-value">{anomaly.deviation > 0 ? '+' : ''}{anomaly.deviation.toFixed(1)}%</span>
          </div>
        </div>
      {anomaly.recommendation && (
        <div className="recommendation">
          <strong>💡 Recommendation:</strong> {anomaly.recommendation}
        </div>
      )}
    </div>
  );
};

interface InsightCardProps {
  insight: Insight;
}

const InsightCard: React.FC<InsightCardProps> = ({ insight }) => {
  const priorityColors = {
    low: '#107c10',
    medium: '#ff8c00',
    high: '#d13438'
  };

  const categoryIcons = {
    performance: '⚡',
    usage: '📊',
    cost: '💰',
    engagement: '👥',
    security: '🔒'
  };

  return (
    <div className="insight-card" style={{ borderLeftColor: priorityColors[insight.priority] }}>
      <div className="insight-header">
        <span className="insight-icon">{categoryIcons[insight.category]}</span>
        <span className={`priority-badge ${insight.priority}`}>{insight.priority}</span>
        {insight.actionable && <span className="actionable-badge">Actionable</span>}
      </div>
      <h4 className="insight-title">{insight.title}</h4>
      <p className="insight-description">{insight.description}</p>
      <div className="insight-metric">
        <span className="metric-name">{insight.metric}:</span>
        <span className="metric-value">{insight.value.toFixed(2)}</span>
        {insight.change !== 0 && (
          <span className={`metric-change ${insight.change > 0 ? 'positive' : 'negative'}`}>
            {insight.change > 0 ? '↑' : '↓'} {Math.abs(insight.change * 100).toFixed(1)}%
          </span>
        )}
      </div>
      {insight.action && (
        <div className="suggested-action">
          <strong>🎯 Action:</strong> {insight.action}
        </div>
      )}
    </div>
  );
};

interface ForecastChartProps {
  forecast: Forecast;
}

const ForecastChart: React.FC<ForecastChartProps> = ({ forecast }) => {
  if (!forecast.predictions.length) return <div className="no-data">No forecast data available</div>;

  const maxValue = Math.max(...forecast.predictions.map(p => p.upperBound));
  const minValue = Math.min(...forecast.predictions.map(p => p.lowerBound));
  const range = maxValue - minValue || 1;

  const trendIcons = {
    up: '📈',
    down: '📉',
    stable: '➡️'
  };

  return (
    <div className="forecast-chart">
      <div className="forecast-header">
        <span className="forecast-trend">{trendIcons[forecast.trend]} {forecast.trend}</span>
        <span className="forecast-accuracy">Accuracy: {(forecast.accuracy * 100).toFixed(0)}%</span>
        {forecast.seasonality && (
          <div className="seasonality-tags">
            {forecast.seasonality.daily && <span className="tag">Daily</span>}
            {forecast.seasonality.weekly && <span className="tag">Weekly</span>}
            {forecast.seasonality.monthly && <span className="tag">Monthly</span>}
          </div>
        )}
      </div>
      
      <svg className="forecast-svg" viewBox={`0 0 ${forecast.predictions.length * 20} 100`} preserveAspectRatio="none">
        {/* Confidence band */}
        <path
          d={`M 0 ${50 - ((forecast.predictions[0].upperBound - minValue) / range) * 80}
            ${forecast.predictions.map((p, i) => `L ${i * 20} ${50 - ((p.upperBound - minValue) / range) * 80}`).join(' ')}
            L ${(forecast.predictions.length - 1) * 20} ${50 - ((forecast.predictions[forecast.predictions.length - 1].lowerBound - minValue) / range) * 80}
            ${forecast.predictions.slice().reverse().map((p, i) => `L ${(forecast.predictions.length - 1 - i) * 20} ${50 - ((p.lowerBound - minValue) / range) * 80}`).join(' ')}
            Z`}
          fill="#e8f4fd"
          opacity="0.5"
        />
        
        {/* Prediction line */}
        <polyline
          fill="none"
          stroke="#0078d4"
          strokeWidth="2"
          points={forecast.predictions.map((p, i) => `${i * 20},${50 - ((p.value - minValue) / range) * 80}`).join(' ')}
        />
        
        {/* Data points */}
        {forecast.predictions.map((p, i) => (
          <circle
            key={i}
            cx={i * 20}
            cy={50 - ((p.value - minValue) / range) * 80}
            r="3"
            fill="#0078d4"
          />
        ))}
      </svg>
      
      <div className="forecast-legend">
        <span className="legend-item prediction">Prediction</span>
        <span className="legend-item confidence">Confidence Band</span>
      </div>
    </div>
  );
};

interface FunnelVisualizationProps {
  funnel: FunnelAnalysis;
}

const FunnelVisualization: React.FC<FunnelVisualizationProps> = ({ funnel }) => {
  const maxUsers = funnel.stages[0]?.users || 1;

  return (
    <div className="funnel-visualization">
      <h4>{funnel.name}</h4>
      <div className="overall-conversion">
        Overall Conversion: {funnel.overallConversion.toFixed(1)}%
      </div>
      
      <div className="funnel-stages">
        {funnel.stages.map((stage, index) => {
          const width = (stage.users / maxUsers) * 100;
          return (
            <div key={index} className="funnel-stage">
              <div 
                className="stage-bar" 
                style={{ width: `${width}%` }}
              >
                <span className="stage-name">{stage.name}</span>
                <span className="stage-users">{stage.users}</span>
              </div>
              {index < funnel.stages.length - 1 && (
                <div className="stage-conversion">
                  ↓ {stage.conversionRate.toFixed(1)}%
                  <span className="dropoff">({stage.dropOff} dropped)</span>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

const AdvancedAnalyticsDashboard: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'anomalies' | 'insights' | 'forecast' | 'correlation' | 'cohorts' | 'funnel' | 'alerts'>('anomalies');
  const [timeRange, setTimeRange] = useState<TimeRange>('7d');
  const [isLoading, setIsLoading] = useState(false);

  // Data states
  const [anomalies, setAnomalies] = useState<Anomaly[]>([]);
  const [insights, setInsights] = useState<Insight[]>([]);
  const [forecast, setForecast] = useState<Forecast | null>(null);
  const [correlations, setCorrelations] = useState<Correlation[]>([]);
  const [cohorts, setCohorts] = useState<CohortData[]>([]);
  const [funnel, setFunnel] = useState<FunnelAnalysis | null>(null);
  const [alerts, setAlerts] = useState<AlertConfig[]>([]);
  const [reports, setReports] = useState<ScheduledReport[]>([]);

  const loadData = useCallback(async () => {
    setIsLoading(true);
    try {
      switch (activeTab) {
        case 'anomalies':
          const newAnomalies = await advancedAnalyticsService.detectAnomalies(timeRange);
          setAnomalies(newAnomalies);
          break;
        case 'insights':
          const newInsights = await advancedAnalyticsService.generateInsights();
          setInsights(newInsights);
          break;
        case 'forecast':
          const newForecast = await advancedAnalyticsService.generateForecast('query_volume', 30);
          setForecast(newForecast);
          break;
        case 'correlation':
          const newCorrelations = await advancedAnalyticsService.analyzeCorrelations();
          setCorrelations(newCorrelations);
          break;
        case 'cohorts':
          const newCohorts = await advancedAnalyticsService.analyzeCohorts();
          setCohorts(newCohorts);
          break;
        case 'funnel':
          const newFunnel = await advancedAnalyticsService.analyzeFunnel();
          setFunnel(newFunnel);
          break;
        case 'alerts':
          setAlerts(advancedAnalyticsService.getAlerts());
          setReports(advancedAnalyticsService.getScheduledReports());
          break;
      }
    } catch (error) {
      logger.error('Failed to load advanced analytics', undefined, error as Error);
    } finally {
      setIsLoading(false);
    }
  }, [activeTab, timeRange]);

  useEffect(() => {
    loadData();
  }, [loadData]);

  const handleCreateAlert = async () => {
    const name = prompt('Alert name:');
    if (!name) return;
    
    await advancedAnalyticsService.createAlert({
      name,
      metric: 'query_volume',
      condition: 'above',
      threshold: 1000,
      timeWindow: '24h',
      enabled: true,
      notificationChannels: ['in_app']
    });
    
    setAlerts(advancedAnalyticsService.getAlerts());
  };

  return (
    <div className="advanced-analytics-dashboard">
      <header className="dashboard-header">
        <h2>🔮 Advanced Analytics</h2>
        
        <div className="dashboard-controls">
          <select 
            value={timeRange} 
            onChange={(e) => setTimeRange(e.target.value as TimeRange)}
            className="time-range-select"
          >
            <option value="24h">Last 24 Hours</option>
            <option value="7d">Last 7 Days</option>
            <option value="30d">Last 30 Days</option>
            <option value="90d">Last 90 Days</option>
          </select>

          <button 
            className="refresh-btn"
            onClick={loadData}
            disabled={isLoading}
          >
            {isLoading ? '🔄' : '↻'} Refresh
          </button>
        </div>
      </header>

      <nav className="dashboard-tabs">
        {[
          { id: 'anomalies', label: 'Anomalies', icon: '⚠️' },
          { id: 'insights', label: 'Insights', icon: '💡' },
          { id: 'forecast', label: 'Forecast', icon: '🔮' },
          { id: 'correlation', label: 'Correlations', icon: '🔗' },
          { id: 'cohorts', label: 'Cohorts', icon: '👥' },
          { id: 'funnel', label: 'Funnel', icon: '📊' },
          { id: 'alerts', label: 'Alerts & Reports', icon: '🔔' }
        ].map(tab => (
          <button
            key={tab.id}
            className={`tab-btn ${activeTab === tab.id ? 'active' : ''}`}
            onClick={() => setActiveTab(tab.id as typeof activeTab)}
          >
            {tab.icon} {tab.label}
          </button>
        ))}
      </nav>

      <div className="dashboard-content">
        {isLoading ? (
          <div className="loading-state">
            <div className="spinner"></div>
            <p>Analyzing data...</p>
          </div>
        ) : (
          <>
            {activeTab === 'anomalies' && (
              <div className="anomalies-section">
                <h3>🚨 Detected Anomalies</h3>
                {anomalies.length === 0 ? (
                  <div className="no-anomalies">
                    <span className="check-icon">✅</span>
                    <p>No anomalies detected in the selected time range</p>
                  </div>
                ) : (
                  <div className="anomalies-list">
                    {anomalies.map(anomaly => (
                      <AnomalyCard key={anomaly.id} anomaly={anomaly} />
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === 'insights' && (
              <div className="insights-section">
                <h3>💡 AI-Generated Insights</h3>
                {insights.length === 0 ? (
                  <div className="no-insights">
                    <p>No insights available. Run analysis to generate insights.</p>
                  </div>
                ) : (
                  <div className="insights-list">
                    {insights.map(insight => (
                      <InsightCard key={insight.id} insight={insight} />
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === 'forecast' && (
              <div className="forecast-section">
                <h3>🔮 30-Day Forecast</h3>
                {forecast ? (
                  <ForecastChart forecast={forecast} />
                ) : (
                  <div className="no-data">No forecast data available</div>
                )}
              </div>
            )}

            {activeTab === 'correlation' && (
              <div className="correlation-section">
                <h3>🔗 Metric Correlations</h3>
                {correlations.length === 0 ? (
                  <div className="no-data">No correlation data available</div>
                ) : (
                  <div className="correlations-list">
                    {correlations.map((corr, idx) => (
                      <div key={idx} className="correlation-card">
                        <div className="correlation-metrics">
                          <span className="metric">{corr.metric1}</span>
                          <span className="arrow">↔️</span>
                          <span className="metric">{corr.metric2}</span>
                        </div>
                        <div className="correlation-strength">
                          <div 
                            className="strength-bar" 
                            style={{ width: `${Math.abs(corr.coefficient) * 100}%` }}
                          />
                          <span className="coefficient">{corr.coefficient.toFixed(3)}</span>
                        </div>
                        <span className={`trend ${corr.trend}`}>{corr.trend}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}

            {activeTab === 'cohorts' && (
              <div className="cohorts-section">
                <h3>👥 Cohort Analysis</h3>
                {cohorts.length === 0 ? (
                  <div className="no-data">No cohort data available</div>
                ) : (
                  <table className="cohorts-table">
                    <thead>
                      <tr>
                        <th>Cohort</th>
                        <th>Size</th>
                        {Array.from({ length: 8 }, (_, i) => (
                          <th key={i}>Week {i}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {cohorts.map(cohort => (
                        <tr key={cohort.cohortDate}>
                          <td>{cohort.cohortDate}</td>
                          <td>{cohort.size}</td>
                          {cohort.retention.map((ret, i) => (
                            <td 
                              key={i} 
                              className="retention-cell"
                              style={{ 
                                backgroundColor: `rgba(0, 120, 212, ${ret / 100})`,
                                color: ret > 50 ? 'white' : 'black'
                              }}
                            >
                              {ret.toFixed(0)}%
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>
            )}

            {activeTab === 'funnel' && (
              <div className="funnel-section">
                {funnel && <FunnelVisualization funnel={funnel} />}
              </div>
            )}

            {activeTab === 'alerts' && (
              <div className="alerts-section">
                <div className="alerts-header">
                  <h3>🔔 Configured Alerts</h3>
                  <button className="create-btn" onClick={handleCreateAlert}>
                    + Create Alert
                  </button>
                </div>
                
                {alerts.length === 0 ? (
                  <p className="no-alerts">No alerts configured</p>
                ) : (
                  <div className="alerts-list">
                    {alerts.map(alert => (
                      <div key={alert.id} className={`alert-card ${alert.enabled ? 'enabled' : 'disabled'}`}>
                        <div className="alert-header">
                          <span className="alert-name">{alert.name}</span>
                          <span className={`status-badge ${alert.enabled ? 'on' : 'off'}`}>
                            {alert.enabled ? 'ON' : 'OFF'}
                          </span>
                        </div>
                        <div className="alert-details">
                          <span>{alert.metric}</span>
                          <span>{alert.condition} {alert.threshold}</span>
                          <span>{alert.timeWindow}</span>
                        </div>
                        <div className="alert-channels">
                          {alert.notificationChannels.map(ch => (
                            <span key={ch} className="channel-tag">{ch}</span>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                <h3>📋 Scheduled Reports</h3>
                {reports.length === 0 ? (
                  <p className="no-reports">No scheduled reports</p>
                ) : (
                  <div className="reports-list">
                    {reports.map(report => (
                      <div key={report.id} className="report-card">
                        <div className="report-name">{report.name}</div>
                        <div className="report-details">
                          <span>{report.frequency}</span>
                          <span>{report.format.toUpperCase()}</span>
                          <span>Next: {new Date(report.nextRun).toLocaleDateString()}</span>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </div>

      <style>{`
        .advanced-analytics-dashboard {
          padding: 20px;
          background: #faf9f8;
          min-height: 100vh;
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        .dashboard-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 20px;
          padding-bottom: 15px;
          border-bottom: 1px solid #e1e1e1;
        }

        .dashboard-header h2 {
          margin: 0;
          color: #323130;
        }

        .dashboard-controls {
          display: flex;
          gap: 10px;
          align-items: center;
        }

        .time-range-select {
          padding: 8px 12px;
          border: 1px solid #8a8886;
          border-radius: 4px;
          font-size: 14px;
        }

        .refresh-btn, .create-btn {
          padding: 8px 16px;
          background: #0078d4;
          color: white;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        }

        .refresh-btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }

        .dashboard-tabs {
          display: flex;
          gap: 5px;
          margin-bottom: 20px;
          border-bottom: 2px solid #e1e1e1;
          flex-wrap: wrap;
        }

        .tab-btn {
          padding: 12px 20px;
          background: none;
          border: none;
          cursor: pointer;
          font-size: 14px;
          color: #605e5c;
          border-bottom: 2px solid transparent;
          margin-bottom: -2px;
          transition: all 0.2s;
        }

        .tab-btn:hover {
          color: #0078d4;
          background: #f3f2f1;
        }

        .tab-btn.active {
          color: #0078d4;
          border-bottom-color: #0078d4;
          font-weight: 600;
        }

        .dashboard-content {
          background: white;
          border-radius: 8px;
          padding: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }

        /* Anomalies */
        .anomaly-card {
          background: #fff;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 15px;
          border-left: 4px solid;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .anomaly-header {
          display: flex;
          align-items: center;
          gap: 10px;
          margin-bottom: 10px;
        }

        .anomaly-icon {
          font-size: 20px;
        }

        .severity-badge {
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
        }

        .severity-badge.low { background: #dff6dd; color: #107c10; }
        .severity-badge.medium { background: #fff4ce; color: #ff8c00; }
        .severity-badge.high { background: #fde7e9; color: #d13438; }
        .severity-badge.critical { background: #d13438; color: white; }

        .anomaly-time {
          color: #605e5c;
          font-size: 12px;
          margin-left: auto;
        }

        .anomaly-title {
          margin: 0 0 10px 0;
          font-size: 15px;
        }

        .anomaly-metrics {
          display: flex;
          gap: 20px;
          margin-bottom: 10px;
        }

        .anomaly-metrics .metric {
          display: flex;
          flex-direction: column;
        }

        .anomaly-metrics label {
          font-size: 11px;
          color: #605e5c;
        }

        .anomaly-metrics value {
          font-weight: 600;
        }

        .deviation.positive { color: #d13438; }
        .deviation.negative { color: #107c10; }

        .recommendation {
          background: #f3f2f1;
          padding: 10px;
          border-radius: 4px;
          font-size: 13px;
        }

        .no-anomalies {
          text-align: center;
          padding: 60px;
        }

        .check-icon {
          font-size: 48px;
          display: block;
          margin-bottom: 15px;
        }

        /* Insights */
        .insight-card {
          background: #fff;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 15px;
          border-left: 4px solid;
          box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }

        .insight-header {
          display: flex;
          align-items: center;
          gap: 10px;
          margin-bottom: 10px;
        }

        .insight-icon {
          font-size: 20px;
        }

        .priority-badge {
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
          font-weight: 600;
          text-transform: uppercase;
        }

        .priority-badge.low { background: #dff6dd; color: #107c10; }
        .priority-badge.medium { background: #fff4ce; color: #ff8c00; }
        .priority-badge.high { background: #fde7e9; color: #d13438; }

        .actionable-badge {
          background: #0078d4;
          color: white;
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
          margin-left: auto;
        }

        .insight-title {
          margin: 0 0 8px 0;
          font-size: 15px;
        }

        .insight-description {
          color: #605e5c;
          margin-bottom: 10px;
          font-size: 13px;
        }

        .insight-metric {
          display: flex;
          gap: 10px;
          align-items: center;
          margin-bottom: 10px;
        }

        .metric-change.positive { color: #107c10; }
        .metric-change.negative { color: #d13438; }

        .suggested-action {
          background: #e8f4fd;
          padding: 10px;
          border-radius: 4px;
          font-size: 13px;
        }

        /* Forecast */
        .forecast-chart {
          padding: 20px;
        }

        .forecast-header {
          display: flex;
          gap: 15px;
          align-items: center;
          margin-bottom: 20px;
        }

        .forecast-trend {
          font-weight: 600;
        }

        .forecast-accuracy {
          color: #605e5c;
        }

        .seasonality-tags {
          display: flex;
          gap: 5px;
        }

        .tag {
          background: #f3f2f1;
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
        }

        .forecast-svg {
          width: 100%;
          height: 200px;
        }

        .forecast-legend {
          display: flex;
          gap: 20px;
          margin-top: 15px;
          font-size: 12px;
        }

        .legend-item::before {
          content: '';
          display: inline-block;
          width: 12px;
          height: 12px;
          margin-right: 5px;
          border-radius: 2px;
        }

        .legend-item.prediction::before {
          background: #0078d4;
        }

        .legend-item.confidence::before {
          background: #e8f4fd;
        }

        /* Correlations */
        .correlation-card {
          background: #f9f9f9;
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 10px;
        }

        .correlation-metrics {
          display: flex;
          align-items: center;
          gap: 10px;
          margin-bottom: 10px;
        }

        .correlation-metrics .metric {
          font-weight: 600;
        }

        .correlation-strength {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .strength-bar {
          height: 8px;
          background: #0078d4;
          border-radius: 4px;
          transition: width 0.3s;
        }

        .trend.positive { color: #107c10; }
        .trend.negative { color: #d13438; }
        .trend.none { color: #605e5c; }

        /* Cohorts */
        .cohorts-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 13px;
        }

        .cohorts-table th, .cohorts-table td {
          padding: 8px;
          text-align: center;
          border: 1px solid #e1e1e1;
        }

        .cohorts-table th {
          background: #f3f2f1;
          font-weight: 600;
        }

        .retention-cell {
          transition: background-color 0.2s;
        }

        /* Funnel */
        .funnel-visualization {
          padding: 20px;
        }

        .overall-conversion {
          font-size: 24px;
          font-weight: 600;
          color: #0078d4;
          margin-bottom: 20px;
        }

        .funnel-stage {
          margin-bottom: 10px;
        }

        .stage-bar {
          background: linear-gradient(90deg, #0078d4, #5c9fd6);
          color: white;
          padding: 12px 15px;
          border-radius: 4px;
          display: flex;
          justify-content: space-between;
          align-items: center;
          transition: width 0.3s;
        }

        .stage-conversion {
          text-align: center;
          padding: 8px;
          color: #605e5c;
          font-size: 13px;
        }

        .dropoff {
          color: #d13438;
          margin-left: 10px;
        }

        /* Alerts */
        .alerts-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 20px;
        }

        .alert-card, .report-card {
          background: #f9f9f9;
          padding: 15px;
          border-radius: 8px;
          margin-bottom: 10px;
        }

        .alert-card.disabled {
          opacity: 0.6;
        }

        .alert-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 8px;
        }

        .alert-name {
          font-weight: 600;
        }

        .status-badge {
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
        }

        .status-badge.on {
          background: #dff6dd;
          color: #107c10;
        }

        .status-badge.off {
          background: #f3f2f1;
          color: #605e5c;
        }

        .alert-details, .report-details {
          display: flex;
          gap: 15px;
          color: #605e5c;
          font-size: 13px;
          margin-bottom: 8px;
        }

        .alert-channels {
          display: flex;
          gap: 5px;
        }

        .channel-tag {
          background: #e8f4fd;
          color: #0078d4;
          padding: 2px 8px;
          border-radius: 12px;
          font-size: 11px;
        }

        /* Loading */
        .loading-state {
          display: flex;
          flex-direction: column;
          align-items: center;
          padding: 60px;
          color: #605e5c;
        }

        .spinner {
          width: 40px;
          height: 40px;
          border: 3px solid #e1e1e1;
          border-top-color: #0078d4;
          border-radius: 50%;
          animation: spin 1s linear infinite;
          margin-bottom: 15px;
        }

        @keyframes spin {
          to { transform: rotate(360deg); }
        }

        .no-data, .no-alerts, .no-reports {
          text-align: center;
          padding: 40px;
          color: #605e5c;
        }
      `}</style>
    </div>
  );
};

export default AdvancedAnalyticsDashboard;
