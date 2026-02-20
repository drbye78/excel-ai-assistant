/**
 * Analytics Dashboard Component
 * Week 15: Analytics Dashboard
 * 
 * Enterprise analytics dashboard for tracking add-in usage,
 * AI query statistics, performance metrics, and user engagement.
 */

import React, { useState, useEffect, useCallback, useRef } from 'react';
import {
  analyticsService,
  AnalyticsEvent,
  UsageMetrics,
  PerformanceMetrics,
  UserEngagement,
  AIUsageStats,
  TimeRange,
  DateRange,
  SlowQuery
} from '../services/analyticsService';

interface MetricCardProps {
  title: string;
  value: string | number;
  subtitle?: string;
  trend?: number;
  icon: string;
  color: string;
}

const MetricCard: React.FC<MetricCardProps> = ({ title, value, subtitle, trend, icon, color }) => (
  <div className="metric-card" style={{ borderLeft: `4px solid ${color}` }}>
    <div className="metric-header">
      <span className="metric-icon">{icon}</span>
      <span className="metric-title">{title}</span>
    </div>
    <div className="metric-value" style={{ color }}>{value}</div>
    {subtitle && <div className="metric-subtitle">{subtitle}</div>}
    {trend !== undefined && (
      <div className={`metric-trend ${trend >= 0 ? 'positive' : 'negative'}`}>
        {trend >= 0 ? '↑' : '↓'} {Math.abs(trend).toFixed(1)}%
      </div>
    )}
  </div>
);

interface ChartProps {
  data: { label: string; value: number }[];
  type: 'bar' | 'line' | 'pie';
  color?: string;
}

const SimpleChart: React.FC<ChartProps> = ({ data, type, color = '#0078d4' }) => {
  if (data.length === 0) return <div className="no-data">No data available</div>;

  const maxValue = Math.max(...data.map(d => d.value));

  if (type === 'bar') {
    return (
      <div className="bar-chart">
        {data.map((item, index) => (
          <div key={index} className="bar-item">
            <div className="bar-label">{item.label}</div>
            <div className="bar-wrapper">
              <div 
                className="bar" 
                style={{ 
                  width: `${(item.value / maxValue) * 100}%`,
                  backgroundColor: color
                }}
              >
                <span className="bar-value">{item.value}</span>
              </div>
            </div>
          </div>
        ))}
      </div>
    );
  }

  if (type === 'line') {
    const points = data.map((d, i) => {
      const x = (i / (data.length - 1)) * 100;
      const y = 100 - ((d.value / maxValue) * 80);
      return `${x},${y}`;
    }).join(' ');

    return (
      <svg className="line-chart" viewBox="0 0 100 100" preserveAspectRatio="none">
        <polyline
          fill="none"
          stroke={color}
          strokeWidth="2"
          points={points}
        />
        {data.map((d, i) => {
          const x = (i / (data.length - 1)) * 100;
          const y = 100 - ((d.value / maxValue) * 80);
          return (
            <circle
              key={i}
              cx={x}
              cy={y}
              r="2"
              fill={color}
            />
          );
        })}
      </svg>
    );
  }

  return null;
};

interface TabPanelProps {
  children: React.ReactNode;
  isActive: boolean;
}

const TabPanel: React.FC<TabPanelProps> = ({ children, isActive }) => (
  <div className={`tab-panel ${isActive ? 'active' : ''}`}>
    {isActive && children}
  </div>
);

const AnalyticsDashboard: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'overview' | 'performance' | 'users' | 'ai' | 'realtime'>('overview');
  const [timeRange, setTimeRange] = useState<TimeRange>('7d');
  const [customRange, setCustomRange] = useState<DateRange | undefined>();
  const [isLoading, setIsLoading] = useState(false);

  // Metrics state
  const [usageMetrics, setUsageMetrics] = useState<UsageMetrics | null>(null);
  const [performanceMetrics, setPerformanceMetrics] = useState<PerformanceMetrics | null>(null);
  const [userEngagement, setUserEngagement] = useState<UserEngagement[]>([]);
  const [aiStats, setAiStats] = useState<AIUsageStats | null>(null);
  const [realtimeMetrics, setRealtimeMetrics] = useState(analyticsService.getRealTimeMetrics());
  const [recentEvents, setRecentEvents] = useState<AnalyticsEvent[]>([]);

  // Real-time updates
  const realtimeIntervalRef = useRef<NodeJS.Timeout | null>(null);

  const loadMetrics = useCallback(async () => {
    setIsLoading(true);
    try {
      const [usage, perf, users, ai] = await Promise.all([
        analyticsService.getUsageMetrics(timeRange, customRange),
        analyticsService.getPerformanceMetrics(timeRange, customRange),
        analyticsService.getUserEngagement(timeRange, customRange),
        analyticsService.getAIUsageStats(timeRange, customRange)
      ]);

      setUsageMetrics(usage);
      setPerformanceMetrics(perf);
      setUserEngagement(users);
      setAiStats(ai);
    } catch (error) {
      console.error('Failed to load metrics:', error);
    } finally {
      setIsLoading(false);
    }
  }, [timeRange, customRange]);

  useEffect(() => {
    loadMetrics();
  }, [loadMetrics]);

  // Real-time updates
  useEffect(() => {
    if (activeTab === 'realtime') {
      realtimeIntervalRef.current = setInterval(() => {
        setRealtimeMetrics(analyticsService.getRealTimeMetrics());
      }, 5000);
    }

    return () => {
      if (realtimeIntervalRef.current) {
        clearInterval(realtimeIntervalRef.current);
      }
    };
  }, [activeTab]);

  // Subscribe to events
  useEffect(() => {
    const unsubscribe = analyticsService.subscribeToEvent('ai_query', (event) => {
      setRecentEvents(prev => [event, ...prev].slice(0, 50));
    });

    return unsubscribe;
  }, []);

  const formatDuration = (ms: number): string => {
    if (ms < 1000) return `${ms.toFixed(0)}ms`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
    return `${(ms / 60000).toFixed(1)}m`;
  };

  const formatNumber = (num: number): string => {
    if (num >= 1000000) return `${(num / 1000000).toFixed(1)}M`;
    if (num >= 1000) return `${(num / 1000).toFixed(1)}K`;
    return num.toFixed(0);
  };

  const handleExport = async (format: 'json' | 'csv') => {
    try {
      const data = await analyticsService.exportAnalytics(format, timeRange, customRange);
      const blob = new Blob([data], { 
        type: format === 'json' ? 'application/json' : 'text/csv' 
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `analytics_${timeRange}.${format}`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error('Export failed:', error);
    }
  };

  return (
    <div className="analytics-dashboard">
      <header className="dashboard-header">
        <h2>📊 Analytics Dashboard</h2>
        
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
            <option value="1y">Last Year</option>
          </select>

          <button 
            className="refresh-btn"
            onClick={loadMetrics}
            disabled={isLoading}
          >
            {isLoading ? '🔄' : '↻'} Refresh
          </button>

          <div className="export-dropdown">
            <button className="export-btn">📥 Export</button>
            <div className="export-menu">
              <button onClick={() => handleExport('json')}>Export as JSON</button>
              <button onClick={() => handleExport('csv')}>Export as CSV</button>
            </div>
          </div>
        </div>
      </header>

      <nav className="dashboard-tabs">
        {[
          { id: 'overview', label: 'Overview', icon: '📈' },
          { id: 'performance', label: 'Performance', icon: '⚡' },
          { id: 'users', label: 'User Engagement', icon: '👥' },
          { id: 'ai', label: 'AI Usage', icon: '🤖' },
          { id: 'realtime', label: 'Real-time', icon: '🔴' }
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
            <p>Loading analytics...</p>
          </div>
        ) : (
          <>
            <TabPanel isActive={activeTab === 'overview'}>
              <div className="metrics-grid">
                <MetricCard
                  title="Total Users"
                  value={formatNumber(usageMetrics?.totalUsers || 0)}
                  subtitle="Active in selected period"
                  icon="👤"
                  color="#0078d4"
                />
                <MetricCard
                  title="Total Sessions"
                  value={formatNumber(usageMetrics?.totalSessions || 0)}
                  subtitle={`Avg ${formatDuration(usageMetrics?.avgSessionDuration || 0)} per session`}
                  icon="🌐"
                  color="#107c10"
                />
                <MetricCard
                  title="AI Queries"
                  value={formatNumber(usageMetrics?.totalQueries || 0)}
                  subtitle={`${((usageMetrics?.successRate || 0) * 100).toFixed(1)}% success rate`}
                  icon="🤖"
                  color="#8661c5"
                />
                <MetricCard
                  title="Avg Response Time"
                  value={formatDuration(performanceMetrics?.avgResponseTime || 0)}
                  subtitle={`P95: ${formatDuration(performanceMetrics?.p95ResponseTime || 0)}`}
                  icon="⚡"
                  color="#ff8c00"
                />
              </div>

              <div className="chart-section">
                <h3>📊 Daily Active Users</h3>
                <SimpleChart
                  data={(usageMetrics?.dailyActiveUsers || []).map(d => ({
                    label: d.date.slice(5),
                    value: d.count
                  }))}
                  type="line"
                  color="#0078d4"
                />
              </div>

              <div className="chart-section">
                <h3>🔥 Query Volume</h3>
                <SimpleChart
                  data={(usageMetrics?.queryVolume || []).map(d => ({
                    label: d.date.slice(5),
                    value: d.count
                  }))}
                  type="bar"
                  color="#8661c5"
                />
              </div>

              <div className="top-features">
                <h3>⭐ Top Features</h3>
                <SimpleChart
                  data={(usageMetrics?.topFeatures || []).slice(0, 5).map(f => ({
                    label: f.feature,
                    value: f.count
                  }))}
                  type="bar"
                  color="#107c10"
                />
              </div>
            </TabPanel>

            <TabPanel isActive={activeTab === 'performance'}>
              <div className="metrics-grid">
                <MetricCard
                  title="P50 Response Time"
                  value={formatDuration(performanceMetrics?.p50ResponseTime || 0)}
                  subtitle="Median response time"
                  icon="⏱️"
                  color="#0078d4"
                />
                <MetricCard
                  title="P95 Response Time"
                  value={formatDuration(performanceMetrics?.p95ResponseTime || 0)}
                  subtitle="95th percentile"
                  icon="⏱️"
                  color="#ff8c00"
                />
                <MetricCard
                  title="P99 Response Time"
                  value={formatDuration(performanceMetrics?.p99ResponseTime || 0)}
                  subtitle="99th percentile"
                  icon="⏱️"
                  color="#d13438"
                />
                <MetricCard
                  title="Error Rate"
                  value={`${((performanceMetrics?.errorRate || 0) * 100).toFixed(2)}%`}
                  subtitle="Failed requests"
                  icon="⚠️"
                  color="#d13438"
                />
              </div>

              <div className="throughput-section">
                <h3>📈 Throughput</h3>
                <MetricCard
                  title="Queries per Minute"
                  value={(performanceMetrics?.throughput || 0).toFixed(1)}
                  icon="🚀"
                  color="#107c10"
                />
              </div>

              <div className="slow-queries-section">
                <h3>🐌 Slow Queries</h3>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Query ID</th>
                      <th>Duration</th>
                      <th>Timestamp</th>
                      <th>User</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(performanceMetrics?.slowQueries || []).map((query: SlowQuery) => (
                      <tr key={query.id}>
                        <td>{query.query}</td>
                        <td>{formatDuration(query.duration)}</td>
                        <td>{new Date(query.timestamp).toLocaleString()}</td>
                        <td>{query.userId}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </TabPanel>

            <TabPanel isActive={activeTab === 'users'}>
              <div className="user-engagement-header">
                <h3>👥 User Engagement</h3>
                <p>Top {userEngagement.length} most engaged users</p>
              </div>

              <div className="user-list">
                {userEngagement.slice(0, 20).map((user) => (
                  <div key={user.userId} className="user-card">
                    <div className="user-header">
                      <span className="user-id">{user.userId}</span>
                      <span className={`engagement-badge score-${Math.floor(user.engagementScore / 20)}`}>
                        {user.engagementScore.toFixed(0)}
                      </span>
                    </div>
                    <div className="user-stats">
                      <div className="stat">
                        <label>Sessions</label>
                        <value>{user.totalSessions}</value>
                      </div>
                      <div className="stat">
                        <label>Queries</label>
                        <value>{user.totalQueries}</value>
                      </div>
                      <div className="stat">
                        <label>Avg/Session</label>
                        <value>{user.avgQueriesPerSession.toFixed(1)}</value>
                      </div>
                    </div>
                    <div className="favorite-features">
                      {user.favoriteFeatures.slice(0, 3).map((feature, i) => (
                        <span key={i} className="feature-tag">{feature}</span>
                      ))}
                    </div>
                    <div className="last-active">
                      Last active: {new Date(user.lastActive).toLocaleString()}
                    </div>
                  </div>
                ))}
              </div>
            </TabPanel>

            <TabPanel isActive={activeTab === 'ai'}>
              <div className="metrics-grid">
                <MetricCard
                  title="Total Tokens"
                  value={formatNumber(aiStats?.totalTokensUsed || 0)}
                  subtitle="Across all models"
                  icon="📝"
                  color="#0078d4"
                />
                <MetricCard
                  title="Estimated Cost"
                  value={`$${(aiStats?.totalCost || 0).toFixed(2)}`}
                  subtitle="For selected period"
                  icon="💰"
                  color="#107c10"
                />
                <MetricCard
                  title="Accuracy Score"
                  value={`${((aiStats?.accuracyScore || 0) * 100).toFixed(1)}%`}
                  subtitle="Successful queries"
                  icon="🎯"
                  color="#8661c5"
                />
                <MetricCard
                  title="Query Types"
                  value={aiStats?.queryTypes.length || 0}
                  subtitle="Different categories"
                  icon="📋"
                  color="#ff8c00"
                />
              </div>

              <div className="model-breakdown">
                <h3>🤖 Model Usage Breakdown</h3>
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Model</th>
                      <th>Tokens</th>
                      <th>Cost</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(aiStats?.modelBreakdown || []).map((model) => (
                      <tr key={model.model}>
                        <td>{model.model}</td>
                        <td>{formatNumber(model.tokens)}</td>
                        <td>${model.cost.toFixed(4)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="query-types">
                <h3>📊 Query Type Distribution</h3>
                <SimpleChart
                  data={(aiStats?.queryTypes || []).map(q => ({
                    label: q.type,
                    value: q.count
                  }))}
                  type="bar"
                  color="#8661c5"
                />
              </div>
            </TabPanel>

            <TabPanel isActive={activeTab === 'realtime'}>
              <div className="metrics-grid realtime">
                <MetricCard
                  title="Active Sessions"
                  value={realtimeMetrics.activeSessions}
                  subtitle="Currently online"
                  icon="🟢"
                  color="#107c10"
                />
                <MetricCard
                  title="Queries/Min"
                  value={realtimeMetrics.queriesPerMinute.toFixed(1)}
                  subtitle="Current rate"
                  icon="⚡"
                  color="#0078d4"
                />
                <MetricCard
                  title="Current Errors"
                  value={realtimeMetrics.currentErrors}
                  subtitle="Last 5 minutes"
                  icon="🔴"
                  color="#d13438"
                />
                <MetricCard
                  title="Avg Response"
                  value={formatDuration(realtimeMetrics.avgResponseTime)}
                  subtitle="Last 5 minutes"
                  icon="⏱️"
                  color="#ff8c00"
                />
              </div>

              <div className="live-events">
                <h3>🔴 Live Events</h3>
                <div className="events-stream">
                  {recentEvents.length === 0 ? (
                    <p className="no-events">Waiting for events...</p>
                  ) : (
                    recentEvents.map((event) => (
                      <div 
                        key={event.id} 
                        className={`event-item ${event.success ? 'success' : 'error'}`}
                      >
                        <span className="event-time">
                          {new Date(event.timestamp).toLocaleTimeString()}
                        </span>
                        <span className="event-type">{event.type}</span>
                        <span className="event-feature">{event.feature}</span>
                        {event.performance && (
                          <span className="event-duration">
                            {formatDuration(event.performance.processingTime)}
                          </span>
                        )}
                      </div>
                    ))
                  )}
                </div>
              </div>
            </TabPanel>
          </>
        )}
      </div>

      <style>{`
        .analytics-dashboard {
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

        .refresh-btn, .export-btn {
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

        .export-dropdown {
          position: relative;
        }

        .export-menu {
          display: none;
          position: absolute;
          top: 100%;
          right: 0;
          background: white;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
          border-radius: 4px;
          min-width: 150px;
          z-index: 100;
        }

        .export-dropdown:hover .export-menu {
          display: block;
        }

        .export-menu button {
          display: block;
          width: 100%;
          padding: 10px 15px;
          background: none;
          border: none;
          text-align: left;
          cursor: pointer;
        }

        .export-menu button:hover {
          background: #f3f2f1;
        }

        .dashboard-tabs {
          display: flex;
          gap: 5px;
          margin-bottom: 20px;
          border-bottom: 2px solid #e1e1e1;
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

        .metrics-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
          gap: 15px;
          margin-bottom: 25px;
        }

        .metric-card {
          background: #fff;
          border-radius: 8px;
          padding: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        }

        .metric-header {
          display: flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 10px;
        }

        .metric-icon {
          font-size: 20px;
        }

        .metric-title {
          font-size: 12px;
          color: #605e5c;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }

        .metric-value {
          font-size: 32px;
          font-weight: 600;
          margin-bottom: 5px;
        }

        .metric-subtitle {
          font-size: 12px;
          color: #605e5c;
        }

        .metric-trend {
          font-size: 12px;
          margin-top: 8px;
          font-weight: 500;
        }

        .metric-trend.positive {
          color: #107c10;
        }

        .metric-trend.negative {
          color: #d13438;
        }

        .chart-section, .top-features, .model-breakdown, .query-types,
        .throughput-section, .slow-queries-section, .user-engagement-header {
          margin-top: 25px;
        }

        .chart-section h3, .top-features h3, .model-breakdown h3, .query-types h3,
        .throughput-section h3, .slow-queries-section h3, .live-events h3 {
          font-size: 16px;
          color: #323130;
          margin-bottom: 15px;
        }

        .bar-chart {
          display: flex;
          flex-direction: column;
          gap: 10px;
        }

        .bar-item {
          display: flex;
          align-items: center;
          gap: 10px;
        }

        .bar-label {
          min-width: 100px;
          font-size: 12px;
          color: #605e5c;
        }

        .bar-wrapper {
          flex: 1;
          background: #f3f2f1;
          border-radius: 4px;
          height: 24px;
          overflow: hidden;
        }

        .bar {
          height: 100%;
          border-radius: 4px;
          display: flex;
          align-items: center;
          justify-content: flex-end;
          padding-right: 8px;
          transition: width 0.3s ease;
        }

        .bar-value {
          color: white;
          font-size: 12px;
          font-weight: 500;
        }

        .line-chart {
          width: 100%;
          height: 200px;
        }

        .data-table {
          width: 100%;
          border-collapse: collapse;
          font-size: 13px;
        }

        .data-table th, .data-table td {
          padding: 10px;
          text-align: left;
          border-bottom: 1px solid #e1e1e1;
        }

        .data-table th {
          background: #f3f2f1;
          font-weight: 600;
          color: #323130;
        }

        .user-list {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
          gap: 15px;
        }

        .user-card {
          background: #f9f9f9;
          border-radius: 8px;
          padding: 15px;
          border: 1px solid #e1e1e1;
        }

        .user-header {
          display: flex;
          justify-content: space-between;
          align-items: center;
          margin-bottom: 10px;
        }

        .user-id {
          font-weight: 600;
          color: #323130;
          font-size: 14px;
        }

        .engagement-badge {
          padding: 4px 10px;
          border-radius: 12px;
          font-size: 12px;
          font-weight: 600;
        }

        .engagement-badge.score-0 { background: #fde7e9; color: #d13438; }
        .engagement-badge.score-1 { background: #fff4ce; color: #ff8c00; }
        .engagement-badge.score-2 { background: #fffbd6; color: #ffc107; }
        .engagement-badge.score-3 { background: #dff6dd; color: #107c10; }
        .engagement-badge.score-4, .engagement-badge.score-5 { 
          background: #d1f0d9; color: #0b6e0b; 
        }

        .user-stats {
          display: flex;
          gap: 20px;
          margin-bottom: 10px;
        }

        .stat {
          text-align: center;
        }

        .stat label {
          display: block;
          font-size: 11px;
          color: #605e5c;
          text-transform: uppercase;
        }

        .stat value {
          display: block;
          font-size: 18px;
          font-weight: 600;
          color: #323130;
        }

        .favorite-features {
          display: flex;
          gap: 5px;
          flex-wrap: wrap;
          margin-bottom: 10px;
        }

        .feature-tag {
          background: #e8f4fd;
          color: #0078d4;
          padding: 3px 8px;
          border-radius: 12px;
          font-size: 11px;
        }

        .last-active {
          font-size: 11px;
          color: #605e5c;
        }

        .events-stream {
          max-height: 400px;
          overflow-y: auto;
          border: 1px solid #e1e1e1;
          border-radius: 4px;
        }

        .event-item {
          display: flex;
          gap: 15px;
          padding: 10px 15px;
          border-bottom: 1px solid #f3f2f1;
          font-size: 13px;
          align-items: center;
        }

        .event-item:last-child {
          border-bottom: none;
        }

        .event-item.success {
          border-left: 3px solid #107c10;
        }

        .event-item.error {
          border-left: 3px solid #d13438;
        }

        .event-time {
          min-width: 80px;
          color: #605e5c;
          font-size: 12px;
        }

        .event-type {
          background: #f3f2f1;
          padding: 2px 8px;
          border-radius: 4px;
          font-size: 11px;
          text-transform: uppercase;
        }

        .event-feature {
          flex: 1;
          color: #323130;
        }

        .event-duration {
          color: #605e5c;
          font-size: 12px;
        }

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

        .no-data, .no-events {
          text-align: center;
          padding: 40px;
          color: #605e5c;
        }

        .tab-panel {
          display: none;
        }

        .tab-panel.active {
          display: block;
        }
      `}</style>
    </div>
  );
};

export default AnalyticsDashboard;
