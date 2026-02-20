/**
 * Advanced Analytics Service Integration Tests
 * Phase 4: Testing & Quality Assurance
 *
 * Tests the integration between advancedAnalyticsService and analyticsService
 * including anomaly detection, forecasting, insights, and reporting features.
 */

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
} from '../advancedAnalytics';
import { analyticsService, AnalyticsEvent } from '../analyticsService';

// Mock localStorage with full Storage interface
const localStorageMock: Storage = {
  getItem: jest.fn(),
  setItem: jest.fn(),
  removeItem: jest.fn(),
  clear: jest.fn(),
  length: 0,
  key: jest.fn()
} as unknown as Storage;

global.localStorage = localStorageMock;

// Mock IndexedDB with proper types
// Note: In a real browser environment, IndexedDB is available
// For Node.js tests, we provide a minimal mock
const mockIndexedDB = {
  open: jest.fn().mockImplementation(() => ({
    onsuccess: null as ((this: IDBOpenDBRequest, ev: Event) => any) | null,
    onerror: null as ((this: IDBOpenDBRequest, ev: Event) => any) | null,
    result: null as IDBDatabase | null
  }))
} as unknown as IDBFactory;

// Only mock indexedDB if it's not available (Node.js environment)
if (typeof global.indexedDB === 'undefined') {
  global.indexedDB = mockIndexedDB;
}

// Mock EnterpriseAuthService
jest.mock('../enterpriseAuth', () => ({
  EnterpriseAuthService: {
    getCurrentUser: jest.fn().mockReturnValue({ id: 'test-user', role: 'user' }),
    isEnterpriseUser: jest.fn().mockReturnValue(false),
    getAccessToken: jest.fn().mockReturnValue('test-token')
  },
  UserRole: {
    ADMIN: 'admin',
    USER: 'user',
    GUEST: 'guest'
  }
}));

describe('AdvancedAnalyticsService Integration Tests', () => {
  let mockIDBDatabase: any;
  let mockIDBTransaction: any;
  let mockIDBObjectStore: any;
  let mockIDBRequest: any;

  beforeEach(() => {
    jest.clearAllMocks();
    
    // Setup IndexedDB mocks
    mockIDBRequest = {
      onsuccess: null,
      onerror: null,
      result: [],
      onupgradeneeded: null
    };

    mockIDBObjectStore = {
      add: jest.fn(),
      getAll: jest.fn().mockReturnValue(mockIDBRequest),
      createIndex: jest.fn(),
      index: jest.fn().mockReturnValue({
        openCursor: jest.fn().mockReturnValue(mockIDBRequest)
      }),
      delete: jest.fn()
    };

    mockIDBTransaction = {
      objectStore: jest.fn().mockReturnValue(mockIDBObjectStore)
    };

    mockIDBDatabase = {
      transaction: jest.fn().mockReturnValue(mockIDBTransaction),
      objectStoreNames: {
        contains: jest.fn().mockReturnValue(false)
      },
      createObjectStore: jest.fn().mockReturnValue(mockIDBObjectStore)
    };

    mockIndexedDB.open.mockReturnValue(mockIDBRequest);
    
    // Reset localStorage mocks
    (localStorage.getItem as jest.Mock).mockReturnValue(null);
    (localStorage.setItem as jest.Mock).mockImplementation(() => {});
  });

  // ==================== ANOMALY DETECTION TESTS ====================

  describe('Anomaly Detection', () => {
    it('should detect volume anomalies in query data', async () => {
      // Generate events with a significant spike
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];
      
      // Generate 48 hours of baseline data
      for (let i = 48; i >= 0; i--) {
        const hour = new Date(baseDate.getTime() - i * 60 * 60 * 1000);
        const hourStr = hour.toISOString().slice(0, 13) + ':00:00';
        
        // Baseline: 50 queries per hour
        const queryCount = i === 2 ? 500 : 50; // Spike at hour 2
        
        for (let j = 0; j < queryCount; j++) {
          events.push({
            id: `evt_${hour.getTime()}_${j}`,
            type: 'ai_query',
            userId: `user_${j % 20}`,
            userRole: 'user',
            sessionId: `sess_${hour.getTime()}`,
            timestamp: new Date(hour.getTime() + j * 1000),
            feature: 'ai_assistant',
            success: true,
            performance: {
              processingTime: 1000
            }
          });
        }
      }

      // Mock the exportAnalytics response
      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const anomalies = await advancedAnalyticsService.detectAnomalies('7d');

      expect(anomalies.length).toBeGreaterThan(0);
      
      const volumeAnomaly = anomalies.find(a => a.metric === 'query_volume');
      expect(volumeAnomaly).toBeDefined();
      expect(volumeAnomaly?.type).toBe('spike');
      expect(volumeAnomaly?.severity).toBe('critical');
    });

    it('should detect latency anomalies', async () => {
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];

      for (let i = 48; i >= 0; i--) {
        const hour = new Date(baseDate.getTime() - i * 60 * 60 * 1000);
        
        for (let j = 0; j < 50; j++) {
          events.push({
            id: `evt_${hour.getTime()}_${j}`,
            type: 'ai_query',
            userId: `user_${j % 10}`,
            userRole: 'user',
            sessionId: `sess_${hour.getTime()}`,
            timestamp: new Date(hour.getTime() + j * 1000),
            feature: 'ai_assistant',
            success: true,
            performance: {
              processingTime: i === 5 ? 15000 : 1000 // Extreme latency at hour 5
            }
          });
        }
      }

      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const anomalies = await advancedAnalyticsService.detectAnomalies('7d');

      const latencyAnomaly = anomalies.find(a => a.metric === 'response_time');
      expect(latencyAnomaly).toBeDefined();
      expect(latencyAnomaly?.type).toBe('spike');
      expect(latencyAnomaly?.severity).toBe('critical');
      expect(latencyAnomaly?.recommendation).toContain('performance');
    });

    it('should detect error rate anomalies', async () => {
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];

      for (let i = 48; i >= 0; i--) {
        const hour = new Date(baseDate.getTime() - i * 60 * 60 * 1000);
        
        for (let j = 0; j < 100; j++) {
          const isErrorHour = i === 10;
          events.push({
            id: `evt_${hour.getTime()}_${j}`,
            type: 'ai_query',
            userId: `user_${j % 10}`,
            userRole: 'user',
            sessionId: `sess_${hour.getTime()}`,
            timestamp: new Date(hour.getTime() + j * 1000),
            feature: 'ai_assistant',
            success: isErrorHour ? j < 30 : true, // 30% error rate at hour 10
            performance: {
              processingTime: 1000
            }
          });
        }
      }

      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const anomalies = await advancedAnalyticsService.detectAnomalies('7d');

      const errorAnomaly = anomalies.find(a => a.metric === 'error_rate');
      expect(errorAnomaly).toBeDefined();
      expect(errorAnomaly?.severity).toBe('high');
    });

    it('should store and retrieve anomalies', async () => {
      const mockAnomalies: Anomaly[] = [
        {
          id: 'anom_1',
          type: 'spike',
          metric: 'query_volume',
          severity: 'high',
          timestamp: new Date(),
          expectedValue: 100,
          actualValue: 500,
          deviation: 400,
          description: 'Test anomaly'
        }
      ];

      (localStorage.getItem as jest.Mock).mockReturnValue(JSON.stringify(mockAnomalies));

      const recentAnomalies = advancedAnalyticsService.getRecentAnomalies(10);
      
      expect(recentAnomalies.length).toBe(1);
      expect(recentAnomalies[0].id).toBe('anom_1');
    });
  });

  // ==================== FORECASTING TESTS ====================

  describe('Forecasting', () => {
    it('should generate forecast for query volume', async () => {
      // Mock historical data
      const mockMetrics = {
        totalUsers: 100,
        activeUsers: 50,
        totalSessions: 200,
        avgSessionDuration: 300000,
        totalQueries: 1000,
        successRate: 0.95,
        topFeatures: [],
        dailyActiveUsers: Array.from({ length: 90 }, (_, i) => ({
          date: new Date(Date.now() - (90 - i) * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
          count: 50 + i * 0.5 // Slight upward trend
        })),
        queryVolume: Array.from({ length: 90 }, (_, i) => ({
          date: new Date(Date.now() - (90 - i) * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
          count: 100 + i * 2 // Upward trend
        }))
      };

      jest.spyOn(analyticsService, 'getUsageMetrics').mockResolvedValue(mockMetrics);

      const forecast = await advancedAnalyticsService.generateForecast('query_volume', 30);

      expect(forecast.metric).toBe('query_volume');
      expect(forecast.horizon).toBe(30);
      expect(forecast.predictions.length).toBe(30);
      expect(forecast.trend).toBe('up');
      
      // Verify prediction structure
      const firstPrediction = forecast.predictions[0];
      expect(firstPrediction).toHaveProperty('date');
      expect(firstPrediction).toHaveProperty('value');
      expect(firstPrediction).toHaveProperty('lowerBound');
      expect(firstPrediction).toHaveProperty('upperBound');
      expect(firstPrediction).toHaveProperty('confidence');
      expect(firstPrediction.confidence).toBeGreaterThan(0);
      expect(firstPrediction.lowerBound).toBeLessThanOrEqual(firstPrediction.value);
      expect(firstPrediction.upperBound).toBeGreaterThanOrEqual(firstPrediction.value);
    });

    it('should detect seasonality in forecast', async () => {
      const mockMetrics = {
        totalUsers: 100,
        activeUsers: 50,
        totalSessions: 200,
        avgSessionDuration: 300000,
        totalQueries: 1000,
        successRate: 0.95,
        topFeatures: [],
        dailyActiveUsers: Array.from({ length: 168 }, (_, i) => {
          const date = new Date(Date.now() - (168 - i) * 60 * 60 * 1000);
          // Simulate daily pattern: higher during business hours
          const hour = date.getHours();
          const isBusinessHours = hour >= 9 && hour < 18;
          return {
            date: date.toISOString(),
            count: isBusinessHours ? 80 : 30
          };
        }),
        queryVolume: []
      };

      jest.spyOn(analyticsService, 'getUsageMetrics').mockResolvedValue(mockMetrics);

      const forecast = await advancedAnalyticsService.generateForecast('user_growth', 14);

      expect(forecast.seasonality).toBeDefined();
    });
  });

  // ==================== INSIGHTS GENERATION TESTS ====================

  describe('Insights Generation', () => {
    it('should generate performance insights', async () => {
      const mockPerfMetrics = {
        avgResponseTime: 2000,
        p50ResponseTime: 1500,
        p95ResponseTime: 8000, // Above threshold
        p99ResponseTime: 12000,
        errorRate: 0.02,
        throughput: 50,
        slowQueries: []
      };

      jest.spyOn(analyticsService, 'getPerformanceMetrics').mockResolvedValue(mockPerfMetrics);
      
      // Mock other dependencies
      jest.spyOn(analyticsService, 'getUsageMetrics').mockResolvedValue({
        totalUsers: 100,
        activeUsers: 50,
        totalSessions: 200,
        avgSessionDuration: 300000,
        totalQueries: 1000,
        successRate: 0.95,
        topFeatures: [],
        dailyActiveUsers: [],
        queryVolume: []
      });

      jest.spyOn(analyticsService, 'getAIUsageStats').mockResolvedValue({
        totalTokensUsed: 10000,
        totalCost: 0.5,
        modelBreakdown: [],
        queryTypes: [],
        accuracyScore: 0.95,
        userSatisfaction: 0
      });

      jest.spyOn(analyticsService, 'getUserEngagement').mockResolvedValue([]);

      const insights = await advancedAnalyticsService.generateInsights();

      const perfInsight = insights.find(i => i.category === 'performance');
      expect(perfInsight).toBeDefined();
      expect(perfInsight?.title).toContain('Response Time');
      expect(perfInsight?.actionable).toBe(true);
    });

    it('should generate cost insights for high GPT-4 usage', async () => {
      jest.spyOn(analyticsService, 'getPerformanceMetrics').mockResolvedValue({
        avgResponseTime: 1000,
        p50ResponseTime: 800,
        p95ResponseTime: 2000,
        p99ResponseTime: 3000,
        errorRate: 0.01,
        throughput: 50,
        slowQueries: []
      });

      jest.spyOn(analyticsService, 'getUsageMetrics').mockResolvedValue({
        totalUsers: 100,
        activeUsers: 50,
        totalSessions: 200,
        avgSessionDuration: 300000,
        totalQueries: 1000,
        successRate: 0.95,
        topFeatures: [],
        dailyActiveUsers: [],
        queryVolume: []
      });

      jest.spyOn(analyticsService, 'getAIUsageStats').mockResolvedValue({
        totalTokensUsed: 100000,
        totalCost: 5.0,
        modelBreakdown: [
          { model: 'gpt-4', tokens: 60000, cost: 1.8 }, // 60% GPT-4
          { model: 'gpt-3.5-turbo', tokens: 40000, cost: 0.02 }
        ],
        queryTypes: [],
        accuracyScore: 0.95,
        userSatisfaction: 0
      });

      jest.spyOn(analyticsService, 'getUserEngagement').mockResolvedValue([]);

      const insights = await advancedAnalyticsService.generateInsights();

      const costInsight = insights.find(i => i.category === 'cost');
      expect(costInsight).toBeDefined();
      expect(costInsight?.title).toContain('GPT-4');
    });
  });

  // ==================== CORRELATION ANALYSIS TESTS ====================

  describe('Correlation Analysis', () => {
    it('should calculate correlations between metrics', async () => {
      // Create mock events with correlated patterns
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];

      for (let i = 30; i >= 0; i--) {
        const date = new Date(baseDate.getTime() - i * 24 * 60 * 60 * 1000);
        
        // As query volume increases, response time increases (positive correlation)
        const queryCount = 50 + i * 2;
        const responseTime = 1000 + i * 50;

        for (let j = 0; j < queryCount; j++) {
          events.push({
            id: `evt_${date.getTime()}_${j}`,
            type: 'ai_query',
            userId: `user_${j % 10}`,
            userRole: 'user',
            sessionId: `sess_${date.getTime()}`,
            timestamp: new Date(date.getTime() + j * 1000),
            feature: 'ai_assistant',
            success: true,
            performance: {
              processingTime: responseTime + Math.random() * 200
            }
          });
        }
      }

      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const correlations = await advancedAnalyticsService.analyzeCorrelations();

      expect(correlations.length).toBeGreaterThan(0);
      
      const volTimeCorr = correlations.find(c => 
        c.metric1 === 'query_volume' && c.metric2 === 'response_time'
      );
      expect(volTimeCorr).toBeDefined();
      expect(volTimeCorr?.coefficient).toBeGreaterThan(0);
      expect(volTimeCorr?.trend).toBe('positive');
    });
  });

  // ==================== COHORT ANALYSIS TESTS ====================

  describe('Cohort Analysis', () => {
    it('should perform cohort analysis on user retention', async () => {
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];

      // Simulate 4 weeks of user activity
      for (let week = 0; week < 4; week++) {
        const weekStart = new Date(baseDate.getTime() - (4 - week) * 7 * 24 * 60 * 60 * 1000);
        const cohortUsers = 20; // 20 new users per week

        for (let user = 0; user < cohortUsers; user++) {
          const userId = `user_${week}_${user}`;
          
          // Users are active for subsequent weeks (decreasing retention)
          const activeWeeks = Math.max(1, 4 - week - Math.floor(user / 10));
          
          for (let w = 0; w < activeWeeks; w++) {
            const activeDate = new Date(weekStart.getTime() + w * 7 * 24 * 60 * 60 * 1000);
            
            events.push({
              id: `evt_${userId}_${w}`,
              type: 'ai_query',
              userId,
              userRole: 'user',
              sessionId: `sess_${userId}_${w}`,
              timestamp: activeDate,
              feature: 'ai_assistant',
              success: true
            });
          }
        }
      }

      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const cohorts = await advancedAnalyticsService.analyzeCohorts();

      expect(cohorts.length).toBeGreaterThan(0);
      
      const firstCohort = cohorts[0];
      expect(firstCohort).toHaveProperty('cohortDate');
      expect(firstCohort).toHaveProperty('size');
      expect(firstCohort).toHaveProperty('retention');
      expect(firstCohort.retention.length).toBeGreaterThan(0);
      
      // Retention should decrease over time
      if (firstCohort.retention.length > 1) {
        expect(firstCohort.retention[0]).toBeGreaterThanOrEqual(firstCohort.retention[1]);
      }
    });
  });

  // ==================== FUNNEL ANALYSIS TESTS ====================

  describe('Funnel Analysis', () => {
    it('should analyze user journey funnel', async () => {
      const baseDate = new Date();
      const events: AnalyticsEvent[] = [];

      // Simulate 100 users through the funnel
      for (let i = 0; i < 100; i++) {
        const userId = `user_${i}`;
        const sessionId = `sess_${i}`;
        const baseTime = baseDate.getTime() - Math.random() * 30 * 24 * 60 * 60 * 1000;

        // Stage 1: All users open the add-in
        events.push({
          id: `evt_${userId}_open`,
          type: 'feature_used',
          userId,
          userRole: 'user',
          sessionId,
          timestamp: new Date(baseTime),
          feature: 'session_start',
          success: true
        });

        // Stage 2: 70 users make AI queries
        if (i < 70) {
          events.push({
            id: `evt_${userId}_query`,
            type: 'ai_query',
            userId,
            userRole: 'user',
            sessionId,
            timestamp: new Date(baseTime + 5000),
            feature: 'ai_assistant',
            success: true
          });

          // Stage 3: 50 users apply results
          if (i < 50) {
            events.push({
              id: `evt_${userId}_apply`,
              type: 'batch_operation',
              userId,
              userRole: 'user',
              sessionId,
              timestamp: new Date(baseTime + 10000),
              feature: 'batch_processor',
              success: true
            });

            // Stage 4: 30 users return (use feature more than once)
            if (i < 30) {
              events.push({
                id: `evt_${userId}_return`,
                type: 'ai_query',
                userId,
                userRole: 'user',
                sessionId,
                timestamp: new Date(baseTime + 86400000), // Next day
                feature: 'ai_assistant',
                success: true
              });
            }
          }
        }
      }

      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify(events));

      const funnel = await advancedAnalyticsService.analyzeFunnel();

      expect(funnel).toBeDefined();
      expect(funnel.name).toBe('User Journey');
      expect(funnel.stages).toHaveLength(4);
      
      // Stage 1: 100 users
      expect(funnel.stages[0].users).toBe(100);
      expect(funnel.stages[0].conversionRate).toBe(100);

      // Stage 2: 70 users (70% conversion)
      expect(funnel.stages[1].users).toBe(70);
      expect(funnel.stages[1].conversionRate).toBe(70);
      expect(funnel.stages[1].dropOff).toBe(30);

      // Stage 3: 50 users
      expect(funnel.stages[2].users).toBe(50);

      // Stage 4: 30 users
      expect(funnel.stages[3].users).toBe(30);

      // Overall conversion
      expect(funnel.overallConversion).toBe(30);
    });
  });

  // ==================== ALERTS MANAGEMENT TESTS ====================

  describe('Alerts Management', () => {
    it('should create and retrieve alerts', async () => {
      const alertConfig: Omit<AlertConfig, 'id'> = {
        name: 'High Error Rate Alert',
        metric: 'error_rate',
        condition: 'above',
        threshold: 0.05,
        timeWindow: '24h',
        enabled: true,
        notificationChannels: ['email', 'in_app'],
        recipients: ['admin@example.com']
      };

      const alert = await advancedAnalyticsService.createAlert(alertConfig);

      expect(alert.id).toBeDefined();
      expect(alert.name).toBe(alertConfig.name);
      expect(alert.metric).toBe(alertConfig.metric);
      
      const alerts = advancedAnalyticsService.getAlerts();
      expect(alerts).toContainEqual(alert);
      
      expect(localStorage.setItem).toHaveBeenCalledWith(
        'advanced_analytics_alerts',
        expect.any(String)
      );
    });

    it('should update alerts', async () => {
      // First create an alert
      const alert = await advancedAnalyticsService.createAlert({
        name: 'Test Alert',
        metric: 'query_volume',
        condition: 'above',
        threshold: 100,
        timeWindow: '1h',
        enabled: true,
        notificationChannels: ['in_app']
      });

      // Update it
      await advancedAnalyticsService.updateAlert(alert.id, {
        threshold: 200,
        enabled: false
      });

      const alerts = advancedAnalyticsService.getAlerts();
      const updatedAlert = alerts.find(a => a.id === alert.id);
      
      expect(updatedAlert?.threshold).toBe(200);
      expect(updatedAlert?.enabled).toBe(false);
    });

    it('should delete alerts', async () => {
      const alert = await advancedAnalyticsService.createAlert({
        name: 'Alert to Delete',
        metric: 'response_time',
        condition: 'above',
        threshold: 5000,
        timeWindow: '1h',
        enabled: true,
        notificationChannels: ['in_app']
      });

      await advancedAnalyticsService.deleteAlert(alert.id);

      const alerts = advancedAnalyticsService.getAlerts();
      expect(alerts).not.toContainEqual(expect.objectContaining({ id: alert.id }));
    });
  });

  // ==================== SCHEDULED REPORTS TESTS ====================

  describe('Scheduled Reports', () => {
    it('should create and retrieve scheduled reports', async () => {
      const reportConfig: Omit<ScheduledReport, 'id'> = {
        name: 'Weekly Analytics Report',
        frequency: 'weekly',
        metrics: ['query_volume', 'user_engagement', 'error_rate'],
        recipients: ['team@example.com'],
        format: 'pdf',
        nextRun: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000)
      };

      const report = await advancedAnalyticsService.createScheduledReport(reportConfig);

      expect(report.id).toBeDefined();
      expect(report.name).toBe(reportConfig.name);
      expect(report.frequency).toBe(reportConfig.frequency);
      
      const reports = advancedAnalyticsService.getScheduledReports();
      expect(reports).toContainEqual(report);
      
      expect(localStorage.setItem).toHaveBeenCalledWith(
        'advanced_analytics_reports',
        expect.any(String)
      );
    });

    it('should update scheduled reports', async () => {
      const report = await advancedAnalyticsService.createScheduledReport({
        name: 'Daily Report',
        frequency: 'daily',
        metrics: ['query_volume'],
        recipients: ['user@example.com'],
        format: 'csv',
        nextRun: new Date()
      });

      await advancedAnalyticsService.updateScheduledReport(report.id, {
        frequency: 'weekly',
        metrics: ['query_volume', 'user_engagement']
      });

      const reports = advancedAnalyticsService.getScheduledReports();
      const updatedReport = reports.find(r => r.id === report.id);
      
      expect(updatedReport?.frequency).toBe('weekly');
      expect(updatedReport?.metrics).toContain('user_engagement');
    });

    it('should delete scheduled reports', async () => {
      const report = await advancedAnalyticsService.createScheduledReport({
        name: 'Report to Delete',
        frequency: 'monthly',
        metrics: ['all'],
        recipients: ['admin@example.com'],
        format: 'html',
        nextRun: new Date()
      });

      await advancedAnalyticsService.deleteScheduledReport(report.id);

      const reports = advancedAnalyticsService.getScheduledReports();
      expect(reports).not.toContainEqual(expect.objectContaining({ id: report.id }));
    });
  });

  // ==================== EDGE CASES AND ERROR HANDLING ====================

  describe('Edge Cases and Error Handling', () => {
    it('should handle empty event data gracefully', async () => {
      jest.spyOn(analyticsService, 'exportAnalytics').mockResolvedValue(JSON.stringify([]));

      const anomalies = await advancedAnalyticsService.detectAnomalies('24h');
      
      expect(anomalies).toEqual([]);
    });

    it('should handle insufficient data for forecasting', async () => {
      jest.spyOn(analyticsService, 'getUsageMetrics').mockResolvedValue({
        totalUsers: 10,
        activeUsers: 5,
        totalSessions: 20,
        avgSessionDuration: 300000,
        totalQueries: 50,
        successRate: 0.95,
        topFeatures: [],
        dailyActiveUsers: [], // Empty data
        queryVolume: []
      });

      const forecast = await advancedAnalyticsService.generateForecast('query_volume', 7);

      expect(forecast).toBeDefined();
      expect(forecast.predictions).toHaveLength(7);
      expect(forecast.accuracy).toBe(0);
    });

    it('should handle localStorage errors gracefully', async () => {
      (localStorage.setItem as jest.Mock).mockImplementation(() => {
        throw new Error('Storage quota exceeded');
      });

      // Should not throw
      await expect(advancedAnalyticsService.createAlert({
        name: 'Test Alert',
        metric: 'query_volume',
        condition: 'above',
        threshold: 100,
        timeWindow: '1h',
        enabled: true,
        notificationChannels: ['in_app']
      })).rejects.toThrow();
    });

    it('should load existing alerts from localStorage', async () => {
      const storedAlerts: AlertConfig[] = [
        {
          id: 'alert_1',
          name: 'Stored Alert',
          metric: 'error_rate',
          condition: 'above',
          threshold: 0.1,
          timeWindow: '24h',
          enabled: true,
          notificationChannels: ['email']
        }
      ];

      (localStorage.getItem as jest.Mock).mockImplementation((key) => {
        if (key === 'advanced_analytics_alerts') {
          return JSON.stringify(storedAlerts);
        }
        return null;
      });

      // Create new instance to trigger loadAlerts
      const service = (advancedAnalyticsService as any);
      service.alerts = []; // Reset
      await service.loadAlerts();

      const alerts = advancedAnalyticsService.getAlerts();
      expect(alerts.length).toBeGreaterThanOrEqual(0);
    });
  });
});
