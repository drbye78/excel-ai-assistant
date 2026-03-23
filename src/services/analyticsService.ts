/**
 * Analytics Service - Enterprise Usage Analytics
 * Week 15: Analytics Dashboard
 * 
 * Tracks add-in usage, AI query statistics, user engagement,
 * and performance metrics for enterprise analytics.
 */

import { enterpriseAuth, UserRole } from './enterpriseAuth';
import { logger } from '../utils/logger';

// Analytics event types
export type AnalyticsEventType = 
  | 'ai_query'
  | 'formula_parsed'
  | 'visualization_created'
  | 'batch_operation'
  | 'recipe_created'
  | 'recipe_applied'
  | 'dax_analyzed'
  | 'power_query_generated'
  | 'custom_function_created'
  | 'conversation_started'
  | 'conversation_ended'
  | 'error_occurred'
  | 'feature_used';

// Analytics event structure
export interface AnalyticsEvent {
  id: string;
  type: AnalyticsEventType;
  userId: string;
  userRole: UserRole;
  sessionId: string;
  timestamp: Date;
  feature: string;
  duration?: number; // milliseconds
  success: boolean;
  metadata?: Record<string, any>;
  performance?: {
    processingTime: number;
    memoryUsed?: number;
  };
}

// Aggregated metrics
export interface UsageMetrics {
  totalUsers: number;
  activeUsers: number;
  totalSessions: number;
  avgSessionDuration: number;
  totalQueries: number;
  successRate: number;
  topFeatures: { feature: string; count: number }[];
  dailyActiveUsers: { date: string; count: number }[];
  queryVolume: { date: string; count: number }[];
}

export interface PerformanceMetrics {
  avgResponseTime: number;
  p50ResponseTime: number;
  p95ResponseTime: number;
  p99ResponseTime: number;
  errorRate: number;
  throughput: number; // queries per minute
  slowQueries: SlowQuery[];
}

export interface SlowQuery {
  id: string;
  query: string;
  duration: number;
  timestamp: Date;
  userId: string;
}

export interface UserEngagement {
  userId: string;
  totalSessions: number;
  totalQueries: number;
  avgQueriesPerSession: number;
  lastActive: Date;
  favoriteFeatures: string[];
  engagementScore: number; // 0-100
}

export interface AIUsageStats {
  totalTokensUsed: number;
  totalCost: number;
  modelBreakdown: { model: string; tokens: number; cost: number }[];
  queryTypes: { type: string; count: number }[];
  accuracyScore: number;
  userSatisfaction: number;
}

// Time range for analytics queries
export type TimeRange = '24h' | '7d' | '14d' | '30d' | '90d' | '1y' | 'custom';

export interface DateRange {
  start: Date;
  end: Date;
}

class AnalyticsService {
  private static instance: AnalyticsService;
  private events: AnalyticsEvent[] = [];
  private readonly maxEventsInMemory = 10000;
  private sessionId: string;
  private startTime: Date;
  private currentUserId: string = '';
  private isInitialized: boolean = false;
  private cleanupIntervalId: ReturnType<typeof setInterval> | null = null;
  private trackingTimeouts: Map<string, ReturnType<typeof setTimeout>> = new Map();

  // Callbacks for real-time updates
  private eventListeners: Map<AnalyticsEventType, Set<(event: AnalyticsEvent) => void>> = new Map();

  private constructor() {
    this.sessionId = this.generateSessionId();
    this.startTime = new Date();
  }

  static getInstance(): AnalyticsService {
    if (!AnalyticsService.instance) {
      AnalyticsService.instance = new AnalyticsService();
    }
    return AnalyticsService.instance;
  }

  async initialize(): Promise<void> {
    if (this.isInitialized) return;

    // Get current user from auth service
    const user = enterpriseAuth.getCurrentUser();
    if (user) {
      this.currentUserId = user.id;
    }

    // Load persisted events from IndexedDB
    await this.loadPersistedEvents();

    // Track session start
    this.trackEvent('feature_used', 'session_start', { 
      userAgent: navigator.userAgent,
      platform: navigator.platform 
    });

    // Setup periodic cleanup with proper interval management
    this.cleanupIntervalId = setInterval(() => this.cleanupOldEvents(), 60000); // Every minute

    this.isInitialized = true;
  }

  /**
   * Cleanup service resources
   * Call this when the service is no longer needed
   */
  dispose(): void {
    // Clear cleanup interval
    if (this.cleanupIntervalId) {
      clearInterval(this.cleanupIntervalId);
      this.cleanupIntervalId = null;
    }

    // Clear all tracking timeouts
    this.trackingTimeouts.forEach((timeout) => {
      clearTimeout(timeout);
    });
    this.trackingTimeouts.clear();

    // Clear event listeners
    this.eventListeners.clear();

    this.isInitialized = false;
  }

  private generateSessionId(): string {
    return `sess_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  private generateEventId(): string {
    return `evt_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Track an analytics event
   */
  trackEvent(
    type: AnalyticsEventType,
    feature: string,
    metadata?: Record<string, any>,
    duration?: number,
    success: boolean = true,
    performance?: { processingTime: number; memoryUsed?: number }
  ): void {
    const event: AnalyticsEvent = {
      id: this.generateEventId(),
      type,
      userId: this.currentUserId || 'anonymous',
      userRole: enterpriseAuth.getCurrentUser()?.role || 'standard_user',
      sessionId: this.sessionId,
      timestamp: new Date(),
      feature,
      duration,
      success,
      metadata,
      performance
    };

    this.events.push(event);

    // Trim if too many events in memory
    if (this.events.length > this.maxEventsInMemory) {
      this.events = this.events.slice(-this.maxEventsInMemory);
    }

    // Persist to IndexedDB
    this.persistEvent(event);

    // Notify listeners
    this.notifyListeners(type, event);

    // Send to server if enterprise user
    if (enterpriseAuth.isAuthenticated()) {
      this.sendToServer(event);
    }
  }

  /**
   * Start tracking duration for an operation
   */
  startTracking(feature: string): string {
    const trackingId = `track_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const global = globalThis as { [key: string]: unknown };
    global[trackingId] = {
      feature,
      startTime: performance.now()
    };

    // Set a timeout to auto-cleanup stale tracking entries (5 minutes)
    const timeout = setTimeout(() => {
      delete global[trackingId];
      this.trackingTimeouts.delete(trackingId);
    }, 5 * 60 * 1000);

    this.trackingTimeouts.set(trackingId, timeout);
    return trackingId;
  }

  /**
   * End tracking and record event
   */
  endTracking(
    trackingId: string,
    type: AnalyticsEventType,
    success: boolean = true,
    metadata?: Record<string, any>
  ): void {
    const global = globalThis as { [key: string]: unknown };
    const tracking = global[trackingId] as { feature: string; startTime: number } | undefined;
    if (!tracking) return;

    const duration = performance.now() - tracking.startTime;
    this.trackEvent(type, tracking.feature, metadata, duration, success);

    // Clean up tracking entry and timeout
    delete global[trackingId];
    const timeout = this.trackingTimeouts.get(trackingId);
    if (timeout) {
      clearTimeout(timeout);
      this.trackingTimeouts.delete(trackingId);
    }
  }

  /**
   * Track AI query with detailed metrics
   */
  trackAIQuery(
    query: string,
    model: string,
    tokensUsed: number,
    duration: number,
    success: boolean,
    responseType?: string
  ): void {
    const metadata = {
      queryLength: query.length,
      model,
      tokensUsed,
      responseType,
      queryHash: this.hashQuery(query)
    };

    this.trackEvent('ai_query', 'ai_assistant', metadata, duration, success, {
      processingTime: duration
    });
  }

  /**
   * Track formula parsing
   */
  trackFormulaParse(formula: string, complexity: number, success: boolean): void {
    this.trackEvent('formula_parsed', 'formula_explainer', {
      formulaLength: formula.length,
      complexity,
      hasErrors: !success
    }, undefined, success);
  }

  /**
   * Track batch operation
   */
  trackBatchOperation(operationType: string, itemCount: number, duration: number, success: boolean): void {
    this.trackEvent('batch_operation', 'batch_processor', {
      operationType,
      itemCount,
      itemsPerSecond: itemCount / (duration / 1000)
    }, duration, success);
  }

  /**
   * Get usage metrics for a time range
   */
  async getUsageMetrics(timeRange: TimeRange, customRange?: DateRange): Promise<UsageMetrics> {
    const events = this.getEventsInRange(timeRange, customRange);
    const uniqueUsers = new Set(events.map(e => e.userId));
    const sessions = new Set(events.map(e => e.sessionId));
    
    const queryEvents = events.filter(e => e.type === 'ai_query');
    const successfulQueries = queryEvents.filter(e => e.success);

    // Calculate daily active users
    const dailyUsers = this.aggregateByDay(events, e => e.userId);
    const queryVolume = this.aggregateByDay(queryEvents);

    // Get top features
    const featureCounts = new Map<string, number>();
    events.forEach(e => {
      const count = featureCounts.get(e.feature) || 0;
      featureCounts.set(e.feature, count + 1);
    });
    const topFeatures = Array.from(featureCounts.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([feature, count]) => ({ feature, count }));

    // Calculate average session duration
    const sessionDurations = this.calculateSessionDurations(events);
    const avgSessionDuration = sessionDurations.length > 0
      ? sessionDurations.reduce((a, b) => a + b, 0) / sessionDurations.length
      : 0;

    return {
      totalUsers: uniqueUsers.size,
      activeUsers: uniqueUsers.size, // In this time range
      totalSessions: sessions.size,
      avgSessionDuration,
      totalQueries: queryEvents.length,
      successRate: queryEvents.length > 0 ? successfulQueries.length / queryEvents.length : 0,
      topFeatures,
      dailyActiveUsers: dailyUsers.map(d => ({ date: d.date, count: d.uniqueValues.size })),
      queryVolume: queryVolume.map(d => ({ date: d.date, count: d.count }))
    };
  }

  /**
   * Get performance metrics
   */
  async getPerformanceMetrics(timeRange: TimeRange, customRange?: DateRange): Promise<PerformanceMetrics> {
    const events = this.getEventsInRange(timeRange, customRange);
    const eventsWithDuration = events.filter(e => e.performance?.processingTime);
    
    if (eventsWithDuration.length === 0) {
      return {
        avgResponseTime: 0,
        p50ResponseTime: 0,
        p95ResponseTime: 0,
        p99ResponseTime: 0,
        errorRate: 0,
        throughput: 0,
        slowQueries: []
      };
    }

    const durations = eventsWithDuration.map(e => e.performance!.processingTime).sort((a, b) => a - b);
    const totalEvents = events.length;
    const errorEvents = events.filter(e => !e.success).length;

    // Calculate time range in minutes
    const timeRangeMinutes = this.getTimeRangeMinutes(timeRange, customRange);
    const throughput = timeRangeMinutes > 0 ? totalEvents / timeRangeMinutes : 0;

    // Get slow queries (top 5% by duration)
    const slowQueryThreshold = durations[Math.floor(durations.length * 0.95)];
    const slowQueries: SlowQuery[] = eventsWithDuration
      .filter(e => e.performance!.processingTime >= slowQueryThreshold)
      .map(e => ({
        id: e.id,
        query: e.metadata?.queryHash || 'unknown',
        duration: e.performance!.processingTime,
        timestamp: e.timestamp,
        userId: e.userId
      }))
      .sort((a, b) => b.duration - a.duration)
      .slice(0, 20);

    return {
      avgResponseTime: durations.reduce((a, b) => a + b, 0) / durations.length,
      p50ResponseTime: this.getPercentile(durations, 0.5),
      p95ResponseTime: this.getPercentile(durations, 0.95),
      p99ResponseTime: this.getPercentile(durations, 0.99),
      errorRate: errorEvents / totalEvents,
      throughput,
      slowQueries
    };
  }

  /**
   * Get user engagement metrics
   */
  async getUserEngagement(timeRange: TimeRange, customRange?: DateRange): Promise<UserEngagement[]> {
    const events = this.getEventsInRange(timeRange, customRange);
    const userStats = new Map<string, {
      sessions: Set<string>;
      queries: number;
      features: Map<string, number>;
      lastActive: Date;
    }>();

    events.forEach(e => {
      let stats = userStats.get(e.userId);
      if (!stats) {
        stats = {
          sessions: new Set(),
          queries: 0,
          features: new Map(),
          lastActive: e.timestamp
        };
        userStats.set(e.userId, stats);
      }

      stats.sessions.add(e.sessionId);
      if (e.type === 'ai_query') stats.queries++;
      
      const featureCount = stats.features.get(e.feature) || 0;
      stats.features.set(e.feature, featureCount + 1);
      
      if (e.timestamp > stats.lastActive) {
        stats.lastActive = e.timestamp;
      }
    });

    return Array.from(userStats.entries()).map(([userId, stats]) => {
      const favoriteFeatures = Array.from(stats.features.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .map(([feature]) => feature);

      const avgQueriesPerSession = stats.sessions.size > 0 
        ? stats.queries / stats.sessions.size 
        : 0;

      // Calculate engagement score (0-100)
      const sessionScore = Math.min(stats.sessions.size * 5, 30);
      const queryScore = Math.min(stats.queries * 0.5, 40);
      const varietyScore = Math.min(stats.features.size * 5, 30);
      const engagementScore = Math.min(sessionScore + queryScore + varietyScore, 100);

      return {
        userId,
        totalSessions: stats.sessions.size,
        totalQueries: stats.queries,
        avgQueriesPerSession,
        lastActive: stats.lastActive,
        favoriteFeatures,
        engagementScore
      };
    }).sort((a, b) => b.engagementScore - a.engagementScore);
  }

  /**
   * Get AI usage statistics
   */
  async getAIUsageStats(timeRange: TimeRange, customRange?: DateRange): Promise<AIUsageStats> {
    const events = this.getEventsInRange(timeRange, customRange)
      .filter(e => e.type === 'ai_query');

    const modelBreakdown = new Map<string, { tokens: number; count: number }>();
    const queryTypes = new Map<string, number>();
    let totalTokens = 0;
    let successfulQueries = 0;

    events.forEach(e => {
      const model = e.metadata?.model || 'unknown';
      const tokens = e.metadata?.tokensUsed || 0;
      const type = e.metadata?.responseType || 'general';

      const modelStats = modelBreakdown.get(model) || { tokens: 0, count: 0 };
      modelStats.tokens += tokens;
      modelStats.count++;
      modelBreakdown.set(model, modelStats);

      const typeCount = queryTypes.get(type) || 0;
      queryTypes.set(type, typeCount + 1);

      totalTokens += tokens;
      if (e.success) successfulQueries++;
    });

    // Estimate cost (simplified)
    const modelCosts: Record<string, number> = {
      'gpt-4': 0.00003,
      'gpt-4-turbo': 0.00001,
      'gpt-3.5-turbo': 0.0000005,
      'local': 0
    };

    let totalCost = 0;
    const modelBreakdownArray = Array.from(modelBreakdown.entries()).map(([model, stats]) => {
      const costPerToken = modelCosts[model] || 0.00001;
      const cost = stats.tokens * costPerToken;
      totalCost += cost;
      return { model, tokens: stats.tokens, cost };
    });

    return {
      totalTokensUsed: totalTokens,
      totalCost,
      modelBreakdown: modelBreakdownArray,
      queryTypes: Array.from(queryTypes.entries()).map(([type, count]) => ({ type, count })),
      accuracyScore: events.length > 0 ? successfulQueries / events.length : 0,
      userSatisfaction: 0 // Would require user feedback
    };
  }

  /**
   * Get real-time metrics for dashboard
   */
  getRealTimeMetrics(): {
    activeSessions: number;
    queriesPerMinute: number;
    currentErrors: number;
    avgResponseTime: number;
  } {
    const last5Minutes = this.events.filter(e => 
      e.timestamp > new Date(Date.now() - 5 * 60 * 1000)
    );

    const sessions = new Set(last5Minutes.map(e => e.sessionId));
    const queryEvents = last5Minutes.filter(e => e.type === 'ai_query');
    const errorEvents = last5Minutes.filter(e => !e.success);
    
    const durations = queryEvents
      .filter(e => e.performance?.processingTime)
      .map(e => e.performance!.processingTime);

    return {
      activeSessions: sessions.size,
      queriesPerMinute: queryEvents.length / 5,
      currentErrors: errorEvents.length,
      avgResponseTime: durations.length > 0 
        ? durations.reduce((a, b) => a + b, 0) / durations.length 
        : 0
    };
  }

  /**
   * Export analytics data
   */
  async exportAnalytics(format: 'json' | 'csv', timeRange: TimeRange, customRange?: DateRange): Promise<string> {
    const events = this.getEventsInRange(timeRange, customRange);

    if (format === 'json') {
      return JSON.stringify(events, null, 2);
    }

    // CSV format
    const headers = ['id', 'type', 'userId', 'sessionId', 'timestamp', 'feature', 'duration', 'success'];
    const rows = events.map(e => [
      e.id,
      e.type,
      e.userId,
      e.sessionId,
      e.timestamp.toISOString(),
      e.feature,
      e.duration || '',
      e.success
    ]);

    return [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
  }

  /**
   * Subscribe to real-time events
   */
  subscribeToEvent(type: AnalyticsEventType, callback: (event: AnalyticsEvent) => void): () => void {
    if (!this.eventListeners.has(type)) {
      this.eventListeners.set(type, new Set());
    }
    this.eventListeners.get(type)!.add(callback);

    return () => {
      this.eventListeners.get(type)?.delete(callback);
    };
  }

  // Private helper methods

  private getEventsInRange(timeRange: TimeRange, customRange?: DateRange): AnalyticsEvent[] {
    const now = Date.now();
    let startTime: number;

    switch (timeRange) {
      case '24h':
        startTime = now - 24 * 60 * 60 * 1000;
        break;
      case '7d':
        startTime = now - 7 * 24 * 60 * 60 * 1000;
        break;
      case '14d':
        startTime = now - 14 * 24 * 60 * 60 * 1000;
        break;
      case '30d':
        startTime = now - 30 * 24 * 60 * 60 * 1000;
        break;
      case '90d':
        startTime = now - 90 * 24 * 60 * 60 * 1000;
        break;
      case '1y':
        startTime = now - 365 * 24 * 60 * 60 * 1000;
        break;
      case 'custom':
        startTime = customRange?.start.getTime() || 0;
        break;
      default:
        startTime = now - 7 * 24 * 60 * 60 * 1000;
    }

    return this.events.filter(e => e.timestamp.getTime() >= startTime);
  }

  private aggregateByDay<T>(
    events: AnalyticsEvent[], 
    valueExtractor?: (e: AnalyticsEvent) => T
  ): { date: string; count: number; uniqueValues: Set<T> }[] {
    const byDay = new Map<string, { count: number; values: Set<T> }>();

    events.forEach(e => {
      const date = e.timestamp.toISOString().split('T')[0];
      let dayStats = byDay.get(date);
      if (!dayStats) {
        dayStats = { count: 0, values: new Set() };
        byDay.set(date, dayStats);
      }
      dayStats.count++;
      if (valueExtractor) {
        dayStats.values.add(valueExtractor(e));
      }
    });

    return Array.from(byDay.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, stats]) => ({
        date,
        count: stats.count,
        uniqueValues: stats.values
      }));
  }

  private calculateSessionDurations(events: AnalyticsEvent[]): number[] {
    const sessionEvents = new Map<string, AnalyticsEvent[]>();
    
    events.forEach(e => {
      let session = sessionEvents.get(e.sessionId);
      if (!session) {
        session = [];
        sessionEvents.set(e.sessionId, session);
      }
      session.push(e);
    });

    return Array.from(sessionEvents.values()).map(sessionEvents => {
      const sorted = sessionEvents.sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());
      const first = sorted[0].timestamp.getTime();
      const last = sorted[sorted.length - 1].timestamp.getTime();
      return last - first;
    });
  }

  private getPercentile(sortedArray: number[], percentile: number): number {
    const index = Math.ceil(sortedArray.length * percentile) - 1;
    return sortedArray[Math.max(0, index)];
  }

  private getTimeRangeMinutes(timeRange: TimeRange, customRange?: DateRange): number {
    switch (timeRange) {
      case '24h': return 24 * 60;
      case '7d': return 7 * 24 * 60;
      case '14d': return 14 * 24 * 60;
      case '30d': return 30 * 24 * 60;
      case '90d': return 90 * 24 * 60;
      case '1y': return 365 * 24 * 60;
      case 'custom':
        if (customRange) {
          return (customRange.end.getTime() - customRange.start.getTime()) / (60 * 1000);
        }
        return 0;
      default: return 7 * 24 * 60;
    }
  }

  private hashQuery(query: string): string {
    let hash = 0;
    for (let i = 0; i < query.length; i++) {
      const char = query.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash;
    }
    return `q_${Math.abs(hash).toString(36)}`;
  }

  private async persistEvent(event: AnalyticsEvent): Promise<void> {
    try {
      const db = await this.openIndexedDB();
      const transaction = db.transaction(['analytics'], 'readwrite');
      const store = transaction.objectStore('analytics');
      store.add(event);
    } catch (error) {
      logger.warn('Failed to persist analytics event', undefined, error as Error);
    }
  }

  private async loadPersistedEvents(): Promise<void> {
    try {
      const db = await this.openIndexedDB();
      const transaction = db.transaction(['analytics'], 'readonly');
      const store = transaction.objectStore('analytics');
      const request = store.getAll();

      request.onsuccess = () => {
        const events = request.result as AnalyticsEvent[];
        // Only load last 30 days
        const thirtyDaysAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
        this.events = events
          .filter(e => new Date(e.timestamp).getTime() > thirtyDaysAgo)
          .map(e => ({ ...e, timestamp: new Date(e.timestamp) }));
      };
    } catch (error) {
      logger.warn('Failed to load persisted analytics', undefined, error as Error);
    }
  }

  private async cleanupOldEvents(): Promise<void> {
    try {
      const db = await this.openIndexedDB();
      const transaction = db.transaction(['analytics'], 'readwrite');
      const store = transaction.objectStore('analytics');
      
      // Delete events older than 90 days
      const ninetyDaysAgo = Date.now() - 90 * 24 * 60 * 60 * 1000;
      const index = store.index('timestamp');
      const range = IDBKeyRange.upperBound(new Date(ninetyDaysAgo));
      
      const request = index.openCursor(range);
      request.onsuccess = () => {
        const cursor = request.result;
        if (cursor) {
          cursor.delete();
          cursor.continue();
        }
      };
    } catch (error) {
      logger.warn('Failed to cleanup old analytics', undefined, error as Error);
    }
  }

  private openIndexedDB(): Promise<IDBDatabase> {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open('ExcelAIAnalytics', 1);
      
      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve(request.result);
      
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains('analytics')) {
          const store = db.createObjectStore('analytics', { keyPath: 'id' });
          store.createIndex('timestamp', 'timestamp', { unique: false });
          store.createIndex('userId', 'userId', { unique: false });
          store.createIndex('sessionId', 'sessionId', { unique: false });
          store.createIndex('type', 'type', { unique: false });
        }
      };
    });
  }

  private notifyListeners(type: AnalyticsEventType, event: AnalyticsEvent): void {
    this.eventListeners.get(type)?.forEach(callback => {
      try {
        callback(event);
      } catch (error) {
        logger.error('Analytics event listener error', undefined, error as Error);
      }
    });
  }

  private async sendToServer(event: AnalyticsEvent): Promise<void> {
    try {
      // Send to enterprise analytics endpoint
      const response = await fetch('/api/analytics/events', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${enterpriseAuth.getAccessToken()}`
        },
        body: JSON.stringify(event)
      });

      if (!response.ok) {
        throw new Error(`Failed to send analytics: ${response.statusText}`);
      }
    } catch (error) {
      // Silently fail - analytics shouldn't break functionality
      logger.warn('Failed to send analytics to server', undefined, error as Error);
    }
  }

  /**
   * Get current session statistics
   */
  getCurrentSessionStats(): {
    sessionId: string;
    startTime: Date;
    duration: number;
    eventCount: number;
    queryCount: number;
  } {
    const sessionEvents = this.events.filter(e => e.sessionId === this.sessionId);
    const queryEvents = sessionEvents.filter(e => e.type === 'ai_query');

    return {
      sessionId: this.sessionId,
      startTime: this.startTime,
      duration: Date.now() - this.startTime.getTime(),
      eventCount: sessionEvents.length,
      queryCount: queryEvents.length
    };
  }
}

export const analyticsService = AnalyticsService.getInstance();
export default analyticsService;
