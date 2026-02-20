/**
 * Cost Tracker Service
 * Tracks AI API token usage and associated costs
 * Provides cost estimation, budgeting, and alerts
 */

import { RateLimitTier } from './rateLimiter';
import { logger } from '../utils/logger';

export interface TokenUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
}

export interface ModelPricing {
  promptPrice: number;  // Price per 1M tokens
  completionPrice: number; // Price per 1M tokens
}

export interface UsageRecord {
  id: string;
  timestamp: Date;
  model: string;
  usage: TokenUsage;
  cost: number;
  requestId?: string;
  conversationId?: string;
  success: boolean;
}

export interface CostBudget {
  daily: number;
  weekly: number;
  monthly: number;
  perRequest: number;
}

export interface CostAlert {
  type: 'warning' | 'limit' | 'threshold';
  threshold: number;
  current: number;
  period: 'daily' | 'weekly' | 'monthly';
  message: string;
}

export interface CostSummary {
  totalCost: number;
  totalTokens: number;
  promptTokens: number;
  completionTokens: number;
  requestCount: number;
  successRate: number;
  averageCostPerRequest: number;
  byModel: Record<string, { cost: number; tokens: number; requests: number }>;
  byDay: Array<{ date: string; cost: number; tokens: number; requests: number }>;
}

/**
 * Model pricing data (as of 2024)
 * Prices are per 1M tokens in USD
 */
const MODEL_PRICING: Record<string, ModelPricing> = {
  // OpenAI models
  'gpt-4': { promptPrice: 30.00, completionPrice: 60.00 },
  'gpt-4-turbo': { promptPrice: 10.00, completionPrice: 30.00 },
  'gpt-4o': { promptPrice: 5.00, completionPrice: 15.00 },
  'gpt-4o-mini': { promptPrice: 0.15, completionPrice: 0.60 },
  'gpt-3.5-turbo': { promptPrice: 0.50, completionPrice: 1.50 },
  'gpt-3.5-turbo-16k': { promptPrice: 3.00, completionPrice: 4.00 },
  
  // Anthropic models (via OpenRouter)
  'anthropic/claude-3-opus-20240229': { promptPrice: 15.00, completionPrice: 75.00 },
  'anthropic/claude-3-sonnet-20240229': { promptPrice: 3.00, completionPrice: 15.00 },
  'anthropic/claude-3-haiku-20240307': { promptPrice: 0.25, completionPrice: 1.25 },
  
  // Meta models (via OpenRouter)
  'meta-llama/llama-3-70b-instruct': { promptPrice: 0.80, completionPrice: 0.80 },
  'meta-llama/llama-3-8b-instruct': { promptPrice: 0.06, completionPrice: 0.06 },
  
  // Mistral models (via OpenRouter)
  'mistralai/mixtral-8x7b-instruct': { promptPrice: 0.24, completionPrice: 0.24 },
  'mistralai/mistral-7b-instruct': { promptPrice: 0.06, completionPrice: 0.06 },
  
  // Google models (via OpenRouter)
  'google/gemini-pro-1.5': { promptPrice: 3.50, completionPrice: 10.50 },
  
  // Default fallback
  'default': { promptPrice: 1.00, completionPrice: 2.00 }
};

/**
 * Default budget limits per tier
 */
const TIER_BUDGETS: Record<RateLimitTier, CostBudget> = {
  free: { daily: 1.00, weekly: 5.00, monthly: 20.00, perRequest: 0.10 },
  basic: { daily: 10.00, weekly: 50.00, monthly: 200.00, perRequest: 1.00 },
  pro: { daily: 50.00, weekly: 250.00, monthly: 1000.00, perRequest: 5.00 },
  enterprise: { daily: 500.00, weekly: 2500.00, monthly: 10000.00, perRequest: 50.00 }
};

export class CostTracker {
  private static instance: CostTracker;
  
  private usageRecords: UsageRecord[] = [];
  private budget: CostBudget;
  private alerts: CostAlert[] = [];
  private alertCallbacks: Array<(alert: CostAlert) => void> = [];
  private storageKey = 'ai_cost_tracker_data';
  
  private constructor() {
    this.budget = TIER_BUDGETS.basic;
    this.loadFromStorage();
  }

  static getInstance(): CostTracker {
    if (!CostTracker.instance) {
      CostTracker.instance = new CostTracker();
    }
    return CostTracker.instance;
  }

  /**
   * Set budget limits
   */
  setBudget(budget: Partial<CostBudget>): void {
    this.budget = { ...this.budget, ...budget };
    this.checkAlerts();
  }

  /**
   * Set budget tier
   */
  setTier(tier: RateLimitTier): void {
    this.budget = TIER_BUDGETS[tier];
    this.checkAlerts();
  }

  /**
   * Get current budget
   */
  getBudget(): CostBudget {
    return { ...this.budget };
  }

  /**
   * Get pricing for a model
   */
  getModelPricing(model: string): ModelPricing {
    // Normalize model name
    const normalizedModel = model.toLowerCase().replace(/^openai\//, '');
    
    // Check exact match first
    if (MODEL_PRICING[model]) {
      return MODEL_PRICING[model];
    }
    
    // Check normalized match
    if (MODEL_PRICING[normalizedModel]) {
      return MODEL_PRICING[normalizedModel];
    }
    
    // Check partial match
    for (const [key, pricing] of Object.entries(MODEL_PRICING)) {
      if (model.toLowerCase().includes(key.toLowerCase()) || 
          key.toLowerCase().includes(normalizedModel)) {
        return pricing;
      }
    }
    
    return MODEL_PRICING['default'];
  }

  /**
   * Calculate cost for token usage
   */
  calculateCost(model: string, usage: TokenUsage): number {
    const pricing = this.getModelPricing(model);
    
    const promptCost = (usage.promptTokens / 1_000_000) * pricing.promptPrice;
    const completionCost = (usage.completionTokens / 1_000_000) * pricing.completionPrice;
    
    return promptCost + completionCost;
  }

  /**
   * Estimate cost before making a request
   */
  estimateCost(model: string, estimatedPromptTokens: number, maxCompletionTokens: number): number {
    const pricing = this.getModelPricing(model);
    
    const promptCost = (estimatedPromptTokens / 1_000_000) * pricing.promptPrice;
    const completionCost = (maxCompletionTokens / 1_000_000) * pricing.completionPrice;
    
    return promptCost + completionCost;
  }

  /**
   * Record token usage
   */
  recordUsage(
    model: string,
    usage: TokenUsage,
    metadata?: { requestId?: string; conversationId?: string; success?: boolean }
  ): UsageRecord {
    const cost = this.calculateCost(model, usage);
    
    const record: UsageRecord = {
      id: `usage_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      timestamp: new Date(),
      model,
      usage,
      cost,
      requestId: metadata?.requestId,
      conversationId: metadata?.conversationId,
      success: metadata?.success ?? true
    };
    
    this.usageRecords.push(record);
    this.saveToStorage();
    this.checkAlerts();
    
    return record;
  }

  /**
   * Check if a request would exceed budget
   */
  wouldExceedBudget(model: string, estimatedTokens: TokenUsage): boolean {
    const estimatedCost = this.calculateCost(model, estimatedTokens);
    
    // Check per-request limit
    if (estimatedCost > this.budget.perRequest) {
      return true;
    }
    
    // Check daily limit
    const dailyCost = this.getCostForPeriod('daily');
    if (dailyCost + estimatedCost > this.budget.daily) {
      return true;
    }
    
    return false;
  }

  /**
   * Check if request is within budget and throw if not
   */
  checkBudget(model: string, estimatedTokens: TokenUsage): void {
    const estimatedCost = this.calculateCost(model, estimatedTokens);
    
    if (estimatedCost > this.budget.perRequest) {
      throw new CostLimitError(
        `Request cost ($${estimatedCost.toFixed(4)}) exceeds per-request limit ($${this.budget.perRequest.toFixed(2)})`,
        'perRequest'
      );
    }
    
    const dailyCost = this.getCostForPeriod('daily');
    if (dailyCost + estimatedCost > this.budget.daily) {
      throw new CostLimitError(
        `Daily budget limit ($${this.budget.daily.toFixed(2)}) would be exceeded`,
        'daily'
      );
    }
  }

  /**
   * Get cost for a time period
   */
  getCostForPeriod(period: 'daily' | 'weekly' | 'monthly'): number {
    const now = new Date();
    let startDate: Date;
    
    switch (period) {
      case 'daily':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case 'weekly':
        startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        break;
      case 'monthly':
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        break;
    }
    
    return this.usageRecords
      .filter(r => r.timestamp >= startDate)
      .reduce((sum, r) => sum + r.cost, 0);
  }

  /**
   * Get token usage for a time period
   */
  getTokensForPeriod(period: 'daily' | 'weekly' | 'monthly'): TokenUsage {
    const now = new Date();
    let startDate: Date;
    
    switch (period) {
      case 'daily':
        startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
        break;
      case 'weekly':
        startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        break;
      case 'monthly':
        startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        break;
    }
    
    const records = this.usageRecords.filter(r => r.timestamp >= startDate);
    
    return {
      promptTokens: records.reduce((sum, r) => sum + r.usage.promptTokens, 0),
      completionTokens: records.reduce((sum, r) => sum + r.usage.completionTokens, 0),
      totalTokens: records.reduce((sum, r) => sum + r.usage.totalTokens, 0)
    };
  }

  /**
   * Get comprehensive cost summary
   */
  getSummary(period?: 'daily' | 'weekly' | 'monthly'): CostSummary {
    let records = this.usageRecords;
    
    if (period) {
      const now = new Date();
      let startDate: Date;
      
      switch (period) {
        case 'daily':
          startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
          break;
        case 'weekly':
          startDate = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          break;
        case 'monthly':
          startDate = new Date(now.getFullYear(), now.getMonth(), 1);
          break;
      }
      
      records = records.filter(r => r.timestamp >= startDate);
    }
    
    const totalCost = records.reduce((sum, r) => sum + r.cost, 0);
    const successCount = records.filter(r => r.success).length;
    
    // Group by model
    const byModel: Record<string, { cost: number; tokens: number; requests: number }> = {};
    for (const record of records) {
      if (!byModel[record.model]) {
        byModel[record.model] = { cost: 0, tokens: 0, requests: 0 };
      }
      byModel[record.model].cost += record.cost;
      byModel[record.model].tokens += record.usage.totalTokens;
      byModel[record.model].requests += 1;
    }
    
    // Group by day
    const byDayMap = new Map<string, { cost: number; tokens: number; requests: number }>();
    for (const record of records) {
      const date = record.timestamp.toISOString().split('T')[0];
      if (!byDayMap.has(date)) {
        byDayMap.set(date, { cost: 0, tokens: 0, requests: 0 });
      }
      const entry = byDayMap.get(date)!;
      entry.cost += record.cost;
      entry.tokens += record.usage.totalTokens;
      entry.requests += 1;
    }
    
    const byDay = Array.from(byDayMap.entries())
      .map(([date, data]) => ({ date, ...data }))
      .sort((a, b) => a.date.localeCompare(b.date));
    
    return {
      totalCost,
      totalTokens: records.reduce((sum, r) => sum + r.usage.totalTokens, 0),
      promptTokens: records.reduce((sum, r) => sum + r.usage.promptTokens, 0),
      completionTokens: records.reduce((sum, r) => sum + r.usage.completionTokens, 0),
      requestCount: records.length,
      successRate: records.length > 0 ? successCount / records.length : 1,
      averageCostPerRequest: records.length > 0 ? totalCost / records.length : 0,
      byModel,
      byDay
    };
  }

  /**
   * Register callback for cost alerts
   */
  onAlert(callback: (alert: CostAlert) => void): void {
    this.alertCallbacks.push(callback);
  }

  /**
   * Check for budget alerts
   */
  private checkAlerts(): void {
    const periods: Array<'daily' | 'weekly' | 'monthly'> = ['daily', 'weekly', 'monthly'];
    const thresholds = [0.5, 0.75, 0.9, 1.0]; // 50%, 75%, 90%, 100%
    
    for (const period of periods) {
      const cost = this.getCostForPeriod(period);
      const limit = this.budget[period];
      
      for (const threshold of thresholds) {
        if (cost >= limit * threshold) {
          const alert: CostAlert = {
            type: threshold >= 1.0 ? 'limit' : threshold >= 0.9 ? 'threshold' : 'warning',
            threshold: limit * threshold,
            current: cost,
            period,
            message: this.getAlertMessage(period, threshold, cost, limit)
          };
          
          // Check if this alert already exists
          const existingAlert = this.alerts.find(
            a => a.period === period && a.threshold === alert.threshold
          );
          
          if (!existingAlert) {
            this.alerts.push(alert);
            this.notifyAlert(alert);
          }
        }
      }
    }
  }

  /**
   * Generate alert message
   */
  private getAlertMessage(
    period: 'daily' | 'weekly' | 'monthly',
    threshold: number,
    current: number,
    limit: number
  ): string {
    const percentage = (threshold * 100).toFixed(0);
    
    if (threshold >= 1.0) {
      return `${period.charAt(0).toUpperCase() + period.slice(1)} budget limit of $${limit.toFixed(2)} has been reached!`;
    }
    
    return `${percentage}% of ${period} budget used ($${current.toFixed(2)} of $${limit.toFixed(2)})`;
  }

  /**
   * Notify alert callbacks
   */
  private notifyAlert(alert: CostAlert): void {
    for (const callback of this.alertCallbacks) {
      try {
        callback(alert);
      } catch (error) {
        logger.error('Alert callback error', { error, alert });
      }
    }
  }

  /**
   * Get active alerts
   */
  getAlerts(): CostAlert[] {
    return [...this.alerts];
  }

  /**
   * Clear alerts
   */
  clearAlerts(): void {
    this.alerts = [];
  }

  /**
   * Save data to localStorage
   */
  private saveToStorage(): void {
    try {
      const data = {
        records: this.usageRecords.slice(-1000), // Keep last 1000 records
        budget: this.budget
      };
      localStorage.setItem(this.storageKey, JSON.stringify(data));
    } catch (error) {
      logger.error('Failed to save cost tracker data', { error });
    }
  }

  /**
   * Load data from localStorage
   */
  private loadFromStorage(): void {
    try {
      const data = localStorage.getItem(this.storageKey);
      if (data) {
        const parsed = JSON.parse(data);
        this.usageRecords = (parsed.records || []).map((r: any) => ({
          ...r,
          timestamp: new Date(r.timestamp)
        }));
        if (parsed.budget) {
          this.budget = parsed.budget;
        }
      }
    } catch (error) {
      logger.error('Failed to load cost tracker data', { error });
    }
  }

  /**
   * Export usage data
   */
  exportData(): string {
    return JSON.stringify({
      exportDate: new Date().toISOString(),
      budget: this.budget,
      records: this.usageRecords
    }, null, 2);
  }

  /**
   * Clear all data
   */
  reset(): void {
    this.usageRecords = [];
    this.alerts = [];
    localStorage.removeItem(this.storageKey);
  }
}

/**
 * Custom error for cost limit exceeded
 */
export class CostLimitError extends Error {
  constructor(
    message: string,
    public period: 'perRequest' | 'daily' | 'weekly' | 'monthly'
  ) {
    super(message);
    this.name = 'CostLimitError';
  }
}

export default CostTracker.getInstance();