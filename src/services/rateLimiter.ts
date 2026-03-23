/**
 * Rate Limiter Service
 * Implements token bucket algorithm for API rate limiting
 * Prevents cost overruns and manages request throttling
 */

export interface RateLimitConfig {
  /** Maximum requests allowed per window */
  maxRequests: number;
  /** Time window in milliseconds */
  windowMs: number;
  /** Maximum tokens that can accumulate */
  maxBurst?: number;
  /** Whether to queue requests when rate limited */
  queueEnabled?: boolean;
  /** Maximum queue size */
  maxQueueSize?: number;
}

export interface RateLimitState {
  tokens: number;
  lastRefill: number;
  queueLength: number;
  totalRequests: number;
  totalThrottled: number;
}

export interface QueuedRequest {
  id: string;
  timestamp: number;
  resolve: () => void;
  reject: (error: Error) => void;
}

export type RateLimitTier = 'free' | 'basic' | 'pro' | 'enterprise';

/**
 * Default rate limit configurations per tier
 */
const TIER_CONFIGS: Record<RateLimitTier, RateLimitConfig> = {
  free: {
    maxRequests: 10,
    windowMs: 60000, // 10 requests per minute
    maxBurst: 5,
    queueEnabled: true,
    maxQueueSize: 20
  },
  basic: {
    maxRequests: 60,
    windowMs: 60000, // 60 requests per minute
    maxBurst: 20,
    queueEnabled: true,
    maxQueueSize: 50
  },
  pro: {
    maxRequests: 300,
    windowMs: 60000, // 300 requests per minute
    maxBurst: 50,
    queueEnabled: true,
    maxQueueSize: 100
  },
  enterprise: {
    maxRequests: 1000,
    windowMs: 60000, // 1000 requests per minute
    maxBurst: 100,
    queueEnabled: true,
    maxQueueSize: 500
  }
};

/**
 * Model-specific rate limits (tokens per minute)
 */
const MODEL_RATE_LIMITS: Record<string, number> = {
  'gpt-4': 10000,
  'gpt-4-turbo': 10000,
  'gpt-4o': 30000,
  'gpt-4o-mini': 200000,
  'gpt-3.5-turbo': 200000,
  'claude-3-opus': 10000,
  'claude-3-sonnet': 40000,
  'claude-3-haiku': 100000,
  'default': 60000
};

export class RateLimiter {
  private static instance: RateLimiter;
  
  private config: RateLimitConfig;
  private tokens: number;
  private lastRefill: number;
  private queue: QueuedRequest[] = [];
  private totalRequests = 0;
  private totalThrottled = 0;
  private modelTokenUsage: Map<string, number> = new Map();
  private processingQueue = false;

  private constructor(config?: RateLimitConfig) {
    this.config = config || TIER_CONFIGS.basic;
    this.tokens = this.config.maxBurst || this.config.maxRequests;
    this.lastRefill = Date.now();
  }

  static getInstance(): RateLimiter {
    if (!RateLimiter.instance) {
      RateLimiter.instance = new RateLimiter();
    }
    return RateLimiter.instance;
  }

  /**
   * Configure rate limiter with specific settings
   */
  configure(config: Partial<RateLimitConfig>): void {
    this.config = { ...this.config, ...config };
    if (this.tokens > (this.config.maxBurst || this.config.maxRequests)) {
      this.tokens = this.config.maxBurst || this.config.maxRequests;
    }
  }

  /**
   * Set rate limit tier
   */
  setTier(tier: RateLimitTier): void {
    this.configure(TIER_CONFIGS[tier]);
  }

  /**
   * Get current configuration
   */
  getConfig(): RateLimitConfig {
    return { ...this.config };
  }

  /**
   * Refill tokens based on elapsed time
   */
  private refillTokens(): void {
    const now = Date.now();
    const elapsed = now - this.lastRefill;
    
    // Calculate tokens to add based on rate
    const tokensPerMs = this.config.maxRequests / this.config.windowMs;
    const tokensToAdd = elapsed * tokensPerMs;
    
    const maxTokens = this.config.maxBurst || this.config.maxRequests;
    this.tokens = Math.min(maxTokens, this.tokens + tokensToAdd);
    this.lastRefill = now;
  }

  /**
   * Try to acquire a token for making a request
   * @returns true if request can proceed, false if rate limited
   */
  tryAcquire(): boolean {
    this.refillTokens();
    
    if (this.tokens >= 1) {
      this.tokens -= 1;
      this.totalRequests++;
      return true;
    }
    
    this.totalThrottled++;
    return false;
  }

  /**
   * Acquire a token, queueing the request if necessary
   * @param requestId - Optional request identifier
   * @returns Promise that resolves when request can proceed
   */
  async acquire(requestId?: string): Promise<void> {
    // Try immediate acquisition
    if (this.tryAcquire()) {
      return;
    }

    // If queueing is disabled, throw error
    if (!this.config.queueEnabled) {
      throw new RateLimitError(
        'Rate limit exceeded. Please try again later.',
        this.getTimeUntilNextToken()
      );
    }

    // Check queue size
    if (this.queue.length >= (this.config.maxQueueSize || 50)) {
      throw new RateLimitError(
        'Rate limit queue is full. Please try again later.',
        this.getEstimatedWaitTime()
      );
    }

    // Queue the request
    return new Promise((resolve, reject) => {
      const request: QueuedRequest = {
        id: requestId || `req_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
        timestamp: Date.now(),
        resolve,
        reject
      };
      
      this.queue.push(request);
      this.processQueue();
    });
  }

  /**
   * Process queued requests
   */
  private processQueue(): void {
    if (this.processingQueue || this.queue.length === 0) {
      return;
    }

    this.processingQueue = true;

    const processNext = () => {
      if (this.queue.length === 0) {
        this.processingQueue = false;
        return;
      }

      this.refillTokens();

      if (this.tokens >= 1) {
        const request = this.queue.shift();
        if (request) {
          this.tokens -= 1;
          this.totalRequests++;
          request.resolve();
        }
      }

      // Continue processing if there are more requests
      if (this.queue.length > 0) {
        const waitTime = this.getTimeUntilNextToken();
        setTimeout(processNext, waitTime);
      } else {
        this.processingQueue = false;
      }
    };

    processNext();
  }

  /**
   * Get time until next token is available (in ms)
   */
  private getTimeUntilNextToken(): number {
    const tokensPerMs = this.config.maxRequests / this.config.windowMs;
    const tokensNeeded = 1 - this.tokens;
    return Math.ceil(tokensNeeded / tokensPerMs);
  }

  /**
   * Get estimated wait time for queued requests
   */
  private getEstimatedWaitTime(): number {
    const tokensPerMs = this.config.maxRequests / this.config.windowMs;
    const tokensNeeded = this.queue.length;
    return Math.ceil(tokensNeeded / tokensPerMs);
  }

  /**
   * Track token usage for a specific model
   */
  trackModelUsage(model: string, tokensUsed: number): void {
    const current = this.modelTokenUsage.get(model) || 0;
    this.modelTokenUsage.set(model, current + tokensUsed);
  }

  /**
   * Check if model has remaining capacity
   */
  hasModelCapacity(model: string, estimatedTokens: number): boolean {
    const limit = MODEL_RATE_LIMITS[model] || MODEL_RATE_LIMITS['default'];
    const current = this.modelTokenUsage.get(model) || 0;
    return current + estimatedTokens <= limit;
  }

  /**
   * Get remaining capacity for a model
   */
  getModelRemainingCapacity(model: string): number {
    const limit = MODEL_RATE_LIMITS[model] || MODEL_RATE_LIMITS['default'];
    const current = this.modelTokenUsage.get(model) || 0;
    return Math.max(0, limit - current);
  }

  /**
   * Reset model usage counters (should be called periodically)
   */
  resetModelUsage(): void {
    this.modelTokenUsage.clear();
  }

  /**
   * Get current state
   */
  getState(): RateLimitState {
    this.refillTokens();
    return {
      tokens: this.tokens,
      lastRefill: this.lastRefill,
      queueLength: this.queue.length,
      totalRequests: this.totalRequests,
      totalThrottled: this.totalThrottled
    };
  }

  /**
   * Get time until rate limit resets
   */
  getTimeUntilReset(): number {
    const now = Date.now();
    const windowStart = Math.floor(now / this.config.windowMs) * this.config.windowMs;
    const windowEnd = windowStart + this.config.windowMs;
    return windowEnd - now;
  }

  /**
   * Clear the request queue
   */
  clearQueue(): void {
    const error = new Error('Request cancelled');
    this.queue.forEach(request => request.reject(error));
    this.queue = [];
  }

  /**
   * Reset the rate limiter
   */
  reset(): void {
    this.tokens = this.config.maxBurst || this.config.maxRequests;
    this.lastRefill = Date.now();
    this.clearQueue();
    this.totalRequests = 0;
    this.totalThrottled = 0;
    this.modelTokenUsage.clear();
  }
}

/**
 * Custom error for rate limiting
 */
export class RateLimitError extends Error {
  constructor(
    message: string,
    public retryAfter: number
  ) {
    super(message);
    this.name = 'RateLimitError';
  }
}

export default RateLimiter.getInstance();