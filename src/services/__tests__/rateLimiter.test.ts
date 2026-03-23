/**
 * Rate Limiter Service Unit Tests
 * Phase 2: Comprehensive Testing
 */

import { RateLimiter, RateLimitError, RateLimitTier, RateLimitConfig } from '../rateLimiter';

describe('RateLimiter', () => {
  let rateLimiter: RateLimiter;

  beforeEach(() => {
    // Reset singleton for each test
    (RateLimiter as any).instance = undefined;
    rateLimiter = RateLimiter.getInstance();
  });

  describe('Singleton Pattern', () => {
    it('should return the same instance', () => {
      const instance1 = RateLimiter.getInstance();
      const instance2 = RateLimiter.getInstance();
      expect(instance1).toBe(instance2);
    });
  });

  describe('Configuration', () => {
    it('should have default configuration', () => {
      const config = rateLimiter.getConfig();
      expect(config.maxRequests).toBeGreaterThan(0);
      expect(config.windowMs).toBeGreaterThan(0);
    });

    it('should update configuration', () => {
      rateLimiter.configure({ maxRequests: 100, windowMs: 60000 });
      const config = rateLimiter.getConfig();
      expect(config.maxRequests).toBe(100);
      expect(config.windowMs).toBe(60000);
    });

    it('should set tier configuration', () => {
      rateLimiter.setTier('pro');
      const config = rateLimiter.getConfig();
      expect(config.maxRequests).toBe(300);
    });

    it.each(['free', 'basic', 'pro', 'enterprise'] as RateLimitTier[])(
      'should configure %s tier correctly',
      (tier) => {
        rateLimiter.setTier(tier);
        const config = rateLimiter.getConfig();
        expect(config.maxRequests).toBeGreaterThan(0);
      }
    );
  });

  describe('Token Acquisition', () => {
    it('should acquire token when available', () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 60000 });
      const acquired = rateLimiter.tryAcquire();
      expect(acquired).toBe(true);
    });

    it('should deny token when rate limit exceeded', () => {
      rateLimiter.configure({ maxRequests: 2, windowMs: 60000, maxBurst: 2 });
      
      // Acquire all available tokens
      expect(rateLimiter.tryAcquire()).toBe(true);
      expect(rateLimiter.tryAcquire()).toBe(true);
      
      // Should be denied now
      expect(rateLimiter.tryAcquire()).toBe(false);
    });

    it('should track total requests', () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 60000 });
      
      rateLimiter.tryAcquire();
      rateLimiter.tryAcquire();
      rateLimiter.tryAcquire();
      
      const state = rateLimiter.getState();
      expect(state.totalRequests).toBe(3);
    });

    it('should track throttled requests', () => {
      rateLimiter.configure({ maxRequests: 1, windowMs: 60000, maxBurst: 1 });
      
      rateLimiter.tryAcquire();
      rateLimiter.tryAcquire(); // Should be throttled
      
      const state = rateLimiter.getState();
      expect(state.totalThrottled).toBe(1);
    });
  });

  describe('Async Acquisition with Queue', () => {
    it('should acquire immediately when tokens available', async () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 60000, queueEnabled: true });
      
      await expect(rateLimiter.acquire()).resolves.toBeUndefined();
    });

    it('should throw error when queue disabled and rate limited', async () => {
      rateLimiter.configure({ 
        maxRequests: 1, 
        windowMs: 60000, 
        maxBurst: 1, 
        queueEnabled: false 
      });
      
      await rateLimiter.acquire(); // First should succeed
      
      await expect(rateLimiter.acquire()).rejects.toThrow(RateLimitError);
    });

    it('should queue requests when rate limited', async () => {
      rateLimiter.configure({ 
        maxRequests: 1, 
        windowMs: 100, // Short window for testing
        maxBurst: 1, 
        queueEnabled: true,
        maxQueueSize: 5
      });
      
      // Acquire the only token
      await rateLimiter.acquire();
      
      // Queue a second request (should not throw immediately)
      const acquirePromise = rateLimiter.acquire();
      
      // Should be queued, not rejected
      await expect(acquirePromise).resolves.toBeUndefined();
    }, 10000);
  });

  describe('Model Usage Tracking', () => {
    it('should track model usage', () => {
      rateLimiter.trackModelUsage('gpt-4', 1000);
      rateLimiter.trackModelUsage('gpt-4', 500);
      
      expect(rateLimiter.hasModelCapacity('gpt-4', 5000)).toBe(true);
    });

    it('should detect when model capacity exceeded', () => {
      rateLimiter.resetModelUsage();
      
      // Use a lot of tokens
      rateLimiter.trackModelUsage('gpt-4', 15000);
      
      expect(rateLimiter.hasModelCapacity('gpt-4', 1000)).toBe(false);
    });

    it('should return remaining capacity', () => {
      rateLimiter.resetModelUsage();
      
      const remaining = rateLimiter.getModelRemainingCapacity('gpt-4');
      expect(remaining).toBeGreaterThan(0);
    });

    it('should handle unknown models with default limit', () => {
      const remaining = rateLimiter.getModelRemainingCapacity('unknown-model');
      expect(remaining).toBeGreaterThan(0);
    });
  });

  describe('State Management', () => {
    it('should return current state', () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 60000 });
      
      const state = rateLimiter.getState();
      
      expect(state).toHaveProperty('tokens');
      expect(state).toHaveProperty('lastRefill');
      expect(state).toHaveProperty('queueLength');
      expect(state).toHaveProperty('totalRequests');
      expect(state).toHaveProperty('totalThrottled');
    });

    it('should reset state', () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 60000 });
      
      rateLimiter.tryAcquire();
      rateLimiter.tryAcquire();
      
      rateLimiter.reset();
      
      const state = rateLimiter.getState();
      expect(state.totalRequests).toBe(0);
      expect(state.totalThrottled).toBe(0);
    });
  });

  describe('Queue Management', () => {
    it('should clear queue', async () => {
      rateLimiter.configure({ 
        maxRequests: 1, 
        windowMs: 60000, 
        maxBurst: 1,
        queueEnabled: true,
        maxQueueSize: 10
      });
      
      await rateLimiter.acquire();
      
      // Queue multiple requests
      const promises = [
        rateLimiter.acquire(),
        rateLimiter.acquire(),
        rateLimiter.acquire()
      ];
      
      // Clear queue
      rateLimiter.clearQueue();
      
      const state = rateLimiter.getState();
      expect(state.queueLength).toBe(0);
    });
  });

  describe('RateLimitError', () => {
    it('should create error with retryAfter', () => {
      const error = new RateLimitError('Rate limit exceeded', 5000);
      
      expect(error.message).toBe('Rate limit exceeded');
      expect(error.retryAfter).toBe(5000);
      expect(error.name).toBe('RateLimitError');
    });
  });

  describe('Edge Cases', () => {
    it('should handle zero maxRequests', () => {
      expect(() => {
        rateLimiter.configure({ maxRequests: 0, windowMs: 60000 });
      }).not.toThrow();
    });

    it('should handle very short window', () => {
      rateLimiter.configure({ maxRequests: 10, windowMs: 100 });
      
      const acquired = rateLimiter.tryAcquire();
      expect(acquired).toBe(true);
    });

    it('should handle concurrent acquisitions', async () => {
      rateLimiter.configure({ 
        maxRequests: 5, 
        windowMs: 1000, 
        maxBurst: 5,
        queueEnabled: true 
      });
      
      const promises = Array(3).fill(null).map(() => rateLimiter.acquire());
      
      await expect(Promise.all(promises)).resolves.toBeDefined();
    });
  });
});