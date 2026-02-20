/**
 * Logger Service Tests
 *
 * Unit tests for the production logging framework
 * @module utils/__tests__/logger
 */

import { Logger, LogLevel, LogEntry } from '../logger';

describe('Logger', () => {
  let logger: Logger;

  beforeEach(() => {
    logger = Logger.getInstance();
    logger.setLevel(LogLevel.DEBUG);
    logger.clear();
  });

  afterEach(() => {
    logger.clear();
  });

  describe('Log Levels', () => {
    test('should log messages at appropriate levels', () => {
      logger.debug('debug message');
      logger.info('info message');
      logger.warn('warn message');
      logger.error('error message');

      const logs = logger.getLogs();
      expect(logs).toHaveLength(4);
      expect(logs[0].level).toBe(LogLevel.DEBUG);
      expect(logs[1].level).toBe(LogLevel.INFO);
      expect(logs[2].level).toBe(LogLevel.WARN);
      expect(logs[3].level).toBe(LogLevel.ERROR);
    });

    test('should respect log level filter', () => {
      logger.setLevel(LogLevel.WARN);
      
      logger.debug('should not appear');
      logger.info('should not appear');
      logger.warn('should appear');
      logger.error('should appear');

      const logs = logger.getLogs();
      expect(logs).toHaveLength(2);
      expect(logs[0].message).toBe('should appear');
      expect(logs[1].message).toBe('should appear');
    });

    test('should respect SILENT log level', () => {
      logger.setLevel(LogLevel.SILENT);
      
      logger.debug('silent');
      logger.info('silent');
      logger.warn('silent');
      logger.error('silent');

      const logs = logger.getLogs();
      expect(logs).toHaveLength(0);
    });
  });

  describe('Log Content', () => {
    test('should include context in logs', () => {
      const context = { userId: '123', action: 'test' };
      logger.info('test message', context);

      const logs = logger.getLogs();
      expect(logs[0].context).toEqual(context);
    });

    test('should include error details', () => {
      const error = new Error('Test error');
      logger.error('operation failed', { component: 'Test' }, error);

      const logs = logger.getLogs();
      expect(logs[0].error).toBe(error);
      expect(logs[0].message).toBe('operation failed');
    });

    test('should include timestamp', () => {
      const before = new Date();
      logger.info('test');
      const after = new Date();

      const logs = logger.getLogs();
      const logTime = new Date(logs[0].timestamp);
      
      expect(logTime.getTime()).toBeGreaterThanOrEqual(before.getTime());
      expect(logTime.getTime()).toBeLessThanOrEqual(after.getTime());
    });
  });

  describe('Log Management', () => {
    test('should limit log history', () => {
      logger.setLevel(LogLevel.DEBUG);
      for (let i = 0; i < 1100; i++) {
        logger.info(`message ${i}`);
      }

      const logs = logger.getLogs();
      expect(logs.length).toBeLessThanOrEqual(1000);
    });

    test('should clear logs', () => {
      logger.info('message 1');
      logger.info('message 2');
      
      expect(logger.getLogs()).toHaveLength(2);
      
      logger.clear();
      
      expect(logger.getLogs()).toHaveLength(0);
    });

    test('should export logs to JSON', () => {
      logger.info('test message', { key: 'value' });
      
      const exported = logger.export();
      const parsed = JSON.parse(exported);
      
      expect(parsed).toHaveLength(1);
      expect(parsed[0].message).toBe('test message');
      expect(parsed[0].context).toEqual({ key: 'value' });
    });
  });

  describe('Log Filtering', () => {
    beforeEach(() => {
      logger.debug('debug log');
      logger.info('info log');
      logger.warn('warn log');
      logger.error('error log');
    });

    test('should filter by level', () => {
      const logs = logger.getLogs({ level: LogLevel.WARN });
      
      expect(logs).toHaveLength(2);
      expect(logs[0].level).toBe(LogLevel.WARN);
      expect(logs[1].level).toBe(LogLevel.ERROR);
    });

    test('should filter by date', () => {
      const since = new Date();
      since.setSeconds(since.getSeconds() + 1);
      
      logger.info('future log');
      
      const logs = logger.getLogs({ since });
      
      expect(logs).toHaveLength(1);
      expect(logs[0].message).toBe('future log');
    });
  });

  describe('Subscriptions', () => {
    test('should notify subscribers', () => {
      const subscriber = jest.fn();
      const unsubscribe = logger.subscribe(subscriber);
      
      logger.info('test');
      
      expect(subscriber).toHaveBeenCalledTimes(1);
      expect(subscriber).toHaveBeenCalledWith(expect.objectContaining({
        message: 'test',
        level: LogLevel.INFO,
      }));
      
      unsubscribe();
    });

    test('should not notify after unsubscribe', () => {
      const subscriber = jest.fn();
      const unsubscribe = logger.subscribe(subscriber);
      
      unsubscribe();
      
      logger.info('test');
      
      expect(subscriber).not.toHaveBeenCalled();
    });

    test('should handle subscriber errors gracefully', () => {
      const errorSubscriber = jest.fn(() => { throw new Error('Subscriber error'); });
      const normalSubscriber = jest.fn();
      
      logger.subscribe(errorSubscriber);
      logger.subscribe(normalSubscriber);
      
      // Should not throw
      expect(() => logger.info('test')).not.toThrow();
      
      // Normal subscriber should still be called
      expect(normalSubscriber).toHaveBeenCalled();
    });
  });

  describe('Singleton Pattern', () => {
    test('should return same instance', () => {
      const instance1 = Logger.getInstance();
      const instance2 = Logger.getInstance();
      
      expect(instance1).toBe(instance2);
    });

    test('should share state across instances', () => {
      const instance1 = Logger.getInstance();
      const instance2 = Logger.getInstance();
      
      instance1.info('shared');
      
      const logs = instance2.getLogs();
      expect(logs).toHaveLength(1);
      expect(logs[0].message).toBe('shared');
    });
  });

  describe('Edge Cases', () => {
    test('should handle empty message', () => {
      logger.info('');
      
      const logs = logger.getLogs();
      expect(logs).toHaveLength(1);
      expect(logs[0].message).toBe('');
    });

    test('should handle undefined context', () => {
      logger.info('test');
      
      const logs = logger.getLogs();
      expect(logs[0].context).toBeUndefined();
    });

    test('should handle circular reference in context', () => {
      const context: any = { key: 'value' };
      context.self = context; // Circular reference
      
      // Should not throw
      expect(() => logger.info('test', context)).not.toThrow();
    });

    test('should handle very long messages', () => {
      const longMessage = 'a'.repeat(10000);
      
      logger.info(longMessage);
      
      const logs = logger.getLogs();
      expect(logs[0].message).toBe(longMessage);
    });
  });
});
