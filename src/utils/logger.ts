/**
 * Production Logger Utility
 * Replaces console.* calls with structured, controllable logging
 * Logs to both console and server-side file
 */

// Type references for Office.js
/// <reference types="@types/office-js" />
/// <reference types="@types/office-runtime" />

import { notificationManager } from './notificationManager';

export enum LogLevel {
  DEBUG = 0,
  INFO = 1,
  WARN = 2,
  ERROR = 3,
  SILENT = 4
}

export interface LogEntry {
  timestamp: string;
  level: LogLevel;
  message: string;
  context?: Record<string, unknown>;
  error?: {
    name?: string;
    message: string;
    stack?: string;
  };
}

export interface LoggerConfig {
  serverUrl?: string;
  batchSize?: number;
  flushInterval?: number;
  enableServerLogging?: boolean;
  minServerLogLevel?: LogLevel;
}

export class Logger {
  private static instance: Logger;
  private logLevel: LogLevel = LogLevel.DEBUG;
  private logs: LogEntry[] = [];
  private maxLogs: number = 1000;
  private subscribers: ((entry: LogEntry) => void)[] = [];
  
  // Server logging configuration
  private config: LoggerConfig = {
    serverUrl: typeof window !== 'undefined' && window.location.origin.includes('localhost')
      ? 'http://localhost:3001'
      : (typeof window !== 'undefined' ? window.location.origin : 'http://localhost:3001'),
    batchSize: 10,
    flushInterval: 5000, // 5 seconds
    enableServerLogging: true,
    minServerLogLevel: LogLevel.INFO
  };
  
  // Batch buffer for server logging
  private pendingLogs: LogEntry[] = [];
  private flushTimer: ReturnType<typeof setInterval> | null = null;
  private isFlushing: boolean = false;

  static getInstance(): Logger {
    if (!Logger.instance) {
      Logger.instance = new Logger();
    }
    return Logger.instance;
  }

  constructor() {
    // Start the flush timer
    this.startFlushTimer();
    
    // Flush on page unload
    if (typeof window !== 'undefined') {
      window.addEventListener('beforeunload', () => {
        this.flushToServer();
      });
    }
  }

  configure(config: Partial<LoggerConfig>): void {
    this.config = { ...this.config, ...config };
    
    // Restart timer if interval changed
    if (config.flushInterval !== undefined) {
      this.stopFlushTimer();
      this.startFlushTimer();
    }
  }

  setLevel(level: LogLevel): void {
    this.logLevel = level;
  }

  debug(message: string, context?: Record<string, unknown>): void {
    this.log(LogLevel.DEBUG, message, context);
  }

  info(message: string, context?: Record<string, unknown>): void {
    this.log(LogLevel.INFO, message, context);
  }

  warn(message: string, context?: Record<string, unknown>, error?: Error): void {
    this.log(LogLevel.WARN, message, context, error);
  }

  error(message: string, context?: Record<string, unknown>, error?: Error): void {
    this.log(LogLevel.ERROR, message, context, error);
    
    if (error && this.shouldNotifyUser(error)) {
      notificationManager.error(message);
    }
  }

  private log(level: LogLevel, message: string, context?: Record<string, unknown>, error?: Error): void {
    if (level < this.logLevel) return;

    const entry: LogEntry = {
      timestamp: new Date().toISOString(),
      level,
      message,
      context,
      error: error ? {
        name: error.name,
        message: error.message,
        stack: error.stack
      } : undefined
    };

    // Add to local logs
    this.logs.push(entry);
    
    if (this.logs.length > this.maxLogs) {
      this.logs = this.logs.slice(-this.maxLogs);
    }

    // Add to pending batch for server
    if (this.config.enableServerLogging && level >= (this.config.minServerLogLevel ?? LogLevel.INFO)) {
      this.pendingLogs.push(entry);
      
      // Flush if batch size reached
      if (this.pendingLogs.length >= (this.config.batchSize ?? 1)) {
        this.flushToServer();
      }
    }

    // Notify subscribers
    this.subscribers.forEach(cb => {
      try {
        cb(entry);
      } catch {
        // Don't let logging failures break the app
      }
    });

    // Always output to console
    this.outputToConsole(entry);
  }

  private outputToConsole(entry: LogEntry): void {
    const prefix = `[${entry.timestamp}] ${LogLevel[entry.level]}:`;
    switch (entry.level) {
      case LogLevel.DEBUG:
        console.debug(prefix, entry.message, entry.context);
        break;
      case LogLevel.INFO:
        console.info(prefix, entry.message, entry.context);
        break;
      case LogLevel.WARN:
        console.warn(prefix, entry.message, entry.context, entry.error);
        break;
      case LogLevel.ERROR:
        console.error(prefix, entry.message, entry.context, entry.error);
        break;
    }
  }

  private shouldNotifyUser(error: Error): boolean {
    const userErrors = ['ValidationError', 'PermissionError', 'NetworkError', 'ExcelAPIError'];
    return userErrors.some(type => error.name.includes(type));
  }

  // ============================================================================
  // SERVER LOGGING METHODS
  // ============================================================================

  private startFlushTimer(): void {
    if (this.flushTimer) return;
    
    this.flushTimer = setInterval(() => {
      this.flushToServer();
    }, this.config.flushInterval ?? 5000);
  }

  private stopFlushTimer(): void {
    if (this.flushTimer) {
      clearInterval(this.flushTimer);
      this.flushTimer = null;
    }
  }

  /**
   * Flush pending logs to the server
   * Uses fetch with keepalive for reliability (better CORS support than sendBeacon)
   */
  async flushToServer(): Promise<void> {
    if (this.isFlushing || this.pendingLogs.length === 0) return;
    
    this.isFlushing = true;
    const logsToSend = [...this.pendingLogs];
    this.pendingLogs = [];
    
    try {
      const serverUrl = this.config.serverUrl || 'http://localhost:3001';
      const url = `${serverUrl}/api/logs`;
      
      await this.sendWithFetch(url, logsToSend);
    } catch (error) {
      // Don't log this error to avoid infinite loop
      // Just put the logs back in the queue
      this.pendingLogs.unshift(...logsToSend);
      console.warn('[Logger] Failed to send logs to server:', error);
    } finally {
      this.isFlushing = false;
    }
  }

  private async sendWithFetch(url: string, logs: LogEntry[]): Promise<void> {
    // Use fetch with keepalive for reliability during page unload
    // keepalive is similar to sendBeacon but with better CORS support
    const fetchOptions: RequestInit & { keepalive?: boolean } = {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      },
      body: JSON.stringify({ entries: logs }),
      keepalive: true, // Allows the request to continue even if page unloads
      mode: 'cors', // Explicitly request CORS mode
      credentials: 'omit' // Don't send cookies for cross-origin requests
    };
    
    const response = await fetch(url, fetchOptions);
    
    if (!response.ok) {
      throw new Error(`Server returned ${response.status}`);
    }
  }

  /**
   * Force flush all pending logs immediately
   */
  async flush(): Promise<void> {
    await this.flushToServer();
  }

  /**
   * Get server log statistics
   */
  async getServerStats(): Promise<{
    files: Array<{ name: string; size: number; sizeMB: string; modified: string }>;
    totalSizeMB: string;
    config: { maxLogSizeMB: number; maxLogFiles: number };
  } | null> {
    try {
      const serverUrl = this.config.serverUrl || 'http://localhost:3001';
      const response = await fetch(`${serverUrl}/api/logs/stats`);
      
      if (!response.ok) return null;
      
      return await response.json();
    } catch {
      return null;
    }
  }

  /**
   * Get logs from server
   */
  async getServerLogs(options?: {
    lines?: number;
    level?: LogLevel;
    since?: string;
  }): Promise<LogEntry[] | null> {
    try {
      const serverUrl = this.config.serverUrl || 'http://localhost:3001';
      const params = new URLSearchParams();
      
      if (options?.lines) params.set('lines', String(options.lines));
      if (options?.level !== undefined) params.set('level', String(options.level));
      if (options?.since) params.set('since', options.since);
      
      const response = await fetch(`${serverUrl}/api/logs?${params}`);
      
      if (!response.ok) return null;
      
      const data = await response.json();
      return data.logs;
    } catch {
      return null;
    }
  }

  /**
   * Clear server logs
   */
  async clearServerLogs(): Promise<boolean> {
    try {
      const serverUrl = this.config.serverUrl || 'http://localhost:3001';
      const response = await fetch(`${serverUrl}/api/logs`, {
        method: 'DELETE'
      });
      
      return response.ok;
    } catch {
      return false;
    }
  }

  subscribe(callback: (entry: LogEntry) => void): () => void {
    this.subscribers.push(callback);
    return () => {
      const index = this.subscribers.indexOf(callback);
      if (index > -1) this.subscribers.splice(index, 1);
    };
  }

  getLogs(filter?: { level?: LogLevel; since?: Date }): LogEntry[] {
    let filtered = this.logs;
    if (filter?.level !== undefined) {
      filtered = filtered.filter(l => l.level >= filter.level!);
    }
    if (filter?.since) {
      filtered = filtered.filter(l => new Date(l.timestamp) >= filter.since!);
    }
    return filtered;
  }

  export(): string {
    return JSON.stringify(this.logs, null, 2);
  }

  clear(): void {
    this.logs = [];
  }
}

export const logger = Logger.getInstance();
export default logger;
