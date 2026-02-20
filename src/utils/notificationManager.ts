/**
 * Notification Manager Utility
 * 
 * Provides a centralized notification system for the Excel AI Assistant.
 * Supports multiple notification types: in-app toasts, Excel message bars,
 * and structured logging.
 * 
 * @module utils/notificationManager
 */

import { logger } from './logger';

export type NotificationType = 'info' | 'success' | 'warning' | 'error' | 'dismissed';

// Declare Excel global for TypeScript
declare const Excel: any;

export interface NotificationOptions {
  /** Duration in milliseconds before auto-dismiss (0 = persistent) */
  duration?: number;
  /** Whether to show in Excel UI */
  showInExcel?: boolean;
  /** Whether to show as toast notification */
  showAsToast?: boolean;
  /** Whether to log to structured logger */
  logToConsole?: boolean;
  /** Custom action button */
  action?: {
    label: string;
    callback: () => void;
  };
  /** Notification ID for deduplication */
  id?: string;
}

export interface Notification {
  id: string;
  type: NotificationType;
  message: string;
  title?: string;
  timestamp: Date;
  options: NotificationOptions;
}

type NotificationListener = (notification: Notification) => void;

/**
 * Notification Manager class
 * Singleton pattern for global notification handling
 */
class NotificationManager {
  private static instance: NotificationManager;
  private listeners: Set<NotificationListener> = new Set();
  private activeNotifications: Map<string, Notification> = new Map();
  private notificationHistory: Notification[] = [];
  private maxHistorySize: number = 100;

  private constructor() {}

  static getInstance(): NotificationManager {
    if (!NotificationManager.instance) {
      NotificationManager.instance = new NotificationManager();
    }
    return NotificationManager.instance;
  }

  // ============================================================================
  // Core Notification Methods
  // ============================================================================

  /**
   * Show an info notification
   */
  info(message: string, title?: string, options: NotificationOptions = {}): string {
    return this.show('info', message, title, options);
  }

  /**
   * Show a success notification
   */
  success(message: string, title?: string, options: NotificationOptions = {}): string {
    return this.show('success', message, title, { duration: 3000, ...options });
  }

  /**
   * Show a warning notification
   */
  warning(message: string, title?: string, options: NotificationOptions = {}): string {
    return this.show('warning', message, title, { duration: 5000, ...options });
  }

  /**
   * Show an error notification
   */
  error(message: string, title?: string, options: NotificationOptions = {}): string {
    return this.show('error', message, title, { duration: 0, ...options });
  }

  /**
   * Show a notification
   */
  show(
    type: NotificationType,
    message: string,
    title?: string,
    options: NotificationOptions = {}
  ): string {
    const {
      duration = 5000,
      showInExcel = false,
      showAsToast = true,
      logToConsole = true,
      id = `notif-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`,
    } = options;

    const notification: Notification = {
      id,
      type,
      message,
      title,
      timestamp: new Date(),
      options: { ...options, duration, showInExcel, showAsToast, logToConsole },
    };

    // Store notification
    this.activeNotifications.set(id, notification);
    this.addToHistory(notification);

    // Log to structured logger
    if (logToConsole) {
      this.logToStructuredLogger(notification);
    }

    // Show in Excel
    if (showInExcel) {
      this.showInExcel(notification);
    }

    // Notify listeners
    this.notifyListeners(notification);

    // Auto-dismiss after duration
    if (duration > 0) {
      setTimeout(() => {
        this.dismiss(id);
      }, duration);
    }

    return id;
  }

  // ============================================================================
  // Excel Integration
  // ============================================================================

  /**
   * Show notification in Excel UI using the Office.js API
   */
  private async showInExcel(notification: Notification): Promise<void> {
    try {
      // Try to use Excel's notification system if available
      if (typeof Excel !== 'undefined' && Excel.run) {
        await Excel.run(async (_context: any) => {
          // Use Excel's UI for notifications
          // Note: Excel doesn't have a native toast API, but we can use:
          // 1. Status bar messages (if available)
          // 2. Task pane communication
          // 3. Custom dialog
          
          // Log to structured logger instead of console
          logger.debug('Excel notification shown', { notification });
        });
      }
    } catch (error) {
      logger.error('Failed to show Excel notification', { notification }, error as Error);
    }
  }

  // ============================================================================
  // Structured Logging
  // ============================================================================

  private logToStructuredLogger(notification: Notification): void {
    const context = { 
      type: notification.type,
      title: notification.title,
      message: notification.message
    };

    switch (notification.type) {
      case 'error':
        logger.error('Notification: ' + notification.message, context);
        break;
      case 'warning':
        logger.warn('Notification: ' + notification.message, context);
        break;
      case 'success':
      case 'info':
      default:
        logger.info('Notification: ' + notification.message, context);
        break;
    }
  }

  // ============================================================================
  // Notification Management
  // ============================================================================

  /**
   * Dismiss a notification
   */
  dismiss(id: string): boolean {
    const notification = this.activeNotifications.get(id);
    if (!notification) return false;

    this.activeNotifications.delete(id);
    this.notifyListeners({ ...notification, type: 'dismissed' as NotificationType });
    return true;
  }

  /**
   * Dismiss all notifications
   */
  dismissAll(): void {
    const ids = Array.from(this.activeNotifications.keys());
    ids.forEach(id => this.dismiss(id));
  }

  /**
   * Get active notifications
   */
  getActiveNotifications(): Notification[] {
    return Array.from(this.activeNotifications.values());
  }

  /**
   * Get notification history
   */
  getHistory(): Notification[] {
    return [...this.notificationHistory];
  }

  /**
   * Clear notification history
   */
  clearHistory(): void {
    this.notificationHistory = [];
  }

  // ============================================================================
  // Listeners
  // ============================================================================

  /**
   * Subscribe to notifications
   */
  subscribe(listener: NotificationListener): () => void {
    this.listeners.add(listener);
    
    // Return unsubscribe function
    return () => {
      this.listeners.delete(listener);
    };
  }

  /**
   * Unsubscribe from notifications
   */
  unsubscribe(listener: NotificationListener): void {
    this.listeners.delete(listener);
  }

  private notifyListeners(notification: Notification): void {
    this.listeners.forEach(listener => {
      try {
        listener(notification);
      } catch (error) {
        logger.error('Notification listener error', {}, error as Error);
      }
    });
  }

  // ============================================================================
  // History Management
  // ============================================================================

  private addToHistory(notification: Notification): void {
    this.notificationHistory.push(notification);
    
    // Trim history if exceeds max size
    if (this.notificationHistory.length > this.maxHistorySize) {
      this.notificationHistory = this.notificationHistory.slice(-this.maxHistorySize);
    }
  }

  // ============================================================================
  // Utility Methods
  // ============================================================================

  /**
   * Show a notification for an async operation
   */
  async withNotification<T>(
    operation: () => Promise<T>,
    messages: {
      pending: string;
      success: string;
      error: string;
    },
    options: NotificationOptions = {}
  ): Promise<T> {
    const pendingId = this.info(messages.pending, undefined, { ...options, duration: 0 });

    try {
      const result = await operation();
      this.dismiss(pendingId);
      this.success(messages.success, undefined, options);
      return result;
    } catch (error) {
      this.dismiss(pendingId);
      this.error(messages.error + ': ' + (error instanceof Error ? error.message : String(error)), undefined, options);
      throw error;
    }
  }

  /**
   * Create a promise-based notification that resolves when dismissed
   */
  showAndWait(
    type: NotificationType,
    message: string,
    title?: string,
    options: NotificationOptions = {}
  ): Promise<void> {
    return new Promise((resolve) => {
      const id = this.show(type, message, title, { ...options, duration: 0 });
      
      const unsubscribe = this.subscribe((notif) => {
        if (notif.id === id && notif.type === 'dismissed') {
          unsubscribe();
          resolve();
        }
      });
    });
  }
}

// Export singleton instance
export const notificationManager = NotificationManager.getInstance();
export default notificationManager;
