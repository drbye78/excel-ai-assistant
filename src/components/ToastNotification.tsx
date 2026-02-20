// Toast Notification System - User feedback for actions
// Provides success, error, info, and warning messages

import React, { createContext, useContext, useState, useCallback } from 'react';
import { Stack, MessageBar, MessageBarType, IconButton } from '@fluentui/react';

export type ToastIntent = 'success' | 'error' | 'info' | 'warning';

export interface Toast {
  id: string;
  message: string;
  intent: ToastIntent;
  duration?: number;
  dismissible?: boolean;
}

interface ToastContextType {
  toasts: Toast[];
  success: (message: string, duration?: number) => void;
  error: (message: string, duration?: number) => void;
  info: (message: string, duration?: number) => void;
  warning: (message: string, duration?: number) => void;
  removeToast: (id: string) => void;
}

const ToastContext = createContext<ToastContextType | undefined>(undefined);

export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [toasts, setToasts] = useState<Toast[]>([]);

  const addToast = useCallback((toast: Omit<Toast, 'id'>) => {
    const id = `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    const newToast: Toast = { ...toast, id, dismissible: toast.dismissible ?? true };
    
    setToasts((prev) => [...prev, newToast]);

    // Auto-dismiss after duration
    if (toast.duration !== 0) {
      setTimeout(() => {
        removeToast(id);
      }, toast.duration || 5000);
    }
  }, []);

  const success = useCallback((message: string, duration?: number) => {
    addToast({ message, intent: 'success', duration });
  }, [addToast]);

  const error = useCallback((message: string, duration?: number) => {
    addToast({ message, intent: 'error', duration });
  }, [addToast]);

  const info = useCallback((message: string, duration?: number) => {
    addToast({ message, intent: 'info', duration });
  }, [addToast]);

  const warning = useCallback((message: string, duration?: number) => {
    addToast({ message, intent: 'warning', duration });
  }, [addToast]);

  const removeToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);

  return (
    <ToastContext.Provider value={{ toasts, success, error, info, warning, removeToast }}>
      {children}
      <ToastContainer toasts={toasts} onDismiss={removeToast} />
    </ToastContext.Provider>
  );
};

export const useToast = (): ToastContextType => {
  const context = useContext(ToastContext);
  if (!context) {
    throw new Error('useToast must be used within a ToastProvider');
  }
  return context;
};

// Toast Container Component
interface ToastContainerProps {
  toasts: Toast[];
  onDismiss: (id: string) => void;
}

const ToastContainer: React.FC<ToastContainerProps> = ({ toasts, onDismiss }) => {
  if (toasts.length === 0) return null;

  return (
    <Stack
      styles={{
        root: {
          position: 'fixed',
          top: 16,
          right: 16,
          zIndex: 1000,
          maxWidth: 400,
          gap: 8,
        },
      }}
    >
      {toasts.map((toast) => (
        <ToastItem key={toast.id} toast={toast} onDismiss={onDismiss} />
      ))}
    </Stack>
  );
};

// Individual Toast Item
interface ToastItemProps {
  toast: Toast;
  onDismiss: (id: string) => void;
}

const ToastItem: React.FC<ToastItemProps> = ({ toast, onDismiss }) => {
  const getMessageBarType = (intent: ToastIntent): MessageBarType => {
    switch (intent) {
      case 'success':
        return MessageBarType.success;
      case 'error':
        return MessageBarType.error;
      case 'warning':
        return MessageBarType.warning;
      case 'info':
      default:
        return MessageBarType.info;
    }
  };

  const getIconName = (intent: ToastIntent): string => {
    switch (intent) {
      case 'success':
        return 'CheckMark';
      case 'error':
        return 'Error';
      case 'warning':
        return 'Warning';
      case 'info':
      default:
        return 'Info';
    }
  };

  return (
    <MessageBar
      messageBarType={getMessageBarType(toast.intent)}
      isMultiline={false}
      onDismiss={toast.dismissible ? () => onDismiss(toast.id) : undefined}
      dismissButtonAriaLabel="Dismiss notification"
      styles={{
        root: {
          boxShadow: '0 4px 12px rgba(0, 0, 0, 0.15)',
          borderRadius: 4,
          animation: 'slideIn 0.3s ease-out',
        },
      }}
    >
      {toast.message}
    </MessageBar>
  );
};

// Convenience hook for common toast patterns
export const useToastHelpers = () => {
  const toast = useToast();

  const notifyAsyncOperation = async <T,>(
    operation: () => Promise<T>,
    {
      loadingMessage = 'Processing...',
      successMessage = 'Completed successfully',
      errorMessage = 'Operation failed',
    }: {
      loadingMessage?: string;
      successMessage?: string | ((result: T) => string);
      errorMessage?: string | ((error: Error) => string);
    } = {}
  ): Promise<T | undefined> => {
    const loadingToastId = `${Date.now()}-loading`;
    
    // Show loading toast (0 duration = don't auto-dismiss)
    toast.info(loadingMessage, 0);

    try {
      const result = await operation();
      
      // Dismiss loading toast and show success
      const message = typeof successMessage === 'function' 
        ? successMessage(result) 
        : successMessage;
      toast.success(message);
      
      return result;
    } catch (error) {
      const message = typeof errorMessage === 'function' && error instanceof Error
        ? errorMessage(error)
        : errorMessage;
      toast.error(message);
      throw error;
    }
  };

  return { notifyAsyncOperation, ...toast };
};

export default ToastProvider;
