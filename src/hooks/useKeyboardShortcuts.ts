// Keyboard Shortcuts Hook
// Provides keyboard shortcut handling for the application

import { useEffect, useCallback } from 'react';

export interface KeyboardShortcut {
  key: string;
  ctrl?: boolean;
  shift?: boolean;
  alt?: boolean;
  action: () => void;
  description?: string;
}

export interface UseKeyboardShortcutsOptions {
  enabled?: boolean;
  shortcuts: KeyboardShortcut[];
}

export const useKeyboardShortcuts = ({ 
  enabled = true, 
  shortcuts 
}: UseKeyboardShortcutsOptions): void => {
  const handleKeyDown = useCallback((event: KeyboardEvent) => {
    if (!enabled) return;

    // Don't trigger shortcuts when typing in input fields
    const target = event.target as HTMLElement;
    const isInput = target.tagName === 'INPUT' || 
                    target.tagName === 'TEXTAREA' || 
                    target.isContentEditable;

    for (const shortcut of shortcuts) {
      const ctrlMatch = shortcut.ctrl ? (event.ctrlKey || event.metaKey) : !event.ctrlKey && !event.metaKey;
      const shiftMatch = shortcut.shift ? event.shiftKey : !event.shiftKey;
      const altMatch = shortcut.alt ? event.altKey : !event.altKey;
      const keyMatch = event.key.toLowerCase() === shortcut.key.toLowerCase();

      // Allow shortcuts in inputs for some keys (like Ctrl+Enter to submit)
      if (keyMatch && ctrlMatch && shiftMatch && altMatch) {
        // If we're in an input and the shortcut requires focus, skip it
        if (isInput && !shortcut.description?.includes('input')) {
          continue;
        }
        
        event.preventDefault();
        shortcut.action();
        break;
      }
    }
  }, [enabled, shortcuts]);

  useEffect(() => {
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleKeyDown]);
};

// Pre-defined shortcuts for the AI Assistant
export const AI_ASSISTANT_SHORTCUTS: KeyboardShortcut[] = [
  {
    key: 'Enter',
    ctrl: true,
    action: () => {}, // Will be connected to send message
    description: 'Send message (in chat input)'
  },
  {
    key: 'k',
    ctrl: true,
    action: () => {}, // Will be connected to quick actions
    description: 'Open quick actions'
  },
  {
    key: 'l',
    ctrl: true,
    shift: true,
    action: () => {}, // Will be connected to clear chat
    description: 'Clear chat history'
  },
  {
    key: 's',
    ctrl: true,
    action: () => {}, // Will be connected to save/export
    description: 'Export conversation'
  },
  {
    key: 'Escape',
    action: () => {}, // Will be connected to close modals
    description: 'Close current view'
  }
];

export default useKeyboardShortcuts;
