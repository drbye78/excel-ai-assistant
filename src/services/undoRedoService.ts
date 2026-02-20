/**
 * Undo/Redo Service
 * Phase 7: Feature Enhancements
 * 
 * Provides undo/redo functionality for AI actions
 */

import { AIAction } from '@/types';

export interface UndoableOperation {
  id: string;
  action: AIAction;
  reverseAction: AIAction;
  timestamp: Date;
  description: string;
  executed: boolean;
}

export interface UndoRedoState {
  canUndo: boolean;
  canRedo: boolean;
  undoCount: number;
  redoCount: number;
  historySize: number;
}

export class UndoRedoService {
  private static instance: UndoRedoService;
  
  private undoStack: UndoableOperation[] = [];
  private redoStack: UndoableOperation[] = [];
  private maxHistorySize = 50;
  
  private constructor() {}

  static getInstance(): UndoRedoService {
    if (!UndoRedoService.instance) {
      UndoRedoService.instance = new UndoRedoService();
    }
    return UndoRedoService.instance;
  }

  /**
   * Set maximum history size
   */
  setMaxHistorySize(size: number): void {
    this.maxHistorySize = size;
    
    // Trim if current size exceeds new max
    if (this.undoStack.length > size) {
      this.undoStack = this.undoStack.slice(-size);
    }
  }

  /**
   * Get reverse action for an action
   */
  private getReverseAction(action: AIAction): AIAction | null {
    const payload = action.payload as Record<string, any>;
    
    switch (action.type) {
      case 'set_values': {
        // To undo setting values, we need the previous values
        // This would need to be captured before the action
        return {
          type: 'set_values',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            values: payload?.previousValues || [],
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'clear_range': {
        // To undo clearing, restore previous values
        return {
          type: 'set_values',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            values: payload?.previousValues || [],
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'format_cells': {
        // To undo formatting, restore previous format
        return {
          type: 'format_cells',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            previousFormat: payload?.previousFormat,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'create_table': {
        // To undo creating a table, delete it
        return {
          type: 'delete_table',
          label: `Undo: ${action.label}`,
          payload: {
            name: payload?.name,
            keepData: true,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'delete_table': {
        // To undo deleting a table, recreate it (if we have the data)
        if (payload?.tableData) {
          return {
            type: 'create_table',
            label: `Undo: ${action.label}`,
            payload: {
              range: payload.tableData.range,
              name: payload.tableData.name,
              hasHeaders: payload.tableData.hasHeaders,
              style: payload.tableData.style,
              worksheetName: payload?.worksheetName
            }
          };
        }
        return null;
      }
      
      case 'create_chart': {
        // To undo creating a chart, delete it
        return {
          type: 'delete_chart',
          label: `Undo: ${action.label}`,
          payload: {
            name: payload?.chartName || payload?.name,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'delete_chart': {
        // Cannot easily undo chart deletion
        return null;
      }
      
      case 'add_worksheet': {
        // To undo adding a worksheet, delete it
        return {
          type: 'delete_worksheet',
          label: `Undo: ${action.label}`,
          payload: {
            name: payload?.createdName || payload?.name
          }
        };
      }
      
      case 'delete_worksheet': {
        // Cannot easily undo worksheet deletion
        return null;
      }
      
      case 'create_named_range': {
        // To undo creating a named range, delete it
        return {
          type: 'delete_named_range',
          label: `Undo: ${action.label}`,
          payload: {
            name: payload?.name
          }
        };
      }
      
      case 'delete_named_range': {
        // To undo deleting a named range, recreate it
        if (payload?.rangeData) {
          return {
            type: 'create_named_range',
            label: `Undo: ${action.label}`,
            payload: {
              name: payload.rangeData.name,
              address: payload.rangeData.address,
              worksheetName: payload.rangeData.worksheetName
            }
          };
        }
        return null;
      }
      
      case 'add_validation': {
        // To undo adding validation, remove it
        return {
          type: 'remove_validation',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'add_comment': {
        // To undo adding a comment, delete it
        return {
          type: 'delete_comment',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      case 'add_hyperlink': {
        // To undo adding a hyperlink, remove it
        return {
          type: 'remove_hyperlink',
          label: `Undo: ${action.label}`,
          payload: {
            address: payload?.address,
            worksheetName: payload?.worksheetName
          }
        };
      }
      
      default:
        return null;
    }
  }

  /**
   * Register an action for potential undo
   * @param action The action that was executed
   * @param description Human-readable description
   * @param previousState State before action (for restoration)
   */
  registerAction(
    action: AIAction,
    description: string,
    previousState?: Record<string, unknown>
  ): UndoableOperation | null {
    const reverseAction = this.getReverseAction(action);
    
    if (!reverseAction) {
      // This action cannot be undone
      return null;
    }
    
    // Add previous state to reverse action if available
    if (previousState) {
      reverseAction.payload = {
        ...reverseAction.payload,
        ...previousState
      };
    }
    
    const operation: UndoableOperation = {
      id: `op_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      action,
      reverseAction,
      timestamp: new Date(),
      description,
      executed: true
    };
    
    // Clear redo stack when new action is registered
    this.redoStack = [];
    
    // Add to undo stack
    this.undoStack.push(operation);
    
    // Trim to max size
    if (this.undoStack.length > this.maxHistorySize) {
      this.undoStack.shift();
    }
    
    return operation;
  }

  /**
   * Get the next undo operation
   */
  peekUndo(): UndoableOperation | null {
    if (this.undoStack.length === 0) {
      return null;
    }
    return this.undoStack[this.undoStack.length - 1];
  }

  /**
   * Get the next redo operation
   */
  peekRedo(): UndoableOperation | null {
    if (this.redoStack.length === 0) {
      return null;
    }
    return this.redoStack[this.redoStack.length - 1];
  }

  /**
   * Pop the next undo operation
   */
  popUndo(): UndoableOperation | null {
    if (this.undoStack.length === 0) {
      return null;
    }
    
    const operation = this.undoStack.pop()!;
    
    // Move to redo stack
    operation.executed = false;
    this.redoStack.push(operation);
    
    return operation;
  }

  /**
   * Pop the next redo operation
   */
  popRedo(): UndoableOperation | null {
    if (this.redoStack.length === 0) {
      return null;
    }
    
    const operation = this.redoStack.pop()!;
    
    // Move back to undo stack
    operation.executed = true;
    this.undoStack.push(operation);
    
    return operation;
  }

  /**
   * Get current state
   */
  getState(): UndoRedoState {
    return {
      canUndo: this.undoStack.length > 0,
      canRedo: this.redoStack.length > 0,
      undoCount: this.undoStack.length,
      redoCount: this.redoStack.length,
      historySize: this.maxHistorySize
    };
  }

  /**
   * Get all undo operations
   */
  getUndoHistory(): UndoableOperation[] {
    return [...this.undoStack].reverse();
  }

  /**
   * Get all redo operations
   */
  getRedoHistory(): UndoableOperation[] {
    return [...this.redoStack].reverse();
  }

  /**
   * Clear all history
   */
  clearHistory(): void {
    this.undoStack = [];
    this.redoStack = [];
  }

  /**
   * Check if an action type is reversible
   */
  isReversible(actionType: string): boolean {
    const reversibleTypes = [
      'set_values',
      'clear_range',
      'format_cells',
      'create_table',
      'delete_table',
      'create_chart',
      'add_worksheet',
      'create_named_range',
      'delete_named_range',
      'add_validation',
      'add_comment',
      'add_hyperlink'
    ];
    
    const nonReversibleTypes = [
      'delete_chart',
      'delete_worksheet',
      'explain',
      'apply_suggestion'
    ];
    
    if (nonReversibleTypes.includes(actionType)) {
      return false;
    }
    
    return reversibleTypes.includes(actionType);
  }

  /**
   * Serialize state for storage
   */
  serialize(): string {
    return JSON.stringify({
      undoStack: this.undoStack,
      redoStack: this.redoStack,
      maxHistorySize: this.maxHistorySize
    });
  }

  /**
   * Deserialize state from storage
   */
  deserialize(data: string): void {
    try {
      const parsed = JSON.parse(data);
      this.undoStack = (parsed.undoStack || []).map((op: any) => ({
        ...op,
        timestamp: new Date(op.timestamp)
      }));
      this.redoStack = (parsed.redoStack || []).map((op: any) => ({
        ...op,
        timestamp: new Date(op.timestamp)
      }));
      this.maxHistorySize = parsed.maxHistorySize || 50;
    } catch (error) {
      console.error('Failed to deserialize undo/redo state:', error);
    }
  }
}

export default UndoRedoService.getInstance();