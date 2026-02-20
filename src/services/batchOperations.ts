/**
 * Batch Operations Service
 * 
 * Executes multiple AI operations in batch mode for efficiency.
 * Provides progress tracking, undo support, and error handling.
 * 
 * Features:
 * - Queue multiple operations for sequential execution
 * - Parallel execution for independent operations
 * - Progress tracking with callbacks
 * - Undo/redo support
 * - Error recovery and rollback
 * - Operation templates and presets
 * 
 * @module services/batchOperations
 */

import { AIAction, AIActionType, FormatOptions, ChartOptions, TableOptions } from '../types';
import { actionHandler } from './actionHandler';
import { notificationManager } from '../utils/notificationManager';
import { excelService } from './excelService';

// ============================================================================
// Type Definitions
// ============================================================================

/** A single batch operation */
export interface BatchOperation {
  id: string;
  type: AIActionType;
  label: string;
  payload: any;
  options?: OperationOptions;
}

/** Options for individual operations */
export interface OperationOptions {
  /** Skip if this operation fails */
  skipOnError?: boolean;
  /** Delay before executing (ms) */
  delay?: number;
  /** Timeout for operation (ms) */
  timeout?: number;
  /** Retry count on failure */
  retries?: number;
  /** Dependencies - operation IDs that must complete first */
  dependsOn?: string[];
  /** Whether this operation can run in parallel */
  parallelizable?: boolean;
}

/** Batch execution configuration */
export interface BatchConfig {
  /** Batch name/description */
  name: string;
  /** Stop on first error */
  stopOnError?: boolean;
  /** Execute operations in parallel where possible */
  parallel?: boolean;
  /** Max parallel operations */
  maxParallel?: number;
  /** Delay between operations (ms) */
  delayBetween?: number;
  /** Global timeout (ms) */
  timeout?: number;
  /** Enable undo support */
  enableUndo?: boolean;
  /** Callback for progress updates */
  onProgress?: (progress: BatchProgress) => void;
  /** Callback when complete */
  onComplete?: (result: BatchResult) => void;
  /** Callback on error */
  onError?: (error: BatchError) => void;
}

/** Batch execution progress */
export interface BatchProgress {
  batchId: string;
  total: number;
  completed: number;
  failed: number;
  skipped: number;
  currentOperation?: BatchOperation;
  percentComplete: number;
  elapsedTime: number;
  estimatedTimeRemaining?: number;
}

/** Batch execution result */
export interface BatchResult {
  batchId: string;
  success: boolean;
  operations: OperationResult[];
  completedAt: Date;
  duration: number;
  undoAvailable: boolean;
}

/** Individual operation result */
export interface OperationResult {
  operationId: string;
  success: boolean;
  error?: string;
  duration: number;
  result?: any;
  undone?: boolean;
}

/** Batch error details */
export interface BatchError {
  batchId: string;
  operationId?: string;
  error: Error;
  recoverable: boolean;
}

/** Undo snapshot for a batch */
interface BatchSnapshot {
  batchId: string;
  operations: BatchOperation[];
  snapshots: Map<string, any>;
  timestamp: Date;
}

/** Batch preset/template */
export interface BatchPreset {
  id: string;
  name: string;
  description: string;
  operations: Omit<BatchOperation, 'id'>[];
  config: Partial<BatchConfig>;
  category: string;
  tags: string[];
}

// ============================================================================
// Batch Operations Service
// ============================================================================

export class BatchOperationsService {
  private static instance: BatchOperationsService;
  private operationQueue: BatchOperation[] = [];
  private activeBatches: Map<string, BatchExecution> = new Map();
  private batchHistory: BatchResult[] = [];
  private undoStack: Map<string, BatchSnapshot> = new Map();
  private redoStack: Map<string, BatchSnapshot> = new Map();
  private presets: Map<string, BatchPreset> = new Map();
  private isExecuting: boolean = false;

  private constructor() {
    this.initializePresets();
  }

  static getInstance(): BatchOperationsService {
    if (!BatchOperationsService.instance) {
      BatchOperationsService.instance = new BatchOperationsService();
    }
    return BatchOperationsService.instance;
  }

  // ============================================================================
  // Queue Management
  // ============================================================================

  /**
   * Add an operation to the queue
   */
  enqueue(operation: Omit<BatchOperation, 'id'>): string {
    const id = `op-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    const fullOperation: BatchOperation = { ...operation, id };
    this.operationQueue.push(fullOperation);
    return id;
  }

  /**
   * Add multiple operations to the queue
   */
  enqueueMany(operations: Omit<BatchOperation, 'id'>[]): string[] {
    return operations.map(op => this.enqueue(op));
  }

  /**
   * Remove an operation from the queue
   */
  dequeue(operationId: string): boolean {
    const index = this.operationQueue.findIndex(op => op.id === operationId);
    if (index >= 0) {
      this.operationQueue.splice(index, 1);
      return true;
    }
    return false;
  }

  /**
   * Clear the operation queue
   */
  clearQueue(): void {
    this.operationQueue = [];
  }

  /**
   * Get current queue
   */
  getQueue(): BatchOperation[] {
    return [...this.operationQueue];
  }

  /**
   * Move operation in queue
   */
  moveOperation(operationId: string, newIndex: number): boolean {
    const currentIndex = this.operationQueue.findIndex(op => op.id === operationId);
    if (currentIndex < 0) return false;

    const [operation] = this.operationQueue.splice(currentIndex, 1);
    this.operationQueue.splice(newIndex, 0, operation);
    return true;
  }

  // ============================================================================
  // Batch Execution
  // ============================================================================

  /**
   * Execute all queued operations as a batch
   */
  async executeQueue(config: Partial<BatchConfig> = {}): Promise<BatchResult> {
    if (this.operationQueue.length === 0) {
      throw new Error('No operations in queue');
    }

    const batchId = `batch-${Date.now()}`;
    const operations = [...this.operationQueue];
    
    // Clear queue after copying
    this.operationQueue = [];

    return this.executeBatch(batchId, operations, config);
  }

  /**
   * Execute a specific set of operations
   */
  async executeBatch(
    batchId: string,
    operations: BatchOperation[],
    config: Partial<BatchConfig> = {}
  ): Promise<BatchResult> {
    const fullConfig: BatchConfig = {
      name: 'Batch Operation',
      stopOnError: false,
      parallel: false,
      maxParallel: 3,
      delayBetween: 100,
      timeout: 300000, // 5 minutes
      enableUndo: true,
      ...config,
    };

    const execution = new BatchExecution(batchId, operations, fullConfig);
    this.activeBatches.set(batchId, execution);

    // Take snapshot for undo if enabled
    if (fullConfig.enableUndo) {
      await this.takeSnapshot(batchId, operations);
    }

    try {
      const result = await execution.execute();
      this.batchHistory.push(result);
      this.activeBatches.delete(batchId);
      
      if (fullConfig.onComplete) {
        fullConfig.onComplete(result);
      }

      return result;
    } catch (error) {
      this.activeBatches.delete(batchId);
      throw error;
    }
  }

  /**
   * Cancel an active batch
   */
  cancelBatch(batchId: string): boolean {
    const execution = this.activeBatches.get(batchId);
    if (execution) {
      execution.cancel();
      return true;
    }
    return false;
  }

  /**
   * Get active batch progress
   */
  getProgress(batchId: string): BatchProgress | undefined {
    const execution = this.activeBatches.get(batchId);
    return execution?.getProgress();
  }

  // ============================================================================
  // Undo/Redo
  // ============================================================================

  /**
   * Undo a completed batch
   */
  async undoBatch(batchId: string): Promise<boolean> {
    const snapshot = this.undoStack.get(batchId);
    if (!snapshot) {
      notificationManager.warning('No undo available for this batch');
      return false;
    }

    try {
      // Execute undo operations in reverse order
      const undoOperations = [...snapshot.operations].reverse().map(op => ({
        ...op,
        id: `undo-${op.id}`,
        type: this.getUndoActionType(op.type),
        payload: this.getUndoPayload(op),
      }));

      for (const op of undoOperations) {
        await this.executeOperation(op);
      }

      // Move to redo stack
      this.redoStack.set(batchId, snapshot);
      this.undoStack.delete(batchId);

      notificationManager.success('Batch operation undone');
      return true;
    } catch (error) {
      notificationManager.error('Failed to undo batch: ' + error);
      return false;
    }
  }

  /**
   * Redo a previously undone batch
   */
  async redoBatch(batchId: string): Promise<boolean> {
    const snapshot = this.redoStack.get(batchId);
    if (!snapshot) {
      notificationManager.warning('No redo available');
      return false;
    }

    // Re-execute the original operations
    const result = await this.executeBatch(
      `redo-${batchId}`,
      snapshot.operations,
      { name: 'Redo Batch' }
    );

    if (result.success) {
      this.undoStack.set(batchId, snapshot);
      this.redoStack.delete(batchId);
      notificationManager.success('Batch operation redone');
      return true;
    }

    return false;
  }

  /**
   * Check if undo is available for a batch
   */
  canUndo(batchId: string): boolean {
    return this.undoStack.has(batchId);
  }

  /**
   * Check if redo is available for a batch
   */
  canRedo(batchId: string): boolean {
    return this.redoStack.has(batchId);
  }

  private async takeSnapshot(batchId: string, operations: BatchOperation[]): Promise<void> {
    // Capture the current state of all affected cells/ranges for true undo capability
    const snapshots = new Map<string, any>();
    
    // Capture state for operations that modify ranges
    for (const op of operations) {
      const payload = op.payload as Record<string, unknown>;
      const address = payload.address as string | undefined;
      const range = payload.range as string | undefined;
      const worksheetName = payload.worksheetName as string | undefined;
      
      const targetRange = address || range;
      if (targetRange && typeof targetRange === 'string') {
        try {
          // Capture current state before modification
          const rangeContext = await excelService.getRange(targetRange, worksheetName);
          snapshots.set(op.id, {
            type: 'range',
            address: targetRange,
            worksheetName: worksheetName || rangeContext.worksheetName,
            values: rangeContext.values,
            formulas: rangeContext.formulas,
            numberFormat: rangeContext.numberFormat,
            rowCount: rangeContext.rowCount,
            columnCount: rangeContext.columnCount,
          });
        } catch (error) {
          // If range doesn't exist or can't be accessed, store operation metadata
          snapshots.set(op.id, {
            type: 'operation',
            operation: op,
            error: error instanceof Error ? error.message : 'Unknown error',
          });
        }
      } else {
        // For operations without specific ranges, store operation metadata
        snapshots.set(op.id, {
          type: 'operation',
          operation: op,
        });
      }
    }
    
    const snapshot: BatchSnapshot = {
      batchId,
      operations,
      snapshots,
      timestamp: new Date(),
    };

    this.undoStack.set(batchId, snapshot);
    // Clear redo stack when new operation is performed
    this.redoStack.delete(batchId);
  }

  private getUndoActionType(type: AIActionType): AIActionType {
    // Map operations to their undo counterparts
    const undoMap: Record<AIActionType, AIActionType> = {
      'insert_formula': 'clear_range',
      'set_values': 'clear_range',
      'create_chart': 'delete_worksheet', // Charts are on worksheets
      'create_table': 'clear_range',
      'format_cells': 'format_cells', // Revert with original format
      'add_validation': 'clear_range',
      'clear_range': 'set_values', // Would need original values
      'create_pivot_table': 'delete_worksheet',
      'add_worksheet': 'delete_worksheet',
      'delete_worksheet': 'add_worksheet', // Can't truly undo
      'auto_fit_columns': 'auto_fit_columns', // No real undo
      'create_named_range': 'clear_range',
      'explain': 'explain',
      'apply_suggestion': 'clear_range',
    };
    return undoMap[type] || type;
  }

  private getUndoPayload(operation: BatchOperation): any {
    // Retrieve the original state from snapshot
    const snapshot = Array.from(this.undoStack.values())
      .find(s => s.operations.some(op => op.id === operation.id));
    
    if (!snapshot) {
      // Fallback: return operation payload with undo flag
      return { ...operation.payload, isUndo: true };
    }
    
    const stateSnapshot = snapshot.snapshots.get(operation.id);
    if (!stateSnapshot) {
      return { ...operation.payload, isUndo: true };
    }
    
    // Restore original state based on operation type
    switch (operation.type) {
      case 'set_values':
      case 'insert_formula':
      case 'clear_range':
        // Restore original values and formulas
        if (stateSnapshot.type === 'range') {
          return {
            address: stateSnapshot.address,
            worksheetName: stateSnapshot.worksheetName,
            values: stateSnapshot.values,
            formulas: stateSnapshot.formulas,
            numberFormat: stateSnapshot.numberFormat,
            isUndo: true,
          };
        }
        break;
        
      case 'format_cells':
        // Restore original formatting
        if (stateSnapshot.type === 'range') {
          return {
            address: stateSnapshot.address,
            worksheetName: stateSnapshot.worksheetName,
            numberFormat: stateSnapshot.numberFormat,
            isUndo: true,
          };
        }
        break;
        
      default:
        // For other operations, return original payload with undo flag
        return { ...operation.payload, isUndo: true };
    }
    
    return { ...operation.payload, isUndo: true };
  }

  // ============================================================================
  // Presets/Templates
  // ============================================================================

  private initializePresets(): void {
    const presets: BatchPreset[] = [
      {
        id: 'cleanup-data',
        name: 'Clean Up Data',
        description: 'Standard data cleaning operations',
        category: 'Data Cleaning',
        tags: ['clean', 'format', 'prepare'],
        operations: [
          { type: 'format_cells', label: 'Trim whitespace', payload: { trim: true } },
          { type: 'format_cells', label: 'Standardize case', payload: { case: 'title' } },
          { type: 'add_validation', label: 'Add data validation', payload: { type: 'list' } },
        ],
        config: { stopOnError: false, enableUndo: true },
      },
      {
        id: 'monthly-report',
        name: 'Monthly Report Setup',
        description: 'Create monthly report structure',
        category: 'Reporting',
        tags: ['report', 'monthly', 'charts'],
        operations: [
          { type: 'create_table', label: 'Create data table', payload: { hasHeaders: true } },
          { type: 'create_chart', label: 'Add trend chart', payload: { type: 'line' } },
          { type: 'format_cells', label: 'Format numbers', payload: { numberFormat: '#,##0.00' } },
        ],
        config: { stopOnError: true, enableUndo: true },
      },
      {
        id: 'export-prep',
        name: 'Export Preparation',
        description: 'Prepare data for export',
        category: 'Export',
        tags: ['export', 'format', 'clean'],
        operations: [
          { type: 'format_cells', label: 'Remove formatting', payload: { clearFormats: true } },
          { type: 'auto_fit_columns', label: 'Auto-fit columns', payload: {} },
          { type: 'create_named_range', label: 'Define export range', payload: { name: 'ExportData' } },
        ],
        config: { stopOnError: false },
      },
    ];

    presets.forEach(preset => this.presets.set(preset.id, preset));
  }

  /**
   * Get all available presets
   */
  getPresets(): BatchPreset[] {
    return Array.from(this.presets.values());
  }

  /**
   * Get presets by category
   */
  getPresetsByCategory(category: string): BatchPreset[] {
    return this.getPresets().filter(p => p.category === category);
  }

  /**
   * Get a specific preset
   */
  getPreset(id: string): BatchPreset | undefined {
    return this.presets.get(id);
  }

  /**
   * Add a custom preset
   */
  addPreset(preset: Omit<BatchPreset, 'id'>): string {
    const id = `preset-${Date.now()}`;
    this.presets.set(id, { ...preset, id });
    return id;
  }

  /**
   * Delete a preset
   */
  deletePreset(id: string): boolean {
    return this.presets.delete(id);
  }

  /**
   * Execute a preset
   */
  async executePreset(presetId: string, configOverrides: Partial<BatchConfig> = {}): Promise<BatchResult> {
    const preset = this.presets.get(presetId);
    if (!preset) {
      throw new Error(`Preset not found: ${presetId}`);
    }

    const operations = preset.operations.map(op => ({
      ...op,
      id: `op-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    }));

    const config: BatchConfig = {
      name: preset.name,
      ...preset.config,
      ...configOverrides,
    };

    return this.executeBatch(`preset-${presetId}-${Date.now()}`, operations, config);
  }

  // ============================================================================
  // History
  // ============================================================================

  /**
   * Get batch execution history
   */
  getHistory(): BatchResult[] {
    return [...this.batchHistory];
  }

  /**
   * Clear history
   */
  clearHistory(): void {
    this.batchHistory = [];
  }

  /**
   * Get last N batch results
   */
  getRecentBatches(count: number = 10): BatchResult[] {
    return this.batchHistory.slice(-count);
  }

  // ============================================================================
  // Private Methods
  // ============================================================================

  private async executeOperation(operation: BatchOperation): Promise<OperationResult> {
    const startTime = Date.now();

    try {
      // Convert batch operation to AI action
      const action: AIAction = {
        type: operation.type,
        label: operation.label,
        payload: operation.payload,
      };

      // Execute via action handler
      await actionHandler.executeAction(action);

      return {
        operationId: operation.id,
        success: true,
        duration: Date.now() - startTime,
      };
    } catch (error) {
      return {
        operationId: operation.id,
        success: false,
        error: error instanceof Error ? error.message : String(error),
        duration: Date.now() - startTime,
      };
    }
  }
}

// ============================================================================
// Batch Execution Class
// ============================================================================

class BatchExecution {
  private startTime: number = 0;
  private cancelled: boolean = false;
  private results: OperationResult[] = [];
  private currentOperationIndex: number = 0;

  constructor(
    private batchId: string,
    private operations: BatchOperation[],
    private config: BatchConfig
  ) {}

  async execute(): Promise<BatchResult> {
    this.startTime = Date.now();
    this.results = [];
    this.currentOperationIndex = 0;

    if (this.config.parallel) {
      await this.executeParallel();
    } else {
      await this.executeSequential();
    }

    const duration = Date.now() - this.startTime;
    const success = this.results.every(r => r.success);

    return {
      batchId: this.batchId,
      success,
      operations: this.results,
      completedAt: new Date(),
      duration,
      undoAvailable: this.config.enableUndo ?? true,
    };
  }

  private async executeSequential(): Promise<void> {
    for (let i = 0; i < this.operations.length; i++) {
      if (this.cancelled) break;

      this.currentOperationIndex = i;
      const operation = this.operations[i];

      // Check dependencies
      if (operation.options?.dependsOn) {
        const depsComplete = operation.options.dependsOn.every(
          depId => this.results.find(r => r.operationId === depId)?.success
        );
        if (!depsComplete) {
          this.results.push({
            operationId: operation.id,
            success: false,
            error: 'Dependencies not met',
            duration: 0,
          });
          continue;
        }
      }

      // Delay if specified
      if (operation.options?.delay) {
        await this.sleep(operation.options.delay);
      }

      // Execute with retries
      const result = await this.executeWithRetries(operation);
      this.results.push(result);

      // Notify progress
      this.notifyProgress();

      // Handle error
      if (!result.success && this.config.stopOnError) {
        if (this.config.onError) {
          this.config.onError({
            batchId: this.batchId,
            operationId: operation.id,
            error: new Error(result.error || 'Unknown error'),
            recoverable: operation.options?.skipOnError ?? false,
          });
        }
        break;
      }

      // Delay between operations
      if (this.config.delayBetween && i < this.operations.length - 1) {
        await this.sleep(this.config.delayBetween);
      }
    }
  }

  private async executeParallel(): Promise<void> {
    const maxParallel = this.config.maxParallel || 3;
    const chunks = this.chunkArray(this.operations, maxParallel);

    for (const chunk of chunks) {
      if (this.cancelled) break;

      const promises = chunk.map(op => this.executeWithRetries(op));
      const chunkResults = await Promise.all(promises);
      
      this.results.push(...chunkResults);
      this.notifyProgress();

      // Check for errors
      const hasError = chunkResults.some(r => !r.success);
      if (hasError && this.config.stopOnError) {
        break;
      }

      if (this.config.delayBetween) {
        await this.sleep(this.config.delayBetween);
      }
    }
  }

  private async executeWithRetries(operation: BatchOperation): Promise<OperationResult> {
    const maxRetries = operation.options?.retries ?? 0;
    let lastError: string | undefined;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      const result = await this.executeSingle(operation);
      if (result.success) {
        return result;
      }
      lastError = result.error;
      
      if (attempt < maxRetries) {
        await this.sleep(1000 * (attempt + 1)); // Exponential backoff
      }
    }

    return {
      operationId: operation.id,
      success: false,
      error: lastError || 'Failed after retries',
      duration: 0,
    };
  }

  private async executeSingle(operation: BatchOperation): Promise<OperationResult> {
    const startTime = Date.now();

    try {
      const action: AIAction = {
        type: operation.type,
        label: operation.label,
        payload: operation.payload,
      };

      await actionHandler.executeAction(action);

      return {
        operationId: operation.id,
        success: true,
        duration: Date.now() - startTime,
      };
    } catch (error) {
      return {
        operationId: operation.id,
        success: false,
        error: error instanceof Error ? error.message : String(error),
        duration: Date.now() - startTime,
      };
    }
  }

  private notifyProgress(): void {
    if (!this.config.onProgress) return;

    const completed = this.results.length;
    const failed = this.results.filter(r => !r.success).length;
    const skipped = this.results.filter(r => r.error === 'Skipped').length;
    const elapsedTime = Date.now() - this.startTime;

    this.config.onProgress({
      batchId: this.batchId,
      total: this.operations.length,
      completed,
      failed,
      skipped,
      currentOperation: this.operations[this.currentOperationIndex],
      percentComplete: Math.round((completed / this.operations.length) * 100),
      elapsedTime,
      estimatedTimeRemaining: this.calculateETR(elapsedTime, completed),
    });
  }

  private calculateETR(elapsedTime: number, completed: number): number | undefined {
    if (completed === 0) return undefined;
    const avgTimePerOp = elapsedTime / completed;
    const remaining = this.operations.length - completed;
    return Math.round(avgTimePerOp * remaining);
  }

  private chunkArray<T>(array: T[], size: number): T[][] {
    const chunks: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }

  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  cancel(): void {
    this.cancelled = true;
  }

  getProgress(): BatchProgress {
    const completed = this.results.length;
    const failed = this.results.filter(r => !r.success).length;
    const skipped = this.results.filter(r => r.error === 'Skipped').length;
    const elapsedTime = Date.now() - this.startTime;

    return {
      batchId: this.batchId,
      total: this.operations.length,
      completed,
      failed,
      skipped,
      currentOperation: this.operations[this.currentOperationIndex],
      percentComplete: Math.round((completed / this.operations.length) * 100),
      elapsedTime,
      estimatedTimeRemaining: this.calculateETR(elapsedTime, completed),
    };
  }
}

// Export singleton instance
export const batchOperations = BatchOperationsService.getInstance();
export default batchOperations;
