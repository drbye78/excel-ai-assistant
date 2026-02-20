/**
 * BatchOperationsPanel Component
 * 
 * UI panel for managing batch operations in Excel.
 * Allows users to queue multiple AI operations and execute them together.
 * 
 * @module BatchOperationsPanel
 */

import React, { useState, useEffect, useCallback, useMemo } from 'react';
import {
  Stack,
  Text,
  IconButton,
  PrimaryButton,
  DefaultButton,
  Separator,
  TooltipHost,
  Toggle,
  Dropdown,
  IDropdownOption,
  TextField,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Pivot,
  PivotItem,
  Label,
  ChoiceGroup,
  IChoiceGroupOption,
  Spinner,
  SpinnerSize,
  ScrollablePane,
  ScrollbarVisibility,
} from '@fluentui/react';
import {
  AddIcon,
  PlayIcon,
  DeleteIcon,
  ClearIcon,
  UndoIcon,
  RedoIcon,
  SettingsIcon,
  HistoryIcon,
  SaveIcon,
  TaskListIcon,
} from '@fluentui/react-icons-mdl2';
import {
  batchOperations,
  BatchOperation,
  BatchConfig,
  BatchProgress,
  BatchResult,
  BatchPreset,
  OperationResult,
} from '../services/batchOperations';
import { AIActionType } from '../types';
import { notificationManager } from '../utils/notificationManager';

interface BatchOperationsPanelProps {
  /** Callback when batch is completed */
  onBatchComplete?: (result: BatchResult) => void;
}

type PanelView = 'queue' | 'presets' | 'history' | 'settings';

const actionTypeOptions: IDropdownOption[] = [
  { key: 'insert_formula', text: 'Insert Formula' },
  { key: 'set_values', text: 'Set Values' },
  { key: 'format_cells', text: 'Format Cells' },
  { key: 'create_table', text: 'Create Table' },
  { key: 'create_chart', text: 'Create Chart' },
  { key: 'create_pivot_table', text: 'Create Pivot Table' },
  { key: 'add_validation', text: 'Add Validation' },
  { key: 'clear_range', text: 'Clear Range' },
  { key: 'auto_fit_columns', text: 'Auto-fit Columns' },
  { key: 'create_named_range', text: 'Create Named Range' },
  { key: 'add_worksheet', text: 'Add Worksheet' },
  { key: 'delete_worksheet', text: 'Delete Worksheet' },
];

const categoryColors: Record<string, string> = {
  'Data Cleaning': '#107c10',
  'Reporting': '#0078d4',
  'Export': '#ffc107',
  'Analysis': '#673ab7',
  'Formatting': '#ff5722',
};

export const BatchOperationsPanel: React.FC<BatchOperationsPanelProps> = ({
  onBatchComplete,
}) => {
  const [activeView, setActiveView] = useState<PanelView>('queue');
  const [queue, setQueue] = useState<BatchOperation[]>([]);
  const [isExecuting, setIsExecuting] = useState(false);
  const [progress, setProgress] = useState<BatchProgress | null>(null);
  const [results, setResults] = useState<BatchResult | null>(null);
  const [presets, setPresets] = useState<BatchPreset[]>([]);
  const [history, setHistory] = useState<BatchResult[]>([]);
  
  // New operation form state
  const [selectedActionType, setSelectedActionType] = useState<AIActionType>('format_cells');
  const [operationLabel, setOperationLabel] = useState('');
  const [operationPayload, setOperationPayload] = useState('{}');
  const [isPayloadValid, setIsPayloadValid] = useState(true);
  
  // Batch settings
  const [batchConfig, setBatchConfig] = useState<Partial<BatchConfig>>({
    name: 'Batch Operation',
    stopOnError: false,
    parallel: false,
    delayBetween: 100,
    enableUndo: true,
  });
  
  // Dialog states
  const [isClearDialogOpen, setIsClearDialogOpen] = useState(false);
  const [isExecuteDialogOpen, setIsExecuteDialogOpen] = useState(false);
  const [selectedHistoryItem, setSelectedHistoryItem] = useState<BatchResult | null>(null);

  // Load initial data
  useEffect(() => {
    refreshQueue();
    setPresets(batchOperations.getPresets());
    setHistory(batchOperations.getHistory());
  }, []);

  // Refresh queue periodically
  useEffect(() => {
    const interval = setInterval(() => {
      if (!isExecuting) {
        refreshQueue();
      }
    }, 1000);
    return () => clearInterval(interval);
  }, [isExecuting]);

  const refreshQueue = useCallback(() => {
    setQueue(batchOperations.getQueue());
  }, []);

  const handleAddOperation = () => {
    if (!operationLabel.trim()) {
      notificationManager.warning('Please enter an operation label');
      return;
    }

    let payload: any = {};
    try {
      payload = JSON.parse(operationPayload);
    } catch (e) {
      notificationManager.error('Invalid JSON payload');
      return;
    }

    batchOperations.enqueue({
      type: selectedActionType,
      label: operationLabel,
      payload,
    });

    setOperationLabel('');
    setOperationPayload('{}');
    refreshQueue();
    notificationManager.success('Operation added to queue');
  };

  const handleRemoveOperation = (id: string) => {
    batchOperations.dequeue(id);
    refreshQueue();
  };

  const handleClearQueue = () => {
    batchOperations.clearQueue();
    refreshQueue();
    setIsClearDialogOpen(false);
    notificationManager.info('Queue cleared');
  };

  const handleExecuteBatch = async () => {
    if (queue.length === 0) {
      notificationManager.warning('Queue is empty');
      return;
    }

    setIsExecuteDialogOpen(false);
    setIsExecuting(true);
    setResults(null);
    setProgress(null);

    try {
      const result = await batchOperations.executeQueue({
        ...batchConfig,
        onProgress: (p) => setProgress(p),
      });

      setResults(result);
      setHistory(batchOperations.getHistory());
      refreshQueue();

      if (result.success) {
        notificationManager.success(`Batch completed: ${result.operations.length} operations`);
      } else {
        notificationManager.warning(`Batch completed with ${result.operations.filter(o => !o.success).length} errors`);
      }

      onBatchComplete?.(result);
    } catch (error) {
      notificationManager.error('Batch execution failed: ' + error);
    } finally {
      setIsExecuting(false);
      setProgress(null);
    }
  };

  const handleExecutePreset = async (preset: BatchPreset) => {
    setIsExecuting(true);
    setResults(null);

    try {
      const result = await batchOperations.executePreset(preset.id, {
        onProgress: (p) => setProgress(p),
      });

      setResults(result);
      setHistory(batchOperations.getHistory());

      notificationManager.success(`Preset "${preset.name}" executed successfully`);
      onBatchComplete?.(result);
    } catch (error) {
      notificationManager.error('Preset execution failed: ' + error);
    } finally {
      setIsExecuting(false);
      setProgress(null);
    }
  };

  const handleUndo = async (batchId: string) => {
    const success = await batchOperations.undoBatch(batchId);
    if (success) {
      setHistory(batchOperations.getHistory());
    }
  };

  const handleRedo = async (batchId: string) => {
    const success = await batchOperations.redoBatch(batchId);
    if (success) {
      setHistory(batchOperations.getHistory());
    }
  };

  const validatePayload = (value: string) => {
    try {
      JSON.parse(value);
      setIsPayloadValid(true);
    } catch {
      setIsPayloadValid(false);
    }
    setOperationPayload(value);
  };

  // Queue columns
  const queueColumns: IColumn[] = useMemo(() => [
    {
      key: 'order',
      name: '#',
      minWidth: 30,
      maxWidth: 40,
      onRender: (_item, index) => <Text variant="small">{(index || 0) + 1}</Text>,
    },
    {
      key: 'type',
      name: 'Type',
      minWidth: 120,
      maxWidth: 150,
      onRender: (item: BatchOperation) => (
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
          {actionTypeOptions.find(o => o.key === item.type)?.text || item.type}
        </Text>
      ),
    },
    {
      key: 'label',
      name: 'Label',
      minWidth: 150,
      isResizable: true,
      onRender: (item: BatchOperation) => <Text variant="small">{item.label}</Text>,
    },
    {
      key: 'actions',
      name: '',
      minWidth: 50,
      maxWidth: 50,
      onRender: (item: BatchOperation) => (
        <TooltipHost content="Remove from queue">
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            onClick={() => handleRemoveOperation(item.id)}
            styles={{ root: { height: 24, width: 24 } }}
          />
        </TooltipHost>
      ),
    },
  ], []);

  // History columns
  const historyColumns: IColumn[] = useMemo(() => [
    {
      key: 'name',
      name: 'Batch',
      minWidth: 150,
      isResizable: true,
      onRender: (item: BatchResult) => (
        <Stack>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            {item.batchId.split('-').slice(0, 2).join('-')}
          </Text>
          <Text variant="tiny" styles={{ root: { color: '#666' } }}>
            {item.completedAt.toLocaleString()}
          </Text>
        </Stack>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 80,
      onRender: (item: BatchResult) => (
        <Text
          variant="small"
          styles={{
            root: {
              color: item.success ? '#107c10' : '#d13438',
              fontWeight: 600,
            },
          }}
        >
          {item.success ? 'Success' : 'Failed'}
        </Text>
      ),
    },
    {
      key: 'operations',
      name: 'Ops',
      minWidth: 50,
      onRender: (item: BatchResult) => (
        <Text variant="small">{item.operations.length}</Text>
      ),
    },
    {
      key: 'duration',
      name: 'Time',
      minWidth: 60,
      onRender: (item: BatchResult) => (
        <Text variant="small">{(item.duration / 1000).toFixed(1)}s</Text>
      ),
    },
    {
      key: 'actions',
      name: '',
      minWidth: 80,
      onRender: (item: BatchResult) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          {item.undoAvailable && batchOperations.canUndo(item.batchId) && (
            <TooltipHost content="Undo">
              <IconButton
                iconProps={{ iconName: 'Undo' }}
                onClick={() => handleUndo(item.batchId)}
                styles={{ root: { height: 24, width: 24 } }}
              />
            </TooltipHost>
          )}
          {batchOperations.canRedo(item.batchId) && (
            <TooltipHost content="Redo">
              <IconButton
                iconProps={{ iconName: 'Redo' }}
                onClick={() => handleRedo(item.batchId)}
                styles={{ root: { height: 24, width: 24 } }}
              />
            </TooltipHost>
          )}
          <TooltipHost content="View details">
            <IconButton
              iconProps={{ iconName: 'View' }}
              onClick={() => setSelectedHistoryItem(item)}
              styles={{ root: { height: 24, width: 24 } }}
            />
          </TooltipHost>
        </Stack>
      ),
    },
  ], []);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '16px', height: '100%' } }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <TaskListIcon />
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            Batch Operations
          </Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          {queue.length > 0 && !isExecuting && (
            <PrimaryButton
              iconProps={{ iconName: 'Play' }}
              onClick={() => setIsExecuteDialogOpen(true)}
              text={`Execute (${queue.length})`}
            />
          )}
          {isExecuting && progress && (
            <DefaultButton disabled>
              <Spinner size={SpinnerSize.small} />
            </DefaultButton>
          )}
        </Stack>
      </Stack>

      <Separator />

      {/* Navigation */}
      <Pivot
        selectedKey={activeView}
        onLinkClick={(item) => item && setActiveView(item.props.itemKey as PanelView)}
      >
        <PivotItem headerText="Queue" itemKey="queue" />
        <PivotItem headerText="Presets" itemKey="presets" />
        <PivotItem headerText="History" itemKey="history" />
        <PivotItem headerText="Settings" itemKey="settings" />
      </Pivot>

      {/* Queue View */}
      {activeView === 'queue' && (
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { flex: 1 } }}>
          {/* Add Operation Form */}
          <Stack tokens={{ childrenGap: 12 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Add Operation
            </Text>
            
            <Dropdown
              label="Action Type"
              selectedKey={selectedActionType}
              options={actionTypeOptions}
              onChange={(_, option) => option && setSelectedActionType(option.key as AIActionType)}
            />
            
            <TextField
              label="Label"
              placeholder="e.g., Format header row"
              value={operationLabel}
              onChange={(_, val) => setOperationLabel(val || '')}
            />
            
            <TextField
              label="Payload (JSON)"
              placeholder='{"address": "A1:D10", "bold": true}'
              value={operationPayload}
              onChange={(_, val) => validatePayload(val || '{}')}
              multiline
              rows={3}
              errorMessage={!isPayloadValid ? 'Invalid JSON' : undefined}
            />
            
            <DefaultButton
              iconProps={{ iconName: 'Add' }}
              onClick={handleAddOperation}
              disabled={!operationLabel.trim() || !isPayloadValid}
              text="Add to Queue"
            />
          </Stack>

          <Separator />

          {/* Queue List */}
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { flex: 1 } }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                Operation Queue ({queue.length})
              </Text>
              {queue.length > 0 && (
                <TooltipHost content="Clear queue">
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setIsClearDialogOpen(true)}
                  />
                </TooltipHost>
              )}
            </Stack>

            {queue.length === 0 ? (
              <MessageBar messageBarType={MessageBarType.info}>
                No operations in queue. Add operations above or use a preset.
              </MessageBar>
            ) : (
              <ScrollablePane
                scrollbarVisibility={ScrollbarVisibility.auto}
                styles={{ root: { height: 200 } }}
              >
                <DetailsList
                  items={queue}
                  columns={queueColumns}
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionMode={SelectionMode.none}
                  compact
                />
              </ScrollablePane>
            )}
          </Stack>

          {/* Execution Progress */}
          {isExecuting && progress && (
            <Stack tokens={{ childrenGap: 8 }}>
              <ProgressIndicator
                label={progress.currentOperation?.label || 'Executing...'}
                description={`${progress.completed} of ${progress.total} operations (${progress.percentComplete}%)`}
                percentComplete={progress.percentComplete / 100}
              />
            </Stack>
          )}

          {/* Results */}
          {results && (
            <MessageBar
              messageBarType={results.success ? MessageBarType.success : MessageBarType.warning}
              isMultiline
            >
              <Stack tokens={{ childrenGap: 4 }}>
                <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                  Batch {results.success ? 'Completed' : 'Completed with Errors'}
                </Text>
                <Text variant="small">
                  {results.operations.filter(o => o.success).length} succeeded,{' '}
                  {results.operations.filter(o => !o.success).length} failed
                  {' '}• {results.duration / 1000}s
                </Text>
              </Stack>
            </MessageBar>
          )}
        </Stack>
      )}

      {/* Presets View */}
      {activeView === 'presets' && (
        <ScrollablePane
          scrollbarVisibility={ScrollbarVisibility.auto}
          styles={{ root: { flex: 1 } }}
        >
          <Stack tokens={{ childrenGap: 12 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Batch Presets
            </Text>
            
            {presets.map((preset) => (
              <Stack
                key={preset.id}
                tokens={{ childrenGap: 8 }}
                styles={{
                  root: {
                    padding: 12,
                    border: '1px solid #e0e0e0',
                    borderRadius: 4,
                    borderLeft: `4px solid ${categoryColors[preset.category] || '#ccc'}`,
                  },
                }}
              >
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                  <Stack>
                    <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                      {preset.name}
                    </Text>
                    <Text variant="small" styles={{ root: { color: '#666' } }}>
                      {preset.description}
                    </Text>
                  </Stack>
                  <PrimaryButton
                    iconProps={{ iconName: 'Play' }}
                    onClick={() => handleExecutePreset(preset)}
                    disabled={isExecuting}
                    text="Run"
                  />
                </Stack>
                
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <Text variant="tiny" styles={{ root: { color: '#666' } }}>
                    {preset.operations.length} operations
                  </Text>
                  <Text variant="tiny" styles={{ root: { color: '#666' } }}>
                    • {preset.category}
                  </Text>
                  {preset.tags.map((tag) => (
                    <Text
                      key={tag}
                      variant="tiny"
                      styles={{
                        root: {
                          backgroundColor: '#f0f0f0',
                          padding: '2px 6px',
                          borderRadius: 3,
                        },
                      }}
                    >
                      {tag}
                    </Text>
                  ))}
                </Stack>
              </Stack>
            ))}
          </Stack>
        </ScrollablePane>
      )}

      {/* History View */}
      {activeView === 'history' && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { flex: 1 } }}>
          <Stack horizontal horizontalAlign="space-between">
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Execution History ({history.length})
            </Text>
            <DefaultButton
              iconProps={{ iconName: 'Clear' }}
              onClick={() => {
                batchOperations.clearHistory();
                setHistory([]);
              }}
              text="Clear"
            />
          </Stack>
          
          {history.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No batch operations executed yet.
            </MessageBar>
          ) : (
            <ScrollablePane
              scrollbarVisibility={ScrollbarVisibility.auto}
              styles={{ root: { flex: 1 } }}
            >
              <DetailsList
                items={history}
                columns={historyColumns}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                compact
              />
            </ScrollablePane>
          )}
        </Stack>
      )}

      {/* Settings View */}
      {activeView === 'settings' && (
        <Stack tokens={{ childrenGap: 16 }}>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
            Batch Execution Settings
          </Text>
          
          <TextField
            label="Batch Name"
            value={batchConfig.name}
            onChange={(_, val) => setBatchConfig({ ...batchConfig, name: val })}
          />
          
          <Toggle
            label="Stop on error"
            checked={batchConfig.stopOnError}
            onChange={(_, checked) => setBatchConfig({ ...batchConfig, stopOnError: checked })}
          />
          
          <Toggle
            label="Execute in parallel"
            checked={batchConfig.parallel}
            onChange={(_, checked) => setBatchConfig({ ...batchConfig, parallel: checked })}
          />
          
          <Toggle
            label="Enable undo"
            checked={batchConfig.enableUndo}
            onChange={(_, checked) => setBatchConfig({ ...batchConfig, enableUndo: checked })}
          />
          
          <TextField
            label="Delay between operations (ms)"
            type="number"
            value={String(batchConfig.delayBetween)}
            onChange={(_, val) => setBatchConfig({ ...batchConfig, delayBetween: parseInt(val || '0') })}
          />
        </Stack>
      )}

      {/* Clear Queue Dialog */}
      <Dialog
        hidden={!isClearDialogOpen}
        onDismiss={() => setIsClearDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Clear Queue',
          subText: `Are you sure you want to remove all ${queue.length} operations from the queue?`,
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setIsClearDialogOpen(false)} text="Cancel" />
          <PrimaryButton onClick={handleClearQueue} text="Clear" />
        </DialogFooter>
      </Dialog>

      {/* Execute Dialog */}
      <Dialog
        hidden={!isExecuteDialogOpen}
        onDismiss={() => setIsExecuteDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Execute Batch',
          subText: `Execute ${queue.length} operations? This may take a moment.`,
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setIsExecuteDialogOpen(false)} text="Cancel" />
          <PrimaryButton onClick={handleExecuteBatch} text="Execute" />
        </DialogFooter>
      </Dialog>

      {/* History Item Details Dialog */}
      {selectedHistoryItem && (
        <Dialog
          hidden={!selectedHistoryItem}
          onDismiss={() => setSelectedHistoryItem(null)}
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: 'Batch Details',
          }}
        >
          <Stack tokens={{ childrenGap: 12 }}>
            <Text variant="medium">
              <strong>Batch ID:</strong> {selectedHistoryItem.batchId}
            </Text>
            <Text variant="medium">
              <strong>Completed:</strong> {selectedHistoryItem.completedAt.toLocaleString()}
            </Text>
            <Text variant="medium">
              <strong>Duration:</strong> {(selectedHistoryItem.duration / 1000).toFixed(2)} seconds
            </Text>
            <Text variant="medium">
              <strong>Status:</strong>{' '}
              <span style={{ color: selectedHistoryItem.success ? '#107c10' : '#d13438' }}>
                {selectedHistoryItem.success ? 'Success' : 'Failed'}
              </span>
            </Text>
            
            <Separator />
            
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Operations ({selectedHistoryItem.operations.length})
            </Text>
            
            <ScrollablePane styles={{ root: { maxHeight: 300 } }}>
              <Stack tokens={{ childrenGap: 8 }}>
                {selectedHistoryItem.operations.map((op, idx) => (
                  <Stack
                    key={op.operationId}
                    horizontal
                    horizontalAlign="space-between"
                    styles={{
                      root: {
                        padding: 8,
                        backgroundColor: op.success ? '#f6f6f6' : '#fde7e9',
                        borderRadius: 4,
                      },
                    }}
                  >
                    <Stack>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                        {idx + 1}. {op.operationId.split('-')[0]}
                      </Text>
                      {op.error && (
                        <Text variant="tiny" styles={{ root: { color: '#d13438' } }}>
                          Error: {op.error}
                        </Text>
                      )}
                    </Stack>
                    <Text variant="small" styles={{ root: { color: '#666' } }}>
                      {op.duration}ms
                    </Text>
                  </Stack>
                ))}
              </Stack>
            </ScrollablePane>
          </Stack>
          
          <DialogFooter>
            <DefaultButton onClick={() => setSelectedHistoryItem(null)} text="Close" />
          </DialogFooter>
        </Dialog>
      )}
    </Stack>
  );
};

export default BatchOperationsPanel;
