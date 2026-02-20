/**
 * VisualHighlightPanel Component
 * 
 * UI panel for managing visual cell highlights in Excel.
 * Provides controls for viewing, navigating, and clearing highlights.
 * Integrates with visualHighlighter service.
 * 
 * @module VisualHighlightPanel
 */

import React, { useState, useEffect, useCallback } from 'react';
import {
  Stack,
  Text,
  IconButton,
  PrimaryButton,
  DefaultButton,
  Separator,
  List,
  TooltipHost,
  Toggle,
  Slider,
  Dropdown,
  IDropdownOption,
  Label,
  MessageBar,
  MessageBarType,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
  ContextualMenu,
  IContextualMenuProps,
  Selection,
  SelectionMode,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  ScrollablePane,
  ScrollbarVisibility,
} from '@fluentui/react';
import {
  NavigateForwardMirroredIcon,
  ClearFilterIcon,
  AutoEnhanceOffIcon,
  ColorIcon,
  DeleteIcon,
  MoreVerticalIcon,
  NavigateBackIcon,
  MapPinIcon,
  FlashOnIcon,
} from '@fluentui/react-icons-mdl2';
import { visualHighlighter, HighlightInfo, HighlightCategory } from '../services/visualHighlighter';
import { cellReferenceParser } from '../utils/cellReferenceParser';
import { notificationManager } from '../utils/notificationManager';

interface HighlightListItem extends HighlightInfo {
  id: string;
  displayName: string;
  cellCount: number;
}

interface VisualHighlightPanelProps {
  /** AI response text to auto-extract highlights from */
  aiResponseText?: string;
  /** Enable auto-highlight from AI responses */
  autoHighlightEnabled?: boolean;
  /** Callback when highlights change */
  onHighlightsChange?: (highlights: HighlightListItem[]) => void;
}

const categoryOptions: IDropdownOption[] = [
  { key: 'ai-reference', text: 'AI Reference', data: { color: '#107c10' } },
  { key: 'error', text: 'Error', data: { color: '#d13438' } },
  { key: 'warning', text: 'Warning', data: { color: '#ffc107' } },
  { key: 'success', text: 'Success', data: { color: '#28a745' } },
  { key: 'info', text: 'Info', data: { color: '#0078d4' } },
  { key: 'anomaly', text: 'Anomaly', data: { color: '#ff5722' } },
  { key: 'trend', text: 'Trend', data: { color: '#673ab7' } },
  { key: 'selection', text: 'Selection', data: { color: '#ff9800' } },
];

const pulseSpeedOptions: IDropdownOption[] = [
  { key: 'slow', text: 'Slow (2s)' },
  { key: 'normal', text: 'Normal (1s)' },
  { key: 'fast', text: 'Fast (0.5s)' },
];

export const VisualHighlightPanel: React.FC<VisualHighlightPanelProps> = ({
  aiResponseText,
  autoHighlightEnabled = true,
  onHighlightsChange,
}) => {
  const [highlights, setHighlights] = useState<HighlightListItem[]>([]);
  const [selectedCategory, setSelectedCategory] = useState<HighlightCategory>('ai-reference');
  const [pulseEnabled, setPulseEnabled] = useState(true);
  const [pulseSpeed, setPulseSpeed] = useState<'slow' | 'normal' | 'fast'>('normal');
  const [autoClearDelay, setAutoClearDelay] = useState(30);
  const [isClearDialogOpen, setIsClearDialogOpen] = useState(false);
  const [customRangeInput, setCustomRangeInput] = useState('');
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const [menuTarget, setMenuTarget] = useState<HTMLElement | undefined>();
  const [selectedHighlightId, setSelectedHighlightId] = useState<string | null>(null);

  // Load highlights on mount and when they change
  useEffect(() => {
    loadHighlights();

    // Set up interval to refresh highlights
    const interval = setInterval(loadHighlights, 1000);

    return () => clearInterval(interval);
  }, []);

  // Auto-extract highlights from AI response
  useEffect(() => {
    if (autoHighlightEnabled && aiResponseText) {
      extractHighlightsFromAIResponse(aiResponseText);
    }
  }, [aiResponseText, autoHighlightEnabled]);

  // Notify parent of changes
  useEffect(() => {
    onHighlightsChange?.(highlights);
  }, [highlights, onHighlightsChange]);

  const loadHighlights = useCallback(() => {
    const currentHighlights = visualHighlighter.getAllHighlights();
    const items: HighlightListItem[] = currentHighlights.map(h => ({
      ...h,
      id: `${h.category}_${h.rangeAddress}_${h.timestamp.getTime()}`,
      displayName: h.rangeAddress,
      cellCount: calculateCellCount(h.rangeAddress),
    }));
    setHighlights(items);
  }, []);

  const calculateCellCount = (rangeAddress: string): number => {
    try {
      const parsed = cellReferenceParser.parseA1Range(rangeAddress);
      if (parsed) {
        return cellReferenceParser.getRangeSize(parsed);
      }
    } catch {
      // Ignore parsing errors
    }
    return 1;
  };

  const extractHighlightsFromAIResponse = async (text: string) => {
    const references = cellReferenceParser.extractReferencesFromAIResponse(text);
    
    for (const ref of references) {
      const rangeStr = cellReferenceParser.rangeToA1(ref.range);
      await visualHighlighter.highlightRange(rangeStr, 'ai-reference', {
        pulse: pulseEnabled,
        pulseSpeed,
        autoClear: autoClearDelay > 0,
        autoClearDelay: autoClearDelay * 1000,
      });
    }

    if (references.length > 0) {
      notificationManager.success(`Highlighted ${references.length} cell reference(s) from AI response`);
      loadHighlights();
    }
  };

  const handleNavigateToHighlight = async (highlightId: string) => {
    const highlight = visualHighlighter.getHighlight(highlightId);
    if (highlight) {
      await visualHighlighter.navigateToHighlight(highlightId);
    }
  };

  const handleClearHighlight = async (highlightId: string) => {
    await visualHighlighter.clearHighlight(highlightId);
    loadHighlights();
  };

  const handleClearAllHighlights = async () => {
    await visualHighlighter.clearAllHighlights();
    loadHighlights();
    setIsClearDialogOpen(false);
    notificationManager.success('All highlights cleared');
  };

  const handleClearByCategory = async (category: HighlightCategory) => {
    await visualHighlighter.clearByCategory(category);
    loadHighlights();
  };

  const handleAddCustomHighlight = async () => {
    if (!customRangeInput.trim()) {
      notificationManager.warning('Please enter a cell range');
      return;
    }

    if (!cellReferenceParser.isValidCellReference(customRangeInput)) {
      notificationManager.error('Invalid cell reference format');
      return;
    }

    await visualHighlighter.highlightRange(customRangeInput, selectedCategory, {
      pulse: pulseEnabled,
      pulseSpeed,
      autoClear: autoClearDelay > 0,
      autoClearDelay: autoClearDelay * 1000,
    });

    setCustomRangeInput('');
    loadHighlights();
    notificationManager.success(`Highlighted ${customRangeInput}`);
  };

  const handleExtractFromSelection = async () => {
    try {
      await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load('address');
        await context.sync();

        await visualHighlighter.highlightRange(selection.address, selectedCategory, {
          pulse: pulseEnabled,
          pulseSpeed,
          autoClear: autoClearDelay > 0,
          autoClearDelay: autoClearDelay * 1000,
        });

        loadHighlights();
        notificationManager.success(`Highlighted ${selection.address}`);
      });
    } catch (error) {
      notificationManager.error('Failed to highlight selection: ' + error);
    }
  };

  const getCategoryColor = (category: HighlightCategory): string => {
    const option = categoryOptions.find(opt => opt.key === category);
    return option?.data?.color || '#cccccc';
  };

  const columns: IColumn[] = [
    {
      key: 'color',
      name: '',
      minWidth: 30,
      maxWidth: 30,
      onRender: (item: HighlightListItem) => (
        <div
          style={{
            width: 16,
            height: 16,
            borderRadius: 3,
            backgroundColor: item.color,
            border: '1px solid #ccc',
          }}
        />
      ),
    },
    {
      key: 'range',
      name: 'Range',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: HighlightListItem) => (
        <TooltipHost content={item.rangeAddress}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            {item.displayName}
          </Text>
        </TooltipHost>
      ),
    },
    {
      key: 'category',
      name: 'Category',
      minWidth: 80,
      maxWidth: 100,
      onRender: (item: HighlightListItem) => (
        <Text variant="small" styles={{ root: { color: getCategoryColor(item.category) } }}>
          {item.category}
        </Text>
      ),
    },
    {
      key: 'cells',
      name: 'Cells',
      minWidth: 50,
      maxWidth: 60,
      onRender: (item: HighlightListItem) => (
        <Text variant="small">{item.cellCount}</Text>
      ),
    },
    {
      key: 'actions',
      name: '',
      minWidth: 80,
      maxWidth: 80,
      onRender: (item: HighlightListItem) => (
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <TooltipHost content="Navigate to range">
            <IconButton
              iconProps={{ iconName: 'NavigateForward' }}
              onClick={() => handleNavigateToHighlight(item.id)}
              styles={{ root: { height: 24, width: 24 } }}
            />
          </TooltipHost>
          <TooltipHost content="Remove highlight">
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              onClick={() => handleClearHighlight(item.id)}
              styles={{ root: { height: 24, width: 24 } }}
            />
          </TooltipHost>
          <TooltipHost content="More actions">
            <IconButton
              iconProps={{ iconName: 'MoreVertical' }}
              onClick={(e) => {
                setSelectedHighlightId(item.id);
                setMenuTarget(e.currentTarget as HTMLElement);
                setIsMenuOpen(true);
              }}
              styles={{ root: { height: 24, width: 24 } }}
            />
          </TooltipHost>
        </Stack>
      ),
    },
  ];

  const menuProps: IContextualMenuProps = {
    target: menuTarget,
    directionalHint: 4, // bottomLeftEdge
    onDismiss: () => setIsMenuOpen(false),
    items: [
      {
        key: 'navigate',
        text: 'Navigate to Range',
        iconProps: { iconName: 'NavigateForward' },
        onClick: () => selectedHighlightId && handleNavigateToHighlight(selectedHighlightId),
      },
      {
        key: 'clear',
        text: 'Remove Highlight',
        iconProps: { iconName: 'Delete' },
        onClick: () => selectedHighlightId && handleClearHighlight(selectedHighlightId),
      },
      {
        key: 'divider1',
        itemType: 1, // Divider
      },
      {
        key: 'clearCategory',
        text: 'Clear by Category',
        iconProps: { iconName: 'ClearFilter' },
        subMenuProps: {
          items: categoryOptions.map(opt => ({
            key: opt.key as string,
            text: opt.text,
            onClick: () => handleClearByCategory(opt.key as HighlightCategory),
          })),
        },
      },
    ],
  };

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '16px' } }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <ColorIcon />
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            Visual Highlights
          </Text>
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <TooltipHost content="Clear all highlights">
            <IconButton
              iconProps={{ iconName: 'ClearFilter' }}
              onClick={() => setIsClearDialogOpen(true)}
              disabled={highlights.length === 0}
            />
          </TooltipHost>
          <TooltipHost content="Refresh">
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              onClick={loadHighlights}
            />
          </TooltipHost>
        </Stack>
      </Stack>

      <Separator />

      {/* Auto-highlight toggle */}
      <Toggle
        label="Auto-highlight from AI responses"
        checked={autoHighlightEnabled}
        onChange={(_, checked) => {}}
        onText="Enabled"
        offText="Disabled"
      />

      {/* New Highlight Section */}
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          Add Highlight
        </Text>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <TextField
            placeholder="Enter range (e.g., A1:D10)"
            value={customRangeInput}
            onChange={(_, val) => setCustomRangeInput(val || '')}
            styles={{ root: { flex: 1 } }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                handleAddCustomHighlight();
              }
            }}
          />
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            onClick={handleAddCustomHighlight}
            disabled={!customRangeInput.trim()}
          >
            Add
          </PrimaryButton>
        </Stack>

        <DefaultButton
          iconProps={{ iconName: 'MapPin' }}
          onClick={handleExtractFromSelection}
          text="Highlight Current Selection"
        />

        <Dropdown
          label="Category"
          selectedKey={selectedCategory}
          options={categoryOptions}
          onChange={(_, option) => option && setSelectedCategory(option.key as HighlightCategory)}
        />

        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <Toggle
            label="Pulse animation"
            checked={pulseEnabled}
            onChange={(_, checked) => setPulseEnabled(!!checked)}
            inlineLabel
          />
          {pulseEnabled && (
            <Dropdown
              selectedKey={pulseSpeed}
              options={pulseSpeedOptions}
              onChange={(_, option) => option && setPulseSpeed(option.key as 'slow' | 'normal' | 'fast')}
              styles={{ root: { width: 120 } }}
            />
          )}
        </Stack>

        <Slider
          label={`Auto-clear after ${autoClearDelay} seconds (0 = never)`}
          min={0}
          max={300}
          step={10}
          value={autoClearDelay}
          onChange={setAutoClearDelay}
          showValue={false}
          valueFormat={(val) => val === 0 ? 'Never' : `${val}s`}
        />
      </Stack>

      <Separator />

      {/* Highlights List */}
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { flex: 1, minHeight: 200 } }}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
            Active Highlights ({highlights.length})
          </Text>
          {highlights.length > 0 && (
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              {highlights.reduce((sum, h) => sum + h.cellCount, 0)} cells highlighted
            </Text>
          )}
        </Stack>

        {highlights.length === 0 ? (
          <MessageBar messageBarType={MessageBarType.info}>
            No active highlights. Add highlights manually or enable auto-highlight to see cell references from AI responses.
          </MessageBar>
        ) : (
          <ScrollablePane
            scrollbarVisibility={ScrollbarVisibility.auto}
            styles={{ root: { height: 250 } }}
          >
            <DetailsList
              items={highlights}
              columns={columns}
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              compact
              isHeaderVisible
            />
          </ScrollablePane>
        )}
      </Stack>

      {/* Quick Actions */}
      {highlights.length > 0 && (
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Quick Actions
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Clear All"
              iconProps={{ iconName: 'ClearFilter' }}
              onClick={() => setIsClearDialogOpen(true)}
            />
            <DefaultButton
              text="Clear AI References"
              iconProps={{ iconName: 'AutoEnhanceOff' }}
              onClick={() => handleClearByCategory('ai-reference')}
            />
            <DefaultButton
              text="Clear Errors"
              iconProps={{ iconName: 'ErrorBadge' }}
              onClick={() => handleClearByCategory('error')}
            />
          </Stack>
        </Stack>
      )}

      {/* Clear All Dialog */}
      <Dialog
        hidden={!isClearDialogOpen}
        onDismiss={() => setIsClearDialogOpen(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Clear All Highlights',
          subText: `Are you sure you want to clear all ${highlights.length} active highlight(s)? This action cannot be undone.`,
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setIsClearDialogOpen(false)} text="Cancel" />
          <PrimaryButton onClick={handleClearAllHighlights} text="Clear All" />
        </DialogFooter>
      </Dialog>

      {/* Context Menu */}
      {isMenuOpen && <ContextualMenu {...menuProps} />}
    </Stack>
  );
};

export default VisualHighlightPanel;
