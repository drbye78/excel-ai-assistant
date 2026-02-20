/**
 * DAX Explainer Component
 *
 * Interactive component for analyzing and explaining DAX formulas
 * in Power Pivot and Excel Data Models. Similar to FormulaExplainer
 * but specialized for DAX context.
 *
 * Features:
 * - Real-time DAX formula analysis
 * - Human-readable explanations
 * - Context transition visualization
 * - Performance optimization hints
 * - Function reference lookup
 * - Dependency tracking
 *
 * @module components/DAXExplainer
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
  Dropdown,
  IDropdownOption,
  TextField,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Label,
  ProgressIndicator,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  ScrollablePane,
  ScrollbarVisibility,
  Shimmer,
  ShimmerElementsGroup,
  ShimmerElementType,
} from '@fluentui/react';
import {
  SearchIcon,
  LightbulbIcon,
  WarningIcon,
  ErrorBadgeIcon,
  InfoIcon,
  ChartIcon,
  FilterSettingsIcon,
  TimerIcon,
  FunctionIcon,
} from '@fluentui/react-icons-mdl2';
import {
  daxParser,
  DAXExplanation,
  DAXFunction,
  DAXComplexity,
  DAXFunctionCategory,
  ParsedDAXFormula,
} from '../services/daxParser';
import { notificationManager } from '../utils/notificationManager';
import { powerPivotService, PowerPivotMeasure } from '../services/powerPivotService';
import { logger } from '../utils/logger';

interface DAXExplainerProps {
  /** Initial formula to analyze */
  initialFormula?: string;
  /** Callback when formula is analyzed */
  onAnalyze?: (explanation: DAXExplanation) => void;
  /** Enable Power Pivot integration */
  enablePowerPivot?: boolean;
}

const complexityColors: Record<DAXComplexity, string> = {
  simple: '#107c10',
  moderate: '#ffc107',
  complex: '#ff9800',
  very_complex: '#d13438',
};

const complexityLabels: Record<DAXComplexity, string> = {
  simple: 'Simple',
  moderate: 'Moderate',
  complex: 'Complex',
  very_complex: 'Very Complex',
};

const categoryOptions: IDropdownOption[] = [
  { key: 'all', text: 'All Categories' },
  { key: 'aggregation', text: 'Aggregation' },
  { key: 'filter', text: 'Filter' },
  { key: 'time_intelligence', text: 'Time Intelligence' },
  { key: 'logical', text: 'Logical' },
  { key: 'mathematical', text: 'Mathematical' },
  { key: 'text', text: 'Text' },
  { key: 'date_time', text: 'Date & Time' },
  { key: 'relationship', text: 'Relationship' },
  { key: 'information', text: 'Information' },
  { key: 'table_manipulation', text: 'Table Manipulation' },
];

export const DAXExplainer: React.FC<DAXExplainerProps> = ({
  initialFormula = '',
  onAnalyze,
  enablePowerPivot = true,
}) => {
  const [formula, setFormula] = useState(initialFormula);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [explanation, setExplanation] = useState<DAXExplanation | null>(null);
  const [parsed, setParsed] = useState<ParsedDAXFormula | null>(null);
  const [errors, setErrors] = useState<string[]>([]);
  const [activeTab, setActiveTab] = useState('explanation');
  
  // Function reference state
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [selectedFunction, setSelectedFunction] = useState<DAXFunction | null>(null);
  const [availableMeasures, setAvailableMeasures] = useState<PowerPivotMeasure[]>([]);
  const [isLoadingMeasures, setIsLoadingMeasures] = useState(false);
  const [showMeasureSelector, setShowMeasureSelector] = useState(false);

  // Analyze formula on mount if provided
  useEffect(() => {
    if (initialFormula) {
      handleAnalyze();
    }
  }, []);

  const handleAnalyze = useCallback(() => {
    if (!formula.trim()) {
      notificationManager.warning('Please enter a DAX formula');
      return;
    }

    setIsAnalyzing(true);
    setErrors([]);

    try {
      // Validate syntax first
      const syntaxErrors = daxParser.validate(formula);
      const errorMessages = syntaxErrors
        .filter(e => e.severity === 'error')
        .map(e => e.message);
      
      if (errorMessages.length > 0) {
        setErrors(errorMessages);
      }

      // Parse and explain
      const parsedResult = daxParser.parse(formula);
      const explanationResult = daxParser.explain(formula);

      setParsed(parsedResult);
      setExplanation(explanationResult);
      onAnalyze?.(explanationResult);

      if (errorMessages.length === 0) {
        notificationManager.success('DAX formula analyzed successfully');
      }
    } catch (error) {
      notificationManager.error('Failed to analyze formula: ' + error);
      setErrors([error instanceof Error ? error.message : String(error)]);
    } finally {
      setIsAnalyzing(false);
    }
  }, [formula, onAnalyze]);

  // Filter functions based on search and category
  const filteredFunctions = useMemo(() => {
    let functions = daxParser.searchFunctions(searchQuery);
    if (selectedCategory !== 'all') {
      functions = functions.filter(f => f.category === selectedCategory);
    }
    return functions;
  }, [searchQuery, selectedCategory]);

  // Function reference columns
  const functionColumns: IColumn[] = useMemo(() => [
    {
      key: 'name',
      name: 'Function',
      minWidth: 120,
      maxWidth: 150,
      onRender: (item: DAXFunction) => (
        <Text variant="small" styles={{ root: { fontWeight: 600, fontFamily: 'monospace' } }}>
          {item.name}
        </Text>
      ),
    },
    {
      key: 'category',
      name: 'Category',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: DAXFunction) => (
        <Text variant="small" styles={{ root: { color: '#666' } }}>
          {item.category.replace('_', ' ')}
        </Text>
      ),
    },
    {
      key: 'description',
      name: 'Description',
      minWidth: 200,
      isResizable: true,
      onRender: (item: DAXFunction) => (
        <Text variant="small">{item.description}</Text>
      ),
    },
  ], []);

  // Dependency columns
  const dependencyColumns: IColumn[] = useMemo(() => [
    {
      key: 'type',
      name: 'Type',
      minWidth: 80,
      onRender: (item) => (
        <Text
          variant="small"
          styles={{
            root: {
              fontWeight: 600,
              color: item.type === 'measure' ? '#0078d4' : item.type === 'column' ? '#107c10' : '#666',
            },
          }}
        >
          {item.type}
        </Text>
      ),
    },
    {
      key: 'name',
      name: 'Name',
      minWidth: 150,
      isResizable: true,
      onRender: (item) => (
        <Text variant="small" styles={{ root: { fontFamily: 'monospace' } }}>
          {item.table ? `${item.table}[${item.name}]` : item.name}
        </Text>
      ),
    },
  ], []);

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { padding: '16px', height: '100%' } }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <FunctionIcon />
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            DAX Formula Explainer
          </Text>
        </Stack>
        {enablePowerPivot && (
          <DefaultButton
            iconProps={{ iconName: 'Database' }}
            text="Get from Power Pivot"
            onClick={async () => {
              setIsLoadingMeasures(true);
              try {
                const measures = await powerPivotService.getMeasures();
                if (measures.length === 0) {
                  notificationManager.warning('No measures found in the data model');
                } else {
                  setAvailableMeasures(measures);
                  setShowMeasureSelector(true);
                  notificationManager.success(`Found ${measures.length} measures`);
                }
              } catch (error) {
                logger.error('Failed to fetch measures from Power Pivot', { error });
                notificationManager.error('Failed to fetch measures: ' + (error instanceof Error ? error.message : String(error)));
              } finally {
                setIsLoadingMeasures(false);
              }
            }}
            disabled={isLoadingMeasures}
          />
        )}
      </Stack>

      <Separator />

      {/* Formula Input */}
      <Stack tokens={{ childrenGap: 8 }}>
        <Label>DAX Formula</Label>
        <TextField
          multiline
          rows={4}
          placeholder="Enter DAX formula (e.g., CALCULATE(SUM(Sales[Amount]), ALL(Dates[Year])))"
          value={formula}
          onChange={(_, val) => setFormula(val || '')}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && e.ctrlKey) {
              handleAnalyze();
            }
          }}
          styles={{
            fieldGroup: { fontFamily: 'Consolas, monospace' },
          }}
        />
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="small" styles={{ root: { color: '#666' } }}>
            Press Ctrl+Enter to analyze
          </Text>
          <PrimaryButton
            iconProps={{ iconName: 'Search' }}
            onClick={handleAnalyze}
            disabled={isAnalyzing || !formula.trim()}
            text={isAnalyzing ? 'Analyzing...' : 'Analyze'}
          />
        </Stack>
      </Stack>

      {/* Errors */}
      {errors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          <Stack tokens={{ childrenGap: 4 }}>
            {errors.map((error, idx) => (
              <Text key={idx} variant="small">
                {error}
              </Text>
            ))}
          </Stack>
        </MessageBar>
      )}

      {/* Results */}
      {explanation && !isAnalyzing && (
        <>
          <Separator />
          
          {/* Complexity Badge */}
          <Stack
            horizontal
            tokens={{ childrenGap: 16 }}
            styles={{
              root: {
                padding: '12px 16px',
                backgroundColor: '#f3f2f1',
                borderRadius: 4,
              },
            }}
          >
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <ChartIcon />
              <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                Complexity:
              </Text>
              <Text
                variant="medium"
                styles={{
                  root: {
                    fontWeight: 600,
                    color: complexityColors[explanation.complexity],
                  },
                }}
              >
                {complexityLabels[explanation.complexity]}
              </Text>
            </Stack>
            
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <FilterSettingsIcon />
              <Text variant="small" styles={{ root: { color: '#666' } }}>
                {explanation.contextInfo.filterModifications} filter mods
              </Text>
            </Stack>
            
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <TimerIcon />
              <Text variant="small" styles={{ root: { color: '#666' } }}>
                {explanation.contextInfo.iteratorFunctions} iterators
              </Text>
            </Stack>
          </Stack>

          {/* Tabs */}
          <Pivot
            selectedKey={activeTab}
            onLinkClick={(item) => item && setActiveTab(item.props.itemKey as string)}
          >
            <PivotItem headerText="Explanation" itemKey="explanation">
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
                {/* Summary */}
                <MessageBar messageBarType={MessageBarType.info}>
                  {explanation.summary}
                </MessageBar>

                {/* Breakdown */}
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                  Formula Breakdown
                </Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  {explanation.breakdown.map((step, idx) => (
                    <Stack
                      key={idx}
                      tokens={{ childrenGap: 4 }}
                      styles={{
                        root: {
                          paddingLeft: step.indentation * 20,
                          borderLeft: '2px solid #e0e0e0',
                          padding: '8px 8px 8px 12px',
                        },
                      }}
                    >
                      <Text variant="small" styles={{ root: { fontWeight: 600, fontFamily: 'monospace' } }}>
                        {step.function}
                      </Text>
                      <Text variant="small">{step.description}</Text>
                      <Text variant="tiny" styles={{ root: { color: '#666', fontStyle: 'italic' } }}>
                        Context: {step.context}
                      </Text>
                    </Stack>
                  ))}
                </Stack>
              </Stack>
            </PivotItem>

            <PivotItem headerText="Optimizations" itemKey="optimizations">
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
                {explanation.optimizations.length === 0 ? (
                  <MessageBar messageBarType={MessageBarType.success}>
                    No optimization suggestions - your formula looks good!
                  </MessageBar>
                ) : (
                  <>
                    {explanation.optimizations.map((opt, idx) => (
                      <MessageBar
                        key={idx}
                        messageBarType={
                          opt.severity === 'high'
                            ? MessageBarType.severeWarning
                            : opt.severity === 'medium'
                            ? MessageBarType.warning
                            : MessageBarType.info
                        }
                        isMultiline
                      >
                        <Stack tokens={{ childrenGap: 4 }}>
                          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                            {opt.type === 'performance'
                              ? '⚡ Performance'
                              : opt.type === 'readability'
                              ? '📖 Readability'
                              : '✨ Best Practice'}
                            : {opt.message}
                          </Text>
                          <Text variant="small">{opt.suggestion}</Text>
                          {opt.example && (
                            <Text
                              variant="small"
                              styles={{
                                root: {
                                  fontFamily: 'monospace',
                                  backgroundColor: '#f0f0f0',
                                  padding: '4px 8px',
                                  borderRadius: 3,
                                },
                              }}
                            >
                              Example: {opt.example}
                            </Text>
                          )}
                        </Stack>
                      </MessageBar>
                    ))}
                    
                    {explanation.performanceHints.length > 0 && (
                      <>
                        <Separator />
                        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                          Performance Notes
                        </Text>
                        {explanation.performanceHints.map((hint, idx) => (
                          <MessageBar key={idx} messageBarType={MessageBarType.info}>
                            {hint}
                          </MessageBar>
                        ))}
                      </>
                    )}
                  </>
                )}
              </Stack>
            </PivotItem>

            <PivotItem headerText="Dependencies" itemKey="dependencies">
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
                {parsed && parsed.dependencies.length > 0 ? (
                  <DetailsList
                    items={parsed.dependencies}
                    columns={dependencyColumns}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    selectionMode={SelectionMode.none}
                    compact
                  />
                ) : (
                  <MessageBar messageBarType={MessageBarType.info}>
                    No external dependencies detected
                  </MessageBar>
                )}
              </Stack>
            </PivotItem>

            <PivotItem headerText="Function Reference" itemKey="reference">
              <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
                {/* Search */}
                <Stack horizontal tokens={{ childrenGap: 8 }}>
                  <TextField
                    placeholder="Search functions..."
                    value={searchQuery}
                    onChange={(_, val) => setSearchQuery(val || '')}
                    styles={{ root: { flex: 1 } }}
                  />
                  <Dropdown
                    selectedKey={selectedCategory}
                    options={categoryOptions}
                    onChange={(_, option) => option && setSelectedCategory(option.key as string)}
                    styles={{ root: { width: 150 } }}
                  />
                </Stack>

                {/* Function List */}
                <ScrollablePane styles={{ root: { height: 300 } }}>
                  <DetailsList
                    items={filteredFunctions}
                    columns={functionColumns}
                    layoutMode={DetailsListLayoutMode.fixedColumns}
                    selectionMode={SelectionMode.single}
                    onActiveItemChanged={(item) => setSelectedFunction(item)}
                    compact
                  />
                </ScrollablePane>

                {/* Selected Function Details */}
                {selectedFunction && (
                  <Stack
                    tokens={{ childrenGap: 8 }}
                    styles={{
                      root: {
                        padding: 12,
                        backgroundColor: '#f9f9f9',
                        borderRadius: 4,
                      },
                    }}
                  >
                    <Stack horizontal horizontalAlign="space-between">
                      <Text variant="medium" styles={{ root: { fontWeight: 600, fontFamily: 'monospace' } }}>
                        {selectedFunction.name}
                      </Text>
                      <Text variant="small" styles={{ root: { color: '#666' } }}>
                        {selectedFunction.category.replace('_', ' ')}
                      </Text>
                    </Stack>
                    <Text variant="small">{selectedFunction.description}</Text>
                    <Text
                      variant="small"
                      styles={{
                        root: {
                          fontFamily: 'monospace',
                          backgroundColor: '#f0f0f0',
                          padding: '4px 8px',
                          borderRadius: 3,
                        },
                      }}
                    >
                      {selectedFunction.syntax}
                    </Text>
                    {selectedFunction.parameters.length > 0 && (
                      <Stack tokens={{ childrenGap: 4 }}>
                        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                          Parameters:
                        </Text>
                        {selectedFunction.parameters.map((param, idx) => (
                          <Text key={idx} variant="small">
                            • {param.name} ({param.type}){param.optional ? ' - optional' : ''}: {param.description}
                          </Text>
                        ))}
                      </Stack>
                    )}
                  </Stack>
                )}
              </Stack>
            </PivotItem>
          </Pivot>
        </>
      )}

      {/* Loading State */}
      {isAnalyzing && (
        <Shimmer
          shimmerElements={[
            { type: ShimmerElementType.line, width: '100%', height: 16 },
            { type: ShimmerElementType.gap, width: '100%', height: 8 },
            { type: ShimmerElementType.line, width: '80%', height: 16 },
            { type: ShimmerElementType.gap, width: '100%', height: 8 },
            { type: ShimmerElementType.line, width: '60%', height: 16 },
          ]}
        />
      )}
    </Stack>
  );
};

export default DAXExplainer;
