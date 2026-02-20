/**
 * Excel AI Assistant - Power Query Builder Component
 * Interactive UI for building and managing Power Queries
 * 
 * @module components/PowerQueryBuilder
 */

import React, { useState, useEffect, useCallback } from "react";
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  IconButton,
  TextField,
  Dropdown,
  IDropdownOption,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Pivot,
  PivotItem,
  Icon,
  TooltipHost,
  Separator,
  Dialog,
  DialogType,
  DialogFooter,
  ActionButton,
  Callout,
  Text as FluentText,
  CommandBar,
  ICommandBarItemProps
} from "@fluentui/react";
import {
  powerQueryService,
  PowerQueryInfo,
  QueryModificationResult,
  QueryDependencyGraph
} from "../services/powerQueryService";
import {
  mCodeGenerator,
  PowerQueryOperation,
  OperationCategory,
  MCodeExplanation,
  ValidationResult
} from "../services/mCodeGenerator";

// ============================================================================
// Types
// ============================================================================

type BuilderTab = "queries" | "builder" | "editor" | "dependencies";

interface PowerQueryBuilderProps {
  onQueryCreated?: (queryName: string) => void;
  onQueryUpdated?: (queryName: string) => void;
}

interface QueryWithMetadata extends PowerQueryInfo {
  explanation?: MCodeExplanation;
  isValid?: boolean;
  validationErrors?: string[];
}

// ============================================================================
// Component
// ============================================================================

export const PowerQueryBuilder: React.FC<PowerQueryBuilderProps> = ({
  onQueryCreated,
  onQueryUpdated
}) => {
  // State
  const [activeTab, setActiveTab] = useState<BuilderTab>("queries");
  const [queries, setQueries] = useState<QueryWithMetadata[]>([]);
  const [selectedQuery, setSelectedQuery] = useState<QueryWithMetadata | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Builder state
  const [operations, setOperations] = useState<Record<OperationCategory, PowerQueryOperation[]>>({
    source: [], transform: [], merge: [], aggregate: [], pivot: [], filter: [], sort: [], addColumn: [], group: []
  });
  const [selectedCategory, setSelectedCategory] = useState<OperationCategory>("source");
  const [selectedOperation, setSelectedOperation] = useState<PowerQueryOperation | null>(null);
  const [operationInputs, setOperationInputs] = useState<Record<string, any>>({});
  const [generatedMCode, setGeneratedMCode] = useState("");

  // Editor state
  const [editorMCode, setEditorMCode] = useState("");
  const [editorQueryName, setEditorQueryName] = useState("");
  const [validationResult, setValidationResult] = useState<ValidationResult | null>(null);
  const [explanation, setExplanation] = useState<MCodeExplanation | null>(null);

  // Dependencies state
  const [dependencyGraph, setDependencyGraph] = useState<QueryDependencyGraph | null>(null);

  // Dialogs
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [showCreateDialog, setShowCreateDialog] = useState(false);
  const [newQueryName, setNewQueryName] = useState("");
  const [newQueryDescription, setNewQueryDescription] = useState("");

  // Load queries on mount
  const loadQueries = useCallback(async () => {
    setIsLoading(true);
    setError(null);

    try {
      const queryInfos = await powerQueryService.getAllQueries();

      // Enhance with metadata
      const enhancedQueries: QueryWithMetadata[] = await Promise.all(
        queryInfos.map(async (q) => {
          const validation = mCodeGenerator.validateMCode(q.formula);
          const explanation = mCodeGenerator.explainMCode(q.formula);

          return {
            ...q,
            isValid: validation.isValid,
            validationErrors: validation.errors.map(e => e.message),
            explanation
          };
        })
      );

      setQueries(enhancedQueries);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load queries");
    } finally {
      setIsLoading(false);
    }
  }, []);

  useEffect(() => {
    loadQueries();
    setOperations(mCodeGenerator.getOperationsByCategory());
  }, [loadQueries]);

  // ============================================================================
  // Query List Tab
  // ============================================================================

  const renderQueriesList = () => {
    const columns: IColumn[] = [
      {
        key: "status",
        name: "",
        fieldName: "isValid",
        minWidth: 30,
        maxWidth: 30,
        onRender: (item: QueryWithMetadata) => (
          <Icon
            iconName={item.isValid ? "CheckMark" : "ErrorBadge"}
            styles={{
              root: {
                color: item.isValid ? "#107c10" : "#d13438",
                fontSize: 16
              }
            }}
          />
        )
      },
      { key: "name", name: "Query Name", fieldName: "name", minWidth: 150, maxWidth: 200 },
      { key: "source", name: "Source", fieldName: "source", minWidth: 100, maxWidth: 150 },
      {
        key: "steps",
        name: "Steps",
        minWidth: 60,
        maxWidth: 80,
        onRender: (item: QueryWithMetadata) => (
          <Text>{item.explanation?.steps.length || 0}</Text>
        )
      },
      {
        key: "actions",
        name: "Actions",
        minWidth: 120,
        maxWidth: 150,
        onRender: (item: QueryWithMetadata) => (
          <Stack horizontal tokens={{ childrenGap: 4 }}>
            <IconButton
              iconProps={{ iconName: "Edit" }}
              title="Edit"
              onClick={() => handleEditQuery(item)}
            />
            <IconButton
              iconProps={{ iconName: "Refresh" }}
              title="Refresh"
              onClick={() => handleRefreshQuery(item.name)}
            />
            <IconButton
              iconProps={{ iconName: "Copy" }}
              title="Duplicate"
              onClick={() => handleDuplicateQuery(item.name)}
            />
            <IconButton
              iconProps={{ iconName: "Delete" }}
              title="Delete"
              onClick={() => {
                setSelectedQuery(item);
                setShowDeleteDialog(true);
              }}
            />
          </Stack>
        )
      }
    ];

    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <CommandBar
          items={[
            {
              key: "new",
              text: "New Query",
              iconProps: { iconName: "Add" },
              onClick: () => {
                setNewQueryName("");
                setNewQueryDescription("");
                setShowCreateDialog(true);
              }
            },
            {
              key: "refresh",
              text: "Refresh All",
              iconProps: { iconName: "Refresh" },
              onClick: handleRefreshAllQueries
            },
            {
              key: "dependencies",
              text: "View Dependencies",
              iconProps: { iconName: "Relationship" },
              onClick: () => setActiveTab("dependencies")
            }
          ]}
        />

        {isLoading ? (
          <ProgressIndicator label="Loading queries..." />
        ) : queries.length === 0 ? (
          <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
            <Icon iconName="DataFlow" styles={{ root: { fontSize: 64, color: "#c8c6c4" } }} />
            <Text variant="large" styles={{ root: { color: "#605e5c", marginTop: 16 } }}>
              No Power Queries found in this workbook
            </Text>
            <PrimaryButton
              text="Create Your First Query"
              iconProps={{ iconName: "Add" }}
              onClick={() => setActiveTab("builder")}
              styles={{ root: { marginTop: 16 } }}
            />
          </Stack>
        ) : (
          <DetailsList
            items={queries}
            columns={columns}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.single}
            onItemInvoked={(item) => handleEditQuery(item as QueryWithMetadata)}
          />
        )}
      </Stack>
    );
  };

  // ============================================================================
  // Query Builder Tab
  // ============================================================================

  const renderBuilder = () => {
    const categoryOptions: IDropdownOption[] = [
      { key: "source", text: "📥 Data Source" },
      { key: "filter", text: "🔍 Filter" },
      { key: "transform", text: "🔧 Transform" },
      { key: "addColumn", text: "➕ Add Column" },
      { key: "group", text: "📊 Group By" },
      { key: "sort", text: "📋 Sort" },
      { key: "pivot", text: "🔄 Pivot/Unpivot" },
      { key: "merge", text: "🔗 Merge/Append" }
    ];

    const operationOptions: IDropdownOption[] = operations[selectedCategory]?.map(op => ({
      key: op.name,
      text: op.name,
      data: op
    })) || [];

    return (
      <Stack tokens={{ childrenGap: 16 }} horizontal>
        {/* Left panel - Operation selection */}
        <Stack styles={{ root: { width: 300, borderRight: "1px solid #e1dfdd", paddingRight: 16 } }}>
          <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 12 } }}>
            Step Builder
          </Text>

          <Dropdown
            label="Category"
            selectedKey={selectedCategory}
            options={categoryOptions}
            onChange={(_, option) => {
              setSelectedCategory(option?.key as OperationCategory);
              setSelectedOperation(null);
              setOperationInputs({});
              setGeneratedMCode("");
            }}
          />

          <Dropdown
            label="Operation"
            selectedKey={selectedOperation?.name}
            options={operationOptions}
            placeholder="Select an operation"
            onChange={(_, option) => {
              const op = option?.data as PowerQueryOperation;
              setSelectedOperation(op);

              // Set default values
              const defaults: Record<string, any> = {};
              op?.inputs.forEach(input => {
                if (input.defaultValue !== undefined) {
                  defaults[input.name] = input.defaultValue;
                }
              });
              setOperationInputs(defaults);
              updateGeneratedMCode(op, defaults);
            }}
            styles={{ root: { marginTop: 8 } }}
          />

          {selectedOperation && (
            <>
              <Separator />
              <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                {selectedOperation.description}
              </Text>

              <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
                {selectedOperation.inputs.map(input => (
                  <Stack key={input.name} tokens={{ childrenGap: 4 }}>
                    <FluentText variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      {input.name} {input.required && <span style={{ color: "#d13438" }}>*</span>}
                    </FluentText>

                    {input.options ? (
                      <Dropdown
                        selectedKey={operationInputs[input.name]}
                        options={input.options.map(o => ({ key: o, text: o }))}
                        onChange={(_, option) => {
                          const newInputs = { ...operationInputs, [input.name]: option?.key };
                          setOperationInputs(newInputs);
                          updateGeneratedMCode(selectedOperation, newInputs);
                        }}
                      />
                    ) : input.type === "boolean" ? (
                      <Dropdown
                        selectedKey={operationInputs[input.name]?.toString()}
                        options={[{ key: "true", text: "Yes" }, { key: "false", text: "No" }]}
                        onChange={(_, option) => {
                          const newInputs = { ...operationInputs, [input.name]: option?.key === "true" };
                          setOperationInputs(newInputs);
                          updateGeneratedMCode(selectedOperation, newInputs);
                        }}
                      />
                    ) : (
                      <TextField
                        value={operationInputs[input.name] || ""}
                        placeholder={input.description}
                        onChange={(_, value) => {
                          const newInputs = { ...operationInputs, [input.name]: value };
                          setOperationInputs(newInputs);
                          updateGeneratedMCode(selectedOperation, newInputs);
                        }}
                      />
                    )}
                  </Stack>
                ))}
              </Stack>
            </>
          )}
        </Stack>

        {/* Right panel - Preview */}
        <Stack styles={{ root: { flex: 1, paddingLeft: 16 } }}>
          <Text variant="large" styles={{ root: { fontWeight: 600, marginBottom: 12 } }}>
            Generated M Code
          </Text>

          <div
            style={{
              backgroundColor: "#f3f2f1",
              padding: 16,
              borderRadius: 4,
              fontFamily: "Consolas, monospace",
              fontSize: 13,
              minHeight: 200,
              whiteSpace: "pre-wrap"
            }}
          >
            {generatedMCode || "Select an operation to generate M code"}
          </div>

          {generatedMCode && (
            <>
              <Separator />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <PrimaryButton
                  text="Add to New Query"
                  iconProps={{ iconName: "Add" }}
                  onClick={() => {
                    setEditorMCode(generatedMCode);
                    setActiveTab("editor");
                  }}
                />
                <DefaultButton
                  text="Copy to Clipboard"
                  iconProps={{ iconName: "Copy" }}
                  onClick={() => navigator.clipboard.writeText(generatedMCode)}
                />
              </Stack>

              <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginTop: 12 } }}>
                This generates a single step. For a complete query, use the Editor tab to build a sequence of steps.
              </MessageBar>
            </>
          )}
        </Stack>
      </Stack>
    );
  };

  // ============================================================================
  // Query Editor Tab
  // ============================================================================

  const renderEditor = () => {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            M Code Editor
          </Text>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <DefaultButton
              text="Format"
              iconProps={{ iconName: "AlignLeft" }}
              onClick={() => {
                const formatted = mCodeGenerator.formatMCode(editorMCode);
                setEditorMCode(formatted);
              }}
            />
            <PrimaryButton
              text="Validate"
              iconProps={{ iconName: "CheckMark" }}
              onClick={handleValidateMCode}
            />
          </Stack>
        </Stack>

        <Stack horizontal tokens={{ childrenGap: 16 }}>
          <Stack styles={{ root: { flex: 1 } }}>
            <TextField
              label="Query Name"
              value={editorQueryName}
              onChange={(_, value) => setEditorQueryName(value || "")}
              placeholder="Enter query name"
            />

            <TextField
              label="M Code"
              multiline
              rows={15}
              value={editorMCode}
              onChange={(_, value) => {
                setEditorMCode(value || "");
                setValidationResult(null);
                setExplanation(null);
              }}
              styles={{
                field: {
                  fontFamily: "Consolas, monospace",
                  fontSize: 13
                }
              }}
            />

            <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
              <PrimaryButton
                text={selectedQuery ? "Update Query" : "Create Query"}
                iconProps={{ iconName: "Save" }}
                disabled={!editorQueryName || !editorMCode}
                onClick={handleSaveQuery}
              />
              {selectedQuery && (
                <DefaultButton
                  text="Reset"
                  onClick={() => {
                    setEditorMCode(selectedQuery.formula);
                    setEditorQueryName(selectedQuery.name);
                    setValidationResult(null);
                    setExplanation(null);
                  }}
                />
              )}
            </Stack>
          </Stack>

          <Stack styles={{ root: { width: 400 } }}>
            {validationResult && (
              <>
                <MessageBar
                  messageBarType={validationResult.isValid ? MessageBarType.success : MessageBarType.error}
                >
                  {validationResult.isValid
                    ? "M code is valid!"
                    : `Found ${validationResult.errors.length} error(s)`}
                </MessageBar>

                {!validationResult.isValid && validationResult.errors.length > 0 && (
                  <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
                    {validationResult.errors.map((error, index) => (
                      <MessageBar key={index} messageBarType={MessageBarType.error}>
                        Line {error.line}: {error.message}
                      </MessageBar>
                    ))}
                  </Stack>
                )}

                {validationResult.warnings.length > 0 && (
                  <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                      Warnings:
                    </Text>
                    {validationResult.warnings.map((warning, index) => (
                      <MessageBar key={index} messageBarType={MessageBarType.warning}>
                        {warning.message}
                      </MessageBar>
                    ))}
                  </Stack>
                )}
              </>
            )}

            {explanation && (
              <>
                <Separator />
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                  Explanation
                </Text>
                <Text>{explanation.summary}</Text>

                <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 12 } }}>
                  Steps ({explanation.steps.length}):
                </Text>
                <Stack tokens={{ childrenGap: 8 }}>
                  {explanation.steps.map((step, index) => (
                    <div
                      key={index}
                      style={{
                        padding: 8,
                        backgroundColor: "#f3f2f1",
                        borderRadius: 4,
                        borderLeft: "3px solid #0078d4"
                      }}
                    >
                      <Text variant="smallPlus" styles={{ root: { fontWeight: 600, color: "#0078d4" } }}>
                        {step.stepName}
                      </Text>
                      <Text variant="small">{step.description}</Text>
                    </div>
                  ))}
                </Stack>

                {explanation.performanceNotes.length > 0 && (
                  <>
                    <Text variant="smallPlus" styles={{ root: { fontWeight: 600, marginTop: 12 } }}>
                      Performance Notes:
                    </Text>
                    {explanation.performanceNotes.map((note, index) => (
                      <MessageBar key={index} messageBarType={MessageBarType.warning}>
                        {note}
                      </MessageBar>
                    ))}
                  </>
                )}
              </>
            )}
          </Stack>
        </Stack>
      </Stack>
    );
  };

  // ============================================================================
  // Dependencies Tab
  // ============================================================================

  const renderDependencies = () => {
    return (
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal horizontalAlign="space-between">
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
            Query Dependencies
          </Text>
          <DefaultButton
            text="Refresh Graph"
            iconProps={{ iconName: "Refresh" }}
            onClick={async () => {
              const graph = await powerQueryService.buildDependencyGraph();
              setDependencyGraph(graph);
            }}
          />
        </Stack>

        {dependencyGraph ? (
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal tokens={{ childrenGap: 32 }}>
              <StatCard label="Total Queries" value={dependencyGraph.queries.length.toString()} />
              <StatCard label="Root Queries" value={dependencyGraph.rootQueries.length.toString()} />
              <StatCard
                label="With Dependencies"
                value={Array.from(dependencyGraph.dependencies.values()).filter(d => d.length > 0).length.toString()}
              />
            </Stack>

            <Separator />

            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Dependency Map
            </Text>

            <Stack tokens={{ childrenGap: 8 }}>
              {dependencyGraph.queries.map(queryName => {
                const deps = dependencyGraph.dependencies.get(queryName) || [];
                return (
                  <Stack
                    key={queryName}
                    horizontal
                    verticalAlign="center"
                    tokens={{ childrenGap: 12, padding: 12 }}
                    styles={{ root: { backgroundColor: "#f3f2f1", borderRadius: 4 } }}
                  >
                    <Icon
                      iconName={deps.length > 0 ? "DependencyAdd" : "Document"}
                      styles={{ root: { fontSize: 20, color: deps.length > 0 ? "#0078d4" : "#605e5c" } }}
                    />
                    <Stack styles={{ root: { flex: 1 } }}>
                      <Text styles={{ root: { fontWeight: 600 } }}>{queryName}</Text>
                      {deps.length > 0 ? (
                        <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                          Depends on: {deps.join(", ")}
                        </Text>
                      ) : (
                        <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                          No dependencies (root query)
                        </Text>
                      )}
                    </Stack>
                    {dependencyGraph.rootQueries.includes(queryName) && (
                      <span
                        style={{
                          backgroundColor: "#107c10",
                          color: "white",
                          padding: "2px 8px",
                          borderRadius: 12,
                          fontSize: 11
                        }}
                      >
                        Root
                      </span>
                    )}
                  </Stack>
                );
              })}
            </Stack>
          </Stack>
        ) : (
          <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
            <Icon iconName="Relationship" styles={{ root: { fontSize: 64, color: "#c8c6c4" } }} />
            <Text variant="large" styles={{ root: { color: "#605e5c", marginTop: 16 } }}>
              Click "Refresh Graph" to analyze query dependencies
            </Text>
          </Stack>
        )}
      </Stack>
    );
  };

  // ============================================================================
  // Helper Components
  // ============================================================================

  interface StatCardProps {
    label: string;
    value: string;
  }

  const StatCard: React.FC<StatCardProps> = ({ label, value }) => (
    <Stack
      horizontalAlign="center"
      tokens={{ padding: 16, childrenGap: 4 }}
      styles={{
        root: {
          backgroundColor: "#f3f2f1",
          borderRadius: 8,
          minWidth: 100
        }
      }}
    >
      <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: "#0078d4" } }}>
        {value}
      </Text>
      <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
        {label}
      </Text>
    </Stack>
  );

  // ============================================================================
  // Event Handlers
  // ============================================================================

  const handleEditQuery = (query: QueryWithMetadata) => {
    setSelectedQuery(query);
    setEditorMCode(query.formula);
    setEditorQueryName(query.name);
    setValidationResult(null);
    setExplanation(null);
    setActiveTab("editor");
  };

  const handleRefreshQuery = async (queryName: string) => {
    setIsLoading(true);
    try {
      const result = await powerQueryService.refreshQuery(queryName);
      if (result.success) {
        // Show success message
      } else {
        setError(result.error || "Refresh failed");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Refresh failed");
    } finally {
      setIsLoading(false);
    }
  };

  const handleRefreshAllQueries = async () => {
    setIsLoading(true);
    try {
      await powerQueryService.refreshAllQueries({
        onProgress: (percent) => {
          // Could update progress indicator here
        }
      });
      await loadQueries();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Refresh failed");
    } finally {
      setIsLoading(false);
    }
  };

  const handleDuplicateQuery = async (queryName: string) => {
    try {
      const result = await powerQueryService.duplicateQuery(queryName);
      if (result.success) {
        await loadQueries();
        onQueryCreated?.(result.queryName!);
      } else {
        setError(result.error || "Duplicate failed");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Duplicate failed");
    }
  };

  const handleDeleteQuery = async () => {
    if (!selectedQuery) return;

    try {
      const success = await powerQueryService.deleteQuery(selectedQuery.name);
      if (success) {
        await loadQueries();
        setSelectedQuery(null);
        setShowDeleteDialog(false);
      } else {
        setError("Failed to delete query");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Delete failed");
    }
  };

  const handleValidateMCode = () => {
    const validation = mCodeGenerator.validateMCode(editorMCode);
    setValidationResult(validation);

    if (validation.isValid) {
      const explanation = mCodeGenerator.explainMCode(editorMCode);
      setExplanation(explanation);
    }
  };

  const handleSaveQuery = async () => {
    if (!editorQueryName || !editorMCode) return;

    try {
      let result: QueryModificationResult;

      if (selectedQuery && selectedQuery.name === editorQueryName) {
        // Update existing
        result = await powerQueryService.updateQuery(editorQueryName, editorMCode);
        if (result.success) {
          onQueryUpdated?.(editorQueryName);
        }
      } else {
        // Create new
        result = await powerQueryService.createQuery(editorQueryName, editorMCode, newQueryDescription);
        if (result.success) {
          onQueryCreated?.(editorQueryName);
        }
      }

      if (result.success) {
        await loadQueries();
        setActiveTab("queries");
      } else {
        setError(result.error || "Save failed");
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : "Save failed");
    }
  };

  const updateGeneratedMCode = (operation: PowerQueryOperation, inputs: Record<string, any>) => {
    let code = operation.mCode;

    // Replace placeholders with actual values
    operation.inputs.forEach(input => {
      const value = inputs[input.name];
      if (value !== undefined) {
        const placeholder = `{${input.name}}`;
        code = code.replace(new RegExp(placeholder, "g"), value);
      }
    });

    setGeneratedMCode(code);
  };

  // ============================================================================
  // Main Render
  // ============================================================================

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline
          onDismiss={() => setError(null)}
        >
          {error}
        </MessageBar>
      )}

      <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey as BuilderTab)}>
        <PivotItem headerText="Queries" itemKey="queries" itemIcon="DataFlow">
          {renderQueriesList()}
        </PivotItem>
        <PivotItem headerText="Builder" itemKey="builder" itemIcon="BuildDefinition">
          {renderBuilder()}
        </PivotItem>
        <PivotItem headerText="Editor" itemKey="editor" itemIcon="Code">
          {renderEditor()}
        </PivotItem>
        <PivotItem headerText="Dependencies" itemKey="dependencies" itemIcon="Relationship">
          {renderDependencies()}
        </PivotItem>
      </Pivot>

      {/* Delete Dialog */}
      <Dialog
        hidden={!showDeleteDialog}
        onDismiss={() => setShowDeleteDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Delete Query",
          subText: `Are you sure you want to delete "${selectedQuery?.name}"? This action cannot be undone.`
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setShowDeleteDialog(false)} text="Cancel" />
          <PrimaryButton onClick={handleDeleteQuery} text="Delete" styles={{ root: { backgroundColor: "#d13438" } }} />
        </DialogFooter>
      </Dialog>

      {/* Create Dialog */}
      <Dialog
        hidden={!showCreateDialog}
        onDismiss={() => setShowCreateDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Create New Query"
        }}
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <TextField
            label="Query Name"
            required
            value={newQueryName}
            onChange={(_, value) => setNewQueryName(value || "")}
            placeholder="MyNewQuery"
          />
          <TextField
            label="Description (optional)"
            multiline
            rows={2}
            value={newQueryDescription}
            onChange={(_, value) => setNewQueryDescription(value || "")}
          />
        </Stack>
        <DialogFooter>
          <DefaultButton onClick={() => setShowCreateDialog(false)} text="Cancel" />
          <PrimaryButton
            onClick={() => {
              if (newQueryName) {
                setEditorQueryName(newQueryName);
                setEditorMCode(`let\n    Source = #table({\"Column1\"}, {{\"Data\"}})\nin\n    Source`);
                setShowCreateDialog(false);
                setActiveTab("editor");
              }
            }}
            text="Create"
            disabled={!newQueryName}
          />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default PowerQueryBuilder;
