/**
 * Excel AI Assistant - Recipe Builder Component
 * Create and edit custom recipes/templates
 * 
 * @module components/RecipeBuilder
 */

import React, { useState, useEffect } from "react";
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  IconButton,
  TextField,
  Dropdown,
  IDropdownOption,
  MessageBar,
  MessageBarType,
  Separator,
  Pivot,
  PivotItem,
  Icon,
  TooltipHost,
  Dialog,
  DialogType,
  DialogFooter,
  CommandBar,
  TagPicker,
  ITag,
  Toggle,
  Label
} from "@fluentui/react";
import {
  recipeService,
  Recipe,
  RecipeCategory,
  RecipeInput
} from "../services/recipeService";

// ============================================================================
// Types
// ============================================================================

type BuilderTab = "details" | "inputs" | "preview";

interface RecipeBuilderProps {
  recipe?: Recipe | null;
  onSave?: (recipe: Recipe) => void;
  onCancel?: () => void;
}

// ============================================================================
// Component
// ============================================================================

export const RecipeBuilder: React.FC<RecipeBuilderProps> = ({
  recipe,
  onSave,
  onCancel
}) => {
  const isEditing = !!recipe;

  // State
  const [activeTab, setActiveTab] = useState<BuilderTab>("details");
  const [name, setName] = useState(recipe?.name || "");
  const [description, setDescription] = useState(recipe?.description || "");
  const [category, setCategory] = useState<RecipeCategory>(recipe?.category || "custom");
  const [tags, setTags] = useState<ITag[]>(recipe?.tags.map(t => ({ key: t, name: t })) || []);
  const [prompt, setPrompt] = useState(recipe?.prompt || "");
  const [inputs, setInputs] = useState<RecipeInput[]>(recipe?.expectedInputs || []);
  const [difficulty, setDifficulty] = useState<"beginner" | "intermediate" | "advanced">(recipe?.difficulty || "intermediate");
  const [estimatedTime, setEstimatedTime] = useState(recipe?.estimatedTime || "");
  const [errors, setErrors] = useState<string[]>([]);
  const [showDiscardDialog, setShowDiscardDialog] = useState(false);

  // Add new input state
  const [newInputName, setNewInputName] = useState("");
  const [newInputType, setNewInputType] = useState<RecipeInput["type"]>("text");
  const [newInputLabel, setNewInputLabel] = useState("");
  const [newInputRequired, setNewInputRequired] = useState(false);
  const [newInputDescription, setNewInputDescription] = useState("");
  const [newInputOptions, setNewInputOptions] = useState("");

  // Options for dropdowns
  const categoryOptions: IDropdownOption[] = recipeService.getAllCategories().map(cat => ({
    key: cat.key,
    text: cat.name
  }));

  const difficultyOptions: IDropdownOption[] = [
    { key: "beginner", text: "Beginner" },
    { key: "intermediate", text: "Intermediate" },
    { key: "advanced", text: "Advanced" }
  ];

  const inputTypeOptions: IDropdownOption[] = [
    { key: "text", text: "Text" },
    { key: "number", text: "Number" },
    { key: "range", text: "Range" },
    { key: "column", text: "Column" },
    { key: "table", text: "Table" },
    { key: "boolean", text: "Boolean" },
    { key: "select", text: "Select (Dropdown)" },
    { key: "multi-select", text: "Multi-Select" }
  ];

  // Validation
  const validate = (): boolean => {
    const newErrors: string[] = [];

    if (!name.trim()) {
      newErrors.push("Recipe name is required");
    }

    if (!description.trim()) {
      newErrors.push("Description is required");
    }

    if (!prompt.trim()) {
      newErrors.push("Prompt is required");
    }

    if (inputs.length === 0) {
      newErrors.push("At least one input is required");
    }

    setErrors(newErrors);
    return newErrors.length === 0;
  };

  // Save recipe
  const handleSave = () => {
    if (!validate()) {
      setActiveTab("details");
      return;
    }

    const recipeData = {
      name: name.trim(),
      description: description.trim(),
      category,
      tags: tags.map(t => t.name),
      prompt: prompt.trim(),
      expectedInputs: inputs,
      difficulty,
      estimatedTime: estimatedTime.trim() || undefined,
      author: "You",
      isPublic: false,
      isBuiltIn: false,
      version: "1.0"
    };

    if (isEditing && recipe) {
      const updated = recipeService.updateRecipe(recipe.id, recipeData);
      if (updated) {
        onSave?.(updated);
      }
    } else {
      const created = recipeService.createRecipe(recipeData);
      onSave?.(created);
    }
  };

  // Add input
  const handleAddInput = () => {
    if (!newInputName.trim() || !newInputLabel.trim()) {
      return;
    }

    const newInput: RecipeInput = {
      name: newInputName.trim(),
      type: newInputType,
      label: newInputLabel.trim(),
      required: newInputRequired,
      description: newInputDescription.trim() || undefined,
      options: newInputType === "select" || newInputType === "multi-select"
        ? newInputOptions.split(",").map(o => o.trim()).filter(o => o)
        : undefined
    };

    setInputs([...inputs, newInput]);

    // Reset form
    setNewInputName("");
    setNewInputType("text");
    setNewInputLabel("");
    setNewInputRequired(false);
    setNewInputDescription("");
    setNewInputOptions("");
  };

  // Remove input
  const handleRemoveInput = (index: number) => {
    setInputs(inputs.filter((_, i) => i !== index));
  };

  // Move input
  const handleMoveInput = (index: number, direction: "up" | "down") => {
    if (direction === "up" && index > 0) {
      const newInputs = [...inputs];
      [newInputs[index - 1], newInputs[index]] = [newInputs[index], newInputs[index - 1]];
      setInputs(newInputs);
    } else if (direction === "down" && index < inputs.length - 1) {
      const newInputs = [...inputs];
      [newInputs[index], newInputs[index + 1]] = [newInputs[index + 1], newInputs[index]];
      setInputs(newInputs);
    }
  };

  // ============================================================================
  // Render Tabs
  // ============================================================================

  const renderDetailsTab = () => (
    <Stack tokens={{ childrenGap: 16 }}>
      {errors.length > 0 && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {errors.map((e, i) => (
            <div key={i}>• {e}</div>
          ))}
        </MessageBar>
      )}

      <TextField
        label="Recipe Name"
        required
        value={name}
        onChange={(_, value) => setName(value || "")}
        placeholder="e.g., Monthly Sales Report"
      />

      <TextField
        label="Description"
        required
        multiline
        rows={3}
        value={description}
        onChange={(_, value) => setDescription(value || "")}
        placeholder="Briefly describe what this recipe does..."
      />

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <Dropdown
          label="Category"
          required
          selectedKey={category}
          options={categoryOptions}
          onChange={(_, option) => setCategory(option?.key as RecipeCategory)}
          styles={{ root: { width: 200 } }}
        />
        <Dropdown
          label="Difficulty"
          required
          selectedKey={difficulty}
          options={difficultyOptions}
          onChange={(_, option) => setDifficulty(option?.key as "beginner" | "intermediate" | "advanced")}
          styles={{ root: { width: 150 } }}
        />
        <TextField
          label="Estimated Time"
          value={estimatedTime}
          onChange={(_, value) => setEstimatedTime(value || "")}
          placeholder="e.g., 5 min"
          styles={{ root: { width: 120 } }}
        />
      </Stack>

      <div>
        <Label>Tags</Label>
        <TagPicker
          onResolveSuggestions={(filter) => {
            const allTags = recipeService.getAllTags();
            return allTags
              .filter(t => t.toLowerCase().includes(filter.toLowerCase()))
              .map(t => ({ key: t, name: t }));
          }}
          selectedItems={tags}
          onChange={(items) => setTags(items || [])}
          itemLimit={10}
          inputProps={{ placeholder: "Add tags..." }}
        />
      </div>
    </Stack>
  );

  const renderInputsTab = () => (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        Recipe Inputs
      </Text>
      <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
        Define the inputs users will provide when running this recipe. These will be substituted into your prompt.
      </Text>

      {/* Existing Inputs */}
      {inputs.length > 0 && (
        <Stack tokens={{ childrenGap: 8 }}>
          {inputs.map((input, index) => (
            <div
              key={index}
              style={{
                padding: 12,
                backgroundColor: "#f3f2f1",
                borderRadius: 4,
                borderLeft: "3px solid #0078d4"
              }}
            >
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack>
                  <Text styles={{ root: { fontWeight: 600 } }}>
                    {input.label}
                    {input.required && <span style={{ color: "#d13438" }}> *</span>}
                  </Text>
                  <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                    Type: {input.type} | Name: {"{"}{input.name}{"}"}
                  </Text>
                  {input.description && (
                    <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                      {input.description}
                    </Text>
                  )}
                </Stack>
                <Stack horizontal>
                  <IconButton
                    iconProps={{ iconName: "Up" }}
                    title="Move Up"
                    disabled={index === 0}
                    onClick={() => handleMoveInput(index, "up")}
                  />
                  <IconButton
                    iconProps={{ iconName: "Down" }}
                    title="Move Down"
                    disabled={index === inputs.length - 1}
                    onClick={() => handleMoveInput(index, "down")}
                  />
                  <IconButton
                    iconProps={{ iconName: "Delete" }}
                    title="Remove"
                    onClick={() => handleRemoveInput(index)}
                  />
                </Stack>
              </Stack>
            </div>
          ))}
        </Stack>
      )}

      <Separator />

      {/* Add New Input */}
      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
        Add New Input
      </Text>

      <Stack horizontal tokens={{ childrenGap: 16 }}>
        <TextField
          label="Input Name"
          required
          value={newInputName}
          onChange={(_, value) => setNewInputName(value || "")}
          placeholder="e.g., dataRange"
          styles={{ root: { width: 150 } }}
        />
        <TextField
          label="Display Label"
          required
          value={newInputLabel}
          onChange={(_, value) => setNewInputLabel(value || "")}
          placeholder="e.g., Data Range"
          styles={{ root: { width: 200 } }}
        />
        <Dropdown
          label="Type"
          required
          selectedKey={newInputType}
          options={inputTypeOptions}
          onChange={(_, option) => setNewInputType(option?.key as RecipeInput["type"])}
          styles={{ root: { width: 180 } }}
        />
      </Stack>

      <TextField
        label="Description"
        value={newInputDescription}
        onChange={(_, value) => setNewInputDescription(value || "")}
        placeholder="Help text shown to users..."
      />

      {(newInputType === "select" || newInputType === "multi-select") && (
        <TextField
          label="Options (comma-separated)"
          value={newInputOptions}
          onChange={(_, value) => setNewInputOptions(value || "")}
          placeholder="e.g., Option 1, Option 2, Option 3"
        />
      )}

      <Toggle
        label="Required"
        checked={newInputRequired}
        onChange={(_, checked) => setNewInputRequired(checked || false)}
      />

      <DefaultButton
        text="Add Input"
        iconProps={{ iconName: "Add" }}
        onClick={handleAddInput}
        disabled={!newInputName.trim() || !newInputLabel.trim()}
      />
    </Stack>
  );

  const renderPromptTab = () => (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        Recipe Prompt
      </Text>
      <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
        Write the instructions that will be sent to the AI. Use {"{"}inputName{"}"} to reference user inputs.
      </Text>

      <TextField
        multiline
        rows={20}
        value={prompt}
        onChange={(_, value) => setPrompt(value || "")}
        placeholder={`Example:
Create a monthly sales report with the following components:

1. Summary table with total sales from {dataRange}
2. Line chart showing daily sales trend for {month}
3. Format all currency values in {currency}
4. Highlight the top 10 products

Available input variables:
${inputs.map(i => `• {${i.name}} - ${i.label}`).join("\n")}
`}
        styles={{
          field: {
            fontFamily: "Consolas, monospace",
            fontSize: 13
          }
        }}
      />

      <MessageBar messageBarType={MessageBarType.info}>
        <strong>Tip:</strong> Use clear, step-by-step instructions for best results. Number your steps and be specific about what Excel operations should be performed.
      </MessageBar>
    </Stack>
  );

  const renderPreviewTab = () => (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
        Recipe Preview
      </Text>

      <div
        style={{
          padding: 24,
          backgroundColor: "#f3f2f1",
          borderRadius: 8
        }}
      >
        <Stack tokens={{ childrenGap: 12 }}>
          <Text variant="xxLarge" styles={{ root: { fontWeight: 600 } }}>
            {name || "Untitled Recipe"}
          </Text>

          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <span
              style={{
                backgroundColor: "#0078d4",
                color: "white",
                padding: "2px 8px",
                borderRadius: 12,
                fontSize: 11
              }}
            >
              {recipeService.getCategoryDisplayName(category)}
            </span>
            <span
              style={{
                backgroundColor: "#f3f2f1",
                border: "1px solid #e1dfdd",
                padding: "2px 8px",
                borderRadius: 12,
                fontSize: 11
              }}
            >
              {difficulty}
            </span>
            {estimatedTime && (
              <span
                style={{
                  backgroundColor: "#f3f2f1",
                  border: "1px solid #e1dfdd",
                  padding: "2px 8px",
                  borderRadius: 12,
                  fontSize: 11
                }}
              >
                <Icon iconName="Clock" styles={{ root: { fontSize: 10, marginRight: 4 } }} />
                {estimatedTime}
              </span>
            )}
          </Stack>

          <Text styles={{ root: { marginTop: 8 } }}>
            {description || "No description provided"}
          </Text>

          <Separator />

          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            Required Inputs ({inputs.length})
          </Text>
          {inputs.length === 0 ? (
            <Text styles={{ root: { color: "#d13438" } }}>
              No inputs defined. Add at least one input.
            </Text>
          ) : (
            <Stack tokens={{ childrenGap: 8 }}>
              {inputs.map((input, index) => (
                <div
                  key={index}
                  style={{
                    padding: 8,
                    backgroundColor: "white",
                    borderRadius: 4,
                    border: "1px solid #e1dfdd"
                  }}
                >
                  <Text styles={{ root: { fontWeight: 600 } }}>
                    {input.label}
                    {input.required && <span style={{ color: "#d13438" }}> *</span>}
                  </Text>
                  <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                    Type: {input.type} | Variable: {"{"}{input.name}{"}"}
                  </Text>
                </div>
              ))}
            </Stack>
          )}

          <Separator />

          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            AI Prompt
          </Text>
          <div
            style={{
              backgroundColor: "white",
              padding: 12,
              borderRadius: 4,
              border: "1px solid #e1dfdd",
              fontFamily: "Consolas, monospace",
              fontSize: 13,
              whiteSpace: "pre-wrap"
            }}
          >
            {prompt || "No prompt defined"}
          </div>
        </Stack>
      </div>
    </Stack>
  );

  // ============================================================================
  // Main Render
  // ============================================================================

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { maxWidth: 900 } }}>
      <CommandBar
        items={[
          {
            key: "save",
            text: isEditing ? "Update Recipe" : "Create Recipe",
            iconProps: { iconName: "Save" },
            onClick: handleSave
          },
          {
            key: "cancel",
            text: "Cancel",
            iconProps: { iconName: "Cancel" },
            onClick: () => {
              if (name || description || prompt) {
                setShowDiscardDialog(true);
              } else {
                onCancel?.();
              }
            }
          }
        ]}
      />

      <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey as BuilderTab)}>
        <PivotItem headerText="Details" itemKey="details" itemIcon="ContactInfo">
          {renderDetailsTab()}
        </PivotItem>
        <PivotItem headerText="Inputs" itemKey="inputs" itemIcon="TextField">
          {renderInputsTab()}
        </PivotItem>
        <PivotItem headerText="Prompt" itemKey="prompt" itemIcon="AlignLeft">
          {renderPromptTab()}
        </PivotItem>
        <PivotItem headerText="Preview" itemKey="preview" itemIcon="View">
          {renderPreviewTab()}
        </PivotItem>
      </Pivot>

      {/* Discard Dialog */}
      <Dialog
        hidden={!showDiscardDialog}
        onDismiss={() => setShowDiscardDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Discard Changes?",
          subText: "You have unsaved changes. Are you sure you want to discard them?"
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setShowDiscardDialog(false)} text="Keep Editing" />
          <PrimaryButton
            onClick={() => {
              setShowDiscardDialog(false);
              onCancel?.();
            }}
            text="Discard"
            styles={{ root: { backgroundColor: "#d13438" } }}
          />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default RecipeBuilder;
