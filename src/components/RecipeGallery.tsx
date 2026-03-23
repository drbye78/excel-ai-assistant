/**
 * Excel AI Assistant - Recipe Gallery Component
 * Browse, search, and discover reusable Excel recipes/templates
 * 
 * @module components/RecipeGallery
 */

import React, { useState, useEffect, useCallback } from "react";
import { logger } from "../utils/logger";
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
  CommandBar,
  SearchBox,
  TagPicker,
  ITag,
  Rating,
  RatingSize
} from "@fluentui/react";
import {
  recipeService,
  Recipe,
  RecipeCategory,
  RecipeFilters,
  UserRecipe
} from "../services/recipeService";

// ============================================================================
// Types
// ============================================================================

type GalleryTab = "all" | "builtin" | "my" | "recent" | "favorites";

interface RecipeGalleryProps {
  onSelectRecipe?: (recipe: Recipe) => void;
  onExecuteRecipe?: (recipe: Recipe) => void;
  onCreateRecipe?: () => void;
}

// ============================================================================
// Component
// ============================================================================

export const RecipeGallery: React.FC<RecipeGalleryProps> = ({
  onSelectRecipe,
  onExecuteRecipe,
  onCreateRecipe
}) => {
  // State
  const [activeTab, setActiveTab] = useState<GalleryTab>("all");
  const [recipes, setRecipes] = useState<Recipe[]>([]);
  const [filteredRecipes, setFilteredRecipes] = useState<Recipe[]>([]);
  const [selectedRecipe, setSelectedRecipe] = useState<Recipe | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState<RecipeCategory | "all">("all");
  const [selectedDifficulty, setSelectedDifficulty] = useState<string>("all");
  const [selectedSort, setSelectedSort] = useState<string>("popular");

  // Dialogs
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [showShareDialog, setShowShareDialog] = useState(false);
  const [shareCode, setShareCode] = useState("");
  const [showImportDialog, setShowImportDialog] = useState(false);
  const [importData, setImportData] = useState("");

  // Load recipes
  const loadRecipes = useCallback(() => {
    setIsLoading(true);

    try {
      let results: Recipe[] = [];

      switch (activeTab) {
        case "builtin":
          results = recipeService.getRecipes({ isBuiltIn: true });
          break;
        case "my":
          results = recipeService.getRecipes({ isBuiltIn: false });
          break;
        case "recent":
          const recent = recipeService.getRecentlyUsedRecipes(20);
          results = recent;
          break;
        case "favorites":
          results = recipeService.getFrequentlyUsedRecipes(20);
          break;
        case "all":
        default:
          results = recipeService.getAllRecipes();
          break;
      }

      setRecipes(results);
      applyFilters(results);
    } catch (error) {
      logger.error("Failed to load recipes", undefined, error as Error);
    } finally {
      setIsLoading(false);
    }
  }, [activeTab]);

  useEffect(() => {
    loadRecipes();
  }, [loadRecipes]);

  // Apply filters
  const applyFilters = (recipeList: Recipe[]) => {
    const filters: RecipeFilters = {};

    if (searchQuery) {
      filters.search = searchQuery;
    }

    if (selectedCategory !== "all") {
      filters.category = selectedCategory as RecipeCategory;
    }

    if (selectedDifficulty !== "all") {
      filters.difficulty = selectedDifficulty as "beginner" | "intermediate" | "advanced";
    }

    filters.sortBy = selectedSort as RecipeFilters["sortBy"];

    let filtered = recipeList;

    if (searchQuery) {
      const searchLower = searchQuery.toLowerCase();
      filtered = filtered.filter(r =>
        r.name.toLowerCase().includes(searchLower) ||
        r.description.toLowerCase().includes(searchLower) ||
        r.tags.some(t => t.toLowerCase().includes(searchLower))
      );
    }

    if (selectedCategory !== "all") {
      filtered = filtered.filter(r => r.category === selectedCategory);
    }

    if (selectedDifficulty !== "all") {
      filtered = filtered.filter(r => r.difficulty === selectedDifficulty);
    }

    // Sort
    switch (selectedSort) {
      case "popular":
        filtered.sort((a, b) => b.usageCount - a.usageCount);
        break;
      case "newest":
        filtered.sort((a, b) => b.createdAt.getTime() - a.createdAt.getTime());
        break;
      case "rating":
        filtered.sort((a, b) => b.rating - a.rating);
        break;
      case "name":
        filtered.sort((a, b) => a.name.localeCompare(b.name));
        break;
    }

    setFilteredRecipes(filtered);
  };

  useEffect(() => {
    applyFilters(recipes);
  }, [searchQuery, selectedCategory, selectedDifficulty, selectedSort, recipes]);

  // ============================================================================
  // Render
  // ============================================================================

  const renderRecipeCard = (recipe: Recipe) => {
    const isUserRecipe = !recipe.isBuiltIn;
    const difficultyColor = {
      beginner: "#107c10",
      intermediate: "#ffc107",
      advanced: "#d13438"
    }[recipe.difficulty];

    return (
      <div
        key={recipe.id}
        style={{
          backgroundColor: "#ffffff",
          borderRadius: 8,
          border: "1px solid #e1dfdd",
          padding: 16,
          cursor: "pointer",
          transition: "box-shadow 0.2s",
          position: "relative"
        }}
        onClick={() => {
          setSelectedRecipe(recipe);
          onSelectRecipe?.(recipe);
        }}
        onMouseEnter={(e) => {
          e.currentTarget.style.boxShadow = "0 4px 12px rgba(0,0,0,0.1)";
        }}
        onMouseLeave={(e) => {
          e.currentTarget.style.boxShadow = "none";
        }}
      >
        {/* Header */}
        <Stack horizontal horizontalAlign="space-between" verticalAlign="start">
          <Stack styles={{ root: { flex: 1 } }}>
            <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
              {recipe.name}
            </Text>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
              <span
                style={{
                  backgroundColor: difficultyColor,
                  color: recipe.difficulty === "intermediate" ? "#323130" : "white",
                  padding: "2px 8px",
                  borderRadius: 12,
                  fontSize: 11,
                  fontWeight: 600
                }}
              >
                {recipe.difficulty}
              </span>
              {recipe.estimatedTime && (
                <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
                  <Icon iconName="Clock" styles={{ root: { fontSize: 12, marginRight: 4 } }} />
                  {recipe.estimatedTime}
                </Text>
              )}
            </Stack>
          </Stack>
          <Stack horizontal>
            {isUserRecipe && (
              <IconButton
                iconProps={{ iconName: "Delete" }}
                title="Delete"
                onClick={(e) => {
                  e.stopPropagation();
                  setSelectedRecipe(recipe);
                  setShowDeleteDialog(true);
                }}
              />
            )}
            <IconButton
              iconProps={{ iconName: "Share" }}
              title="Share"
              onClick={(e) => {
                e.stopPropagation();
                handleShareRecipe(recipe);
              }}
            />
          </Stack>
        </Stack>

        {/* Description */}
        <Text
          variant="small"
          styles={{
            root: {
              color: "#605e5c",
              marginTop: 8,
              display: "-webkit-box",
              WebkitLineClamp: 2,
              WebkitBoxOrient: "vertical",
              overflow: "hidden"
            }
          }}
        >
          {recipe.description}
        </Text>

        {/* Tags */}
        <Stack horizontal wrap tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 12 } }}>
          {recipe.tags.slice(0, 3).map((tag, index) => (
            <span
              key={index}
              style={{
                backgroundColor: "#f3f2f1",
                color: "#605e5c",
                padding: "2px 8px",
                borderRadius: 12,
                fontSize: 11
              }}
            >
              {tag}
            </span>
          ))}
          {recipe.tags.length > 3 && (
            <span style={{ color: "#605e5c", fontSize: 11 }}>+{recipe.tags.length - 3}</span>
          )}
        </Stack>

        {/* Footer */}
        <Separator styles={{ root: { margin: "12px 0" } }} />

        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Stack horizontal tokens={{ childrenGap: 16 }}>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
              <Icon iconName="FavoriteStarFill" styles={{ root: { color: "#ffc107", fontSize: 14 } }} />
              <Text variant="small">{recipe.rating.toFixed(1)}</Text>
              <Text variant="small" styles={{ root: { color: "#605e5c" } }}>({recipe.ratingCount})</Text>
            </Stack>
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
              <Icon iconName="Play" styles={{ root: { color: "#0078d4", fontSize: 14 } }} />
              <Text variant="small">{recipe.usageCount.toLocaleString()}</Text>
            </Stack>
          </Stack>
          <PrimaryButton
            text="Use Recipe"
            iconProps={{ iconName: "Play" }}
            onClick={(e) => {
              e.stopPropagation();
              onExecuteRecipe?.(recipe);
            }}
          />
        </Stack>
      </div>
    );
  };

  const renderFilters = () => {
    const categoryOptions: IDropdownOption[] = [
      { key: "all", text: "All Categories" },
      ...recipeService.getAllCategories().map(cat => ({
        key: cat.key,
        text: cat.name
      }))
    ];

    const difficultyOptions: IDropdownOption[] = [
      { key: "all", text: "All Levels" },
      { key: "beginner", text: "Beginner" },
      { key: "intermediate", text: "Intermediate" },
      { key: "advanced", text: "Advanced" }
    ];

    const sortOptions: IDropdownOption[] = [
      { key: "popular", text: "Most Popular" },
      { key: "newest", text: "Newest" },
      { key: "rating", text: "Highest Rated" },
      { key: "name", text: "Name (A-Z)" }
    ];

    return (
      <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 16 } }}>
        <SearchBox
          placeholder="Search recipes..."
          value={searchQuery}
          onChange={(_, value) => setSearchQuery(value || "")}
          onClear={() => setSearchQuery("")}
          styles={{ root: { width: 300 } }}
        />
        <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
          <Dropdown
            label="Category"
            selectedKey={selectedCategory}
            options={categoryOptions}
            onChange={(_, option) => setSelectedCategory(option?.key as RecipeCategory | "all")}
            styles={{ root: { width: 150 } }}
          />
          <Dropdown
            label="Difficulty"
            selectedKey={selectedDifficulty}
            options={difficultyOptions}
            onChange={(_, option) => setSelectedDifficulty(option?.key as string)}
            styles={{ root: { width: 130 } }}
          />
          <Dropdown
            label="Sort By"
            selectedKey={selectedSort}
            options={sortOptions}
            onChange={(_, option) => setSelectedSort(option?.key as string)}
            styles={{ root: { width: 140 } }}
          />
          <DefaultButton
            text="Clear Filters"
            iconProps={{ iconName: "Clear" }}
            onClick={() => {
              setSearchQuery("");
              setSelectedCategory("all");
              setSelectedDifficulty("all");
              setSelectedSort("popular");
            }}
            styles={{ root: { alignSelf: "flex-end" } }}
          />
        </Stack>
      </Stack>
    );
  };

  const renderRecipeGrid = () => {
    if (isLoading) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 32 }}>
          <ProgressIndicator label="Loading recipes..." />
        </Stack>
      );
    }

    if (filteredRecipes.length === 0) {
      return (
        <Stack horizontalAlign="center" tokens={{ padding: 48 }}>
          <Icon iconName="ClipboardList" styles={{ root: { fontSize: 64, color: "#c8c6c4" } }} />
          <Text variant="large" styles={{ root: { color: "#605e5c", marginTop: 16 } }}>
            No recipes found
          </Text>
          <Text styles={{ root: { color: "#605e5c", marginTop: 8 } }}>
            Try adjusting your filters or create a new recipe
          </Text>
          {activeTab === "all" && (
            <PrimaryButton
              text="Create Recipe"
              iconProps={{ iconName: "Add" }}
              onClick={onCreateRecipe}
              styles={{ root: { marginTop: 16 } }}
            />
          )}
        </Stack>
      );
    }

    return (
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))",
          gap: 16
        }}
      >
        {filteredRecipes.map(renderRecipeCard)}
      </div>
    );
  };

  // ============================================================================
  // Handlers
  // ============================================================================

  const handleShareRecipe = (recipe: Recipe) => {
    const code = recipeService.shareRecipe(recipe.id);
    if (code) {
      setShareCode(code);
      setShowShareDialog(true);
    }
  };

  const handleDeleteRecipe = () => {
    if (!selectedRecipe) return;

    const success = recipeService.deleteRecipe(selectedRecipe.id);
    if (success) {
      loadRecipes();
      setShowDeleteDialog(false);
      setSelectedRecipe(null);
    }
  };

  const handleImportRecipe = () => {
    if (!importData.trim()) return;

    // Try as JSON first
    let recipe = recipeService.importRecipe(importData);

    // Try as share code if JSON fails
    if (!recipe) {
      recipe = recipeService.importFromShareCode(importData);
    }

    if (recipe) {
      loadRecipes();
      setShowImportDialog(false);
      setImportData("");
    } else {
      alert("Invalid recipe data. Please check your input.");
    }
  };

  // ============================================================================
  // Main Render
  // ============================================================================

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {/* Header */}
      <CommandBar
        items={[
          {
            key: "new",
            text: "New Recipe",
            iconProps: { iconName: "Add" },
            onClick: onCreateRecipe
          },
          {
            key: "import",
            text: "Import",
            iconProps: { iconName: "Download" },
            onClick: () => setShowImportDialog(true)
          }
        ]}
      />

      {/* Tabs */}
      <Pivot selectedKey={activeTab} onLinkClick={(item) => setActiveTab(item?.props.itemKey as GalleryTab)}>
        <PivotItem headerText="All Recipes" itemKey="all" itemIcon="AllApps">
          {renderFilters()}
          {renderRecipeGrid()}
        </PivotItem>
        <PivotItem headerText="Built-in" itemKey="builtin" itemIcon="Certificate">
          {renderFilters()}
          {renderRecipeGrid()}
        </PivotItem>
        <PivotItem headerText="My Recipes" itemKey="my" itemIcon="Contact">
          {renderFilters()}
          {renderRecipeGrid()}
        </PivotItem>
        <PivotItem headerText="Recently Used" itemKey="recent" itemIcon="History">
          {renderRecipeGrid()}
        </PivotItem>
        <PivotItem headerText="Most Used" itemKey="favorites" itemIcon="FavoriteStar">
          {renderRecipeGrid()}
        </PivotItem>
      </Pivot>

      {/* Delete Dialog */}
      <Dialog
        hidden={!showDeleteDialog}
        onDismiss={() => setShowDeleteDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Delete Recipe",
          subText: `Are you sure you want to delete "${selectedRecipe?.name}"? This action cannot be undone.`
        }}
      >
        <DialogFooter>
          <DefaultButton onClick={() => setShowDeleteDialog(false)} text="Cancel" />
          <PrimaryButton onClick={handleDeleteRecipe} text="Delete" styles={{ root: { backgroundColor: "#d13438" } }} />
        </DialogFooter>
      </Dialog>

      {/* Share Dialog */}
      <Dialog
        hidden={!showShareDialog}
        onDismiss={() => setShowShareDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Share Recipe",
          subText: "Copy this code to share the recipe with others:"
        }}
      >
        <TextField
          value={shareCode}
          readOnly
          multiline
          rows={4}
        />
        <DialogFooter>
          <DefaultButton onClick={() => setShowShareDialog(false)} text="Close" />
          <PrimaryButton
            text="Copy to Clipboard"
            iconProps={{ iconName: "Copy" }}
            onClick={() => {
              navigator.clipboard.writeText(shareCode);
              setShowShareDialog(false);
            }}
          />
        </DialogFooter>
      </Dialog>

      {/* Import Dialog */}
      <Dialog
        hidden={!showImportDialog}
        onDismiss={() => setShowImportDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Import Recipe",
          subText: "Paste recipe JSON or share code:"
        }}
      >
        <TextField
          value={importData}
          onChange={(_, value) => setImportData(value || "")}
          multiline
          rows={6}
          placeholder="Paste JSON or share code here..."
        />
        <DialogFooter>
          <DefaultButton onClick={() => setShowImportDialog(false)} text="Cancel" />
          <PrimaryButton
            text="Import"
            iconProps={{ iconName: "Upload" }}
            onClick={handleImportRecipe}
            disabled={!importData.trim()}
          />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default RecipeGallery;
