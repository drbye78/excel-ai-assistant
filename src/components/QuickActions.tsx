// Quick Actions Panel Component
// Pre-built prompts for common Excel tasks

import * as React from "react";
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  SearchBox,
  Pivot,
  PivotItem,
  ActionButton,
  Separator
} from "@fluentui/react";
import { lightTheme } from "@/theme";

const { colors, spacing } = lightTheme;

export interface QuickAction {
  id: string;
  title: string;
  description: string;
  prompt: string;
  icon?: string;
  category: 'formulas' | 'formatting' | 'data' | 'analysis' | 'charts';
}

export const QUICK_ACTIONS: QuickAction[] = [
  // Formulas
  {
    id: "sum-range",
    title: "Sum Range",
    description: "Calculate sum of selected cells",
    prompt: "Write a formula to sum the selected range",
    category: "formulas"
  },
  {
    id: "average",
    title: "Average",
    description: "Calculate average of selected cells",
    prompt: "Write a formula to calculate the average of selected cells",
    category: "formulas"
  },
  {
    id: "vlookup",
    title: "VLOOKUP",
    description: "Look up a value in a table",
    prompt: "Write a VLOOKUP formula to find a value",
    category: "formulas"
  },
  {
    id: "if-formula",
    title: "IF Statement",
    description: "Conditional logic formula",
    prompt: "Write an IF formula for conditional calculation",
    category: "formulas"
  },
  // Formatting
  {
    id: "conditional-format",
    title: "Conditional Formatting",
    description: "Highlight cells based on rules",
    prompt: "Add conditional formatting to highlight cells",
    category: "formatting"
  },
  {
    id: "number-format",
    title: "Number Format",
    description: "Format numbers as currency/percentage",
    prompt: "Format the selected range as currency",
    category: "formatting"
  },
  {
    id: "table-style",
    title: "Table Style",
    description: "Apply a table style",
    prompt: "Convert the range to a formatted table",
    category: "formatting"
  },
  // Data
  {
    id: "remove-duplicates",
    title: "Remove Duplicates",
    description: "Remove duplicate rows",
    prompt: "Remove duplicate rows from the selected data",
    category: "data"
  },
  {
    id: "sort-range",
    title: "Sort Data",
    description: "Sort by column",
    prompt: "Sort the selected data by a column",
    category: "data"
  },
  {
    id: "filter-data",
    title: "Filter Data",
    description: "Add filters to columns",
    prompt: "Add auto-filter to the selected range",
    category: "data"
  },
  // Analysis
  {
    id: "pivot-table",
    title: "Create Pivot Table",
    description: "Generate pivot table analysis",
    prompt: "Create a pivot table to summarize the data",
    category: "analysis"
  },
  {
    id: "statistics",
    title: "Statistical Analysis",
    description: "Calculate stats (mean, median, std dev)",
    prompt: "Calculate statistical analysis of the selected data",
    category: "analysis"
  },
  {
    id: "trend-analysis",
    title: "Trend Analysis",
    description: "Analyze data trends",
    prompt: "Analyze trends in the selected data",
    category: "analysis"
  },
  // Charts
  {
    id: "bar-chart",
    title: "Bar Chart",
    description: "Create a bar chart",
    prompt: "Create a bar chart from the selected data",
    category: "charts"
  },
  {
    id: "line-chart",
    title: "Line Chart",
    description: "Create a line chart",
    prompt: "Create a line chart showing trends",
    category: "charts"
  },
  {
    id: "pie-chart",
    title: "Pie Chart",
    description: "Create a pie chart",
    prompt: "Create a pie chart to show distribution",
    category: "charts"
  }
];

interface QuickActionsProps {
  isOpen: boolean;
  onClose: () => void;
  onSelectAction: (prompt: string) => void;
}

export const QuickActions: React.FC<QuickActionsProps> = ({
  isOpen,
  onClose,
  onSelectAction
}) => {
  const [searchQuery, setSearchQuery] = React.useState("");
  const [selectedCategory, setSelectedCategory] = React.useState<string>("all");

  const handleCategoryChange = (key: string | undefined) => {
    if (key) {
      setSelectedCategory(key);
    }
  };

  const filteredActions = QUICK_ACTIONS.filter(action => {
    const matchesSearch = action.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
                         action.description.toLowerCase().includes(searchQuery.toLowerCase());
    const matchesCategory = selectedCategory === "all" || action.category === selectedCategory;
    return matchesSearch && matchesCategory;
  });

  const handleActionClick = (action: QuickAction) => {
    onSelectAction(action.prompt);
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(0, 0, 0, 0.4)',
        display: 'flex',
        alignItems: 'flex-start',
        justifyContent: 'center',
        paddingTop: '100px',
        zIndex: 1000
      }}
      onClick={onClose}
      role="dialog"
      aria-modal="true"
      aria-label="Quick Actions"
    >
      <div
        style={{
          backgroundColor: colors.background.card,
          borderRadius: '8px',
          boxShadow: '0 8px 32px rgba(0, 0, 0, 0.2)',
          width: '600px',
          maxHeight: '70vh',
          overflow: 'hidden',
          display: 'flex',
          flexDirection: 'column'
        }}
        onClick={(e) => e.stopPropagation()}
        role="document"
      >
        {/* Header */}
        <Stack
          horizontalAlign="space-between"
          verticalAlign="center"
          styles={{
            root: {
              padding: spacing.md,
              borderBottom: `1px solid ${colors.border.default}`
            }
          }}
        >
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
            ⚡ Quick Actions
          </Text>
          <DefaultButton
            text="✕"
            onClick={onClose}
            aria-label="Close quick actions"
            styles={{ root: { minWidth: 32, padding: '0 8px' } }}
          />
        </Stack>

        {/* Search */}
        <div style={{ padding: spacing.md, borderBottom: `1px solid ${colors.border.default}` }}>
          <SearchBox
            placeholder="Search actions..."
            value={searchQuery}
            onChange={(_, newValue) => setSearchQuery(newValue || "")}
            styles={{ root: { width: '100%' } }}
          />
        </div>

        {/* Category Tabs */}
        <Stack horizontal tokens={{ childrenGap: spacing.sm }} styles={{ root: { padding: `${spacing.sm} ${spacing.md}`, flexWrap: 'wrap' } }}>
          {['all', 'formulas', 'formatting', 'data', 'analysis', 'charts'].map(category => (
            <DefaultButton
              key={category}
              text={category.charAt(0).toUpperCase() + category.slice(1)}
              onClick={() => setSelectedCategory(category)}
              aria-pressed={selectedCategory === category}
              styles={{
                root: selectedCategory === category ? {
                  backgroundColor: colors.brand.primary,
                  color: colors.text.onBrand
                } : {}
              }}
            />
          ))}
        </Stack>

        {/* Actions List */}
        <div style={{ 
          overflowY: 'auto', 
          flex: 1,
          padding: spacing.md 
        }}>
          <Stack tokens={{ childrenGap: spacing.sm }}>
            {filteredActions.map(action => (
              <ActionButton
                key={action.id}
                onClick={() => handleActionClick(action)}
                styles={{
                  root: {
                    textAlign: 'left',
                    padding: spacing.md,
                    height: 'auto',
                    justifyContent: 'flex-start'
                  }
                }}
              >
                <Stack tokens={{ childrenGap: spacing.xs }}>
                  <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                    {action.title}
                  </Text>
                  <Text variant="small" styles={{ root: { color: colors.text.secondary } }}>
                    {action.description}
                  </Text>
                </Stack>
              </ActionButton>
            ))}
            
            {filteredActions.length === 0 && (
              <Stack horizontalAlign="center" styles={{ root: { padding: spacing.xl } }}>
                <Text>No actions found matching your search.</Text>
              </Stack>
            )}
          </Stack>
        </div>

        {/* Footer */}
        <Stack
          horizontal
          horizontalAlign="center"
          verticalAlign="center"
          styles={{
            root: {
              padding: spacing.md,
              borderTop: `1px solid ${colors.border.default}`,
              backgroundColor: colors.background.page
            }
          }}
        >
          <Text variant="small" styles={{ root: { color: colors.text.secondary } }}>
            Press <kbd style={{ 
              background: colors.neutral.gray20, 
              padding: '2px 6px', 
              borderRadius: '4px',
              fontFamily: 'monospace'
            }}>Ctrl+K</kbd> to open anytime
          </Text>
        </Stack>
      </div>
    </div>
  );
};

export default QuickActions;
