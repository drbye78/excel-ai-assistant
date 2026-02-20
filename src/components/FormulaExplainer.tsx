import * as React from "react";
import { useState, useEffect } from "react";
import {
  explainFormula,
  FormulaExplanation,
  parseFormula
} from "@/services/formulaParser";
import {
  Stack,
  Text,
  DefaultButton,
  PrimaryButton,
  Separator,
  IconButton,
  TooltipHost,
  IStackTokens,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from "@fluentui/react";
import { tokens } from "@fluentui/react-theme";

interface FormulaExplainerProps {
  formula?: string;
  onClose?: () => void;
}

const stackTokens: IStackTokens = {
  childrenGap: 15
};

export const FormulaExplainer: React.FC<FormulaExplainerProps> = ({
  formula: initialFormula,
  onClose
}) => {
  const [formula, setFormula] = useState(initialFormula || "");
  const [explanation, setExplanation] = useState<FormulaExplanation | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedCell, setSelectedCell] = useState<string | null>(null);

  useEffect(() => {
    if (initialFormula) {
      analyzeFormula(initialFormula);
    }
  }, [initialFormula]);

  // Listen for selection changes in Excel
  useEffect(() => {
    const checkSelection = async () => {
      try {
        await Excel.run(async (context) => {
          const selection = context.workbook.getSelectedRange();
          selection.load("formulas, address");
          await context.sync();

          const selectedFormula = selection.formulas[0][0];
          if (selectedFormula && selectedFormula.startsWith("=")) {
            setSelectedCell(selection.address);
            setFormula(selectedFormula);
          }
        });
      } catch (e) {
        // Excel not available (development mode)
      }
    };

    // Check every 2 seconds for selection changes
    const interval = setInterval(checkSelection, 2000);
    return () => clearInterval(interval);
  }, []);

  const analyzeFormula = async (formulaToAnalyze: string) => {
    if (!formulaToAnalyze || !formulaToAnalyze.startsWith("=")) {
      setError("Please select a cell containing a formula");
      return;
    }

    setIsLoading(true);
    setError(null);

    try {
      const result = explainFormula(formulaToAnalyze);
      setExplanation(result);
    } catch (err) {
      setError(`Failed to analyze formula: ${err.message}`);
    } finally {
      setIsLoading(false);
    }
  };

  const getComplexityColor = (complexity: string): string => {
    switch (complexity) {
      case "Simple":
        return "#107c10"; // Green
      case "Moderate":
        return "#ffc107"; // Yellow
      case "Complex":
        return "#d83b01"; // Red
      default:
        return "#605e5c";
    }
  };

  const handleExplainSelected = async () => {
    try {
      await Excel.run(async (context) => {
        const selection = context.workbook.getSelectedRange();
        selection.load("formulas");
        await context.sync();

        const selectedFormula = selection.formulas[0][0];
        if (selectedFormula && selectedFormula.startsWith("=")) {
          analyzeFormula(selectedFormula);
        } else {
          setError("The selected cell does not contain a formula");
        }
      });
    } catch (e) {
      setError("Could not access Excel. Please ensure you're running in Excel.");
    }
  };

  return (
    <Stack tokens={stackTokens} styles={{ root: { padding: "20px", maxWidth: "600px" } }}>
      {/* Header */}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          🔍 Formula Explainer
        </Text>
        {onClose && (
          <IconButton
            iconProps={{ iconName: "Cancel" }}
            onClick={onClose}
            title="Close"
          />
        )}
      </Stack>

      <Separator />

      {/* Input Section */}
      <Stack tokens={{ childrenGap: 10 }}>
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
          Formula
        </Text>

        {selectedCell && (
          <Text variant="small" styles={{ root: { color: "#605e5c" } }}>
            Selected cell: {selectedCell}
          </Text>
        )}

        <div
          style={{
            backgroundColor: "#f3f2f1",
            padding: "12px",
            borderRadius: "4px",
            fontFamily: "Consolas, monospace",
            fontSize: "14px",
            wordBreak: "break-all"
          }}
        >
          {formula || "Select a cell with a formula..."}
        </div>

        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text="Explain Selected Cell"
            onClick={handleExplainSelected}
            disabled={isLoading}
          />
          {explanation && (
            <DefaultButton
              text="Re-analyze"
              onClick={() => analyzeFormula(formula)}
              disabled={isLoading}
            />
          )}
        </Stack>
      </Stack>

      {/* Error Message */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel="Close"
        >
          {error}
        </MessageBar>
      )}

      {/* Loading */}
      {isLoading && (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 10 }}>
          <Spinner size={SpinnerSize.medium} />
          <Text>Analyzing formula...</Text>
        </Stack>
      )}

      {/* Explanation Results */}
      {explanation && !isLoading && (
        <Stack tokens={{ childrenGap: 15 }}>
          <Separator />

          {/* Summary */}
          <Stack tokens={{ childrenGap: 5 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Summary
            </Text>
            <Text styles={{ root: { fontSize: "14px", lineHeight: "1.5" } }}>
              {explanation.summary}
            </Text>
          </Stack>

          {/* Complexity Badge */}
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }}>
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              Complexity:
            </Text>
            <div
              style={{
                backgroundColor: getComplexityColor(explanation.complexity),
                color: "white",
                padding: "4px 12px",
                borderRadius: "12px",
                fontSize: "12px",
                fontWeight: 600
              }}
            >
              {explanation.complexity}
            </div>
          </Stack>

          {/* Dependencies */}
          {explanation.dependencies.length > 0 && (
            <Stack tokens={{ childrenGap: 5 }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                Dependencies ({explanation.dependencies.length})
              </Text>
              <div style={{ display: "flex", flexWrap: "wrap", gap: "5px" }}>
                {explanation.dependencies.map((dep, idx) => (
                  <TooltipHost key={idx} content={`Click to navigate to ${dep}`}>
                    <DefaultButton
                      text={dep}
                      onClick={() => navigateToCell(dep)}
                      styles={{
                        root: {
                          fontSize: "11px",
                          height: "24px",
                          padding: "0 8px"
                        }
                      }}
                    />
                  </TooltipHost>
                ))}
              </div>
            </Stack>
          )}

          <Separator />

          {/* Step-by-Step Breakdown */}
          <Stack tokens={{ childrenGap: 10 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Step-by-Step Breakdown
            </Text>

            {explanation.breakdown.map((step, idx) => (
              <div
                key={idx}
                style={{
                  marginLeft: `${step.depth * 20}px`,
                  padding: "10px",
                  backgroundColor: step.depth === 0 ? "#e3f2fd" : "#f5f5f5",
                  borderLeft: `3px solid ${step.depth === 0 ? "#2196f3" : "#9e9e9e"}`,
                  borderRadius: "0 4px 4px 0"
                }}
              >
                <Text
                  variant="small"
                  styles={{
                    root: {
                      fontFamily: "Consolas, monospace",
                      fontWeight: 600,
                      color: "#1565c0",
                      marginBottom: "4px"
                    }
                  }}
                >
                  {step.expression}
                </Text>
                <Text variant="small" styles={{ root: { color: "#424242" } }}>
                  {step.description}
                </Text>
              </div>
            ))}
          </Stack>

          {/* Optimizations */}
          {explanation.optimizations && explanation.optimizations.length > 0 && (
            <>
              <Separator />
              <Stack tokens={{ childrenGap: 10 }}>
                <Text
                  variant="mediumPlus"
                  styles={{ root: { fontWeight: 600, color: "#d83b01" } }}
                >
                  💡 Optimization Suggestions
                </Text>
                {explanation.optimizations.map((opt, idx) => (
                  <MessageBar
                    key={idx}
                    messageBarType={MessageBarType.warning}
                    styles={{ root: { fontSize: "12px" } }}
                  >
                    {opt}
                  </MessageBar>
                ))}
              </Stack>
            </>
          )}
        </Stack>
      )}
    </Stack>
  );
};

// Helper to navigate to a cell in Excel
async function navigateToCell(address: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("worksheet");
      await context.sync();

      const worksheet = range.worksheet;
      const targetRange = worksheet.getRange(address);
      targetRange.select();
      await context.sync();
    });
  } catch (e) {
    console.error("Failed to navigate to cell:", e);
  }
}

export default FormulaExplainer;
