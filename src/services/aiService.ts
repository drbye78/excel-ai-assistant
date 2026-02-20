// AI Service - Handles communication with OpenAI-compatible APIs
import { AIRequest, AIResponse, AISettings, AIAction, ExcelContext } from "@/types";
import OpenRouterService, { OpenRouterModel } from "./openRouterService";
import SettingsService from "./settingsService";
import logger from "@/utils/logger";

export class AIService {
  private static instance: AIService;
  private settings: AISettings;
  private isSettingsInitialized: boolean = false;

  private constructor() {
    // Default settings - will be overridden by persisted configuration
    this.settings = {
      apiUrl: "https://openrouter.ai/api/v1",
      apiKey: "",
      model: "openai/gpt-4",
      temperature: 0.7,
      maxTokens: 4000
    };
  }

  static getInstance(): AIService {
    if (!AIService.instance) {
      AIService.instance = new AIService();
    }
    return AIService.instance;
  }

  /**
   * Initialize and load persisted settings
   */
  async initialize(): Promise<void> {
    try {
      await SettingsService.initialize();
      const persistedSettings = SettingsService.getSettings();
      this.settings = { ...this.settings, ...persistedSettings };
      this.isSettingsInitialized = true;
    } catch (error) {
      logger.warn('Failed to initialize settings, using defaults', undefined, error instanceof Error ? error : new Error(String(error)));
      this.isSettingsInitialized = true;
    }
  }

  /**
   * Check if settings are initialized
   */
  isInitialized(): boolean {
    return this.isSettingsInitialized;
  }

  updateSettings(settings: Partial<AISettings>, persist: boolean = true, scope: 'global' | 'workbook' | 'auto' = 'auto'): Promise<void> {
    // Update local state immediately
    this.settings = { ...this.settings, ...settings };
    
    // Persist if requested
    if (persist) {
      return SettingsService.updateSettings(settings, scope);
    }
    return Promise.resolve();
  }

  getSettings(): AISettings {
    return { ...this.settings };
  }

  /**
   * Get current settings scope info
   */
  getSettingsScope(): { type: string; source: string } {
    return SettingsService.getScope();
  }

  /**
   * Get available models from OpenRouter
   * Returns models based on current API configuration
   */
  async getAvailableModels(): Promise<OpenRouterModel[]> {
    const settings = this.settings;
    
    // Check if using OpenRouter
    if (!OpenRouterService.isOpenRouterUrl(settings.apiUrl)) {
      return [];
    }

    return OpenRouterService.getAvailableModels(settings.apiUrl, settings.apiKey);
  }

  /**
   * Get default models (fallback when API unavailable)
   */
  getDefaultModels(): OpenRouterModel[] {
    return OpenRouterService.getDefaultModels();
  }

  /**
   * Force refresh models from OpenRouter API
   */
  async refreshModels(): Promise<OpenRouterModel[]> {
    const settings = this.settings;
    
    if (!OpenRouterService.isOpenRouterUrl(settings.apiUrl)) {
      return [];
    }

    return OpenRouterService.refreshModels(settings.apiUrl, settings.apiKey);
  }

  /**
   * Check if the API URL is for OpenRouter
   */
  private isOpenRouter(url: string): boolean {
    return url.includes('openrouter.ai');
  }

  /**
   * Get headers for the API request, including OpenRouter-specific headers
   */
  private getHeaders(activeSettings: AISettings): HeadersInit {
    const headers: HeadersInit = {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${activeSettings.apiKey}`
    };

    // Add OpenRouter-specific required headers
    if (this.isOpenRouter(activeSettings.apiUrl)) {
      // Use window location as the app URL, with fallback
      // process.env may not be available in browser environment
      const APP_URL = typeof window !== 'undefined' 
        ? window.location.origin 
        : "https://excel-ai-assistant.com";
      const APP_NAME = "Excel AI Assistant";
      
      (headers as Record<string, string>)["HTTP-Referer"] = APP_URL;
      (headers as Record<string, string>)["X-Title"] = APP_NAME;
    }

    return headers;
  }

  /**
   * Transform model name for OpenRouter if needed
   */
  private getModelForRequest(activeSettings: AISettings): string {
    // If using OpenRouter, prefix with provider if not already prefixed
    if (this.isOpenRouter(activeSettings.apiUrl)) {
      const model = activeSettings.model;
      // Check if model already has provider prefix (e.g., "openai/gpt-4")
      if (!model.includes('/')) {
        // Map common models to OpenRouter format
        const modelMap: Record<string, string> = {
          'gpt-4': 'openai/gpt-4',
          'gpt-4-turbo': 'openai/gpt-4-turbo',
          'gpt-3.5-turbo': 'openai/gpt-3.5-turbo',
          'gpt-3.5-turbo-16k': 'openai/gpt-3.5-turbo-16k',
          'claude-3-opus': 'anthropic/claude-3-opus-20240229',
          'claude-3-sonnet': 'anthropic/claude-3-sonnet-20240229',
          'claude-3-haiku': 'anthropic/claude-3-haiku-20240307',
          'llama-3-70b': 'meta-llama/llama-3-70b-instruct',
          'llama-3-8b': 'meta-llama/llama-3-8b-instruct',
          'mistral-7b': 'mistralai/mistral-7b-instruct',
          'mixtral-8x7b': 'mistralai/mixtral-8x7b-instruct'
        };
        return modelMap[model.toLowerCase()] || `openai/${model}`;
      }
    }
    return activeSettings.model;
  }

  async sendMessage(request: AIRequest): Promise<AIResponse> {
    const { message, context, conversationHistory, settings } = request;

    // Use provided settings or fall back to instance settings
    const activeSettings = settings || this.settings;

    // Validate settings
    if (!activeSettings.apiKey) {
      throw new Error("API key is required. Please configure your settings.");
    }
    if (!activeSettings.apiUrl) {
      throw new Error("API URL is required. Please configure your settings.");
    }

    // Build system prompt with Excel context
    const systemPrompt = this.buildSystemPrompt(context);

    // Build messages array
    const messages = [
      { role: "system", content: systemPrompt },
      ...conversationHistory.map((msg) => ({
        role: msg.role,
        content: msg.content
      })),
      { role: "user", content: this.buildUserPrompt(message, context) }
    ];

    try {
      // Get model (transformed for OpenRouter if needed)
      const model = this.getModelForRequest(activeSettings);

      // Get headers (with OpenRouter-specific headers if needed)
      const headers = this.getHeaders(activeSettings);

      const response = await fetch(`${activeSettings.apiUrl}/chat/completions`, {
        method: "POST",
        headers,
        body: JSON.stringify({
          model,
          messages,
          temperature: activeSettings.temperature,
          max_tokens: activeSettings.maxTokens,
          tools: this.getAvailableTools(),
          tool_choice: "auto"
        })
      });

      if (!response.ok) {
        let errorMessage = `API Error: ${response.status}`;
        try {
          const errorData = await response.json();
          // OpenRouter error format
          if (errorData.error) {
            errorMessage = errorData.error.message || errorData.error;
          } else if (errorData.error?.message) {
            errorMessage = errorData.error.message;
          }
        } catch {
          // Response wasn't JSON, use status text
          errorMessage = `${response.status}: ${response.statusText}`;
        }
        throw new Error(errorMessage);
      }

      const data = await response.json();
      
      // Check for valid response structure
      if (!data.choices || !data.choices[0]) {
        throw new Error("Invalid response from AI API");
      }

      const assistantMessage = data.choices[0].message;

      // Parse tool calls if present
      const actions = this.parseToolCalls(assistantMessage.tool_calls);

      return {
        message: assistantMessage.content || this.formatFunctionResult(actions),
        actions,
        suggestedPrompts: this.generateSuggestedPrompts(context),
        requiresConfirmation: actions.some((a) => this.requiresConfirmation(a.type))
      };
    } catch (error) {
      // Re-throw with more context
      if (error instanceof Error) {
        logger.error("AI Service Error", undefined, error);
        throw error;
      }
      logger.error("AI Service Error", undefined, new Error(String(error)));
      throw new Error(String(error));
    }
  }

  private buildSystemPrompt(context: ExcelContext): string {
    return `You are an expert Excel AI assistant. You help users work with Excel workbooks through natural language.

CURRENT WORKBOOK STATE:
- Workbook has ${context.workbook.worksheets.length} worksheets: ${context.workbook.worksheets.join(", ")}
- Active worksheet: ${context.activeWorksheet.name}
- Tables: ${context.tables.map((t) => t.name).join(", ") || "None"}
- Charts: ${context.charts.map((c) => c.name).join(", ") || "None"}
- Pivot Tables: ${context.pivotTables.map((p) => p.name).join(", ") || "None"}
${context.selection ? `- Current selection: ${context.selection.address} (${context.selection.rowCount} rows x ${context.selection.columnCount} columns)` : ""}

AVAILABLE CAPABILITIES:
1. Cell Operations: Read/write values, formulas, formatting
2. Range Operations: Copy, clear, format, auto-fit
3. Table Operations: Create, modify, add rows, delete
4. Chart Operations: Create charts (column, line, pie, bar, scatter, etc.)
5. Pivot Tables: Create with row/column/data/filter fields
6. Data Validation: Add validation rules, dropdown lists
7. Named Ranges: Create and manage named ranges
8. Worksheet Management: Add, delete, rename worksheets
9. VBA Macros: Create, explain, and refactor VBA code

VBA MACRO CAPABILITIES:
- Generate VBA code from natural language descriptions
- Explain existing VBA code line by line
- Refactor/optimize VBA code for better performance
- Create macros for automation tasks (data processing, formatting, etc.)

When providing VBA code, use proper VBA syntax and wrap in triple backticks with 'vba' language tag:
\`\`\`vba
Sub MyMacro()
    ' Your code here
End Sub
\`\`\`

When providing formulas, wrap them in triple backticks like:
\`\`\`=SUM(A1:A10)\`\`\`

When suggesting actions, be specific about cell addresses and worksheet names.

Always confirm destructive operations (delete, clear) before executing.`;
  }

  private buildUserPrompt(message: string, context: ExcelContext): string {
    let prompt = message;

    if (context.selection) {
      prompt += `\n\n[Context: Currently selected range is ${context.selection.address} with data: ${JSON.stringify(context.selection.values.slice(0, 5))}]`;
    }

    return prompt;
  }

  private getAvailableTools(): any[] {
    return [
      {
        type: "function",
        function: {
          name: "insert_formula",
          description: "Insert an Excel formula into a cell or range",
          parameters: {
            type: "object",
            properties: {
              formula: { type: "string", description: "The Excel formula to insert" },
              address: { type: "string", description: "Cell address (e.g., 'A1', 'B2:D10')" },
              worksheetName: { type: "string", description: "Worksheet name (optional, defaults to active)" }
            },
            required: ["formula", "address"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "set_values",
          description: "Set values in a range",
          parameters: {
            type: "object",
            properties: {
              values: { type: "array", description: "2D array of values" },
              address: { type: "string", description: "Cell address" },
              worksheetName: { type: "string" }
            },
            required: ["values", "address"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "create_table",
          description: "Create an Excel table",
          parameters: {
            type: "object",
            properties: {
              range: { type: "string", description: "Range address for the table" },
              name: { type: "string", description: "Table name" },
              hasHeaders: { type: "boolean" },
              style: { type: "string", description: "Table style (e.g., 'TableStyleMedium2')" },
              worksheetName: { type: "string" }
            },
            required: ["range", "hasHeaders"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "create_chart",
          description: "Create a chart",
          parameters: {
            type: "object",
            properties: {
              dataRange: { type: "string", description: "Data range for the chart" },
              type: { type: "string", enum: ["columnClustered", "line", "pie", "barClustered", "scatter"] },
              title: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["dataRange", "type"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "create_pivot_table",
          description: "Create a pivot table",
          parameters: {
            type: "object",
            properties: {
              name: { type: "string" },
              sourceData: { type: "string", description: "Source data range" },
              destination: { type: "string", description: "Destination cell" },
              rowFields: { type: "array", items: { type: "string" } },
              columnFields: { type: "array", items: { type: "string" } },
              dataFields: { type: "array", items: { type: "object" } },
              worksheetName: { type: "string" }
            },
            required: ["name", "sourceData"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "format_cells",
          description: "Format cells",
          parameters: {
            type: "object",
            properties: {
              address: { type: "string" },
              numberFormat: { type: "string" },
              bold: { type: "boolean" },
              italic: { type: "boolean" },
              color: { type: "string" },
              fillColor: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["address"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "add_validation",
          description: "Add data validation",
          parameters: {
            type: "object",
            properties: {
              address: { type: "string" },
              type: { type: "string", enum: ["list", "whole", "decimal", "date", "textLength", "custom"] },
              formula1: { type: "string" },
              formula2: { type: "string" },
              allowBlank: { type: "boolean" },
              inputMessage: { type: "string" },
              errorMessage: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["address", "type"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "clear_range",
          description: "Clear a range",
          parameters: {
            type: "object",
            properties: {
              address: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["address"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "add_worksheet",
          description: "Add a new worksheet",
          parameters: {
            type: "object",
            properties: {
              name: { type: "string" }
            }
          }
        }
      },
      {
        type: "function",
        function: {
          name: "delete_worksheet",
          description: "Delete a worksheet",
          parameters: {
            type: "object",
            properties: {
              name: { type: "string" }
            },
            required: ["name"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "auto_fit_columns",
          description: "Auto-fit column widths",
          parameters: {
            type: "object",
            properties: {
              address: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["address"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "create_named_range",
          description: "Create a named range",
          parameters: {
            type: "object",
            properties: {
              name: { type: "string" },
              address: { type: "string" },
              worksheetName: { type: "string" }
            },
            required: ["name", "address"]
          }
        }
      },
      // VBA Macro Tools
      {
        type: "function",
        function: {
          name: "create_vba_macro",
          description: "Create a VBA macro from natural language description. Generates VBA code for automation tasks like data processing, formatting, report generation, etc.",
          parameters: {
            type: "object",
            properties: {
              macroName: { type: "string", description: "Name for the VBA macro" },
              description: { type: "string", description: "Natural language description of what the macro should do" },
              worksheetName: { type: "string", description: "Optional worksheet name to work with" }
            },
            required: ["macroName", "description"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "explain_vba_code",
          description: "Explain existing VBA code line by line with detailed comments",
          parameters: {
            type: "object",
            properties: {
              vbaCode: { type: "string", description: "The VBA code to explain" },
              detailLevel: { type: "string", enum: ["brief", "detailed", "comprehensive"], description: "Level of detail for the explanation" }
            },
            required: ["vbaCode"]
          }
        }
      },
      {
        type: "function",
        function: {
          name: "refactor_vba_code",
          description: "Refactor and optimize VBA code for better performance, readability, and best practices",
          parameters: {
            type: "object",
            properties: {
              vbaCode: { type: "string", description: "The VBA code to refactor" },
              goals: { type: "array", items: { type: "string" }, description: "Optimization goals: performance, readability, error_handling, or modern_vba" }
            },
            required: ["vbaCode"]
          }
        }
      }
    ];
  }

  private parseToolCalls(toolCalls: any[]): AIAction[] {
    if (!toolCalls || !Array.isArray(toolCalls) || toolCalls.length === 0) return [];

    const actions: AIAction[] = [];

    for (const toolCall of toolCalls) {
      if (toolCall.type === "function" && toolCall.function) {
        const func = toolCall.function;
        try {
          const args = JSON.parse(func.arguments);
          actions.push({
            type: func.name as any,
            label: this.getActionLabel(func.name),
            payload: args
          });
        } catch (e) {
          logger.error("Failed to parse tool arguments", undefined, e instanceof Error ? e : new Error(String(e)));
        }
      }
    }

    return actions;
  }

  private getActionLabel(functionName: string): string {
    const labels: Record<string, string> = {
      insert_formula: "Insert Formula",
      set_values: "Set Values",
      create_table: "Create Table",
      create_chart: "Create Chart",
      create_pivot_table: "Create Pivot Table",
      format_cells: "Format Cells",
      add_validation: "Add Validation",
      clear_range: "Clear Range",
      add_worksheet: "Add Worksheet",
      delete_worksheet: "Delete Worksheet",
      auto_fit_columns: "Auto-fit Columns",
      create_named_range: "Create Named Range",
      // VBA Macro Labels
      create_vba_macro: "Create VBA Macro",
      explain_vba_code: "Explain VBA Code",
      refactor_vba_code: "Refactor VBA Code"
    };
    return labels[functionName] || functionName;
  }

  private formatFunctionResult(actions: AIAction[]): string {
    if (actions.length === 0) return "";

    return `I'll perform the following action${actions.length > 1 ? "s" : ""}:\n${actions
      .map((a) => `- ${a.label}`)
      .join("\n")}`;
  }

  private requiresConfirmation(actionType: string): boolean {
    const destructiveActions = ["delete_worksheet", "clear_range", "delete_table"];
    return destructiveActions.includes(actionType);
  }

  private generateSuggestedPrompts(context: ExcelContext): string[] {
    const prompts = [
      "Create a summary table from the selected data",
      "Generate a chart showing trends",
      "Add data validation to prevent invalid entries",
      "Format this range as currency",
      "Create a pivot table for analysis",
      // VBA-related prompts
      "Create a macro to format all sheets",
      "Explain this VBA code",
      "Refactor this macro for better performance"
    ];

    if (context.tables.length > 0) {
      prompts.push(`Add a new row to table "${context.tables[0].name}"`);
    }

    return prompts.slice(0, 3);
  }
}

export default AIService.getInstance();
