/**
 * Local AI Service
 *
 * Provides privacy-focused AI capabilities using local models via Ollama
 * or Web Workers for on-device processing. No data leaves the user's machine.
 *
 * Features:
 * - Ollama integration for local LLMs
 * - Web Worker processing for non-blocking UI
 * - Model management and switching
 * - Offline capability
 * - Privacy-preserving data analysis
 *
 * @module services/localAIService
 */

import { notificationManager } from '../utils/notificationManager';

// ============================================================================
// Type Definitions
// ============================================================================

/** Local AI provider types */
export type LocalAIProvider = 'ollama' | 'webworker' | 'mock';

/** Local model information */
export interface LocalModel {
  id: string;
  name: string;
  provider: LocalAIProvider;
  size: string;
  description: string;
  capabilities: string[];
  isDownloaded: boolean;
  downloadProgress?: number;
}

/** AI request for local processing */
export interface LocalAIRequest {
  prompt: string;
  systemPrompt?: string;
  context?: string[];
  maxTokens?: number;
  temperature?: number;
  model?: string;
}

/** AI response from local model */
export interface LocalAIResponse {
  text: string;
  tokensUsed: number;
  duration: number;
  model: string;
}

/** Service configuration */
export interface LocalAIConfig {
  provider: LocalAIProvider;
  ollamaUrl: string;
  defaultModel: string;
  maxTokens: number;
  temperature: number;
}

/** Processing status */
export interface ProcessingStatus {
  isProcessing: boolean;
  progress: number;
  message: string;
}

// ============================================================================
// Local AI Service
// ============================================================================

export class LocalAIService {
  private static instance: LocalAIService;
  private config: LocalAIConfig;
  private worker: Worker | null = null;
  private workerBlobUrl: string | null = null;
  private isInitialized: boolean = false;
  private processingCallbacks: Map<string, (response: LocalAIResponse) => void> = new Map();

  private constructor() {
    this.config = {
      provider: 'ollama',
      ollamaUrl: 'http://localhost:11434',
      defaultModel: 'llama2',
      maxTokens: 2048,
      temperature: 0.7,
    };
  }

  static getInstance(): LocalAIService {
    if (!LocalAIService.instance) {
      LocalAIService.instance = new LocalAIService();
    }
    return LocalAIService.instance;
  }

  // ============================================================================
  // Initialization
  // ============================================================================

  /**
   * Initialize the service
   */
  async initialize(config?: Partial<LocalAIConfig>): Promise<boolean> {
    if (config) {
      this.config = { ...this.config, ...config };
    }

    try {
      switch (this.config.provider) {
        case 'ollama':
          return await this.initializeOllama();
        case 'webworker':
          return this.initializeWebWorker();
        case 'mock':
          this.isInitialized = true;
          return true;
        default:
          return false;
      }
    } catch (error) {
      notificationManager.error('Failed to initialize Local AI: ' + error);
      return false;
    }
  }

  private async initializeOllama(): Promise<boolean> {
    try {
      const response = await fetch(`${this.config.ollamaUrl}/api/tags`);
      if (response.ok) {
        this.isInitialized = true;
        notificationManager.success('Connected to Ollama');
        return true;
      }
    } catch {
      notificationManager.warning('Ollama not available at ' + this.config.ollamaUrl);
    }
    return false;
  }

  private initializeWebWorker(): boolean {
    try {
      // Create inline worker for simple tasks
      const workerScript = `
        self.onmessage = function(e) {
          const { id, prompt, type } = e.data;
          
          // Simple pattern matching for Excel-related queries
          let response = '';
          
          if (type === 'formula') {
            response = generateFormulaExplanation(prompt);
          } else if (type === 'analysis') {
            response = generateAnalysis(prompt);
          } else {
            response = 'I can help with Excel formulas and data analysis. Please specify what you need.';
          }
          
          self.postMessage({
            id,
            text: response,
            tokensUsed: response.length / 4,
            duration: 0,
            model: 'webworker-local'
          });
        };
        
        function generateFormulaExplanation(prompt) {
          // Simple rule-based formula explanation
          if (prompt.includes('SUM')) {
            return 'The SUM function adds all the numbers in a range of cells.';
          } else if (prompt.includes('VLOOKUP')) {
            return 'VLOOKUP searches for a value in the first column of a table and returns a value in the same row.';
          } else if (prompt.includes('IF')) {
            return 'The IF function makes logical comparisons between values.';
          }
          return 'I can explain Excel formulas. Which specific formula would you like help with?';
        }
        
        function generateAnalysis(prompt) {
          return 'For detailed data analysis, consider using the built-in Data Analysis tools in the AI Assistant.';
        }
      `;

      // Clean up existing worker if any
      this.disposeWorker();

      const blob = new Blob([workerScript], { type: 'application/javascript' });
      this.workerBlobUrl = URL.createObjectURL(blob);
      this.worker = new Worker(this.workerBlobUrl);
      
      this.worker.onmessage = (e) => {
        const { id, ...response } = e.data;
        const callback = this.processingCallbacks.get(id);
        if (callback) {
          callback(response);
          this.processingCallbacks.delete(id);
        }
      };

      this.isInitialized = true;
      return true;
    } catch (error) {
      notificationManager.error('Failed to initialize Web Worker: ' + error);
      return false;
    }
  }

  /**
   * Clean up worker resources to prevent memory leaks
   */
  private disposeWorker(): void {
    if (this.worker) {
      this.worker.terminate();
      this.worker = null;
    }
    if (this.workerBlobUrl) {
      URL.revokeObjectURL(this.workerBlobUrl);
      this.workerBlobUrl = null;
    }
    this.processingCallbacks.clear();
  }

  // ============================================================================
  // Model Management
  // ============================================================================

  /**
   * Get available models
   */
  async getAvailableModels(): Promise<LocalModel[]> {
    const models: LocalModel[] = [];

    if (this.config.provider === 'ollama') {
      try {
        const response = await fetch(`${this.config.ollamaUrl}/api/tags`);
        if (response.ok) {
          const data = await response.json();
          for (const model of data.models || []) {
            models.push({
              id: model.name,
              name: model.name,
              provider: 'ollama',
              size: this.formatSize(model.size),
              description: `Ollama model: ${model.name}`,
              capabilities: ['text-generation', 'code'],
              isDownloaded: true,
            });
          }
        }
      } catch {
        // Ollama not available
      }
    }

    // Add mock models for testing
    models.push({
      id: 'mock-excel-assistant',
      name: 'Excel Assistant (Offline)',
      provider: 'mock',
      size: '0 MB',
      description: 'Rule-based assistant for offline use',
      capabilities: ['formula-explanation', 'basic-analysis'],
      isDownloaded: true,
    });

    return models;
  }

  private formatSize(bytes: number): string {
    const gb = bytes / (1024 * 1024 * 1024);
    return gb >= 1 ? `${gb.toFixed(1)} GB` : `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  /**
   * Download a model
   */
  async downloadModel(modelId: string): Promise<boolean> {
    if (this.config.provider !== 'ollama') {
      notificationManager.warning('Model download only available with Ollama');
      return false;
    }

    try {
      notificationManager.info(`Downloading model: ${modelId}...`);
      
      const response = await fetch(`${this.config.ollamaUrl}/api/pull`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: modelId }),
      });

      if (response.ok) {
        notificationManager.success(`Model ${modelId} downloaded successfully`);
        return true;
      }
    } catch (error) {
      notificationManager.error('Failed to download model: ' + error);
    }
    return false;
  }

  // ============================================================================
  // AI Processing
  // ============================================================================

  /**
   * Process a prompt with local AI
   */
  async process(request: LocalAIRequest): Promise<LocalAIResponse> {
    if (!this.isInitialized) {
      await this.initialize();
    }

    const startTime = Date.now();

    switch (this.config.provider) {
      case 'ollama':
        return this.processWithOllama(request, startTime);
      case 'webworker':
        return this.processWithWebWorker(request, startTime);
      case 'mock':
        return this.processWithMock(request, startTime);
      default:
        throw new Error('Unknown provider: ' + this.config.provider);
    }
  }

  private async processWithOllama(
    request: LocalAIRequest,
    startTime: number
  ): Promise<LocalAIResponse> {
    const response = await fetch(`${this.config.ollamaUrl}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: request.model || this.config.defaultModel,
        prompt: this.buildPrompt(request),
        stream: false,
        options: {
          temperature: request.temperature || this.config.temperature,
          num_predict: request.maxTokens || this.config.maxTokens,
        },
      }),
    });

    if (!response.ok) {
      throw new Error(`Ollama request failed: ${response.statusText}`);
    }

    const data = await response.json();
    
    return {
      text: data.response,
      tokensUsed: data.eval_count || 0,
      duration: Date.now() - startTime,
      model: request.model || this.config.defaultModel,
    };
  }

  private processWithWebWorker(
    request: LocalAIRequest,
    startTime: number
  ): Promise<LocalAIResponse> {
    return new Promise((resolve, reject) => {
      if (!this.worker) {
        reject(new Error('Web Worker not initialized'));
        return;
      }

      const id = `req-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
      
      this.processingCallbacks.set(id, (response) => {
        resolve({
          ...response,
          duration: Date.now() - startTime,
        });
      });

      // Set timeout
      setTimeout(() => {
        if (this.processingCallbacks.has(id)) {
          this.processingCallbacks.delete(id);
          reject(new Error('Request timeout'));
        }
      }, 30000);

      this.worker.postMessage({
        id,
        prompt: request.prompt,
        type: this.detectRequestType(request.prompt),
      });
    });
  }

  private processWithMock(
    request: LocalAIRequest,
    startTime: number
  ): Promise<LocalAIResponse> {
    return new Promise((resolve) => {
      setTimeout(() => {
        resolve({
          text: this.generateMockResponse(request.prompt),
          tokensUsed: request.prompt.length / 4,
          duration: Date.now() - startTime,
          model: 'mock-local',
        });
      }, 500);
    });
  }

  private buildPrompt(request: LocalAIRequest): string {
    let prompt = '';
    
    if (request.systemPrompt) {
      prompt += `System: ${request.systemPrompt}\n\n`;
    }
    
    if (request.context && request.context.length > 0) {
      prompt += `Context:\n${request.context.join('\n')}\n\n`;
    }
    
    prompt += `User: ${request.prompt}\n\nAssistant:`;
    
    return prompt;
  }

  private detectRequestType(prompt: string): string {
    const lower = prompt.toLowerCase();
    if (lower.includes('formula') || lower.includes('function') || lower.includes('=')) {
      return 'formula';
    } else if (lower.includes('analyze') || lower.includes('data') || lower.includes('statistics')) {
      return 'analysis';
    }
    return 'general';
  }

  private generateMockResponse(prompt: string): string {
    const lower = prompt.toLowerCase();
    
    if (lower.includes('formula') || lower.includes('sum') || lower.includes('vlookup')) {
      return "I can help you with Excel formulas. For SUM, you would use =SUM(range). For VLOOKUP, use =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]). Would you like me to explain a specific formula?";
    } else if (lower.includes('privacy') || lower.includes('local')) {
      return "This is a privacy-focused local AI mode. Your data never leaves your machine. All processing happens locally through Ollama or Web Workers.";
    } else if (lower.includes('hello') || lower.includes('hi')) {
      return "Hello! I'm your local Excel AI assistant. I can help with formulas, data analysis, and Excel tips - all while keeping your data private.";
    }
    
    return "I'm running in local/offline mode. I can help with Excel formulas, basic data analysis, and general Excel questions. For more advanced features, please use the cloud AI mode.";
  }

  // ============================================================================
  // Configuration
  // ============================================================================

  /**
   * Update configuration
   */
  updateConfig(config: Partial<LocalAIConfig>): void {
    this.config = { ...this.config, ...config };
    this.isInitialized = false; // Require re-initialization
  }

  /**
   * Get current configuration
   */
  getConfig(): LocalAIConfig {
    return { ...this.config };
  }

  /**
   * Check if service is initialized
   */
  isReady(): boolean {
    return this.isInitialized;
  }

  /**
   * Get current provider
   */
  getProvider(): LocalAIProvider {
    return this.config.provider;
  }
}

// Export singleton instance
export const localAIService = LocalAIService.getInstance();
export default localAIService;
