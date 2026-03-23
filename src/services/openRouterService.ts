// OpenRouter Service - Dynamic Model Management
// Fetches available models from OpenRouter API with caching

import { logger } from '../utils/logger';

export interface OpenRouterModel {
  id: string;
  name: string;
  description?: string;
  pricing: {
    prompt: number;  // per 1M tokens
    completion: number;
  };
  context_length: number;
  supports_function_calling: boolean;
  supports_vision?: boolean;
}

export interface OpenRouterModelsResponse {
  data: Array<{
    id: string;
    name: string;
    description?: string;
    pricing: {
      prompt: string;
      completion: string;
    };
    context_length: number;
    supports_function_calling: boolean;
    supports_vision?: boolean;
  }>;
}

export interface CachedModels {
  models: OpenRouterModel[];
  cachedAt: number;
}

// Default models to use when API is unavailable
const DEFAULT_MODELS: OpenRouterModel[] = [
  {
    id: 'openai/gpt-4',
    name: 'GPT-4',
    description: 'OpenAI GPT-4 - Most capable model',
    pricing: { prompt: 30, completion: 60 },
    context_length: 8192,
    supports_function_calling: true
  },
  {
    id: 'openai/gpt-4-turbo',
    name: 'GPT-4 Turbo',
    description: 'OpenAI GPT-4 Turbo - Faster and cheaper',
    pricing: { prompt: 10, completion: 30 },
    context_length: 128000,
    supports_function_calling: true
  },
  {
    id: 'openai/gpt-3.5-turbo',
    name: 'GPT-3.5 Turbo',
    description: 'OpenAI GPT-3.5 - Fast and affordable',
    pricing: { prompt: 0.5, completion: 1.5 },
    context_length: 16385,
    supports_function_calling: true
  },
  {
    id: 'anthropic/claude-3-opus-20240229',
    name: 'Claude 3 Opus',
    description: 'Anthropic Claude 3 Opus - Most capable',
    pricing: { prompt: 15, completion: 75 },
    context_length: 200000,
    supports_function_calling: false
  },
  {
    id: 'anthropic/claude-3-sonnet-20240229',
    name: 'Claude 3 Sonnet',
    description: 'Anthropic Claude 3 Sonnet - Balanced',
    pricing: { prompt: 3, completion: 15 },
    context_length: 200000,
    supports_function_calling: false
  },
  {
    id: 'meta-llama/llama-3-70b-instruct',
    name: 'Llama 3 70B',
    description: 'Meta Llama 3 70B - Open source',
    pricing: { prompt: 0.8, completion: 0.8 },
    context_length: 8192,
    supports_function_calling: false
  },
  {
    id: 'mistralai/mixtral-8x7b-instruct',
    name: 'Mixtral 8x7B',
    description: 'Mistral Mixtral - Expert mixture',
    pricing: { prompt: 0.6, completion: 0.6 },
    context_length: 32000,
    supports_function_calling: false
  },
  {
    id: 'google/gemini-pro-1.5',
    name: 'Gemini Pro 1.5',
    description: 'Google Gemini Pro - Large context',
    pricing: { prompt: 1.25, completion: 5 },
    context_length: 1000000,
    supports_function_calling: false
  }
];

const CACHE_KEY = 'openrouter_models_cache';
const CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

export class OpenRouterService {
  private static instance: OpenRouterService;
  private cachedModels: CachedModels | null = null;
  private isFetching: boolean = false;
  private fetchPromise: Promise<OpenRouterModel[]> | null = null;

  private constructor() {
    // Load cache from localStorage on initialization
    this.loadFromCache();
  }

  static getInstance(): OpenRouterService {
    if (!OpenRouterService.instance) {
      OpenRouterService.instance = new OpenRouterService();
    }
    return OpenRouterService.instance;
  }

  /**
   * Check if cache is valid
   */
  private isCacheValid(): boolean {
    if (!this.cachedModels) return false;
    return Date.now() - this.cachedModels.cachedAt < CACHE_TTL_MS;
  }

  /**
   * Load cached models from localStorage
   */
  private loadFromCache(): void {
    try {
      const cached = localStorage.getItem(CACHE_KEY);
      if (cached) {
        this.cachedModels = JSON.parse(cached);
      }
    } catch (error) {
      logger.warn('Failed to load cached models', undefined, error as Error);
    }
  }

  /**
   * Save models to localStorage cache
   */
  private saveToCache(models: OpenRouterModel[]): void {
    try {
      this.cachedModels = {
        models,
        cachedAt: Date.now()
      };
      localStorage.setItem(CACHE_KEY, JSON.stringify(this.cachedModels));
    } catch (error) {
      logger.warn('Failed to save models to cache', undefined, error as Error);
    }
  }

  /**
   * Check if the API URL is OpenRouter (static method for easy access)
   */
  static isOpenRouterUrl(url: string): boolean {
    return url.includes('openrouter.ai');
  }

  /**
   * Get headers for OpenRouter API requests
   */
  getHeaders(apiKey: string): HeadersInit {
    return {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
      'HTTP-Referer': 'https://excel-ai-assistant.com',
      'X-Title': 'Excel AI Assistant'
    };
  }

  /**
   * Fetch available models from OpenRouter API
   */
  async fetchModels(apiUrl: string, apiKey: string): Promise<OpenRouterModel[]> {
    // Return cached models if valid
    if (this.isCacheValid() && this.cachedModels?.models) {
      return this.cachedModels.models;
    }

    // Return cached models if already fetching
    if (this.isFetching && this.fetchPromise) {
      return this.fetchPromise;
    }

    this.isFetching = true;

    this.fetchPromise = this.doFetchModels(apiUrl, apiKey);

    try {
      const models = await this.fetchPromise;
      this.saveToCache(models);
      return models;
    } finally {
      this.isFetching = false;
      this.fetchPromise = null;
    }
  }

  /**
   * Actually fetch models from API
   */
  private async doFetchModels(apiUrl: string, apiKey: string): Promise<OpenRouterModel[]> {
    // Replace /v1/chat/completions with /v1/models for the models endpoint
    const modelsUrl = apiUrl.replace('/v1/chat/completions', '/v1/models');

    const response = await fetch(modelsUrl, {
      method: 'GET',
      headers: this.getHeaders(apiKey)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to fetch models: ${response.status} - ${errorText}`);
    }

    const data: OpenRouterModelsResponse = await response.json();

    // Transform API response to our format
    const models: OpenRouterModel[] = data.data.map(model => ({
      id: model.id,
      name: this.formatModelName(model.id, model.name),
      description: model.description,
      pricing: {
        prompt: parseFloat(model.pricing.prompt) * 1000000, // Convert from per-token to per-1M
        completion: parseFloat(model.pricing.completion) * 1000000
      },
      context_length: model.context_length,
      supports_function_calling: model.supports_function_calling,
      supports_vision: model.supports_vision
    }));

    // Sort by pricing (cheapest first) for easier selection
    models.sort((a, b) => a.pricing.prompt - b.pricing.prompt);

    return models;
  }

  /**
   * Format model ID into a readable name
   */
  private formatModelName(id: string, apiName: string): string {
    // Extract provider and model name
    const parts = id.split('/');
    if (parts.length >= 2) {
      const provider = parts[0];
      const modelName = parts[1];
      
      // Format provider name
      const providerNames: Record<string, string> = {
        'openai': 'OpenAI',
        'anthropic': 'Anthropic',
        'meta-llama': 'Meta Llama',
        'mistralai': 'Mistral',
        'google': 'Google',
        'cohere': 'Cohere',
        'ai21': 'AI21',
        'togetherai': 'Together'
      };
      
      const providerDisplay = providerNames[provider] || provider;
      return `${providerDisplay} ${this.formatModelNameOnly(modelName)}`;
    }
    
    return apiName || id;
  }

  /**
   * Format just the model name part
   */
  private formatModelNameOnly(name: string): string {
    return name
      .replace(/-/g, ' ')
      .replace(/instruct/i, '(Instruct)')
      .replace(/chat/i, '(Chat)')
      .replace(/preview/i, '(Preview)')
      .replace(/v(\d+)/i, 'v$1');
  }

  /**
   * Get available models - fetches from API or returns defaults
   */
  async getAvailableModels(apiUrl: string, apiKey: string): Promise<OpenRouterModel[]> {
    if (!OpenRouterService.isOpenRouterUrl(apiUrl)) {
      // Not OpenRouter, return empty (will use defaults)
      return [];
    }

    try {
      return await this.fetchModels(apiUrl, apiKey);
    } catch (error) {
      logger.error('Failed to fetch OpenRouter models', undefined, error as Error);
      // Return default models on error
      return DEFAULT_MODELS;
    }
  }

  /**
   * Get models that support function calling (for tool use)
   */
  getFunctionCallingModels(models: OpenRouterModel[]): OpenRouterModel[] {
    return models.filter(m => m.supports_function_calling);
  }

  /**
   * Get models sorted by price
   */
  getModelsByPrice(models: OpenRouterModel[]): OpenRouterModel[] {
    return [...models].sort((a, b) => a.pricing.prompt - b.pricing.prompt);
  }

  /**
   * Get models that support vision
   */
  getVisionModels(models: OpenRouterModel[]): OpenRouterModel[] {
    return models.filter(m => m.supports_vision);
  }

  /**
   * Get default models (fallback when API unavailable)
   */
  getDefaultModels(): OpenRouterModel[] {
    return DEFAULT_MODELS;
  }

  /**
   * Format price for display
   */
  formatPrice(pricePerMillion: number): string {
    if (pricePerMillion < 0.01) {
      return `$${pricePerMillion.toFixed(4)}/M`;
    }
    return `$${pricePerMillion.toFixed(2)}/M`;
  }

  /**
   * Format context length for display
   */
  formatContextLength(tokens: number): string {
    if (tokens >= 1000000) {
      return `${(tokens / 1000000).toFixed(0)}M`;
    }
    if (tokens >= 1000) {
      return `${(tokens / 1000).toFixed(0)}K`;
    }
    return tokens.toString();
  }

  /**
   * Clear the cache
   */
  clearCache(): void {
    this.cachedModels = null;
    localStorage.removeItem(CACHE_KEY);
  }

  /**
   * Force refresh models from API
   */
  async refreshModels(apiUrl: string, apiKey: string): Promise<OpenRouterModel[]> {
    this.clearCache();
    return this.getAvailableModels(apiUrl, apiKey);
  }
}

export default OpenRouterService.getInstance();
