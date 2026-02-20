// Conversation Storage Service
// Manages persistent storage of chat conversations using IndexedDB

import { Message } from "@/types";

export interface Conversation {
  id: string;
  workbookId: string;
  workbookName: string;
  title: string;
  messages: Message[];
  createdAt: Date;
  updatedAt: Date;
  messageCount: number;
  isPinned?: boolean;
  tags?: string[];
}

export interface ConversationSummary {
  id: string;
  title: string;
  workbookName: string;
  lastMessage: string;
  lastMessageDate: Date;
  messageCount: number;
  isPinned: boolean;
}

export interface StorageStats {
  totalConversations: number;
  totalMessages: number;
  oldestConversation: Date;
  storageUsed: number; // in bytes
}

const DB_NAME = 'ExcelAIAssistantDB';
const DB_VERSION = 1;
const CONVERSATION_STORE = 'conversations';
const MAX_STORAGE_SIZE = 50 * 1024 * 1024; // 50MB limit

export class ConversationStorage {
  private db: IDBDatabase | null = null;
  private initPromise: Promise<void> | null = null;

  // Initialize IndexedDB
  async initialize(): Promise<void> {
    if (this.initPromise) {
      return this.initPromise;
    }

    this.initPromise = new Promise((resolve, reject) => {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onerror = () => {
        reject(new Error('Failed to open IndexedDB'));
      };

      request.onsuccess = () => {
        this.db = request.result;
        resolve();
      };

      request.onupgradeneeded = (event) => {
        const db = (event.target as IDBOpenDBRequest).result;

        // Create conversations store
        if (!db.objectStoreNames.contains(CONVERSATION_STORE)) {
          const store = db.createObjectStore(CONVERSATION_STORE, { keyPath: 'id' });
          store.createIndex('workbookId', 'workbookId', { unique: false });
          store.createIndex('updatedAt', 'updatedAt', { unique: false });
          store.createIndex('isPinned', 'isPinned', { unique: false });
        }
      };
    });

    return this.initPromise;
  }

  // Get database instance
  private async getDB(): Promise<IDBDatabase> {
    if (!this.db) {
      await this.initialize();
    }
    if (!this.db) {
      throw new Error('Database not initialized');
    }
    return this.db;
  }

  // Generate unique ID
  private generateId(): string {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  // Create a new conversation
  async createConversation(
    workbookId: string,
    workbookName: string,
    title: string = 'New Conversation'
  ): Promise<Conversation> {
    const db = await this.getDB();

    const conversation: Conversation = {
      id: this.generateId(),
      workbookId,
      workbookName,
      title,
      messages: [],
      createdAt: new Date(),
      updatedAt: new Date(),
      messageCount: 0,
      isPinned: false,
      tags: []
    };

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readwrite');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.add(conversation);

      request.onsuccess = () => resolve(conversation);
      request.onerror = () => reject(new Error('Failed to create conversation'));
    });
  }

  // Save or update a conversation
  async saveConversation(conversation: Conversation): Promise<void> {
    const db = await this.getDB();

    // Update metadata
    conversation.updatedAt = new Date();
    conversation.messageCount = conversation.messages.length;

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readwrite');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.put(conversation);

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Failed to save conversation'));
    });
  }

  // Get a conversation by ID
  async getConversation(id: string): Promise<Conversation | null> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readonly');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.get(id);

      request.onsuccess = () => {
        const result = request.result;
        if (result) {
          // Convert dates back from ISO strings
          result.createdAt = new Date(result.createdAt);
          result.updatedAt = new Date(result.updatedAt);
          result.messages.forEach((msg: Message) => {
            msg.timestamp = new Date(msg.timestamp);
          });
          resolve(result);
        } else {
          resolve(null);
        }
      };
      request.onerror = () => reject(new Error('Failed to get conversation'));
    });
  }

  // Get all conversations for a workbook
  async getConversationsByWorkbook(workbookId: string): Promise<Conversation[]> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readonly');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const index = store.index('workbookId');
      const request = index.getAll(workbookId);

      request.onsuccess = () => {
        const results = request.result;
        // Convert dates
        results.forEach((conv: Conversation) => {
          conv.createdAt = new Date(conv.createdAt);
          conv.updatedAt = new Date(conv.updatedAt);
          conv.messages.forEach((msg: Message) => {
            msg.timestamp = new Date(msg.timestamp);
          });
        });
        resolve(results);
      };
      request.onerror = () => reject(new Error('Failed to get conversations'));
    });
  }

  // Get all conversation summaries (lightweight)
  async getConversationSummaries(): Promise<ConversationSummary[]> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readonly');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const index = store.index('updatedAt');
      const request = index.openCursor(null, 'prev'); // Most recent first

      const summaries: ConversationSummary[] = [];

      request.onsuccess = () => {
        const cursor = request.result;
        if (cursor) {
          const conv = cursor.value;
          const lastMessage = conv.messages[conv.messages.length - 1];

          summaries.push({
            id: conv.id,
            title: conv.title,
            workbookName: conv.workbookName,
            lastMessage: lastMessage ? lastMessage.content.substring(0, 100) : '',
            lastMessageDate: new Date(conv.updatedAt),
            messageCount: conv.messageCount,
            isPinned: conv.isPinned || false
          });

          cursor.continue();
        } else {
          resolve(summaries);
        }
      };

      request.onerror = () => reject(new Error('Failed to get summaries'));
    });
  }

  // Delete a conversation
  async deleteConversation(id: string): Promise<void> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readwrite');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.delete(id);

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Failed to delete conversation'));
    });
  }

  // Delete all conversations for a workbook
  async deleteConversationsByWorkbook(workbookId: string): Promise<number> {
    const conversations = await this.getConversationsByWorkbook(workbookId);

    for (const conv of conversations) {
      await this.deleteConversation(conv.id);
    }

    return conversations.length;
  }

  // Update conversation title
  async updateTitle(id: string, title: string): Promise<void> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    conversation.title = title;
    await this.saveConversation(conversation);
  }

  // Pin/unpin conversation
  async setPinned(id: string, isPinned: boolean): Promise<void> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    conversation.isPinned = isPinned;
    await this.saveConversation(conversation);
  }

  // Add tags to conversation
  async addTags(id: string, tags: string[]): Promise<void> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    conversation.tags = [...new Set([...(conversation.tags || []), ...tags])];
    await this.saveConversation(conversation);
  }

  // Remove tags from conversation
  async removeTags(id: string, tags: string[]): Promise<void> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    conversation.tags = (conversation.tags || []).filter(tag => !tags.includes(tag));
    await this.saveConversation(conversation);
  }

  // Search conversations
  async searchConversations(query: string): Promise<Conversation[]> {
    const allSummaries = await this.getConversationSummaries();
    const matchingIds = allSummaries
      .filter(summary =>
        summary.title.toLowerCase().includes(query.toLowerCase()) ||
        summary.lastMessage.toLowerCase().includes(query.toLowerCase()) ||
        summary.workbookName.toLowerCase().includes(query.toLowerCase())
      )
      .map(summary => summary.id);

    const conversations: Conversation[] = [];
    for (const id of matchingIds) {
      const conv = await this.getConversation(id);
      if (conv) {
        conversations.push(conv);
      }
    }

    return conversations;
  }

  // Export conversation as JSON
  async exportConversation(id: string): Promise<string> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    return JSON.stringify(conversation, null, 2);
  }

  // Export conversation as Markdown
  async exportConversationAsMarkdown(id: string): Promise<string> {
    const conversation = await this.getConversation(id);
    if (!conversation) {
      throw new Error('Conversation not found');
    }

    let markdown = `# ${conversation.title}\n\n`;
    markdown += `**Workbook:** ${conversation.workbookName}\n`;
    markdown += `**Date:** ${conversation.createdAt.toLocaleString()}\n`;
    markdown += `**Messages:** ${conversation.messageCount}\n\n`;
    markdown += `---\n\n`;

    for (const message of conversation.messages) {
      const role = message.role === 'user' ? '**You**' : '**AI Assistant**';
      const time = message.timestamp.toLocaleString();
      markdown += `${role} (${time}):\n\n`;
      markdown += `${message.content}\n\n`;
      markdown += `---\n\n`;
    }

    return markdown;
  }

  // Import conversation from JSON
  async importConversation(json: string): Promise<Conversation> {
    try {
      const data = JSON.parse(json);

      // Validate required fields
      if (!data.title || !data.messages) {
        throw new Error('Invalid conversation format');
      }

      // Create new conversation with new ID
      const conversation: Conversation = {
        id: this.generateId(),
        workbookId: data.workbookId || 'imported',
        workbookName: data.workbookName || 'Imported',
        title: data.title + ' (Imported)',
        messages: data.messages,
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: data.messages.length,
        isPinned: false,
        tags: [...(data.tags || []), 'imported']
      };

      await this.saveConversation(conversation);
      return conversation;
    } catch (error) {
      throw new Error(`Failed to import conversation: ${error.message}`);
    }
  }

  // Get storage statistics
  async getStats(): Promise<StorageStats> {
    const summaries = await this.getConversationSummaries();

    let totalMessages = 0;
    let oldestDate = new Date();

    for (const summary of summaries) {
      totalMessages += summary.messageCount;
      if (summary.lastMessageDate < oldestDate) {
        oldestDate = summary.lastMessageDate;
      }
    }

    // Estimate storage used (rough calculation)
    const storageUsed = await this.estimateStorageSize();

    return {
      totalConversations: summaries.length,
      totalMessages,
      oldestConversation: oldestDate,
      storageUsed
    };
  }

  // Estimate storage size
  private async estimateStorageSize(): Promise<number> {
    // This is a rough estimate
    const allConversations = await this.getAllConversations();
    const json = JSON.stringify(allConversations);
    return new Blob([json]).size;
  }

  // Get all conversations (for admin/backup)
  private async getAllConversations(): Promise<Conversation[]> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readonly');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.getAll();

      request.onsuccess = () => {
        const results = request.result;
        results.forEach((conv: Conversation) => {
          conv.createdAt = new Date(conv.createdAt);
          conv.updatedAt = new Date(conv.updatedAt);
          conv.messages.forEach((msg: Message) => {
            msg.timestamp = new Date(msg.timestamp);
          });
        });
        resolve(results);
      };
      request.onerror = () => reject(new Error('Failed to get all conversations'));
    });
  }

  // Clean up old conversations
  async cleanupOldConversations(daysToKeep: number = 30): Promise<number> {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

    const allConversations = await this.getAllConversations();
    const toDelete = allConversations.filter(
      conv => !conv.isPinned && conv.updatedAt < cutoffDate
    );

    for (const conv of toDelete) {
      await this.deleteConversation(conv.id);
    }

    return toDelete.length;
  }

  // Clear all data (use with caution)
  async clearAllData(): Promise<void> {
    const db = await this.getDB();

    return new Promise((resolve, reject) => {
      const transaction = db.transaction([CONVERSATION_STORE], 'readwrite');
      const store = transaction.objectStore(CONVERSATION_STORE);
      const request = store.clear();

      request.onsuccess = () => resolve();
      request.onerror = () => reject(new Error('Failed to clear data'));
    });
  }
}

// Singleton instance
let storageInstance: ConversationStorage | null = null;

export function getConversationStorage(): ConversationStorage {
  if (!storageInstance) {
    storageInstance = new ConversationStorage();
  }
  return storageInstance;
}

// Auto-save functionality
export class AutoSaveManager {
  private storage: ConversationStorage;
  private currentConversation: Conversation | null = null;
  private saveInterval: number = 30000; // 30 seconds
  private intervalId: number | null = null;

  constructor(storage: ConversationStorage) {
    this.storage = storage;
  }

  startAutoSave(conversation: Conversation): void {
    this.currentConversation = conversation;
    this.stopAutoSave();

    this.intervalId = window.setInterval(async () => {
      if (this.currentConversation) {
        try {
          await this.storage.saveConversation(this.currentConversation);
          console.log('Auto-saved conversation:', this.currentConversation.id);
        } catch (error) {
          console.error('Auto-save failed:', error);
        }
      }
    }, this.saveInterval);
  }

  stopAutoSave(): void {
    if (this.intervalId) {
      clearInterval(this.intervalId);
      this.intervalId = null;
    }
  }

  setInterval(intervalMs: number): void {
    this.saveInterval = intervalMs;
    if (this.currentConversation) {
      this.startAutoSave(this.currentConversation);
    }
  }

  async forceSave(): Promise<void> {
    if (this.currentConversation) {
      await this.storage.saveConversation(this.currentConversation);
    }
  }
}

export default getConversationStorage;
