/**
 * Unit Tests for ConversationStorage
 * Tests conversation persistence, auto-save, and debouncing
 */

import { ConversationStorage, AutoSaveManager, Conversation } from '../conversationStorage';

// Mock logger
jest.mock('../../utils/logger', () => ({
  logger: {
    info: jest.fn(),
    error: jest.fn(),
    warn: jest.fn(),
    debug: jest.fn()
  }
}));

// Mock IndexedDB
const mockIndexedDB = {
  open: jest.fn(),
  deleteDatabase: jest.fn()
};

const mockDB = {
  createObjectStore: jest.fn(),
  transaction: jest.fn(),
  objectStoreNames: { contains: jest.fn() }
};

const mockStore = {
  add: jest.fn(),
  put: jest.fn(),
  get: jest.fn(),
  delete: jest.fn(),
  getAll: jest.fn(),
  createIndex: jest.fn(),
  clear: jest.fn()
};

const mockTransaction = {
  objectStore: jest.fn().mockReturnValue(mockStore)
};

const mockRequest = {
  onsuccess: null as any,
  onerror: null as any,
  onupgradeneeded: null as any,
  result: mockDB,
  error: null
};

beforeAll(() => {
  global.indexedDB = mockIndexedDB as any;
});

beforeEach(() => {
  jest.clearAllMocks();
  mockIndexedDB.open.mockReturnValue(mockRequest);
  mockDB.transaction.mockReturnValue(mockTransaction);
  mockDB.objectStoreNames.contains.mockReturnValue(true);
});

describe('ConversationStorage', () => {
  let storage: ConversationStorage;

  beforeEach(() => {
    storage = new ConversationStorage();
  });

  describe('initialization', () => {
    it('should initialize IndexedDB', async () => {
      const initPromise = storage.initialize();
      
      // Simulate successful open
      mockRequest.onsuccess?.();
      
      await initPromise;
      expect(mockIndexedDB.open).toHaveBeenCalledWith('ExcelAIAssistantDB', 1);
    });

    it('should handle initialization errors', async () => {
      const initPromise = storage.initialize();
      
      // Simulate error
      mockRequest.onerror?.();
      
      await expect(initPromise).rejects.toThrow('Failed to open IndexedDB');
    });

    it('should only initialize once', async () => {
      const promise1 = storage.initialize();
      mockRequest.onsuccess?.();
      await promise1;

      const promise2 = storage.initialize();
      mockRequest.onsuccess?.();
      await promise2;

      expect(mockIndexedDB.open).toHaveBeenCalledTimes(1);
    });
  });

  describe('createConversation', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should create a new conversation', async () => {
      const mockConversation = {
        id: expect.any(String),
        workbookId: 'workbook-1',
        workbookName: 'Test Workbook',
        title: 'New Conversation',
        messages: [],
        createdAt: expect.any(Date),
        updatedAt: expect.any(Date),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const addRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.add.mockReturnValue(addRequest);

      const createPromise = storage.createConversation('workbook-1', 'Test Workbook');
      addRequest.onsuccess?.();

      const conversation = await createPromise;
      expect(conversation).toMatchObject({
        workbookId: 'workbook-1',
        workbookName: 'Test Workbook',
        title: 'New Conversation',
        messages: [],
        messageCount: 0
      });
      expect(mockStore.add).toHaveBeenCalled();
    });

    it('should handle creation errors', async () => {
      const addRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.add.mockReturnValue(addRequest);

      const createPromise = storage.createConversation('workbook-1', 'Test');
      addRequest.onerror?.();

      await expect(createPromise).rejects.toThrow('Failed to create conversation');
    });
  });

  describe('saveConversation', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should save an existing conversation', async () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test Conversation',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      const savePromise = storage.saveConversation(conversation);
      putRequest.onsuccess?.();

      await savePromise;
      expect(mockStore.put).toHaveBeenCalled();
    });

    it('should update message count on save', async () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [
          { id: '1', role: 'user', content: 'Hello', timestamp: new Date() },
          { id: '2', role: 'assistant', content: 'Hi', timestamp: new Date() }
        ],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      const savePromise = storage.saveConversation(conversation);
      putRequest.onsuccess?.();

      await savePromise;
      expect(conversation.messageCount).toBe(2);
    });
  });

  describe('getConversation', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should retrieve a conversation by ID', async () => {
      const mockData = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        messageCount: 0
      };

      const getRequest = { onsuccess: null as any, onerror: null as any, result: mockData };
      mockStore.get.mockReturnValue(getRequest);

      const getPromise = storage.getConversation('conv-1');
      getRequest.onsuccess?.();

      const conversation = await getPromise;
      expect(conversation).toBeDefined();
      expect(conversation?.id).toBe('conv-1');
    });

    it('should return null for non-existent conversation', async () => {
      const getRequest = { onsuccess: null as any, onerror: null as any, result: null };
      mockStore.get.mockReturnValue(getRequest);

      const getPromise = storage.getConversation('non-existent');
      getRequest.onsuccess?.();

      const conversation = await getPromise;
      expect(conversation).toBeNull();
    });
  });

  describe('deleteConversation', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should delete a conversation', async () => {
      const deleteRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.delete.mockReturnValue(deleteRequest);

      const deletePromise = storage.deleteConversation('conv-1');
      deleteRequest.onsuccess?.();

      await deletePromise;
      expect(mockStore.delete).toHaveBeenCalledWith('conv-1');
    });
  });

  describe('importConversation', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should import a valid conversation JSON', async () => {
      const jsonData = JSON.stringify({
        title: 'Imported Conversation',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        messages: [
          { id: '1', role: 'user', content: 'Hello', timestamp: new Date().toISOString() }
        ]
      });

      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      const importPromise = storage.importConversation(jsonData);
      putRequest.onsuccess?.();

      const conversation = await importPromise;
      expect(conversation.title).toBe('Imported Conversation (Imported)');
      expect(conversation.messages.length).toBe(1);
    });

    it('should reject invalid JSON', async () => {
      await expect(storage.importConversation('invalid json')).rejects.toThrow();
    });

    it('should reject conversation without required fields', async () => {
      const jsonData = JSON.stringify({ invalid: true });
      await expect(storage.importConversation(jsonData)).rejects.toThrow('Invalid conversation format');
    });
  });

  describe('clearAllData', () => {
    beforeEach(async () => {
      const initPromise = storage.initialize();
      mockRequest.onsuccess?.();
      await initPromise;
    });

    it('should clear all data', async () => {
      const clearRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.clear = jest.fn().mockReturnValue(clearRequest);

      const clearPromise = storage.clearAllData();
      clearRequest.onsuccess?.();

      await clearPromise;
      expect(mockStore.clear).toHaveBeenCalled();
    });
  });
});

describe('AutoSaveManager', () => {
  let storage: ConversationStorage;
  let autoSaveManager: AutoSaveManager;

  beforeEach(async () => {
    storage = new ConversationStorage();
    const initPromise = storage.initialize();
    mockRequest.onsuccess?.();
    await initPromise;

    autoSaveManager = new AutoSaveManager(storage);
    jest.useFakeTimers();
  });

  afterEach(() => {
    jest.useRealTimers();
    autoSaveManager.stopAutoSave();
  });

  describe('startAutoSave', () => {
    it('should start auto-save for a conversation', () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      autoSaveManager.startAutoSave(conversation);
      expect(autoSaveManager.isDirtyConversation()).toBe(false);
    });
  });

  describe('markDirty', () => {
    it('should mark conversation as dirty', () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      autoSaveManager.startAutoSave(conversation);
      autoSaveManager.markDirty();
      expect(autoSaveManager.isDirtyConversation()).toBe(true);
    });
  });

  describe('debouncing', () => {
    it('should debounce saves', async () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const saveSpy = jest.spyOn(storage, 'saveConversation').mockResolvedValue();

      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      autoSaveManager.startAutoSave(conversation);
      
      // Mark dirty multiple times rapidly
      autoSaveManager.markDirty();
      autoSaveManager.markDirty();
      autoSaveManager.markDirty();

      // Advance time but not past debounce
      jest.advanceTimersByTime(3000);
      expect(saveSpy).not.toHaveBeenCalled();

      // Advance past debounce
      jest.advanceTimersByTime(3000);
      // Save should be called once
      expect(saveSpy).toHaveBeenCalledTimes(1);
    });

    it('should set debounce interval', () => {
      autoSaveManager.setDebounceInterval(10000);
      
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const saveSpy = jest.spyOn(storage, 'saveConversation').mockResolvedValue();
      
      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      autoSaveManager.startAutoSave(conversation);
      autoSaveManager.markDirty();

      // Should not save after default 5 seconds
      jest.advanceTimersByTime(5000);
      expect(saveSpy).not.toHaveBeenCalled();

      // Should save after new 10 second interval
      jest.advanceTimersByTime(6000);
      expect(saveSpy).toHaveBeenCalledTimes(1);
    });
  });

  describe('forceSave', () => {
    it('should force immediate save', async () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      const saveSpy = jest.spyOn(storage, 'saveConversation').mockResolvedValue();
      
      const putRequest = { onsuccess: null as any, onerror: null as any };
      mockStore.put.mockReturnValue(putRequest);

      autoSaveManager.startAutoSave(conversation);
      autoSaveManager.markDirty();

      const forcePromise = autoSaveManager.forceSave();
      putRequest.onsuccess?.();
      await forcePromise;

      expect(saveSpy).toHaveBeenCalled();
      expect(autoSaveManager.isDirtyConversation()).toBe(false);
    });
  });

  describe('stopAutoSave', () => {
    it('should stop auto-save', () => {
      const conversation: Conversation = {
        id: 'conv-1',
        workbookId: 'workbook-1',
        workbookName: 'Test',
        title: 'Test',
        messages: [],
        createdAt: new Date(),
        updatedAt: new Date(),
        messageCount: 0,
        isPinned: false,
        tags: []
      };

      autoSaveManager.startAutoSave(conversation);
      autoSaveManager.markDirty();
      autoSaveManager.stopAutoSave();

      expect(autoSaveManager.isDirtyConversation()).toBe(false);
    });
  });
});