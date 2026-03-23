/**
 * Jest Test Setup
 * Phase 2: Comprehensive Testing
 * 
 * This file is run before each test file and sets up the testing environment
 */

import '@testing-library/jest-dom';

// ============================================================================
// POLYFILLS
// ============================================================================

// TextEncoder/TextDecoder polyfill for Node.js environment
import { TextEncoder, TextDecoder } from 'util';
global.TextEncoder = TextEncoder as any;
global.TextDecoder = TextDecoder as any;

// Also add to global scope for crypto mock
if (typeof global.TextEncoder === 'undefined') {
  global.TextEncoder = TextEncoder as any;
}
if (typeof global.TextDecoder === 'undefined') {
  global.TextDecoder = TextDecoder as any;
}

// ============================================================================
// GLOBAL MOCKS
// ============================================================================

// Mock Office.js
declare global {
  interface Window {
    Office: any;
    Excel: any;
  }
}

// Office.js mock
const mockOffice = {
  context: {
    document: {
      url: 'test-workbook.xlsx',
      settings: {
        get: jest.fn(),
        set: jest.fn(),
        remove: jest.fn(),
        saveAsync: jest.fn((callback?: (result: any) => void) => {
          if (callback) callback({ status: 'succeeded' });
        })
      }
    },
    workbook: {
      name: 'TestWorkbook'
    }
  },
  initialize: jest.fn((callback: () => void) => callback()),
  onReady: jest.fn((callback: (info: any) => void) => callback({ host: 'Excel' })),
  run: jest.fn(async (callback: (context: any) => Promise<any>) => {
    const mockContext = createMockExcelContext();
    return callback(mockContext);
  })
};

// Excel API mock
const mockExcel = {
  run: mockOffice.run,
  CalculationType: {
    automatic: 'automatic',
    full: 'full'
  },
  ChartType: {
    columnClustered: 'columnClustered',
    barClustered: 'barClustered',
    line: 'line',
    pie: 'pie',
    xyscatter: 'xyscatter'
  },
  RangeCopyType: {
    all: 'all',
    values: 'values',
    formats: 'formats'
  },
  DataValidationType: {
    list: 'list',
    whole: 'whole',
    decimal: 'decimal',
    date: 'date',
    textLength: 'textLength',
    custom: 'custom'
  },
  DataValidationAlertStyle: {
    stop: 'stop',
    warning: 'warning',
    information: 'information'
  },
  BorderLineStyle: {
    none: 'none',
    continuous: 'continuous',
    dash: 'dash'
  },
  BorderWeight: {
    hairline: 'hairline',
    thin: 'thin',
    medium: 'medium',
    thick: 'thick'
  },
  BorderIndex: {
    edgeTop: 'edgeTop',
    edgeBottom: 'edgeBottom',
    edgeLeft: 'edgeLeft',
    edgeRight: 'edgeRight'
  },
  ChartLegendPosition: {
    bottom: 'bottom',
    top: 'top',
    left: 'left',
    right: 'right'
  },
  HorizontalAlignment: {
    left: 'left',
    center: 'center',
    right: 'right'
  },
  VerticalAlignment: {
    top: 'top',
    center: 'center',
    bottom: 'bottom'
  },
  FillPatternType: {
    solid: 'solid'
  }
};

// Create mock Excel context
function createMockExcelContext() {
  const mockRange = {
    address: 'Sheet1!A1:B2',
    values: [['A', 'B'], ['1', '2']],
    formulas: [['A', 'B'], ['1', '2']],
    numberFormat: [['General', 'General']],
    rowCount: 2,
    columnCount: 2,
    worksheet: { name: 'Sheet1' },
    load: jest.fn().mockReturnThis(),
    clear: jest.fn(),
    copyFrom: jest.fn(),
    format: {
      font: { name: 'Calibri', size: 11, bold: false, italic: false, color: '#000000' },
      fill: { pattern: 'solid', color: '#FFFFFF' },
      borders: {
        getItem: jest.fn().mockReturnThis(),
        items: [],
        load: jest.fn().mockReturnThis()
      },
      horizontalAlignment: 'general',
      verticalAlignment: 'bottom',
      wrapText: false,
      autoFitColumns: jest.fn(),
      autoFitRows: jest.fn(),
      protection: { locked: true, hidden: false }
    },
    dataValidation: {
      rule: {},
      ignoreBlanks: true,
      prompt: {},
      errorAlert: {},
      clear: jest.fn()
    },
    merge: jest.fn(),
    unmerge: jest.fn()
  };

  const mockTable = {
    name: 'Table1',
    range: mockRange,
    style: 'TableStyleMedium2',
    rowCount: 10,
    columns: { count: 5, getItem: jest.fn() },
    rows: { add: jest.fn() },
    showTotals: false,
    getHeaderRowRange: jest.fn().mockReturnValue(mockRange),
    getDataBodyRange: jest.fn().mockReturnValue(mockRange),
    delete: jest.fn(),
    load: jest.fn().mockReturnThis()
  };

  const mockChart = {
    name: 'Chart1',
    type: 'columnClustered',
    title: { text: 'Test Chart' },
    legend: { position: 'bottom' },
    dataLabels: { visible: false },
    delete: jest.fn(),
    load: jest.fn().mockReturnThis()
  };

  const mockPivotTable = {
    name: 'PivotTable1',
    rowHierarchies: { add: jest.fn() },
    columnHierarchies: { add: jest.fn() },
    dataHierarchies: { add: jest.fn() },
    filterHierarchies: { add: jest.fn() },
    hierarchies: { getItem: jest.fn() },
    refresh: jest.fn(),
    load: jest.fn().mockReturnThis()
  };

  const mockWorksheet = {
    name: 'Sheet1',
    index: 0,
    tables: {
      items: [mockTable],
      getItem: jest.fn().mockReturnValue(mockTable),
      add: jest.fn().mockReturnValue(mockTable),
      load: jest.fn().mockReturnThis()
    },
    charts: {
      items: [mockChart],
      getItem: jest.fn().mockReturnValue(mockChart),
      add: jest.fn().mockReturnValue(mockChart),
      load: jest.fn().mockReturnThis()
    },
    pivotTables: {
      items: [mockPivotTable],
      getItem: jest.fn().mockReturnValue(mockPivotTable),
      add: jest.fn().mockReturnValue(mockPivotTable),
      load: jest.fn().mockReturnThis()
    },
    getRange: jest.fn().mockReturnValue(mockRange),
    getUsedRange: jest.fn().mockReturnValue(mockRange),
    getRangeByIndexes: jest.fn().mockReturnValue(mockRange),
    delete: jest.fn(),
    calculate: jest.fn(),
    protection: {
      protected: false,
      protect: jest.fn(),
      unprotect: jest.fn()
    },
    load: jest.fn().mockReturnThis()
  };

  return {
    workbook: {
      worksheets: {
        items: [mockWorksheet],
        getActiveWorksheet: jest.fn().mockReturnValue(mockWorksheet),
        getItem: jest.fn().mockReturnValue(mockWorksheet),
        add: jest.fn().mockReturnValue(mockWorksheet),
        load: jest.fn().mockReturnThis()
      },
      getSelectedRange: jest.fn().mockReturnValue(mockRange),
      names: {
        items: [],
        add: jest.fn(),
        getItem: jest.fn(),
        load: jest.fn().mockReturnThis()
      },
      tables: []
    },
    application: {
      calculate: jest.fn()
    },
    sync: jest.fn().mockResolvedValue(undefined),
    load: jest.fn().mockReturnThis(),
    trackedObjects: {
      add: jest.fn(),
      remove: jest.fn()
    }
  };
}

// Assign mocks to global
global.Office = mockOffice;
global.Excel = mockExcel;

// Mock Web Crypto API
const mockCrypto = {
  subtle: {
    importKey: jest.fn().mockResolvedValue({}),
    deriveKey: jest.fn().mockResolvedValue({}),
    encrypt: jest.fn().mockResolvedValue(new ArrayBuffer(16)),
    decrypt: jest.fn().mockResolvedValue(new TextEncoder().encode('decrypted')),
    digest: jest.fn().mockResolvedValue(new ArrayBuffer(32))
  },
  getRandomValues: jest.fn((arr: Uint8Array) => {
    for (let i = 0; i < arr.length; i++) {
      arr[i] = Math.floor(Math.random() * 256);
    }
    return arr;
  })
};

Object.defineProperty(global, 'crypto', {
  value: mockCrypto
});

// Mock localStorage
const localStorageMock = (() => {
  let store: Record<string, string> = {};
  return {
    getItem: jest.fn((key: string) => store[key] || null),
    setItem: jest.fn((key: string, value: string) => {
      store[key] = value;
    }),
    removeItem: jest.fn((key: string) => {
      delete store[key];
    }),
    clear: jest.fn(() => {
      store = {};
    }),
    get length() {
      return Object.keys(store).length;
    },
    key: jest.fn((index: number) => Object.keys(store)[index] || null)
  };
})();

Object.defineProperty(global, 'localStorage', {
  value: localStorageMock
});

// Mock sessionStorage
Object.defineProperty(global, 'sessionStorage', {
  value: localStorageMock
});

// Mock fetch
global.fetch = jest.fn();

// Mock ResizeObserver
class ResizeObserverMock {
  observe = jest.fn();
  unobserve = jest.fn();
  disconnect = jest.fn();
}
global.ResizeObserver = ResizeObserverMock as any;

// Mock IntersectionObserver
class IntersectionObserverMock {
  observe = jest.fn();
  unobserve = jest.fn();
  disconnect = jest.fn();
}
global.IntersectionObserver = IntersectionObserverMock as any;

// Mock matchMedia
Object.defineProperty(window, 'matchMedia', {
  value: jest.fn().mockImplementation((query: string) => ({
    matches: false,
    media: query,
    onchange: null,
    addListener: jest.fn(),
    removeListener: jest.fn(),
    addEventListener: jest.fn(),
    removeEventListener: jest.fn(),
    dispatchEvent: jest.fn()
  }))
});

// Mock clipboard API
Object.defineProperty(navigator, 'clipboard', {
  value: {
    writeText: jest.fn().mockResolvedValue(undefined),
    readText: jest.fn().mockResolvedValue('')
  }
});

// ============================================================================
// JEST EXTENSIONS
// ============================================================================

// Extend Jest matchers
expect.extend({
  toBeValidCellAddress(received: string) {
    const pass = /^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/.test(received);
    return {
      pass,
      message: () => `expected ${received} ${pass ? 'not to be' : 'to be'} a valid cell address`
    };
  },
  toBeWithinRange(received: number, floor: number, ceiling: number) {
    const pass = received >= floor && received <= ceiling;
    return {
      pass,
      message: () => `expected ${received} ${pass ? 'not to be' : 'to be'} within range ${floor} - ${ceiling}`
    };
  }
});

// ============================================================================
// TEST UTILITIES
// ============================================================================

// Helper to wait for async operations
export const waitFor = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

// Helper to flush promises
export const flushPromises = () => new Promise(setImmediate);

// Helper to create mock event
export const createMockEvent = (overrides: Partial<Event> = {}) => ({
  preventDefault: jest.fn(),
  stopPropagation: jest.fn(),
  ...overrides
} as Event);

// Helper to create mock keyboard event
export const createMockKeyboardEvent = (key: string, overrides: Partial<KeyboardEvent> = {}) => ({
  key,
  preventDefault: jest.fn(),
  stopPropagation: jest.fn(),
  ...overrides
} as unknown as KeyboardEvent);

// Reset all mocks after each test
afterEach(() => {
  jest.clearAllMocks();
  localStorageMock.clear();
});

// Clean up after all tests
afterAll(() => {
  jest.restoreAllMocks();
});

export { createMockExcelContext };