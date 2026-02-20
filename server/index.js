/**
 * Excel AI Assistant - Express Server
 * Handles API proxying and backend services
 */

const express = require('express');
const path = require('path');
const fs = require('fs');
const readline = require('readline');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3001;

// ============================================================================
// LOGGING CONFIGURATION
// ============================================================================

const LOG_DIR = path.join(__dirname, 'logs');
const LOG_FILE = path.join(LOG_DIR, 'app.log');
const MAX_LOG_SIZE_MB = parseInt(process.env.MAX_LOG_SIZE_MB || '10', 10);
const MAX_LOG_FILES = parseInt(process.env.MAX_LOG_FILES || '5', 10);

// Ensure log directory exists
if (!fs.existsSync(LOG_DIR)) {
  fs.mkdirSync(LOG_DIR, { recursive: true });
}

/**
 * Server-side Logger class
 * Handles file logging with rotation
 */
class ServerLogger {
  constructor() {
    this.logStream = null;
    this.currentLogFile = LOG_FILE;
    this.initLogStream();
  }

  initLogStream() {
    this.logStream = fs.createWriteStream(this.currentLogFile, { flags: 'a' });
  }

  formatEntry(entry) {
    const levelNames = ['DEBUG', 'INFO', 'WARN', 'ERROR', 'SILENT'];
    const levelName = levelNames[entry.level] || 'UNKNOWN';
    const context = entry.context ? ` | Context: ${JSON.stringify(entry.context)}` : '';
    const error = entry.error ? ` | Error: ${entry.error.stack || entry.error.message}` : '';
    return `[${entry.timestamp}] [${levelName}] ${entry.message}${context}${error}\n`;
  }

  checkRotation() {
    try {
      const stats = fs.statSync(this.currentLogFile);
      const sizeMB = stats.size / (1024 * 1024);
      
      if (sizeMB >= MAX_LOG_SIZE_MB) {
        this.rotateLogs();
      }
    } catch (err) {
      // File might not exist yet, that's fine
    }
  }

  rotateLogs() {
    // Close current stream
    if (this.logStream) {
      this.logStream.end();
    }

    // Rotate existing log files
    for (let i = MAX_LOG_FILES - 1; i >= 1; i--) {
      const oldFile = path.join(LOG_DIR, `app.${i}.log`);
      const newFile = path.join(LOG_DIR, `app.${i + 1}.log`);
      
      if (fs.existsSync(oldFile)) {
        if (i === MAX_LOG_FILES - 1) {
          // Delete oldest file
          fs.unlinkSync(oldFile);
        } else {
          fs.renameSync(oldFile, newFile);
        }
      }
    }

    // Rename current log to .1
    if (fs.existsSync(this.currentLogFile)) {
      fs.renameSync(this.currentLogFile, path.join(LOG_DIR, 'app.1.log'));
    }

    // Reinitialize stream
    this.initLogStream();
  }

  log(entry) {
    this.checkRotation();
    const formatted = this.formatEntry(entry);
    
    // Write to file
    this.logStream.write(formatted);
    
    // Also output to console
    const levelNames = ['DEBUG', 'INFO', 'WARN', 'ERROR'];
    const consoleMethod = ['debug', 'info', 'warn', 'error'][entry.level] || 'log';
    console[consoleMethod](`[SERVER] ${entry.message}`, entry.context || '', entry.error || '');
  }

  getLogs(options = {}) {
    return new Promise((resolve, reject) => {
      const { lines = 100, level, since } = options;
      const logs = [];
      
      if (!fs.existsSync(this.currentLogFile)) {
        resolve([]);
        return;
      }

      const rl = readline.createInterface({
        input: fs.createReadStream(this.currentLogFile),
        crlfDelay: Infinity
      });

      rl.on('line', (line) => {
        try {
          // Parse log line: [timestamp] [LEVEL] message
          const match = line.match(/\[([^\]]+)\] \[(\w+)\] (.+)/);
          if (match) {
            const entry = {
              timestamp: match[1],
              level: levelNames.indexOf(match[2]),
              message: match[2]
            };
            
            if (level !== undefined && entry.level < level) return;
            logs.push(entry);
          }
        } catch {
          // Skip unparseable lines
        }
      });

      rl.on('close', () => {
        // Return last N lines
        resolve(logs.slice(-lines));
      });

      rl.on('error', reject);
    });
  }

  clearLogs() {
    // Delete all log files
    for (let i = 1; i <= MAX_LOG_FILES; i++) {
      const file = path.join(LOG_DIR, `app.${i}.log`);
      if (fs.existsSync(file)) {
        fs.unlinkSync(file);
      }
    }
    
    // Clear current log
    if (fs.existsSync(this.currentLogFile)) {
      fs.writeFileSync(this.currentLogFile, '');
    }
  }
}

const levelNames = ['DEBUG', 'INFO', 'WARN', 'ERROR', 'SILENT'];
const serverLogger = new ServerLogger();

// Global error handlers
process.on('uncaughtException', (error) => {
  serverLogger.log({
    timestamp: new Date().toISOString(),
    level: 3, // ERROR
    message: 'Uncaught Exception',
    error: { message: error.message, stack: error.stack }
  });
});

process.on('unhandledRejection', (reason, promise) => {
  serverLogger.log({
    timestamp: new Date().toISOString(),
    level: 3, // ERROR
    message: 'Unhandled Rejection',
    error: { message: String(reason) }
  });
});


// Middleware
app.use(express.json({ limit: '10mb' }));

// Also parse text/plain for sendBeacon requests
app.use(express.text({ limit: '10mb', type: 'text/plain' }));

// Enhanced CORS headers - must be before all routes
app.use((req, res, next) => {
  // Allow any origin
  res.header('Access-Control-Allow-Origin', '*');
  
  // Allow common methods
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  
  // Allow common headers including Content-Type
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization, Cache-Control');
  
  // Allow credentials (optional, but needed for some scenarios)
  res.header('Access-Control-Allow-Credentials', 'true');
  
  // Handle preflight OPTIONS requests immediately
  if (req.method === 'OPTIONS') {
    console.log(`[CORS] Preflight OPTIONS request for ${req.path}`);
    return res.status(204).end();
  }
  
  next();
});

// Additional CORS middleware specifically for /api/logs
app.use('/api/logs', (req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  
  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }
  next();
});

// Health check endpoint
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ============================================================================
// LOGGING API ENDPOINTS
// ============================================================================

/**
 * POST /api/logs
 * Receive log entries from client-side and write to server log file
 * Supports both JSON (from fetch) and text/plain (from sendBeacon)
 */
app.post('/api/logs', (req, res) => {
  try {
    let entries = null;
    
    // Handle text/plain content-type from sendBeacon
    if (typeof req.body === 'string') {
      try {
        const parsed = JSON.parse(req.body);
        entries = parsed.entries;
      } catch (e) {
        console.error('Failed to parse text body as JSON:', e);
        return res.status(400).json({ error: 'Invalid JSON in request body' });
      }
    } else {
      // Handle application/json
      entries = req.body.entries;
    }
    
    if (Array.isArray(entries)) {
      entries.forEach(entry => {
        serverLogger.log(entry);
      });
      res.json({ success: true, received: entries.length });
    } else if (req.body?.entry) {
      // Single entry format
      serverLogger.log(req.body.entry);
      res.json({ success: true, received: 1 });
    } else {
      res.status(400).json({ error: 'No log entries provided' });
    }
  } catch (error) {
    console.error('Logging error:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * GET /api/logs
 * Retrieve log entries from server log file
 * Query params: lines (default 100), level (0-4), since (ISO date)
 */
app.get('/api/logs', async (req, res) => {
  try {
    const options = {
      lines: parseInt(req.query.lines || '100', 10),
      level: req.query.level !== undefined ? parseInt(req.query.level, 10) : undefined,
      since: req.query.since ? new Date(req.query.since) : undefined
    };
    
    const logs = await serverLogger.getLogs(options);
    res.json({ logs });
  } catch (error) {
    console.error('Error retrieving logs:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * DELETE /api/logs
 * Clear all log files
 */
app.delete('/api/logs', (req, res) => {
  try {
    serverLogger.clearLogs();
    res.json({ success: true, message: 'Logs cleared' });
  } catch (error) {
    console.error('Error clearing logs:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * GET /api/logs/download
 * Download the current log file
 */
app.get('/api/logs/download', (req, res) => {
  try {
    if (!fs.existsSync(LOG_FILE)) {
      return res.status(404).json({ error: 'No log file found' });
    }
    
    res.download(LOG_FILE, `excel-ai-assistant-${new Date().toISOString().split('T')[0]}.log`);
  } catch (error) {
    console.error('Error downloading logs:', error);
    res.status(500).json({ error: error.message });
  }
});

/**
 * GET /api/logs/stats
 * Get log file statistics
 */
app.get('/api/logs/stats', (req, res) => {
  try {
    const stats = {
      logDir: LOG_DIR,
      files: [],
      totalSize: 0
    };
    
    if (fs.existsSync(LOG_DIR)) {
      const files = fs.readdirSync(LOG_DIR);
      files.forEach(file => {
        const filePath = path.join(LOG_DIR, file);
        const fileStats = fs.statSync(filePath);
        stats.files.push({
          name: file,
          size: fileStats.size,
          sizeMB: (fileStats.size / (1024 * 1024)).toFixed(2),
          modified: fileStats.mtime
        });
        stats.totalSize += fileStats.size;
      });
    }
    
    stats.totalSizeMB = (stats.totalSize / (1024 * 1024)).toFixed(2);
    stats.config = {
      maxLogSizeMB: MAX_LOG_SIZE_MB,
      maxLogFiles: MAX_LOG_FILES
    };
    
    res.json(stats);
  } catch (error) {
    console.error('Error getting log stats:', error);
    res.status(500).json({ error: error.message });
  }
});

// Proxy endpoint for AI API calls (optional - helps with CORS)
app.post('/api/chat', async (req, res) => {
  try {
    const { message, settings } = req.body;
    
    const response = await fetch(`${settings.apiUrl}/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${settings.apiKey}`
      },
      body: JSON.stringify({
        model: settings.model,
        messages: [
          { role: 'system', content: 'You are an Excel AI assistant.' },
          { role: 'user', content: message }
        ],
        temperature: settings.temperature || 0.7,
        max_tokens: settings.maxTokens || 4000
      })
    });

    if (!response.ok) {
      const error = await response.json();
      return res.status(response.status).json({ error: error.error?.message || 'API Error' });
    }

    const data = await response.json();
    res.json(data);
  } catch (error) {
    console.error('API Proxy Error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Serve static files from dist folder (for production)
app.use(express.static(path.join(__dirname, '../dist')));

// Start server
app.listen(PORT, () => {
  console.log(`✓ Excel AI Assistant server running on port ${PORT}`);
  console.log(`✓ Health check: http://localhost:${PORT}/api/health`);
});
