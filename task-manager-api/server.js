const path = require('path');
require('dotenv').config({ path: path.join(__dirname, '.env') });
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');

const Task = require('./models/Task');
const User = require('./models/User');
const authMiddleware = require('./middleware/auth');
const { validateCreateTask, validateUpdateTask } = require('./middleware/validateTask');
const { cache, getCache, setCache, invalidateTaskCache, getDiagnostics } = require('./cache');
const taskEvents = require('./events');
const { getEventAuditLog } = require('./listeners');

const app = express();
const PORT = process.env.PORT || 5000;
const MONGO_URI = process.env.MONGO_URI || 'mongodb://127.0.0.1:27017/task_manager_db';
const JWT_SECRET = process.env.JWT_SECRET || 'charusat_awdf_jwt_secret_key_sem5_2026';
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '1h';

// Enable CORS for frontend integration
app.use(cors());

// ============================================================================
// 1. DATABASE CONNECTION
// ============================================================================
mongoose
  .connect(MONGO_URI)
  .then(() => {
    console.log('MongoDB connected successfully');
    console.log(`Database URI: ${MONGO_URI}`);
  })
  .catch((err) => {
    console.error('MongoDB connection error:', err.message);
  });

// ============================================================================
// 2. GLOBAL MIDDLEWARES
// ============================================================================

// Global Request Logger with High-Resolution Response Time Measurement
app.use((req, res, next) => {
  const start = process.hrtime();
  const timestamp = new Date().toISOString();

  res.on('finish', () => {
    const diff = process.hrtime(start);
    const timeMs = (diff[0] * 1e3 + diff[1] * 1e-6).toFixed(2);
    const cacheHeader = res.getHeader('X-Cache') || 'N/A';
    console.log(`[${timestamp}] ${req.method} ${req.url} -> ${res.statusCode} (${timeMs} ms) [Cache: ${cacheHeader}]`);
  });

  next();
});

// Enforce Content-Type: application/json on POST/PUT/PATCH (except debug utility endpoints)
const requireJsonContent = (req, res, next) => {
  const isUtilityEndpoint = ['/cache/clear', '/events/test-error'].includes(req.path);
  if (['POST', 'PUT', 'PATCH'].includes(req.method) && !isUtilityEndpoint) {
    const contentType = req.headers['content-type'];
    if (!contentType || !contentType.toLowerCase().includes('application/json')) {
      return res.status(415).json({
        success: false,
        error: 'Unsupported Media Type',
        message: 'Requests with payload (POST/PUT) must include Content-Type: application/json header',
      });
    }
  }
  next();
};
app.use(requireJsonContent);

// Body Parser
app.use(express.json());

// ============================================================================
// 3. ROOT & CACHE DIAGNOSTIC ENDPOINTS (Supplementary Problem 2)
// ============================================================================
app.get('/', (req, res) => {
  res.status(200).json({
    success: true,
    message: 'Task Manager REST API with Native EventEmitter Background Processing is running',
    student: 'Priyansh Pravinbhai Panchal',
    course: 'Advance Web Development Frameworks',
    practical: 10,
    backgroundProcessing: {
      engine: 'Node.js native events (EventEmitter)',
      events: ['task-created (async + 2s simulated worker)', 'task-deleted', 'error'],
      eventsLogEndpoint: 'GET /events/log',
      testErrorEndpoint: 'POST /events/test-error',
    },
    caching: {
      engine: 'node-cache',
      stdTTL: 60,
      statsEndpoint: 'GET /cache/stats',
    },
    database: mongoose.connection.readyState === 1 ? 'Connected to MongoDB' : 'Connecting...',
    endpoints: {
      auth: {
        register: 'POST /auth/register (or POST /register)',
        login: 'POST /auth/login (or POST /login)',
        me: 'GET /auth/me (or GET /me) [Protected]',
      },
      tasks: {
        getAllTasks: 'GET /tasks [Protected + Cached 60s]',
        getTaskById: 'GET /tasks/:id [Protected + Cached 60s]',
        createTask: 'POST /tasks [Protected + Validated + Emits task-created]',
        updateTask: 'PUT /tasks/:id [Protected + Validated + Cache Invalidation]',
        deleteTask: 'DELETE /tasks/:id [Protected + Emits task-deleted]',
      },
      events: {
        eventsLog: 'GET /events/log',
        debugEvents: 'GET /debug/events',
        testError: 'POST /events/test-error',
      },
      diagnostics: {
        cacheStats: 'GET /cache/stats',
        clearCache: 'POST /cache/clear',
      },
    },
  });
});

// Cache performance diagnostics endpoint (Supplementary Problem 2)
app.get(['/cache/stats', '/debug/cache'], (req, res) => {
  const stats = getDiagnostics();
  res.status(200).json({
    success: true,
    message: 'In-Memory Cache (node-cache) Performance Metrics',
    timestamp: new Date().toISOString(),
    data: stats,
  });
});

// Cache flush/clear endpoint (GET or POST)
app.all('/cache/clear', (req, res) => {
  cache.flushAll();
  res.status(200).json({
    success: true,
    message: 'All in-memory cache keys cleared successfully',
    timestamp: new Date().toISOString(),
  });
});

// EventEmitter Background Audit Log Endpoint (Practical 10)
app.get(['/events/log', '/debug/events'], (req, res) => {
  const logs = getEventAuditLog();
  res.status(200).json({
    success: true,
    message: 'Node.js native EventEmitter background processing audit logs',
    totalEvents: logs.length,
    events: logs,
  });
});

// Test error event emission (Supplementary Problem 2)
app.all('/events/test-error', (req, res) => {
  taskEvents.emit('error', new Error('Manual test error emitted for Practical 10 verification'));
  res.status(200).json({
    success: true,
    message: 'Manual error event emitted and caught gracefully by taskEvents error listener',
  });
});

// ============================================================================
// 4. AUTHENTICATION HANDLERS & ROUTES
// ============================================================================

const handleRegister = async (req, res, next) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        error: 'Validation Error',
        message: 'Both email and password are required for registration.',
      });
    }

    if (password.length < 6) {
      return res.status(400).json({
        success: false,
        error: 'Validation Error',
        message: 'Password must be at least 6 characters long.',
      });
    }

    const existingUser = await User.findOne({ email: email.toLowerCase().trim() });
    if (existingUser) {
      return res.status(409).json({
        success: false,
        error: 'Conflict',
        message: 'A user with this email address already exists.',
      });
    }

    const hashedPassword = await bcrypt.hash(password, 10);
    const user = await User.create({
      email: email.toLowerCase().trim(),
      password: hashedPassword,
    });

    res.status(201).json({
      success: true,
      message: 'User registered successfully',
      user: {
        id: user._id,
        email: user.email,
        createdAt: user.createdAt,
      },
    });
  } catch (err) {
    next(err);
  }
};

const handleLogin = async (req, res, next) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res.status(400).json({
        success: false,
        error: 'Validation Error',
        message: 'Both email and password are required for login.',
      });
    }

    const user = await User.findOne({ email: email.toLowerCase().trim() });
    if (!user) {
      return res.status(401).json({
        success: false,
        error: 'Invalid Credentials',
        message: 'Invalid email or password.',
      });
    }

    const isMatch = await bcrypt.compare(password, user.password);
    if (!isMatch) {
      return res.status(401).json({
        success: false,
        error: 'Invalid Credentials',
        message: 'Invalid email or password.',
      });
    }

    const token = jwt.sign(
      { id: user._id, email: user.email },
      JWT_SECRET,
      { expiresIn: JWT_EXPIRES_IN }
    );

    res.status(200).json({
      success: true,
      message: 'Login successful',
      token,
      user: {
        id: user._id,
        email: user.email,
      },
    });
  } catch (err) {
    next(err);
  }
};

const handleMe = async (req, res, next) => {
  try {
    const user = await User.findById(req.user.id).select('-password');
    if (!user) {
      return res.status(404).json({
        success: false,
        error: 'Not Found',
        message: 'User account not found.',
      });
    }

    res.status(200).json({
      success: true,
      message: 'Authenticated user profile retrieved',
      user: {
        id: user._id,
        email: user.email,
        createdAt: user.createdAt,
      },
    });
  } catch (err) {
    next(err);
  }
};

app.post('/register', handleRegister);
app.post('/auth/register', handleRegister);
app.post('/login', handleLogin);
app.post('/auth/login', handleLogin);
app.get('/me', authMiddleware, handleMe);
app.get('/auth/me', authMiddleware, handleMe);

// ============================================================================
// 5. CACHED & PROTECTED RESTful TASK CRUD ENDPOINTS (Practical 9 - Steps 3 & 4)
// ============================================================================

// 1. GET /tasks - Read all tasks with in-memory caching (HTTP 200, Protected, Cached 60s)
app.get('/tasks', authMiddleware, async (req, res, next) => {
  try {
    const { completed, priority, noCache } = req.query;

    // Build specific cache key
    const hasFilter = completed !== undefined || priority !== undefined;
    const cacheKey = hasFilter
      ? `tasks_${req.user.id}_${completed || 'all'}_${priority || 'all'}`
      : `all_tasks_${req.user.id}`;

    // Step 3: Check cache first (unless ?noCache=true is passed for testing)
    if (noCache !== 'true') {
      const cachedResponse = getCache(cacheKey);
      if (cachedResponse) {
        // Cache HIT: Return immediately from memory (sub-millisecond)
        res.setHeader('X-Cache', 'HIT');
        res.setHeader('X-Cache-Key', cacheKey);
        res.setHeader('X-Cache-TTL', '60');
        return res.status(200).json(cachedResponse);
      }
    }

    // Cache MISS: Query MongoDB
    res.setHeader('X-Cache', 'MISS');
    res.setHeader('X-Cache-Key', cacheKey);

    const filter = {};
    if (completed !== undefined) {
      filter.completed = completed === 'true';
    }
    if (priority) {
      filter.priority = priority.toLowerCase();
    }

    filter.$or = [
      { userId: req.user.id },
      { userId: null },
      { userId: { $exists: false } },
    ];

    const tasks = await Task.find(filter).sort({ createdAt: -1 });

    const responsePayload = {
      success: true,
      count: tasks.length,
      data: tasks,
      source: 'database',
    };

    // Store in node-cache with 60-second TTL (Step 3)
    setCache(cacheKey, { ...responsePayload, source: 'cache' }, 60);

    res.status(200).json(responsePayload);
  } catch (err) {
    next(err);
  }
});

// 2. GET /tasks/:id - Read task by ID with single-task caching (Supplementary Problem 1)
app.get('/tasks/:id', authMiddleware, async (req, res, next) => {
  try {
    const { noCache } = req.query;
    const cacheKey = `task_${req.params.id}`;

    // Check individual task cache
    if (noCache !== 'true') {
      const cachedTask = getCache(cacheKey);
      if (cachedTask) {
        res.setHeader('X-Cache', 'HIT');
        res.setHeader('X-Cache-Key', cacheKey);
        return res.status(200).json(cachedTask);
      }
    }

    // Cache MISS: Query database
    res.setHeader('X-Cache', 'MISS');
    res.setHeader('X-Cache-Key', cacheKey);

    const task = await Task.findById(req.params.id);

    if (!task) {
      return res.status(404).json({
        success: false,
        error: 'Task Not Found',
        message: `No task found with ID: ${req.params.id}`,
      });
    }

    const responsePayload = {
      success: true,
      data: task,
      source: 'database',
    };

    // Store single task in cache with 60-second TTL
    setCache(cacheKey, { ...responsePayload, source: 'cache' }, 60);

    res.status(200).json(responsePayload);
  } catch (err) {
    next(err);
  }
});

// 3. POST /tasks - Create task + invalidate cache + emit async task-created event (Practical 10 - Step 3)
app.post('/tasks', authMiddleware, validateCreateTask, async (req, res, next) => {
  try {
    const taskData = {
      ...req.body,
      userId: req.user.id,
    };

    const task = await Task.create(taskData);

    // Invalidate cache
    invalidateTaskCache();

    // Step 3 & Problem Definition: Log exact timestamp at API response point
    const apiResponseTime = new Date().toISOString();
    console.log(`[API Response] Task creation responded HTTP 201 at ${apiResponseTime}`);

    // Return immediate HTTP 201 response to client BEFORE background worker finishes
    res.status(201).json({
      success: true,
      message: 'Task created successfully in MongoDB (Background processing initiated)',
      apiResponseTime,
      data: task,
    });

    // Step 3: Emit task-created event handled asynchronously by background listener
    setImmediate(() => {
      taskEvents.emit('task-created', {
        ...task.toObject(),
        userEmail: req.user.email,
        apiResponseTime,
        simulateError: req.body.simulateError === true,
      });
    });
  } catch (err) {
    next(err);
  }
});

// 4. PUT /tasks/:id - Update task + invalidate cache (HTTP 200, Protected, Validated)
app.put('/tasks/:id', authMiddleware, validateUpdateTask, async (req, res, next) => {
  try {
    if (req.body.title) {
      req.body.title = req.body.title.trim();
    }

    const task = await Task.findByIdAndUpdate(
      req.params.id,
      req.body,
      { new: true, runValidators: true }
    );

    if (!task) {
      return res.status(404).json({
        success: false,
        error: 'Task Not Found',
        message: `No task found with ID: ${req.params.id}`,
      });
    }

    // Step 4 & Supplementary 1: Invalidate collection cache AND specific task cache
    invalidateTaskCache(req.params.id);

    res.status(200).json({
      success: true,
      message: 'Task updated successfully (Cache invalidated)',
      data: task,
    });
  } catch (err) {
    next(err);
  }
});

// 5. DELETE /tasks/:id - Delete task + invalidate cache + emit async task-deleted event (Supplementary Problem 1)
app.delete('/tasks/:id', authMiddleware, async (req, res, next) => {
  try {
    const task = await Task.findByIdAndDelete(req.params.id);

    if (!task) {
      return res.status(404).json({
        success: false,
        error: 'Task Not Found',
        message: `No task found with ID: ${req.params.id}`,
      });
    }

    // Invalidate collection cache AND specific task cache
    invalidateTaskCache(req.params.id);

    const apiResponseTime = new Date().toISOString();
    console.log(`[API Response: DELETE] Task deletion responded HTTP 200 at ${apiResponseTime}`);

    res.status(200).json({
      success: true,
      message: 'Task deleted successfully from MongoDB (Background cleanup initiated)',
      apiResponseTime,
      data: task,
    });

    // Supplementary Problem 1: Emit task-deleted event handled asynchronously
    setImmediate(() => {
      taskEvents.emit('task-deleted', {
        id: task._id,
        title: task.title,
        userEmail: req.user.email,
        apiResponseTime,
      });
    });
  } catch (err) {
    next(err);
  }
});

// ============================================================================
// 5.8 AI-POWERED TASK ENHANCEMENT ROUTE (Practical 13)
// ============================================================================
const { generateTaskDescription } = require('./aiService');

const handleAiGenerate = async (req, res, next) => {
  try {
    const { title } = req.body;
    if (!title || typeof title !== 'string' || !title.trim()) {
      return res.status(400).json({
        success: false,
        error: 'Validation Error',
        message: 'A non-empty task title is required to generate an AI suggestion.',
      });
    }

    // Supplementary Problem 2: Rate limit tracking per client/IP (1 request / min)
    const clientId = (req.user && req.user.id) || req.ip || 'anonymous-client';
    const result = await generateTaskDescription(title.trim(), clientId);

    if (result.error === 'RateLimitExceeded') {
      return res.status(429).json({
        success: false,
        error: 'Too Many Requests',
        message: result.message,
        waitSeconds: result.waitSeconds,
      });
    }

    // Step 3 & Step 5: Deliver formatted suggestion, priority, and fallback metadata
    res.status(200).json({
      success: true,
      description: result.description,
      priority: result.priority,
      fallback: result.fallback,
      provider: result.provider,
      disclaimer: result.disclaimer,
      data: result,
    });
  } catch (err) {
    // Step 6: Graceful error handling - core app continues functioning
    console.error('[AI Endpoint Error]', err.message);
    res.status(200).json({
      success: true,
      description: `Review requirements and complete: ${req.body.title || 'Untitled'}`,
      priority: 'medium',
      fallback: true,
      message: 'AI service encountered an internal error. Fallback applied.',
    });
  }
};

app.post('/api/ai/generate-description', handleAiGenerate);
app.post('/ai/generate-description', handleAiGenerate);

// ============================================================================
// 6. 404 NOT FOUND HANDLER FOR UNDEFINED ROUTES
// ============================================================================
app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Route Not Found',
    message: `The requested endpoint '${req.method} ${req.originalUrl}' does not exist on this server.`,
  });
});

// ============================================================================
// 7. GLOBAL ERROR HANDLING MIDDLEWARE
// ============================================================================
app.use((err, req, res, next) => {
  if (err.code === 11000) {
    const field = Object.keys(err.keyValue || {})[0] || 'field';
    return res.status(409).json({
      success: false,
      error: 'Duplicate Key Error',
      message: `A record with this ${field} already exists.`,
    });
  }

  if (err.name === 'ValidationError') {
    const errors = Object.keys(err.errors).map((key) => ({
      field: key,
      message: err.errors[key].message,
    }));

    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: 'One or more fields failed Mongoose schema validation',
      details: errors,
    });
  }

  if (err.name === 'CastError') {
    return res.status(400).json({
      success: false,
      error: 'Invalid ID Format',
      message: `Invalid MongoDB ObjectId: '${err.value}'`,
    });
  }

  console.error('[Global Error Handler]:', err.message);
  res.status(500).json({
    success: false,
    error: 'Internal Server Error',
    message: err.message || 'Something went wrong on the server',
  });
});

// ============================================================================
// START SERVER (Only when executed directly, not during Jest testing)
// ============================================================================
if (process.env.NODE_ENV !== 'test') {
  app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
    console.log(`API Base URL: http://localhost:${PORT}`);
  });
}

module.exports = app;
