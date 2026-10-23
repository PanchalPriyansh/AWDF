// Central API service connecting React to Express & MongoDB backend with JWT Auth (Practicals 6 & 7)
const BASE_URL = 'http://localhost:5000';

const TOKEN_KEY = 'task_manager_jwt';
const USER_KEY = 'task_manager_user';

/**
 * Helper to retrieve stored JWT token from localStorage
 */
export function getStoredToken() {
  return localStorage.getItem(TOKEN_KEY);
}

/**
 * Helper to retrieve stored User profile from localStorage
 */
export function getStoredUser() {
  const user = localStorage.getItem(USER_KEY);
  return user ? JSON.parse(user) : null;
}

/**
 * Returns HTTP headers including Content-Type and Authorization Bearer token
 */
export function getAuthHeaders() {
  const token = getStoredToken();
  const headers = {
    'Content-Type': 'application/json',
  };
  if (token) {
    headers['Authorization'] = `Bearer ${token}`;
  }
  return headers;
}

/**
 * Handles 401 response: clears tokens and dispatches an event for frontend redirection (Supplementary 2)
 */
function handleAuthError(data) {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
  window.dispatchEvent(new CustomEvent('auth:expired', { detail: data.message }));
  throw new Error(data.message || 'Unauthorized: Session expired. Please log in again.');
}

/**
 * Register a new user with hashed password (Practical 7 - Step 3)
 */
export async function registerUser(email, password) {
  const res = await fetch(`${BASE_URL}/auth/register`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.message || 'Registration failed');
  }
  return data;
}

/**
 * Log in existing user and receive JWT token (Practical 7 - Step 4)
 */
export async function loginUser(email, password) {
  const res = await fetch(`${BASE_URL}/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });

  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.message || 'Login failed');
  }

  // Store token and user details in client storage
  if (data.token) {
    localStorage.setItem(TOKEN_KEY, data.token);
  }
  if (data.user) {
    localStorage.setItem(USER_KEY, JSON.stringify(data.user));
  }

  return data;
}

/**
 * Fetch currently logged in user profile from /me endpoint (Supplementary Problem 1)
 */
export async function getMe() {
  const res = await fetch(`${BASE_URL}/auth/me`, {
    headers: getAuthHeaders(),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to fetch current user profile');
  }
  return data;
}

/**
 * Logout mechanism that clears token and user from client (Supplementary Problem 3)
 */
export function logoutUser() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
  window.dispatchEvent(new CustomEvent('auth:logout'));
}

/**
 * Fetch background EventEmitter audit logs (Practical 10)
 */
export async function getEventsLog() {
  const res = await fetch(`${BASE_URL}/events/log`);
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.message || 'Failed to fetch event audit logs');
  }
  return data;
}

/**
 * Fetch server-side NodeCache performance diagnostics (Practical 9 - Supplementary 2)
 */
export async function getCacheStats() {
  const res = await fetch(`${BASE_URL}/cache/stats`);
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.message || 'Failed to fetch cache diagnostics');
  }
  return data;
}

/**
 * Fetch all tasks from MongoDB for authenticated user (Practical 7 - Step 6)
 */
export async function getTasks(filter = {}) {
  const query = new URLSearchParams(filter).toString();
  const url = query ? `${BASE_URL}/tasks?${query}` : `${BASE_URL}/tasks`;

  const res = await fetch(url, {
    headers: getAuthHeaders(),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to fetch tasks from server');
  }
  return data;
}

/**
 * Fetch a single task by its MongoDB ObjectId (Protected)
 */
export async function getTaskById(id) {
  const res = await fetch(`${BASE_URL}/tasks/${id}`, {
    headers: getAuthHeaders(),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to fetch task');
  }
  return data;
}

/**
 * Create a new task in MongoDB with validation (Practical 7 - Step 7)
 */
export async function createTask(taskData) {
  const res = await fetch(`${BASE_URL}/tasks`, {
    method: 'POST',
    headers: getAuthHeaders(),
    body: JSON.stringify(taskData),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to create task');
  }
  return data;
}

/**
 * Update an existing task in MongoDB with validation (Protected + Validated)
 */
export async function updateTask(id, updateData) {
  const res = await fetch(`${BASE_URL}/tasks/${id}`, {
    method: 'PUT',
    headers: getAuthHeaders(),
    body: JSON.stringify(updateData),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to update task');
  }
  return data;
}

/**
 * Delete a task from MongoDB (Protected)
 */
export async function deleteTask(id) {
  const res = await fetch(`${BASE_URL}/tasks/${id}`, {
    method: 'DELETE',
    headers: getAuthHeaders(),
  });

  const data = await res.json();
  if (res.status === 401) {
    handleAuthError(data);
  }
  if (!res.ok) {
    throw new Error(data.message || 'Failed to delete task');
  }
  return data;
}
