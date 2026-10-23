import { useState, useEffect, lazy, Suspense } from 'react';
import {
  getTasks,
  createTask,
  updateTask,
  deleteTask,
  getMe,
  logoutUser,
  getCacheStats,
  getEventsLog,
  getStoredToken,
  getStoredUser,
} from '../services/api';
import Spinner from '../components/Spinner';
import ErrorMessage from '../components/ErrorMessage';
import ConfirmModal from '../components/ConfirmModal';
import Toast from '../components/Toast';
import AuthCard from '../components/AuthCard';
import DelayedFallback from '../components/DelayedFallback';

// Supplementary Problem 1 (Practical 8): Lazy load heavy third-party Chart.js component on demand
const TaskAnalyticsChart = lazy(() => import('../components/TaskAnalyticsChart'));

function Projects() {
  // Authentication states (Practical 7)
  const [token, setToken] = useState(() => getStoredToken());
  const [currentUser, setCurrentUser] = useState(() => getStoredUser());
  const [profileModal, setProfileModal] = useState(null);

  // In-memory cache diagnostics state (Practical 9 Supplementary 2)
  const [cacheModal, setCacheModal] = useState(null);

  // Background EventEmitter logs state (Practical 10)
  const [eventsModal, setEventsModal] = useState(null);

  // Lazy analytics chart toggle state (Practical 8 Supplementary 1)
  const [showAnalytics, setShowAnalytics] = useState(false);

  // State for task data, loading, and initial fetch error
  const [tasks, setTasks] = useState([]);
  const [loading, setLoading] = useState(false);
  const [fetchError, setFetchError] = useState(null);

  // Form states for creating a new task
  const [newTitle, setNewTitle] = useState('');
  const [newDesc, setNewDesc] = useState('');
  const [newPriority, setNewPriority] = useState('medium');
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Filter & Search states
  const [searchQuery, setSearchQuery] = useState('');
  const [filterPriority, setFilterPriority] = useState('all');

  // Supplementary Problem 2 (Practical 6): Delete confirmation modal state
  const [taskToDelete, setTaskToDelete] = useState(null);
  const [isDeleting, setIsDeleting] = useState(false);

  // Toast notification state
  const [toast, setToast] = useState({ message: '', type: 'success' });

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
  };

  // Reload trigger for manual re-fetching
  const [reloadTrigger, setReloadTrigger] = useState(0);

  // Listen for global auth events (Supplementary Problems 2 & 3)
  useEffect(() => {
    const handleAuthExpired = (e) => {
      setToken(null);
      setCurrentUser(null);
      setTasks([]);
      setLoading(false);
      showToast(e.detail || 'Session expired (401). Please log in again.', 'error');
    };

    const handleAuthLogout = () => {
      setToken(null);
      setCurrentUser(null);
      setTasks([]);
      setLoading(false);
      showToast('Logged out successfully. JWT cleared from client storage.', 'success');
    };

    window.addEventListener('auth:expired', handleAuthExpired);
    window.addEventListener('auth:logout', handleAuthLogout);

    return () => {
      window.removeEventListener('auth:expired', handleAuthExpired);
      window.removeEventListener('auth:logout', handleAuthLogout);
    };
  }, []);

  // Fetch tasks when user is authenticated (Practical 7 - Step 6)
  useEffect(() => {
    if (!token) return;

    let isCancelled = false;

    getTasks()
      .then((response) => {
        if (!isCancelled && response && response.data) {
          setTasks(response.data);
          setFetchError(null);
        }
      })
      .catch((err) => {
        if (!isCancelled) {
          setFetchError(
            err.message || 'Unable to connect to Express backend. Please ensure server is running.'
          );
        }
      })
      .finally(() => {
        if (!isCancelled) {
          setLoading(false);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [token, reloadTrigger]);

  const loadTasks = () => {
    setLoading(true);
    setFetchError(null);
    setReloadTrigger((prev) => prev + 1);
  };

  // Handle successful login/registration
  const handleAuthSuccess = (user, jwtToken) => {
    setToken(jwtToken);
    setCurrentUser(user);
    loadTasks();
  };

  // Handle logout (Supplementary Problem 3)
  const handleLogout = () => {
    logoutUser();
  };

  // Verify /me endpoint (Supplementary Problem 1)
  const handleVerifyProfile = async () => {
    try {
      const data = await getMe();
      setProfileModal(data.user);
      showToast('Profile successfully verified using decoded JWT token!', 'success');
    } catch (err) {
      showToast(err.message || 'Failed to verify profile', 'error');
    }
  };

  // Inspect /cache/stats endpoint (Practical 9 Supplementary 2)
  const handleInspectCache = async () => {
    try {
      const res = await getCacheStats();
      setCacheModal(res.data);
      showToast('Live in-memory cache diagnostics retrieved!', 'success');
    } catch (err) {
      showToast(err.message || 'Failed to fetch cache stats', 'error');
    }
  };

  // Inspect /events/log endpoint (Practical 10)
  const handleInspectEvents = async () => {
    try {
      const res = await getEventsLog();
      setEventsModal(res.events || []);
      showToast('Live EventEmitter background logs retrieved!', 'success');
    } catch (err) {
      showToast(err.message || 'Failed to fetch event logs', 'error');
    }
  };

  // Task Creation with server-side validation and Optimistic UI update
  const handleCreateTask = async (e) => {
    e.preventDefault();
    if (!newTitle.trim()) {
      showToast('Validation Error: Task title is required.', 'error');
      return;
    }

    const trimmedTitle = newTitle.trim();
    const taskPayload = {
      title: trimmedTitle,
      description: newDesc.trim(),
      priority: newPriority,
      completed: false,
    };

    // Optimistic UI update
    const tempId = `optimistic-${Date.now()}`;
    const optimisticTask = {
      _id: tempId,
      ...taskPayload,
      createdAt: new Date().toISOString(),
      isOptimistic: true,
    };

    setTasks((prev) => [optimisticTask, ...prev]);
    setNewTitle('');
    setNewDesc('');
    setNewPriority('medium');
    setIsSubmitting(true);

    try {
      const result = await createTask(taskPayload);
      if (result && result.data) {
        setTasks((prev) =>
          prev.map((t) => (t._id === tempId ? result.data : t))
        );
        showToast('Task created and saved to MongoDB!', 'success');
      }
    } catch (err) {
      // Rollback optimistic update
      setTasks((prev) => prev.filter((t) => t._id !== tempId));
      showToast(err.message || 'Server rejected task creation.', 'error');
    } finally {
      setIsSubmitting(false);
    }
  };

  // Task Toggle Completion
  const handleToggleCompleted = async (task) => {
    const originalStatus = task.completed;
    const updatedStatus = !originalStatus;

    setTasks((prev) =>
      prev.map((t) => (t._id === task._id ? { ...t, completed: updatedStatus } : t))
    );

    try {
      await updateTask(task._id, { completed: updatedStatus });
      showToast(
        `Task marked as ${updatedStatus ? 'completed' : 'pending'}.`,
        'success'
      );
    } catch (err) {
      // Rollback on failure
      setTasks((prev) =>
        prev.map((t) => (t._id === task._id ? { ...t, completed: originalStatus } : t))
      );
      showToast(err.message || 'Failed to update task status.', 'error');
    }
  };

  // Task Deletion Flow
  const triggerDelete = (task) => {
    setTaskToDelete(task);
  };

  const handleConfirmDelete = async () => {
    if (!taskToDelete) return;
    setIsDeleting(true);

    try {
      await deleteTask(taskToDelete._id);
      setTasks((prev) => prev.filter((t) => t._id !== taskToDelete._id));
      showToast(`Task "${taskToDelete.title}" deleted from MongoDB.`, 'success');
    } catch (err) {
      showToast(err.message || 'Failed to delete task.', 'error');
    } finally {
      setIsDeleting(false);
      setTaskToDelete(null);
    }
  };

  const handleCancelDelete = () => {
    setTaskToDelete(null);
  };

  // Filter tasks based on search and priority
  const filteredTasks = tasks.filter((task) => {
    const matchesSearch =
      task.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      (task.description && task.description.toLowerCase().includes(searchQuery.toLowerCase()));
    const matchesPriority =
      filterPriority === 'all' || task.priority === filterPriority;
    return matchesSearch && matchesPriority;
  });

  const getPriorityColor = (priority) => {
    switch (priority) {
      case 'high':
        return { bg: '#fee2e2', text: '#991b1b', border: '#f87171' };
      case 'low':
        return { bg: '#e0f2fe', text: '#0369a1', border: '#7dd3fc' };
      default:
        return { bg: '#fef3c7', text: '#92400e', border: '#fcd34d' };
    }
  };

  return (
    <div className="projects-container" style={{ maxWidth: '960px', margin: '0 auto', padding: '1rem' }}>
      {/* Toast Notification */}
      {toast.message && (
        <Toast
          message={toast.message}
          type={toast.type}
          onClose={() => setToast({ message: '', type: 'success' })}
        />
      )}

      {/* Delete Confirmation Modal */}
      {taskToDelete && (
        <ConfirmModal
          isOpen={true}
          title="Confirm Task Deletion"
          message={`Are you sure you want to permanently delete "${taskToDelete.title}" from MongoDB?`}
          confirmLabel="Delete Task"
          isProcessing={isDeleting}
          onConfirm={handleConfirmDelete}
          onCancel={handleCancelDelete}
        />
      )}

      {/* Profile Modal for /me verification (Supplementary Problem 1) */}
      {profileModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
        }}>
          <div style={{
            backgroundColor: '#ffffff',
            borderRadius: '12px',
            padding: '24px',
            maxWidth: '450px',
            width: '90%',
            boxShadow: '0 10px 25px rgba(0,0,0,0.2)',
            color: '#1e293b',
          }}>
            <h3 style={{ marginTop: 0, color: '#1d4ed8' }}>👤 Decoded JWT User Profile (/me)</h3>
            <p style={{ fontSize: '0.9rem', color: '#64748b' }}>
              Retrieved directly from the protected <code>GET /auth/me</code> route using the current Bearer token.
            </p>
            <table style={{ width: '100%', fontSize: '0.875rem', marginBottom: '1.5rem', borderCollapse: 'collapse' }}>
              <tbody>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>User ID:</td>
                  <td style={{ padding: '8px 0', fontFamily: 'monospace' }}>{profileModal.id}</td>
                </tr>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Email:</td>
                  <td style={{ padding: '8px 0' }}>{profileModal.email}</td>
                </tr>
                <tr>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Created At:</td>
                  <td style={{ padding: '8px 0' }}>{new Date(profileModal.createdAt).toLocaleString()}</td>
                </tr>
              </tbody>
            </table>
            <button
              type="button"
              onClick={() => setProfileModal(null)}
              style={{
                width: '100%',
                padding: '10px',
                backgroundColor: '#1d4ed8',
                color: '#ffffff',
                border: 'none',
                borderRadius: '6px',
                fontWeight: 600,
                cursor: 'pointer',
              }}
            >
              Close
            </button>
          </div>
        </div>
      )}

      {/* Cache Diagnostics Modal (Practical 9 Supplementary 2) */}
      {cacheModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
        }}>
          <div style={{
            backgroundColor: '#ffffff',
            borderRadius: '12px',
            padding: '24px',
            maxWidth: '450px',
            width: '90%',
            boxShadow: '0 10px 25px rgba(0,0,0,0.2)',
            color: '#1e293b',
          }}>
            <h3 style={{ marginTop: 0, color: '#166534' }}>⚡ NodeCache Diagnostics (/cache/stats)</h3>
            <p style={{ fontSize: '0.9rem', color: '#64748b' }}>
              In-memory server cache performance metrics and hit-rate monitoring (Practical 9).
            </p>
            <table style={{ width: '100%', fontSize: '0.875rem', marginBottom: '1.5rem', borderCollapse: 'collapse' }}>
              <tbody>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Standard TTL:</td>
                  <td style={{ padding: '8px 0' }}>{cacheModal.stdTTL} seconds</td>
                </tr>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Cache Hits:</td>
                  <td style={{ padding: '8px 0', color: '#166534', fontWeight: 600 }}>{cacheModal.hits}</td>
                </tr>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Cache Misses:</td>
                  <td style={{ padding: '8px 0', color: '#dc2626' }}>{cacheModal.misses}</td>
                </tr>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Hit Rate:</td>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>{cacheModal.hitRate}</td>
                </tr>
                <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Invalidations:</td>
                  <td style={{ padding: '8px 0' }}>{cacheModal.invalidations}</td>
                </tr>
                <tr>
                  <td style={{ padding: '8px 0', fontWeight: 600 }}>Active Cached Keys:</td>
                  <td style={{ padding: '8px 0' }}>{cacheModal.keyCount}</td>
                </tr>
              </tbody>
            </table>
            <button
              type="button"
              onClick={() => setCacheModal(null)}
              style={{
                width: '100%',
                padding: '10px',
                backgroundColor: '#166534',
                color: '#ffffff',
                border: 'none',
                borderRadius: '6px',
                fontWeight: 600,
                cursor: 'pointer',
              }}
            >
              Close Diagnostics
            </button>
          </div>
        </div>
      )}

      {/* Background EventEmitter Modal (Practical 10) */}
      {eventsModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
        }}>
          <div style={{
            backgroundColor: '#ffffff',
            borderRadius: '12px',
            padding: '24px',
            maxWidth: '560px',
            width: '92%',
            maxHeight: '80vh',
            overflowY: 'auto',
            boxShadow: '0 10px 25px rgba(0,0,0,0.2)',
            color: '#1e293b',
          }}>
            <h3 style={{ marginTop: 0, color: '#92400e' }}>🔔 Background Event Logs (EventEmitter)</h3>
            <p style={{ fontSize: '0.875rem', color: '#64748b' }}>
              Shows asynchronous execution of background listeners decoupled from the main HTTP response (Practical 10).
            </p>
            {eventsModal.length === 0 ? (
              <p style={{ color: '#94a3b8', fontStyle: 'italic' }}>No background events logged yet. Create or delete a task to trigger background events.</p>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginBottom: '1.5rem' }}>
                {eventsModal.slice(0, 10).map((ev, i) => (
                  <div key={i} style={{
                    padding: '10px',
                    borderRadius: '6px',
                    backgroundColor: ev.event === 'error' ? '#fef2f2' : '#f8fafc',
                    border: '1px solid #e2e8f0',
                    fontSize: '0.8rem',
                  }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontWeight: 600 }}>
                      <span style={{ color: ev.event === 'error' ? '#dc2626' : '#1d4ed8' }}>
                        ● {ev.event}
                      </span>
                      <span style={{ color: '#166534' }}>{ev.status}</span>
                    </div>
                    {ev.title && <div>Title: <strong>{ev.title}</strong></div>}
                    {ev.apiDispatchedAt && <div>API Responded: {new Date(ev.apiDispatchedAt).toLocaleTimeString()}</div>}
                    {ev.backgroundCompletedAt && <div>Worker Completed: {new Date(ev.backgroundCompletedAt).toLocaleTimeString()} ({ev.simulatedDelay || 'async'})</div>}
                  </div>
                ))}
              </div>
            )}
            <button
              type="button"
              onClick={() => setEventsModal(null)}
              style={{
                width: '100%',
                padding: '10px',
                backgroundColor: '#92400e',
                color: '#ffffff',
                border: 'none',
                borderRadius: '6px',
                fontWeight: 600,
                cursor: 'pointer',
              }}
            >
              Close Event Logs
            </button>
          </div>
        </div>
      )}

      {/* Main Page Header */}
      <header className="projects-header" style={{ textAlign: 'center', marginBottom: '2rem' }}>
        <h1 style={{ fontSize: '2rem', marginBottom: '0.5rem' }}>
          Full-Stack Task Manager (Practical 7)
        </h1>
        <p style={{ color: 'var(--text-muted, #64748b)' }}>
          Node.js, Express, MongoDB, Mongoose, and JWT-Based Authentication Pipeline
        </p>
      </header>

      {/* Render Authentication Card if not authenticated */}
      {!token ? (
        <AuthCard onAuthSuccess={handleAuthSuccess} showToast={showToast} />
      ) : (
        <>
          {/* Authenticated User Session Bar */}
          <div style={{
            display: 'flex',
            flexWrap: 'wrap',
            justifyContent: 'space-between',
            alignItems: 'center',
            backgroundColor: 'var(--card-bg, #f8fafc)',
            padding: '12px 18px',
            borderRadius: '8px',
            marginBottom: '1.5rem',
            border: '1px solid var(--border-color, #e2e8f0)',
            gap: '12px',
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '1.2rem' }}>👤</span>
              <div>
                <strong style={{ fontSize: '0.95rem' }}>{currentUser?.email || 'Authenticated User'}</strong>
                <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginTop: '2px' }}>
                  <span style={{
                    fontSize: '0.75rem',
                    backgroundColor: '#dcfce7',
                    color: '#166534',
                    padding: '2px 8px',
                    borderRadius: '12px',
                    fontWeight: 600,
                  }}>
                    ● JWT Active (1h)
                  </span>
                </div>
              </div>
            </div>

            <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
              <button
                type="button"
                onClick={() => setShowAnalytics((prev) => !prev)}
                style={{
                  padding: '6px 12px',
                  backgroundColor: showAnalytics ? '#dbeafe' : '#f8fafc',
                  color: '#1d4ed8',
                  border: '1px solid #93c5fd',
                  borderRadius: '6px',
                  fontSize: '0.85rem',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                {showAnalytics ? 'Hide Analytics' : '📊 Task Analytics (Lazy Chart)'}
              </button>
              <button
                type="button"
                onClick={handleInspectCache}
                style={{
                  padding: '6px 12px',
                  backgroundColor: '#f0fdf4',
                  color: '#166534',
                  border: '1px solid #bbf7d0',
                  borderRadius: '6px',
                  fontSize: '0.85rem',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                ⚡ Cache Stats (/cache)
              </button>
              <button
                type="button"
                onClick={handleInspectEvents}
                style={{
                  padding: '6px 12px',
                  backgroundColor: '#fef3c7',
                  color: '#92400e',
                  border: '1px solid #fde68a',
                  borderRadius: '6px',
                  fontSize: '0.85rem',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                🔔 Event Logs (/events)
              </button>
              <button
                type="button"
                onClick={handleVerifyProfile}
                style={{
                  padding: '6px 12px',
                  backgroundColor: '#eff6ff',
                  color: '#1d4ed8',
                  border: '1px solid #bfdbfe',
                  borderRadius: '6px',
                  fontSize: '0.85rem',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                Inspect Token (/me)
              </button>
              <button
                type="button"
                onClick={handleLogout}
                style={{
                  padding: '6px 12px',
                  backgroundColor: '#fee2e2',
                  color: '#991b1b',
                  border: '1px solid #fecaca',
                  borderRadius: '6px',
                  fontSize: '0.85rem',
                  fontWeight: 600,
                  cursor: 'pointer',
                }}
              >
                Logout
              </button>
            </div>
          </div>

          {/* Lazy-Loaded Heavy Analytics Component (Supplementary Problem 1 & 2) */}
          {showAnalytics && (
            <Suspense fallback={<DelayedFallback delay={150} label="Loading Chart.js chunk..." />}>
              <TaskAnalyticsChart tasks={tasks} />
            </Suspense>
          )}

          {/* Task Creation Form (with Server-side Validation) */}
          <section
            className="task-form-card"
            style={{
              backgroundColor: 'var(--card-bg, #ffffff)',
              borderRadius: '8px',
              padding: '1.5rem',
              boxShadow: '0 2px 8px rgba(0,0,0,0.06)',
              marginBottom: '2rem',
              border: '1px solid var(--border-color, #e2e8f0)',
            }}
          >
            <h2 style={{ fontSize: '1.25rem', marginTop: 0, marginBottom: '1rem' }}>
              Add New Task (Protected & Validated)
            </h2>
            <form onSubmit={handleCreateTask} style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
              <div>
                <input
                  type="text"
                  placeholder="Task title (required)..."
                  value={newTitle}
                  onChange={(e) => setNewTitle(e.target.value)}
                  style={{
                    width: '100%',
                    padding: '10px 14px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    fontSize: '1rem',
                    boxSizing: 'border-box',
                  }}
                />
              </div>
              <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap' }}>
                <input
                  type="text"
                  placeholder="Task description (optional)..."
                  value={newDesc}
                  onChange={(e) => setNewDesc(e.target.value)}
                  style={{
                    flex: '1 1 200px',
                    padding: '10px 14px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    fontSize: '0.95rem',
                  }}
                />
                <select
                  value={newPriority}
                  onChange={(e) => setNewPriority(e.target.value)}
                  style={{
                    padding: '10px 14px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    fontSize: '0.95rem',
                    backgroundColor: 'inherit',
                    color: 'inherit',
                  }}
                >
                  <option value="low">Low Priority</option>
                  <option value="medium">Medium Priority</option>
                  <option value="high">High Priority</option>
                </select>
                <button
                  type="submit"
                  disabled={isSubmitting}
                  style={{
                    backgroundColor: '#1d4ed8',
                    color: '#ffffff',
                    border: 'none',
                    borderRadius: '6px',
                    padding: '10px 20px',
                    fontWeight: '600',
                    cursor: isSubmitting ? 'not-allowed' : 'pointer',
                    opacity: isSubmitting ? 0.7 : 1,
                  }}
                >
                  {isSubmitting ? 'Saving...' : 'Add Task'}
                </button>
              </div>
            </form>
          </section>

          {/* Controls: Search, Filter, Refresh */}
          <div
            className="task-controls"
            style={{
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              flexWrap: 'wrap',
              gap: '1rem',
              marginBottom: '1.5rem',
            }}
          >
            <div style={{ display: 'flex', gap: '0.5rem', flex: '1 1 250px' }}>
              <input
                type="text"
                placeholder="Search tasks..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{
                  width: '100%',
                  padding: '8px 12px',
                  borderRadius: '6px',
                  border: '1px solid #cbd5e1',
                  fontSize: '0.9rem',
                }}
              />
            </div>

            <div style={{ display: 'flex', gap: '0.75rem', alignItems: 'center' }}>
              <select
                value={filterPriority}
                onChange={(e) => setFilterPriority(e.target.value)}
                style={{
                  padding: '8px 12px',
                  borderRadius: '6px',
                  border: '1px solid #cbd5e1',
                  fontSize: '0.9rem',
                  backgroundColor: 'inherit',
                  color: 'inherit',
                }}
              >
                <option value="all">All Priorities</option>
                <option value="high">High</option>
                <option value="medium">Medium</option>
                <option value="low">Low</option>
              </select>

              <button
                type="button"
                onClick={loadTasks}
                style={{
                  backgroundColor: '#f1f5f9',
                  color: '#334155',
                  border: '1px solid #cbd5e1',
                  borderRadius: '6px',
                  padding: '8px 14px',
                  cursor: 'pointer',
                  fontSize: '0.9rem',
                  fontWeight: '500',
                }}
              >
                ↻ Refresh
              </button>
            </div>
          </div>

          {/* Task List / States */}
          {loading ? (
            <Spinner label="Loading tasks from MongoDB..." />
          ) : fetchError ? (
            <ErrorMessage message={fetchError} onRetry={loadTasks} />
          ) : filteredTasks.length === 0 ? (
            <div
              style={{
                textAlign: 'center',
                padding: '3rem 1rem',
                backgroundColor: 'var(--card-bg, #ffffff)',
                borderRadius: '8px',
                border: '1px dashed #cbd5e1',
              }}
            >
              <p style={{ margin: '0 0 0.5rem', fontSize: '1.1rem', color: '#64748b' }}>
                No tasks found.
              </p>
              <p style={{ margin: 0, fontSize: '0.85rem', color: '#94a3b8' }}>
                Create a task above to persist it in MongoDB under your user account.
              </p>
            </div>
          ) : (
            <ul style={{ listStyle: 'none', padding: 0, margin: 0, display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
              {filteredTasks.map((task) => {
                const priorityStyle = getPriorityColor(task.priority);
                return (
                  <li
                    key={task._id}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      padding: '1rem 1.25rem',
                      backgroundColor: 'var(--card-bg, #ffffff)',
                      borderRadius: '8px',
                      border: '1px solid var(--border-color, #e2e8f0)',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
                      opacity: task.isOptimistic ? 0.6 : 1,
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', flex: 1 }}>
                      <input
                        type="checkbox"
                        checked={task.completed}
                        onChange={() => handleToggleCompleted(task)}
                        disabled={task.isOptimistic}
                        style={{ width: '18px', height: '18px', cursor: 'pointer' }}
                      />
                      <div style={{ flex: 1 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                          <span
                            style={{
                              fontSize: '1.05rem',
                              fontWeight: '600',
                              textDecoration: task.completed ? 'line-through' : 'none',
                              color: task.completed ? '#94a3b8' : 'inherit',
                            }}
                          >
                            {task.title}
                          </span>
                          <span
                            style={{
                              fontSize: '0.75rem',
                              padding: '2px 8px',
                              borderRadius: '12px',
                              backgroundColor: priorityStyle.bg,
                              color: priorityStyle.text,
                              border: `1px solid ${priorityStyle.border}`,
                              fontWeight: 'bold',
                              textTransform: 'uppercase',
                            }}
                          >
                            {task.priority || 'medium'}
                          </span>
                          {task.isOptimistic && (
                            <span style={{ fontSize: '0.75rem', color: '#1d4ed8', fontStyle: 'italic' }}>
                              (Saving...)
                            </span>
                          )}
                        </div>
                        {task.description && (
                          <p
                            style={{
                              margin: '0.25rem 0 0',
                              fontSize: '0.875rem',
                              color: 'var(--text-muted, #64748b)',
                            }}
                          >
                            {task.description}
                          </p>
                        )}
                        <span style={{ fontSize: '0.75rem', color: '#94a3b8' }}>
                          Created: {new Date(task.createdAt).toLocaleString()}
                        </span>
                      </div>
                    </div>

                    <div style={{ display: 'flex', gap: '0.5rem' }}>
                      <button
                        type="button"
                        onClick={() => triggerDelete(task)}
                        disabled={task.isOptimistic}
                        style={{
                          backgroundColor: '#fee2e2',
                          color: '#dc2626',
                          border: '1px solid #fecaca',
                          borderRadius: '6px',
                          padding: '6px 12px',
                          fontSize: '0.85rem',
                          cursor: 'pointer',
                          fontWeight: '500',
                        }}
                      >
                        Delete
                      </button>
                    </div>
                  </li>
                );
              })}
            </ul>
          )}
        </>
      )}
    </div>
  );
}

export default Projects;
