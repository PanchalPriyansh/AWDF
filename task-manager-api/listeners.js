const taskEvents = require('./events');

// In-memory background event log buffer for verification & diagnostics
const eventAuditLog = [];

/**
 * Register Event Listeners (Practical 10 - Step 2 & Supplementary Problems 1, 2, 3)
 */

// 1. Task Created Listener (Step 2 & Supplementary 3)
taskEvents.on('task-created', (task) => {
  const handlerStartTime = new Date().toISOString();
  console.log(`[Listener: task-created] Handler received event for: "${task.title}" at ${handlerStartTime}`);
  console.log(`[Notification Details] Title: "${task.title}" | User: ${task.userEmail || task.userId} | API Responded At: ${task.apiResponseTime || 'N/A'}`);

  // Supplementary Problem 3: Simulate heavy background processing (2-second asynchronous delay)
  // This demonstrates that the main HTTP request/response returned immediately without waiting for this 2000ms worker
  setTimeout(() => {
    try {
      // Simulate potential failure check (Supplementary 2)
      if (task.simulateError) {
        throw new Error('Simulated external email notification service timeout (504 Gateway)');
      }

      const handlerFinishTime = new Date().toISOString();
      console.log(`[Background Worker Finished (+2000ms)] Async email delivery complete for "${task.title}" at ${handlerFinishTime}`);

      eventAuditLog.unshift({
        event: 'task-created',
        taskId: task._id,
        title: task.title,
        user: task.userEmail || task.userId,
        apiDispatchedAt: task.apiResponseTime,
        handlerReceivedAt: handlerStartTime,
        backgroundCompletedAt: handlerFinishTime,
        simulatedDelay: '2000ms',
        status: 'SUCCESS',
      });

      // Keep last 50 events
      if (eventAuditLog.length > 50) eventAuditLog.pop();
    } catch (err) {
      // Forward unhandled exceptions to the 'error' listener (Supplementary 2)
      taskEvents.emit('error', err);
    }
  }, 2000);
});

// 2. Task Deleted Listener (Supplementary Problem 1)
taskEvents.on('task-deleted', (data) => {
  const timestamp = new Date().toISOString();
  console.log(`[Listener: task-deleted] Task "${data.title}" (ID: ${data.id}) was permanently deleted by ${data.userEmail || 'User'} at ${timestamp}`);
  console.log(`[Audit Notification] Cleanup event recorded in background without blocking DELETE API response.`);

  eventAuditLog.unshift({
    event: 'task-deleted',
    taskId: data.id,
    title: data.title,
    user: data.userEmail,
    apiDispatchedAt: data.apiResponseTime,
    backgroundCompletedAt: timestamp,
    status: 'SUCCESS',
  });

  if (eventAuditLog.length > 50) eventAuditLog.pop();
});

// 3. Error Event Listener (Supplementary Problem 2)
taskEvents.on('error', (err) => {
  const timestamp = new Date().toISOString();
  console.error(`[Event Error Listener] Handled background event processing error at ${timestamp}: ${err.message}`);

  eventAuditLog.unshift({
    event: 'error',
    errorMessage: err.message,
    timestamp,
    status: 'CAUGHT_ERROR',
  });

  if (eventAuditLog.length > 50) eventAuditLog.pop();
});

module.exports = {
  getEventAuditLog: () => eventAuditLog,
};
