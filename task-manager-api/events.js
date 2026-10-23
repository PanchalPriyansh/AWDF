const EventEmitter = require('events');

/**
 * Dedicated TaskEvents EventEmitter Module (Practical 10 - Step 1)
 * Built using Node.js native 'events' module without external dependencies.
 */
class TaskEvents extends EventEmitter {}

// Export singleton instance
const taskEvents = new TaskEvents();

module.exports = taskEvents;
