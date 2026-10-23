/**
 * Task Input Validation Middleware (Practical 7 - Step 7 & Problem Definition)
 * Rejects malformed requests before they reach the controller or database.
 */

// Validation middleware for POST /tasks (Creation)
const validateCreateTask = (req, res, next) => {
  const { title, priority } = req.body;

  // Validation 1: Required title field
  if (!title || typeof title !== 'string' || title.trim().length === 0) {
    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: 'Task title is required and cannot be empty or blank whitespace.',
    });
  }

  // Validation 2: Optional priority enum check
  if (priority && !['low', 'medium', 'high'].includes(priority.toLowerCase())) {
    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: `Invalid priority '${priority}'. Allowed priorities are: 'low', 'medium', 'high'.`,
    });
  }

  next();
};

// Validation middleware for PUT /tasks/:id (Update)
const validateUpdateTask = (req, res, next) => {
  const { title, priority } = req.body;

  // If title is present in update payload, ensure it's not empty whitespace
  if (title !== undefined && (typeof title !== 'string' || title.trim().length === 0)) {
    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: 'Updated task title cannot be an empty or blank string.',
    });
  }

  // If priority is present in update payload, ensure valid enum
  if (priority && !['low', 'medium', 'high'].includes(priority.toLowerCase())) {
    return res.status(400).json({
      success: false,
      error: 'Validation Error',
      message: `Invalid priority '${priority}'. Allowed priorities are: 'low', 'medium', 'high'.`,
    });
  }

  next();
};

module.exports = {
  validateCreateTask,
  validateUpdateTask,
};
