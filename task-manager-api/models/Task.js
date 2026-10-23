const mongoose = require('mongoose');

// Task Schema definition (Practicals 5 & 7)
const taskSchema = new mongoose.Schema({
  title: {
    type: String,
    required: [true, 'Task title is required'],
  },
  description: {
    type: String,
    default: '',
  },
  completed: {
    type: Boolean,
    default: false,
  },
  createdAt: {
    type: Date,
    default: Date.now,
  },
  // Supplementary Problem 1 (Practical 5): Priority field restricted to enum
  priority: {
    type: String,
    enum: {
      values: ['low', 'medium', 'high'],
      message: '{VALUE} is not a valid priority. Allowed priorities: low, medium, high',
    },
    default: 'medium',
  },
  // User association (Practical 7)
  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'User',
    default: null,
  },
});

// Supplementary Problem 2 (Practical 5): Pre-save hook that automatically trims whitespace from title
taskSchema.pre('save', function () {
  if (this.title) {
    this.title = this.title.trim();
  }
});

module.exports = mongoose.model('Task', taskSchema);
