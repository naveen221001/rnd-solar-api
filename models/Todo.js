const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const Todo = sequelize.define('Todo', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  issue: {
    type: DataTypes.TEXT,
    allowNull: false,
    validate: {
      notEmpty: true,
      len: [1, 1000]
    }
  },
  responsibility: {
    type: DataTypes.TEXT,
    allowNull: false,
    comment: 'Comma-separated list of team member codes'
  },
  status: {
    type: DataTypes.ENUM('Pending', 'In Progress', 'Review', 'On Hold', 'Done'),
    defaultValue: 'Pending',
    allowNull: false
  },
  priority: {
    type: DataTypes.ENUM('High', 'Medium', 'Low'),
    defaultValue: 'Medium',
    allowNull: false
  },
  category: {
    type: DataTypes.ENUM('Testing', 'Documentation', 'Reporting', 'General', 'Process', 'Safety'),
    defaultValue: 'General',
    allowNull: false
  },
  dueDate: {
    type: DataTypes.DATE,
    allowNull: true
  },
  createdBy: {
    type: DataTypes.STRING,
    allowNull: false,
    comment: 'Email of the user who created this todo'
  }
}, {
  tableName: 'todos',
  timestamps: true, // Adds createdAt and updatedAt
  indexes: [
    {
      fields: ['status']
    },
    {
      fields: ['createdBy']
    },
    {
      fields: ['dueDate']
    }
  ]
});

module.exports = Todo;