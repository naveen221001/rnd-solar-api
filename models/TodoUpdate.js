const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const TodoUpdate = sequelize.define('TodoUpdate', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  todoId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: {
      model: 'todos',
      key: 'id'
    },
    onUpdate: 'CASCADE',
    onDelete: 'CASCADE'
  },
  status: {
    type: DataTypes.ENUM('Pending', 'In Progress', 'Review', 'On Hold', 'Done'),
    allowNull: false
  },
  note: {
    type: DataTypes.TEXT,
    allowNull: false,
    validate: {
      notEmpty: true
    }
  },
  meetingDate: {
    type: DataTypes.DATE,
    allowNull: true
  },
  updatedBy: {
    type: DataTypes.STRING,
    allowNull: false,
    comment: 'Email of the user who made this update'
  }
}, {
  tableName: 'todo_updates',
  timestamps: true,
  indexes: [
    {
      fields: ['todoId']
    },
    {
      fields: ['meetingDate']
    }
  ]
});

module.exports = TodoUpdate;