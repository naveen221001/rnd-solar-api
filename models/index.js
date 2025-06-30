const sequelize = require('../config/database');
const Todo = require('./Todo');
const TodoUpdate = require('./Todoupdate');
const Meeting = require('./Meeting');

// Define associations
Todo.hasMany(TodoUpdate, {
  foreignKey: 'todoId',
  as: 'updates',
  onDelete: 'CASCADE'
});

TodoUpdate.belongsTo(Todo, {
  foreignKey: 'todoId',
  as: 'todo'
});

// Export models and sequelize
module.exports = {
  sequelize,
  Todo,
  TodoUpdate,
  Meeting
};