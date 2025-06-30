const { DataTypes } = require('sequelize');
const sequelize = require('../config/database');

const Meeting = sequelize.define('Meeting', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  meetingDate: {
    type: DataTypes.DATE,
    allowNull: false
  },
  attendees: {
    type: DataTypes.TEXT,
    allowNull: false,
    comment: 'Comma-separated list of team member codes'
  },
  topicsDiscussed: {
    type: DataTypes.INTEGER,
    defaultValue: 0,
    comment: 'Number of todos discussed in this meeting'
  },
  notes: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  createdBy: {
    type: DataTypes.STRING,
    allowNull: false,
    comment: 'Email of the user who created this meeting record'
  }
}, {
  tableName: 'meetings',
  timestamps: true,
  indexes: [
    {
      fields: ['meetingDate']
    },
    {
      fields: ['createdBy']
    }
  ]
});

module.exports = Meeting;