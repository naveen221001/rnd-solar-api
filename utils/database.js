const { sequelize, Todo, TodoUpdate, Meeting } = require('../models');

async function initializeDatabase() {
  try {
    console.log('🔄 Connecting to PostgreSQL database...');
    
    // Test the connection
    await sequelize.authenticate();
    console.log('✅ Database connection established successfully');
    
    // Sync all models (create tables if they don't exist)
    await sequelize.sync({ 
      alter: process.env.NODE_ENV === 'development', // Only alter in development
      force: false // Never drop tables in production
    });
    
    console.log('✅ Database tables synchronized successfully');
    
    // Log database info
    const todoCount = await Todo.count();
    const updateCount = await TodoUpdate.count();
    const meetingCount = await Meeting.count();
    
    console.log(`📊 Database Status:`);
    console.log(`   - Todos: ${todoCount}`);
    console.log(`   - Updates: ${updateCount}`);
    console.log(`   - Meetings: ${meetingCount}`);
    
    return true;
    
  } catch (error) {
    console.error('❌ Database initialization failed:', error);
    console.error('Database URL:', process.env.DATABASE_URL ? 'Set' : 'Not set');
    return false;
  }
}

async function testDatabaseConnection() {
  try {
    await sequelize.authenticate();
    return true;
  } catch (error) {
    console.error('❌ Database connection test failed:', error);
    return false;
  }
}

module.exports = {
  initializeDatabase,
  testDatabaseConnection,
  sequelize
};