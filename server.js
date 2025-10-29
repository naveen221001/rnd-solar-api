


// Complete server.js - Express server with Microsoft Authentication
const userMap = {
  "naveen.chamaria@vikramsolar.com": "Naveen Kumar Chamaria",
  "aritra.de@vikramsolar.com": "Aritra De",
  "bidisha.saha@vikramsolar.com": "Bidisha Saha",
  "arup.mahapatra@vikramsolar.com": "Arup Mahapatra",
  "tannu@vikramsolar.com": "Tannu",
  "tanushree.roy@vikramsolar.com": "Tanushree Roy",
  "soumya.ghosal@vikramsolar.com": "Soumya Ghosal",
  "krishanu.ghosh@vikramsolar.com": "Krishanu Ghosh",
  "Samaresh.Banerjee@vikramsolar.com": "Samaresh Banerjee",
  "gopal.kumar@vikramsolar.com": "Gopal Kumar",
  "jai.jaiswal@vikramsolar.com": "Jai Jaiswal",
  "shakya.acharya@vikramsolar.com": "Shakya Acharya",
  "deepanjana.adak@vikramsolar.com": "Deepanjana Adak",
  "sumit.kumar@vikramsolar.com": "Sumit Kumar",
  "rnd.lab@vikramsolar.com": "R&D Lab",
  "aman.srivastava@vikramsolar.com" : "Aman Srivastava",
  "samaresh.banerjee89@gmail.com" : "Samaresh Banerjee"
};

// Authorized R&D team emails
const AUTHORIZED_EMAILS = [
  "naveen.chamaria@vikramsolar.com",
  "aritra.de@vikramsolar.com",
  "aman.srivastava@vikramsolar.com",
  "bidisha.saha@vikramsolar.com",
  "arup.mahapatra@vikramsolar.com",
  "tannu@vikramsolar.com",
  "tanushree.roy@vikramsolar.com",
  "soumya.ghosal@vikramsolar.com",
  "krishanu.ghosh@vikramsolar.com",
  "samaresh.banerjee@vikramsolar.com",
  "gopal.kumar@vikramsolar.com",
  "jai.jaiswal@vikramsolar.com",
  "shakya.acharya@vikramsolar.com",
  "deepanjana.adak@vikramsolar.com",
  // 2 spaces reserved for future team members
  "sumit.kumar@vikramsolar.com",
  "rnd.lab@vikramsolar.com",
  "samaresh.banerjee89@gmail.com"
];

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const { initializeDatabase } = require('./utils/database');
const { Todo, TodoUpdate, Meeting } = require('./models');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
const PORT = process.env.PORT || 3001;

// Microsoft Azure AD configuration
const client = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`
});

function getKey(header, callback) {
  client.getSigningKey(header.kid, function (err, key) {
    const signingKey = key.getPublicKey();
    callback(null, signingKey);
  });
}

// Authentication middleware for Microsoft tokens
function authenticateMicrosoftToken(req, res, next) {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ message: 'Token missing' });
  }

  jwt.verify(token, getKey, {}, (err, decoded) => {
    if (err) {
      console.error("❌ Microsoft token verification failed:", err);
      return res.status(403).json({ message: 'Invalid token' });
    }
    
    // Check if user email is authorized
    const userEmail = decoded.preferred_username || decoded.upn || decoded.email;
    
    if (!AUTHORIZED_EMAILS.includes(userEmail)) {
      console.log(`❌ Unauthorized email attempted access: ${userEmail}`);
      return res.status(403).json({ message: 'Unauthorized email address' });
    }
    
    console.log(`✅ Authorized user authenticated: ${userMap[userEmail] || userEmail}`);
    req.user = decoded;
    next();
  });
}

// Enable CORS for your domains (updated to include Authorization header)
app.use(cors({
  origin: [
    'https://vikramsolar-rnd-rm-dashboard-naveen.netlify.app',
    'http://localhost:3000',
    'http://localhost:3001'
  ],
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'],
  credentials: true,
  optionsSuccessStatus: 200
}));

app.options('*', cors({
  origin: [
    'https://vikramsolar-rnd-rm-dashboard-naveen.netlify.app',
    'http://localhost:3000',
    'http://localhost:3001'
  ],
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'Accept'],
  credentials: true
}));

// Disable response caching
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Helper function to safely parse dates from Excel
function parseExcelDate(dateValue) {
  if (dateValue == null) return null;
  if (dateValue instanceof Date) return dateValue;
  
  try {
    if (typeof dateValue === 'number') {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const millisecondsPerDay = 24 * 60 * 60 * 1000;
      return new Date(excelEpoch.getTime() + dateValue * millisecondsPerDay);
    }
    
    if (typeof dateValue === 'string' && dateValue.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      const [day, month, year] = dateValue.split('/');
      return new Date(Date.UTC(year, month - 1, day));
    }
    
    return new Date(String(dateValue));
  } catch (e) {
    console.error('Error parsing date:', dateValue, e);
    return null;
  }
}

// Add this helper function after the parseExcelDate function
function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}


function parseExcelDate(dateValue) {
  if (dateValue == null) return null;
  if (dateValue instanceof Date) return dateValue;
  
  try {
    if (typeof dateValue === 'number') {
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const millisecondsPerDay = 24 * 60 * 60 * 1000;
      return new Date(excelEpoch.getTime() + dateValue * millisecondsPerDay);
    }
    
    if (typeof dateValue === 'string' && dateValue.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      const [day, month, year] = dateValue.split('/');
      return new Date(Date.UTC(year, month - 1, day));
    }
    
    return new Date(String(dateValue));
  } catch (e) {
    console.error('Error parsing date:', dateValue, e);
    return null;
  }
}

async function backupTodosToGitHub() {
  try {
    console.log('🔄 Starting todos backup to GitHub...');
    
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    if (!fileInfo.exists) {
      console.log('❌ No todos file to backup');
      return false;
    }
    
    // Read the current file
    const filePath = fileInfo.path;
    const fileExists = fs.existsSync(filePath);
    
    if (fileExists) {
      const stats = fs.statSync(filePath);
      console.log(`📊 Backing up todos file: ${stats.size} bytes, modified: ${stats.mtime}`);
      
      // Create a backup timestamp
      const backupInfo = {
        timestamp: new Date().toISOString(),
        fileSize: stats.size,
        lastModified: stats.mtime,
        backupReason: 'Scheduled backup after todo changes'
      };
      
      // Save backup info
      const backupInfoPath = path.join(__dirname, 'data', '.backup_info.json');
      fs.writeFileSync(backupInfoPath, JSON.stringify(backupInfo, null, 2));
      
      console.log('✅ Todos backup completed');
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('❌ Error backing up todos:', error);
    return false;
  }
}

async function createExcelBackupFromDatabase() {
  try {
    console.log('🔄 Creating Excel backup from database data...');
    
    // Get all data from database
    const todos = await Todo.findAll({ order: [['id', 'ASC']] });
    const updates = await TodoUpdate.findAll({ order: [['todoId', 'ASC'], ['createdAt', 'ASC']] });
    const meetings = await Meeting.findAll({ order: [['meetingDate', 'DESC']] });
    
    // Convert to Excel format
    const todosExcelData = todos.map(todo => ({
      'ID': todo.id,
      'ISSUE': todo.issue,
      'RESPONSIBILITY': todo.responsibility,
      'STATUS': todo.status,
      'PRIORITY': todo.priority,
      'CATEGORY': todo.category,
      'DUE_DATE': todo.dueDate,
      'CREATED_DATE': todo.createdAt,
      'CREATED_BY': todo.createdBy
    }));
    
    const updatesExcelData = updates.map(update => ({
      'TODO_ID': update.todoId,
      'UPDATE_DATE': update.createdAt,
      'STATUS': update.status,
      'NOTE': update.note,
      'MEETING_DATE': update.meetingDate,
      'UPDATED_BY': update.updatedBy
    }));
    
    const meetingsExcelData = meetings.map(meeting => ({
      'MEETING_DATE': meeting.meetingDate,
      'ATTENDEES': meeting.attendees,
      'TOPICS_DISCUSSED': meeting.topicsDiscussed,
      'NOTES': meeting.notes,
      'CREATED_BY': meeting.createdBy,
      'CREATED_AT': meeting.createdAt
    }));
    
    // Create Excel workbook
    const workbook = xlsx.utils.book_new();
    
    const todosSheet = xlsx.utils.json_to_sheet(todosExcelData);
    const updatesSheet = xlsx.utils.json_to_sheet(updatesExcelData);
    const meetingsSheet = xlsx.utils.json_to_sheet(meetingsExcelData);
    
    xlsx.utils.book_append_sheet(workbook, todosSheet, 'Todos');
    xlsx.utils.book_append_sheet(workbook, updatesSheet, 'Updates');
    xlsx.utils.book_append_sheet(workbook, meetingsSheet, 'Meetings');
    
    // Ensure data directory exists
    const dataDir = path.join(__dirname, 'data');
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }
    
    // Write Excel file
    const excelPath = path.join(dataDir, 'RND_Todos.xlsx');
    xlsx.writeFile(workbook, excelPath);
    
    console.log(`✅ Excel backup created: ${todos.length} todos, ${updates.length} updates, ${meetings.length} meetings`);
    console.log(`📁 Excel file saved to: ${excelPath}`);
    
    return true;
    
  } catch (error) {
    console.error('❌ Error creating Excel backup from database:', error);
    return false;
  }
}


function calculateChamberData(rawData, worksheet) {
  // Get the column range from the worksheet
  const range = xlsx.utils.decode_range(worksheet['!ref']);
  
  // Column Q is column 16 (0-indexed, so Q = 16)
  const columnQ_Index = 16; // Q is the 17th column (0-indexed = 16)
  
  return rawData.map((row, index) => {
    const moduleId = row['Module ID'] || '';
    const vslId = row['VSL ID'] || '';
    const bomUnderTest = row['BOM under test'] || '';
    const testName = row['Test Name'] || '';
    const count = parseFloat(row['Count']) || 0;
    const startDate = parseExcelDate(row['Start Date']);
    const actualEndDate = parseExcelDate(row['Actual End Date']);

    // Calculate Done (HR) - sum all values from column Q onwards
    let doneHr = 0;
    let nonZeroCount = 0;
    let columnCount = 0;
    
    // Get all column keys
    const allColumns = Object.keys(row);
    
    // Filter and sum columns from Q onwards
    allColumns.forEach(col => {
      // Skip the known non-date columns (first 16 columns A-P)
      const knownColumns = [
        'Module ID', 'VSL ID', 'BOM under test', 'Test Name', 'Count',
        'Start Date', 'Progress', 'Done (Hr)', 'Remaining (Hr)',
        'Done (Cycles)', 'Remaining (Cycles)', 'Tentative End',
        'Lag', 'Type', 'Cycle Time (Hr)', 'Total Duration (Hr)',
        'Actual End Date', 'Status'
      ];
      
      // If it's not a known column and contains a value, it's likely a date column
      if (!knownColumns.includes(col)) {
        const value = parseFloat(row[col]) || 0;
        if (!isNaN(value)) {
          doneHr += value;
          columnCount++;
          if (value > 0) nonZeroCount++;
        }
      }
    });
    
    doneHr = Math.round(doneHr * 100) / 100; // Round to 2 decimal places
    
    // Debug logging for first few rows
    if (index < 3) {
      console.log(`\nRow ${index + 1} Debug Info:`);
      console.log(`- Module ID: ${moduleId}, Test Name: ${testName}`);
      console.log(`- Count: ${count}`);
      console.log(`- Columns with numeric values: ${columnCount}`);
      console.log(`- Non-zero entries: ${nonZeroCount}`);
      console.log(`- Calculated Done (Hr): ${doneHr}`);
    }

    // Calculate Type first (needed for other calculations)
    let type;
    if (testName === 'TC' || testName === 'HF') {
      type = 'Cycle';
    } else if (testName === 'LID') {
      type = 'kWHr';
    } else {
      type = 'Hr';
    }

    // Calculate Cycle Time (Hr)
    let cycleTimeHr;
    if (type === 'Hr') {
      cycleTimeHr = 'NA';
    } else if (testName === 'TC') {
      cycleTimeHr = 2.1;
    } else if (testName === 'LETID') {
      cycleTimeHr = 162;
    } else if (testName === 'LID') {
      cycleTimeHr = 1;
    } else {
      cycleTimeHr = 24;
    }

    // Calculate Total Duration (HR)
    let totalDurationHr;
    if (cycleTimeHr === 'NA') {
      totalDurationHr = count;
    } else {
      totalDurationHr = count * cycleTimeHr;
    }

    // Calculate Remaining (HR)
    let remainingHr;
    if (doneHr >= totalDurationHr) {
      remainingHr = 'DONE';
    } else {
      remainingHr = Math.round((totalDurationHr - doneHr) * 100) / 100;
    }

    // Calculate Done (Cycles)
    let doneCycles;
    if (testName === 'TC') {
      doneCycles = Math.round((doneHr / 2.1) * 100) / 100;
    } else if (testName === 'HF') {
      doneCycles = Math.round((doneHr / 24) * 100) / 100;
    } else {
      doneCycles = '-';
    }

    // Calculate Remaining (Cycles)
    let remainingCycles;
    if (remainingHr === 'DONE') {
      remainingCycles = '0';
    } else if (testName === 'TC') {
      remainingCycles = Math.round((remainingHr / 2.1) * 100) / 100;
    } else if (testName === 'HF') {
      remainingCycles = Math.round((remainingHr / 24) * 100) / 100;
    } else {
      remainingCycles = '-';
    }

    // Calculate Tentative End Date
    let tentativeEndDate = null;
    
    // Check conditions: if LID test, or start date is blank, or sum of daily hours < 1
    const isLID = testName === 'LID';
    const isStartDateBlank = !startDate;
    const isDailyHoursLessThanOne = doneHr < 1;
    
    if (isLID || isStartDateBlank || isDailyHoursLessThanOne) {
      tentativeEndDate = null;
    } else {
      if (remainingHr === 'DONE') {
        const daysToAdd = Math.ceil(totalDurationHr / 24);
        tentativeEndDate = addDays(startDate, daysToAdd);
      } else {
        // For now, just add remaining hours / 24 to today
        const today = new Date();
        const daysToAdd = Math.ceil(remainingHr / 24);
        tentativeEndDate = addDays(today, daysToAdd);
      }
    }

    // Calculate Lag
    let lag = null;
    if (tentativeEndDate && actualEndDate) {
      const diffTime = actualEndDate - tentativeEndDate;
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays >= 1) {
        lag = diffDays;
      } else {
        lag = 0;
      }
    }

    return {
      id: index + 1,
      moduleId,
      vslId,
      bomUnderTest,
      testName,
      count,
      startDate,
      doneHr,
      remainingHr,
      doneCycles,
      remainingCycles,
      tentativeEndDate,
      lag,
      type,
      cycleTimeHr,
      totalDurationHr,
      actualEndDate
    };
  });
}

// Define standard test durations manually as a backup
const HARDCODED_STD_DURATIONS = {
  'ADHESION - PCT - ADHESION': 6,
  'GEL TEST': 2,
  'TENSILE TEST': 2,
  'SHRINKAGE TEST': 2,
  'GSM TEST': 1,
  'SML': 2,
  'DML': 2,
  'SKINFREE TEST': 1,
  'CURING TEST': 3,
  'HARDNESS TEST': 1,
  'BYPASS DIODE TEST': 2,
  'THERMAL AND FUCNTIONAL': 2,
  'PCT': 2,
  'RESISTANCE TEST': 1,
  'PEEL STRENGTH': 2
};


// Add this function after the calculateChamberData function in server.js

// Function to calculate shrinkage test results from Shrinkage format
// UPDATED: Fix calculateShrinkageResults function in server.js
// IMPROVED: More robust shrinkage calculation function
// REPLACE your calculateShrinkageResults function with this enhanced version:

function calculateShrinkageResults(shrinkageData) {
  console.log('\n🧪 CALCULATING SHRINKAGE RESULTS:');
  console.log('==================================');
  
  return shrinkageData.map((row, index) => {
    console.log(`\nProcessing row ${index + 1}:`);
    console.log('Available columns:', Object.keys(row));
    
    // Get vendor and encapsulant type with multiple possible column names
    const vendorName = row['VENDOR NAME'] || row['Vendor Name'] || row['VENDOR_NAME'] || 
                      row['vendor name'] || row['Vendor'] || '';
    
    const encapsulantType = row['ENCAPSULANT TYPE'] || row['Encapsulant Type'] || 
                           row['ENCAPSULANT_TYPE'] || row['encapsulant type'] || 
                           row['Type'] || row['TYPE'] || '';
    
    console.log(`Vendor: "${vendorName}", Type: "${encapsulantType}"`);
    
    // Parse values for WITHOUT HEAT (columns C,D,E,F)
    // Try multiple possible column names for Excel variations
    const td1WithoutHeat = parseFloat(
      row['TD1'] || row['td1'] || row['TD1_WITHOUT'] || 
      row['TD1 WITHOUT HEAT'] || row['TD1_WO_HEAT'] || 
      row['__EMPTY'] || row['__EMPTY_2'] || 0  // Excel sometimes uses __EMPTY for unlabeled columns
    );
    
    const td2WithoutHeat = parseFloat(
      row['TD2'] || row['td2'] || row['TD2_WITHOUT'] || 
      row['TD2 WITHOUT HEAT'] || row['TD2_WO_HEAT'] ||
      row['__EMPTY_1'] || row['__EMPTY_3'] || 0
    );
    
    const md1WithoutHeat = parseFloat(
      row['MD1'] || row['md1'] || row['MD1_WITHOUT'] || 
      row['MD1 WITHOUT HEAT'] || row['MD1_WO_HEAT'] ||
      row['__EMPTY_2'] || row['__EMPTY_4'] || 0
    );
    
    const md2WithoutHeat = parseFloat(
      row['MD2'] || row['md2'] || row['MD2_WITHOUT'] || 
      row['MD2 WITHOUT HEAT'] || row['MD2_WO_HEAT'] ||
      row['__EMPTY_3'] || row['__EMPTY_5'] || 0
    );
    
    // Parse values for WITH HEAT (columns G,H,I,J)
    // Handle Excel's automatic column renaming (__1, .1, etc.)
    const td1WithHeat = parseFloat(
      row['TD1__1'] || row['TD1.1'] || row['TD1_1'] || 
      row['TD1_WITH'] || row['TD1_HEAT'] || row['TD1 WITH HEAT'] ||
      row['__EMPTY_6'] || row['__EMPTY_4'] || 0
    );
    
    const td2WithHeat = parseFloat(
      row['TD2__1'] || row['TD2.1'] || row['TD2_1'] || 
      row['TD2_WITH'] || row['TD2_HEAT'] || row['TD2 WITH HEAT'] ||
      row['__EMPTY_7'] || row['__EMPTY_5'] || 0
    );
    
    const md1WithHeat = parseFloat(
      row['MD1__1'] || row['MD1.1'] || row['MD1_1'] || 
      row['MD1_WITH'] || row['MD1_HEAT'] || row['MD1 WITH HEAT'] ||
      row['__EMPTY_8'] || row['__EMPTY_6'] || 0
    );
    
    const md2WithHeat = parseFloat(
      row['MD2__1'] || row['MD2.1'] || row['MD2_1'] || 
      row['MD2_WITH'] || row['MD2_HEAT'] || row['MD2 WITH HEAT'] ||
      row['__EMPTY_9'] || row['__EMPTY_7'] || 0
    );
    
    console.log('WITHOUT HEAT:', { td1: td1WithoutHeat, td2: td2WithoutHeat, md1: md1WithoutHeat, md2: md2WithoutHeat });
    console.log('WITH HEAT:', { td1: td1WithHeat, td2: td2WithHeat, md1: md1WithHeat, md2: md2WithHeat });
    
    // Check if we have valid data
    const hasValidWithoutHeatData = td1WithoutHeat > 0 || td2WithoutHeat > 0 || md1WithoutHeat > 0 || md2WithoutHeat > 0;
    const hasValidWithHeatData = td1WithHeat > 0 || td2WithHeat > 0 || md1WithHeat > 0 || md2WithHeat > 0;
    
    if (!hasValidWithoutHeatData || !hasValidWithHeatData) {
      console.log('⚠️ Missing measurement data - checking alternative column names...');
      
      // List all non-empty numeric columns for debugging
      const numericColumns = {};
      Object.keys(row).forEach(colName => {
        const value = parseFloat(row[colName]);
        if (!isNaN(value) && value > 0) {
          numericColumns[colName] = value;
        }
      });
      console.log('Available numeric columns:', numericColumns);
    }
    
    // Calculate means (handle zero values gracefully)
    const tdMeanWithoutHeat = td1WithoutHeat > 0 && td2WithoutHeat > 0 ? 
      (td1WithoutHeat + td2WithoutHeat) / 2 : 
      Math.max(td1WithoutHeat, td2WithoutHeat);
      
    const tdMeanWithHeat = td1WithHeat > 0 && td2WithHeat > 0 ? 
      (td1WithHeat + td2WithHeat) / 2 : 
      Math.max(td1WithHeat, td2WithHeat);
      
    const mdMeanWithoutHeat = md1WithoutHeat > 0 && md2WithoutHeat > 0 ? 
      (md1WithoutHeat + md2WithoutHeat) / 2 : 
      Math.max(md1WithoutHeat, md2WithoutHeat);
      
    const mdMeanWithHeat = md1WithHeat > 0 && md2WithHeat > 0 ? 
      (md1WithHeat + md2WithHeat) / 2 : 
      Math.max(md1WithHeat, md2WithHeat);
    
    // Calculate differences (shrinkage percentage)
    const tdDifference = Math.abs(tdMeanWithHeat - tdMeanWithoutHeat);
    const mdDifference = Math.abs(mdMeanWithHeat - mdMeanWithoutHeat);
    
    // Apply pass/fail criteria (less than 1% shrinkage = pass)
    const tdStatus = tdDifference < 1.0 ? 'PASS' : 'FAIL';
    const mdStatus = mdDifference < 1.0 ? 'PASS' : 'FAIL';
    const finalStatus = (tdStatus === 'PASS' && mdStatus === 'PASS') ? 'PASS' : 'FAIL';
    
    console.log('CALCULATIONS:');
    console.log(`TD Mean Without Heat: ${tdMeanWithoutHeat.toFixed(2)}`);
    console.log(`TD Mean With Heat: ${tdMeanWithHeat.toFixed(2)}`);
    console.log(`TD Difference: ${tdDifference.toFixed(2)}% (${tdStatus})`);
    console.log(`MD Mean Without Heat: ${mdMeanWithoutHeat.toFixed(2)}`);
    console.log(`MD Mean With Heat: ${mdMeanWithHeat.toFixed(2)}`);
    console.log(`MD Difference: ${mdDifference.toFixed(2)}% (${mdStatus})`);
    console.log(`FINAL RESULT: ${finalStatus}`);
    
    return {
      id: index + 1,
      vendorName,
      encapsulantType,
      
      // Raw measurements
      td1WithoutHeat, td2WithoutHeat, md1WithoutHeat, md2WithoutHeat,
      td1WithHeat, td2WithHeat, md1WithHeat, md2WithHeat,
      
      // Calculated means
      tdMeanWithoutHeat: Math.round(tdMeanWithoutHeat * 100) / 100,
      tdMeanWithHeat: Math.round(tdMeanWithHeat * 100) / 100,
      mdMeanWithoutHeat: Math.round(mdMeanWithoutHeat * 100) / 100,
      mdMeanWithHeat: Math.round(mdMeanWithHeat * 100) / 100,
      
      // Differences and results
      tdDifference: Math.round(tdDifference * 100) / 100,
      mdDifference: Math.round(mdDifference * 100) / 100,
      tdStatus, 
      mdStatus, 
      finalStatus
    };
  });
}

// ALSO UPDATE the shrinkage sheet reading section in your API endpoint:
// Replace the existing shrinkage reading section with this:



// FIXED: Function to update test results in Test Data based on shrinkage results
function updateTestDataWithShrinkageResults(testData, shrinkageResults) {
  console.log('🔍 Starting shrinkage integration...');
  console.log(`Processing ${testData.length} test records`);
  console.log(`Available shrinkage results: ${shrinkageResults.length}`);
  
  return testData.map(testRow => {
    // Check if this is a shrinkage test - IMPROVED MATCHING
    const testName = (testRow['TEST NAME'] || '').toUpperCase();
    const vendorName = testRow['VENDOR NAME'] || '';
    
    console.log(`Checking test: "${testName}" for vendor: "${vendorName}"`);
    
    if (testName.includes('SHRINKAGE')) {
      console.log(`✅ Found shrinkage test for vendor: ${vendorName}`);
      
      // Find matching shrinkage results for this vendor
      const vendorShrinkageData = shrinkageResults.filter(shrinkage => {
        const shrinkageVendor = shrinkage.vendorName || '';
        console.log(`Comparing: "${vendorName}" with "${shrinkageVendor}"`);
        return shrinkageVendor === vendorName;
      });
      
      console.log(`Found ${vendorShrinkageData.length} shrinkage results for ${vendorName}`);
      
      if (vendorShrinkageData.length > 0) {
        // Check if both FRONT EPE and BACK EVA pass for this vendor
        const frontEpeResult = vendorShrinkageData.find(item => 
          item.encapsulantType === 'FRONT EPE'
        );
        const backEvaResult = vendorShrinkageData.find(item => 
          item.encapsulantType === 'BACK EVA'
        );
        
        console.log(`Front EPE result:`, frontEpeResult);
        console.log(`Back EVA result:`, backEvaResult);
        
        let overallResult = 'FAIL';
        
        // FIXED LOGIC: Both FRONT EPE and BACK EVA must pass for overall pass
        if (frontEpeResult && backEvaResult) {
          if (frontEpeResult.finalStatus === 'PASS' && backEvaResult.finalStatus === 'PASS') {
            overallResult = 'PASS';
            console.log(`✅ Both FRONT EPE and BACK EVA passed for ${vendorName}`);
          } else {
            console.log(`❌ One or both failed - Front: ${frontEpeResult.finalStatus}, Back: ${backEvaResult.finalStatus}`);
          }
        } else if (frontEpeResult && frontEpeResult.finalStatus === 'PASS' && !backEvaResult) {
          // Only FRONT EPE tested and it passed
          overallResult = 'PASS';
          console.log(`✅ Only FRONT EPE tested and passed for ${vendorName}`);
        } else if (backEvaResult && backEvaResult.finalStatus === 'PASS' && !frontEpeResult) {
          // Only BACK EVA tested and it passed
          overallResult = 'PASS';
          console.log(`✅ Only BACK EVA tested and passed for ${vendorName}`);
        } else {
          console.log(`❌ No valid results found for ${vendorName}`);
        }
        
        // IMPORTANT: Update the TEST RESULT field directly
        testRow['TEST RESULT'] = overallResult;
        testRow['SHRINKAGE_CALCULATION_DETAILS'] = {
          frontEpe: frontEpeResult,
          backEva: backEvaResult,
          overallResult: overallResult,
          calculatedAt: new Date().toISOString(),
          vendorName: vendorName
        };
        
        console.log(`🔄 Updated test result for ${vendorName}: ${overallResult}`);
      } else {
        console.log(`⚠️ No shrinkage calculation data found for vendor: ${vendorName}`);
        // Keep the original test result if no calculation data is available
      }
    }
    
    return testRow;
  });
}

// Add this function after the calculateShrinkageResults function in server.js

// Function to calculate adhesion test results from Adhesion sheet
// 5. UPDATED: Adhesion calculation - specify header row
// UPDATED: Fix calculateAdhesionResults function in server.js
function calculateAdhesionResults(adhesionData) {
  return adhesionData.map((row, index) => {
    // Based on your Excel: BOM, VENDOR NAME, then the measurement columns
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    
    // Parse values for PRE PCT section (columns C-F in Excel)
    // GLASS to ENCAPSULANT: max, min
    const prePctGlassToEncapMax = parseFloat(row['max']) || 0;
    const prePctGlassToEncapMin = parseFloat(row['min']) || 0;
    
    // BACKSHEET to ENCAPSULANT: max, min
    const prePctBacksheetToEncapMax = parseFloat(row['max__1']) || parseFloat(row['max.1']) || 0;
    const prePctBacksheetToEncapMin = parseFloat(row['min__1']) || parseFloat(row['min.1']) || 0;
    
    // Parse values for POST PCT section (columns G-J in Excel)
    // GLASS to ENCAPSULANT: max, min
    const postPctGlassToEncapMax = parseFloat(row['max__2']) || parseFloat(row['max.2']) || 0;
    const postPctGlassToEncapMin = parseFloat(row['min__2']) || parseFloat(row['min.2']) || 0;
    
    // BACKSHEET to ENCAPSULANT: max, min  
    const postPctBacksheetToEncapMax = parseFloat(row['max__3']) || parseFloat(row['max.3']) || 0;
    const postPctBacksheetToEncapMin = parseFloat(row['min__3']) || parseFloat(row['min.3']) || 0;
    
    // Debug logging for first few rows
    if (index < 2) {
      console.log(`Adhesion row ${index + 1} raw data:`, {
        vendor: vendorName,
        bom: bom,
        prePct: { glassMax: prePctGlassToEncapMax, glassMin: prePctGlassToEncapMin, 
                  backMax: prePctBacksheetToEncapMax, backMin: prePctBacksheetToEncapMin },
        postPct: { glassMax: postPctGlassToEncapMax, glassMin: postPctGlassToEncapMin,
                   backMax: postPctBacksheetToEncapMax, backMin: postPctBacksheetToEncapMin }
      });
    }
    
    // Calculate averages
    const prePctGlassToEncapAvg = (prePctGlassToEncapMax + prePctGlassToEncapMin) / 2;
    const prePctBacksheetToEncapAvg = (prePctBacksheetToEncapMax + prePctBacksheetToEncapMin) / 2;
    const postPctGlassToEncapAvg = (postPctGlassToEncapMax + postPctGlassToEncapMin) / 2;
    const postPctBacksheetToEncapAvg = (postPctBacksheetToEncapMax + postPctBacksheetToEncapMin) / 2;
    
    // Apply criteria: Glass > 60, Backsheet > 40
    const prePctGlassToEncapStatus = prePctGlassToEncapAvg > 60 ? 'PASS' : 'FAIL';
    const prePctBacksheetToEncapStatus = prePctBacksheetToEncapAvg > 40 ? 'PASS' : 'FAIL';
    const postPctGlassToEncapStatus = postPctGlassToEncapAvg > 60 ? 'PASS' : 'FAIL';
    const postPctBacksheetToEncapStatus = postPctBacksheetToEncapAvg > 40 ? 'PASS' : 'FAIL';
    
    const finalStatus = (prePctGlassToEncapStatus === 'PASS' && 
                        prePctBacksheetToEncapStatus === 'PASS' && 
                        postPctGlassToEncapStatus === 'PASS' && 
                        postPctBacksheetToEncapStatus === 'PASS') ? 'PASS' : 'FAIL';
    
    return {
      id: index + 1, vendorName, bom,
      prePctGlassToEncapMax, prePctGlassToEncapMin,
      prePctBacksheetToEncapMax, prePctBacksheetToEncapMin,
      postPctGlassToEncapMax, postPctGlassToEncapMin,
      postPctBacksheetToEncapMax, postPctBacksheetToEncapMin,
      prePctGlassToEncapAvg: Math.round(prePctGlassToEncapAvg * 100) / 100,
      prePctBacksheetToEncapAvg: Math.round(prePctBacksheetToEncapAvg * 100) / 100,
      postPctGlassToEncapAvg: Math.round(postPctGlassToEncapAvg * 100) / 100,
      postPctBacksheetToEncapAvg: Math.round(postPctBacksheetToEncapAvg * 100) / 100,
      prePctGlassToEncapStatus, prePctBacksheetToEncapStatus,
      postPctGlassToEncapStatus, postPctBacksheetToEncapStatus,
      finalStatus
    };
  });
}

// Function to update test results in Test Data based on adhesion results
function updateTestDataWithAdhesionResults(testData, adhesionResults) {
  return testData.map(testRow => {
    // Check if this is an adhesion test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('ADHESION')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching adhesion results for this vendor and BOM
      const adhesionResult = adhesionResults.find(result => 
        result.vendorName === vendorName && result.bom === bomType
      );
      
      if (adhesionResult) {
        // Update the test result
        testRow['TEST RESULT'] = adhesionResult.finalStatus;
        testRow['ADHESION_CALCULATION_DETAILS'] = {
          prePctGlassToEncap: {
            avg: adhesionResult.prePctGlassToEncapAvg,
            status: adhesionResult.prePctGlassToEncapStatus,
            criteria: '> 60'
          },
          prePctBacksheetToEncap: {
            avg: adhesionResult.prePctBacksheetToEncapAvg,
            status: adhesionResult.prePctBacksheetToEncapStatus,
            criteria: '> 40'
          },
          postPctGlassToEncap: {
            avg: adhesionResult.postPctGlassToEncapAvg,
            status: adhesionResult.postPctGlassToEncapStatus,
            criteria: '> 60'
          },
          postPctBacksheetToEncap: {
            avg: adhesionResult.postPctBacksheetToEncapAvg,
            status: adhesionResult.postPctBacksheetToEncapStatus,
            criteria: '> 40'
          },
          overallResult: adhesionResult.finalStatus
        };
      }
    }
    
    return testRow;
  });
}


// Add this function after the calculateAdhesionResults function in server.js

// Function to calculate tensile strength test results from Tensile Strength sheet
// UPDATED: Fix calculateTensileStrengthResults function in server.js
function calculateTensileStrengthResults(tensileData) {
  return tensileData.map((row, index) => {
    // Based on your Excel: BOM, VENDOR NAME, BREAK, CHANGE IN ELONGATION %
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    
    // Parse Break value (should be > 10 MPa)
    const breakValue = parseFloat(row['BREAK']) || 0;
    
    // Parse Change in Elongation % (should be >= 450%)
    let changeInElongationPercent = parseFloat(row['CHANGE IN ELONGATION %']) || 0;
    
    // Debug logging for first few rows
    if (index < 2) {
      console.log(`Tensile row ${index + 1} raw data:`, {
        vendor: vendorName,
        bom: bom,
        break: breakValue,
        elongation: changeInElongationPercent
      });
    }
    
    // Apply pass/fail criteria
    const breakStatus = breakValue > 10 ? 'PASS' : 'FAIL';
    const elongationStatus = changeInElongationPercent >= 450 ? 'PASS' : 'FAIL';
    
    // Final status: PASS only if BOTH break and elongation pass
    const finalStatus = (breakStatus === 'PASS' && elongationStatus === 'PASS') ? 'PASS' : 'FAIL';
    
    return {
      id: index + 1,
      vendorName,
      bom,
      
      // Raw values
      breakValue: Math.round(breakValue * 100) / 100,
      changeInElongationPercent: Math.round(changeInElongationPercent * 100) / 100,
      
      // Individual status
      breakStatus,
      elongationStatus,
      
      // Criteria info
      breakCriteria: '> 10 MPa',
      elongationCriteria: '>= 450%',
      
      // Final status
      finalStatus
    };
  });
}

// Add this function to your server.js file 
// Place it after the calculateTensileStrengthResults function

// Function to update test results in Test Data based on tensile strength results
function updateTestDataWithTensileStrengthResults(testData, tensileStrengthResults) {
  return testData.map(testRow => {
    // Check if this is a tensile strength test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('TENSILE')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching tensile strength results for this vendor and BOM
      const tensileResult = tensileStrengthResults.find(result => 
        result.vendorName === vendorName && result.bom === bomType
      );
      
      if (tensileResult) {
        // Update the test result
        testRow['TEST RESULT'] = tensileResult.finalStatus;
        testRow['TENSILE_CALCULATION_DETAILS'] = {
          break: {
            value: tensileResult.breakValue,
            status: tensileResult.breakStatus,
            criteria: tensileResult.breakCriteria
          },
          elongation: {
            percent: tensileResult.changeInElongationPercent,
            status: tensileResult.elongationStatus,
            criteria: tensileResult.elongationCriteria,
            calculationMethod: tensileResult.initialLength && tensileResult.finalLength ? 
              'Calculated from lengths' : 'Direct percentage'
          },
          overallResult: tensileResult.finalStatus
        };
        
        // Add length details if available
        if (tensileResult.initialLength && tensileResult.finalLength) {
          testRow['TENSILE_CALCULATION_DETAILS'].elongation.initialLength = tensileResult.initialLength;
          testRow['TENSILE_CALCULATION_DETAILS'].elongation.finalLength = tensileResult.finalLength;
        }
        
        console.log(`Updated tensile test result for ${vendorName} ${bomType}: ${tensileResult.finalStatus} (break: ${tensileResult.breakValue} MPa, elongation: ${tensileResult.changeInElongationPercent}%)`);
      } else {
        console.log(`No tensile calculation found for ${vendorName} ${bomType}`);
      }
    }
    
    return testRow;
  });
}

// Function to update test results in Test Data based on tensile strength results
// 2. UPDATED: GSM calculation - matches your column names
// UPDATED: Fix calculateGSMResults function in server.js
function calculateGSMResults(gsmData) {
  return gsmData.map((row, index) => {
    // Based on your Excel: BOM, VENDOR NAME, Type, min value 1-5, Mean
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    const category = row['Type'] || ''; // This determines OLD vs NEW range
    
    // Parse the 5 measurement values
    const value1 = parseFloat(row['min value 1']) || 0;
    const value2 = parseFloat(row['min value 2']) || 0;
    const value3 = parseFloat(row['min value 3']) || 0;
    const value4 = parseFloat(row['min value 4']) || 0;
    const value5 = parseFloat(row['min value 5']) || 0;
    
    // Use provided mean if available, otherwise calculate
    let average = parseFloat(row['Mean']) || 0;
    if (average === 0) {
      const measurements = [value1, value2, value3, value4, value5];
      const validMeasurements = measurements.filter(val => val > 0);
      average = validMeasurements.length > 0 ? 
        validMeasurements.reduce((sum, val) => sum + val, 0) / validMeasurements.length : 0;
    }
    
    // Debug logging for first few rows
    if (index < 2) {
      console.log(`GSM row ${index + 1} raw data:`, {
        vendor: vendorName,
        bom: bom,
        category: category,
        values: [value1, value2, value3, value4, value5],
        mean: average
      });
    }
    
    // Define ranges based on category
    let minRange, maxRange, rangeName;
    
    if (category.toUpperCase().includes('OLD')) {
      minRange = 420; maxRange = 480; rangeName = 'OLD: 420 to 480';
    } else if (category.toUpperCase().includes('NEW')) {
      minRange = 380; maxRange = 440; rangeName = 'NEW: 380 to 440';
    } else {
      // Default to OLD range if unclear
      minRange = 420; maxRange = 480; rangeName = 'OLD: 420 to 480 (default)';
    }
    
    const isWithinRange = average >= minRange && average <= maxRange;
    const finalStatus = isWithinRange ? 'PASS' : 'FAIL';
    
    return {
      id: index + 1, vendorName, bom, category,
      value1, value2, value3, value4, value5,
      average: Math.round(average * 100) / 100,
      validMeasurementCount: [value1, value2, value3, value4, value5].filter(val => val > 0).length,
      minRange, maxRange, rangeName, isWithinRange, finalStatus,
      criteria: `${minRange} ≤ average ≤ ${maxRange}`
    };
  });
}

// Function to update test results in Test Data based on GSM results
function updateTestDataWithGSMResults(testData, gsmResults) {
  return testData.map(testRow => {
    // Check if this is a GSM test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('GSM')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching GSM results for this vendor and BOM
      const gsmResult = gsmResults.find(result => 
        result.vendorName === vendorName && result.bom === bomType
      );
      
      if (gsmResult) {
        // Update the test result
        testRow['TEST RESULT'] = gsmResult.finalStatus;
        testRow['GSM_CALCULATION_DETAILS'] = {
          measurements: {
            value1: gsmResult.value1,
            value2: gsmResult.value2,
            value3: gsmResult.value3,
            value4: gsmResult.value4,
            value5: gsmResult.value5,
            validCount: gsmResult.validMeasurementCount
          },
          calculation: {
            average: gsmResult.average,
            range: gsmResult.rangeName,
            minRange: gsmResult.minRange,
            maxRange: gsmResult.maxRange,
            criteria: gsmResult.criteria
          },
          status: {
            isWithinRange: gsmResult.isWithinRange,
            finalStatus: gsmResult.finalStatus
          },
          category: gsmResult.category
        };
        
        console.log(`Updated GSM test result for ${vendorName} ${bomType}: ${gsmResult.finalStatus} (avg: ${gsmResult.average}, range: ${gsmResult.rangeName})`);
      } else {
        console.log(`No GSM calculation found for ${vendorName} ${bomType}`);
      }
    }
    
    return testRow;
  });
}

// 3. UPDATED: Resistance calculation - handles your structure
// UPDATED: Fix processResistanceResults function in server.js
function processResistanceResults(resistanceData) {
  return resistanceData.map((row, index) => {
    // Based on your Excel: BOM, TYPE, VENDOR NAME, MEASURED VALUE
    const bom = row['BOM'] || '';
    const type = row['TYPE'] || '';
    const vendorName = row['VENDOR NAME'] || '';
    const measuredValue = parseFloat(row['MEASURED VALUE']) || 0;
    
    // Look for result column (might need manual entry)
    const testResult = row['TEST RESULT'] || row['RESULT'] || row['STATUS'] || 'Pending';
    
    // Debug logging for first few rows
    if (index < 2) {
      console.log(`Resistance row ${index + 1} raw data:`, {
        bom: bom,
        type: type,
        vendor: vendorName,
        measured: measuredValue,
        result: testResult
      });
    }
    
    return {
      id: index + 1, bom, type, vendorName,
      measuredValue: Math.round(measuredValue * 1000000) / 1000000, // 6 decimal places for resistance
      testResult: testResult.toUpperCase() === 'PASS' ? 'PASS' : 
                 testResult.toUpperCase() === 'FAIL' ? 'FAIL' : 'Pending',
      ribbonType: type, 
      hasValidMeasurement: measuredValue > 0,
      notes: row['NOTES'] || row['REMARKS'] || ''
    };
  });
}

// Function to update test results in Test Data based on resistance results
function updateTestDataWithResistanceResults(testData, resistanceResults) {
  return testData.map(testRow => {
    // Check if this is a resistance test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('RESISTANCE')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching resistance results for this vendor and BOM
      // Note: There might be multiple results (BUS RIBBON and INTERCONNECT RIBBON)
      const vendorResistanceData = resistanceResults.filter(result => 
        result.vendorName === vendorName && result.bom === bomType
      );
      
      if (vendorResistanceData.length > 0) {
        // Separate BUS RIBBON and INTERCONNECT RIBBON results
        const busRibbonResult = vendorResistanceData.find(item => 
          item.type.toUpperCase().includes('BUS')
        );
        const interconnectRibbonResult = vendorResistanceData.find(item => 
          item.type.toUpperCase().includes('INTERCONNECT')
        );
        
        // Determine overall result based on manual entries
        let overallResult = 'Pending';
        
        // Logic for overall result:
        // - If both types are tested and both are PASS, then PASS
        // - If any type is FAIL, then FAIL
        // - If any type is Pending, then Pending
        
        const busStatus = busRibbonResult ? busRibbonResult.testResult : 'Not Tested';
        const interconnectStatus = interconnectRibbonResult ? interconnectRibbonResult.testResult : 'Not Tested';
        
        if (busStatus === 'FAIL' || interconnectStatus === 'FAIL') {
          overallResult = 'FAIL';
        } else if (busStatus === 'PASS' && interconnectStatus === 'PASS') {
          overallResult = 'PASS';
        } else if (busStatus === 'PASS' && interconnectStatus === 'Not Tested') {
          overallResult = 'PASS'; // Only one type tested and passed
        } else if (interconnectStatus === 'PASS' && busStatus === 'Not Tested') {
          overallResult = 'PASS'; // Only one type tested and passed
        } else if (busStatus === 'PASS' || interconnectStatus === 'PASS') {
          overallResult = 'Pending'; // Mixed results
        } else {
          overallResult = 'Pending'; // Default
        }
        
        // Update the test result
        testRow['TEST RESULT'] = overallResult;
        testRow['RESISTANCE_CALCULATION_DETAILS'] = {
          busRibbon: busRibbonResult ? {
            measuredValue: busRibbonResult.measuredValue,
            result: busRibbonResult.testResult,
            notes: busRibbonResult.notes
          } : null,
          interconnectRibbon: interconnectRibbonResult ? {
            measuredValue: interconnectRibbonResult.measuredValue,
            result: interconnectRibbonResult.testResult,
            notes: interconnectRibbonResult.notes
          } : null,
          overallResult: overallResult,
          testMethod: 'Manual Assessment',
          criteria: 'No specific criteria - manually assessed',
          totalMeasurements: vendorResistanceData.length
        };
        
        console.log(`Updated resistance test result for ${vendorName} ${bomType}: ${overallResult}`);
        console.log(`- BUS RIBBON: ${busStatus}`);
        console.log(`- INTERCONNECT RIBBON: ${interconnectStatus}`);
      } else {
        console.log(`No resistance measurements found for ${vendorName} ${bomType}`);
      }
    }
    
    return testRow;
  });
}

// Function to process bypass diode test results from BYPASS DIODE TEST sheet
// 4. UPDATED: Bypass Diode calculation - handles your structure  
// UPDATED: Fix processBypassDiodeResults function in server.js
function processBypassDiodeResults(bypassDiodeData) {
  return bypassDiodeData.map((row, index) => {
    // Based on your Excel: BOM, VENDOR NAME, MAX TEMPERATURE OF DIODE(Tj)
    const bom = row['BOM'] || '';
    const vendorName = row['VENDOR NAME'] || '';
    const maxTemperatureTj = parseFloat(row['MAX TEMPERATURE OF DIODE(Tj)']) || 0;
    
    // Look for manual test result entry
    const testResult = row['TEST RESULT'] || row['RESULT'] || row['STATUS'] || 'Pending';
    
    // Debug logging for first few rows
    if (index < 2) {
      console.log(`Bypass Diode row ${index + 1} raw data:`, {
        bom: bom,
        vendor: vendorName,
        temperature: maxTemperatureTj,
        result: testResult
      });
    }
    
    let normalizedResult = 'Pending';
    if (testResult) {
      const resultUpper = testResult.toString().toUpperCase();
      if (resultUpper === 'PASS' || resultUpper === 'P') {
        normalizedResult = 'PASS';
      } else if (resultUpper === 'FAIL' || resultUpper === 'F') {
        normalizedResult = 'FAIL';
      }
    }
    
    return {
      id: index + 1, bom, vendorName,
      maxTemperatureTj: Math.round(maxTemperatureTj * 100) / 100,
      testResult: normalizedResult,
      hasValidTemperature: maxTemperatureTj > 0,
      temperatureUnit: '°C',
      notes: row['NOTES'] || row['REMARKS'] || '',
      testDate: row['TEST DATE'] || row['DATE'] || '',
      testedBy: row['TESTED BY'] || row['OPERATOR'] || '',
      assessmentMethod: 'Manual Operator Assessment',
      criteria: 'Manual evaluation based on Max Temperature (Tj)'
    };
  });
}

// Function to update test results in Test Data based on bypass diode results
function updateTestDataWithBypassDiodeResults(testData, bypassDiodeResults) {
  return testData.map(testRow => {
    // Check if this is a bypass diode test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('BYPASS')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching bypass diode results for this vendor and BOM
      const bypassDiodeResult = bypassDiodeResults.find(result => 
        result.vendorName === vendorName && result.bom === bomType
      );
      
      if (bypassDiodeResult) {
        // Update the test result with the manual assessment
        testRow['TEST RESULT'] = bypassDiodeResult.testResult;
        testRow['BYPASS_DIODE_CALCULATION_DETAILS'] = {
          temperature: {
            maxTemperatureTj: bypassDiodeResult.maxTemperatureTj,
            unit: bypassDiodeResult.temperatureUnit,
            hasValidReading: bypassDiodeResult.hasValidTemperature
          },
          assessment: {
            result: bypassDiodeResult.testResult,
            method: bypassDiodeResult.assessmentMethod,
            criteria: bypassDiodeResult.criteria,
            assessedBy: bypassDiodeResult.testedBy,
            assessmentDate: bypassDiodeResult.testDate
          },
          notes: bypassDiodeResult.notes,
          overallResult: bypassDiodeResult.testResult
        };
        
        console.log(`Updated bypass diode test result for ${vendorName} ${bomType}: ${bypassDiodeResult.testResult} (Tj: ${bypassDiodeResult.maxTemperatureTj}°C)`);
      } else {
        console.log(`No bypass diode assessment found for ${vendorName} ${bomType}`);
      }
    }
    
    return testRow;
  });
}

// Alternative function for more flexible column name handling
function processBypassDiodeResultsFlexible(bypassDiodeData) {
  return bypassDiodeData.map((row, index) => {
    const bom = row['BOM'] || '';
    const vendorName = row['VENDOR NAME'] || '';
    
    // Handle multiple possible column names for temperature
    let maxTemperatureTj = 0;
    const possibleTempColumns = [
      'MAX TEMPERATURE OF DIODE(Tj)', 
      'MAX TEMPERATURE OF DIODE (Tj)',
      'MAX TEMPERATURE OF DIODE',
      'TEMPERATURE (Tj)',
      'Tj',
      'MAX TEMP',
      'DIODE TEMPERATURE',
      'JUNCTION TEMPERATURE'
    ];
    
    for (const colName of possibleTempColumns) {
      if (row[colName] !== undefined && row[colName] !== null && row[colName] !== '') {
        const parsedTemp = parseFloat(row[colName]);
        if (!isNaN(parsedTemp) && parsedTemp > 0) {
          maxTemperatureTj = parsedTemp;
          break;
        }
      }
    }
    
    // Handle multiple possible column names for test result
    let testResult = 'Pending';
    const possibleResultColumns = [
      'TEST RESULT', 'RESULT', 'STATUS', 'PASS/FAIL', 'OUTCOME',
      'ASSESSMENT', 'MANUAL RESULT', 'OPERATOR RESULT'
    ];
    
    for (const colName of possibleResultColumns) {
      if (row[colName] !== undefined && row[colName] !== null && row[colName] !== '') {
        const result = String(row[colName]).toUpperCase();
        if (result === 'PASS' || result === 'FAIL' || result === 'P' || result === 'F') {
          testResult = result === 'P' ? 'PASS' : result === 'F' ? 'FAIL' : result;
          break;
        }
      }
    }
    
    return {
      id: index + 1,
      bom,
      vendorName,
      maxTemperatureTj: Math.round(maxTemperatureTj * 100) / 100,
      testResult,
      hasValidTemperature: maxTemperatureTj > 0,
      temperatureUnit: '°C',
      notes: row['NOTES'] || row['REMARKS'] || row['COMMENTS'] || '',
      testDate: row['TEST DATE'] || row['DATE'] || '',
      testedBy: row['TESTED BY'] || row['OPERATOR'] || '',
      assessmentMethod: 'Manual Operator Assessment',
      criteria: 'Manual evaluation based on Max Temperature (Tj) and operational requirements'
    };
  });
}


// Improved file check function to provide more details
function checkExcelFile(filename) {
  const excelFilePath = path.join(__dirname, 'data', filename);
  const exists = fs.existsSync(excelFilePath);
  
  let fileInfo = {
    exists,
    path: excelFilePath,
    size: null,
    lastModified: null,
    lastChecked: new Date().toISOString()
  };
  
  if (exists) {
    try {
      const stats = fs.statSync(excelFilePath);
      fileInfo.size = stats.size;
      fileInfo.lastModified = stats.mtime;
      console.log(`Excel file found at ${excelFilePath}, size: ${stats.size} bytes, last modified: ${stats.mtime}`);
    } catch (error) {
      console.error(`Error getting file stats: ${error.message}`);
    }
  } else {
    console.warn(`Excel file not found at ${excelFilePath}`);
  }
  
  return fileInfo;
}

// MODIFIED API endpoint for test data - WITH SHRINKAGE INTEGRATION
// Replace the existing app.get('/api/test-data', ...) endpoint with this:

app.get('/api/test-data', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/test-data from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      query: req.query,
      headers: {
        'user-agent': req.headers['user-agent'],
        'cache-control': req.headers['cache-control']
      }
    };
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        fileInfo,
        requestInfo
      });
    }
    
    // Read the Excel file with force reload
    const excelFilePath = fileInfo.path;
    
    // Use try/catch specifically for file reading
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message,
        fileInfo,
        requestInfo
      });
    }
    
    // Log available sheets
    console.log('Available sheets in workbook:', workbook.SheetNames);
    
    // Read the "Test Data" sheet
    const testDataSheetName = 'Test Data';
    const worksheet = workbook.Sheets[testDataSheetName];
    if (!worksheet) {
      console.error(`Sheet "${testDataSheetName}" not found. Available sheets:`, workbook.SheetNames);
      return res.status(404).json({ 
        error: `${testDataSheetName} sheet not found in Excel file`,
        availableSheets: workbook.SheetNames,
        fileInfo,
        requestInfo
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${testDataSheetName} sheet`);
    
    // Check for shrinkage tests and read Shrinkage if needed
    // UPDATED: Fix Shrinkage sheet reading in server.js
// Replace the shrinkage reading section in /api/test-data endpoint

// Check for shrinkage tests and read Shrinkage if needed
// REPLACE this section in your /api/test-data endpoint:
// Starting from "Check for shrinkage tests and read Shrinkage if needed"

const hasShrinkageTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('SHRINKAGE')
);

let shrinkageResults = [];
if (hasShrinkageTests) {
  const shrinkageSheetName = 'Shrinkage';
  const shrinkageSheet = workbook.Sheets[shrinkageSheetName];
  
  if (shrinkageSheet) {
    console.log('📊 Found shrinkage tests, reading Shrinkage sheet...');
    
    // FIXED: Read the sheet structure more carefully
    try {
      // Read raw data to understand the structure
      const fullShrinkageData = xlsx.utils.sheet_to_json(shrinkageSheet, { 
        header: 1, // Get raw array format
        defval: ''
      });
      
      console.log('Shrinkage sheet structure:');
      console.log('Row 0 (first row):', fullShrinkageData[0]);
      console.log('Row 1 (second row):', fullShrinkageData[1]);
      console.log('Row 2 (third row):', fullShrinkageData[2]);
      console.log('Row 3 (data row):', fullShrinkageData[3]);
      
      // Try reading from row 3 (where data starts) with auto-generated headers
      const shrinkageRawData = xlsx.utils.sheet_to_json(shrinkageSheet, { 
        range: 3, // Start from row 4 (0-indexed = 3) where data begins
        defval: ''
        // Let xlsx auto-generate headers from the structure
      });
      
      console.log(`Read ${shrinkageRawData.length} shrinkage data rows`);
      if (shrinkageRawData.length > 0) {
        console.log('Available columns:', Object.keys(shrinkageRawData[0]));
        console.log('Sample shrinkage row:', shrinkageRawData[0]);
      }
      
      // Filter for valid data rows
      const validShrinkageData = shrinkageRawData.filter(row => {
        // Check multiple possible column names for vendor
        const vendorName = row['VENDOR NAME'] || row['Vendor Name'] || row['VENDOR_NAME'] || 
                          row['vendor name'] || row['Vendor'] || '';
        
        // Check multiple possible column names for type
        const encapsulantType = row['ENCAPSULANT TYPE'] || row['Encapsulant Type'] || 
                               row['ENCAPSULANT_TYPE'] || row['encapsulant type'] || 
                               row['Type'] || row['TYPE'] || '';
        
        const hasVendor = vendorName && vendorName.trim() !== '';
        const hasValidType = encapsulantType && 
          (encapsulantType.includes('FRONT EPE') || encapsulantType.includes('BACK EVA') ||
           encapsulantType.includes('Front EPE') || encapsulantType.includes('Back EVA'));
        
        console.log(`Row validation - Vendor: "${vendorName}", Type: "${encapsulantType}", Valid: ${hasVendor && hasValidType}`);
        
        return hasVendor && hasValidType;
      });
      
      console.log(`✅ Found ${validShrinkageData.length} valid shrinkage data rows`);
      
      if (validShrinkageData.length > 0) {
        shrinkageResults = calculateShrinkageResults(validShrinkageData);
        console.log(`🎯 Calculated shrinkage results for ${shrinkageResults.length} entries`);
        
        // Log the calculated results for debugging
        shrinkageResults.forEach((result, index) => {
          if (index < 3) {
            console.log(`Shrinkage result ${index + 1}:`, {
              vendor: result.vendorName,
              type: result.encapsulantType,
              tdDiff: result.tdDifference,
              mdDiff: result.mdDifference,
              final: result.finalStatus
            });
          }
        });
      } else {
        console.log('⚠️ No valid shrinkage data found after filtering');
        console.log('Available columns in first row:', Object.keys(shrinkageRawData[0] || {}));
      }
    } catch (shrinkageError) {
      console.error('❌ Error processing shrinkage sheet:', shrinkageError);
      console.log('Available sheet names:', workbook.SheetNames);
    }
  } else {
    console.warn('❌ Shrinkage tests found but Shrinkage sheet not available for calculations');
  }
}

    // Add this section after the shrinkage calculation section in the /api/test-data endpoint:

    // Check for adhesion tests and read Adhesion sheet if needed
    // UPDATED: Fix Adhesion sheet reading in server.js
// Replace the adhesion reading section in /api/test-data endpoint

// Check for adhesion tests and read Adhesion sheet if needed
const hasAdhesionTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('ADHESION')
);

let adhesionResults = [];
if (hasAdhesionTests) {
  const adhesionSheetName = 'Adhesion';
  const adhesionSheet = workbook.Sheets[adhesionSheetName];
  
  if (adhesionSheet) {
    console.log('Found adhesion tests, reading Adhesion sheet...');
    
    // Read from row 4 since headers are in row 2-3, data starts from row 4
    const adhesionRawData = xlsx.utils.sheet_to_json(adhesionSheet, {
      range: 3, // Start from row 4 (0-indexed = 3)
      defval: ''
    });
    
    console.log('Sample adhesion data structure:', Object.keys(adhesionRawData[0] || {}));
    console.log('First adhesion row:', adhesionRawData[0]);
    
    // Filter for valid data
    const validAdhesionData = adhesionRawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    console.log(`Found ${validAdhesionData.length} valid adhesion data rows`);
    
    if (validAdhesionData.length > 0) {
      adhesionResults = calculateAdhesionResults(validAdhesionData);
      console.log(`Calculated adhesion results for ${adhesionResults.length} entries`);
      
      // Log calculated results for debugging
      adhesionResults.forEach((result, index) => {
        if (index < 3) {
          console.log(`Adhesion result ${index + 1}:`, {
            vendor: result.vendorName,
            bom: result.bom,
            prePctGlassAvg: result.prePctGlassToEncapAvg,
            prePctBacksheetAvg: result.prePctBacksheetToEncapAvg,
            postPctGlassAvg: result.postPctGlassToEncapAvg,
            postPctBacksheetAvg: result.postPctBacksheetToEncapAvg,
            final: result.finalStatus
          });
        }
      });
    }
  } else {
    console.warn('Adhesion tests found but Adhesion sheet not available for calculations');
  }
}


    // Add this section to the main /api/test-data endpoint after the adhesion section:

// Check for tensile strength tests and read Tensile Strength sheet if needed
const hasTensileStrengthTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('TENSILE')
);

let tensileStrengthResults = [];
if (hasTensileStrengthTests) {
  // Read Tensile Strength sheet for tensile data
  const tensileStrengthSheetName = 'Tensile Strength';
  const tensileStrengthSheet = workbook.Sheets[tensileStrengthSheetName];
  
  if (tensileStrengthSheet) {
    console.log('Found tensile strength tests, reading Tensile Strength sheet for tensile data...');
    const tensileStrengthRawData = xlsx.utils.sheet_to_json(tensileStrengthSheet);
    
    // Log the raw data structure for debugging
    if (tensileStrengthRawData.length > 0) {
      console.log('Sample tensile strength raw data structure:', Object.keys(tensileStrengthRawData[0]));
      console.log('First tensile strength row:', tensileStrengthRawData[0]);
    }
    
    // Filter out empty rows and header rows
    const validTensileStrengthData = tensileStrengthRawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    console.log(`Found ${validTensileStrengthData.length} valid tensile strength data rows`);
    
    if (validTensileStrengthData.length > 0) {
      tensileStrengthResults = calculateTensileStrengthResults(validTensileStrengthData);
      console.log(`Calculated tensile strength results for ${tensileStrengthResults.length} entries`);
      
      // Log calculated results for debugging
      tensileStrengthResults.forEach((result, index) => {
        if (index < 3) { // Log first 3 results for debugging
          console.log(`Tensile strength result ${index + 1}:`, {
            vendor: result.vendorName,
            bom: result.bom,
            breakValue: result.breakValue,
            breakStatus: result.breakStatus,
            elongationPercent: result.changeInElongationPercent,
            elongationStatus: result.elongationStatus,
            final: result.finalStatus
          });
        }
      });
    }
  } else {
    console.warn('Tensile strength tests found but Tensile Strength sheet not available for calculations');
  }
}

// Add this section to your main /api/test-data endpoint after the tensile strength section:

// Check for GSM tests and read GSM sheet if needed
const hasGSMTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('GSM')
);

let gsmResults = [];
if (hasGSMTests) {
  const gsmSheetName = 'GSM';
  const gsmSheet = workbook.Sheets[gsmSheetName];
  
  if (gsmSheet) {
    console.log('Found GSM tests, reading GSM sheet...');
    
    // Read from row 2 since your headers are in row 2
    const gsmRawData = xlsx.utils.sheet_to_json(gsmSheet, {
      range: 1, // Start from row 2 (0-indexed)
      defval: ''
    });
    
    console.log('Sample GSM data structure:', Object.keys(gsmRawData[0] || {}));
    
    // Filter for valid data
    const validGSMData = gsmRawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    console.log(`Found ${validGSMData.length} valid GSM data rows`);
    
    if (validGSMData.length > 0) {
      gsmResults = calculateGSMResults(validGSMData);
      console.log(`Calculated GSM results for ${gsmResults.length} entries`);
      
      // Log calculated results for debugging
      gsmResults.forEach((result, index) => {
        if (index < 3) { // Log first 3 results for debugging
          console.log(`GSM result ${index + 1}:`, {
            vendor: result.vendorName,
            bom: result.bom,
            category: result.category,
            average: result.average,
            range: result.rangeName,
            final: result.finalStatus
          });
        }
      });
    }
  } else {
    console.warn('GSM tests found but GSM sheet not available for calculations');
  }
}

// Add this section to your main /api/test-data endpoint after the GSM section:

// Check for resistance tests and read Resistance sheet if needed
const hasResistanceTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('RESISTANCE')
);

let resistanceResults = [];
if (hasResistanceTests) {
  const resistanceSheetName = 'Resistance';
  const resistanceSheet = workbook.Sheets[resistanceSheetName];
  
  if (resistanceSheet) {
    console.log('Found resistance tests, reading Resistance sheet...');
    
    // Read from row 2 since your headers are in row 2
    const resistanceRawData = xlsx.utils.sheet_to_json(resistanceSheet, {
      range: 1, // Start from row 2 (0-indexed)
      defval: ''
    });
    
    console.log('Sample resistance data structure:', Object.keys(resistanceRawData[0] || {}));
    
    // Filter for valid data
    const validResistanceData = resistanceRawData.filter(row => 
      row['VENDOR NAME'] && row['BOM'] && row['TYPE']
    );
    
    console.log(`Found ${validResistanceData.length} valid resistance data rows`);
    
    if (validResistanceData.length > 0) {
      resistanceResults = processResistanceResults(validResistanceData);
      console.log(`Processed resistance results for ${resistanceResults.length} entries`);
      
      // Log processed results for debugging
      resistanceResults.forEach((result, index) => {
        if (index < 3) { // Log first 3 results for debugging
          console.log(`Resistance result ${index + 1}:`, {
            vendor: result.vendorName,
            bom: result.bom,
            type: result.type,
            measuredValue: result.measuredValue,
            result: result.testResult
          });
        }
      });
      
      // Group by vendor and BOM to show summary
      const resistanceSummary = {};
      resistanceResults.forEach(result => {
        const key = `${result.vendorName}-${result.bom}`;
        if (!resistanceSummary[key]) {
          resistanceSummary[key] = {
            vendor: result.vendorName,
            bom: result.bom,
            busRibbon: null,
            interconnectRibbon: null
          };
        }
        
        if (result.type.includes('BUS')) {
          resistanceSummary[key].busRibbon = {
            value: result.measuredValue,
            result: result.testResult
          };
        } else if (result.type.includes('INTERCONNECT')) {
          resistanceSummary[key].interconnectRibbon = {
            value: result.measuredValue,
            result: result.testResult
          };
        }
      });
      
      console.log('Resistance test summary by vendor/BOM:', resistanceSummary);
    }
  } else {
    console.warn('Resistance tests found but Resistance sheet not available for processing');
  }
}

// Add this section to your main /api/test-data endpoint after the resistance section:

// Check for bypass diode tests and read BYPASS DIODE TEST sheet if needed
const hasBypassDiodeTests = rawData.some(row => 
  row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('BYPASS')
);

let bypassDiodeResults = [];
if (hasBypassDiodeTests) {
  const bypassDiodeSheetName = 'BYPASS DIODE TEST';
  const bypassDiodeSheet = workbook.Sheets[bypassDiodeSheetName];
  
  if (bypassDiodeSheet) {
    console.log('Found bypass diode tests, reading BYPASS DIODE TEST sheet...');
    
    // Read from row 2 since your headers are in row 2
    const bypassDiodeRawData = xlsx.utils.sheet_to_json(bypassDiodeSheet, {
      range: 1, // Start from row 2 (0-indexed)
      defval: ''
    });
    
    console.log('Sample bypass diode data structure:', Object.keys(bypassDiodeRawData[0] || {}));
    
    // Filter for valid data
    const validBypassDiodeData = bypassDiodeRawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    console.log(`Found ${validBypassDiodeData.length} valid bypass diode data rows`);
    
    if (validBypassDiodeData.length > 0) {
      bypassDiodeResults = processBypassDiodeResults(validBypassDiodeData);
      console.log(`Processed bypass diode results for ${bypassDiodeResults.length} entries`);
      
      // Log processed results for debugging
      bypassDiodeResults.forEach((result, index) => {
        if (index < 3) { // Log first 3 results for debugging
          console.log(`Bypass Diode result ${index + 1}:`, {
            vendor: result.vendorName,
            bom: result.bom,
            maxTemperatureTj: result.maxTemperatureTj,
            testResult: result.testResult,
            hasValidTemp: result.hasValidTemperature
          });
        }
      });
      
      // Summary of bypass diode test results
      const bypassDiodeSummary = {
        total: bypassDiodeResults.length,
        passed: bypassDiodeResults.filter(r => r.testResult === 'PASS').length,
        failed: bypassDiodeResults.filter(r => r.testResult === 'FAIL').length,
        pending: bypassDiodeResults.filter(r => r.testResult === 'Pending').length,
        withValidTemperature: bypassDiodeResults.filter(r => r.hasValidTemperature).length
      };
      
      console.log('Bypass Diode test summary:', bypassDiodeSummary);
    }
  } else {
    console.warn('Bypass diode tests found but BYPASS DIODE TEST sheet not available for processing');
  }
}
    
    // Read the Standard Test Times sheet for reference
    const standardsSheetName = 'Standard Test Times';
    const standardsSheet = workbook.Sheets[standardsSheetName];
    let standardsData = [];
    
    if (standardsSheet) {
      standardsData = xlsx.utils.sheet_to_json(standardsSheet);
      console.log(`Processed ${standardsData.length} rows from ${standardsSheetName} sheet`);
    } else {
      console.warn(`Sheet "${standardsSheetName}" not found. Will use hardcoded values for efficiency calculations.`);
    }
    
    // Create a lookup for standard test durations from Excel data
    const standardDurations = {};
    standardsData.forEach(item => {
      const testName = item['TEST NAME'];
      const duration = item['STANDARD DURATION (DAYS)'];
      
      if (testName && duration !== undefined) {
        standardDurations[testName] = parseFloat(duration);
        console.log(`Standard duration for "${testName}": ${duration} days (from Excel)`);
      }
    });
    
    // Fallback to hardcoded durations for any missing values
    for (const [testName, duration] of Object.entries(HARDCODED_STD_DURATIONS)) {
      if (standardDurations[testName] === undefined) {
        standardDurations[testName] = duration;
        console.log(`Standard duration for "${testName}": ${duration} days (from hardcoded values)`);
      }
    }
    
    // Default duration if not found in either source
    const DEFAULT_STD_DURATION = 2; // 2 days
    
    // Update test data with shrinkage results if available
    let updatedRawData = rawData;
    if (shrinkageResults.length > 0) {
      updatedRawData = updateTestDataWithShrinkageResults(rawData, shrinkageResults);
      console.log('Updated test data with shrinkage calculation results');
      
      // Log how many shrinkage tests were updated
      const updatedShrinkageTests = updatedRawData.filter(row => 
        row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('SHRINKAGE') && 
        row['SHRINKAGE_CALCULATION_DETAILS']
      );
      console.log(`Updated ${updatedShrinkageTests.length} shrinkage test results`);
    }

    // And add this after the shrinkage update section:

    // Update test data with adhesion results if available
    if (adhesionResults.length > 0) {
      updatedRawData = updateTestDataWithAdhesionResults(updatedRawData, adhesionResults);
      console.log('Updated test data with adhesion calculation results');
      
      // Log how many adhesion tests were updated
      const updatedAdhesionTests = updatedRawData.filter(row => 
        row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('ADHESION') && 
        row['ADHESION_CALCULATION_DETAILS']
      );
      console.log(`Updated ${updatedAdhesionTests.length} adhesion test results`);
    }


    // And add this section after the adhesion update section:

// Update test data with tensile strength results if available
if (tensileStrengthResults.length > 0) {
  updatedRawData = updateTestDataWithTensileStrengthResults(updatedRawData, tensileStrengthResults);
  console.log('Updated test data with tensile strength calculation results');
  
  // Log how many tensile strength tests were updated
  const updatedTensileStrengthTests = updatedRawData.filter(row => 
    row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('TENSILE') && 
    row['TENSILE_CALCULATION_DETAILS']
  );
  console.log(`Updated ${updatedTensileStrengthTests.length} tensile strength test results`);
}

// Update test data with GSM results if available (add this after tensile strength update)
if (gsmResults.length > 0) {
  updatedRawData = updateTestDataWithGSMResults(updatedRawData, gsmResults);
  console.log('Updated test data with GSM calculation results');
  
  // Log how many GSM tests were updated
  const updatedGSMTests = updatedRawData.filter(row => 
    row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('GSM') && 
    row['GSM_CALCULATION_DETAILS']
  );
  console.log(`Updated ${updatedGSMTests.length} GSM test results`);
}

if (resistanceResults.length > 0) {
  updatedRawData = updateTestDataWithResistanceResults(updatedRawData, resistanceResults);
  console.log('Updated test data with resistance results');
  
  // Log how many resistance tests were updated
  const updatedResistanceTests = updatedRawData.filter(row => 
    row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('RESISTANCE') && 
    row['RESISTANCE_CALCULATION_DETAILS']
  );
  console.log(`Updated ${updatedResistanceTests.length} resistance test results`);
}

// Update test data with bypass diode results if available (add this after resistance update)
if (bypassDiodeResults.length > 0) {
  updatedRawData = updateTestDataWithBypassDiodeResults(updatedRawData, bypassDiodeResults);
  console.log('Updated test data with bypass diode results');
  
  // Log how many bypass diode tests were updated
  const updatedBypassDiodeTests = updatedRawData.filter(row => 
    row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('BYPASS') && 
    row['BYPASS_DIODE_CALCULATION_DETAILS']
  );
  console.log(`Updated ${updatedBypassDiodeTests.length} bypass diode test results`);
}

    
    // Process the data to match the dashboard's expected format
    const processedData = updatedRawData.map((row, index) => {
      const grnTimeValue = row['GRN GENERATION TIME'];
      const testStartValue = row['TEST START DATE AND TIME'];
      const testEndValue = row['TEST END DATE AND TIME'];
      
      const grnTime = parseExcelDate(grnTimeValue);
      const startTime = parseExcelDate(testStartValue);
      const endTime = parseExcelDate(testEndValue);
      
      if (index < 5) {
        console.log(`Row ${index + 1}:`, {
          testName: row['TEST NAME'],
          vendor: row['VENDOR NAME'],
          testResult: row['TEST RESULT'],
          hasShrinkageDetails: !!row['SHRINKAGE_CALCULATION_DETAILS']
        });
      }
      
      let actualDuration = 0;
      if (startTime && endTime) {
        const diffTime = Math.abs(endTime - startTime);
        actualDuration = Math.max(1, Math.ceil(diffTime / (1000 * 60 * 60 * 24)));
      }
      
      const testName = row['TEST NAME'] || '';
      const standardDuration = standardDurations[testName] || DEFAULT_STD_DURATION;
      
      let efficiency = null;
      if (actualDuration > 0 && standardDuration > 0) {
        efficiency = Math.round((standardDuration / actualDuration) * 100);
      }
      
      return {
        id: index + 1,
        bom: row['BOM'] || '',
        test: testName,
        vendor: row['VENDOR NAME'] || '',
        grnTime: grnTime,
        startTime: startTime,
        endTime: endTime,
        status: row['STATUS'] || 'Pending',
        result: row['TEST RESULT'] || 'Pending',
        actualDuration: actualDuration,
        standardDuration: standardDuration,
        efficiency: efficiency,
        shrinkageDetails: row['SHRINKAGE_CALCULATION_DETAILS'] || null
      };
    });
    
    res.json(processedData);
  } catch (error) {
    console.error('Error processing request:', error);
    res.status(500).json({ 
      error: 'Failed to process request', 
      details: error.message,
      stack: error.stack 
    });
  }
});

// API endpoint for line trials data - NOW WITH AUTHENTICATION
app.get('/api/line-trials', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/line-trials from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      query: req.query,
      headers: {
        'user-agent': req.headers['user-agent'],
        'cache-control': req.headers['cache-control']
      }
    };
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Line_Trials.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Line Trials Excel file not found',
        message: 'The Line Trials file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        fileInfo,
        requestInfo
      });
    }
    
    // Read the Excel file with force reload
    const excelFilePath = fileInfo.path;
    
    // Use try/catch specifically for file reading
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Line Trials Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Line Trials Excel file',
        details: readError.message,
        fileInfo,
        requestInfo
      });
    }
    
    // Log available sheets
    console.log('Available sheets in Line Trials workbook:', workbook.SheetNames);
    
    // Assuming the main sheet is the first one or named "Line Trials"
    const sheetName = workbook.SheetNames[1]; // First sheet or specify the exact name
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      console.error(`Sheet not found. Available sheets:`, workbook.SheetNames);
      return res.status(404).json({ 
        error: `Sheet not found in Line Trials Excel file`,
        availableSheets: workbook.SheetNames,
        fileInfo,
        requestInfo
      });
    }
    
    // Convert to JSON
    const lineTrialsData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${lineTrialsData.length} rows from Line Trials sheet`);
    
    // Process the data as needed for your frontend
    const processedData = lineTrialsData.map((row, index) => {
      const startDate = parseExcelDate(row['START_DATE']);
      const endDate = parseExcelDate(row['END_DATE']);
      // Assuming your Excel has columns: vendor, material, status, remarks
      return {
        id: index + 1,
        vendor: row['VENDOR'] || '',
        bomUnderTrial: row['BOM_UNDER_TRIAL'] || '',
        status: row['STATUS'] || '',
        startDate: startDate,
        endDate: endDate,
        remarks: row['REMARKS'] || '',
        orderno: row['ORDER_NO.'] || ''
      };
    });
    
    const sortedData = processedData.sort((a, b) => {
      if ((a.status === 'Completed' || a.status === 'Failed') && 
          (b.status === 'Completed' || b.status === 'Failed')) {
        // Both items are completed or failed, sort by end date
        return b.endDate - a.endDate;
      } else if ((a.status === 'Completed' || a.status === 'Failed') && 
                !(b.status === 'Completed' || b.status === 'Failed')) {
        // A is completed or failed, B is in progress or pending
        return -1; // A comes first
      } else if (!(a.status === 'Completed' || a.status === 'Failed') && 
                (b.status === 'Completed' || b.status === 'Failed')) {
        // A is in progress or pending, B is completed or failed
        return 1; // B comes first
      } else {
        // Both are in progress or pending, sort by start date
        return b.startDate - a.startDate;
      }
    });
    
    res.json(sortedData);
  } catch (error) {
    console.error('Error processing Line Trials request:', error);
    res.status(500).json({ 
      error: 'Failed to process Line Trials request', 
      details: error.message,
      stack: error.stack 
    });
  }
});

// API endpoint for certifications data - NOW WITH AUTHENTICATION
app.get('/api/certifications', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/certifications from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      query: req.query,
      headers: {
        'user-agent': req.headers['user-agent'],
        'cache-control': req.headers['cache-control']
      }
    };
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Certifications.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Certifications Excel file not found',
        message: 'The Certifications file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        fileInfo,
        requestInfo
      });
    }
    
    // Read the Excel file with force reload
    const excelFilePath = fileInfo.path;
    
    // Use try/catch specifically for file reading
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Certifications Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Certifications Excel file',
        details: readError.message,
        fileInfo,
        requestInfo
      });
    }
    
    // Log available sheets
    console.log('Available sheets in Certifications workbook:', workbook.SheetNames);
    
    // Process each sheet for certification data
    const certificationDetails = {
      completed: [],
      inProcess: [],
      pending: []
    };
    
    // Process Completed sheet
    if (workbook.SheetNames.includes('COMPLETED')) {
      const completedSheet = workbook.Sheets['COMPLETED'];
      const completedData = xlsx.utils.sheet_to_json(completedSheet);
      
      completedData.forEach(row => {
        const completionDate = parseExcelDate(row['COMPLETION_DATE']);
        
        certificationDetails.completed.push({
          product: row['PRODUCT'] || '',
          certName: row['CERTIFICATION'] || '',
          agency: row['AGENCY'] || '',
          completionDate: completionDate,
          notes: row['NOTE'] || '',
          wattpeak: row['WATTPEAK'] || '',
          plantName: row['PLANT_NAME'] || ''
        });
      });
      
      // Sort by completion date (newest first)
      certificationDetails.completed.sort((a, b) => b.completionDate - a.completionDate);
    }
    
    // Process In Process sheet
    if (workbook.SheetNames.includes('In Process')) {
      const inProcessSheet = workbook.Sheets['In Process'];
      const inProcessData = xlsx.utils.sheet_to_json(inProcessSheet);
      
      inProcessData.forEach(row => {
        const startDate = parseExcelDate(row['START_DATE']);
        const expectedCompletion = parseExcelDate(row['EXPECTED_COMPLETION']);
        
        certificationDetails.inProcess.push({
          product: row['PRODUCT'] || '',
          certName: row['CERTIFICATION'] || '',
          agency: row['AGENCY'] || '',
          startDate: startDate,
          expectedCompletion: expectedCompletion,
          status: row['STATUS'] || '',
          wattpeak: row['WATTPEAK'] || '',
          plantName: row['PLANT_NAME'] || ''
        });
      });
      
      // Sort by start date (newest first)
      certificationDetails.inProcess.sort((a, b) => b.startDate - a.startDate);
    }
    
    // Process Pending sheet
    if (workbook.SheetNames.includes('Pending')) {
      const pendingSheet = workbook.Sheets['Pending'];
      const pendingData = xlsx.utils.sheet_to_json(pendingSheet);
      
      pendingData.forEach(row => {
        const plannedStart = parseExcelDate(row['PLANNED_START']);
        
        certificationDetails.pending.push({
          product: row['PRODUCT'] || '',
          certName: row['CERTIFICATION'] || '',
          agency: row['AGENCY'] || '',
          plannedStart: plannedStart,
          priority: row['PRIORITY'] || '',
          wattpeak: row['WATTPEAK'] || '',
          plantName: row['PLANT_NAME'] || ''
        });
      });
      
      // Sort by planned start date (newest first)
      certificationDetails.pending.sort((a, b) => b.plannedStart - a.plannedStart);
    }
    
    // Create summary data by product
    const certificationData = [];
    const productSummary = {};
    
    // Process all certification data to build product summary
    ['completed', 'inProcess', 'pending'].forEach(status => {
      certificationDetails[status].forEach(cert => {
        const product = cert.product;
        
        if (!productSummary[product]) {
          productSummary[product] = {
            product: product,
            completed: 0,
            inProcess: 0,
            pending: 0
          };
        }
        
        // Increment the appropriate counter
        productSummary[product][status]++;
      });
    });
    
    // Convert summary to array
    for (const product in productSummary) {
      certificationData.push(productSummary[product]);
    }
    
    res.json({
      certificationData,
      certificationDetails
    });
  } catch (error) {
    console.error('Error processing Certifications request:', error);
    res.status(500).json({ 
      error: 'Failed to process Certifications request', 
      details: error.message
    });
  }
});

// API endpoint for chamber data - NOW WITH AUTHENTICATION
app.get('/api/chamber-data', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/chamber-data from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    const requestInfo = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      query: req.query,
      headers: {
        'user-agent': req.headers['user-agent'],
        'cache-control': req.headers['cache-control']
      }
    };
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Chamber_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Chamber Tests Excel file not found',
        message: 'The Chamber Tests file has not been synced yet from OneDrive.',
        fileInfo,
        requestInfo
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Chamber Tests Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Chamber Tests Excel file',
        details: readError.message,
        fileInfo,
        requestInfo
      });
    }
    
    console.log('Available sheets in Chamber Tests workbook:', workbook.SheetNames);
    
    // Use the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) {
      console.error(`Sheet not found. Available sheets:`, workbook.SheetNames);
      return res.status(404).json({ 
        error: `Sheet not found in Chamber Tests Excel file`,
        availableSheets: workbook.SheetNames,
        fileInfo,
        requestInfo
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from Chamber Tests sheet`);
    
    if (rawData.length === 0) {
      return res.json([]);
    }
    
    // Log total columns for debugging
    if (rawData.length > 0) {
      const totalColumns = Object.keys(rawData[0]).length;
      console.log(`Total columns in data: ${totalColumns}`);
    }
    
    // Calculate all derived fields with the simplified approach
    const processedData = calculateChamberData(rawData, worksheet);
    
    // Sort by status and date (active tests first, then by start date)
    const sortedData = processedData.sort((a, b) => {
      // Completed tests go to the end
      if (a.remainingHr === 'DONE' && b.remainingHr !== 'DONE') return 1;
      if (a.remainingHr !== 'DONE' && b.remainingHr === 'DONE') return -1;
      
      // Among non-completed, sort by start date (most recent first)
      if (a.startDate && b.startDate) {
        return new Date(b.startDate) - new Date(a.startDate);
      }
      
      return 0;
    });
    
    console.log(`Returning ${sortedData.length} processed chamber test records`);
    res.json(sortedData);
    
  } catch (error) {
    console.error('Error processing Chamber Tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process Chamber Tests request', 
      details: error.message,
      stack: error.stack 
    });
  }
});


// Add these API endpoints to your existing server.js file

// API endpoint for R&D todos data - WITH AUTHENTICATION
app.get('/api/todos', authenticateMicrosoftToken, async (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`📥 Loading todos from database for ${userMap[userEmail] || userEmail}`);
    
    // Query todos with their updates
    const todos = await Todo.findAll({
      include: [{
        model: TodoUpdate,
        as: 'updates',
        required: false, // LEFT JOIN
        order: [['createdAt', 'ASC']]
      }],
      order: [['createdAt', 'DESC']]
    });
    
    console.log(`📊 Found ${todos.length} todos in database`);
    
    // Transform to frontend format
    const processedTodos = todos.map(todo => {
      const updates = todo.updates || [];
      
      // Get unique meeting dates from updates
      const meetingDates = [...new Set(
        updates
          .filter(update => update.meetingDate)
          .map(update => update.meetingDate.toISOString().split('T')[0])
      )];
      
      return {
        id: todo.id,
        issue: todo.issue,
        responsibility: todo.responsibility.split(',').map(r => r.trim()).filter(r => r),
        status: todo.status,
        priority: todo.priority,
        category: todo.category,
        dueDate: todo.dueDate ? todo.dueDate.toISOString().split('T')[0] : null,
        createdDate: todo.createdAt.toISOString().split('T')[0],
        updates: updates.map(update => ({
          date: update.createdAt.toISOString().split('T')[0],
          status: update.status,
          note: update.note,
          meetingDate: update.meetingDate ? update.meetingDate.toISOString().split('T')[0] : null
        })),
        meetingDates: meetingDates
      };
    });
    
    console.log(`✅ Returning ${processedTodos.length} processed todos`);
    res.json(processedTodos);
    
  } catch (error) {
    console.error('❌ Error loading todos from database:', error);
    res.status(500).json({
      error: 'Failed to load todos from database',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// API endpoint to add new todo - WITH AUTHENTICATION
// Replace your existing POST /api/todos endpoint with this working version
// that actually writes to the Excel file

// Add new todo endpoint (REPLACE EXISTING)
app.post('/api/todos', authenticateMicrosoftToken, async (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const todoData = req.body;
    
    console.log(`📝 Creating new todo in database for ${userMap[userEmail] || userEmail}:`, todoData);
    
    // Validation
    if (!todoData.issue || !todoData.issue.trim()) {
      return res.status(400).json({
        error: 'Invalid issue field',
        message: 'Issue description is required and must be non-empty'
      });
    }
    
    if (!todoData.responsibility || !Array.isArray(todoData.responsibility) || todoData.responsibility.length === 0) {
      return res.status(400).json({
        error: 'Invalid responsibility field',
        message: 'At least one team member must be assigned'
      });
    }
    
    // Create todo in database
    const newTodo = await Todo.create({
      issue: todoData.issue.trim(),
      responsibility: todoData.responsibility.join(','),
      status: 'Pending',
      priority: todoData.priority || 'Medium',
      category: todoData.category || 'General',
      dueDate: todoData.dueDate ? new Date(todoData.dueDate) : null,
      createdBy: userEmail
    });
    
    console.log(`✅ Todo created in database with ID: ${newTodo.id}`);
    
    // Create Excel backup asynchronously
    setTimeout(async () => {
      try {
        await createExcelBackupFromDatabase();
        console.log('✅ Excel backup created after todo creation');
      } catch (error) {
        console.error('❌ Excel backup failed:', error);
      }
    }, 1000);
    
    // Return formatted response
    const responseData = {
      id: newTodo.id,
      issue: newTodo.issue,
      responsibility: newTodo.responsibility.split(',').map(r => r.trim()),
      status: newTodo.status,
      priority: newTodo.priority,
      category: newTodo.category,
      dueDate: newTodo.dueDate ? newTodo.dueDate.toISOString().split('T')[0] : null,
      createdDate: newTodo.createdAt.toISOString().split('T')[0],
      updates: [],
      meetingDates: []
    };
    
    res.json(responseData);
    
  } catch (error) {
    console.error('❌ Error creating todo in database:', error);
    
    // Handle specific database errors
    if (error.name === 'SequelizeValidationError') {
      return res.status(400).json({
        error: 'Validation error',
        message: error.errors.map(err => err.message).join(', '),
        details: error.errors
      });
    }
    
    res.status(500).json({
      error: 'Failed to create todo in database',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});


// Add this new endpoint to check data persistence status:
app.get('/api/todos/persistence-status', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`Checking persistence status for ${userMap[userEmail] || userEmail}`);
    
    // Check if todos file exists
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    
    // Check backup info
    const backupInfoPath = path.join(__dirname, 'data', '.backup_info.json');
    let backupInfo = null;
    
    if (fs.existsSync(backupInfoPath)) {
      try {
        const backupData = fs.readFileSync(backupInfoPath, 'utf8');
        backupInfo = JSON.parse(backupData);
      } catch (e) {
        console.warn('Could not read backup info:', e.message);
      }
    }
    
    // Get server uptime and memory
    const uptimeHours = Math.floor(process.uptime() / 3600);
    const memoryUsage = process.memoryUsage();
    
    res.json({
      dataFile: {
        exists: fileInfo.exists,
        path: fileInfo.exists ? fileInfo.path : null,
        size: fileInfo.exists ? fileInfo.size : 0,
        lastModified: fileInfo.exists ? fileInfo.lastModified : null
      },
      backup: backupInfo,
      server: {
        uptime: `${uptimeHours} hours`,
        memory: {
          used: Math.round(memoryUsage.heapUsed / 1024 / 1024) + ' MB',
          total: Math.round(memoryUsage.heapTotal / 1024 / 1024) + ' MB'
        },
        platform: process.platform,
        nodeVersion: process.version
      },
      sync: {
        toOneDrive: 'Manual via GitHub Actions',
        frequency: 'Every 2 hours during weekdays',
        lastSync: 'Check GitHub Actions for last run'
      }
    });
    
  } catch (error) {
    console.error('Error checking persistence status:', error);
    res.status(500).json({
      error: 'Failed to check persistence status',
      details: error.message
    });
  }
});

// FIX 5: Add a manual sync trigger endpoint

// Add this endpoint to manually trigger sync:
app.post('/api/todos/trigger-sync', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`Manual sync triggered by ${userMap[userEmail] || userEmail}`);
    
    // Trigger backup
    backupTodosToGitHub()
      .then(success => {
        res.json({
          success: success,
          message: success ? 'Backup completed successfully' : 'Backup failed',
          timestamp: new Date().toISOString(),
          triggeredBy: userEmail,
          note: 'GitHub Actions will sync to OneDrive within 2 hours during weekdays'
        });
      })
      .catch(error => {
        res.status(500).json({
          success: false,
          error: 'Backup failed',
          details: error.message,
          timestamp: new Date().toISOString(),
          triggeredBy: userEmail
        });
      });
    
  } catch (error) {
    console.error('Error triggering manual sync:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to trigger sync',
      details: error.message
    });
  }
});

// Also, let's fix the PUT endpoint to actually update the Excel file
app.put('/api/todos/:id', authenticateMicrosoftToken, async (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const todoId = parseInt(req.params.id);
    const updateData = req.body;
    
    console.log(`📝 Updating todo ${todoId} in database for ${userMap[userEmail] || userEmail}:`, updateData);
    
    // Validation
    if (!updateData.status || !updateData.note) {
      return res.status(400).json({
        error: 'Missing required fields',
        message: 'Status and note are required for updates'
      });
    }
    
    // Check if todo exists
    const existingTodo = await Todo.findByPk(todoId);
    if (!existingTodo) {
      return res.status(404).json({
        error: 'Todo not found',
        message: `Todo with ID ${todoId} does not exist`
      });
    }
    
    // Update the todo status
    await Todo.update(
      { status: updateData.status },
      { 
        where: { id: todoId },
        returning: true // For PostgreSQL
      }
    );
    
    // Create update record
    const todoUpdate = await TodoUpdate.create({
      todoId: todoId,
      status: updateData.status,
      note: updateData.note,
      meetingDate: updateData.meetingDate ? new Date(updateData.meetingDate) : new Date(),
      updatedBy: userEmail
    });
    
    console.log(`✅ Todo ${todoId} updated in database, update record ID: ${todoUpdate.id}`);
    
    // Create Excel backup asynchronously
    setTimeout(async () => {
      try {
        await createExcelBackupFromDatabase();
        console.log('✅ Excel backup updated after todo update');
      } catch (error) {
        console.error('❌ Excel backup failed:', error);
      }
    }, 1000);
    
    res.json({
      success: true,
      message: 'Todo updated successfully',
      update: {
        todoId: todoId,
        date: updateData.meetingDate || new Date().toISOString().split('T')[0],
        status: updateData.status,
        note: updateData.note,
        updatedBy: userEmail
      }
    });
    
  } catch (error) {
    console.error('❌ Error updating todo in database:', error);
    res.status(500).json({
      error: 'Failed to update todo in database',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// API endpoint to edit todo fields (separate from status updates)
app.put('/api/todos/:id/edit', authenticateMicrosoftToken, async (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const todoId = parseInt(req.params.id);
    const editData = req.body;
    
    console.log(`📝 Editing todo ${todoId} fields for ${userMap[userEmail] || userEmail}:`, editData);
    
    // Validation
    if (!editData.issue || !editData.issue.trim()) {
      return res.status(400).json({
        error: 'Missing required fields',
        message: 'Issue description is required'
      });
    }
    
    if (!editData.responsibility || !Array.isArray(editData.responsibility) || editData.responsibility.length === 0) {
      return res.status(400).json({
        error: 'Missing required fields',
        message: 'At least one team member must be assigned'
      });
    }
    
    // Check if todo exists
    const existingTodo = await Todo.findByPk(todoId);
    if (!existingTodo) {
      return res.status(404).json({
        error: 'Todo not found',
        message: `Todo with ID ${todoId} does not exist`
      });
    }
    
    // Update the todo fields
    await Todo.update(
      { 
        issue: editData.issue.trim(),
        responsibility: editData.responsibility.join(','),
        priority: editData.priority,
        category: editData.category,
        dueDate: editData.dueDate ? new Date(editData.dueDate) : null
      },
      { 
        where: { id: todoId },
        returning: true
      }
    );
    
    console.log(`✅ Todo ${todoId} fields updated successfully`);
    
    // Create Excel backup asynchronously
    setTimeout(async () => {
      try {
        await createExcelBackupFromDatabase();
        console.log('✅ Excel backup updated after todo edit');
      } catch (error) {
        console.error('❌ Excel backup failed:', error);
      }
    }, 1000);
    
    res.json({
      success: true,
      message: 'Todo updated successfully',
      todoId: todoId,
      updatedBy: userEmail
    });
    
  } catch (error) {
    console.error('❌ Error editing todo:', error);
    res.status(500).json({
      error: 'Failed to edit todo',
      details: error.message,
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
});

// API endpoint to delete todo - WITH AUTHENTICATION
app.delete('/api/todos/:id', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const todoId = parseInt(req.params.id);
    
    console.log(`API request to delete todo ${todoId} from ${userMap[userEmail] || userEmail}`);
    
    // In a real implementation, you would:
    // 1. Read the existing Excel file
    // 2. Remove the todo from the Todos sheet
    // 3. Remove related updates from the Updates sheet
    // 4. Save the Excel file
    // 5. Trigger a sync to OneDrive
    
    // Mock response
    const deleteResponse = {
      success: true,
      message: 'Todo deleted successfully',
      deletedId: todoId,
      deletedBy: userEmail
    };
    
    console.log(`Mock todo deletion:`, deleteResponse);
    res.json(deleteResponse);
    
  } catch (error) {
    console.error('Error deleting todo:', error);
    res.status(500).json({ 
      error: 'Failed to delete todo', 
      details: error.message
    });
  }
});

// API endpoint for meeting summary - WITH AUTHENTICATION
app.get('/api/meetings', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/meetings from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'R&D Todos Excel file not found',
        message: 'The R&D Todos file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading R&D Todos Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read R&D Todos Excel file',
        details: readError.message
      });
    }
    
    // Read Meetings sheet (optional)
    const meetingsSheetName = 'Meetings';
    const meetingsSheet = workbook.Sheets[meetingsSheetName];
    
    let meetingsData = [];
    if (meetingsSheet) {
      const meetingsRawData = xlsx.utils.sheet_to_json(meetingsSheet);
      
      meetingsData = meetingsRawData.map(meeting => ({
        date: parseExcelDate(meeting['MEETING_DATE']),
        attendees: meeting['ATTENDEES'] ? meeting['ATTENDEES'].split(',').map(a => a.trim()) : [],
        topicsDiscussed: meeting['TOPICS_DISCUSSED'] || 0,
        notes: meeting['NOTES'] || ''
      })).filter(meeting => meeting.date).sort((a, b) => new Date(b.date) - new Date(a.date));
    } else {
      console.warn(`${meetingsSheetName} sheet not found. Using empty meetings data.`);
    }
    
    // Also get meeting dates from Updates sheet to show meeting activity
    const updatesSheet = workbook.Sheets['Updates'];
    let meetingDatesFromUpdates = [];
    
    if (updatesSheet) {
      const updatesRawData = xlsx.utils.sheet_to_json(updatesSheet);
      meetingDatesFromUpdates = [...new Set(
        updatesRawData
          .map(update => parseExcelDate(update['MEETING_DATE']))
          .filter(Boolean)
          .map(date => date.toISOString().split('T')[0])
      )].sort((a, b) => new Date(b) - new Date(a));
    }
    
    console.log(`Returning ${meetingsData.length} meetings and ${meetingDatesFromUpdates.length} meeting dates from updates`);
    
    res.json({
      meetings: meetingsData.map(meeting => ({
        ...meeting,
        date: meeting.date ? meeting.date.toISOString().split('T')[0] : null
      })),
      meetingDatesFromUpdates: meetingDatesFromUpdates,
      totalMeetings: meetingsData.length + meetingDatesFromUpdates.length
    });
    
  } catch (error) {
    console.error('Error processing meetings request:', error);
    res.status(500).json({ 
      error: 'Failed to process meetings request', 
      details: error.message
    });
  }
});

app.post('/api/meetings', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const meetingData = req.body;
    
    console.log(`API request to save meeting from ${userMap[userEmail] || userEmail}:`, meetingData);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'R&D Todos Excel file not found'
      });
    }
    
    // Read the existing Excel file
    const excelFilePath = fileInfo.path;
    let workbook;
    
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file for meeting:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Get or create Meetings sheet
    let meetingsSheet = workbook.Sheets['Meetings'];
    let meetings = [];
    
    if (meetingsSheet) {
      meetings = xlsx.utils.sheet_to_json(meetingsSheet);
    } else {
      console.log('Meetings sheet not found, creating new one...');
      meetingsSheet = xlsx.utils.json_to_sheet([]);
      workbook.Sheets['Meetings'] = meetingsSheet;
    }
    
    // Add new meeting record
    const newMeeting = {
      'MEETING_DATE': new Date(meetingData.meetingDate),
      'ATTENDEES': Array.isArray(meetingData.attendees) ? meetingData.attendees.join(',') : meetingData.attendees,
      'TOPICS_DISCUSSED': meetingData.topicsDiscussed || 0,
      'NOTES': meetingData.notes || `Meeting held with ${meetingData.updates?.length || 0} todo updates`,
      'CREATED_BY': userEmail,
      'CREATED_AT': new Date()
    };
    
    meetings.push(newMeeting);
    
    // Update the Meetings sheet
    const newMeetingsSheet = xlsx.utils.json_to_sheet(meetings);
    workbook.Sheets['Meetings'] = newMeetingsSheet;
    
    // Write back to file
    try {
      xlsx.writeFile(workbook, excelFilePath);
      console.log('✅ Successfully saved meeting to Excel file');
    } catch (writeError) {
      console.error('Error writing meeting to Excel file:', writeError);
      return res.status(500).json({
        error: 'Failed to save meeting to Excel file',
        details: writeError.message
      });
    }
    
    res.json({
      success: true,
      message: 'Meeting saved successfully',
      meeting: newMeeting
    });
    
  } catch (error) {
    console.error('Error saving meeting:', error);
    res.status(500).json({ 
      error: 'Failed to save meeting', 
      details: error.message
    });
  }
});

// API endpoint to get todo statistics - WITH AUTHENTICATION
app.get('/api/todos/stats', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/todos/stats from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'R&D Todos Excel file not found',
        message: 'The R&D Todos file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading R&D Todos Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read R&D Todos Excel file',
        details: readError.message
      });
    }
    
    // Read Todos sheet
    const todosSheet = workbook.Sheets['Todos'];
    if (!todosSheet) {
      return res.status(404).json({ 
        error: 'Todos sheet not found in Excel file'
      });
    }
    
    const todosRawData = xlsx.utils.sheet_to_json(todosSheet);
    
    // Calculate statistics
    const stats = {
      total: todosRawData.length,
      byStatus: {},
      byPriority: {},
      byCategory: {},
      byResponsibility: {},
      overdue: 0,
      dueThisWeek: 0,
      dueToday: 0,
      completedThisMonth: 0,
      avgCompletionTime: 0
    };
    
    const today = new Date();
    const oneWeekFromNow = new Date(today.getTime() + (7 * 24 * 60 * 60 * 1000));
    const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    
    let completedTodos = [];
    
    todosRawData.forEach(todo => {
      const status = todo['STATUS'] || 'Pending';
      const priority = todo['PRIORITY'] || 'Medium';
      const category = todo['CATEGORY'] || 'General';
      const responsibility = todo['RESPONSIBILITY'] || '';
      const dueDate = parseExcelDate(todo['DUE_DATE']);
      const createdDate = parseExcelDate(todo['CREATED_DATE']);
      
      // Count by status
      stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
      
      // Count by priority
      stats.byPriority[priority] = (stats.byPriority[priority] || 0) + 1;
      
      // Count by category
      stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;
      
      // Count by responsibility
      if (responsibility) {
        responsibility.split(',').forEach(person => {
          const trimmedPerson = person.trim();
          if (trimmedPerson) {
            stats.byResponsibility[trimmedPerson] = (stats.byResponsibility[trimmedPerson] || 0) + 1;
          }
        });
      }
      
      // Check due dates
      if (dueDate && status !== 'Done') {
        const dueDateObj = new Date(dueDate);
        
        if (dueDateObj < today) {
          stats.overdue++;
        } else if (dueDateObj <= oneWeekFromNow) {
          stats.dueThisWeek++;
        }
        
        if (dueDateObj.toDateString() === today.toDateString()) {
          stats.dueToday++;
        }
      }
      
      // Check completed this month
      if (status === 'Done' && createdDate) {
        const createdDateObj = new Date(createdDate);
        if (createdDateObj >= startOfMonth) {
          stats.completedThisMonth++;
        }
        
        // For average completion time calculation
        if (dueDate) {
          completedTodos.push({
            created: createdDateObj,
            due: new Date(dueDate)
          });
        }
      }
    });
    
    // Calculate average completion time (for completed todos)
    if (completedTodos.length > 0) {
      const totalDays = completedTodos.reduce((sum, todo) => {
        const days = Math.abs((todo.due - todo.created) / (1000 * 60 * 60 * 24));
        return sum + days;
      }, 0);
      stats.avgCompletionTime = Math.round(totalDays / completedTodos.length);
    }
    
    console.log(`Returning statistics for ${stats.total} todos`);
    res.json(stats);
    
  } catch (error) {
    console.error('Error processing todos stats request:', error);
    res.status(500).json({ 
      error: 'Failed to process todos stats request', 
      details: error.message
    });
  }
});

// Update your server startup section to include the new todos endpoints
// Add this to your console.log statements when the server starts:
console.log(`📋 R&D Todos API available at http://localhost:${PORT}/api/todos`);
console.log(`📊 Todos Statistics API available at http://localhost:${PORT}/api/todos/stats`);
console.log(`📅 Meetings API available at http://localhost:${PORT}/api/meetings`);

// Also update your health check endpoint to include todos file info
// In your existing health check endpoint, add this to the excelFiles object:
// rnbTodos: checkExcelFile('RND_Todos.xlsx'),

// Add this helper function for Excel writing (for future use)
function updateTodosExcel(todosData, updatesData, meetingsData) {
  try {
    const filePath = path.join(__dirname, 'data', 'RND_Todos.xlsx');
    
    // Create a new workbook
    const workbook = xlsx.utils.book_new();
    
    // Create Todos sheet
    const todosSheet = xlsx.utils.json_to_sheet(todosData);
    xlsx.utils.book_append_sheet(workbook, todosSheet, 'Todos');
    
    // Create Updates sheet
    if (updatesData && updatesData.length > 0) {
      const updatesSheet = xlsx.utils.json_to_sheet(updatesData);
      xlsx.utils.book_append_sheet(workbook, updatesSheet, 'Updates');
    }
    
    // Create Meetings sheet
    if (meetingsData && meetingsData.length > 0) {
      const meetingsSheet = xlsx.utils.json_to_sheet(meetingsData);
      xlsx.utils.book_append_sheet(workbook, meetingsSheet, 'Meetings');
    }
    
    // Write the file
    xlsx.writeFile(workbook, filePath);
    
    console.log(`Updated RND_Todos.xlsx with ${todosData.length} todos, ${updatesData?.length || 0} updates, ${meetingsData?.length || 0} meetings`);
    return true;
    
  } catch (error) {
    console.error('Error updating todos Excel file:', error);
    return false;
  }
}

// API endpoint to export todos data - WITH AUTHENTICATION
app.get('/api/todos/export', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/todos/export from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('RND_Todos.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'R&D Todos Excel file not found'
      });
    }
    
    // Send the file for download
    res.download(fileInfo.path, 'RND_Todos_Export.xlsx', (err) => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).json({ error: 'Failed to export file' });
      } else {
        console.log('File exported successfully');
      }
    });
    
  } catch (error) {
    console.error('Error exporting todos file:', error);
    res.status(500).json({ 
      error: 'Failed to export todos file', 
      details: error.message
    });
  }
});

// Add these sections to your existing server.js file

// API endpoint for daily updates data - WITH AUTHENTICATION
app.get('/api/daily-updates', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/daily-updates from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
      user: userEmail,
      query: req.query,
      headers: {
        'user-agent': req.headers['user-agent'],
        'cache-control': req.headers['cache-control']
      }
    };
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Daily_Updates.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Daily Updates Excel file not found',
        message: 'The Daily Updates file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        fileInfo,
        requestInfo
      });
    }
    
    // Read the Excel file with force reload
    const excelFilePath = fileInfo.path;
    
    // Use try/catch specifically for file reading
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Daily Updates Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Daily Updates Excel file',
        details: readError.message,
        fileInfo,
        requestInfo
      });
    }
    
    // Log available sheets
    console.log('Available sheets in Daily Updates workbook:', workbook.SheetNames);
    
    // Read the "Daily Updates" sheet
    const dailyUpdatesSheetName = 'Daily Updates';
    const worksheet = workbook.Sheets[dailyUpdatesSheetName];
    if (!worksheet) {
      console.error(`Sheet "${dailyUpdatesSheetName}" not found. Available sheets:`, workbook.SheetNames);
      return res.status(404).json({ 
        error: `${dailyUpdatesSheetName} sheet not found in Excel file`,
        availableSheets: workbook.SheetNames,
        fileInfo,
        requestInfo
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${dailyUpdatesSheetName} sheet`);
    
    // Process the data to match the frontend's expected format
    const processedData = rawData.map((row, index) => {
      const date = parseExcelDate(row['DATE']);
      
      return {
        id: row['ID'] || index + 1,
        date: date ? date.toISOString().split('T')[0] : null,
        section: row['SECTION'] || '',
        update: row['UPDATE'] || '',
        priority: row['PRIORITY'] || 'Medium',
        status: row['STATUS'] || 'Pending',
        assignedTo: row['ASSIGNED_TO'] || '',
        createdBy: row['CREATED_BY'] || '',
        createdDate: parseExcelDate(row['CREATED_DATE']) || date
      };
    }).filter(item => item.section && item.update); // Filter out empty rows
    
    // Sort by date (newest first) and then by ID
    const sortedData = processedData.sort((a, b) => {
      const dateComparison = new Date(b.date) - new Date(a.date);
      if (dateComparison !== 0) return dateComparison;
      return b.id - a.id;
    });
    
    console.log(`Returning ${sortedData.length} processed daily updates`);
    res.json(sortedData);
    
  } catch (error) {
    console.error('Error processing daily updates request:', error);
    res.status(500).json({ 
      error: 'Failed to process daily updates request', 
      details: error.message,
      stack: error.stack 
    });
  }
});

// API endpoint to add new daily update - WITH AUTHENTICATION
app.post('/api/daily-updates', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const updateData = req.body;
    
    console.log(`API request to add daily update from ${userMap[userEmail] || userEmail}:`, updateData);
    
    // Validation
    if (!updateData.section || !updateData.update) {
      return res.status(400).json({
        error: 'Missing required fields',
        message: 'Section and update description are required'
      });
    }
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Daily_Updates.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Daily Updates Excel file not found',
        message: 'Cannot add update - Excel file is not available'
      });
    }
    
    // Read the existing Excel file
    const excelFilePath = fileInfo.path;
    let workbook;
    
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file for daily update:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Get the Daily Updates sheet
    const dailyUpdatesSheet = workbook.Sheets['Daily Updates'];
    if (!dailyUpdatesSheet) {
      return res.status(404).json({
        error: 'Daily Updates sheet not found in Excel file'
      });
    }
    
    // Read existing data
    let existingData = [];
    try {
      existingData = xlsx.utils.sheet_to_json(dailyUpdatesSheet);
    } catch (e) {
      console.warn('Could not read existing data, starting with empty array');
    }
    
    // Find the next ID
    const maxId = existingData.length > 0 ? 
      Math.max(...existingData.map(row => parseInt(row['ID']) || 0)) : 0;
    const nextId = maxId + 1;
    
    // Create new update record
    const newUpdate = {
      'ID': nextId,
      'DATE': new Date(updateData.date || new Date().toISOString().split('T')[0]),
      'SECTION': updateData.section,
      'UPDATE': updateData.update,
      'PRIORITY': updateData.priority || 'Medium',
      'STATUS': updateData.status || 'In Progress',
      'ASSIGNED_TO': updateData.assignedTo || '',
      'CREATED_BY': userEmail,
      'CREATED_DATE': new Date()
    };
    
    // Add new update to existing data
    existingData.push(newUpdate);
    
    // Update the sheet
    const newSheet = xlsx.utils.json_to_sheet(existingData);
    workbook.Sheets['Daily Updates'] = newSheet;
    
    // Write back to file
    try {
      xlsx.writeFile(workbook, excelFilePath);
      console.log('Successfully added daily update to Excel file');
    } catch (writeError) {
      console.error('Error writing daily update to Excel file:', writeError);
      return res.status(500).json({
        error: 'Failed to save daily update to Excel file',
        details: writeError.message
      });
    }
    
    res.json({
      success: true,
      message: 'Daily update added successfully',
      update: newUpdate
    });
    
  } catch (error) {
    console.error('Error adding daily update:', error);
    res.status(500).json({ 
      error: 'Failed to add daily update', 
      details: error.message
    });
  }
});

// API endpoint for daily updates statistics - WITH AUTHENTICATION
app.get('/api/daily-updates/stats', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/daily-updates/stats from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Daily_Updates.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Daily Updates Excel file not found',
        message: 'The Daily Updates file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Daily Updates Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Daily Updates Excel file',
        details: readError.message
      });
    }
    
    // Read Daily Updates sheet
    const dailyUpdatesSheet = workbook.Sheets['Daily Updates'];
    if (!dailyUpdatesSheet) {
      return res.status(404).json({ 
        error: 'Daily Updates sheet not found in Excel file'
      });
    }
    
    const rawData = xlsx.utils.sheet_to_json(dailyUpdatesSheet);
    
    // Calculate statistics
    const stats = {
      total: rawData.length,
      bySection: {},
      byPriority: {},
      byStatus: {},
      todayUpdates: 0,
      thisWeekUpdates: 0,
      thisMonthUpdates: 0,
      completedThisWeek: 0
    };
    
    const today = new Date();
    const todayStr = today.toISOString().split('T')[0];
    const oneWeekAgo = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
    const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    
    rawData.forEach(update => {
      const section = update['SECTION'] || 'Other';
      const priority = update['PRIORITY'] || 'Medium';
      const status = update['STATUS'] || 'Pending';
      const updateDate = parseExcelDate(update['DATE']);
      
      // Count by section
      stats.bySection[section] = (stats.bySection[section] || 0) + 1;
      
      // Count by priority
      stats.byPriority[priority] = (stats.byPriority[priority] || 0) + 1;
      
      // Count by status
      stats.byStatus[status] = (stats.byStatus[status] || 0) + 1;
      
      // Time-based counts
      if (updateDate) {
        const updateDateStr = updateDate.toISOString().split('T')[0];
        
        if (updateDateStr === todayStr) {
          stats.todayUpdates++;
        }
        
        if (updateDate >= oneWeekAgo) {
          stats.thisWeekUpdates++;
          
          if (status === 'Completed') {
            stats.completedThisWeek++;
          }
        }
        
        if (updateDate >= startOfMonth) {
          stats.thisMonthUpdates++;
        }
      }
    });
    
    console.log(`Returning statistics for ${stats.total} daily updates`);
    res.json(stats);
    
  } catch (error) {
    console.error('Error processing daily updates stats request:', error);
    res.status(500).json({ 
      error: 'Failed to process daily updates stats request', 
      details: error.message
    });
  }
});
// Separate API endpoint for detailed shrinkage test analysis - WITH AUTHENTICATION
app.get('/api/shrinkage-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/shrinkage-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "Shrinkage" for shrinkage data
    const shrinkageSheetName = 'Shrinkage';
    const worksheet = workbook.Sheets[shrinkageSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${shrinkageSheetName} sheet not found in Excel file`,
        message: `Please add a "${shrinkageSheetName}" sheet to your Excel file with shrinkage test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${shrinkageSheetName} sheet`);
    
    // Filter out empty rows and header rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && 
      row['ENCAPSULANT TYPE'] && 
      (row['ENCAPSULANT TYPE'] === 'FRONT EPE' || row['ENCAPSULANT TYPE'] === 'BACK EVA')
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          passRate: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = calculateShrinkageResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.finalStatus === 'PASS').length;
    const failedTests = totalTests - passedTests;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Group by vendor for additional insights
    const vendorSummary = {};
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0,
          frontEpe: { total: 0, passed: 0 },
          backEva: { total: 0, passed: 0 }
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      
      if (item.finalStatus === 'PASS') {
        vendor.passed++;
      } else {
        vendor.failed++;
      }
      
      // Track by encapsulant type
      if (item.encapsulantType === 'FRONT EPE') {
        vendor.frontEpe.total++;
        if (item.finalStatus === 'PASS') vendor.frontEpe.passed++;
      } else if (item.encapsulantType === 'BACK EVA') {
        vendor.backEva.total++;
        if (item.finalStatus === 'PASS') vendor.backEva.passed++;
      }
    });
    
    console.log(`Returning ${processedData.length} shrinkage test records with ${passedTests} passed and ${failedTests} failed`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        passRate
      },
      vendorSummary
    });
    
  } catch (error) {
    console.error('Error processing shrinkage tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process shrinkage tests request', 
      details: error.message
    });
  }
});

// Add these API endpoints to your server.js file

// API endpoint for detailed adhesion test analysis - WITH AUTHENTICATION
app.get('/api/adhesion-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/adhesion-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "Adhesion" sheet
    const adhesionSheetName = 'Adhesion';
    const worksheet = workbook.Sheets[adhesionSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${adhesionSheetName} sheet not found in Excel file`,
        message: `Please add an "${adhesionSheetName}" sheet to your Excel file with adhesion test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${adhesionSheetName} sheet`);
    
    // Filter out empty rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          passRate: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = calculateAdhesionResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.finalStatus === 'PASS').length;
    const failedTests = totalTests - passedTests;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Group by vendor for additional insights
    const vendorSummary = {};
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      
      if (item.finalStatus === 'PASS') {
        vendor.passed++;
      } else {
        vendor.failed++;
      }
    });
    
    console.log(`Returning ${processedData.length} adhesion test records with ${passedTests} passed and ${failedTests} failed`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        passRate
      },
      vendorSummary
    });
    
  } catch (error) {
    console.error('Error processing adhesion tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process adhesion tests request', 
      details: error.message
    });
  }
});

// API endpoint for detailed tensile strength test analysis - WITH AUTHENTICATION
app.get('/api/tensile-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/tensile-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "Tensile Strength" sheet
    // Read the "Tensile Strength" sheet
    const tensileSheetName = 'Tensile Strength';
    const worksheet = workbook.Sheets[tensileSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${tensileSheetName} sheet not found in Excel file`,
        message: `Please add a "${tensileSheetName}" sheet to your Excel file with tensile test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${tensileSheetName} sheet`);
    
    // Filter out empty rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          passRate: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = calculateTensileStrengthResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.finalStatus === 'PASS').length;
    const failedTests = totalTests - passedTests;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Group by vendor for additional insights
    const vendorSummary = {};
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0,
          avgBreakValue: 0,
          avgElongation: 0
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      vendor.avgBreakValue += item.breakValue;
      vendor.avgElongation += item.changeInElongationPercent;
      
      if (item.finalStatus === 'PASS') {
        vendor.passed++;
      } else {
        vendor.failed++;
      }
    });
    
    // Calculate averages
    Object.keys(vendorSummary).forEach(vendor => {
      vendorSummary[vendor].avgBreakValue = (vendorSummary[vendor].avgBreakValue / vendorSummary[vendor].total).toFixed(2);
      vendorSummary[vendor].avgElongation = (vendorSummary[vendor].avgElongation / vendorSummary[vendor].total).toFixed(2);
    });
    
    console.log(`Returning ${processedData.length} tensile test records with ${passedTests} passed and ${failedTests} failed`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        passRate
      },
      vendorSummary
    });
    
  } catch (error) {
    console.error('Error processing tensile tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process tensile tests request', 
      details: error.message
    });
  }
});

// API endpoint for detailed GSM test analysis - WITH AUTHENTICATION
app.get('/api/gsm-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/gsm-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "GSM" sheet
    const gsmSheetName = 'GSM';
    const worksheet = workbook.Sheets[gsmSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${gsmSheetName} sheet not found in Excel file`,
        message: `Please add a "${gsmSheetName}" sheet to your Excel file with GSM test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${gsmSheetName} sheet`);
    
    // Filter out empty rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          passRate: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = calculateGSMResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.finalStatus === 'PASS').length;
    const failedTests = totalTests - passedTests;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Group by vendor and category
    const vendorSummary = {};
    const categorySummary = { OLD: { total: 0, passed: 0 }, NEW: { total: 0, passed: 0 } };
    
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0,
          avgValue: 0
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      vendor.avgValue += item.average;
      
      if (item.finalStatus === 'PASS') {
        vendor.passed++;
      } else {
        vendor.failed++;
      }
      
      // Category summary
      const category = item.category.toUpperCase().includes('OLD') ? 'OLD' : 'NEW';
      categorySummary[category].total++;
      if (item.finalStatus === 'PASS') {
        categorySummary[category].passed++;
      }
    });
    
    // Calculate averages
    Object.keys(vendorSummary).forEach(vendor => {
      vendorSummary[vendor].avgValue = (vendorSummary[vendor].avgValue / vendorSummary[vendor].total).toFixed(2);
    });
    
    console.log(`Returning ${processedData.length} GSM test records with ${passedTests} passed and ${failedTests} failed`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        passRate
      },
      vendorSummary,
      categorySummary
    });
    
  } catch (error) {
    console.error('Error processing GSM tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process GSM tests request', 
      details: error.message
    });
  }
});

// API endpoint for detailed resistance test analysis - WITH AUTHENTICATION
app.get('/api/resistance-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/resistance-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "Resistance" sheet
    const resistanceSheetName = 'Resistance';
    const worksheet = workbook.Sheets[resistanceSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${resistanceSheetName} sheet not found in Excel file`,
        message: `Please add a "${resistanceSheetName}" sheet to your Excel file with resistance test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${resistanceSheetName} sheet`);
    
    // Filter out empty rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && row['BOM'] && row['TYPE']
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          pendingTests: 0,
          passRate: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = processResistanceResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.testResult === 'PASS').length;
    const failedTests = processedData.filter(item => item.testResult === 'FAIL').length;
    const pendingTests = processedData.filter(item => item.testResult === 'Pending').length;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Group by vendor and ribbon type
    const vendorSummary = {};
    const ribbonTypeSummary = { 
      'BUS RIBBON': { total: 0, passed: 0, failed: 0, pending: 0 }, 
      'INTERCONNECT RIBBON': { total: 0, passed: 0, failed: 0, pending: 0 } 
    };
    
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0,
          pending: 0,
          busRibbon: { total: 0, passed: 0 },
          interconnectRibbon: { total: 0, passed: 0 }
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      
      if (item.testResult === 'PASS') {
        vendor.passed++;
      } else if (item.testResult === 'FAIL') {
        vendor.failed++;
      } else {
        vendor.pending++;
      }
      
      // Ribbon type summary
      if (item.type.includes('BUS')) {
        vendor.busRibbon.total++;
        ribbonTypeSummary['BUS RIBBON'].total++;
        if (item.testResult === 'PASS') {
          vendor.busRibbon.passed++;
          ribbonTypeSummary['BUS RIBBON'].passed++;
        } else if (item.testResult === 'FAIL') {
          ribbonTypeSummary['BUS RIBBON'].failed++;
        } else {
          ribbonTypeSummary['BUS RIBBON'].pending++;
        }
      } else if (item.type.includes('INTERCONNECT')) {
        vendor.interconnectRibbon.total++;
        ribbonTypeSummary['INTERCONNECT RIBBON'].total++;
        if (item.testResult === 'PASS') {
          vendor.interconnectRibbon.passed++;
          ribbonTypeSummary['INTERCONNECT RIBBON'].passed++;
        } else if (item.testResult === 'FAIL') {
          ribbonTypeSummary['INTERCONNECT RIBBON'].failed++;
        } else {
          ribbonTypeSummary['INTERCONNECT RIBBON'].pending++;
        }
      }
    });
    
    console.log(`Returning ${processedData.length} resistance test records with ${passedTests} passed, ${failedTests} failed, and ${pendingTests} pending`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        pendingTests,
        passRate
      },
      vendorSummary,
      ribbonTypeSummary
    });
    
  } catch (error) {
    console.error('Error processing resistance tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process resistance tests request', 
      details: error.message
    });
  }
});

// API endpoint for detailed bypass diode test analysis - WITH AUTHENTICATION
app.get('/api/bypass-tests', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/bypass-tests from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive.'
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        dateNF: 'yyyy-mm-dd',
        cellNF: true,
        cellStyles: true,
        type: 'binary',
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message
      });
    }
    
    // Read the "BYPASS DIODE TEST" sheet
    const bypassSheetName = 'BYPASS DIODE TEST';
    const worksheet = workbook.Sheets[bypassSheetName];
    if (!worksheet) {
      return res.status(404).json({ 
        error: `${bypassSheetName} sheet not found in Excel file`,
        message: `Please add a "${bypassSheetName}" sheet to your Excel file with bypass diode test data.`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${bypassSheetName} sheet`);
    
    // Filter out empty rows
    const validData = rawData.filter(row => 
      row['VENDOR NAME'] && row['BOM']
    );
    
    if (validData.length === 0) {
      return res.json({
        data: [],
        summary: {
          totalTests: 0,
          passedTests: 0,
          failedTests: 0,
          pendingTests: 0,
          passRate: 0,
          avgTemperature: 0
        }
      });
    }
    
    // Process the data and calculate results
    const processedData = processBypassDiodeResults(validData);
    
    // Calculate summary statistics
    const totalTests = processedData.length;
    const passedTests = processedData.filter(item => item.testResult === 'PASS').length;
    const failedTests = processedData.filter(item => item.testResult === 'FAIL').length;
    const pendingTests = processedData.filter(item => item.testResult === 'Pending').length;
    const passRate = totalTests > 0 ? Math.round((passedTests / totalTests) * 100) : 0;
    
    // Calculate average temperature
    const validTemperatures = processedData.filter(item => item.hasValidTemperature);
    const avgTemperature = validTemperatures.length > 0 ? 
      (validTemperatures.reduce((sum, item) => sum + item.maxTemperatureTj, 0) / validTemperatures.length).toFixed(2) : 0;
    
    // Group by vendor
    const vendorSummary = {};
    processedData.forEach(item => {
      if (!vendorSummary[item.vendorName]) {
        vendorSummary[item.vendorName] = {
          total: 0,
          passed: 0,
          failed: 0,
          pending: 0,
          avgTemperature: 0,
          tempCount: 0
        };
      }
      
      const vendor = vendorSummary[item.vendorName];
      vendor.total++;
      
      if (item.testResult === 'PASS') {
        vendor.passed++;
      } else if (item.testResult === 'FAIL') {
        vendor.failed++;
      } else {
        vendor.pending++;
      }
      
      if (item.hasValidTemperature) {
        vendor.avgTemperature += item.maxTemperatureTj;
        vendor.tempCount++;
      }
    });
    
    // Calculate vendor average temperatures
    Object.keys(vendorSummary).forEach(vendor => {
      if (vendorSummary[vendor].tempCount > 0) {
        vendorSummary[vendor].avgTemperature = (vendorSummary[vendor].avgTemperature / vendorSummary[vendor].tempCount).toFixed(2);
      }
    });
    
    console.log(`Returning ${processedData.length} bypass diode test records with ${passedTests} passed, ${failedTests} failed, and ${pendingTests} pending`);
    
    res.json({
      data: processedData,
      summary: {
        totalTests,
        passedTests,
        failedTests,
        pendingTests,
        passRate,
        avgTemperature
      },
      vendorSummary
    });
    
  } catch (error) {
    console.error('Error processing bypass diode tests request:', error);
    res.status(500).json({ 
      error: 'Failed to process bypass diode tests request', 
      details: error.message
    });
  }
});

// Add these endpoints before your existing endpoints in server.js
// Make sure to add them after all the calculation functions are defined

// Fix for the api/solar-data endpoint - NOW WITH AUTHENTICATION
app.get('/api/solar-data', authenticateMicrosoftToken, (req, res) => {
  console.log('Received request to /api/solar-data, redirecting to /api/test-data');
  // This redirects api/solar-data requests to api/test-data
  req.url = '/api/test-data';
  app.handle(req, res);
});

// Metadata endpoint with improved error handling - NOW WITH AUTHENTICATION
app.get('/api/metadata', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/metadata from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check if the Excel file exists
    const fileInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        fileInfo
      });
    }
    
    // Read the Excel file
    const excelFilePath = fileInfo.path;
    
    let workbook;
    try {
      workbook = xlsx.readFile(excelFilePath, {
        cellDates: true,
        cache: false
      });
    } catch (readError) {
      console.error('Error reading Excel file for metadata:', readError);
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message,
        fileInfo
      });
    }
    
    // Read the Material Tests Map sheet
    const materialTestsSheetName = 'Material Tests Map';
    const materialTestsSheet = workbook.Sheets[materialTestsSheetName];
    let materialTestsData = [];
    
    if (materialTestsSheet) {
      materialTestsData = xlsx.utils.sheet_to_json(materialTestsSheet);
      console.log(`Processed ${materialTestsData.length} rows from ${materialTestsSheetName} sheet`);
    } else {
      console.warn(`Sheet "${materialTestsSheetName}" not found. Using Test Data sheet for metadata.`);
      
      // If Material Tests Map not found, extract from Test Data as fallback
      const testDataSheet = workbook.Sheets['Test Data'];
      if (testDataSheet) {
        const testData = xlsx.utils.sheet_to_json(testDataSheet);
        // Create simplified mapping from actual test data
        testData.forEach(row => {
          if (row['BOM'] && row['TEST NAME']) {
            materialTestsData.push({
              'BOM': row['BOM'],
              'TEST NAME': row['TEST NAME']
            });
          }
        });
      }
    }
    
    // Process to get BOM-tests mapping
    const bomTests = [];
    const uniqueBoms = new Set();
    const uniqueTests = new Set();
    
    materialTestsData.forEach(item => {
      if (item['BOM'] && item['TEST NAME']) {
        uniqueBoms.add(item['BOM']);
        uniqueTests.add(item['TEST NAME']);
        
        // Find or create BOM entry
        let bomEntry = bomTests.find(b => b.bom === item['BOM']);
        if (!bomEntry) {
          bomEntry = { bom: item['BOM'], tests: [] };
          bomTests.push(bomEntry);
        }
        
        // Add test to BOM entry if not already added
        if (!bomEntry.tests.includes(item['TEST NAME'])) {
          bomEntry.tests.push(item['TEST NAME']);
        }
      }
    });
    
    const result = {
      bomTests: bomTests,
      uniqueBoms: Array.from(uniqueBoms),
      uniqueTests: Array.from(uniqueTests),
      metadata: {
        fileInfo: fileInfo,
        generatedAt: new Date().toISOString(),
        requestedBy: userEmail
      }
    };
    
    res.json(result);
  } catch (error) {
    console.error('Error reading Excel metadata:', error);
    res.status(500).json({ 
      error: 'Failed to read Excel metadata',
      details: error.message
    });
  }
});

// Enhanced file status endpoint - NOW WITH AUTHENTICATION
app.get('/api/data-status', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/data-status from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check all Excel files
    const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
    const certificationsInfo = checkExcelFile('Certifications.xlsx');
    const chamberTestsInfo = checkExcelFile('Chamber_Tests.xlsx');
    const rndTodosInfo = checkExcelFile('RND_Todos.xlsx');
    const dailyUpdatesInfo = checkExcelFile('Daily_Updates.xlsx'); // ADD THIS LINE
    
    if (!solarLabInfo.exists && !lineTrialsInfo.exists && !certificationsInfo.exists) {
      return res.status(404).json({
        success: false,
        message: 'No Excel files found',
        files: {
          solarLabInfo,
          lineTrialsInfo,
          certificationsInfo,
          chamberTestsInfo,
          rndTodosInfo
        },
        user: userEmail
      });
    }
    
    // Get sheet names from each file
    const fileDetails = {};
    
    // Process Solar Lab Tests file
    if (solarLabInfo.exists) {
      try {
        const workbook = xlsx.readFile(solarLabInfo.path, { 
          bookSheets: true,
          cache: false
        });
        fileDetails.solarLabTests = {
          lastUpdated: solarLabInfo.lastModified,
          fileSize: solarLabInfo.size,
          sheets: workbook.SheetNames || []
        };
      } catch (e) {
        fileDetails.solarLabTests = {
          error: `Error reading sheet names: ${e.message}`
        };
      }
    }
    
    // Process Line Trials file
    if (lineTrialsInfo.exists) {
      try {
        const workbook = xlsx.readFile(lineTrialsInfo.path, { 
          bookSheets: true,
          cache: false
        });
        fileDetails.lineTrials = {
          lastUpdated: lineTrialsInfo.lastModified,
          fileSize: lineTrialsInfo.size,
          sheets: workbook.SheetNames || []
        };
      } catch (e) {
        fileDetails.lineTrials = {
          error: `Error reading sheet names: ${e.message}`
        };
      }
    }
    
    // Process Certifications file
    if (certificationsInfo.exists) {
      try {
        const workbook = xlsx.readFile(certificationsInfo.path, { 
          bookSheets: true,
          cache: false
        });
        fileDetails.certifications = {
          lastUpdated: certificationsInfo.lastModified,
          fileSize: certificationsInfo.size,
          sheets: workbook.SheetNames || []
        };
      } catch (e) {
        fileDetails.certifications = {
          error: `Error reading sheet names: ${e.message}`
        };
      }
    }
    
    if (chamberTestsInfo.exists) {
      try {
        const workbook = xlsx.readFile(chamberTestsInfo.path, { 
          bookSheets: true,
          cache: false
        });
        fileDetails.chamberTests = {
          lastUpdated: chamberTestsInfo.lastModified,
          fileSize: chamberTestsInfo.size,
          sheets: workbook.SheetNames || []
        };
      } catch (e) {
        fileDetails.chamberTests = {
          error: `Error reading sheet names: ${e.message}`
        };
      }
    }
    
    // ADD THIS SECTION - Process RND Todos file
    if (rndTodosInfo.exists) {
      try {
        const workbook = xlsx.readFile(rndTodosInfo.path, { 
          bookSheets: true,
          cache: false
        });
        fileDetails.rndTodos = {
          lastUpdated: rndTodosInfo.lastModified,
          fileSize: rndTodosInfo.size,
          sheets: workbook.SheetNames || []
        };
      } catch (e) {
        fileDetails.rndTodos = {
          error: `Error reading sheet names: ${e.message}`
        };
      }
    }

    // ADD THIS SECTION - Process Daily Updates file
if (dailyUpdatesInfo.exists) {
  try {
    const workbook = xlsx.readFile(dailyUpdatesInfo.path, { 
      bookSheets: true,
      cache: false
    });
    fileDetails.dailyUpdates = {
      lastUpdated: dailyUpdatesInfo.lastModified,
      fileSize: dailyUpdatesInfo.size,
      sheets: workbook.SheetNames || []
    };
  } catch (e) {
    fileDetails.dailyUpdates = {
      error: `Error reading sheet names: ${e.message}`
    };
  }
}


    res.json({
      success: true,
      files: fileDetails,
      serverTime: new Date().toISOString(),
      user: userEmail
    });
  } catch (error) {
    console.error('Error checking file status:', error);
    res.status(500).json({
      success: false,
      message: 'Error checking file status',
      error: error.message
    });
  }
});

// New file info endpoint - NOW WITH AUTHENTICATION
app.get('/api/file-info', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    console.log(`API request received for /api/file-info from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    // Check all Excel files
    const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
    const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
    const certificationsInfo = checkExcelFile('Certifications.xlsx');
    
    const fileInfos = {
      solarLabTests: solarLabInfo,
      lineTrials: lineTrialsInfo,
      certifications: certificationsInfo
    };
    
    // Try to get more detailed info for each file
    for (const [key, fileInfo] of Object.entries(fileInfos)) {
      if (fileInfo.exists) {
        try {
          // Read the file as binary to get a hash
          const buffer = fs.readFileSync(fileInfo.path);
          const hash = require('crypto')
            .createHash('md5')
            .update(buffer)
            .digest('hex');
          
          fileInfo.md5 = hash;
          fileInfo.contentSample = buffer.slice(0, 100).toString('hex');
        } catch (e) {
          console.error(`Error getting file hash for ${key}:`, e);
          fileInfo.error = e.message;
        }
      }
    }
    
    res.json({
      files: fileInfos,
      serverInfo: {
        time: new Date().toISOString(),
        pid: process.pid,
        platform: process.platform,
        nodeVersion: process.version,
        memoryUsage: process.memoryUsage(),
        requestedBy: userEmail
      }
    });
  } catch (error) {
    console.error('Error getting file info:', error);
    res.status(500).json({
      success: false,
      message: 'Error getting file info',
      error: error.message
    });
  }
});

// Debug endpoint with improved error handling - NOW WITH AUTHENTICATION
app.get('/api/debug/excel/:file?', authenticateMicrosoftToken, (req, res) => {
  try {
    const userEmail = req.user.preferred_username || req.user.upn || req.user.email;
    const fileName = req.params.file || 'Solar_Lab_Tests.xlsx';
    console.log(`API request received for /api/debug/excel/${fileName} from ${userMap[userEmail] || userEmail} at ${new Date().toISOString()}`);
    
    const fileInfo = checkExcelFile(fileName);
    
    if (!fileInfo.exists) {
      return res.status(404).json({ 
        error: `Excel file ${fileName} not found`,
        fileInfo 
      });
    }
    
    let workbook;
    try {
      workbook = xlsx.readFile(fileInfo.path, { 
        cellDates: true,
        cache: false 
      });
    } catch (readError) {
      return res.status(500).json({
        error: 'Failed to read Excel file',
        details: readError.message,
        fileInfo
      });
    }
    
    // Get information about each sheet
    const sheetsInfo = {};
    workbook.SheetNames.forEach(name => {
      try {
        const sheet = workbook.Sheets[name];
        const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1:A1');
        const sampleData = xlsx.utils.sheet_to_json(sheet, { 
          header: 1, 
          range: 0, 
          defval: null 
        }).slice(0, 2);
        
        sheetsInfo[name] = {
          rowCount: range.e.r - range.s.r + 1,
          columnCount: range.e.c - range.s.c + 1,
          columnNames: sampleData[0] || [],
          sampleRow: sampleData[1] || []
        };
      } catch (sheetError) {
        sheetsInfo[name] = {
          error: `Error reading sheet: ${sheetError.message}`
        };
      }
    });
    
    res.json({
      fileName: path.basename(fileInfo.path),
      fileInfo: fileInfo,
      sheets: sheetsInfo,
      serverTime: new Date().toISOString(),
      requestedBy: userEmail
    });
    
  } catch (error) {
    res.status(500).json({ 
      error: 'Failed to read Excel structure', 
      details: error.message 
    });
  }
});

// Enhanced health check endpoint with user tracking
app.get('/health', (req, res) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];
  let userInfo = 'Anonymous';
  
  if (token) {
    try {
      const decoded = jwt.decode(token);
      const userEmail = decoded?.preferred_username || decoded?.upn || decoded?.email;
      userInfo = userMap[userEmail] || userEmail || 'Unknown User';
    } catch (e) {
      userInfo = 'Token Error';
    }
  }
  
  // Check all Excel files
  const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
  const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
  const certificationsInfo = checkExcelFile('Certifications.xlsx');
  const chamberTestsInfo = checkExcelFile('Chamber_Tests.xlsx');
  const rndTodosInfo = checkExcelFile('RND_Todos.xlsx'); // ADD THIS LINE
  
  
  
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    requestedBy: userInfo,
    excelFiles: {
      solarLabTests: solarLabInfo,
      lineTrials: lineTrialsInfo,
      certifications: certificationsInfo,
      chamberTests: chamberTestsInfo,
      rndTodos: rndTodosInfo,
      dailyUpdates: checkExcelFile('Daily_Updates.xlsx')// ADD THIS LINE
    },
    memory: process.memoryUsage(),
    environment: process.env.NODE_ENV || 'development',
    authenticationEnabled: true,
    authorizedUsers: AUTHORIZED_EMAILS.filter(email => email).length
  });
});

// Start the server with more debug info
app.listen(PORT, async () => {
  console.log(`🚀 Server running on port ${PORT} at ${new Date().toISOString()}`);
  console.log(`🔐 Microsoft Authentication ENABLED`);
  console.log(`👥 Authorized users: ${AUTHORIZED_EMAILS.filter(email => email).length} team members`);
  console.log(`🌐 API available at http://localhost:${PORT}/api/test-data`);
  console.log(`📊 Line Trials API available at http://localhost:${PORT}/api/line-trials`);
  console.log(`📋 Certifications API available at http://localhost:${PORT}/api/certifications`);
  console.log(`🏠 Chamber Tests API available at http://localhost:${PORT}/api/chamber-data`);
  console.log(`🔍 Excel debug endpoint available at http://localhost:${PORT}/api/debug/excel`);
  console.log(`📊 Test Data API: http://localhost:${PORT}/api/test-data`);
  console.log(`📋 R&D Todos API: http://localhost:${PORT}/api/todos`);
  console.log(`📅 Meetings API: http://localhost:${PORT}/api/meetings`);
  console.log(`📁 File info endpoint available at http://localhost:${PORT}/api/file-info`);
  console.log(`❤️ Health check available at http://localhost:${PORT}/health`);

  console.log('\n📊 Initializing Database...');
  const dbConnected = await initializeDatabase();

  if (dbConnected) {
    console.log('🎯 ✅ DATABASE READY - Todos will persist across deployments!');
    
    // Create initial Excel backup
    try {
      await createExcelBackupFromDatabase();
      console.log('✅ Initial Excel backup created');
    } catch (error) {
      console.warn('⚠️ Could not create initial Excel backup:', error.message);
    }
    
  } else {
    console.error('❌ DATABASE CONNECTION FAILED - Todos will be lost on restart!');
    console.error('   Please check your DATABASE_URL environment variable');
  }
  
  // Check Excel files on startup
  const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
  const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
  const certificationsInfo = checkExcelFile('Certifications.xlsx');
  
  if (solarLabInfo.exists) {
    console.log(`✅ Solar Lab Tests Excel file is ready at ${solarLabInfo.path}, size: ${solarLabInfo.size} bytes`);
  } else {
    console.log(`⏳ Waiting for Solar Lab Tests Excel file to be synced to ${solarLabInfo.path}`);
  }
  
  if (lineTrialsInfo.exists) {
    console.log(`✅ Line Trials Excel file is ready at ${lineTrialsInfo.path}, size: ${lineTrialsInfo.size} bytes`);
  } else {
    console.log(`⏳ Waiting for Line Trials Excel file to be synced to ${lineTrialsInfo.path}`);
  }
  
  if (certificationsInfo.exists) {
    console.log(`✅ Certifications Excel file is ready at ${certificationsInfo.path}, size: ${certificationsInfo.size} bytes`);
  } else {
    console.log(`⏳ Waiting for Certifications Excel file to be synced to ${certificationsInfo.path}`);
  }
  
  console.log('🔒 All API endpoints are now protected with Microsoft Authentication');
  console.log('📧 Only authorized Vikram Solar R&D team members can access the API');
});

// Handle graceful shutdown
process.on('SIGTERM', () => {
  console.log('SIGTERM signal received: closing server');
  process.exit(0);
});

process.on('SIGINT', () => {
  console.log('SIGINT signal received: closing server');
  process.exit(0);
});
