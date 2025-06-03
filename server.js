// Complete server.js - Express server with Microsoft Authentication
const userMap = {
  "naveen.chamaria@vikramsolar.com": "Naveen Kumar Chamaria",
  "aritra.de@vikramsolar.com": "Aritra De",
  "arindam.halder@vikramsolar.com": "Arindam Halder",
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
  "rnd.lab@vikramsolar.com": "R&D Lab"
};

// Authorized R&D team emails
const AUTHORIZED_EMAILS = [
  "naveen.chamaria@vikramsolar.com",
  "aritra.de@vikramsolar.com",
  "arindam.halder@vikramsolar.com",
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
  "rnd.lab@vikramsolar.com"
];

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const app = express();
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
  origin: process.env.NODE_ENV === 'production' ? 'https://vikramsolar-rnd-rm-dashboard-naveen.netlify.app' : '*',
  methods: ['GET', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'] // Added Authorization header
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

// Function to calculate shrinkage test results from Sheet1 format
function calculateShrinkageResults(shrinkageData) {
  return shrinkageData.map((row, index) => {
    // Parse numeric values from your exact column structure
    const td1WithoutHit = parseFloat(row['TD1']) || 0; // Column C
    const td2WithoutHit = parseFloat(row['TD2']) || 0; // Column D  
    const md1WithoutHit = parseFloat(row['MD1']) || 0; // Column E
    const md2WithoutHit = parseFloat(row['MD2']) || 0; // Column F
    
    // Handle WITH HIT columns (Excel may add __1 suffix for duplicate headers)
    const td1WithHit = parseFloat(row['TD1__1'] || row['TD1_1'] || row['Column7']) || 0; // Column G
    const td2WithHit = parseFloat(row['TD2__1'] || row['TD2_1'] || row['Column8']) || 0; // Column H
    const md1WithHit = parseFloat(row['MD1__1'] || row['MD1_1'] || row['Column9']) || 0; // Column I
    const md2WithHit = parseFloat(row['MD2__1'] || row['MD2_1'] || row['Column10']) || 0; // Column J
    
    // Calculate means
    const tdMeanWithoutHit = (td1WithoutHit + td2WithoutHit) / 2;
    const tdMeanWithHit = (td1WithHit + td2WithHit) / 2;
    const mdMeanWithoutHit = (md1WithoutHit + md2WithoutHit) / 2;
    const mdMeanWithHit = (md1WithHit + md2WithHit) / 2;
    
    // Calculate absolute differences
    const tdDifference = Math.abs(tdMeanWithHit - tdMeanWithoutHit);
    const mdDifference = Math.abs(mdMeanWithHit - mdMeanWithoutHit);
    
    // Determine pass/fail status (less than 1% = pass)
    const tdStatus = tdDifference < 1.0 ? 'PASS' : 'FAIL';
    const mdStatus = mdDifference < 1.0 ? 'PASS' : 'FAIL';
    
    // Final status: PASS only if both TD and MD are PASS
    const finalStatus = (tdStatus === 'PASS' && mdStatus === 'PASS') ? 'PASS' : 'FAIL';
    
    return {
      id: index + 1,
      vendorName: row['VENDOR NAME'] || '',
      encapsulantType: row['ENCAPSULANT TYPE'] || '',
      
      // Raw values
      td1WithoutHit,
      td2WithoutHit,
      md1WithoutHit,
      md2WithoutHit,
      td1WithHit,
      td2WithHit,
      md1WithHit,
      md2WithHit,
      
      // Calculated values
      tdMeanWithoutHit: Math.round(tdMeanWithoutHit * 100) / 100,
      tdMeanWithHit: Math.round(tdMeanWithHit * 100) / 100,
      mdMeanWithoutHit: Math.round(mdMeanWithoutHit * 100) / 100,
      mdMeanWithHit: Math.round(mdMeanWithHit * 100) / 100,
      
      // Differences
      tdDifference: Math.round(tdDifference * 100) / 100,
      mdDifference: Math.round(mdDifference * 100) / 100,
      
      // Status
      tdStatus,
      mdStatus,
      finalStatus
    };
  });
}

// Function to update test results in Test Data based on shrinkage results
function updateTestDataWithShrinkageResults(testData, shrinkageResults) {
  return testData.map(testRow => {
    // Check if this is a shrinkage test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('SHRINKAGE')) {
      const vendorName = testRow['VENDOR NAME'];
      
      // Find matching shrinkage results for this vendor
      const vendorShrinkageData = shrinkageResults.filter(shrinkage => 
        shrinkage.vendorName === vendorName
      );
      
      if (vendorShrinkageData.length > 0) {
        // Check if both FRONT EPE and BACK EVA pass for this vendor
        const frontEpeResult = vendorShrinkageData.find(item => 
          item.encapsulantType === 'FRONT EPE'
        );
        const backEvaResult = vendorShrinkageData.find(item => 
          item.encapsulantType === 'BACK EVA'
        );
        
        let overallResult = 'FAIL';
        
        // Both FRONT EPE and BACK EVA must pass for overall pass
        if (frontEpeResult && backEvaResult) {
          if (frontEpeResult.finalStatus === 'PASS' && backEvaResult.finalStatus === 'PASS') {
            overallResult = 'PASS';
          }
        } else if (frontEpeResult && frontEpeResult.finalStatus === 'PASS') {
          // Only one type tested and it passed
          overallResult = 'PASS';
        } else if (backEvaResult && backEvaResult.finalStatus === 'PASS') {
          // Only one type tested and it passed
          overallResult = 'PASS';
        }
        
        // Update the test result
        testRow['TEST RESULT'] = overallResult;
        testRow['SHRINKAGE_CALCULATION_DETAILS'] = {
          frontEpe: frontEpeResult,
          backEva: backEvaResult,
          overallResult: overallResult
        };
      }
    }
    
    return testRow;
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
    
    // Check for shrinkage tests and read Sheet1 if needed
    const hasShrinkageTests = rawData.some(row => 
      row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('SHRINKAGE')
    );
    
    let shrinkageResults = [];
    if (hasShrinkageTests) {
      // Read Sheet1 for shrinkage data
      const shrinkageSheetName = 'Sheet1';
      const shrinkageSheet = workbook.Sheets[shrinkageSheetName];
      
      if (shrinkageSheet) {
        console.log('Found shrinkage tests, reading Sheet1 for shrinkage data...');
        const shrinkageRawData = xlsx.utils.sheet_to_json(shrinkageSheet);
        
        // Log the raw data structure for debugging
        if (shrinkageRawData.length > 0) {
          console.log('Sample shrinkage raw data structure:', Object.keys(shrinkageRawData[0]));
          console.log('First shrinkage row:', shrinkageRawData[0]);
        }
        
        // Filter out empty rows and header rows
        const validShrinkageData = shrinkageRawData.filter(row => 
          row['VENDOR NAME'] && 
          row['ENCAPSULANT TYPE'] && 
          (row['ENCAPSULANT TYPE'] === 'FRONT EPE' || row['ENCAPSULANT TYPE'] === 'BACK EVA')
        );
        
        console.log(`Found ${validShrinkageData.length} valid shrinkage data rows`);
        
        if (validShrinkageData.length > 0) {
          shrinkageResults = calculateShrinkageResults(validShrinkageData);
          console.log(`Calculated shrinkage results for ${shrinkageResults.length} entries`);
          
          // Log calculated results for debugging
          shrinkageResults.forEach((result, index) => {
            if (index < 3) { // Log first 3 results for debugging
              console.log(`Shrinkage result ${index + 1}:`, {
                vendor: result.vendorName,
                type: result.encapsulantType,
                tdDiff: result.tdDifference,
                mdDiff: result.mdDifference,
                final: result.finalStatus
              });
            }
          });
        }
      } else {
        console.warn('Shrinkage tests found but Sheet1 not available for calculations');
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
    
    // Read the "Sheet1" for shrinkage data
    const shrinkageSheetName = 'Sheet1';
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
    
    if (!solarLabInfo.exists && !lineTrialsInfo.exists && !certificationsInfo.exists) {
      return res.status(404).json({
        success: false,
        message: 'No Excel files found',
        files: {
          solarLabInfo,
          lineTrialsInfo,
          certificationsInfo
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
  
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    requestedBy: userInfo,
    excelFiles: {
      solarLabTests: solarLabInfo,
      lineTrials: lineTrialsInfo,
      certifications: certificationsInfo
    },
    memory: process.memoryUsage(),
    environment: process.env.NODE_ENV || 'development',
    authenticationEnabled: true,
    authorizedUsers: AUTHORIZED_EMAILS.filter(email => email).length
  });
});

// Start the server with more debug info
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT} at ${new Date().toISOString()}`);
  console.log(`🔐 Microsoft Authentication ENABLED`);
  console.log(`👥 Authorized users: ${AUTHORIZED_EMAILS.filter(email => email).length} team members`);
  console.log(`🌐 API available at http://localhost:${PORT}/api/test-data`);
  console.log(`📊 Line Trials API available at http://localhost:${PORT}/api/line-trials`);
  console.log(`📋 Certifications API available at http://localhost:${PORT}/api/certifications`);
  console.log(`🏠 Chamber Tests API available at http://localhost:${PORT}/api/chamber-data`);
  console.log(`🔍 Excel debug endpoint available at http://localhost:${PORT}/api/debug/excel`);
  console.log(`📁 File info endpoint available at http://localhost:${PORT}/api/file-info`);
  console.log(`❤️ Health check available at http://localhost:${PORT}/health`);
  
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
