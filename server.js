// server.js - Express server to serve Excel data as API
const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3001;

// Enable CORS for your Netlify domain
app.use(cors({
  origin: '*' // In production, change to your specific Netlify URL
}));

// Helper function to safely parse dates from Excel
function parseExcelDate(dateValue) {
  // Your existing parseExcelDate function is excellent - keeping it as is
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

// Define standard test durations manually as a backup
const HARDCODED_STD_DURATIONS = {
  // Your existing durations - keeping them as is
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

// Add a file check function to verify Excel file exists
function checkExcelFileExists() {
  const excelFilePath = path.join(__dirname, 'data', 'Solar_Lab_Tests.xlsx');
  const exists = fs.existsSync(excelFilePath);
  
  if (!exists) {
    console.warn(`Excel file not found at ${excelFilePath}`);
  } else {
    console.log(`Excel file found at ${excelFilePath}`);
  }
  
  return {
    exists,
    path: excelFilePath
  };
}

// API endpoints (your existing endpoints)
app.get('/api/test-data', (req, res) => {
  try {
    console.log('API request received for /api/test-data');
    
    // Check if the Excel file exists
    const fileCheck = checkExcelFileExists();
    if (!fileCheck.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        path: fileCheck.path
      });
    }
    
    // Rest of your existing code...
    const excelFilePath = fileCheck.path;
    const workbook = xlsx.readFile(excelFilePath, {
      cellDates: true,
      dateNF: 'yyyy-mm-dd',
      cellNF: true,
      cellStyles: true
    });
    
    // Your existing processing logic...
    console.log('Available sheets in workbook:', workbook.SheetNames);
    
    // Read the "Test Data" sheet
    const testDataSheetName = 'Test Data';
    const worksheet = workbook.Sheets[testDataSheetName];
    if (!worksheet) {
      console.error(`Sheet "${testDataSheetName}" not found. Available sheets:`, workbook.SheetNames);
      return res.status(404).json({ 
        error: `${testDataSheetName} sheet not found in Excel file`,
        availableSheets: workbook.SheetNames
      });
    }
    
    // Convert to JSON
    const rawData = xlsx.utils.sheet_to_json(worksheet);
    console.log(`Processed ${rawData.length} rows from ${testDataSheetName} sheet`);
    
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
      // Make sure to use the exact column headers from your Excel sheet
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
    
    // Check for missing standard durations
    const allTestNames = new Set(rawData.map(row => row['TEST NAME']).filter(Boolean));
    const missingStandards = [...allTestNames].filter(test => standardDurations[test] === undefined);
    if (missingStandards.length > 0) {
      console.warn('Tests missing standard durations (will use default value):', missingStandards);
    }
    
    // Default duration if not found in either source
    const DEFAULT_STD_DURATION = 2; // 2 days
    
    // Process the data to match the dashboard's expected format
    const processedData = rawData.map((row, index) => {
      // Your existing data processing code...
      const grnTimeValue = row['GRN GENERATION TIME'];
      const testStartValue = row['TEST START DATE AND TIME'];
      const testEndValue = row['TEST END DATE AND TIME'];
      
      const grnTime = parseExcelDate(grnTimeValue);
      const startTime = parseExcelDate(testStartValue);
      const endTime = parseExcelDate(testEndValue);
      
      if (index < 5) {
        console.log(`Row ${index + 1}:`, {
          testName: row['TEST NAME'],
          grnTime: grnTime ? grnTime.toISOString() : null,
          startTime: startTime ? startTime.toISOString() : null,
          endTime: endTime ? endTime.toISOString() : null,
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
        efficiency: efficiency
      };
    });
    
    if (processedData.length > 0) {
      // Your existing debugging code...
      const adhesionTests = processedData.filter(item => item.test === 'ADHESION - PCT - ADHESION');
      if (adhesionTests.length > 0) {
        console.log('ADHESION - PCT - ADHESION tests:', adhesionTests.map(t => ({ 
          id: t.id, 
          standardDuration: t.standardDuration 
        })));
      }
      
      console.log('Sample of processed data:', processedData.slice(0, 3));
    }
    
    res.json(processedData);
  } catch (error) {
    console.error('Error reading Excel file:', error);
    res.status(500).json({ 
      error: 'Failed to read Excel data', 
      details: error.message,
      stack: error.stack 
    });
  }
});

// Your existing metadata endpoint
app.get('/api/metadata', (req, res) => {
  try {
    console.log('API request received for /api/metadata');
    
    // Check if the Excel file exists
    const fileCheck = checkExcelFileExists();
    if (!fileCheck.exists) {
      return res.status(404).json({ 
        error: 'Excel file not found',
        message: 'The Excel file has not been synced yet from OneDrive. Please wait for the GitHub Action to run.',
        path: fileCheck.path
      });
    }
    
    // Rest of your existing code...
    const excelFilePath = fileCheck.path;
    const workbook = xlsx.readFile(excelFilePath, {
      cellDates: true
    });
    
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
      uniqueTests: Array.from(uniqueTests)
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

// Add a new endpoint to check the Excel file status
app.get('/api/data-status', (req, res) => {
  try {
    const fileCheck = checkExcelFileExists();
    
    if (!fileCheck.exists) {
      return res.status(404).json({
        success: false,
        message: 'Excel file not found',
        path: fileCheck.path
      });
    }
    
    const stats = fs.statSync(fileCheck.path);
    
    res.json({
      success: true,
      lastUpdated: stats.mtime,
      fileSize: stats.size,
      fileName: 'Solar_Lab_Tests.xlsx',
      path: fileCheck.path
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

// Your existing debug endpoint
app.get('/api/debug/excel', (req, res) => {
  try {
    const fileCheck = checkExcelFileExists();
    if (!fileCheck.exists) {
      return res.status(404).json({ error: 'Excel file not found' });
    }
    
    const workbook = xlsx.readFile(fileCheck.path, { cellDates: true });
    
    // Get information about each sheet
    const sheetsInfo = {};
    workbook.SheetNames.forEach(name => {
      const sheet = workbook.Sheets[name];
      const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1:A1');
      const sampleData = xlsx.utils.sheet_to_json(sheet, { header: 1, range: 0, defval: null }).slice(0, 2);
      
      sheetsInfo[name] = {
        rowCount: range.e.r - range.s.r + 1,
        columnCount: range.e.c - range.s.c + 1,
        columnNames: sampleData[0] || [],
        sampleRow: sampleData[1] || []
      };
    });
    
    res.json({
      fileName: 'Solar_Lab_Tests.xlsx',
      sheets: sheetsInfo
    });
    
  } catch (error) {
    res.status(500).json({ 
      error: 'Failed to read Excel structure', 
      details: error.message 
    });
  }
});

// Your existing health check endpoint
app.get('/health', (req, res) => {
  // Add Excel file status to health check
  const fileCheck = checkExcelFileExists();
  
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    excelFile: {
      exists: fileCheck.exists,
      path: fileCheck.path,
      lastUpdated: fileCheck.exists ? fs.statSync(fileCheck.path).mtime : null
    }
  });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`API available at http://localhost:${PORT}/api/test-data`);
  console.log(`Excel debug endpoint available at http://localhost:${PORT}/api/debug/excel`);
  
  // Check Excel file on startup
  const fileCheck = checkExcelFileExists();
  if (fileCheck.exists) {
    console.log(`Excel file is ready at ${fileCheck.path}`);
  } else {
    console.log(`Waiting for Excel file to be synced to ${fileCheck.path}`);
  }
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