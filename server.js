// server.js - Express server to serve Excel data as API
const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3001;

// Enable CORS for your Netlify domain - replace with your actual Netlify URL
app.use(cors({
  origin: '*' // In production, change to your specific Netlify URL
}));

// Helper function to safely parse dates from Excel
function parseExcelDate(dateValue) {
  // If the value is null or undefined, return null
  if (dateValue == null) return null;
  
  // If it's already a Date object, return it
  if (dateValue instanceof Date) return dateValue;
  
  try {
    // If it's a number, it's likely an Excel serial date
    if (typeof dateValue === 'number') {
      // Excel date serial numbers start from January 1, 1900
      // Need to adjust by the Excel leap year bug
      const excelEpoch = new Date(Date.UTC(1899, 11, 30));
      const millisecondsPerDay = 24 * 60 * 60 * 1000;
      return new Date(excelEpoch.getTime() + dateValue * millisecondsPerDay);
    }
    
    // If it's a string in the format DD/MM/YYYY, convert to YYYY-MM-DD for parsing
    if (typeof dateValue === 'string' && dateValue.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
      const [day, month, year] = dateValue.split('/');
      return new Date(Date.UTC(year, month - 1, day));
    }
    
    // Otherwise try to parse as a regular date string
    return new Date(String(dateValue));
  } catch (e) {
    console.error('Error parsing date:', dateValue, e);
    return null;
  }
}

// Define standard test durations manually as a backup
// This ensures we have the correct values even if the Excel sheet doesn't load properly
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

// API endpoint to read Excel data
app.get('/api/test-data', (req, res) => {
  try {
    console.log('API request received for /api/test-data');
    
    // Path to your Excel file
    const excelFilePath = path.join(__dirname, 'data', 'Solar_Lab_Tests.xlsx');
    
    // Check if file exists
    if (!fs.existsSync(excelFilePath)) {
      console.error('Excel file not found at path:', excelFilePath);
      return res.status(404).json({ error: 'Excel file not found' });
    }
    
    // Read the Excel file with options for better date handling
    const workbook = xlsx.readFile(excelFilePath, {
      cellDates: true,  // This option tells xlsx to parse dates
      dateNF: 'yyyy-mm-dd',  // Date format
      cellNF: true,     // Keep number formats
      cellStyles: true  // Keep cell styles
    });
    
    // Log available sheets for debugging
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
      // Get column names based on your Excel sheet (case sensitive)
      const grnTimeValue = row['GRN GENERATION TIME'];
      const testStartValue = row['TEST START DATE AND TIME'];
      const testEndValue = row['TEST END DATE AND TIME'];
      
      // Parse dates from Excel format to JavaScript Date objects
      const grnTime = parseExcelDate(grnTimeValue);
      const startTime = parseExcelDate(testStartValue);
      const endTime = parseExcelDate(testEndValue);
      
      // Debug first few rows
      if (index < 5) {
        console.log(`Row ${index + 1}:`, {
          testName: row['TEST NAME'],
          grnTime: grnTime ? grnTime.toISOString() : null,
          startTime: startTime ? startTime.toISOString() : null,
          endTime: endTime ? endTime.toISOString() : null,
        });
      }
      
      // Calculate actual test duration in days
      let actualDuration = 0;
      if (startTime && endTime) {
        const diffTime = Math.abs(endTime - startTime);
        actualDuration = Math.max(1, Math.ceil(diffTime / (1000 * 60 * 60 * 24))); // Minimum 1 day
      }
      
      // Get standard duration for this test (with exact name matching)
      const testName = row['TEST NAME'] || '';
      const standardDuration = standardDurations[testName] || DEFAULT_STD_DURATION;
      
      // Calculate efficiency (standard/actual * 100)
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
    
    // Debug: Show sample of processed data
    if (processedData.length > 0) {
      // Check specific test durations
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

// API endpoint to get all unique BOMs and tests
app.get('/api/metadata', (req, res) => {
  try {
    console.log('API request received for /api/metadata');
    
    // Path to your Excel file
    const excelFilePath = path.join(__dirname, 'data', 'Solar_Lab_Tests.xlsx');
    
    // Read the Excel file
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

// Add a debug endpoint to see Excel structure
app.get('/api/debug/excel', (req, res) => {
  try {
    const excelFilePath = path.join(__dirname, 'data', 'Solar_Lab_Tests.xlsx');
    if (!fs.existsSync(excelFilePath)) {
      return res.status(404).json({ error: 'Excel file not found' });
    }
    
    const workbook = xlsx.readFile(excelFilePath, { cellDates: true });
    
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

// Add a health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  console.log(`API available at http://localhost:${PORT}/api/test-data`);
  console.log(`Excel debug endpoint available at http://localhost:${PORT}/api/debug/excel`);
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