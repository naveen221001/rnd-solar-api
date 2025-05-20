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
  origin: process.env.NODE_ENV === 'production' ? 'https://vikramsolar-rnd-rm-dashboard-naveen.netlify.app' : '*',
  methods: ['GET', 'OPTIONS'],
  allowedHeaders: ['Content-Type']
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

// API endpoint for test data
app.get('/api/test-data', (req, res) => {
  try {
    console.log(`API request received for /api/test-data at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
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
        // Set type to 'binary' for better handling of large files
        type: 'binary',
        // Force a reload of the file
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

// API endpoint for line trials data
app.get('/api/line-trials', (req, res) => {
  try {
    console.log(`API request received for /api/line-trials at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
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

// API endpoint for certifications data
app.get('/api/certifications', (req, res) => {
  try {
    console.log(`API request received for /api/certifications at ${new Date().toISOString()}`);
    
    // Add request info to response for debugging
    const requestInfo = {
      timestamp: new Date().toISOString(),
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
    
    // We expect multiple sheets for different certification statuses
    const certSheets = {
      completed: workbook.SheetNames.find(name => name.toLowerCase().includes('completed')) || workbook.SheetNames[0],
      inProcess: workbook.SheetNames.find(name => name.toLowerCase().includes('process') || name.toLowerCase().includes('progress')) || workbook.SheetNames[1],
      pending: workbook.SheetNames.find(name => name.toLowerCase().includes('pending')) || workbook.SheetNames[2]
    };
    
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
    wattpeak: row['WATTPEAK'] || '', // Add wattpeak
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
          wattpeak: row['WATTPEAK'] || '', // Add wattpeak
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
          wattpeak: row['WATTPEAK'] || '', // Add wattpeak
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

// Fix for the api/solar-data endpoint
app.get('/api/solar-data', (req, res) => {
  console.log('Received request to /api/solar-data, redirecting to /api/test-data');
  // This redirects api/solar-data requests to api/test-data
  req.url = '/api/test-data';
  app.handle(req, res);
});

// Metadata endpoint with improved error handling
app.get('/api/metadata', (req, res) => {
  try {
    console.log(`API request received for /api/metadata at ${new Date().toISOString()}`);
    
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
        generatedAt: new Date().toISOString()
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

// Enhanced file status endpoint
app.get('/api/data-status', (req, res) => {
  try {
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
        }
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
      serverTime: new Date().toISOString()
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

// New file info endpoint
app.get('/api/file-info', (req, res) => {
  try {
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
        memoryUsage: process.memoryUsage()
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

// Debug endpoint with improved error handling
app.get('/api/debug/excel/:file?', (req, res) => {
  try {
    const fileName = req.params.file || 'Solar_Lab_Tests.xlsx';
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
      serverTime: new Date().toISOString()
    });
    
  } catch (error) {
    res.status(500).json({ 
      error: 'Failed to read Excel structure', 
      details: error.message 
    });
  }
});

// Enhanced health check endpoint
app.get('/health', (req, res) => {
  // Check all Excel files
  const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
  const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
  const certificationsInfo = checkExcelFile('Certifications.xlsx');
  
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    uptime: process.uptime(),
    excelFiles: {
      solarLabTests: solarLabInfo,
      lineTrials: lineTrialsInfo,
      certifications: certificationsInfo
    },
    memory: process.memoryUsage()
  });
});

// Start the server with more debug info
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT} at ${new Date().toISOString()}`);
  console.log(`API available at http://localhost:${PORT}/api/test-data`);
  console.log(`Line Trials API available at http://localhost:${PORT}/api/line-trials`);
  console.log(`Certifications API available at http://localhost:${PORT}/api/certifications`);
  console.log(`Excel debug endpoint available at http://localhost:${PORT}/api/debug/excel`);
  console.log(`File info endpoint available at http://localhost:${PORT}/api/file-info`);
  
  // Check Excel files on startup
  const solarLabInfo = checkExcelFile('Solar_Lab_Tests.xlsx');
  const lineTrialsInfo = checkExcelFile('Line_Trials.xlsx');
  const certificationsInfo = checkExcelFile('Certifications.xlsx');
  
  if (solarLabInfo.exists) {
    console.log(`Solar Lab Tests Excel file is ready at ${solarLabInfo.path}, size: ${solarLabInfo.size} bytes`);
  } else {
    console.log(`Waiting for Solar Lab Tests Excel file to be synced to ${solarLabInfo.path}`);
  }
  
  if (lineTrialsInfo.exists) {
    console.log(`Line Trials Excel file is ready at ${lineTrialsInfo.path}, size: ${lineTrialsInfo.size} bytes`);
  } else {
    console.log(`Waiting for Line Trials Excel file to be synced to ${lineTrialsInfo.path}`);
  }
  
  if (certificationsInfo.exists) {
    console.log(`Certifications Excel file is ready at ${certificationsInfo.path}, size: ${certificationsInfo.size} bytes`);
  } else {
    console.log(`Waiting for Certifications Excel file to be synced to ${certificationsInfo.path}`);
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
