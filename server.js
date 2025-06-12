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
    const td1WithoutHeat = parseFloat(row['TD1']) || 0; // Column C
    const td2WithoutHeat = parseFloat(row['TD2']) || 0; // Column D  
    const md1WithoutHeat = parseFloat(row['MD1']) || 0; // Column E
    const md2WithoutHeat = parseFloat(row['MD2']) || 0; // Column F
    
    // Handle WITH Heat columns (Excel may add __1 suffix for duplicate headers)
    const td1WithHeat = parseFloat(row['TD1__1'] || row['TD1_1'] || row['Column7']) || 0; // Column G
    const td2WithHeat = parseFloat(row['TD2__1'] || row['TD2_1'] || row['Column8']) || 0; // Column H
    const md1WithHeat = parseFloat(row['MD1__1'] || row['MD1_1'] || row['Column9']) || 0; // Column I
    const md2WithHeat = parseFloat(row['MD2__1'] || row['MD2_1'] || row['Column10']) || 0; // Column J
    
    // Calculate means
    const tdMeanWithoutHeat = (td1WithoutHeat + td2WithoutHeat) / 2;
    const tdMeanWithHeat = (td1WithHeat + td2WithHeat) / 2;
    const mdMeanWithoutHeat = (md1WithoutHeat + md2WithoutHeat) / 2;
    const mdMeanWithHeat = (md1WithHeat + md2WithHeat) / 2;
    
    // Calculate absolute differences
    const tdDifference = Math.abs(tdMeanWithHeat - tdMeanWithoutHeat);
    const mdDifference = Math.abs(mdMeanWithHeat - mdMeanWithoutHeat);
    
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
      td1WithoutHeat,
      td2WithoutHeat,
      md1WithoutHeat,
      md2WithoutHeat,
      td1WithHeat,
      td2WithHeat,
      md1WithHeat,
      md2WithHeat,
      
      // Calculated values
      tdMeanWithoutHeat: Math.round(tdMeanWithoutHeat * 100) / 100,
      tdMeanWithHeat: Math.round(tdMeanWithHeat * 100) / 100,
      mdMeanWithoutHeat: Math.round(mdMeanWithoutHeat * 100) / 100,
      mdMeanWithHeat: Math.round(mdMeanWithHeat * 100) / 100,
      
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

// Add this function after the calculateShrinkageResults function in server.js

// Function to calculate adhesion test results from Adhesion sheet
function calculateAdhesionResults(adhesionData) {
  return adhesionData.map((row, index) => {
    // Parse values from the Adhesion sheet columns
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    
    // PRE PCT values (Columns C-F)
    const prePctGlassToEncapMax = parseFloat(row['max'] || row['Column3']) || 0;      // Column C
    const prePctGlassToEncapMin = parseFloat(row['min'] || row['Column4']) || 0;      // Column D  
    const prePctBacksheetToEncapMax = parseFloat(row['max__1'] || row['Column5']) || 0; // Column E
    const prePctBacksheetToEncapMin = parseFloat(row['min__1'] || row['Column6']) || 0; // Column F
    
    // POST PCT values (Columns G-J)
    const postPctGlassToEncapMax = parseFloat(row['max__2'] || row['Column7']) || 0;     // Column G
    const postPctGlassToEncapMin = parseFloat(row['min__2'] || row['Column8']) || 0;     // Column H
    const postPctBacksheetToEncapMax = parseFloat(row['max__3'] || row['Column9']) || 0; // Column I
    const postPctBacksheetToEncapMin = parseFloat(row['min__3'] || row['Column10']) || 0; // Column J
    
    // Calculate averages for each category
    const prePctGlassToEncapAvg = (prePctGlassToEncapMax + prePctGlassToEncapMin) / 2;
    const prePctBacksheetToEncapAvg = (prePctBacksheetToEncapMax + prePctBacksheetToEncapMin) / 2;
    const postPctGlassToEncapAvg = (postPctGlassToEncapMax + postPctGlassToEncapMin) / 2;
    const postPctBacksheetToEncapAvg = (postPctBacksheetToEncapMax + postPctBacksheetToEncapMin) / 2;
    
    // Apply pass/fail criteria
    const prePctGlassToEncapStatus = prePctGlassToEncapAvg > 60 ? 'PASS' : 'FAIL';
    const prePctBacksheetToEncapStatus = prePctBacksheetToEncapAvg > 40 ? 'PASS' : 'FAIL';
    const postPctGlassToEncapStatus = postPctGlassToEncapAvg > 60 ? 'PASS' : 'FAIL';
    const postPctBacksheetToEncapStatus = postPctBacksheetToEncapAvg > 40 ? 'PASS' : 'FAIL';
    
    // Final status: PASS only if ALL four pass
    const finalStatus = (
      prePctGlassToEncapStatus === 'PASS' && 
      prePctBacksheetToEncapStatus === 'PASS' && 
      postPctGlassToEncapStatus === 'PASS' && 
      postPctBacksheetToEncapStatus === 'PASS'
    ) ? 'PASS' : 'FAIL';
    
    return {
      id: index + 1,
      vendorName,
      bom,
      
      // Raw values
      prePctGlassToEncapMax,
      prePctGlassToEncapMin,
      prePctBacksheetToEncapMax,
      prePctBacksheetToEncapMin,
      postPctGlassToEncapMax,
      postPctGlassToEncapMin,
      postPctBacksheetToEncapMax,
      postPctBacksheetToEncapMin,
      
      // Calculated averages
      prePctGlassToEncapAvg: Math.round(prePctGlassToEncapAvg * 100) / 100,
      prePctBacksheetToEncapAvg: Math.round(prePctBacksheetToEncapAvg * 100) / 100,
      postPctGlassToEncapAvg: Math.round(postPctGlassToEncapAvg * 100) / 100,
      postPctBacksheetToEncapAvg: Math.round(postPctBacksheetToEncapAvg * 100) / 100,
      
      // Individual status
      prePctGlassToEncapStatus,
      prePctBacksheetToEncapStatus,
      postPctGlassToEncapStatus,
      postPctBacksheetToEncapStatus,
      
      // Final status
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
function calculateTensileStrengthResults(tensileData) {
  return tensileData.map((row, index) => {
    // Parse values from the Tensile Strength sheet columns
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    
    // Parse Break value (Column C - should be > 10 MPa)
    const breakValue = parseFloat(row['BREAK']) || 0;
    
    // Parse Change in Elongation % (Column D)
    let changeInElongationPercent = parseFloat(row['CHANGE IN ELONGATION %']) || 0;
    
    // Check if we have initial and final length values instead
    const initialLength = parseFloat(row['INITIAL LENGTH'] || row['Column5']) || 0;
    const finalLength = parseFloat(row['FINAL LENGTH'] || row['Column6']) || 0;
    
    // Calculate elongation % if initial and final lengths are provided
    if (initialLength > 0 && finalLength > 0 && changeInElongationPercent === 0) {
      changeInElongationPercent = ((finalLength - initialLength) * 100) / initialLength;
      console.log(`Calculated elongation % for ${vendorName}: ${changeInElongationPercent}% (Initial: ${initialLength}, Final: ${finalLength})`);
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
      initialLength: initialLength > 0 ? Math.round(initialLength * 100) / 100 : null,
      finalLength: finalLength > 0 ? Math.round(finalLength * 100) / 100 : null,
      
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

// Function to update test results in Test Data based on tensile strength results
function updateTestDataWithTensileStrengthResults(testData, tensileResults) {
  return testData.map(testRow => {
    // Check if this is a tensile strength test
    if (testRow['TEST NAME'] && testRow['TEST NAME'].toUpperCase().includes('TENSILE')) {
      const vendorName = testRow['VENDOR NAME'];
      const bomType = testRow['BOM'];
      
      // Find matching tensile strength results for this vendor and BOM
      const tensileResult = tensileResults.find(result => 
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
      }
    }
    
    return testRow;
  });
}

// Function to calculate GSM test results from GSM sheet
function calculateGSMResults(gsmData) {
  return gsmData.map((row, index) => {
    // Parse values from the GSM sheet columns
    const vendorName = row['VENDOR NAME'] || '';
    const bom = row['BOM'] || '';
    const category = row['CATEGORY'] || row['TYPE'] || ''; // OLD or NEW
    
    // Parse the 5 measurement values
    const value1 = parseFloat(row['VALUE 1'] || row['MIN VALUE 1'] || row['MEASUREMENT 1']) || 0;
    const value2 = parseFloat(row['VALUE 2'] || row['MIN VALUE 2'] || row['MEASUREMENT 2']) || 0;
    const value3 = parseFloat(row['VALUE 3'] || row['MIN VALUE 3'] || row['MEASUREMENT 3']) || 0;
    const value4 = parseFloat(row['VALUE 4'] || row['MIN VALUE 4'] || row['MEASUREMENT 4']) || 0;
    const value5 = parseFloat(row['VALUE 5'] || row['MIN VALUE 5'] || row['MEASUREMENT 5']) || 0;
    
    // Calculate average
    const measurements = [value1, value2, value3, value4, value5];
    const validMeasurements = measurements.filter(val => val > 0);
    const average = validMeasurements.length > 0 ? 
      validMeasurements.reduce((sum, val) => sum + val, 0) / validMeasurements.length : 0;
    
    // Define ranges based on category
    let minRange, maxRange, rangeName;
    
    if (category.toUpperCase().includes('OLD')) {
      minRange = 420;
      maxRange = 480;
      rangeName = 'OLD: 420 to 480';
    } else if (category.toUpperCase().includes('NEW')) {
      minRange = 380;
      maxRange = 440;
      rangeName = 'NEW: 380 to 440';
    } else {
      // Try to parse range from the data itself if category is not clear
      const rangeText = row['RANGE'] || '';
      if (rangeText.includes('420') && rangeText.includes('480')) {
        minRange = 420;
        maxRange = 480;
        rangeName = 'OLD: 420 to 480';
      } else if (rangeText.includes('380') && rangeText.includes('440')) {
        minRange = 380;
        maxRange = 440;
        rangeName = 'NEW: 380 to 440';
      } else {
        // Default to OLD range if unclear
        minRange = 420;
        maxRange = 480;
        rangeName = 'OLD: 420 to 480 (default)';
      }
    }
    
    // Determine pass/fail status
    const isWithinRange = average >= minRange && average <= maxRange;
    const finalStatus = isWithinRange ? 'PASS' : 'FAIL';
    
    // Debug logging for first few rows
    if (index < 3) {
      console.log(`\nGSM Row ${index + 1} Debug:`);
      console.log(`- Vendor: ${vendorName}, BOM: ${bom}`);
      console.log(`- Category: ${category}`);
      console.log(`- Measurements: [${measurements.join(', ')}]`);
      console.log(`- Valid measurements: ${validMeasurements.length}`);
      console.log(`- Average: ${average}`);
      console.log(`- Range: ${rangeName} (${minRange}-${maxRange})`);
      console.log(`- Status: ${finalStatus}`);
    }
    
    return {
      id: index + 1,
      vendorName,
      bom,
      category,
      
      // Raw measurement values
      value1,
      value2,
      value3,
      value4,
      value5,
      
      // Calculated values
      average: Math.round(average * 100) / 100,
      validMeasurementCount: validMeasurements.length,
      
      // Range information
      minRange,
      maxRange,
      rangeName,
      
      // Status
      isWithinRange,
      finalStatus,
      
      // Additional info
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

// Function to process resistance test results from Resistance sheet
function processResistanceResults(resistanceData) {
  return resistanceData.map((row, index) => {
    // Parse values from the Resistance sheet columns
    const bom = row['BOM'] || '';
    const type = row['TYPE'] || ''; // BUS RIBBON or INTERCONNECT RIBBON
    const vendorName = row['VENDOR NAME'] || '';
    const measuredValue = parseFloat(row['MEASURED VALUE']) || 0;
    
    // Since there's no specific criteria, we'll look for a result column
    // or assume it will be manually updated in the Test Data sheet
    const testResult = row['TEST RESULT'] || row['RESULT'] || row['STATUS'] || 'Pending';
    
    // Debug logging for first few rows
    if (index < 3) {
      console.log(`\nResistance Row ${index + 1} Debug:`);
      console.log(`- BOM: ${bom}`);
      console.log(`- Type: ${type}`);
      console.log(`- Vendor: ${vendorName}`);
      console.log(`- Measured Value: ${measuredValue}`);
      console.log(`- Test Result: ${testResult}`);
    }
    
    return {
      id: index + 1,
      bom,
      type, // BUS RIBBON or INTERCONNECT RIBBON
      vendorName,
      measuredValue: Math.round(measuredValue * 1000000) / 1000000, // Round to 6 decimal places for precision
      testResult: testResult.toUpperCase() === 'PASS' ? 'PASS' : 
                 testResult.toUpperCase() === 'FAIL' ? 'FAIL' : 'Pending',
      
      // Additional metadata
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
function processBypassDiodeResults(bypassDiodeData) {
  return bypassDiodeData.map((row, index) => {
    // Parse values from the BYPASS DIODE TEST sheet columns
    const bom = row['BOM'] || '';
    const vendorName = row['VENDOR NAME'] || '';
    const maxTemperatureTj = parseFloat(row['MAX TEMPERATURE OF DIODE(Tj)'] || row['MAX TEMPERATURE OF DIODE (Tj)'] || row['Tj']) || 0;
    
    // Look for manual test result entry
    const testResult = row['TEST RESULT'] || row['RESULT'] || row['STATUS'] || row['PASS/FAIL'] || 'Pending';
    
    // Additional fields that might be present
    const notes = row['NOTES'] || row['REMARKS'] || row['COMMENTS'] || '';
    const testDate = row['TEST DATE'] || row['DATE'] || '';
    const testedBy = row['TESTED BY'] || row['OPERATOR'] || '';
    
    // Debug logging for first few rows
    if (index < 3) {
      console.log(`\nBypass Diode Row ${index + 1} Debug:`);
      console.log(`- BOM: ${bom}`);
      console.log(`- Vendor: ${vendorName}`);
      console.log(`- Max Temperature (Tj): ${maxTemperatureTj}°C`);
      console.log(`- Test Result: ${testResult}`);
      console.log(`- Notes: ${notes}`);
    }
    
    // Normalize test result
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
      id: index + 1,
      bom,
      vendorName,
      maxTemperatureTj: Math.round(maxTemperatureTj * 100) / 100, // Round to 2 decimal places
      testResult: normalizedResult,
      
      // Additional metadata
      hasValidTemperature: maxTemperatureTj > 0,
      temperatureUnit: '°C',
      notes,
      testDate,
      testedBy,
      
      // Assessment metadata
      assessmentMethod: 'Manual Operator Assessment',
      criteria: 'Manual evaluation based on Max Temperature (Tj) and operational requirements'
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

    // Add this section after the shrinkage calculation section in the /api/test-data endpoint:

    // Check for adhesion tests and read Adhesion sheet if needed
    const hasAdhesionTests = rawData.some(row => 
      row['TEST NAME'] && row['TEST NAME'].toUpperCase().includes('ADHESION')
    );
    
    let adhesionResults = [];
    if (hasAdhesionTests) {
      // Read Adhesion sheet for adhesion data
      const adhesionSheetName = 'Adhesion';
      const adhesionSheet = workbook.Sheets[adhesionSheetName];
      
      if (adhesionSheet) {
        console.log('Found adhesion tests, reading Adhesion sheet for adhesion data...');
        const adhesionRawData = xlsx.utils.sheet_to_json(adhesionSheet);
        
        // Log the raw data structure for debugging
        if (adhesionRawData.length > 0) {
          console.log('Sample adhesion raw data structure:', Object.keys(adhesionRawData[0]));
          console.log('First adhesion row:', adhesionRawData[0]);
        }
        
        // Filter out empty rows and header rows
        const validAdhesionData = adhesionRawData.filter(row => 
          row['VENDOR NAME'] && row['BOM']
        );
        
        console.log(`Found ${validAdhesionData.length} valid adhesion data rows`);
        
        if (validAdhesionData.length > 0) {
          adhesionResults = calculateAdhesionResults(validAdhesionData);
          console.log(`Calculated adhesion results for ${adhesionResults.length} entries`);
          
          // Log calculated results for debugging
          adhesionResults.forEach((result, index) => {
            if (index < 3) { // Log first 3 results for debugging
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
  // Read GSM sheet for GSM test data
  const gsmSheetName = 'GSM';
  const gsmSheet = workbook.Sheets[gsmSheetName];
  
  if (gsmSheet) {
    console.log('Found GSM tests, reading GSM sheet for GSM data...');
    const gsmRawData = xlsx.utils.sheet_to_json(gsmSheet);
    
    // Log the raw data structure for debugging
    if (gsmRawData.length > 0) {
      console.log('Sample GSM raw data structure:', Object.keys(gsmRawData[0]));
      console.log('First GSM row:', gsmRawData[0]);
    }
    
    // Filter out empty rows and header rows
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
  // Read Resistance sheet for resistance test data
  const resistanceSheetName = 'Resistance';
  const resistanceSheet = workbook.Sheets[resistanceSheetName];
  
  if (resistanceSheet) {
    console.log('Found resistance tests, reading Resistance sheet for resistance data...');
    const resistanceRawData = xlsx.utils.sheet_to_json(resistanceSheet);
    
    // Log the raw data structure for debugging
    if (resistanceRawData.length > 0) {
      console.log('Sample resistance raw data structure:', Object.keys(resistanceRawData[0]));
      console.log('First resistance row:', resistanceRawData[0]);
    }
    
    // Filter out empty rows and header rows
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
  // Read BYPASS DIODE TEST sheet for bypass diode test data
  const bypassDiodeSheetName = 'BYPASS DIODE TEST';
  const bypassDiodeSheet = workbook.Sheets[bypassDiodeSheetName];
  
  if (bypassDiodeSheet) {
    console.log('Found bypass diode tests, reading BYPASS DIODE TEST sheet for bypass diode data...');
    const bypassDiodeRawData = xlsx.utils.sheet_to_json(bypassDiodeSheet);
    
    // Log the raw data structure for debugging
    if (bypassDiodeRawData.length > 0) {
      console.log('Sample bypass diode raw data structure:', Object.keys(bypassDiodeRawData[0]));
      console.log('First bypass diode row:', bypassDiodeRawData[0]);
    }
    
    // Filter out empty rows and header rows
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
