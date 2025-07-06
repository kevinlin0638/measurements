/**
 * Google Apps Script 代碼
 * 用於處理 HTML 表單提交的資料到 Google Sheets
 * 
 * 設置步驟：
 * 1. 打開 https://script.google.com/
 * 2. 創建新專案
 * 3. 將此代碼貼入 Code.gs 文件
 * 4. 修改下面的 SPREADSHEET_ID 為您的 Google Sheets ID
 * 5. 部署為 Web App
 */

// 請替換為您的 Google Sheets ID
const SPREADSHEET_ID = '13RKGU7mhy-NaFkokHT87tSnPXp5ae0WC0tJjL0wOLKw';

// 工作表名稱
const BOM_SHEET_NAME = 'bom';
const MEASUREMENTS_SHEET_NAME = 'measurements';

/**
 * Handle OPTIONS requests for CORS
 */
function doOptions(e) {
  console.log('=== doOptions called ===');
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE')
    .setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization, Cache-Control')
    .setHeader('Access-Control-Allow-Credentials', 'true')
    .setHeader('Access-Control-Max-Age', '86400');
}

/**
 * Handle POST requests
 */
function doPost(e) {
  console.log('doPost called with:', e);
  
  try {
    // Parse request data from different sources
    let data;
    
    // Check if data is in URL parameters (from URL-encoded request)
    if (e.parameter && e.parameter.data) {
      try {
        data = JSON.parse(e.parameter.data);
        console.log('Parsed request data from parameter:', data);
      } catch (parseError) {
        console.error('JSON parse error from parameter:', parseError);
        return createResponse(false, 'Invalid JSON in parameter data');
      }
    } 
    // Check if data is in POST body (from JSON request)
    else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
        console.log('Parsed request data from postData:', data);
      } catch (parseError) {
        console.error('JSON parse error from postData:', parseError);
        return createResponse(false, 'Invalid JSON in request body');
      }
    } 
    // No data found
    else {
      console.error('No request data found');
      return createResponse(false, 'No request data provided');
    }
    
    const action = data.action;
    const table = data.table;
    const rowData = data.data;
    
    console.log('Request parameters:', { action, table, rowData });
    
    // Validate required parameters
    if (!action) {
      return createResponse(false, 'Missing action parameter');
    }
    
    // Handle different actions
    let result;
    switch(action) {
      case 'add':
        if (!table || !rowData) {
          return createResponse(false, 'Missing table or data parameter');
        }
        if (table === 'bom') {
          result = addBomData(rowData);
        } else if (table === 'measurements') {
          result = addMeasurementsData(rowData);
        } else {
          return createResponse(false, 'Invalid table type: ' + table);
        }
        break;
        
      case 'get':
        if (!table) {
          return createResponse(false, 'Missing table parameter');
        }
        if (table === 'bom') {
          result = getBomData();
        } else if (table === 'measurements') {
          result = getMeasurementsData(rowData?.serial_number);
        } else if (table === 'bom_tree') {
          result = getBomTree();
        } else if (table === 'serial_numbers') {
          result = getSerialNumbers();
        } else {
          return createResponse(false, 'Invalid table type: ' + table);
        }
        break;
        
      case 'update':
        if (!table || !rowData || !rowData.id) {
          return createResponse(false, 'Missing table, data, or id parameter');
        }
        if (table === 'measurements') {
          result = updateMeasurementData(rowData);
        } else {
          return createResponse(false, 'Update not supported for table: ' + table);
        }
        break;
        
      case 'delete':
        if (!table || !rowData || !rowData.id) {
          return createResponse(false, 'Missing table, data, or id parameter');
        }
        if (table === 'measurements') {
          result = deleteMeasurementData(rowData.id);
        } else {
          return createResponse(false, 'Delete not supported for table: ' + table);
        }
        break;
        
      default:
        return createResponse(false, 'Invalid action: ' + action);
    }
    
    console.log('Operation result:', result);
    return createResponse(true, 'Operation completed successfully', result);
    
  } catch (error) {
    console.error('Unexpected error in doPost:', error);
    console.error('Error stack:', error.stack);
    return createResponse(false, 'Server error: ' + error.message + (error.stack ? ' | Stack: ' + error.stack : ''));
  }
}

/**
 * 處理 GET 請求（用於測試）
 */
function doGet(e) {
  return createResponse(true, 'Google Apps Script is running properly');
}

/**
 * Add BOM data
 */
function addBomData(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(BOM_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(BOM_SHEET_NAME);
      // Set header row
      sheet.getRange(1, 1, 1, 5).setValues([
        ['id', 'part_no', 'serial_number', 'assembly_serial_number', 'created_at']
      ]);
    }
    
    // Get next available ID
    const lastRow = sheet.getLastRow();
    const nextId = lastRow === 1 ? 1 : lastRow; // Start from 1 if only header row
    
    // Prepare data to insert
    const rowValues = [
      nextId,
      data.part_no,
      data.serial_number,
      data.assembly_serial_number || '',
      data.created_at
    ];
    
    // Insert data
    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([rowValues]);
    
    return {
      id: nextId,
      table: 'bom',
      row: lastRow + 1
    };
    
  } catch (error) {
    console.error('Error in addBomData:', error);
    throw new Error('Unable to add BOM data: ' + error.message);
  }
}

/**
 * Add Measurements data
 */
function addMeasurementsData(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(MEASUREMENTS_SHEET_NAME);
      // Set header row
      sheet.getRange(1, 1, 1, 5).setValues([
        ['id', 'serial_number', 'para_name', 'para_value', 'created_at']
      ]);
    }
    
    // Get next available ID
    const lastRow = sheet.getLastRow();
    const nextId = lastRow === 1 ? 1 : lastRow; // Start from 1 if only header row
    
    // Prepare data to insert
    const rowValues = [
      nextId,
      data.serial_number,
      data.para_name,
      data.para_value,
      data.created_at
    ];
    
    // Insert data
    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([rowValues]);
    
    return {
      id: nextId,
      table: 'measurements',
      row: lastRow + 1
    };
    
  } catch (error) {
    console.error('Error in addMeasurementsData:', error);
    throw new Error('Unable to add Measurements data: ' + error.message);
  }
}

/**
 * Create unified response format
 */
function createResponse(success, message, data = null) {
  const response = {
    success: success,
    message: message,
    timestamp: new Date().toISOString()
  };
  
  if (data !== null && data !== undefined) {
    response.data = data;
  }
  
  const jsonResponse = JSON.stringify(response);
  console.log('Creating response:', jsonResponse);
  
  return ContentService
    .createTextOutput(jsonResponse)
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, DELETE')
    .setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization, Cache-Control')
    .setHeader('Access-Control-Allow-Credentials', 'true')
    .setHeader('Access-Control-Max-Age', '86400');
}

/**
 * Test function - Can be executed in Google Apps Script editor
 */
function testScript() {
  // Test BOM data
  const bomTestData = {
    part_no: 'TEST001',
    serial_number: 'SN001',
    assembly_serial_number: 'ASN001',
    created_at: new Date().toISOString()
  };
  
  console.log('Testing BOM data insertion...');
  try {
    const bomResult = addBomData(bomTestData);
    console.log('BOM test result:', bomResult);
  } catch (error) {
    console.error('BOM test failed:', error);
  }
  
  // Test Measurements data
  const measurementsTestData = {
    serial_number: 'SN001',
    para_name: 'Temperature',
    para_value: 25.5,
    created_at: new Date().toISOString()
  };
  
  console.log('Testing Measurements data insertion...');
  try {
    const measurementsResult = addMeasurementsData(measurementsTestData);
    console.log('Measurements test result:', measurementsResult);
  } catch (error) {
    console.error('Measurements test failed:', error);
  }
}

/**
 * Initialize sheets (if needed)
 */
function initializeSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Initialize BOM sheet
    let bomSheet = spreadsheet.getSheetByName(BOM_SHEET_NAME);
    if (!bomSheet) {
      bomSheet = spreadsheet.insertSheet(BOM_SHEET_NAME);
      bomSheet.getRange(1, 1, 1, 5).setValues([
        ['id', 'part_no', 'serial_number', 'assembly_serial_number', 'created_at']
      ]);
      console.log('BOM sheet created and initialized');
    }
    
    // Initialize Measurements sheet
    let measurementsSheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    if (!measurementsSheet) {
      measurementsSheet = spreadsheet.insertSheet(MEASUREMENTS_SHEET_NAME);
      measurementsSheet.getRange(1, 1, 1, 5).setValues([
        ['id', 'serial_number', 'para_name', 'para_value', 'created_at']
      ]);
      console.log('Measurements sheet created and initialized');
    }
    
    console.log('Sheets initialization completed');
    
  } catch (error) {
    console.error('Error initializing sheets:', error);
  }
}

/**
 * Get BOM data
 */
function getBomData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(BOM_SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }
    
    // Convert to object array
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
  } catch (error) {
    console.error('Error getting BOM data:', error);
    throw new Error('Unable to retrieve BOM data: ' + error.message);
  }
}

/**
 * Get Measurements data
 */
function getMeasurementsData(serialNumber = null) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }
    
    // Convert to object array
    const headers = data[0];
    const rows = data.slice(1);
    
    let measurements = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    // Filter by serial number if specified
    if (serialNumber) {
      measurements = measurements.filter(m => m.serial_number === serialNumber);
    }
    
    return measurements;
    
  } catch (error) {
    console.error('Error getting measurements data:', error);
    throw new Error('Unable to retrieve measurements data: ' + error.message);
  }
}

/**
 * Get BOM tree structure
 */
function getBomTree() {
  try {
    const bomData = getBomData();
    
    // Build tree structure
    const tree = {};
    const serialToPartMap = {};
    
    // Build serial number to part number mapping
    bomData.forEach(item => {
      serialToPartMap[item.serial_number] = item.part_no;
    });
    
    // Build tree structure
    bomData.forEach(item => {
      const partNo = item.part_no;
      const serialNumber = item.serial_number;
      const assemblySerialNumber = item.assembly_serial_number;
      
      if (!tree[partNo]) {
        tree[partNo] = {
          part_no: partNo,
          serial_numbers: [],
          children: {}
        };
      }
      
      tree[partNo].serial_numbers.push(serialNumber);
      
      // If has assembly serial number, build parent-child relationship
      if (assemblySerialNumber) {
        const parentPartNo = serialToPartMap[assemblySerialNumber];
        if (parentPartNo) {
          if (!tree[parentPartNo]) {
            tree[parentPartNo] = {
              part_no: parentPartNo,
              serial_numbers: [],
              children: {}
            };
          }
          tree[parentPartNo].children[partNo] = tree[partNo];
        }
      }
    });
    
    // Remove items that have become child nodes
    const rootNodes = {};
    Object.keys(tree).forEach(partNo => {
      let isRoot = true;
      Object.keys(tree).forEach(parentPartNo => {
        if (tree[parentPartNo].children[partNo]) {
          isRoot = false;
        }
      });
      if (isRoot) {
        rootNodes[partNo] = tree[partNo];
      }
    });
    
    return rootNodes;
    
  } catch (error) {
    console.error('Error getting BOM tree:', error);
    throw new Error('Unable to retrieve BOM tree: ' + error.message);
  }
}

/**
 * Get all serial numbers
 */
function getSerialNumbers() {
  try {
    const measurements = getMeasurementsData();
    const serialNumbers = [...new Set(measurements.map(m => m.serial_number))];
    return serialNumbers.sort();
    
  } catch (error) {
    console.error('Error getting serial numbers:', error);
    throw new Error('Unable to retrieve serial numbers: ' + error.message);
  }
}

/**
 * Update Measurement data
 */
function updateMeasurementData(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Measurements sheet not found');
    }
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    
    // Find the row to update
    let targetRow = -1;
    let originalSerialNumber = null;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] == data.id) { // id is in the first column
        targetRow = i + 1; // Convert to 1-based index
        originalSerialNumber = allData[i][1]; // serial_number is in the second column
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error('Record not found with ID: ' + data.id);
    }
    
    // Update data - keep original serial number if not provided
    const updatedValues = [
      data.id,
      data.serial_number || originalSerialNumber,
      data.para_name,
      data.para_value,
      new Date().toISOString()
    ];
    
    console.log('Updating row', targetRow, 'with values:', updatedValues);
    sheet.getRange(targetRow, 1, 1, 5).setValues([updatedValues]);
    
    return {
      id: data.id,
      updated: true,
      row: targetRow,
      serial_number: updatedValues[1],
      para_name: updatedValues[2],
      para_value: updatedValues[3]
    };
    
  } catch (error) {
    console.error('Error updating measurement data:', error);
    throw new Error('Unable to update measurement data: ' + error.message);
  }
}

/**
 * Delete Measurement data
 */
function deleteMeasurementData(id) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Measurements sheet not found');
    }
    
    const allData = sheet.getDataRange().getValues();
    
    // Find the row to delete
    let targetRow = -1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] == id) { // id is in the first column
        targetRow = i + 1; // Convert to 1-based index
        break;
      }
    }
    
    if (targetRow === -1) {
      throw new Error('Record not found');
    }
    
    // Delete row
    sheet.deleteRow(targetRow);
    
    return {
      id: id,
      deleted: true,
      row: targetRow
    };
    
  } catch (error) {
    console.error('Error deleting measurement data:', error);
    throw new Error('Unable to delete measurement data: ' + error.message);
  }
} 