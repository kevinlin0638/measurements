/**
 * Google Apps Script 代碼 - 包含 HTML Hosting 功能
 * 完全解決 CORS 問題的方案
 * 
 * 設置步驟：
 * 1. 打開 https://script.google.com/
 * 2. 創建新專案
 * 3. 將此代碼貼入 Code.gs 文件
 * 4. 創建 HTML 文件（index.html）並將 gas-html-page.html 內容貼入
 * 5. 修改下面的 SPREADSHEET_ID 為您的 Google Sheets ID
 * 6. 部署為 Web App
 */

// 請替換為您的 Google Sheets ID
const SPREADSHEET_ID = '13RKGU7mhy-NaFkokHT87tSnPXp5ae0WC0tJjL0wOLKw';

// 工作表名稱
const BOM_SHEET_NAME = 'bom';
const MEASUREMENTS_SHEET_NAME = 'measurements';

/**
 * 處理 GET 請求 - 返回 HTML 頁面
 */
function doGet(e) {
  console.log('=== doGet called - serving HTML page ===');
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Measurement & BOM Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 處理客戶端 API 調用
 */
function handleApiCall(requestData) {
  console.log('=== handleApiCall called ===');
  console.log('Request data:', requestData);
  
  try {
    const action = requestData.action;
    const table = requestData.table;
    const data = requestData.data;
    
    console.log('Processing:', { action, table, data });
    
    // Validate required parameters
    if (!action) {
      throw new Error('Missing action parameter');
    }
    
    // Handle different actions
    let result;
    switch(action) {
      case 'add':
        if (!table || !data) {
          throw new Error('Missing table or data parameter');
        }
        if (table === 'bom') {
          result = addBomData(data);
        } else if (table === 'measurements') {
          result = addMeasurementsData(data);
        } else {
          throw new Error('Invalid table type: ' + table);
        }
        break;
        
      case 'get':
        if (!table) {
          throw new Error('Missing table parameter');
        }
        if (table === 'bom') {
          result = getBomData();
        } else if (table === 'measurements') {
          result = getMeasurementsData(data?.serial_number);
        } else if (table === 'bom_tree') {
          result = getBomTree();
        } else if (table === 'serial_numbers') {
          result = getSerialNumbers();
        } else {
          throw new Error('Invalid table type: ' + table);
        }
        break;
        
      case 'update':
        if (!table || !data || !data.id) {
          throw new Error('Missing table, data, or id parameter');
        }
        if (table === 'measurements') {
          result = updateMeasurementData(data);
        } else {
          throw new Error('Update not supported for table: ' + table);
        }
        break;
        
      case 'delete':
        if (!table || !data || !data.id) {
          throw new Error('Missing table, data, or id parameter');
        }
        if (table === 'measurements') {
          result = deleteMeasurementData(data.id);
        } else {
          throw new Error('Delete not supported for table: ' + table);
        }
        break;
        
      case 'upsert':
        if (!table || !data) {
          throw new Error('Missing table or data parameter');
        }
        if (table === 'measurements') {
          result = upsertMeasurementData(data);
        } else {
          throw new Error('Upsert not supported for table: ' + table);
        }
        break;
        
      case 'add_bom_if_not_exists':
        if (!table || !data) {
          throw new Error('Missing table or data parameter');
        }
        if (table === 'bom') {
          result = addBomDataIfNotExists(data);
        } else {
          throw new Error('add_bom_if_not_exists not supported for table: ' + table);
        }
        break;
        
      default:
        throw new Error('Invalid action: ' + action);
    }
    
    console.log('Operation result:', result);
    return result;
    
  } catch (error) {
    console.error('Error in handleApiCall:', error);
    throw error;
  }
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
        ['id', 'part_no', 'serial_number', 'sub_assembly_serial_number', 'created_at']
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
      data.sub_assembly_serial_number || '',
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
 * Add BOM data if not exists - prevents duplicate BOM relationships
 */
function addBomDataIfNotExists(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(BOM_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = spreadsheet.insertSheet(BOM_SHEET_NAME);
      // Set header row
      sheet.getRange(1, 1, 1, 5).setValues([
        ['id', 'part_no', 'serial_number', 'sub_assembly_serial_number', 'created_at']
      ]);
    }
    
    // Check if relationship already exists
    const existingData = sheet.getDataRange().getValues();
    if (existingData.length > 1) {
      // Skip header row
      const dataRows = existingData.slice(1);
      
      // Check for existing relationship
      const existingRelationship = dataRows.find(row => {
        const [id, part_no, serial_number, sub_assembly_serial_number, created_at] = row;
        return part_no === data.part_no && 
               serial_number === data.serial_number && 
               sub_assembly_serial_number === data.sub_assembly_serial_number;
      });
      
      if (existingRelationship) {
        console.log('BOM relationship already exists:', existingRelationship);
        return {
          created: false,
          existing: true,
          message: 'BOM relationship already exists',
          data: {
            id: existingRelationship[0],
            part_no: existingRelationship[1],
            serial_number: existingRelationship[2],
            sub_assembly_serial_number: existingRelationship[3],
            created_at: existingRelationship[4]
          }
        };
      }
    }
    
    // If not exists, create new relationship
    const lastRow = sheet.getLastRow();
    const nextId = lastRow === 1 ? 1 : lastRow; // Start from 1 if only header row
    
    // Prepare data to insert
    const rowValues = [
      nextId,
      data.part_no,
      data.serial_number,
      data.sub_assembly_serial_number || '',
      data.created_at
    ];
    
    // Insert data
    sheet.getRange(lastRow + 1, 1, 1, 5).setValues([rowValues]);
    
    return {
      created: true,
      existing: false,
      id: nextId,
      table: 'bom',
      row: lastRow + 1
    };
    
  } catch (error) {
    console.error('Error in addBomDataIfNotExists:', error);
    throw new Error('Unable to add BOM data: ' + error.message);
  }
}

/**
 * Add measurements data
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
    throw new Error('Unable to add measurements data: ' + error.message);
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
    
    // Build tree structure where parent contains children as sub_assembly_serial_number
    const tree = {};
    const serialToPartMap = {};
    const childSerials = new Set(); // Track which serial numbers are children
    
    // Build serial number to part number mapping
    bomData.forEach(item => {
      serialToPartMap[item.serial_number] = item.part_no;
    });
    
    // First pass: Create nodes for all parts
    bomData.forEach(item => {
      const partNo = item.part_no;
      const serialNumber = item.serial_number;
      
      // Create node for this part if it doesn't exist
      if (!tree[partNo]) {
        tree[partNo] = {
          part_no: partNo,
          serial_numbers: [],
          children: {}
        };
      }
      
      // Add serial number to part
      if (!tree[partNo].serial_numbers.includes(serialNumber)) {
        tree[partNo].serial_numbers.push(serialNumber);
      }
    });
    
    // Second pass: Build hierarchy - parent contains children
    bomData.forEach(item => {
      const partNo = item.part_no;
      const subAssemblySerialNumber = item.sub_assembly_serial_number;
      
      if (subAssemblySerialNumber && serialToPartMap[subAssemblySerialNumber]) {
        const childPartNo = serialToPartMap[subAssemblySerialNumber];
        
        // Mark this serial number as a child
        childSerials.add(subAssemblySerialNumber);
        
        // Ensure parent node exists
        if (!tree[partNo]) {
          tree[partNo] = {
            part_no: partNo,
            serial_numbers: [],
            children: {}
          };
        }
        
        // Add child part to parent
        if (tree[childPartNo]) {
          tree[partNo].children[childPartNo] = {
            part_no: tree[childPartNo].part_no,
            serial_numbers: [...tree[childPartNo].serial_numbers],
            children: { ...tree[childPartNo].children }
          };
        }
      }
    });
    
    // Third pass: Remove child parts from root level
    childSerials.forEach(serialNumber => {
      const childPartNo = serialToPartMap[serialNumber];
      if (childPartNo) {
        delete tree[childPartNo];
      }
    });
    
    return tree;
    
  } catch (error) {
    console.error('Error getting BOM tree:', error);
    throw new Error('Unable to retrieve BOM tree: ' + error.message);
  }
}

/**
 * Get serial numbers
 */
function getSerialNumbers() {
  try {
    const measurementsData = getMeasurementsData();
    const serialNumbers = [...new Set(measurementsData.map(m => m.serial_number))];
    return serialNumbers.sort();
  } catch (error) {
    console.error('Error getting serial numbers:', error);
    throw new Error('Unable to retrieve serial numbers: ' + error.message);
  }
}

/**
 * Update measurement data
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
    const idIndex = headers.indexOf('id');
    const paraNameIndex = headers.indexOf('para_name');
    const paraValueIndex = headers.indexOf('para_value');
    
    if (idIndex === -1 || paraNameIndex === -1 || paraValueIndex === -1) {
      throw new Error('Required columns not found');
    }
    
    // Find the row with matching ID
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idIndex] == data.id) {
        // Update the row
        sheet.getRange(i + 1, paraNameIndex + 1).setValue(data.para_name);
        sheet.getRange(i + 1, paraValueIndex + 1).setValue(data.para_value);
        
        return {
          id: data.id,
          table: 'measurements',
          row: i + 1,
          updated: true
        };
      }
    }
    
    throw new Error('Measurement with ID ' + data.id + ' not found');
    
  } catch (error) {
    console.error('Error updating measurement data:', error);
    throw new Error('Unable to update measurement data: ' + error.message);
  }
}

/**
 * Delete measurement data
 */
function deleteMeasurementData(id) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MEASUREMENTS_SHEET_NAME);
    
    if (!sheet) {
      throw new Error('Measurements sheet not found');
    }
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const idIndex = headers.indexOf('id');
    
    if (idIndex === -1) {
      throw new Error('ID column not found');
    }
    
    // Find the row with matching ID
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][idIndex] == id) {
        // Delete the row
        sheet.deleteRow(i + 1);
        
        return {
          id: id,
          table: 'measurements',
          row: i + 1,
          deleted: true
        };
      }
    }
    
    throw new Error('Measurement with ID ' + id + ' not found');
    
  } catch (error) {
    console.error('Error deleting measurement data:', error);
    throw new Error('Unable to delete measurement data: ' + error.message);
  }
}

/**
 * Upsert measurement data (Update if exists, Insert if not)
 */
function upsertMeasurementData(data) {
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
    
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const idIndex = headers.indexOf('id');
    const serialNumberIndex = headers.indexOf('serial_number');
    const paraNameIndex = headers.indexOf('para_name');
    const paraValueIndex = headers.indexOf('para_value');
    const createdAtIndex = headers.indexOf('created_at');
    
    if (idIndex === -1 || serialNumberIndex === -1 || paraNameIndex === -1 || paraValueIndex === -1 || createdAtIndex === -1) {
      throw new Error('Required columns not found');
    }
    
    // Find existing row with matching serial_number and para_name
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][serialNumberIndex] === data.serial_number && 
          allData[i][paraNameIndex] === data.para_name) {
        // Update existing row
        sheet.getRange(i + 1, paraValueIndex + 1).setValue(data.para_value);
        sheet.getRange(i + 1, createdAtIndex + 1).setValue(data.created_at);
        
        return {
          id: allData[i][idIndex],
          table: 'measurements',
          row: i + 1,
          action: 'updated'
        };
      }
    }
    
    // No existing row found, insert new one
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
      row: lastRow + 1,
      action: 'inserted'
    };
    
  } catch (error) {
    console.error('Error in upsertMeasurementData:', error);
    throw new Error('Unable to upsert measurement data: ' + error.message);
  }
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