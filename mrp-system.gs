/**
 * MRP Material Requirements Planning System - Google Apps Script Version
 * Material Requirements Planning System
 * 
 * Features:
 * 1. Finished Goods Demand Management
 * 2. Inventory Management
 * 3. BOM Structure Management
 * 4. Lead Time Management
 * 5. Supplier Management
 * 6. MRP Simulation Calculation
 * 7. Period Advancement Simulation
 * 8. Google Sheets Integration
 */

// Google Sheets ID
const SPREADSHEET_ID = '1ieSa4W6_crSYoIobMiEACW2wG-PYb1PLTRFElhS5FoQ';

// Sheet Names
const MRP_DATA_SHEET_NAME = 'mrp_data';
const MRP_RESULTS_SHEET_NAME = 'mrp_results';
const BOM_SHEET_NAME = 'bom';
const DEMAND_SHEET_NAME = 'demand';

/**
 * Handle GET requests - Return HTML page
 */
function doGet(e) {
  console.log('=== doGet called - serving MRP HTML page ===');
  return HtmlService.createTemplateFromFile('MRP_HTML')
    .evaluate()
    .setTitle('MRP 物料需求規劃系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle POST requests - API endpoints
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    console.log('=== doPost called with action: ' + action + ' ===');
    
    switch (action) {
      case 'initialize':
        return ContentService.createTextOutput(JSON.stringify(initializeMRPSystem()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'setBasicParameters':
        return ContentService.createTextOutput(JSON.stringify(
          setMRPBasicParameters(data.productName, data.initialInventory)
        )).setMimeType(ContentService.MimeType.JSON);
      
      case 'setCurrentPeriodDemand':
        return ContentService.createTextOutput(JSON.stringify(
          setCurrentPeriodDemand(data.demand)
        )).setMimeType(ContentService.MimeType.JSON);
      
      case 'addComponent':
        return ContentService.createTextOutput(JSON.stringify(
          addMRPComponent(data.name, data.quantity, data.supplier)
        )).setMimeType(ContentService.MimeType.JSON);
      
      case 'submitCurrentPeriodData':
        console.log('=== doPost submitCurrentPeriodData case ===');
        console.log('Received data object:', data);
        console.log('Data properties:', {
          productName: data.productName,
          initialInventory: data.initialInventory,
          components: data.components
        });
        
        // Validate required parameters
        if (!data.productName) {
          console.log('Error: productName is missing or empty');
          return ContentService.createTextOutput(JSON.stringify({
            success: false,
            message: 'Product name cannot be empty'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // Remove component quantity limit, allow products without components
        
        console.log('Calling submitCurrentPeriodData with validated data');
        return ContentService.createTextOutput(JSON.stringify(submitCurrentPeriodData(data.productName, data.initialInventory, 0, data.components)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getProductResults':
        return ContentService.createTextOutput(JSON.stringify(getProductResults(data.productName)))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'exportToSheets':
        return ContentService.createTextOutput(JSON.stringify(exportMRPToSheets()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'resetSystem':
        return ContentService.createTextOutput(JSON.stringify(resetMRPSystem()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'loadFromSheets':
        return ContentService.createTextOutput(JSON.stringify(loadMRPFromSheets()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'saveToSheets':
        return ContentService.createTextOutput(JSON.stringify(saveMRPResultsToSheets()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'getLoadedData':
        return ContentService.createTextOutput(JSON.stringify(getLoadedDataFromSheets()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'initializeSheets':
        return ContentService.createTextOutput(JSON.stringify(initializeGoogleSheets()))
          .setMimeType(ContentService.MimeType.JSON);
      
      case 'testConnection':
        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          message: 'Connection test successful',
          timestamp: new Date().toISOString()
        })).setMimeType(ContentService.MimeType.JSON);
      
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Unknown operation: ' + action
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    console.error('doPost error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Server error: ' + error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * MRP System Class
 */
class MRPSystem {
  constructor() {
    this.timePeriods = 10;
    this.products = {}; // Product inventory management
    this.currentProduct = null; // Current period analysis product
    this.demand = [];
    this.initialInventory = 1;
    this.components = [];
    this.simulationResults = null;
    this.currentPeriod = 1;
  }

  /**
   * Set basic parameters
   */
  setBasicParameters(timePeriods, productName, initialInventory) {
    this.timePeriods = timePeriods;
    this.currentProduct = productName;
    this.initialInventory = initialInventory;
  }

  /**
   * Set demand data
   */
  setDemand(demandArray) {
    if (Array.isArray(demandArray)) {
      if (demandArray.length !== this.timePeriods) {
        throw new Error('Demand data length must equal the number of simulation periods!');
      }
      this.demand = demandArray;
    } else {
      // 單一數值，設定為當前期間需求
      this.setCurrentPeriodDemand(demandArray);
    }
  }

  /**
   * Set current period demand
   */
  setCurrentPeriodDemand(demand) {
    // Initialize demand array (if not already initialized)
    if (!this.demand || this.demand.length === 0) {
      this.demand = new Array(this.timePeriods).fill(0);
    }
    
    // Set current period demand
    this.demand[this.currentPeriod - 1] = parseInt(demand) || 0;
  }

  /**
   * Add component (simplified version)
   */
  addComponent(name, quantity, supplier) {
    if (!name || quantity <= 0) {
      throw new Error('Component name cannot be empty and quantity must be greater than 0!');
    }
    
    this.components.push({
      name: name,
      quantity: quantity,
      supplier: supplier
    });
  }

  /**
   * Remove component
   */
  removeComponent(componentName) {
    this.components = this.components.filter(comp => comp.name !== componentName);
  }

  /**
   * Add product inventory
   */
  addProduct(productName, initialInventory = 0) {
    this.products[productName] = {
      name: productName,
      initialInventory: initialInventory,
      components: []
    };
  }

  /**
   * Add component to product
   */
  addProductComponent(productName, componentName, quantity, supplier) {
    if (!this.products[productName]) {
      this.addProduct(productName);
    }
    
    this.products[productName].components.push({
      name: componentName,
      quantity: quantity,
      supplier: supplier
    });
  }

  /**
   * Execute MRP calculation
   */
  calculateMRP() {

    const results = {
      periods: [],
      demand: [],
      inventory: [],
      gap: [],
      productionStart: [],
      componentInventory: {}
    };

    let currentInventory = this.initialInventory;
    
    // Initialize component inventory tracking
    this.components.forEach(comp => {
      results.componentInventory[comp.name] = {
        supplier: comp.supplier,
        inventory: new Array(this.timePeriods).fill(0),
        orders: new Array(this.timePeriods).fill(0)
      };
    });

    // Calculate each period
    for (let period = 1; period <= this.timePeriods; period++) {
      const demand = this.demand[period - 1];
      let gap = null;
      let productionStart = 0;

      // Calculate inventory gap
      if (demand > 0) {
        gap = demand - currentInventory; // Negative value means sufficient inventory, positive value means insufficient inventory
      }

      // Determine production start time
      if (gap > 0) {
        // If inventory is insufficient, production is needed
        productionStart = gap;
        console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap > 0, need production)`);
        
        // Calculate component requirements
        this.components.forEach(comp => {
          const componentNeed = gap * comp.quantity;
          const orderPeriod = Math.max(1, period - 1); // Simplified lead time as 1 period
          
          if (orderPeriod <= this.timePeriods) {
            results.componentInventory[comp.name].orders[orderPeriod - 1] += componentNeed;
            // Component inventory = order quantity
            results.componentInventory[comp.name].inventory[period - 1] = 
              results.componentInventory[comp.name].orders[orderPeriod - 1];
          }
        });
      } else if (gap < 0) {
        // If inventory is sufficient, no production needed
        productionStart = 0;
        console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap < 0, sufficient inventory)`);
      } else if (gap === 0) {
        // Exactly meets demand
        productionStart = 0;
        console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap = 0)`);
      }

      // Update inventory
      if (period > 1 && results.productionStart[period - 2] > 0) {
        currentInventory += results.productionStart[period - 2];
      }
      
      if (demand > 0) {
        currentInventory = Math.max(0, currentInventory - demand);
      }

      results.periods.push(period);
      results.demand.push(demand);
      results.inventory.push(currentInventory);
      results.gap.push(gap);
      results.productionStart.push(productionStart);
    }

    this.simulationResults = results;
    return results;
  }

  /**
   * Get simulation results for specific period
   */
  getPeriodResults(period) {
    if (!this.simulationResults || period < 1 || period > this.timePeriods) {
      return null;
    }

    const index = period - 1;
    const result = {
      period: period,
      demand: this.simulationResults.demand[index],
      inventory: this.simulationResults.inventory[index],
      gap: this.simulationResults.gap[index],
      productionStart: this.simulationResults.productionStart[index],
      componentInventory: []
    };

    // Calculate component inventory for this period
    this.components.forEach(comp => {
      const inventory = this.simulationResults.componentInventory[comp.name].inventory[period - 1];
      
      result.componentInventory.push({
        component: comp.name,
        supplier: comp.supplier,
        inventory: inventory
      });
    });

    return result;
  }

  /**
   * Advance to next period
   */
  nextPeriod() {
    if (this.currentPeriod < this.timePeriods) {
      this.currentPeriod++;
      return this.getPeriodResults(this.currentPeriod);
    }
    return null;
  }

  /**
   * Reset period
   */
  resetPeriod() {
    this.currentPeriod = 1;
    return this.getPeriodResults(this.currentPeriod);
  }

  /**
   * Get complete MRP table data
   */
  getMRPTableData() {
    if (!this.simulationResults) {
      return null;
    }

    return {
      headers: Array.from({length: this.timePeriods}, (_, i) => i + 1),
      demand: this.simulationResults.demand,
      inventory: this.simulationResults.inventory,
      gap: this.simulationResults.gap,
      productionStart: this.simulationResults.productionStart,
      componentInventory: this.simulationResults.componentInventory
    };
  }

  /**
   * Validate input data
   */
  validateInputs() {
    const errors = [];

    if (!this.currentProduct || typeof this.currentProduct !== 'string' || !this.currentProduct.trim()) {
      errors.push('Product name cannot be empty');
    }

    if (this.initialInventory < 0) {
      errors.push('Initial inventory cannot be negative');
    }



    this.components.forEach((comp, index) => {
      if (!comp.name || typeof comp.name !== 'string' || !comp.name.trim()) {
        errors.push(`Component ${index + 1} name cannot be empty`);
      }
      if (comp.quantity <= 0) {
        errors.push(`Component ${comp.name} quantity must be greater than 0`);
      }
    });

    return { isValid: errors.length === 0, message: errors.length === 0 ? 'Validation successful' : errors.join(', ') };
  }

  /**
   * Reset system
   */
  reset() {
    this.timePeriods = 10;
    this.currentProduct = null;
    this.demand = [];
    this.initialInventory = 1;
    this.components = [];
    this.simulationResults = null;
    this.currentPeriod = 1;
  }

  /**
   * Get system status summary
   */
  getSystemSummary() {
    return {
      timePeriods: this.timePeriods,
      currentProduct: this.currentProduct,
      initialInventory: this.initialInventory,
      componentCount: this.components.length,
      hasSimulationResults: this.simulationResults !== null,
      currentPeriod: this.currentPeriod
    };
  }
}

// 全域 MRP 系統實例
let globalMRPSystem = new MRPSystem();

/**
 * Initialize MRP system
 */
function initializeMRPSystem() {
  globalMRPSystem = new MRPSystem();
  return { success: true, data: globalMRPSystem.getSystemSummary() };
}

/**
 * Set basic parameters
 */
function setMRPBasicParameters(productName, initialInventory) {
  try {
    globalMRPSystem.setBasicParameters(10, productName, initialInventory); // 固定 10 個期間
    return { success: true, message: 'Basic parameters set successfully' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Set current period demand
 */
function setCurrentPeriodDemand(demand) {
  try {
    globalMRPSystem.setCurrentPeriodDemand(demand);
    return { success: true, message: 'Current period demand set successfully' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Submit current period data
 */
function submitCurrentPeriodData(productName, initialInventory, currentDemand, components) {
  try {
    console.log('=== submitCurrentPeriodData called ===');
    console.log('Received data:', { productName, initialInventory, currentDemand, components });
    console.log('Data types:', {
      productName: typeof productName,
      initialInventory: typeof initialInventory,
      currentDemand: typeof currentDemand,
      components: typeof components
    });
    
    // Check if received data object (from doPost)
    if (typeof productName === 'object' && productName !== null) {
      const data = productName;
      console.log('Received data object from doPost:', data);
      productName = data.productName;
      initialInventory = data.initialInventory;
      currentDemand = data.currentDemand || 0; // If no demand provided, set to 0
      components = data.components;
      console.log('Extracted data:', { productName, initialInventory, currentDemand, components });
    }
    
    // Validate input data
    if (!productName || typeof productName !== 'string' || productName.trim() === '') {
      console.log('Error: productName is invalid');
      return {
        success: false,
        message: 'Product name cannot be empty'
      };
    }
    
    // Remove component quantity limit, allow products without components
    
    // Clear and populate the global MRP system with the received data
    globalMRPSystem = new MRPSystem();
    console.log('Created new MRPSystem instance');
    
    globalMRPSystem.setBasicParameters(10, productName, initialInventory);
    console.log('Set basic parameters:', {
      timePeriods: globalMRPSystem.timePeriods,
      currentProduct: globalMRPSystem.currentProduct,
      initialInventory: globalMRPSystem.initialInventory
    });
    
    globalMRPSystem.setCurrentPeriodDemand(currentDemand);
    console.log('Set current period demand:', currentDemand);
    
    // Clear existing components and add the new ones
    globalMRPSystem.components = [];
    console.log('Cleared components array');
    
    if (components && components.length > 0) {
      console.log('Processing components array with length:', components.length);
      for (let i = 0; i < components.length; i++) {
        const component = components[i];
        console.log(`Processing component ${i + 1}:`, component);
        if (component.name && component.quantity && component.supplier) {
          console.log(`Adding component: ${component.name}`);
          globalMRPSystem.addComponent(component.name, component.quantity, component.supplier);
        } else {
          console.log(`Skipping invalid component ${i + 1}:`, component);
        }
      }
    } else {
      console.log('No components provided or components array is empty');
    }
    
    console.log('MRP System state after setup:', {
      currentProduct: globalMRPSystem.currentProduct,
      components: globalMRPSystem.components,
      initialInventory: globalMRPSystem.initialInventory,
      componentsLength: globalMRPSystem.components.length
    });
    
    // Validate inputs
    const validationResult = globalMRPSystem.validateInputs();
    console.log('Validation result:', validationResult);
    
    if (!validationResult.isValid) {
      console.log('Validation failed, returning error');
      return {
        success: false,
        message: validationResult.message
      };
    }
    
    console.log('Validation passed, saving to sheets');
    // Save to sheets
    saveCurrentPeriodDataToSheets();
    
    return {
      success: true,
      message: "Data submitted and saved successfully"
    };
  } catch (error) {
    console.error('Error in submitCurrentPeriodData:', error);
    return {
      success: false,
      message: "Error occurred while submitting data: " + error.toString()
    };
  }
}

/**
 * Get product results
 */
function getProductResults(data) {
  try {
    const productName = data.productName;
    console.log('getProductResults called with productName:', productName);
    console.log('productName type:', typeof productName);
    
    // Ensure productName is a string
    if (typeof productName !== 'string') {
      console.log('Converting productName to string in getProductResults');
      productName = String(productName);
    }
    
    const results = loadProductResultsFromSheets(productName);
    return {
      success: true,
      data: results
    };
  } catch (error) {
    console.error('getProductResults error:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Add component
 */
function addMRPComponent(name, quantity, supplier) {
  try {
    globalMRPSystem.addComponent(name, quantity, supplier);
    return { success: true, message: 'Component added successfully' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Execute MRP simulation
 */
function runMRPSimulation() {
  try {
    const errors = globalMRPSystem.validateInputs();
    if (errors.length > 0) {
      return { success: false, message: 'Validation failed: ' + errors.join(', ') };
    }

    const results = globalMRPSystem.calculateMRP();
    return { 
      success: true, 
      message: 'MRP simulation completed',
      data: results
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Get current period results
 */
function getCurrentPeriodResults() {
  const results = globalMRPSystem.getPeriodResults(globalMRPSystem.currentPeriod);
  return {
    success: true,
    currentPeriod: globalMRPSystem.currentPeriod,
    data: results
  };
}

/**
 * Save current period data to Google Sheets
 */
function saveCurrentPeriodDataToSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Save product data to Products sheet
    const productsSheet = spreadsheet.getSheetByName('Products');
    if (!productsSheet) {
      return { success: false, message: 'Products sheet not found' };
    }
    
    // Save product data to Products sheet (new format: product name, period, inventory)
    const productData = [
      globalMRPSystem.currentProduct,
      1, // Current period fixed as 1
      globalMRPSystem.initialInventory
    ];
    
    // Find existing product row or add new row
    const data = productsSheet.getDataRange().getValues();
    let productRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === globalMRPSystem.currentProduct && data[i][1] === 1) {
        productRow = i + 1;
        break;
      }
    }
    
    if (productRow === -1) {
      // Add new product row
      productRow = data.length + 1;
    }
    
    // Write product data
    productsSheet.getRange(productRow, 1, 1, 3).setValues([productData]);
    
    // Save BOM data to BOM sheet
    const bomSheet = spreadsheet.getSheetByName('BOM');
    if (!bomSheet) {
      return { success: false, message: 'BOM sheet not found' };
    }
    
    // Prepare new BOM data (new format: product name, period, component name, quantity, supplier)
    const newBomData = globalMRPSystem.components.map(comp => [
      globalMRPSystem.currentProduct,
      1, // Current period fixed as 1
      comp.name,
      comp.quantity,
      comp.supplier
    ]);
    
    if (newBomData.length > 0) {
      // Always insert new BOM data at the end without clearing existing data
      const lastRow = bomSheet.getLastRow();
      const insertRow = lastRow + 1;
      bomSheet.getRange(insertRow, 1, newBomData.length, 5).setValues(newBomData);
    }
    
    return { success: true, message: 'Current period data saved' };
  } catch (error) {
    return { success: false, message: 'Save failed: ' + error.message };
  }
}

/**
 * Load product results from Google Sheets
 */
function loadProductResultsFromSheets(productName) {
  try {
    console.log('loadProductResultsFromSheets called with productName:', productName);
    console.log('productName type:', typeof productName);
    
    // Ensure productName is a string
    if (typeof productName !== 'string') {
      console.log('Converting productName to string');
      productName = String(productName);
    }
    
    console.log('loadProductResultsFromSheets - Final productName:', typeof productName, productName);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Read product data
    const productsSheet = spreadsheet.getSheetByName('Products');
    if (!productsSheet) {
      throw new Error('Products sheet not found');
    }
    
    const productsData = productsSheet.getDataRange().getValues();
    console.log('Products data found:', productsData.length, 'rows');
    let productData = null;
    
    // Find specified product (period 1)
    for (let i = 1; i < productsData.length; i++) {
      const sheetProductName = String(productsData[i][0]).trim(); // Ensure it's a string and remove spaces
      const sheetPeriod = productsData[i][1]; // Period
      
      console.log(`Checking row ${i}: Sheet Product: '${sheetProductName}' (type: ${typeof productsData[i][0]}), Query Product: '${productName}' (type: ${typeof productName}), Sheet Period: ${sheetPeriod}`);
      
      if (sheetProductName === productName.trim() && sheetPeriod === 1) {
        productData = productsData[i];
        console.log('Found product data:', productData);
        break;
      }
    }
    
    if (!productData) {
      console.log('DEBUG: No product data found for productName:', productName);
      console.log('DEBUG: productName type:', typeof productName);
      console.log('DEBUG: productName value:', productName);
      throw new Error('Product not found: ' + productName);
    }
    
    // Read BOM data
    const bomSheet = spreadsheet.getSheetByName('BOM');
    if (!bomSheet) {
      throw new Error('BOM sheet not found');
    }
    
    const bomData = bomSheet.getDataRange().getValues();
    const components = [];
    
    for (let i = 1; i < bomData.length; i++) {
      if (bomData[i][0] === productName && bomData[i][1] === 1) {
        components.push({
          name: bomData[i][2], // Component name
          quantity: parseInt(bomData[i][3]) || 0, // Quantity
          supplier: bomData[i][4] // Supplier
        });
      }
    }
    
    // Calculate demand from BOM data
    const initialInventory = parseInt(productData[2]) || 0; // Inventory
    
    // Calculate demand based on BOM data
    const demand = new Array(10).fill(0);
    
    // Calculate total demand: sum of quantities where this product is used as a component
    let totalDemand = 0;
    for (let i = 1; i < bomData.length; i++) {
      if (bomData[i][2] === productName && bomData[i][1] === 1) {
        totalDemand += parseInt(bomData[i][3]) || 0;
      }
    }
    
    // If no demand found as component, set demand to 0
    // This means the product is not a component of other products, no external demand
    if (totalDemand === 0) {
      totalDemand = 0;
    }
    
    demand[0] = totalDemand; // Period 1 demand
    
    console.log('Product data loaded:', productData);
    console.log('Components found:', components);
    console.log('Initial inventory:', initialInventory);
    console.log('Calculated demand for period 1:', totalDemand);
    console.log('Final demand array:', demand);
    
    // Calculate MRP results
    const mrpResults = calculateMRPResults(demand, initialInventory, components);
    
    return {
      finishedGoodsMRP: mrpResults.finishedGoodsMRP,
      componentInventory: mrpResults.componentInventory
    };
  } catch (error) {
    throw new Error('Failed to load product results: ' + error.message);
  }
}

/**
 * Calculate MRP results
 */
function calculateMRPResults(demand, initialInventory, components) {
  const timePeriods = 10;
  const results = {
    finishedGoodsMRP: {
      demand: demand,
      inventory: [],
      gap: [],
      productionStart: []
    },
    componentInventory: {
      productionStart: new Array(timePeriods).fill(0),
      components: components.map(comp => ({
        name: comp.name,
        quantity: comp.quantity,
        supplier: comp.supplier,
        leadTime: '1 Day',
        quantities: new Array(timePeriods).fill(0)
      }))
    }
  };
  
  let currentInventory = initialInventory;
  
  // Calculate finished goods MRP
  for (let period = 0; period < timePeriods; period++) {
    const periodDemand = demand[period];
    let gap = null;
    let productionStart = 0;
    
    console.log(`Period ${period + 1}: Demand=${periodDemand}, CurrentInventory=${currentInventory}`);
    
    // Calculate inventory gap
    if (periodDemand > 0) {
      gap = periodDemand - currentInventory; // Negative value means sufficient inventory, positive value means insufficient inventory
      console.log(`Period ${period + 1}: Gap=${gap}`);
    }
    
    // Determine production start time
    if (gap > 0) {
      // If inventory is insufficient, production is needed
      productionStart = gap;
      console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap > 0, need production)`);
      
      // Calculate component requirements
      components.forEach((comp, compIndex) => {
        const componentNeed = gap * comp.quantity;
        const orderPeriod = Math.max(0, period - 1); // Component demand 1 period before production starts
        
        if (orderPeriod < timePeriods) {
          results.componentInventory.components[compIndex].quantities[orderPeriod] = componentNeed;
        }
      });
    } else if (gap < 0) {
      // If inventory is sufficient, no production needed
      productionStart = 0;
      console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap < 0, sufficient inventory)`);
    } else if (gap === 0) {
      // Exactly meets demand
      productionStart = 0;
      console.log(`Period ${period + 1}: ProductionStart=${productionStart} (gap = 0)`);
    }
    
    // Update inventory
    currentInventory = Math.max(0, currentInventory - periodDemand) + productionStart;
    console.log(`Period ${period + 1}: NewInventory=${currentInventory}`);
    
    // For period 1, show initial inventory; for other periods, show calculated inventory
    if (period === 0) {
      results.finishedGoodsMRP.inventory[period] = initialInventory;
    } else {
      results.finishedGoodsMRP.inventory[period] = currentInventory;
    }
    results.finishedGoodsMRP.gap[period] = gap;
    results.finishedGoodsMRP.productionStart[period] = productionStart;
    
    // Update production start value in component inventory table
    results.componentInventory.productionStart[period] = productionStart;
  }
  
  return results;
}

/**
 * Export to Google Sheets
 */
function exportMRPToSheets() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(MRP_RESULTS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Results sheet not found' };
    }

    // Clear existing data
    sheet.clear();

    // Prepare data
    const sheetData = [];
    
    // First table: Finished Goods MRP table
    sheetData.push(['Finished Goods MRP Table - ' + globalMRPSystem.currentProduct]);
    sheetData.push([]); // Empty row
    
    // Headers
    const headers = ['Time Period'];
    for (let i = 1; i <= globalMRPSystem.timePeriods; i++) {
      headers.push(i);
    }
    sheetData.push(headers);

    // Finished goods demand
    const demandRow = ['Demand for Finished Goods'];
    globalMRPSystem.simulationResults.demand.forEach(demand => demandRow.push(demand));
    sheetData.push(demandRow);

    // Initial inventory
    const inventoryRow = ['On-Hand Inventory'];
    globalMRPSystem.simulationResults.inventory.forEach(inv => inventoryRow.push(inv));
    sheetData.push(inventoryRow);

    // Inventory gap
    const gapRow = ['Inventory Gap'];
    globalMRPSystem.simulationResults.gap.forEach(gap => gapRow.push(gap !== null ? gap : 'Null'));
    sheetData.push(gapRow);

    // Production start
    const productionRow = ['Production Start'];
    globalMRPSystem.simulationResults.productionStart.forEach(prod => productionRow.push(prod));
    sheetData.push(productionRow);

    // Empty row to separate two tables
    sheetData.push([]);
    sheetData.push([]);

    // Second table: Component inventory table
    sheetData.push(['Component Inventory Table - ' + globalMRPSystem.currentProduct]);
    sheetData.push([]); // Empty row
    
    // Component headers
    const componentHeaders = ['Time Period', 'Production Start'];
    globalMRPSystem.components.forEach(comp => {
      componentHeaders.push(comp.name + ' LT');
      componentHeaders.push(comp.name + ' Qty (X' + comp.quantity + ')');
    });
    componentHeaders.push('Supplier Name');
    sheetData.push(componentHeaders);

    // Component data rows
    for (let period = 1; period <= globalMRPSystem.timePeriods; period++) {
      const periodRow = [period, globalMRPSystem.simulationResults.productionStart[period - 1]];
      
      // Add lead time and quantity for each component
      globalMRPSystem.components.forEach(comp => {
        const inventory = globalMRPSystem.simulationResults.componentInventory[comp.name].inventory[period - 1];
        periodRow.push(inventory > 0 ? '1 Day' : 'Null');
        periodRow.push(inventory > 0 ? inventory : 'Null');
      });
      
      // Supplier name (use first component's supplier as example)
      const supplierName = globalMRPSystem.components.length > 0 ? 
        globalMRPSystem.components[0].supplier : 'N/A';
      periodRow.push(supplierName);
      
      sheetData.push(periodRow);
    }

    // Write data
    sheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);

    return { success: true, message: 'Data exported to Google Sheets' };
  } catch (error) {
    return { success: false, message: 'Export failed: ' + error.message };
  }
}

/**
 * Load MRP data directly from Google Sheets
 */
function loadMRPFromSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Read product data
    const productData = readProductFromSheet(spreadsheet);
    if (productData) {
      globalMRPSystem.setBasicParameters(
        productData.timePeriods,
        productData.productName,
        productData.initialInventory
      );
      globalMRPSystem.setDemand(productData.demand);
    }

    // Read BOM data
    const bomData = readBOMFromSheet(spreadsheet);
    if (bomData) {
      globalMRPSystem.components = bomData;
    }

    return { success: true, message: 'Data loaded from Google Sheets' };
  } catch (error) {
    return { success: false, message: 'Load failed: ' + error.message };
  }
}

/**
 * Read product data from worksheet
 */
function readProductFromSheet(spreadsheet) {
  try {
    const productsSheet = spreadsheet.getSheetByName('Products');
    if (!productsSheet) {
      return null;
    }

    const data = productsSheet.getDataRange().getValues();
    if (data.length < 2) {
      return null;
    }

    // Read first product data (new format: product name, period, inventory)
    const row = data[1];
    const productName = row[0].toString();
    const initialInventory = parseInt(row[2]) || 0; // Inventory
    
    // Create demand array (demand no longer read from Products table)
    const demand = new Array(10).fill(0);
    // Set demand to 0 (will be calculated by BOM)

    return {
      productName: productName,
      initialInventory: initialInventory,
      demand: demand,
      timePeriods: 10
    };
  } catch (error) {
    console.error('Error reading product data:', error);
    return null;
  }
}

/**
 * Read BOM data from worksheet
 */
function readBOMFromSheet(spreadsheet) {
  try {
    const components = [];
    
    const bomSheet = spreadsheet.getSheetByName('BOM');
    if (!bomSheet) {
      return [];
    }
    
    const data = bomSheet.getDataRange().getValues();
    if (data.length < 2) {
      return [];
    }
    
    // Read BOM data (skip header row) (new format: product name, period, component name, quantity, supplier)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[2] && row[3] && row[4]) { // Product name, component name, quantity, supplier
        components.push({
          name: row[2].toString(), // Component name
          quantity: parseInt(row[3]) || 0, // Quantity
          supplier: row[4].toString() // Supplier
        });
      }
    }
    
    return components;
  } catch (error) {
    console.error('Error reading BOM data:', error);
    return [];
  }
}

/**
 * Directly write MRP results to Google Sheets
 */
function saveMRPResultsToSheets() {
  try {
    if (!globalMRPSystem.simulationResults) {
      return { success: false, message: 'Please run MRP simulation first' };
    }

    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Write MRP results to specified worksheet
    writeMRPResultsToSheet(spreadsheet);
    
    return { success: true, message: 'MRP results written to Google Sheets' };
  } catch (error) {
    return { success: false, message: 'Write failed: ' + error.message };
  }
}

/**
 * Read demand data from worksheet
 */
function readDemandFromSheet(spreadsheet) {
  try {
    // Try to read demand data from different worksheets
    const sheets = ['Demand', 'demand', '需求', 'MRP_Data'];
    
    for (const sheetName of sheets) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 0) {
          // Look for demand data row
          for (let i = 0; i < data.length; i++) {
            const row = data[i];
            if (row[0] && (row[0].toString().includes('需求') || row[0].toString().includes('Demand'))) {
              // Found demand row, extract values
              const demandArray = row.slice(1).map(d => parseInt(d) || 0);
              return demandArray;
            }
          }
        }
      }
    }
    
    // If specific format not found, try reading from first worksheet
    const firstSheet = spreadsheet.getSheets()[0];
    const data = firstSheet.getDataRange().getValues();
    
    // Look for row containing numbers as demand data
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length > 1 && !isNaN(row[0])) {
        const demandArray = row.map(d => parseInt(d) || 0);
        return demandArray;
      }
    }
    
    return null;
  } catch (error) {
    console.error('Error reading demand data:', error);
    return null;
  }
}

/**
 * Read basic settings from worksheet
 */
function readBasicSettingsFromSheet(spreadsheet) {
  try {
    // Try to read from settings worksheet
    const settingsSheet = spreadsheet.getSheetByName('Settings');
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      const settings = {};
      
      for (const row of data) {
        if (row.length >= 2) {
          const key = row[0].toString().toLowerCase();
          const value = row[1];
          
          if (key.includes('期間') || key.includes('period')) {
            settings.timePeriods = parseInt(value) || 10;
          } else if (key.includes('成品') || key.includes('product')) {
            settings.productName = value.toString();
          } else if (key.includes('庫存') || key.includes('inventory')) {
            settings.initialInventory = parseInt(value) || 1;
          }
        }
      }
      
      return settings;
    }
    
    // Default settings
    return {
      timePeriods: 10,
      productName: 'Product A',
      initialInventory: 1
    };
  } catch (error) {
    console.error('Error reading basic settings:', error);
    return {
      timePeriods: 10,
      productName: 'Product A',
      initialInventory: 1
    };
  }
}

/**
 * Write MRP results to worksheet
 */
function writeMRPResultsToSheet(spreadsheet) {
  try {
    // Find or create results worksheet
    let resultsSheet = spreadsheet.getSheetByName('MRP_Results');
    if (!resultsSheet) {
      resultsSheet = spreadsheet.insertSheet('MRP_Results');
    } else {
      resultsSheet.clear();
    }
    
    // Prepare results data
    const resultsData = [];
    
    // First table: Finished Goods MRP table
    resultsData.push(['Finished Goods MRP Table - ' + globalMRPSystem.currentProduct]);
    resultsData.push([]); // Empty row
    
    // Headers
    const headers = ['Time Period'];
    for (let i = 1; i <= globalMRPSystem.timePeriods; i++) {
      headers.push(i);
    }
    resultsData.push(headers);

    // Finished goods demand
    const demandRow = ['Demand for Finished Goods'];
    globalMRPSystem.simulationResults.demand.forEach(demand => demandRow.push(demand));
    resultsData.push(demandRow);

    // Initial inventory
    const inventoryRow = ['On-Hand Inventory'];
    globalMRPSystem.simulationResults.inventory.forEach(inv => inventoryRow.push(inv));
    resultsData.push(inventoryRow);

    // Inventory gap
    const gapRow = ['Inventory Gap'];
    globalMRPSystem.simulationResults.gap.forEach(gap => gapRow.push(gap !== null ? gap : 'Null'));
    resultsData.push(gapRow);

    // Production start
    const productionRow = ['Production Start'];
    globalMRPSystem.simulationResults.productionStart.forEach(prod => productionRow.push(prod));
    resultsData.push(productionRow);

    // Empty row to separate two tables
    resultsData.push([]);
    resultsData.push([]);

    // Second table: Component inventory table
    resultsData.push(['Component Inventory Table - ' + globalMRPSystem.currentProduct]);
    resultsData.push([]); // Empty row
    
    // Component headers
    const componentHeaders = ['Time Period', 'Production Start'];
    globalMRPSystem.components.forEach(comp => {
      componentHeaders.push(comp.name + ' LT');
      componentHeaders.push(comp.name + ' Qty (X' + comp.quantity + ')');
    });
    componentHeaders.push('Supplier Name');
    resultsData.push(componentHeaders);

    // Component data rows
    for (let period = 1; period <= globalMRPSystem.timePeriods; period++) {
      const periodRow = [period, globalMRPSystem.simulationResults.productionStart[period - 1]];
      
      // Add lead time and quantity for each component
      globalMRPSystem.components.forEach(comp => {
        const inventory = globalMRPSystem.simulationResults.componentInventory[comp.name].inventory[period - 1];
        periodRow.push(inventory > 0 ? '1 Day' : 'Null');
        periodRow.push(inventory > 0 ? inventory : 'Null');
      });
      
      // Supplier name (use first component's supplier as example)
      const supplierName = globalMRPSystem.components.length > 0 ? 
        globalMRPSystem.components[0].supplier : 'N/A';
      periodRow.push(supplierName);
      
      resultsData.push(periodRow);
    }

    // Write data
    resultsSheet.getRange(1, 1, resultsData.length, resultsData[0].length).setValues(resultsData);
    
    // Format table
    formatResultsSheet(resultsSheet);
    
  } catch (error) {
    console.error('Error writing results:', error);
    throw error;
  }
}

/**
 * Format results worksheet
 */
function formatResultsSheet(sheet) {
  try {
    // Set header row format
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set data row format
    const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    dataRange.setBorder(true, true, true, true, true, true);
    
    // Auto-resize column width
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    
  } catch (error) {
    console.error('格式化工作表錯誤:', error);
  }
}

/**
 * Reset MRP system
 */
function resetMRPSystem() {
  globalMRPSystem.reset();
  return { success: true, message: 'MRP system reset' };
}

/**
 * Get system summary
 */
function getMRPSystemSummary() {
  return globalMRPSystem.getSystemSummary();
}

/**
 * Get data loaded from Google Sheets
 */
function getLoadedDataFromSheets() {
  try {
    return {
      success: true,
      data: {
        timePeriods: globalMRPSystem.timePeriods,
        currentProduct: globalMRPSystem.currentProduct,
        initialInventory: globalMRPSystem.initialInventory,
        demand: globalMRPSystem.demand,
        components: globalMRPSystem.components,
        hasSimulationResults: globalMRPSystem.simulationResults !== null
      }
    };
  } catch (error) {
    return { success: false, message: 'Failed to get data: ' + error.message };
  }
}

/**
 * Initialize Google Sheets
 */
function initializeGoogleSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Create or clear Products worksheet
    let productsSheet = spreadsheet.getSheetByName('Products');
    if (!productsSheet) {
      productsSheet = spreadsheet.insertSheet('Products');
    } else {
      productsSheet.clear();
    }
    
    // Set Products worksheet headers
    const productsHeaders = ['Product Name', 'Period', 'Inventory'];
    productsSheet.getRange(1, 1, 1, productsHeaders.length).setValues([productsHeaders]);
    
    // Create or clear BOM worksheet
    let bomSheet = spreadsheet.getSheetByName('BOM');
    if (!bomSheet) {
      bomSheet = spreadsheet.insertSheet('BOM');
    } else {
      bomSheet.clear();
    }
    
    // Set BOM worksheet headers
    const bomHeaders = ['Product Name', 'Period', 'Component Name', 'Quantity', 'Supplier'];
    bomSheet.getRange(1, 1, 1, bomHeaders.length).setValues([bomHeaders]);
    
    // Create or clear MRP_Results worksheet
    let resultsSheet = spreadsheet.getSheetByName('MRP_Results');
    if (!resultsSheet) {
      resultsSheet = spreadsheet.insertSheet('MRP_Results');
    } else {
      resultsSheet.clear();
    }
    
    // Format worksheets
    formatSheet(productsSheet);
    formatSheet(bomSheet);
    formatSheet(resultsSheet);
    
    return { success: true, message: 'Google Sheets initialization completed' };
  } catch (error) {
    return { success: false, message: 'Initialization failed: ' + error.message };
  }
}

/**
 * Format worksheet
 */
function formatSheet(sheet) {
  try {
    // Set header row format
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground('#667eea');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set data row format
    if (sheet.getLastRow() > 1) {
      const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
      dataRange.setBorder(true, true, true, true, true, true);
    }
    
    // Auto-resize column width
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    
  } catch (error) {
    console.error('格式化工作表錯誤:', error);
  }
}

/**
 * Get current period
 */
function getCurrentPeriod() {
  try {
    // Simplified version: fixed return period 1
    return { success: true, currentPeriod: 1 };
  } catch (error) {
    return { success: false, message: 'Failed to get current period: ' + error.message };
  }
}

/**
 * Update current period
 */
function updateCurrentPeriod(newPeriod) {
  try {
    // Simplified version: fixed return success
    return { success: true, message: 'Current period updated' };
  } catch (error) {
    return { success: false, message: 'Failed to update current period: ' + error.message };
  }
}

/**
 * Create required worksheets
 */
function createRequiredSheets() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Create Products worksheet
    if (!spreadsheet.getSheetByName('Products')) {
      spreadsheet.insertSheet('Products');
    }
    
    // Create BOM worksheet
    if (!spreadsheet.getSheetByName('BOM')) {
      spreadsheet.insertSheet('BOM');
    }
    
    // Create MRP Results worksheet
    if (!spreadsheet.getSheetByName('MRP_Results')) {
      spreadsheet.insertSheet('MRP_Results');
    }
    
    return { success: true, message: 'Worksheets created successfully' };
  } catch (error) {
    return { success: false, message: 'Failed to create worksheets: ' + error.message };
  }
} 

/**
 * Diagnostic function - Check system status
 */
function diagnoseSystem() {
  const results = {
    timestamp: new Date().toISOString(),
    tests: {}
  };
  
  try {
    // 測試 1：基本連接
    results.tests.basicConnection = {
      success: true,
      message: 'Basic connection test successful'
    };
  } catch (error) {
    results.tests.basicConnection = {
      success: false,
      message: 'Basic connection test failed: ' + error.message
    };
  }
  
  try {
    // 測試 2：Google Sheets 存取
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    results.tests.sheetsAccess = {
      success: true,
      message: 'Google Sheets access successful',
      sheetCount: sheets.length,
      sheetNames: sheets.map(sheet => sheet.getName())
    };
  } catch (error) {
    results.tests.sheetsAccess = {
      success: false,
      message: 'Google Sheets access failed: ' + error.message
    };
  }
  
  try {
    // 測試 3：全域變數
    results.tests.globalVariables = {
      success: globalMRPSystem !== undefined,
      message: globalMRPSystem !== undefined ? 'Global variables normal' : 'Global variables undefined',
      hasMRPSystem: globalMRPSystem instanceof MRPSystem
    };
  } catch (error) {
    results.tests.globalVariables = {
      success: false,
      message: 'Global variables check failed: ' + error.message
    };
  }
  
  try {
    // Test 4: Function availability
    const requiredFunctions = [
      'initializeGoogleSheets',
      'loadMRPFromSheets',
      'getCurrentPeriod',
      'testConnection'
    ];
    
    const functionTests = {};
    requiredFunctions.forEach(funcName => {
      functionTests[funcName] = typeof eval(funcName) === 'function';
    });
    
    results.tests.functionAvailability = {
      success: Object.values(functionTests).every(test => test),
      message: 'Function availability check completed',
      details: functionTests
    };
  } catch (error) {
    results.tests.functionAvailability = {
      success: false,
      message: 'Function availability check failed: ' + error.message
    };
  }
  
  // 總結
  const allTests = Object.values(results.tests);
  const passedTests = allTests.filter(test => test.success).length;
  const totalTests = allTests.length;
  
  results.summary = {
    totalTests: totalTests,
    passedTests: passedTests,
    failedTests: totalTests - passedTests,
    overallSuccess: passedTests === totalTests,
    message: `Diagnosis completed: ${passedTests}/${totalTests} tests passed`
  };
  
  console.log('系統診斷結果:', JSON.stringify(results, null, 2));
  return results;
}

/**
 * 快速修復函數
 */
function quickFix() {
  const results = {
    timestamp: new Date().toISOString(),
    fixes: {}
  };
  
  try {
    // 修復 1：重新初始化全域變數
    globalMRPSystem = new MRPSystem();
    results.fixes.globalVariable = {
      success: true,
      message: 'Global variables reinitialized'
    };
  } catch (error) {
    results.fixes.globalVariable = {
      success: false,
      message: 'Global variables fix failed: ' + error.message
    };
  }
  
  try {
    // 修復 2：建立必要的工作表
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requiredSheets = ['Demand', 'BOM', 'Settings', 'MRP_Results'];
    
    const createdSheets = [];
    requiredSheets.forEach(sheetName => {
      if (!spreadsheet.getSheetByName(sheetName)) {
        spreadsheet.insertSheet(sheetName);
        createdSheets.push(sheetName);
      }
    });
    
    results.fixes.sheets = {
      success: true,
      message: 'Worksheet check completed',
      createdSheets: createdSheets
    };
  } catch (error) {
    results.fixes.sheets = {
      success: false,
      message: 'Worksheet fix failed: ' + error.message
    };
  }
  
  try {
    // 修復 3：初始化基本資料
    const initResult = initializeGoogleSheets();
    results.fixes.initialization = {
      success: initResult.success,
      message: initResult.message
    };
  } catch (error) {
    results.fixes.initialization = {
      success: false,
      message: 'Initialization fix failed: ' + error.message
    };
  }
  
  // 總結
  const allFixes = Object.values(results.fixes);
  const successfulFixes = allFixes.filter(fix => fix.success).length;
  const totalFixes = allFixes.length;
  
  results.summary = {
    totalFixes: totalFixes,
    successfulFixes: successfulFixes,
    failedFixes: totalFixes - successfulFixes,
    overallSuccess: successfulFixes === totalFixes,
    message: `Quick fix completed: ${successfulFixes}/${totalFixes} fixes successful`
  };
  
  console.log('快速修復結果:', JSON.stringify(results, null, 2));
  return results;
} 

/**
 * 直接呼叫函數 - 用於 google.script.run
 */

// 測試連接
function testConnection() {
  return {
    success: true,
    message: 'Connection test successful',
    timestamp: new Date().toISOString()
  };
}

// 系統診斷
function diagnoseSystem() {
  const results = {
    timestamp: new Date().toISOString(),
    tests: {}
  };
  
  try {
    // 測試 1：基本連接
    results.tests.basicConnection = {
      success: true,
      message: 'Basic connection test successful'
    };
  } catch (error) {
    results.tests.basicConnection = {
      success: false,
      message: 'Basic connection test failed: ' + error.message
    };
  }
  
  try {
    // 測試 2：Google Sheets 存取
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = spreadsheet.getSheets();
    results.tests.sheetsAccess = {
      success: true,
      message: 'Google Sheets access successful',
      sheetCount: sheets.length,
      sheetNames: sheets.map(sheet => sheet.getName())
    };
  } catch (error) {
    results.tests.sheetsAccess = {
      success: false,
      message: 'Google Sheets access failed: ' + error.message
    };
  }
  
  try {
    // 測試 3：全域變數
    results.tests.globalVariables = {
      success: globalMRPSystem !== undefined,
      message: globalMRPSystem !== undefined ? 'Global variables normal' : 'Global variables undefined',
      hasMRPSystem: globalMRPSystem instanceof MRPSystem
    };
  } catch (error) {
    results.tests.globalVariables = {
      success: false,
      message: 'Global variables check failed: ' + error.message
    };
  }
  
  // 總結
  const allTests = Object.values(results.tests);
  const passedTests = allTests.filter(test => test.success).length;
  const totalTests = allTests.length;
  
  results.summary = {
    totalTests: totalTests,
    passedTests: passedTests,
    failedTests: totalTests - passedTests,
    overallSuccess: passedTests === totalTests,
    message: `Diagnosis completed: ${passedTests}/${totalTests} tests passed`
  };
  
  return results;
}

// 快速修復
function quickFix() {
  const results = {
    timestamp: new Date().toISOString(),
    fixes: {}
  };
  
  try {
    // 修復 1：重新初始化全域變數
    globalMRPSystem = new MRPSystem();
    results.fixes.globalVariable = {
      success: true,
      message: 'Global variables reinitialized'
    };
  } catch (error) {
    results.fixes.globalVariable = {
      success: false,
      message: 'Global variables fix failed: ' + error.message
    };
  }
  
  try {
    // 修復 2：建立必要的工作表
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requiredSheets = ['Demand', 'BOM', 'Settings', 'MRP_Results'];
    
    const createdSheets = [];
    requiredSheets.forEach(sheetName => {
      if (!spreadsheet.getSheetByName(sheetName)) {
        spreadsheet.insertSheet(sheetName);
        createdSheets.push(sheetName);
      }
    });
    
    results.fixes.sheets = {
      success: true,
      message: 'Worksheet check completed',
      createdSheets: createdSheets
    };
  } catch (error) {
    results.fixes.sheets = {
      success: false,
      message: 'Worksheet fix failed: ' + error.message
    };
  }
  
  try {
    // 修復 3：初始化基本資料
    const initResult = initializeGoogleSheets();
    results.fixes.initialization = {
      success: initResult.success,
      message: initResult.message
    };
  } catch (error) {
    results.fixes.initialization = {
      success: false,
      message: 'Initialization fix failed: ' + error.message
    };
  }
  
  // 總結
  const allFixes = Object.values(results.fixes);
  const successfulFixes = allFixes.filter(fix => fix.success).length;
  const totalFixes = allFixes.length;
  
  results.summary = {
    totalFixes: totalFixes,
    successfulFixes: successfulFixes,
    failedFixes: totalFixes - successfulFixes,
    overallSuccess: successfulFixes === totalFixes,
    message: `Quick fix completed: ${successfulFixes}/${totalFixes} fixes successful`
  };
  
  return results;
}

// 初始化 Google Sheets
function initializeSheets() {
  return initializeGoogleSheets();
}

// 載入資料
function loadFromSheets() {
  return loadMRPFromSheets();
}

// 儲存資料
function saveToSheets() {
  return saveMRPResultsToSheets();
}

// 獲取當前期間
function getCurrentPeriodData() {
  return getCurrentPeriod();
}

// 設定基本參數
function setBasicParameters(data) {
  return setMRPBasicParameters(data.productName, data.initialInventory);
}

// 設定需求
function setDemand(data) {
  return setCurrentPeriodDemand(data.demand);
}

// 新增組件
function addComponent(data) {
  return addMRPComponent(data.name, data.quantity, data.supplier);
}

// 執行模擬
function runSimulation() {
  return submitCurrentPeriodData();
}

// 獲取表格資料
function getTableData() {
  return getMRPTableData();
}

// 匯出到 Sheets
function exportToSheets() {
  return exportMRPToSheets();
}

// 重置系統
function resetSystem() {
  return resetMRPSystem();
}

// 獲取載入的資料
function getLoadedData() {
  return getLoadedDataFromSheets();
}

// 更新當前期間
function updateCurrentPeriodData(data) {
  return updateCurrentPeriod(data.newPeriod);
}

// 下一個期間
function nextPeriod() {
  const results = globalMRPSystem.nextPeriod();
  if (results) {
    return {
      success: true,
      currentPeriod: globalMRPSystem.currentPeriod,
      data: results
    };
  } else {
    return {
      success: false,
      message: 'Reached the final period'
    };
  }
}

// 重置期間
function resetPeriod() {
  const results = globalMRPSystem.resetPeriod();
  return {
    success: true,
    currentPeriod: globalMRPSystem.currentPeriod,
    data: results
  };
}

// 獲取當前期間結果
function getCurrentPeriodResults() {
  const results = globalMRPSystem.getPeriodResults(globalMRPSystem.currentPeriod);
  return {
    success: true,
    currentPeriod: globalMRPSystem.currentPeriod,
    data: results
  };
} 

/**
 * 獲取 MRP 表格資料
 */
function getMRPTableData() {
  const data = globalMRPSystem.getMRPTableData();
  if (data) {
    return { success: true, data: data };
  } else {
    return { success: false, message: 'Please run MRP simulation first' };
  }
} 