# BOM and Measurements Management System Setup Instructions

## 📋 System Overview

This system contains two main files:
- `index.html` - Static web page for data input and management
- `gas-script.js` - Google Apps Script code for handling data operations

## 🆕 New Features

### Enhanced User Interface
- **Tab-based interface** with two main sections:
  - **Data Input**: For entering measurements and BOM data
  - **Query & Update**: For viewing, editing, and managing data

### Improved Data Flow
1. **Measurements First**: Start by entering measurements for serial numbers
2. **Multiple Parameters**: Add multiple measurements for each serial number
3. **Flexible Values**: Parameter values can be any text or number
4. **Smart BOM Creation**: BOM entries use existing serial numbers from measurements

### Advanced Query Features
- **Tree Structure Display**: Visual BOM hierarchy
- **Interactive Navigation**: Click parts to expand/collapse
- **Measurement Details**: Click serial numbers to view measurements
- **Edit & Delete**: Update or remove measurements directly

## 🚀 Setup Steps

### Step 1: Setup Google Apps Script

1. **Open Google Apps Script**
   - Go to [https://script.google.com/](https://script.google.com/)
   - Sign in with your Google account

2. **Create New Project**
   - Click "New Project"
   - Name it "BOM-Measurements-API"

3. **Paste Code**
   - Delete the default `myFunction()` code
   - Copy all code from `gas-script.js` file and paste it
   - Confirm the `SPREADSHEET_ID` is set to your Google Sheets ID:
     ```javascript
     const SPREADSHEET_ID = '13RKGU7mhy-NaFkokHT87tSnPXp5ae0WC0tJjL0wOLKw';
     ```

4. **Save Project**
   - Press `Ctrl+S` or click the save icon

### Step 2: Initialize Sheets (Optional)

1. **Run Initialization Function**
   - In Google Apps Script editor
   - Select function: `initializeSheets`
   - Click "Run" button
   - Allow all permissions if prompted

2. **Verify Sheets**
   - Check your Google Sheets
   - You should see two new sheets: `bom` and `measurements`
   - Each sheet has appropriate headers

### Step 3: Deploy as Web App

1. **Start Deployment**
   - Click "Deploy" → "New deployment"

2. **Select Type**
   - Choose "Web app"

3. **Configure Deployment**
   - **Description**: Enter "BOM Measurements API v2"
   - **Execute as**: Select "Me"
   - **Access**: Select "Anyone"

4. **Deploy**
   - Click "Deploy"
   - Allow all permissions if prompted

5. **Get Web App URL**
   - Copy the provided Web App URL
   - Format similar to: `https://script.google.com/macros/s/AKfycbxxx.../exec`

### Step 4: Configure HTML Page

1. **Edit index.html**
   - Open `index.html` file
   - Find this line:
     ```javascript
     const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbw-ab-nSaFu2FuyYsxV75RzCrcyVg50jqJMLBJz2d8ygJtdfU8mUSxgQcOtfsDRUJ68/exec';
     ```
   - Replace with your new Web App URL from Step 3

2. **Save File**
   - Save the `index.html` file

## 🎯 Usage Guide

### Opening the Page
- Double-click `index.html` file
- Or open the file in your browser

### Data Input Tab

#### 1. Measurements Input
1. Enter a **Serial Number**
2. Add **Parameter Names** and **Values**
3. Click "Add Measurement" to add more parameters
4. Click "Submit Measurements" to save

#### 2. BOM Input
1. Enter **Part Number**
2. Select **Serial Number** from dropdown (populated from measurements)
3. Optionally select **Assembly Serial Number** for hierarchical relationships
4. Click "Submit BOM" to save

### Query & Update Tab

#### 1. View BOM Tree
1. Click "Refresh BOM Tree" to load structure
2. Click on part numbers to expand/collapse
3. Click on serial numbers to view measurements

#### 2. Manage Measurements
1. Click any serial number in the tree
2. View all measurements for that serial number
3. Use "Edit" button to modify measurements
4. Use "Delete" button to remove measurements

## 📊 Data Structure

### BOM Table Structure
| Field | Type | Description |
|-------|------|-------------|
| id | Number | Auto-generated unique ID |
| part_no | Text | Part number |
| serial_number | Text | Serial number |
| assembly_serial_number | Text | Assembly serial number (optional) |
| created_at | DateTime | Creation timestamp |

### Measurements Table Structure
| Field | Type | Description |
|-------|------|-------------|
| id | Number | Auto-generated unique ID |
| serial_number | Text | Serial number |
| para_name | Text | Parameter name |
| para_value | Text | Parameter value (any text/number) |
| created_at | DateTime | Creation timestamp |

## 🔧 Testing

### Test Google Apps Script
1. In Google Apps Script editor
2. Select function: `testScript`
3. Click "Run"
4. Check execution log for no errors
5. Verify test data appears in Google Sheets

### Test Web Page
1. Open `index.html`
2. Try submitting measurements
3. Try creating BOM entries
4. Switch to Query tab and test tree navigation
5. Test edit/delete functionality

## 🚨 Troubleshooting

### Common Issues

1. **"Please set Google Apps Script URL!" Error**
   - Check if `SCRIPT_URL` is correctly replaced
   - Verify URL format is correct

2. **"Submit Failed" Error**
   - Check network connection
   - Verify Google Apps Script is deployed correctly
   - Check browser developer tools for error messages

3. **Data Not Appearing in Google Sheets**
   - Confirm Google Sheets ID is correct
   - Check sheet names are `bom` and `measurements`
   - Verify Google Apps Script has permissions

4. **Authorization Issues**
   - Redeploy Google Apps Script
   - Ensure access permissions are set to "Anyone"

### Debugging Steps
1. **Confirm Google Apps Script URL**
   - Visit the Web App URL directly
   - Should see JSON response like:
     ```json
     {
       "success": true,
       "message": "Google Apps Script is running properly",
       "timestamp": "2024-01-XX..."
     }
     ```

2. **Check Browser Developer Tools**
   - Press F12 to open developer tools
   - Look at Console and Network tabs for errors
   - Verify request status codes are 200, not 405

3. **Check Google Apps Script Execution Log**
   - In Google Apps Script editor, view execution logs
   - Look for any error messages

## 🎯 Success Indicators

When setup is successful, you should see:
- ✅ No 405 errors
- ✅ Form submissions show success messages
- ✅ Data appears correctly in Google Sheets
- ✅ BOM tree displays properly
- ✅ Edit/delete functions work
- ✅ Browser developer tools show 200 status codes

## 📝 Important Notes

1. **Data Input Flow**: Always enter measurements first, then create BOM relationships
2. **Serial Number Validation**: BOM entries can only use existing serial numbers from measurements
3. **Parameter Values**: Can be any text or number (not restricted to numbers)
4. **Tree Structure**: BOM creates hierarchical relationships based on assembly serial numbers
5. **Real-time Updates**: Changes are reflected immediately in Google Sheets

## 🔄 Updating Deployment

If you modify the Google Apps Script code:
1. Save the code
2. Click "Deploy" → "Manage deployments"
3. Click the "Edit" icon
4. Select "New version"
5. Update description
6. Click "Deploy"

## 📞 Support

If you encounter issues:
1. Check this troubleshooting section
2. Verify all setup steps are completed correctly
3. Check browser developer tools for detailed error messages
4. Review Google Apps Script execution logs

---

**Important Reminder:** You must redeploy Google Apps Script after any code changes for the modifications to take effect. 