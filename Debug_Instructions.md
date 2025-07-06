# Debug Instructions for BOM and Measurements System

## 🔧 Current Issues

You're experiencing:
1. **Measurements submit successfully** to Google Sheets but show "Fail to fetch" in the web page
2. **Serial numbers request** is not getting proper response

## 📝 Steps to Debug

### Step 1: Check Google Apps Script Logs

1. **Open Google Apps Script**
   - Go to [https://script.google.com/](https://script.google.com/)
   - Open your project

2. **View Execution Logs**
   - Click "Executions" in the left sidebar
   - Look for recent executions
   - Check if there are any errors or warnings

3. **Enable Detailed Logging**
   - The updated code now includes detailed console logging
   - You should see logs like:
     ```
     doPost called with: [object]
     Parsed request data: {action: "add", table: "measurements", data: {...}}
     Creating response: {"success":true,"message":"Operation completed successfully",...}
     ```

### Step 2: Test Browser Console

1. **Open Browser Developer Tools**
   - Press `F12` or right-click → "Inspect"
   - Go to "Console" tab

2. **Submit a Measurement**
   - Enter test data and submit
   - Watch the console for detailed logs:
     ```
     API Call: {action: "add", table: "measurements", data: {...}}
     Response status: 200
     Response text: {"success":true,"message":"Operation completed successfully",...}
     Parsed result: {success: true, message: "Operation completed successfully", ...}
     ```

3. **Check for Errors**
   - Look for any red error messages
   - Pay attention to JSON parsing errors
   - Note any CORS-related errors

### Step 3: Test Serial Numbers Request

1. **Switch to Query Tab**
   - The system should automatically load serial numbers
   - Check console for logs about serial numbers request

2. **Manual Test**
   - Open browser console
   - Run this command:
     ```javascript
     apiCall('get', 'serial_numbers').then(result => console.log('Serial numbers result:', result));
     ```

### Step 4: Verify Google Apps Script Deployment

1. **Test Direct URL Access**
   - Open your Google Apps Script URL directly in browser:
     ```
     https://script.google.com/macros/s/AKfycbxuZLQ_-hOCHYHcwOo6ZpqXq6VHIu8_BaqVqHmw6_B2sThHVqZGMnhmMgYOfNfmGnAh/exec
     ```
   - You should see:
     ```json
     {
       "success": true,
       "message": "Google Apps Script is running properly",
       "timestamp": "2024-01-XX..."
     }
     ```

2. **Check Deployment Version**
   - In Google Apps Script, go to "Deploy" → "Manage deployments"
   - Ensure you're using the latest version
   - The description should show recent updates

### Step 5: Common Fixes

#### Fix 1: Redeploy Google Apps Script
```bash
1. Save your Google Apps Script code
2. Go to "Deploy" → "Manage deployments"
3. Click the edit icon (pencil)
4. Select "New version"
5. Add description: "Fixed CORS and error handling"
6. Click "Deploy"
```

#### Fix 2: Clear Browser Cache
```bash
1. Press Ctrl+Shift+Delete
2. Select "All time"
3. Check "Cookies and other site data"
4. Check "Cached images and files"
5. Click "Clear data"
6. Refresh the page
```

#### Fix 3: Check Network Tab
```bash
1. Open Developer Tools (F12)
2. Go to "Network" tab
3. Submit a measurement
4. Look for the POST request to your Google Apps Script URL
5. Check the request and response details
```

## 🎯 Expected Behavior

### Successful Measurement Submission
```javascript
// Console should show:
API Call: {action: "add", table: "measurements", data: {serial_number: "SN001", para_name: "Temperature", para_value: "25.5", created_at: "..."}}
Response status: 200
Response text: {"success":true,"message":"Operation completed successfully","timestamp":"...","data":{"id":1,"table":"measurements","row":2}}
Parsed result: {success: true, message: "Operation completed successfully", timestamp: "...", data: {...}}

// Web page should show:
"Measurements submitted successfully!"
```

### Successful Serial Numbers Loading
```javascript
// Console should show:
API Call: {action: "get", table: "serial_numbers", data: null}
Response status: 200
Response text: {"success":true,"message":"Operation completed successfully","timestamp":"...","data":["SN001","SN002"]}
Parsed result: {success: true, message: "Operation completed successfully", timestamp: "...", data: [...]}

// BOM dropdown should populate with serial numbers
```

## 🚨 Troubleshooting Specific Errors

### Error: "Fail to fetch"
**Possible Causes:**
- CORS headers not set correctly
- Google Apps Script not deployed properly
- Network connectivity issues
- Browser blocking the request

**Solutions:**
1. Redeploy Google Apps Script with latest code
2. Clear browser cache
3. Try different browser
4. Check if corporate firewall is blocking

### Error: "Invalid JSON response"
**Possible Causes:**
- Google Apps Script returning HTML error page
- Authentication required
- Script execution timeout

**Solutions:**
1. Check Google Apps Script logs for errors
2. Verify script permissions
3. Test script URL directly in browser

### Error: Network request failed
**Possible Causes:**
- Script URL incorrect
- Google Apps Script service down
- Internet connectivity issues

**Solutions:**
1. Verify script URL is correct
2. Test internet connection
3. Try again later

## 📞 Next Steps

1. **Follow the debug steps above**
2. **Check both Google Apps Script logs and browser console**
3. **Share any error messages you find**
4. **Test the direct Google Apps Script URL**

The enhanced logging should help identify exactly where the issue is occurring. 