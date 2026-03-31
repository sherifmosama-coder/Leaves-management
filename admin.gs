/* ============================================================================
   HOLIDAY MANAGEMENT SYSTEM - ADMIN FUNCTIONS
   System reset, employee management, and admin operations
   ============================================================================ */
function canModifyAdmin(actingAdminEmail, targetAdminEmail) {
  try {
    console.log('=== CHECK MODIFY PERMISSION ===');
    console.log('Acting admin:', actingAdminEmail);
    console.log('Target admin:', targetAdminEmail);
    
    // Get both admin details
    const actingAdmin = getAdminByEmail(actingAdminEmail);
    const targetAdmin = getAdminByEmail(targetAdminEmail);
    
    if (!actingAdmin) {
      return {
        canModify: false,
        reason: 'Acting administrator not found'
      };
    }
    
    if (!targetAdmin) {
      return {
        canModify: false,
        reason: 'Target administrator not found'
      };
    }
    
    // Get target admin's creator from Admins sheet
    const adminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins');
    if (!adminsSheet) {
      return {
        canModify: false,
        reason: 'Admins sheet not found'
      };
    }
    
    const data = adminsSheet.getDataRange().getValues();
    let targetCreator = '';
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][1] ? data[i][1].toString().toLowerCase() : '';
      if (rowEmail === targetAdminEmail.toLowerCase()) {
        targetCreator = data[i][6] || ''; // Column G - Created By
        break;
      }
    }
    
    console.log('Target admin creator:', targetCreator);
    console.log('Target admin has full permission:', targetAdmin.fullPermission);
    
    // Rule: If target has full permission, only their creator can modify them
    if (targetAdmin.fullPermission) {
      const actingAdminName = actingAdmin.name || actingAdminEmail;
      
      // Check if acting admin is the creator
      if (targetCreator && targetCreator.toLowerCase() === actingAdminName.toLowerCase()) {
        console.log('✅ Acting admin is the creator - CAN MODIFY');
        return {
          canModify: true,
          reason: ''
        };
      } else {
        console.log('❌ Not the creator - CANNOT MODIFY');
        return {
          canModify: false,
          reason: `This administrator has full permissions and can only be modified by ${targetCreator || 'their creator'}`
        };
      }
    }
    
    // If target doesn't have full permission, any active admin can modify
    console.log('✅ Target has no full permission - CAN MODIFY');
    return {
      canModify: true,
      reason: ''
    };
    
  } catch (error) {
    console.error('Error checking modify permission:', error);
    return {
      canModify: false,
      reason: 'Error checking permissions: ' + error.toString()
    };
  }
}

/**
 * Check if admin can change full permission setting
 * Returns: { canChange: boolean, reason: string }
 */
function canChangeFullPermission(actingAdminEmail, targetAdminEmail) {
  try {
    console.log('=== CHECK FULL PERMISSION CHANGE ===');
    
    // Get target admin details
    const targetAdmin = getAdminByEmail(targetAdminEmail);
    
    if (!targetAdmin) {
      return {
        canChange: false,
        reason: 'Target administrator not found'
      };
    }
    
    // If target currently has full permission, only creator can change it
    if (targetAdmin.fullPermission) {
      const permissionCheck = canModifyAdmin(actingAdminEmail, targetAdminEmail);
      
      if (!permissionCheck.canModify) {
        return {
          canChange: false,
          reason: 'Only the creator can remove full permissions from this administrator'
        };
      }
    }
    
    // If target doesn't have full permission, any admin can grant it
    return {
      canChange: true,
      reason: ''
    };
    
  } catch (error) {
    console.error('Error checking full permission change:', error);
    return {
      canChange: false,
      reason: 'Error checking permissions: ' + error.toString()
    };
  }
}

/* ============================================================================
   NEW SYSTEM RESET LOGIC (PHASE 1)
   ============================================================================ */

/* ============================================================================
   NEW SYSTEM RESET LOGIC (PHASE 1) - REVISED
   ============================================================================ */

/**
 * PHASE 1: Verify Admin Password & Permissions
 * REVISED: Matches structure of 'Admins.csv'
 * Col B (Index 1) = Email
 * Col C (Index 2) = Password
 * Col D (Index 3) = Permission (Must be "Yes")
 */
function newSR_verifyAdminPassword(password, adminEmail) {
  try {
    // 1. Strict Input Validation
    if (!adminEmail || typeof adminEmail !== 'string') {
      return { success: false, error: 'System Error: Admin identity is missing.' };
    }
    
    const targetEmail = adminEmail.trim().toLowerCase();
    const targetPassword = String(password).trim(); 

    // 2. Access Data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const adminsSheet = ss.getSheetByName('Admins');
    if (!adminsSheet) return { success: false, error: 'Critical: Admins sheet not found.' };

    const data = adminsSheet.getDataRange().getValues();
    
    // 3. Locate the Admin
    let adminFound = false;
    
    // Start loop at i=1 to skip headers
    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][1]).toLowerCase().trim(); // Column B
      
      if (rowEmail === targetEmail) {
        adminFound = true;
        const storedPassword = String(data[i][2]).trim(); // Column C
        const permissionValue = String(data[i][3]).trim().toLowerCase(); // Column D
        
        // 4. Verify Credentials
        if (storedPassword === targetPassword) {
           // 5. Verify Permissions (Looking for "yes")
           if (permissionValue === 'yes') {
             return { success: true };
           } else {
             return { success: false, error: 'Access Denied: You do not have "Yes" in Full Permission (Column D).' };
           }
        } else {
          return { success: false, error: 'Incorrect Password.' };
        }
      }
    }
    
    if (!adminFound) {
      return { success: false, error: `Admin account not found for: ${targetEmail}` };
    }

  } catch (error) {
    console.error('newSR_verifyAdminPassword Error:', error);
    return { success: false, error: 'Server Error: ' + error.message };
  }
}

/* ============================================================================
   NEW SYSTEM RESET LOGIC (PHASE 2 UPDATE)
   ============================================================================ */

/**
 * PHASE 2: Fetch Data for Review (UPDATED)
 * Now includes Global Default Balances from Config
 */
function newSR_getPreResetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    
    // --- 1. GET CONFIGURATION & DEFAULTS ---
    const configSheet = ss.getSheetByName('Config');
    // Dates
    const configEndDateValue = configSheet.getRange('B4').getValue(); 
    const systemEndDateStr = Utilities.formatDate(new Date(configEndDateValue), tz, 'yyyy-MM-dd');
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    
    // Default Balances (Based on your verification)
    const defaultAnnual = configSheet.getRange('B9').getValue() || 0;
    const defaultSick = configSheet.getRange('B10').getValue() || 0;
    const defaultEmergency = configSheet.getRange('B11').getValue() || 0;
    
    // --- 2. GET ACTIVE EMPLOYEES ---
    const empSheet = ss.getSheetByName('Employees');
    const empData = empSheet.getDataRange().getValues();
    const employees = []; 

    for (let i = 1; i < empData.length; i++) {
      const row = empData[i];
      const status = String(row[11]); // Col L
      
      if (status === 'Active') {
        employees.push({
          id: String(row[0]), 
          name: row[1],       
          annualBal: 0,       
          sickBal: 0,
          emergencyBal: 0,
          outstandingLeaves: 0,
          outstandingDetails: [] 
        });
      }
    }
    
    // --- 3. GET CURRENT BALANCES (Calendar Logic) ---
    const balSheet = ss.getSheetByName('Balance - Calendar');
    if (balSheet) {
      const balData = balSheet.getDataRange().getValues();
      const headerRow = balData[0];
      
      let targetRowIndex = -1;
      const formatDateSafe = (d) => (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';

      for (let r = 4; r < balData.length; r++) {
        if (formatDateSafe(balData[r][1]) === todayStr) { targetRowIndex = r; break; }
      }
      if (targetRowIndex === -1) {
        for (let r = 4; r < balData.length; r++) {
          if (formatDateSafe(balData[r][1]) === systemEndDateStr) { targetRowIndex = r; break; }
        }
      }
      
      if (targetRowIndex !== -1) {
        const balanceRow = balData[targetRowIndex];
        employees.forEach(emp => {
          const colIndex = headerRow.findIndex(h => String(h) === emp.id);
          if (colIndex > -1) {
            emp.annualBal = balanceRow[colIndex];
            emp.sickBal = balanceRow[colIndex + 1];
            emp.emergencyBal = balanceRow[colIndex + 2];
          }
        });
      }
    }
    
    // --- 4. COUNT OUTSTANDING LEAVES ---
    const leaveSheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    leaveSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        const empId = String(data[r][1]);
        const leaveDateVal = data[r][3]; 
        const status = String(data[r][8]).trim(); 
        
        if (status === 'Pending' || status === '') {
          const empObj = employees.find(e => e.id === empId);
          if (empObj) {
            empObj.outstandingLeaves++;
            const niceDate = (leaveDateVal instanceof Date) ? Utilities.formatDate(leaveDateVal, tz, 'dd MMM') : String(leaveDateVal);
            const shortType = sheetName.replace(' leaves', '');
            empObj.outstandingDetails.push(`${shortType}: ${niceDate}`);
          }
        }
      }
    });

    return { 
      success: true, 
      employees: employees,
      defaults: {
        annual: defaultAnnual,
        sick: defaultSick,
        emergency: defaultEmergency
      }
    };
    
  } catch (error) {
    console.error('newSR_getPreResetData Error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * PHASE 2: Finalize Reset (DESTRUCTIVE ACTION)
 * 1. Updates Config Dates
 * 2. Deletes Inactive Employees & Updates Balances/Dates for Active ones
 * 3. Clears Leave History outside new range
 */
function newSR_finalizeReset(newStartStr, newEndStr, newBalancesMap) {
  const lock = LockService.getScriptLock();
  try {
    // Prevent concurrent resets (30 sec timeout)
    lock.waitLock(30000); 
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    
    // Parse Dates
    const newStartDate = new Date(newStartStr);
    const newEndDate = new Date(newEndStr);
    
    // --- 1. UPDATE CONFIG ---
    const configSheet = ss.getSheetByName('Config');
    configSheet.getRange('B3').setValue(newStartDate);
    configSheet.getRange('B4').setValue(newEndDate);
    
    // --- 2. UPDATE EMPLOYEES SHEET ---
    const empSheet = ss.getSheetByName('Employees');
    const empRange = empSheet.getDataRange();
    const empData = empRange.getValues();
    const newEmpData = [];
    
    // Add Header Row (Preserve existing headers)
    newEmpData.push(empData[0]);
    
    // Process Employee Rows
    for (let i = 1; i < empData.length; i++) {
      const row = empData[i];
      const empId = String(row[0]);
      const status = String(row[11]); // Col L (Index 11) = Active/Inactive
      
      // LOGIC: Only keep 'Active' employees
      if (status === 'Active') {
        // Update Validity Dates (Col E=4, Col F=5)
        row[4] = newStartDate;
        row[5] = newEndDate;
        
        // Update Balances if provided in map
        if (newBalancesMap && newBalancesMap[empId]) {
          // Annual (Col D = Index 3)
          row[3] = newBalancesMap[empId].annual;
          // Emergency (Col M = Index 12)
          row[12] = newBalancesMap[empId].emergency;
          // Sick (Col N = Index 13)
          row[13] = newBalancesMap[empId].sick;
        }
        
        newEmpData.push(row);
      }
    }
    
    // Write back to Employees Sheet (Overwrite entire sheet with filtered data)
    empSheet.clearContents();
    if (newEmpData.length > 0) {
      empSheet.getRange(1, 1, newEmpData.length, newEmpData[0].length).setValues(newEmpData);
    }
    
    // --- 3. CLEAN UP LEAVE SHEETS ---
    const leaveSheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    
    leaveSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const retainedRows = [headers];
      
      // Scan rows
      for (let r = 1; r < data.length; r++) {
        const rowDateStr = data[r][3]; // Col D = Date
        const rowDate = new Date(rowDateStr);
        
        // LOGIC: Keep row ONLY if date is INSIDE the New System Dates
        // (e.g. Future leaves already booked for the new year)
        if (rowDate >= newStartDate && rowDate <= newEndDate) {
          retainedRows.push(data[r]);
        }
      }
      
      // Write back
      sheet.clearContents();
      if (retainedRows.length > 0) {
        sheet.getRange(1, 1, retainedRows.length, retainedRows[0].length).setValues(retainedRows);
      }
    });
    
    return { success: true };
    
  } catch (error) {
    console.error('newSR_finalizeReset Error:', error);
    return { success: false, error: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * PHASE 1: Archive System
 * REVISED: Logs exactly to the 9 columns of "System versions"
 */
function newSR_archiveSystem(newStartStr, newEndStr, adminEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const configSheet = ss.getSheetByName('Config');
    
    // 1. Get Old Dates (Current System Period)
    // Assuming B3 = Start Date, B4 = End Date in Config
    const oldStartDate = configSheet.getRange('B3').getValue();
    const oldEndDate = configSheet.getRange('B4').getValue();
    
    // Format for File Name and Log
    const fmtOldStart = Utilities.formatDate(new Date(oldStartDate), tz, 'dd-MM-yyyy');
    const fmtOldEnd = Utilities.formatDate(new Date(oldEndDate), tz, 'dd-MM-yyyy');
    
    // 2. Create Archived Copy
    const archiveName = `Holiday System Archive - ${fmtOldStart} to ${fmtOldEnd}`;
    const archiveFile = ss.copy(archiveName);
    const archiveUrl = archiveFile.getUrl();
    
    // 3. Log to 'System versions' Sheet
    const versionSheet = ss.getSheetByName('System versions');
    if (versionSheet) {
      const lastRow = versionSheet.getLastRow();
      let newId = 1; // Default if sheet is empty
      
      // Calculate Next ID (Column A)
      if (lastRow > 1) { // Assuming row 1 is headers
        const lastId = versionSheet.getRange(lastRow, 1).getValue();
        if (!isNaN(lastId)) {
          newId = lastId + 1;
        }
      }
      
      // Prepare the Row Data (9 Columns)
      // Headers: #, link, date from, date to, Reset Date, Reset By, New Start, New End, Type
      const rowData = [
        newId,                                                  // A: #
        archiveUrl,                                             // B: Link
        Utilities.formatDate(new Date(oldStartDate), tz, 'MM/dd/yyyy'), // C: Date From
        Utilities.formatDate(new Date(oldEndDate), tz, 'MM/dd/yyyy'),   // D: Date To
        Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy HH:mm:ss'),    // E: Reset Date (Timestamp)
        adminEmail,                                             // F: Reset By
        newStartStr,                                            // G: New Start Date (Input)
        newEndStr,                                              // H: New End Date (Input)
        "Immediate"                                             // I: Reset Type
      ];
      
      versionSheet.appendRow(rowData);
    }
    
    return {
      success: true,
      archiveUrl: archiveUrl,
      archiveName: archiveName
    };
    
  } catch (error) {
    console.error('newSR_archiveSystem Error:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * FETCH SYSTEM VERSIONS LOG
 * Reads the 'System versions' sheet to display history in Settings
 */
function getSystemVersionsLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('System versions');
    
    if (!sheet) {
      return { success: false, error: "'System versions' sheet not found." };
    }
    
    const data = sheet.getDataRange().getValues();
    const versions = [];
    
    // Structure based on CSV: 
    // A=#, B=Link, C=From, D=To, E=ResetDate, F=ResetBy
    
    // Skip header (i=1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Basic validation: must have a link
      if (row[1]) { 
        versions.push({
          id: row[0],
          url: row[1],
          dateFrom: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'dd MMM yyyy') : row[2],
          dateTo: row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'dd MMM yyyy') : row[3],
          resetDate: row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') : row[4],
          resetBy: row[5]
        });
      }
    }
    
    // Sort by ID descending (newest first)
    versions.sort((a, b) => b.id - a.id);
    
    return { success: true, versions: versions };
    
  } catch (error) {
    console.error('getSystemVersionsLog Error:', error);
    return { success: false, error: error.toString() };
  }
}


/* ============================================================================
   SYSTEM RESET HELPER FUNCTIONS
   ============================================================================ */

/**
 * Create archived copy of current spreadsheet
 */
/**
 * Create archived copy of current spreadsheet - SIMPLIFIED
 */
function createArchivedCopy(ss, currentDates, admin) {
  try {
    // Use simple string dates (no formatting needed)
    const startDateStr = currentDates.startDate || 'unknown';
    const endDateStr = currentDates.endDate || 'unknown';
    
    const archiveName = `Holiday System Archive (${startDateStr} to ${endDateStr})`;
    
    // Create a simple copy
    const archive = ss.copy(archiveName);
    
    return {
      success: true,
      archiveUrl: archive.getUrl(),
      archiveId: archive.getId()
    };
    
  } catch (error) {
    console.error('Error creating archived copy:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Log system version in System versions sheet
 */
function logSystemVersion(archiveUrl, oldStartDate, oldEndDate, newStartDate, newEndDate, admin, resetType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const versionsSheet = ss.getSheetByName('System versions');
    
    if (!versionsSheet) {
      return {
        success: false,
        error: 'System versions sheet not found'
      };
    }
    
    const lastRow = versionsSheet.getLastRow();
    const serialNumber = lastRow > 1 ? versionsSheet.getRange(lastRow, 1).getValue() + 1 : 1;
    
    const newRow = [
      serialNumber,
      archiveUrl,
      oldStartDate,
      oldEndDate,
      new Date(),
      admin.email,
      newStartDate,
      newEndDate,
      resetType
    ];
    
    versionsSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    return { success: true };
    
  } catch (error) {
    console.error('Error logging system version:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}


/**
 * Clear all leave requests
 */
function clearAllLeaveRequests() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const leaveSheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    
    for (const sheetName of leaveSheets) {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          // Clear all data except headers
          sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
        }
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error clearing leave requests:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Update reset configuration
 */
function updateResetConfig(admin, resetDate) {
  try {
    const configSheet = getSheet('Config');
    if (configSheet) {
      configSheet.getRange('b21').setValue(resetDate); // Last Reset Date
    }
  } catch (error) {
    console.error('Error updating reset config:', error);
  }
}

/**
 * Generate OTP code
 */
function generateOTPCode() {
  return Math.floor(100000 + Math.random() * 900000).toString();
}

/**
 * Clean up OTP data
 */
function cleanupOTPData(otpToken) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty(`otp_${otpToken}`);
    properties.deleteProperty(`otp_exp_${otpToken}`);
    properties.deleteProperty(`otp_email_${otpToken}`);
    properties.deleteProperty(`otp_verified_${otpToken}`);
  } catch (error) {
    console.error('Error cleaning up OTP data:', error);
  }
}

/* ============================================================================
   NOTIFICATION PREFERENCES MANAGEMENT
   ============================================================================ */

/**
 * Get notification preferences from Config sheet B8
 */
function getNotificationPreferences() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: true,
        data: {
          notifyEnabled: false,
          reminderDays: 7
        }
      };
    }
    
    // Get value from B8 - Alert before system end (days)
    const daysValue = configSheet.getRange('B8').getValue();
    
    // If B8 is blank or 0, notifications are disabled
    const notifyEnabled = daysValue && daysValue > 0;
    const reminderDays = notifyEnabled ? parseInt(daysValue) : 7;
    
    return {
      success: true,
      data: {
        notifyEnabled: notifyEnabled,
        reminderDays: reminderDays
      }
    };
    
  } catch (error) {
    console.error('Error getting notification preferences:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Save notification preferences to Config sheet B8
 */
function saveNotificationPreferences(preferences) {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    // If notifications enabled, save the days value, otherwise clear the cell
    if (preferences.notifyEnabled) {
      configSheet.getRange('B8').setValue(preferences.reminderDays);
    } else {
      configSheet.getRange('B8').clearContent();
    }
    
    // Log action
    // logAdminAction('Notification Preferences Updated');
    
    return { success: true };
    
  } catch (error) {
    console.error('Error saving notification preferences:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   ADMINS DATA MANAGEMENT
   ============================================================================ */

function getAdminsData() {
  try {
    console.log('=== GET ADMINS DATA START ===');
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      console.error('❌ Admins sheet not found');
      return "ERROR:Admins sheet not found";
    }
    
    console.log('✅ Admins sheet found');
    
    const data = adminsSheet.getDataRange().getValues();
    console.log('📊 Total rows in sheet:', data.length);
    
    if (data.length <= 1) {
      console.log('⚠️ No admin data found (empty or header only)');
      return "EMPTY:No administrators found";
    }
    
    const adminsDataArray = [];
    
    // Process each admin (skip header row)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      console.log(`🔍 Processing admin row ${i}:`, row);
      
      // Use standardized column mapping with null safety
      const name = row[ADMIN_COLUMNS.NAME] ? row[ADMIN_COLUMNS.NAME].toString().trim() : 'Unknown';
      const email = row[ADMIN_COLUMNS.EMAIL] ? row[ADMIN_COLUMNS.EMAIL].toString().trim() : '';
      const fullPermission = row[ADMIN_COLUMNS.FULL_PERMISSION] === true || row[ADMIN_COLUMNS.FULL_PERMISSION] === 'Yes' || row[ADMIN_COLUMNS.FULL_PERMISSION] === 'TRUE';
      const receiveRequests = row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === true || row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'Yes' || row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'TRUE';
      
      // Handle dates safely
      let dateCreated = 'Unknown';
      if (row[ADMIN_COLUMNS.DATE_CREATED]) {
        try {
          dateCreated = new Date(row[ADMIN_COLUMNS.DATE_CREATED]).toISOString();
        } catch (e) {
          dateCreated = row[ADMIN_COLUMNS.DATE_CREATED].toString();
        }
      }
      
      const createdBy = row[ADMIN_COLUMNS.CREATED_BY] ? row[ADMIN_COLUMNS.CREATED_BY].toString().trim() : 'System';
      
      let lastLogin = '';
      if (row[ADMIN_COLUMNS.LAST_LOGIN]) {
        try {
          lastLogin = new Date(row[ADMIN_COLUMNS.LAST_LOGIN]).toISOString();
        } catch (e) {
          lastLogin = row[ADMIN_COLUMNS.LAST_LOGIN].toString();
        }
      }
      
      const status = row[ADMIN_COLUMNS.STATUS] ? row[ADMIN_COLUMNS.STATUS].toString().trim().toLowerCase() : 'active';
      const actualStatus = status === '' || status === 'active' ? 'active' : status;
      
      let statusChangedOn = '';
      if (row[ADMIN_COLUMNS.STATUS_CHANGED_ON]) {
        try {
          statusChangedOn = new Date(row[ADMIN_COLUMNS.STATUS_CHANGED_ON]).toISOString();
        } catch (e) {
          statusChangedOn = row[ADMIN_COLUMNS.STATUS_CHANGED_ON].toString();
        }
      }
      
      console.log(`   📋 Processed admin: ${name} (${email}) - Status: ${actualStatus}`);
      
      // Create admin data string - all values as strings separated by |
      const adminString = [
        name,                           // 0: Name
        email,                          // 1: Email  
        fullPermission ? 'true' : 'false', // 2: Full Permission
        receiveRequests ? 'true' : 'false', // 3: Receive Requests
        dateCreated,                    // 4: Date Created
        createdBy,                      // 5: Created By
        lastLogin,                      // 6: Last Login
        actualStatus,                   // 7: Status
        statusChangedOn,                // 8: Status Changed On
        (i + 1).toString()              // 9: Row Index for updates
      ].join('|');
      
      adminsDataArray.push(adminString);
      console.log(`   ✅ Admin string created: ${adminString.substring(0, 100)}...`);
    }
    
    console.log('📋 Total admins processed:', adminsDataArray.length);
    
    if (adminsDataArray.length === 0) {
      console.log('⚠️ No valid admin data found');
      return "EMPTY:No valid administrators found";
    }
    
    // Return as SUCCESS with semicolon-separated admin strings
    const result = "SUCCESS:" + adminsDataArray.join(';');
    console.log('✅ Returning result length:', result.length);
    console.log('=== GET ADMINS DATA END ===');
    
    return result;
    
  } catch (error) {
    console.error('❌ ERROR in getAdminsData:', error);
    return "ERROR:Error retrieving admins data - " + error.toString();
  }
}

/**
 * Update admin field - CORRECTED to accept acting admin email
 */
function updateAdminField(updateData, actingAdminEmail) {
  try {
    console.log('✏️ Updating admin field:', updateData);
    console.log('👤 Updated by:', actingAdminEmail);
    
    if (!actingAdminEmail) {
      return "ERROR:Acting admin email is required";
    }
    
    // SPECIAL CHECK: If changing full permission
    if (updateData.field === 'fullPermission') {
      const permissionCheck = canChangeFullPermission(actingAdminEmail, updateData.adminEmail);
      
      if (!permissionCheck.canChange) {
        console.error('❌ Permission denied:', permissionCheck.reason);
        return "ERROR:" + permissionCheck.reason;
      }
    } else {
      // For other fields, check general modify permission
      const permissionCheck = canModifyAdmin(actingAdminEmail, updateData.adminEmail);
      
      if (!permissionCheck.canModify) {
        console.error('❌ Permission denied:', permissionCheck.reason);
        return "ERROR:" + permissionCheck.reason;
      }
    }
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      return "ERROR:Admins sheet not found";
    }
    
    // Find admin by email
    const data = adminsSheet.getDataRange().getValues();
    let targetRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][1] ? data[i][1].toString().toLowerCase() : '';
      if (rowEmail === updateData.adminEmail.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      return "ERROR:Administrator not found";
    }
    
    // Map field names to column indices
    const ADMIN_COLUMNS = {
      NAME: 0,
      EMAIL: 1,
      PASSWORD: 2,
      FULL_PERMISSION: 3,
      RECEIVE_REQUESTS: 4,
      DATE_CREATED: 5,
      CREATED_BY: 6,
      LAST_LOGIN: 7,
      STATUS: 8,
      STATUS_CHANGED_ON: 9
    };
    
    const fieldMapping = {
      'name': ADMIN_COLUMNS.NAME + 1,
      'fullPermission': ADMIN_COLUMNS.FULL_PERMISSION + 1,
      'receiveRequests': ADMIN_COLUMNS.RECEIVE_REQUESTS + 1
    };
    
    const columnIndex = fieldMapping[updateData.field];
    if (!columnIndex) {
      return "ERROR:Invalid field name - " + updateData.field;
    }
    
    // Update the field
    let cellValue = updateData.value;
    if (updateData.field === 'fullPermission' || updateData.field === 'receiveRequests') {
      cellValue = updateData.value ? 'Yes' : 'No';
    }
    
    adminsSheet.getRange(targetRow, columnIndex).setValue(cellValue);
    
    console.log('✅ Admin field updated successfully');
    
    return "SUCCESS:Administrator updated successfully";
    
  } catch (error) {
    console.error('❌ Error updating admin field:', error);
    return "ERROR:Error updating field - " + error.toString();
  }
}

/**
 * Save admin data (add new or edit existing) - CORRECTED to accept acting admin email
 */
function saveAdminData(adminData, actingAdminEmail) {
  try {
    console.log('=== SAVE ADMIN DATA START ===');
    console.log('💾 Admin data received:', adminData);
    console.log('👤 Acting admin:', actingAdminEmail);
    
    if (!actingAdminEmail) {
      console.error('❌ Acting admin email is required');
      return "ERROR:Acting admin email is required";
    }
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      console.error('❌ Admins sheet not found');
      return "ERROR:Admins sheet not found";
    }
    
    // Get acting admin info
    const currentAdmin = getAdminByEmail(actingAdminEmail);
    if (!currentAdmin) {
      console.error('❌ Acting admin not found');
      return "ERROR:Acting admin not found";
    }
    
    if (adminData.mode === 'add') {
      console.log('➕ Adding new admin...');
      return addNewAdmin(adminData, currentAdmin, actingAdminEmail);
    } else if (adminData.mode === 'edit') {
      console.log('✏️ Editing existing admin...');
      return updateExistingAdmin(adminData, currentAdmin, actingAdminEmail);
    } else {
      console.error('❌ Invalid operation mode:', adminData.mode);
      return "ERROR:Invalid operation mode";
    }
    
  } catch (error) {
    console.error('❌ Error saving admin data:', error);
    return "ERROR:Error saving admin data - " + error.toString();
  }
}

/**
 * Add new admin - CORRECTED to accept acting admin email
 */
function addNewAdmin(adminData, currentAdmin, actingAdminEmail) {
  try {
    console.log('➕ Adding new admin:', adminData.email);
    console.log('👤 Added by:', actingAdminEmail);
    
    const adminsSheet = getSheet('Admins');
    
    // Check if email already exists
    const data = adminsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      if (rowEmail === adminData.email.toLowerCase()) {
        console.error('❌ Admin with email already exists');
        return "ERROR:An administrator with this email address already exists";
      }
    }
    
    // Generate password: first 2 letters of name + 4 random digits
    const namePrefix = adminData.name.substring(0, 2).toLowerCase();
    const randomDigits = Math.floor(1000 + Math.random() * 9000);
    const generatedPassword = namePrefix + randomDigits;
    
    console.log('🔑 Generated password for', adminData.name);
    
    // Prepare row data using standardized column mapping
    const newRow = Array(10).fill(null); // Initialize with 10 columns
    newRow[ADMIN_COLUMNS.NAME] = adminData.name;                           // A: Name
    newRow[ADMIN_COLUMNS.EMAIL] = adminData.email;                         // B: Email  
    newRow[ADMIN_COLUMNS.PASSWORD] = generatedPassword;                    // C: Password
    newRow[ADMIN_COLUMNS.FULL_PERMISSION] = adminData.fullPermission ? 'Yes' : 'No'; // D: Full Permission
    newRow[ADMIN_COLUMNS.RECEIVE_REQUESTS] = adminData.receiveRequests ? 'Yes' : 'No'; // E: Receive Requests
    newRow[ADMIN_COLUMNS.DATE_CREATED] = new Date();                       // F: Date Created
    newRow[ADMIN_COLUMNS.CREATED_BY] = currentAdmin.name || 'System';      // G: Created By
    newRow[ADMIN_COLUMNS.LAST_LOGIN] = null;                               // H: Last Login
    newRow[ADMIN_COLUMNS.STATUS] = 'active';                               // I: Status
    newRow[ADMIN_COLUMNS.STATUS_CHANGED_ON] = null;                        // J: Status Changed On
    
    // Add to sheet
    adminsSheet.appendRow(newRow);
    console.log('✅ Admin added to sheet');
    
    // Send welcome email with credentials
    const emailResult = sendAdminWelcomeEmail(
    adminData.email,                                    // adminEmail (string)
    adminData.name,                                     // adminName (string)
    generatedPassword,                                  // password (string)
    adminData.fullPermission ? 'Yes' : 'No',           // fullPermission (string)
    adminData.receiveRequests ? 'Yes' : 'No',          // receiveRequests (string)
    currentAdmin.name || 'System'                       // createdBy (string)
  );
    console.log('📧 Welcome email result:', emailResult);
    
    // Log action with acting admin email
    // logAdminAction('Added New Admin', adminData.email, actingAdminEmail);
    
    console.log('✅ Successfully added new admin:', adminData.email);
    return "SUCCESS:Administrator added successfully";
    
  } catch (error) {
    console.error('❌ Error adding new admin:', error);
    return "ERROR:Error adding new administrator - " + error.toString();
  }
}

/**
 * Update existing admin - CORRECTED to accept acting admin email
 */
function updateExistingAdmin(adminData, currentAdmin, actingAdminEmail) {
  try {
    console.log('✏️ Updating existing admin:', adminData.email);
    console.log('👤 Updated by:', actingAdminEmail);
    
    const adminsSheet = getSheet('Admins');
    const data = adminsSheet.getDataRange().getValues();
    let targetRow = -1;
    
    // Find admin by email using standardized mapping
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      if (rowEmail === adminData.email.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) {
      console.error('❌ Administrator not found');
      return "ERROR:Administrator not found";
    }
    
    console.log('✅ Found admin at row:', targetRow);
    
    // Update fields using standardized mapping
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.NAME + 1).setValue(adminData.name); // Name
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.FULL_PERMISSION + 1).setValue(adminData.fullPermission ? 'Yes' : 'No'); // Full Permission
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.RECEIVE_REQUESTS + 1).setValue(adminData.receiveRequests ? 'Yes' : 'No'); // Receive Requests
    
    // Log action with acting admin email
    // logAdminAction('Updated Admin', adminData.email, actingAdminEmail);
    
    console.log('✅ Successfully updated admin:', adminData.email);
    return "SUCCESS:Administrator updated successfully";
    
  } catch (error) {
    console.error('❌ Error updating admin:', error);
    return "ERROR:Error updating administrator - " + error.toString();
  }
}

/**
 * Change admin status (activate/deactivate) - COMPLETE FIX
 * File: admin.gs
 * REPLACE the entire changeAdminStatus function with this
 */
function changeAdminStatus(statusData, actingAdminEmail) {
  try {
    console.log('🔄 Changing admin status:', statusData);
    console.log('👤 Changed by:', actingAdminEmail);
    
    if (!actingAdminEmail) {
      return "ERROR:Acting admin email is required";
    }
    
    // CHECK PERMISSION FIRST
    const permissionCheck = canModifyAdmin(actingAdminEmail, statusData.adminEmail);
    
    if (!permissionCheck.canModify) {
      console.error('❌ Permission denied:', permissionCheck.reason);
      return "ERROR:" + permissionCheck.reason;
    }
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      return "ERROR:Admins sheet not found";
    }
    
    // Find admin by email
    const data = adminsSheet.getDataRange().getValues();
    let targetRow = -1;
    let adminData = null;
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      
      if (rowEmail === statusData.adminEmail.toLowerCase()) {
        targetRow = i + 1;
        adminData = {
          name: data[i][ADMIN_COLUMNS.NAME],
          email: data[i][ADMIN_COLUMNS.EMAIL],
          fullPermission: data[i][ADMIN_COLUMNS.FULL_PERMISSION] === 'Yes' || data[i][ADMIN_COLUMNS.FULL_PERMISSION] === true,
          receiveRequests: data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'Yes' || data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === true
        };
        break;
      }
    }
    
    if (targetRow === -1) {
      return "ERROR:Administrator not found";
    }
    
    const newStatus = statusData.newStatus; // Use newStatus directly from frontend
    const currentDate = new Date();
    
    // Update status in sheet
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.STATUS + 1).setValue(newStatus); // Column I - Status
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.STATUS_CHANGED_ON + 1).setValue(currentDate); // Column J - Status Changed On
    
    console.log('✅ Admin status updated to:', newStatus);
    
    // If activating, generate new password and send email
    if (newStatus === 'active') {
      console.log('🔑 Generating new password for activation...');
      
      // Generate new password using standardized format
      const namePrefix = adminData.name.substring(0, 2).toLowerCase();
      const randomDigits = Math.floor(1000 + Math.random() * 9000);
      const newPassword = namePrefix + randomDigits;
      
      // Update password in sheet
      adminsSheet.getRange(targetRow, ADMIN_COLUMNS.PASSWORD + 1).setValue(newPassword);
      console.log('✅ New password generated and saved');
      
      // Get acting admin name
      const actingAdminInfo = getAdminByEmail(actingAdminEmail);
      const actingAdminName = actingAdminInfo ? actingAdminInfo.name : 'Administrator';
      
      // Send reactivation email with new password
      const emailResult = sendAdminReactivationEmail(
        adminData.email,
        adminData.name,
        adminData.fullPermission ? 'Yes' : 'No',
        adminData.receiveRequests ? 'Yes' : 'No',
        newPassword,
        actingAdminName
      );
      
      console.log('📧 Reactivation email result:', emailResult);
      
      if (!emailResult.success) {
        console.error('⚠️ Email failed but status was updated');
        return "SUCCESS:Administrator activated but email failed to send";
      }
    }
    
    console.log('✅ Admin status changed successfully');
    return "SUCCESS:Administrator status changed successfully";
    
  } catch (error) {
    console.error('❌ Error changing admin status:', error);
    return "ERROR:Error changing status - " + error.toString();
  }
}

/* ============================================================================
   HELPER FUNCTIONS
   ============================================================================ */

/**
 * Get current admin info - CORRECTED to accept admin email parameter
 */
function getCurrentAdminInfo(adminEmail) {
  try {
    // Use provided admin email instead of Session.getActiveUser().getEmail()
    if (!adminEmail) {
      console.error('❌ Admin email parameter is required');
      return { 
        name: 'Unknown Administrator', 
        email: 'unknown@example.com',
        fullPermission: false,
        receiveRequests: false 
      };
    }
    
    console.log('🔍 Getting current admin info for:', adminEmail);
    
    // Use the existing getAdminByEmail function which is already corrected
    const adminInfo = getAdminByEmail(adminEmail);
    
    if (adminInfo) {
      console.log('✅ Found current admin info:', adminInfo);
      return adminInfo;
    } else {
      console.warn('⚠️ Admin not found, returning default info');
      return { 
        name: 'System Administrator', 
        email: adminEmail,
        fullPermission: false,
        receiveRequests: false 
      };
    }
    
  } catch (error) {
    console.error('❌ Error getting current admin info:', error);
    return { 
      name: 'System Administrator', 
      email: adminEmail || 'system@example.com',
      fullPermission: false,
      receiveRequests: false 
    };
  }
}
/**
 * Get admin by email
 */
function getAdminByEmail(email) {
  try {
    if (!email) {
      console.error('❌ No email provided');
      return null;
    }
    
    console.log('🔍 Getting admin by email:', email);
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      console.error('❌ Admins sheet not found');
      return null;
    }
    
    const data = adminsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      
      if (rowEmail === email.toLowerCase()) {
        const admin = {
          name: data[i][ADMIN_COLUMNS.NAME] || 'Unknown',
          email: data[i][ADMIN_COLUMNS.EMAIL],
          fullPermission: data[i][ADMIN_COLUMNS.FULL_PERMISSION] === 'Yes' || data[i][ADMIN_COLUMNS.FULL_PERMISSION] === true,
          receiveRequests: data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'Yes' || data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === true,
          status: data[i][ADMIN_COLUMNS.STATUS] || 'active'
        };
        
        console.log('✅ Found admin:', admin);
        return admin;
      }
    }
    
    console.log('❌ Admin not found with email:', email);
    return null;
    
  } catch (error) {
    console.error('❌ Error getting admin by email:', error);
    return null;
  }
}

/* ============================================================================
   EMAIL FUNCTIONS
   ============================================================================ */

/**
 * Send welcome email to new admin
 */
function sendAdminWelcomeEmail(adminEmail, adminName, password, fullPermission, receiveRequests, createdBy) {
  // Convert to strings if boolean received (backward compatibility)
  const fullPermStr = (fullPermission === true || fullPermission === 'Yes' || fullPermission === 'yes') ? 'Yes' : 'No';
  const receiveReqStr = (receiveRequests === true || receiveRequests === 'Yes' || receiveRequests === 'yes') ? 'Yes' : 'No';
  
  const credentials = buildCredentialsCard({
    'Email Address': adminEmail,
    'Temporary Password': password,
    'Full Permission': fullPermStr === 'Yes' ? 'Yes - Can manage other admins' : 'No - Limited access',
    'Receive Requests': receiveReqStr === 'Yes' ? 'Yes - Will receive notifications' : 'No - No notifications'
  });
  
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const systemUrl = configSheet ? configSheet.getRange('B1').getValue() : '#';
  const loginButton = buildEmailButton('Access Admin Portal', systemUrl, 'primary');
  
  return standardSendEmail({
    receiver: adminEmail,
    subject: 'Welcome to Holiday Management System - Admin Access',
    subtitle: 'Administrator Account Created',
    emailType: 'Admin Welcome',
    actualSender: createdBy,
    body_greeting: `Hello ${adminName},`,
    body_intro: `Welcome to the Holiday Management System! Your administrator account has been created by ${createdBy}.`,
    body_content: credentials + loginButton,
    body_securitytips: 'Please change your password immediately after your first login. Never share your credentials with anyone.',
    body_footer: 'Best regards,\nHoliday Management System',
    logDescription: `Welcome email sent to new admin ${adminName}`
  });
}

/**
 * Send reactivation email to admin
 */
function sendAdminReactivationEmail(adminEmail, adminName, fullPermission, receiveRequests, reactivatedBy) {
  // Convert to strings if boolean received
  const fullPermStr = (fullPermission === true || fullPermission === 'Yes' || fullPermission === 'yes') ? 'Yes' : 'No';
  const receiveReqStr = (receiveRequests === true || receiveRequests === 'Yes' || receiveRequests === 'yes') ? 'Yes' : 'No';
  
  const statusCard = buildEmailCard(
    'Account Reactivated',
    `Your administrator account has been reactivated.\n\nFull Permission: ${fullPermStr}\nReceive Requests: ${receiveReqStr}`,
    'success'
  );
  
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const systemUrl = configSheet ? configSheet.getRange('B1').getValue() : '#';
  const loginButton = buildEmailButton('Access Admin Portal', systemUrl, 'success');
  
  return standardSendEmail({
    receiver: adminEmail,
    subject: 'Admin Account Reactivated',
    subtitle: 'Account Status Update',
    emailType: 'Admin Reactivation',
    actualSender: reactivatedBy,
    body_greeting: `Hello ${adminName},`,
    body_intro: 'Good news! Your administrator account has been reactivated and you can now access the system.',
    body_content: statusCard + loginButton,
    body_footer: 'Best regards,\nHoliday Management System',
    logDescription: `Reactivation email sent to admin ${adminName}`
  });
}

/**
 * Send password change PIN email
 */
function sendPasswordChangePINEmail(adminEmail, adminName, pin) {
  try {
    console.log('Sending password change PIN email to:', adminEmail);
    
    const emailData = {
      to: adminEmail,
      name: adminName,
      pin: pin
    };
    
    // Call email sender function
    return sendPasswordChangePIN_Email(emailData);
    
  } catch (error) {
    console.error('Error sending password change PIN email:', error);
    return {
      success: false,
      error: 'Error sending PIN email: ' + error.toString()
    };
  }
}

/**
 * Send password change confirmation email
 */
function sendPasswordChangeConfirmationEmail(adminEmail, adminName) {
  try {
    console.log('Sending password change confirmation email to:', adminEmail);
    
    const emailData = {
      to: adminEmail,
      name: adminName,
      timestamp: new Date()
    };
    
    // Call email sender function
    return sendPasswordChangeConfirmation_Email(emailData);
    
  } catch (error) {
    console.error('Error sending password change confirmation email:', error);
    return {
      success: false,
      error: 'Error sending confirmation email: ' + error.toString()
    };
  }
}

/**
 * Reset all employee balances - Updated for 19-column structure
 */
function resetEmployeeBalances() {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const lastRow = employeesSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true }; // No employees to reset
    }
    
    // Reset opening balance to default (this could be configurable)
    const defaultBalance = 21; // Default annual leave balance
    const defaultEmergencyLeaves = 3; // Default emergency leaves
    const defaultSickLeaves = 3; // Default sick leaves
    const currentDate = new Date();
    const nextYear = new Date(currentDate.getFullYear() + 1, currentDate.getMonth(), currentDate.getDate());
    
    // Update balance, usable from, usable to dates, and leave balances for all ACTIVE employees only
    for (let i = 2; i <= lastRow; i++) {
      // Check if employee is deactivated (column J - Deactivated on)
      const deactivatedDate = employeesSheet.getRange(i, 10).getValue();
      
      if (!deactivatedDate) { // Only reset for active employees
        employeesSheet.getRange(i, 4).setValue(defaultBalance); // D - Opening Balance
        employeesSheet.getRange(i, 5).setValue(currentDate); // E - Usable from
        employeesSheet.getRange(i, 6).setValue(nextYear); // F - Usable to
        employeesSheet.getRange(i, 13).setValue(defaultEmergencyLeaves); // M - Opening emergency leaves
        employeesSheet.getRange(i, 14).setValue(defaultSickLeaves); // N - Opening Sick Leaves
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error resetting employee balances:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get employee statistics for dashboard - New function for 19-column structure
 */
function getEmployeeStatistics() {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        stats: {
          total: 0,
          active: 0,
          deactivated: 0,
          recentlyAdded: 0,
          pendingEmails: 0
        }
      };
    }
    
    let total = 0;
    let active = 0;
    let deactivated = 0;
    let recentlyAdded = 0;
    let pendingEmails = 0;
    
    const oneWeekAgo = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    
    for (let i = 1; i < data.length; i++) {
      total++;
      
      // Check if deactivated (Column J)
      const deactivatedDate = data[i][9];
      if (deactivatedDate) {
        deactivated++;
      } else {
        active++;
      }
      
      // Check if recently added (Column H)
      const addedDate = data[i][7];
      if (addedDate && new Date(addedDate) > oneWeekAgo) {
        recentlyAdded++;
      }
      
      // Check email status (Column I)
      const emailStatus = data[i][8];
      if (emailStatus && emailStatus.toString().toLowerCase() === 'pending') {
        pendingEmails++;
      }
    }
    
    return {
      success: true,
      stats: {
        total: total,
        active: active,
        deactivated: deactivated,
        recentlyAdded: recentlyAdded,
        pendingEmails: pendingEmails
      }
    };
    
  } catch (error) {
    console.error('Error getting employee statistics:', error);
    return {
      success: false,
      error: 'Error retrieving employee statistics: ' + error.toString()
    };
  }
}

/**
 * Bulk update employee weekly holidays - New function for 19-column structure
 */
function bulkUpdateWeeklyHolidays(weeklyHoliday) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const validDays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    if (!validDays.includes(weeklyHoliday)) {
      return {
        success: false,
        error: 'Invalid weekly holiday day'
      };
    }
    
    const lastRow = employeesSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true }; // No employees to update
    }
    
    // Update weekly holiday for all ACTIVE employees only
    for (let i = 2; i <= lastRow; i++) {
      // Check if employee is deactivated (column J - Deactivated on)
      const deactivatedDate = employeesSheet.getRange(i, 10).getValue();
      
      if (!deactivatedDate) { // Only update for active employees
        employeesSheet.getRange(i, 15).setValue(weeklyHoliday); // O - Weekly Holiday
      }
    }
    
    // Log action
    // logAdminAction('Bulk Weekly Holiday Update', `Set to: ${weeklyHoliday}`);
    
    return {
      success: true,
      message: `Weekly holiday updated to ${weeklyHoliday} for all active employees`
    };
    
  } catch (error) {
    console.error('Error bulk updating weekly holidays:', error);
    return {
      success: false,
      error: 'Error updating weekly holidays: ' + error.toString()
    };
  }
}

/**
 * Generate employee report - New function for 19-column structure
 */
function generateEmployeeReport() {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        report: 'No employees found'
      };
    }
    
    let report = 'EMPLOYEE REPORT\n';
    report += '================\n\n';
    
    // Count statistics
    let activeCount = 0;
    let deactivatedCount = 0;
    const weeklyHolidayCounts = {};
    const emergencyLeaveTotals = { total: 0, count: 0 };
    const sickLeaveTotals = { total: 0, count: 0 };
    const annualLeaveTotals = { total: 0, count: 0 };
    
    for (let i = 1; i < data.length; i++) {
      const deactivatedDate = data[i][9]; // J - Deactivated on
      const weeklyHoliday = data[i][14] || 'Not Set'; // O - Weekly Holiday
      const emergencyLeaves = data[i][12] || 0; // M - Emergency Leaves
      const sickLeaves = data[i][13] || 0; // N - Sick Leaves
      const annualLeaves = data[i][3] || 0; // D - Opening Balance
      
      if (deactivatedDate) {
        deactivatedCount++;
      } else {
        activeCount++;
        
        // Only count active employees for leave totals
        emergencyLeaveTotals.total += Number(emergencyLeaves);
        emergencyLeaveTotals.count++;
        
        sickLeaveTotals.total += Number(sickLeaves);
        sickLeaveTotals.count++;
        
        annualLeaveTotals.total += Number(annualLeaves);
        annualLeaveTotals.count++;
      }
      
      weeklyHolidayCounts[weeklyHoliday] = (weeklyHolidayCounts[weeklyHoliday] || 0) + 1;
    }
    
    report += `Total Employees: ${data.length - 1}\n`;
    report += `Active: ${activeCount}\n`;
    report += `Deactivated: ${deactivatedCount}\n\n`;
    
    report += 'LEAVE BALANCE STATISTICS (Active Employees Only):\n';
    report += `Average Annual Leave: ${activeCount > 0 ? (annualLeaveTotals.total / annualLeaveTotals.count).toFixed(1) : 0} days\n`;
    report += `Average Emergency Leave: ${activeCount > 0 ? (emergencyLeaveTotals.total / emergencyLeaveTotals.count).toFixed(1) : 0} days\n`;
    report += `Average Sick Leave: ${activeCount > 0 ? (sickLeaveTotals.total / sickLeaveTotals.count).toFixed(1) : 0} days\n\n`;
    
    report += 'WEEKLY HOLIDAY DISTRIBUTION:\n';
    for (const [day, count] of Object.entries(weeklyHolidayCounts)) {
      report += `${day}: ${count} employees\n`;
    }
    
    report += '\n\nGenerated on: ' + new Date().toLocaleString();
    
    return {
      success: true,
      report: report
    };
    
  } catch (error) {
    console.error('Error generating employee report:', error);
    return {
      success: false,
      error: 'Error generating report: ' + error.toString()
    };
  }
}

/**
 * Search employees by criteria - New function for 19-column structure
 */
function searchEmployees(searchCriteria) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        data: []
      };
    }
    
    const searchTerm = searchCriteria.toLowerCase();
    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      const employee = {
        'ID': data[i][0],                           // A
        'Name': data[i][1],                         // B
        'E-mail': data[i][2],                      // C
        'Opening Balance': data[i][3],              // D
        'Usable from': data[i][4],                 // E
        'Usable to': data[i][5],                   // F
        'Password': data[i][6],                    // G
        'Added on': data[i][7],                    // H
        'Email status': data[i][8],                // I
        'Deactivated on': data[i][9],              // J
        'Last Login': data[i][10],                 // K
        'Active/Inactive': data[i][11],            // L
        'Emergency Leaves': data[i][12],           // M
        'Sick Leaves': data[i][13],               // N
        'Weekly Holiday': data[i][14],             // O
        'Added by': data[i][15],                   // P
        'Deactivated BY': data[i][16],             // Q
        'Reactivated on': data[i][17],             // R
        'Reactivated by': data[i][18],             // S
        'Status': data[i][9] ? 'Deactivated' : 'Active'
      };
      
      // Search in name, email, ID, status, weekly holiday, or added by
      const searchableText = `${employee.ID} ${employee.Name} ${employee['E-mail']} ${employee.Status} ${employee['Weekly Holiday']} ${employee['Added by'] || ''}`.toLowerCase();
      
      if (searchableText.includes(searchTerm)) {
        employees.push(employee);
      }
    }
    
    return {
      success: true,
      data: employees
    };
    
  } catch (error) {
    console.error('Error searching employees:', error);
    return {
      success: false,
      error: 'Error searching employees: ' + error.toString()
    };
  }
}

/**
 * Update employee email status - New function for 19-column structure
 */
function updateEmployeeEmailStatus(employeeId, emailStatus) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const validStatuses = ['Pending', 'Sent', 'Failed', 'Delivered'];
    if (!validStatuses.includes(emailStatus)) {
      return {
        success: false,
        error: 'Invalid email status'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    let targetRow = -1;
    
    // Find employee by ID
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === parseInt(employeeId)) {
        targetRow = i + 1; // Convert to 1-based row index
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }
    
    // Update email status
    employeesSheet.getRange(targetRow, 9).setValue(emailStatus); // I - Email status
    
    return {
      success: true,
      message: 'Email status updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating email status:', error);
    return {
      success: false,
      error: 'Error updating email status: ' + error.toString()
    };
  }
}

/**
 * Get employees by status - New function for 19-column structure
 */
function getEmployeesByStatus(status) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        data: []
      };
    }
    
    const employees = [];
    const isActiveFilter = status.toLowerCase() === 'active';
    
    for (let i = 1; i < data.length; i++) {
      const deactivatedDate = data[i][9]; // J - Deactivated on
      const isActive = !deactivatedDate;
      
      if (isActive === isActiveFilter) {
        const employee = {
          'ID': data[i][0],                           // A
          'Name': data[i][1],                         // B
          'E-mail': data[i][2],                      // C
          'Opening Balance': data[i][3],              // D
          'Usable from': data[i][4],                 // E
          'Usable to': data[i][5],                   // F
          'Password': data[i][6],                    // G
          'Added on': data[i][7],                    // H
          'Email status': data[i][8],                // I
          'Deactivated on': data[i][9],              // J
          'Last Login': data[i][10],                 // K
          'Active/Inactive': data[i][11],            // L
          'Emergency Leaves': data[i][12],           // M
          'Sick Leaves': data[i][13],               // N
          'Weekly Holiday': data[i][14],             // O
          'Added by': data[i][15],                   // P
          'Deactivated BY': data[i][16],             // Q
          'Reactivated on': data[i][17],             // R
          'Reactivated by': data[i][18],             // S
          'Status': isActive ? 'Active' : 'Deactivated'
        };
        employees.push(employee);
      }
    }
    
    return {
      success: true,
      data: employees
    };
    
  } catch (error) {
    console.error('Error getting employees by status:', error);
    return {
      success: false,
      error: 'Error retrieving employees: ' + error.toString()
    };
  }
}

/**
 * Bulk update employee leave balances - New function for 19-column structure
 */
function bulkUpdateLeaveBalances(leaveType, newBalance) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const validLeaveTypes = ['annual', 'emergency', 'sick'];
    if (!validLeaveTypes.includes(leaveType.toLowerCase())) {
      return {
        success: false,
        error: 'Invalid leave type. Must be: annual, emergency, or sick'
      };
    }
    
    if (newBalance < 0 || newBalance > 50) {
      return {
        success: false,
        error: 'Leave balance must be between 0 and 50 days'
      };
    }
    
    const lastRow = employeesSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true }; // No employees to update
    }
    
    let columnIndex;
    let leaveTypeName;
    
    switch (leaveType.toLowerCase()) {
      case 'annual':
        columnIndex = 4; // D - Opening Balance
        leaveTypeName = 'Annual Leave';
        break;
      case 'emergency':
        columnIndex = 13; // M - Opening emergency leaves
        leaveTypeName = 'Emergency Leave';
        break;
      case 'sick':
        columnIndex = 14; // N - Opening Sick Leaves
        leaveTypeName = 'Sick Leave';
        break;
    }
    
    // Update leave balance for all ACTIVE employees only
    let updatedCount = 0;
    for (let i = 2; i <= lastRow; i++) {
      // Check if employee is deactivated (column J - Deactivated on)
      const deactivatedDate = employeesSheet.getRange(i, 10).getValue();
      
      if (!deactivatedDate) { // Only update for active employees
        employeesSheet.getRange(i, columnIndex).setValue(newBalance);
        updatedCount++;
      }
    }
    
    // Log action
    // logAdminAction('Bulk Leave Balance Update', `${leaveTypeName}: ${newBalance} days for ${updatedCount} employees`);
    
    return {
      success: true,
      message: `${leaveTypeName} balance updated to ${newBalance} days for ${updatedCount} active employees`
    };
    
  } catch (error) {
    console.error('Error bulk updating leave balances:', error);
    return {
      success: false,
      error: 'Error updating leave balances: ' + error.toString()
    };
  }
}

/**
 * Get employee activity log - New function for tracking employee changes
 */
function getEmployeeActivityLog(employeeId) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    let employeeData = null;
    
    // Find employee by ID
    for (let i = 1; i < data.length; i++) {
      if (parseInt(data[i][0]) === parseInt(employeeId)) {
        employeeData = data[i];
        break;
      }
    }
    
    if (!employeeData) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }
    
    const activity = [];
    
    // Add creation activity
    if (employeeData[7]) { // H - Added on
      activity.push({
        date: employeeData[7],
        action: 'Employee Added',
        admin: employeeData[15] || 'Unknown', // P - Added by
        details: `Employee created with ID ${employeeData[0]}`
      });
    }
    
    // Add deactivation activity
    if (employeeData[9]) { // J - Deactivated on
      activity.push({
        date: employeeData[9],
        action: 'Employee Deactivated',
        admin: employeeData[16] || 'Unknown', // Q - Deactivated BY
        details: 'Employee account deactivated'
      });
    }
    
    // Add reactivation activity
    if (employeeData[17]) { // R - Reactivated on
      activity.push({
        date: employeeData[17],
        action: 'Employee Reactivated',
        admin: employeeData[18] || 'Unknown', // S - Reactivated by
        details: 'Employee account reactivated'
      });
    }
    
    // Add last login activity
    if (employeeData[10]) { // K - Last Login
      activity.push({
        date: employeeData[10],
        action: 'Last Login',
        admin: 'System',
        details: 'Employee logged into the system'
      });
    }
    
    // Sort by date (newest first)
    activity.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    return {
      success: true,
      employee: {
        id: employeeData[0],
        name: employeeData[1],
        email: employeeData[2]
      },
      activity: activity
    };
    
  } catch (error) {
    console.error('Error getting employee activity log:', error);
    return {
      success: false,
      error: 'Error retrieving activity log: ' + error.toString()
    };
  }
}

/* ============================================================================
   PHASE 2: ADD EMPLOYEE BACKEND FUNCTIONS
   Add these functions to admin.gs file
   ============================================================================ */

/**
 * Get system dates and default values from Config sheet
 * Returns strings only to avoid parsing issues
 */
function getSystemDatesAndDefaults() {
  try {
    const configSheet = getSheet('Config');
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    // Get system dates (B3 and B4)
    const systemStartDate = configSheet.getRange('B3').getValue();
    const systemEndDate = configSheet.getRange('B4').getValue();
    
    // Get default leave values (B9, B10, B11)
    const annualDefault = configSheet.getRange('B9').getValue() || 21;
    const sickDefault = configSheet.getRange('B10').getValue() || 7;
    const emergencyDefault = configSheet.getRange('B11').getValue() || 3;
    
    // Format dates properly for HTML date inputs (YYYY-MM-DD)
    let startDateStr = '';
    let endDateStr = '';
    
    if (systemStartDate && systemStartDate instanceof Date) {
      startDateStr = Utilities.formatDate(systemStartDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    if (systemEndDate && systemEndDate instanceof Date) {
      endDateStr = Utilities.formatDate(systemEndDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    
    console.log('System dates and defaults retrieved successfully');
    
    return {
      success: true,
      systemStartDate: startDateStr,
      systemEndDate: endDateStr,
      annualDefault: annualDefault.toString(),
      sickDefault: sickDefault.toString(),
      emergencyDefault: emergencyDefault.toString()
    };
    
  } catch (error) {
    console.error('Error getting system dates and defaults:', error);
    return {
      success: false,
      error: 'Error retrieving system configuration: ' + error.toString()
    };
  }
}

/**
 * Check if email is unique in the system
 * Returns simple object with boolean result
 */
function checkEmailUniqueness(email) {
  try {
    if (!email || typeof email !== 'string') {
      return { isUnique: false };
    }
    
    const emailLower = email.toLowerCase().trim();
    const employeesSheet = getSheet('Employees');
    
    if (employeesSheet) {
      const data = employeesSheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][2] && data[i][2].toString().toLowerCase() === emailLower) {
          return { isUnique: false };
        }
      }
    }
    
    return { isUnique: true };
    
  } catch (error) {
    console.error('Error checking email uniqueness:', error);
    return { isUnique: false };
  }
}

/**
 * Generate employee password using convention: first 2 letters + 4 random digits
 */
function generateEmployeePassword(employeeName) {
  try {
    // Get first 2 letters of name (remove spaces and special characters)
    const cleanName = employeeName.replace(/[^a-zA-Z]/g, '').toLowerCase();
    const firstTwoLetters = cleanName.substring(0, 2);
    
    // Generate 4 random digits
    const randomDigits = Math.floor(1000 + Math.random() * 9000).toString();
    
    // Combine and capitalize first letters
    const password = firstTwoLetters.charAt(0).toUpperCase() + 
                    firstTwoLetters.charAt(1) + 
                    randomDigits;
    
    return password;
    
  } catch (error) {
    console.error('Error generating employee password:', error);
    // Fallback to simple random password
    return 'Emp' + Math.floor(1000 + Math.random() * 9000).toString();
  }
}

/**
 * Add new employee to the system
 * All parameters are strings to avoid parsing issues
 */
function addNewEmployee(name, email, weeklyHolidays, usableFrom, usableTo, annualDays, emergencyDays, sickDays, currentAdminString) {
  try {
    // Convert all parameters to strings explicitly
    const nameStr = String(name || '');
    const emailStr = String(email || '');
    const weeklyHolidaysStr = String(weeklyHolidays || '');
    const usableFromStr = String(usableFrom || '');
    const usableToStr = String(usableTo || '');
    const annualDaysStr = String(annualDays || '');
    const emergencyDaysStr = String(emergencyDays || '');
    const sickDaysStr = String(sickDays || '');
    const adminStr = String(currentAdminString || '');
    
    console.log('addNewEmployee called with string parameters:', {
      name: nameStr,
      email: emailStr,
      weeklyHolidays: weeklyHolidaysStr,
      usableFrom: usableFromStr,
      usableTo: usableToStr,
      annualDays: annualDaysStr,
      emergencyDays: emergencyDaysStr,
      sickDays: sickDaysStr,
      admin: adminStr
    });
    
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    // Validate admin context
    if (!adminStr || adminStr.trim() === '') {
      return {
        success: false,
        error: 'Admin context missing - please refresh and try again'
      };
    }
    
    console.log('Employee being added by admin:', adminStr);
    
    // Double-check email uniqueness
    const emailCheck = checkEmailUniqueness(emailStr);
    if (!emailCheck.isUnique) {
      return {
        success: false,
        error: 'Email address already exists in the system'
      };
    }
    
    // Generate employee ID - get highest existing ID and add 1
    let nextId = 1;
    const lastRow = employeesSheet.getLastRow();
    
    if (lastRow > 1) {
      const data = employeesSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const currentId = parseInt(data[i][0]);
        if (!isNaN(currentId) && currentId >= nextId) {
          nextId = currentId + 1;
        }
      }
    }
    
    // Generate password using convention: first 2 letters + 4 random digits
    const generatedPassword = generateEmployeePassword(nameStr.trim());
    
    // Parse weekly holidays (comma-separated string)
    const weeklyHolidaysArray = weeklyHolidaysStr ? weeklyHolidaysStr.split(',').map(day => day.trim()) : [];
    
    // Prepare row data for 19-column structure (A-S)
    const rowData = [
      nextId,                                      // A - ID
      nameStr.trim(),                              // B - Name
      emailStr.trim().toLowerCase(),               // C - E-mail
      parseInt(annualDaysStr) || 21,               // D - Opening Balance
      new Date(usableFromStr),                     // E - Usable from
      new Date(usableToStr),                       // F - Usable to
      generatedPassword,                           // G - Password
      new Date(),                                  // H - Added on
      'Pending',                                   // I - Email status (will be updated after email)
      null,                                        // J - Deactivated on
      null,                                        // K - Last Login
      'Active',                                    // L - Active/Inactive
      parseInt(emergencyDaysStr) || 3,             // M - Opening emergency leaves
      parseInt(sickDaysStr) || 3,                  // N - Opening Sick Leaves
      '',                                          // O - Weekly Holiday (IGNORED - kept for compatibility)
      adminStr,                                    // P - Added by
      null,                                        // Q - Deactivated BY
      null,                                        // R - Reactivated on
      null                                         // S - Reactivated BY
    ];
    
    // Add the new employee to the sheet
    employeesSheet.appendRow(rowData);
    
    // Create weekly holiday period in "Weekly Holidays" sheet
    if (weeklyHolidaysArray.length > 0) {
      const weeklyHolidayResult = createWeeklyHolidayPeriod(
        String(nextId),                     // Employee ID as string
        nameStr.trim(),                     // Employee Name as string
        usableFromStr,                      // Start Date as string
        null,                               // End Date (null for standard periods)
        'standard',                         // Period Type as string
        weeklyHolidaysStr,                  // Holiday Days String (already comma-separated)
        adminStr                            // Admin Email as string
      );
      
      if (!weeklyHolidayResult.success) {
        console.error('Warning: Failed to create weekly holiday period:', weeklyHolidayResult.error);
        // Don't fail the entire operation - employee is already added
      } else {
        console.log('✅ Weekly holiday period created:', weeklyHolidayResult.periodId);
      }
    }
    
    // Send welcome email and update email status
    let emailResult = { success: false, message: 'Email sending skipped' };
    try {
      emailResult = sendEmployeeWelcomeEmail(
        nameStr.trim(), 
        emailStr.trim(), 
        String(nextId), 
        generatedPassword,
        adminStr
      );
      
      // Update email status based on email sending result
      if (emailResult.success) {
        employeesSheet.getRange(lastRow + 1, 9).setValue("Sent");
      } else {
        employeesSheet.getRange(lastRow + 1, 9).setValue("Failed");
      }
    } catch (emailError) {
      console.error('Error sending welcome email:', emailError);
      employeesSheet.getRange(lastRow + 1, 9).setValue("Failed");
    }
    
    console.log(`Employee ${nameStr} added successfully by ${adminStr} with ID ${nextId}`);
    
    return {
      success: true,
      message: 'Employee added successfully',
      employeeId: String(nextId),
      emailSent: emailResult.success,
      emailMessage: emailResult.message || 'Email status unknown'
    };
    
  } catch (error) {
    console.error('Error adding new employee:', error);
    return {
      success: false,
      error: 'Error adding employee: ' + error.toString()
    };
  }
}

/**
 * Send welcome email to new employee
 * This function should be added to email-sender.gs
 */
function sendWelcomeEmail(employeeName, employeeEmail, employeeId, temporaryPassword) {
  try {
    // Get current year for template
    const currentYear = new Date().getFullYear().toString();
    
    // Get system URL (you may need to configure this in Config sheet)
    const systemUrl = getSystemUrl(); // We'll create this helper function
    
    // Prepare email template with replacements
    const emailHtml = getWelcomeEmailTemplate()
      .replace(/{{EMPLOYEE_NAME}}/g, employeeName)
      .replace(/{{EMPLOYEE_ID}}/g, employeeId)
      .replace(/{{EMPLOYEE_EMAIL}}/g, employeeEmail)
      .replace(/{{TEMPORARY_PASSWORD}}/g, temporaryPassword)
      .replace(/{{SYSTEM_URL}}/g, systemUrl)
      .replace(/{{CURRENT_YEAR}}/g, currentYear);
    
    // Send email
    GmailApp.sendEmail(
      employeeEmail,
      'Welcome to Holiday Management System - Your Account Details',
      '', // Plain text version (empty for HTML email)
      {
        htmlBody: emailHtml,
        name: 'Holiday Management System'
      }
    );
    
    // Save to email history
    saveEmailHistory(
      employeeEmail,
      'Welcome Email',
      `Welcome email sent to new employee: ${employeeName}`,
      'Sent'
    );
    
    return {
      success: true,
      message: 'Welcome email sent successfully'
    };
    
  } catch (error) {
    console.error('Error sending welcome email:', error);
    return {
      success: false,
      error: 'Error sending welcome email: ' + error.toString()
    };
  }
}

/**
 * Get system URL from Config sheet or return default
 */
function getSystemUrl() {
  try {
    // You can store the system URL in Config sheet (e.g., B20)
    // For now, return the current web app URL
    return ScriptApp.getService().getUrl();
  } catch (error) {
    console.error('Error getting system URL:', error);
    return 'https://your-system-url.com'; // Fallback URL
  }
}

/**
 * Save email to Email History sheet
 */
function saveEmailHistoryWithString(recipientEmail, emailType, description, status, currentAdminString) {
  try {
    let emailHistorySheet = getSheet('Email History');
    
    // Create Email History sheet if it doesn't exist
    if (!emailHistorySheet) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      emailHistorySheet = ss.insertSheet('Email History');
      
      const headers = [
        'Timestamp',           // A
        'Recipient Email',     // B
        'Email Type',          // C
        'Description',         // D
        'Status',              // E
        'Sent By Admin',       // F
        'System User'          // G
      ];
      
      emailHistorySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      const headerRange = emailHistorySheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#3498db');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(12);
      
      emailHistorySheet.autoResizeColumns(1, headers.length);
    }
    
    const adminEmail = currentAdminString || 'System';
    const systemUser = Session.getActiveUser().getEmail();
    
    const rowData = [
      new Date(),                        // Timestamp
      recipientEmail || 'Unknown',       // Recipient Email
      emailType || 'General',            // Email Type
      description || 'No description',   // Description
      status || 'Unknown',               // Status
      adminEmail,                        // Sent By Admin
      systemUser                         // System User
    ];
    
    emailHistorySheet.appendRow(rowData);
    
    console.log(`Email history saved - Admin: ${adminEmail}, Recipient: ${recipientEmail}`);
    
    return { 
      success: true, 
      message: 'Email history saved successfully' 
    };
    
  } catch (error) {
    console.error('Error saving email history:', error);
    return { 
      success: false, 
      error: 'Failed to save email history: ' + error.toString() 
    };
  }
}

function toggleEmployeeStatus(employeeId, currentAdminString) {
  try {
    if (!currentAdminString) {
      return {
        success: false,
        error: 'Admin context missing - please refresh and try again'
      };
    }
    
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    console.log('Employee status change requested by admin:', currentAdminString);
    
    // Find employee row
    const lastRow = employeesSheet.getLastRow();
    const data = employeesSheet.getRange(2, 1, lastRow - 1, 19).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == employeeId) {
        const rowNum = i + 2;
        const row = data[i];
        
        // Determine current status based on latest timestamps
        const deactivatedOn = row[9];  // J - Deactivated on
        const reactivatedOn = row[17]; // R - Reactivated on
        
        let currentStatus = 'Active';
        if (deactivatedOn && (!reactivatedOn || new Date(deactivatedOn) > new Date(reactivatedOn))) {
          currentStatus = 'Inactive';
        }
        
        const currentDateTime = new Date();
        
        if (currentStatus === 'Active') {
          // Deactivate employee
          employeesSheet.getRange(rowNum, 10).setValue(currentDateTime);     // J - Deactivated on
          employeesSheet.getRange(rowNum, 17).setValue(currentAdminString);  // Q - Deactivated By
          employeesSheet.getRange(rowNum, 12).setValue('Inactive');            // L - Active/Inactive
          
          // Log action
          // logAdminActionWithString(currentAdminString, 'Employee Deactivated', `ID: ${employeeId}, Name: ${row[1]}`);
          
          return {
            success: true,
            message: 'Employee deactivated successfully',
            newStatus: 'Inactive'
          };
        } else {
          // Reactivate employee
          employeesSheet.getRange(rowNum, 18).setValue(currentDateTime);     // R - Reactivated on
          employeesSheet.getRange(rowNum, 19).setValue(currentAdminString);  // S - Reactivated BY
          employeesSheet.getRange(rowNum, 12).setValue('Active');          // L - Active/Inactive
          
          // Log action
          // logAdminActionWithString(currentAdminString, 'Employee Reactivated', `ID: ${employeeId}, Name: ${row[1]}`);
          
          return {
            success: true,
            message: 'Employee reactivated successfully',
            newStatus: 'Active'
          };
        }
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error toggling employee status:', error);
    return {
      success: false,
      error: 'Error updating employee status: ' + error.toString()
    };
  }
}

/**
 * Get employee by ID for editing
 */
function getEmployeeById(employeeId) {
  try {
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == employeeId) {
        const row = data[i];
        
        // Format dates for HTML inputs (YYYY-MM-DD)
        const usableFrom = row[4] ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        const usableTo = row[5] ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
        
        return {
          success: true,
          employee: {
            id: row[0].toString(),                // A - ID
            name: row[1] || '',                   // B - Name
            email: row[2] || '',                  // C - E-mail
            annualDays: row[3] || 0,              // D - Annual Opening Balance
            usableFrom: usableFrom,               // E - Usable from (formatted)
            usableTo: usableTo,                   // F - Usable to (formatted)
            weeklyHoliday: row[14] || '',         // O - Weekly Holiday
            emergencyDays: row[12] || 0,          // M - Opening emergency leaves
            sickDays: row[13] || 0                // N - Opening Sick Leaves
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error getting employee by ID:', error);
    return {
      success: false,
      error: 'Error retrieving employee: ' + error.toString()
    };
  }
}

/**
 * Update employee data
 */
function updateEmployee(employeeId, name, email, usableFrom, usableTo, annualDays, emergencyDays, sickDays, currentAdminString) {
  try {
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    // Validate admin context
    if (!currentAdminString) {
      return {
        success: false,
        error: 'Admin context missing'
      };
    }
    
    console.log('Updating employee by admin:', currentAdminString);
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == employeeId) {
        const rowNum = i + 1;
        
        // Convert dates
        const usableFromDate = new Date(usableFrom);
        const usableToDate = new Date(usableTo);
        
        // Validate dates
        if (isNaN(usableFromDate.getTime()) || isNaN(usableToDate.getTime())) {
          return {
            success: false,
            error: 'Invalid date format'
          };
        }
        
        // Update only allowed fields (B, C, E, F, D, M, N)
        employeesSheet.getRange(rowNum, 2).setValue(name.trim() || '');                    // B - Name
        employeesSheet.getRange(rowNum, 3).setValue(email.trim().toLowerCase() || '');    // C - E-mail
        employeesSheet.getRange(rowNum, 4).setValue(parseInt(annualDays) || 0);           // D - Annual Opening Balance
        employeesSheet.getRange(rowNum, 5).setValue(usableFromDate);                      // E - Usable from
        employeesSheet.getRange(rowNum, 6).setValue(usableToDate);                        // F - Usable to
        employeesSheet.getRange(rowNum, 13).setValue(parseInt(emergencyDays) || 0);       // M - Opening emergency leaves
        employeesSheet.getRange(rowNum, 14).setValue(parseInt(sickDays) || 0);            // N - Opening Sick Leaves
        
        // Log admin action
        // logAdminActionWithString(currentAdminString, 'Employee Updated', `ID: ${employeeId}, Name: ${name.trim()}`);
        
        console.log(`Employee ${employeeId} updated by ${currentAdminString}`);
        
        return {
          success: true,
          message: 'Employee updated successfully'
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error updating employee:', error);
    return {
      success: false,
      error: 'Error updating employee: ' + error.toString()
    };
  }
}

function updateEmployeeWithPasswordReset(employeeId, name, email, usableFrom, usableTo, annualDays, emergencyDays, sickDays, resetPassword, currentAdminString) {
  try {
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    // Validate admin context
    if (!currentAdminString) {
      return {
        success: false,
        error: 'Admin context missing'
      };
    }
    
    console.log('Updating employee by admin:', currentAdminString);
    console.log('Password reset requested:', resetPassword);
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == employeeId) {
        const rowNum = i + 1;
        
        // Convert dates
        const usableFromDate = new Date(usableFrom);
        const usableToDate = new Date(usableTo);
        
        // Validate dates
        if (isNaN(usableFromDate.getTime()) || isNaN(usableToDate.getTime())) {
          return {
            success: false,
            error: 'Invalid date format'
          };
        }
        
        let newPassword = null;
        let passwordReset = false;
        
        // Generate new password if requested
        if (resetPassword === 'true') {
          newPassword = generateEmployeePassword(name);
          employeesSheet.getRange(rowNum, 7).setValue(newPassword); // G - Password
          passwordReset = true;
          console.log('New password generated for employee:', employeeId);
        }
        
        // Update employee data
        employeesSheet.getRange(rowNum, 2).setValue(name.trim() || '');                    // B - Name
        employeesSheet.getRange(rowNum, 3).setValue(email.trim().toLowerCase() || '');    // C - E-mail
        employeesSheet.getRange(rowNum, 4).setValue(parseInt(annualDays) || 0);           // D - Annual Opening Balance
        employeesSheet.getRange(rowNum, 5).setValue(usableFromDate);                      // E - Usable from
        employeesSheet.getRange(rowNum, 6).setValue(usableToDate);                        // F - Usable to
        employeesSheet.getRange(rowNum, 13).setValue(parseInt(emergencyDays) || 0);       // M - Opening emergency leaves
        employeesSheet.getRange(rowNum, 14).setValue(parseInt(sickDays) || 0);            // N - Opening Sick Leaves
        
        // Send password email if password was reset
        if (passwordReset && newPassword) {
          try {
            const emailResult = sendEmployeeWelcomeEmail(
              name.trim(),
              email.trim(),
              employeeId,
              newPassword,
              currentAdminString
            );
            
            // Update email status
            if (emailResult.success) {
              employeesSheet.getRange(rowNum, 9).setValue("Sent"); // I - Email status
            } else {
              employeesSheet.getRange(rowNum, 9).setValue("Failed"); // I - Email status
            }
          } catch (emailError) {
            console.error('Error sending password reset email:', emailError);
            employeesSheet.getRange(rowNum, 9).setValue("Failed"); // I - Email status
          }
        }
        
        // Log admin action
        const actionDetails = passwordReset ? 
          `ID: ${employeeId}, Name: ${name.trim()}, Password Reset: Yes` : 
          `ID: ${employeeId}, Name: ${name.trim()}`;
        
        // logAdminActionWithString(currentAdminString, 'Employee Updated', actionDetails);
        
        console.log(`Employee ${employeeId} updated by ${currentAdminString}`);
        
        return {
          success: true,
          message: 'Employee updated successfully',
          passwordReset: passwordReset
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error updating employee:', error);
    return {
      success: false,
      error: 'Error updating employee: ' + error.toString()
    };
  }
}

/* ============================================================================
   OFFICIAL HOLIDAYS ASSIGNMENT MANAGEMENT - ADD TO END OF admin.gs
   Functions for managing holiday employee assignments and notifications
   ============================================================================ */

/**
 * Get active employees for holiday assignment with weekly holiday info from Weekly Holidays sheet
 * File: admin.gs
 * REPLACE the entire function
 */
function getActiveEmployeesForAssignment() {
  try {
    console.log('Getting active employees for assignment...');
    
    // Get employees sheet
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return "ERROR:Employees sheet not found";
    }
    
    const empData = employeesSheet.getDataRange().getValues();
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    const currentDate = new Date();
    
    const enhancedEmployees = [];
    
    // Loop through all employees
    for (let i = 1; i < empData.length; i++) {
      const row = empData[i];
      const empId = row[0];
      const empName = row[1];
      const empEmail = row[2];
      const deactivatedOn = row[9]; // Column J
      
      // Skip deactivated employees
      if (deactivatedOn) {
        continue;
      }
      
      // Get weekly holiday for this employee from Weekly Holidays sheet
      let weeklyHoliday = 'None'; // Default if no active period found
      
      if (weeklyHolidaysSheet) {
        const weeklyData = weeklyHolidaysSheet.getDataRange().getValues();
        
        // Find active weekly holiday period for this employee
        for (let j = 1; j < weeklyData.length; j++) {
          const periodRow = weeklyData[j];
          const periodEmpId = periodRow[1]; // Column B - Employee ID
          const startDate = periodRow[3] ? new Date(periodRow[3]) : null; // Column D
          const endDate = periodRow[4] ? new Date(periodRow[4]) : null; // Column E
          const periodType = periodRow[6]; // Column G - Period Type
          
          // Check if this period matches employee and is currently active
          if (String(periodEmpId) === String(empId)) {
            // Check if current date falls within this period
            const isActive = startDate && currentDate >= startDate && (!endDate || currentDate <= endDate);
            
            if (isActive) {
              // Extract holiday days from columns H-N (Friday to Thursday)
              const dayNames = ['Fri', 'Sat', 'Sun', 'Mon', 'Tue', 'Wed', 'Thu'];
              const activeDays = [];
              
              for (let k = 0; k < dayNames.length; k++) {
                const colValue = periodRow[7 + k]; // Columns H-N
                if (colValue === true || colValue === 'True' || colValue === 'true') {
                  activeDays.push(dayNames[k]);
                }
              }
              
              // Join multiple days with comma, or show 'None' if empty
              weeklyHoliday = activeDays.length > 0 ? activeDays.join(', ') : 'None';
              
              // Temp periods override standard, so if we found a temp, use it and stop
              if (periodType === 'temp') {
                break;
              }
            }
          }
        }
      }
      
      // Format: name|email|id|weeklyHoliday
      enhancedEmployees.push(`${empName}|${empEmail}|${empId}|${weeklyHoliday}`);
    }
    
    if (enhancedEmployees.length === 0) {
      return "SUCCESS:EMPTY";
    }
    
    console.log(`✅ Found ${enhancedEmployees.length} active employees with weekly holidays`);
    return "SUCCESS:" + enhancedEmployees.join(';');
    
  } catch (error) {
    console.error('Error getting active employees for assignment:', error);
    return "ERROR:" + error.toString();
  }
}

/**
 * Send single holiday assignment notification
 */
function sendSingleHolidayNotification(data) {
  try {
    const template = getHolidayAssignmentEmailTemplate();
    const isWorkingDay = data.employee.status === 'working day';
    
    // Calculate holiday dates for display
    const startDate = new Date(data.startDate);
    const endDate = new Date(data.endDate);
    const duration = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    
    const emailData = {
      to: data.employee.email,
      subject: `🏢 ${isWorkingDay ? 'Work Assignment' : 'Holiday Notice'} - ${data.holidayName}`,
      htmlBody: template
        .replace(/{{EMPLOYEE_NAME}}/g, data.employee.name)
        .replace(/{{HOLIDAY_NAME}}/g, data.holidayName)
        .replace(/{{START_DATE}}/g, formatLongDateForDisplay(data.startDate))
        .replace(/{{END_DATE}}/g, formatLongDateForDisplay(data.endDate))
        .replace(/{{DURATION}}/g, duration === 1 ? '1 day' : `${duration} days`)
        .replace(/{{ASSIGNMENT_STATUS}}/g, isWorkingDay ? 'Working Day' : 'Day Off')
        .replace(/{{ASSIGNMENT_MESSAGE}}/g, isWorkingDay ? 
          'You are scheduled to work during this official holiday. Compensation days will be added to your leave balance as per company policy.' :
          'You have the day off during this official holiday. Enjoy your time!')
        .replace(/{{ADMIN_NAME}}/g, data.adminName)
        .replace(/{{SYSTEM_URL}}/g, getSystemUrl())
        .replace(/{{CURRENT_YEAR}}/g, new Date().getFullYear())
        .replace(/{{EMAIL_SUBTITLE}}/g, isWorkingDay ? '💼 Holiday Work Assignment' : '🎉 Holiday Notice'),
      trigger: 'Holiday Assignment'
    };
    
    return sendEmail(emailData);
    
  } catch (error) {
    console.error('Error sending single holiday notification:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   WEEKLY HOLIDAYS MANAGEMENT - ADD TO END OF admin.gs
   Backend functions for managing employee weekly holiday periods
   ============================================================================ */

/**
 * Get all weekly holiday periods with calculated status
 */
function getWeeklyHolidaysData(adminEmail) {
  try {
    console.log('🔍 Getting weekly holidays data for admin:', adminEmail);
    
    if (!adminEmail) {
      return {
        success: false,
        error: 'Admin email is required'
      };
    }
    
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return {
        success: false,
        error: 'Weekly Holidays sheet not found'
      };
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        periods: [],
        message: 'No weekly holiday periods found'
      };
    }
    
    const currentDate = new Date();
    const periods = [];
    
    // Process each period row (starting from row 2)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      const periodId = row[0] || '';
      const employeeId = row[1] || '';
      const employeeName = row[2] || '';
      const startDate = row[3] ? new Date(row[3]) : null;
      const endDate = row[4] ? new Date(row[4]) : null;
      const periodType = row[6] || '';
      
      // Extract holiday days (columns H-N: Fri, Sat, Sun, Mon, Tue, Wed, Thu)
      const holidayDays = [];
      const dayColumns = ['Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];
      
      for (let j = 0; j < dayColumns.length; j++) {
        if (row[7 + j] === 'True' || row[7 + j] === true) {
          holidayDays.push(dayColumns[j]);
        }
      }
      
      // Calculate status based on current date and period boundaries
      let status = 'Inactive';
      
      if (periodType === 'standard') {
        if (startDate && startDate <= currentDate) {
          if (!endDate || endDate >= currentDate) {
            // Check if any temp period is overriding this standard period
            const hasActiveTempPeriod = hasActiveTemporaryPeriod(employeeId, currentDate, data);
            status = hasActiveTempPeriod ? 'Inactive' : 'Active';
          } else {
            status = 'Expired'; // Has end date and it's in the past
          }
        } else if (startDate && startDate > currentDate) {
          status = 'Inactive'; // Future start date
        }
      } else if (periodType === 'temp') {
        if (startDate && endDate) {
          if (currentDate >= startDate && currentDate <= endDate) {
            status = 'Active';
          } else if (currentDate < startDate) {
            status = 'Inactive'; // Future start date
          } else {
            status = 'Expired'; // Past end date
          }
        }
      }
      
      // Format dates for display
      const startDateStr = startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      const endDateStr = endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      
      periods.push({
        periodId: periodId.toString(),
        employeeId: employeeId.toString(),
        employeeName: employeeName.toString(),
        startDate: startDateStr,
        endDate: endDateStr,
        periodType: periodType.toString(),
        status: status,
        holidayDays: holidayDays.join(','),
        rowIndex: i + 1 // For update operations
      });
    }
    
    console.log('✅ Weekly holidays data retrieved successfully. Found', periods.length, 'periods');
    
    return {
      success: true,
      periods: periods
    };
    
  } catch (error) {
    console.error('Error getting weekly holidays data:', error);
    return {
      success: false,
      error: 'Error retrieving weekly holidays: ' + error.toString()
    };
  }
}

/**
 * Helper function to check if employee has active temporary period on specific date
 */
function hasActiveTemporaryPeriod(employeeId, targetDate, allData) {
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const empId = row[1] || '';
    const startDate = row[3] ? new Date(row[3]) : null;
    const endDate = row[4] ? new Date(row[4]) : null;
    const periodType = row[6] || '';
    
    if (empId.toString() === employeeId.toString() && 
        periodType === 'temp' && 
        startDate && endDate &&
        targetDate >= startDate && targetDate <= endDate) {
      return true;
    }
  }
  return false;
}

/**
 * Generate next sequential Period ID (W001, W002, etc.)
 */
function generateNextPeriodId() {
  try {
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return 'W001'; // Default if sheet doesn't exist
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    let maxNumber = 0;
    
    // Find the highest existing period number
    for (let i = 1; i < data.length; i++) {
      const periodId = data[i][0] || '';
      if (periodId.toString().startsWith('W')) {
        const numberPart = periodId.toString().substring(1);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) {
          maxNumber = number;
        }
      }
    }
    
    // Generate next ID with zero padding
    const nextNumber = maxNumber + 1;
    return 'W' + nextNumber.toString().padStart(3, '0');
    
  } catch (error) {
    console.error('Error generating period ID:', error);
    return 'W001';
  }
}

/**
 * Create new weekly holiday period
 */
function createWeeklyHolidayPeriod(employeeId, employeeName, startDate, endDate, periodType, holidayDaysStr, adminEmail) {
  try {
    console.log('📄 Creating weekly holiday period:', {
      employeeId, periodType, startDate, endDate, holidayDaysStr
    });
    
    // Convert all parameters to strings
    const employeeIdStr = String(employeeId || '');
    const employeeNameStr = String(employeeName || '');
    const startDateStr = String(startDate || '');
    const endDateStr = endDate ? String(endDate) : null;
    const periodTypeStr = String(periodType || '');
    const holidayDaysString = String(holidayDaysStr || '');
    const adminEmailStr = String(adminEmail || '');
    
    if (!adminEmailStr) {
      return {
        success: false,
        error: 'Admin email is required'
      };
    }
    
    // Validate required fields
    if (!employeeIdStr || !employeeNameStr || !startDateStr || !periodTypeStr || !holidayDaysString) {
      return {
        success: false,
        error: 'Missing required fields'
      };
    }
    
    // Validate period type
    if (periodTypeStr !== 'standard' && periodTypeStr !== 'temp') {
      return {
        success: false,
        error: 'Invalid period type. Must be "standard" or "temp"'
      };
    }
    
    // Validate temp period has end date
    if (periodTypeStr === 'temp' && !endDateStr) {
      return {
        success: false,
        error: 'Temporary periods must have an end date'
      };
    }
    
    // Validate period conflicts
    const validationResult = validateWeeklyHolidayPeriod(employeeIdStr, startDateStr, endDateStr, periodTypeStr, '');
    if (!validationResult.valid) {
      return {
        success: false,
        error: validationResult.error
      };
    }
    
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return {
        success: false,
        error: 'Weekly Holidays sheet not found'
      };
    }
    
    // Handle standard period replacement
    if (periodTypeStr === 'standard') {
      const updateResult = handleStandardPeriodReplacement(employeeIdStr, startDateStr);
      if (!updateResult.success) {
        console.error('Warning: Could not update existing standard period:', updateResult.error);
      }
    }
    
    // Generate new Period ID
    const periodId = generateNextPeriodId();
    
    // Parse holiday days - support BOTH formats: "Friday,Saturday" or "friday,saturday"
    const holidayDaysArray = holidayDaysString.split(',').map(day => day.trim()).filter(day => day !== '');
    
    // Map all possible formats to full capitalized names
    const dayMapping = {
      // Short names
      'Fri': 'Friday', 'Sat': 'Saturday', 'Sun': 'Sunday', 'Mon': 'Monday',
      'Tue': 'Tuesday', 'Wed': 'Wednesday', 'Thu': 'Thursday',
      // Lowercase full names
      'friday': 'Friday', 'saturday': 'Saturday', 'sunday': 'Sunday', 
      'monday': 'Monday', 'tuesday': 'Tuesday', 'wednesday': 'Wednesday', 'thursday': 'Thursday',
      // Lowercase short names
      'fri': 'Friday', 'sat': 'Saturday', 'sun': 'Sunday', 'mon': 'Monday',
      'tue': 'Tuesday', 'wed': 'Wednesday', 'thu': 'Thursday'
    };
    
    // Normalize to full capitalized day names
    const normalizedDays = holidayDaysArray.map(day => {
      // If it's already a capitalized full name, return it
      if (['Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'].includes(day)) {
        return day;
      }
      // Otherwise, try to map from any format
      return dayMapping[day] || dayMapping[day.toLowerCase()] || day;
    });
    
    // Create new row data
    const newRow = [
      periodId,                                    // A - Period ID
      employeeIdStr,                               // B - Employee ID
      employeeNameStr,                             // C - Employee Name
      new Date(startDateStr),                      // D - Start Date
      endDateStr ? new Date(endDateStr) : '',      // E - End Date
      'Active',                                    // F - Status
      periodTypeStr,                               // G - Period Type
      '', '', '', '', '', '', ''                   // H-N - Days (will be set below)
    ];
    
    // CORRECTED: Set holiday days using FULL day names to match columns
    // Columns H-N are: Fri, Sat, Sun, Mon, Tue, Wed, Thu (headers)
    // But we check against FULL names in our data
    const dayColumns = ['Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];
    for (let i = 0; i < dayColumns.length; i++) {
      newRow[7 + i] = normalizedDays.includes(dayColumns[i]) ? 'True' : '';
    }
    
    // Add to sheet
    const lastRow = weeklyHolidaysSheet.getLastRow();
    weeklyHolidaysSheet.getRange(lastRow + 1, 1, 1, newRow.length).setValues([newRow]);
    
    // Log admin action
    logAdminAction('Weekly Holiday Created', 
      `Period ID: ${periodId}, Employee: ${employeeNameStr} (${employeeIdStr}), Type: ${periodTypeStr}`, 
      adminEmailStr);
    
    console.log('✅ Weekly holiday period created successfully:', periodId);
    
    return {
      success: true,
      periodId: String(periodId),
      message: 'Weekly holiday period created successfully'
    };
    
  } catch (error) {
    console.error('Error creating weekly holiday period:', error);
    return {
      success: false,
      error: 'Error creating period: ' + String(error.toString())
    };
  }
}

/**
 * Handle standard period replacement - set end date on existing standard period
 */
function handleStandardPeriodReplacement(employeeId, newStartDate) {
  try {
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return { success: false, error: 'Sheet not found' };
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    const newStartDateObj = new Date(newStartDate);
    const endDate = new Date(newStartDateObj.getTime() - 24 * 60 * 60 * 1000); // Previous day
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = row[1] || '';
      const periodType = row[6] || '';
      const existingEndDate = row[4];
      
      if (empId.toString() === employeeId.toString() && 
          periodType === 'standard' && 
          !existingEndDate) { // Only update if no end date exists
        
        // Set end date to new start date - 1 day
        weeklyHolidaysSheet.getRange(i + 1, 5).setValue(endDate);
        
        console.log('✅ Updated existing standard period end date for employee:', employeeId);
        return { success: true };
      }
    }
    
    return { success: true }; // No existing standard period found, which is OK
    
  } catch (error) {
    console.error('Error handling standard period replacement:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Update existing weekly holiday period
 */
function updateWeeklyHolidayPeriod(periodId, employeeName, startDate, endDate, periodType, holidayDaysStr, adminEmail) {
  try {
    console.log('🔄 Updating weekly holiday period:', periodId);
    
    if (!adminEmail || !periodId) {
      return {
        success: false,
        error: 'Admin email and period ID are required'
      };
    }
    
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return {
        success: false,
        error: 'Weekly Holidays sheet not found'
      };
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    let foundRowIndex = -1;
    let employeeId = '';
    
    // Find the period to update
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === periodId.toString()) {
        foundRowIndex = i + 1; // Convert to 1-based index
        employeeId = data[i][1] || '';
        break;
      }
    }
    
    if (foundRowIndex === -1) {
      return {
        success: false,
        error: 'Period not found'
      };
    }
    
    // Validate period conflicts (excluding current period)
    const validationResult = validateWeeklyHolidayPeriod(employeeId, startDate, endDate, periodType, periodId);
    if (!validationResult.valid) {
      return {
        success: false,
        error: validationResult.error
      };
    }
    
    // Parse holiday days
    const holidayDays = holidayDaysStr.split(',').filter(day => day.trim() !== '');
    
    // Update the row
    weeklyHolidaysSheet.getRange(foundRowIndex, 3).setValue(employeeName);          // C - Employee Name
    weeklyHolidaysSheet.getRange(foundRowIndex, 4).setValue(new Date(startDate));   // D - Start Date
    weeklyHolidaysSheet.getRange(foundRowIndex, 5).setValue(endDate ? new Date(endDate) : ''); // E - End Date
    weeklyHolidaysSheet.getRange(foundRowIndex, 7).setValue(periodType);            // G - Period Type
    
    // Update holiday days (columns H-N)
    const dayColumns = ['Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];
    for (let i = 0; i < dayColumns.length; i++) {
      const value = holidayDays.includes(dayColumns[i]) ? 'True' : '';
      weeklyHolidaysSheet.getRange(foundRowIndex, 8 + i).setValue(value);
    }
    
    // Log admin action
    logAdminAction('Weekly Holiday Updated', 
      `Period ID: ${periodId}, Employee: ${employeeName}`, 
      adminEmail);
    
    console.log('✅ Weekly holiday period updated successfully:', periodId);
    
    return {
      success: true,
      message: 'Weekly holiday period updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating weekly holiday period:', error);
    return {
      success: false,
      error: 'Error updating period: ' + error.toString()
    };
  }
}

/**
 * Delete weekly holiday period
 */
function deleteWeeklyHolidayPeriod(periodId, adminEmail) {
  try {
    console.log('🗑️ Deleting weekly holiday period:', periodId);
    
    if (!adminEmail || !periodId) {
      return {
        success: false,
        error: 'Admin email and period ID are required'
      };
    }
    
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return {
        success: false,
        error: 'Weekly Holidays sheet not found'
      };
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    let foundRowIndex = -1;
    let employeeName = '';
    
    // Find the period to delete
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === periodId.toString()) {
        foundRowIndex = i + 1; // Convert to 1-based index
        employeeName = data[i][2] || '';
        break;
      }
    }
    
    if (foundRowIndex === -1) {
      return {
        success: false,
        error: 'Period not found'
      };
    }
    
    // Delete the row
    weeklyHolidaysSheet.deleteRow(foundRowIndex);
    
    // Log admin action
    logAdminAction('Weekly Holiday Deleted', 
      `Period ID: ${periodId}, Employee: ${employeeName}`, 
      adminEmail);
    
    console.log('✅ Weekly holiday period deleted successfully:', periodId);
    
    return {
      success: true,
      message: 'Weekly holiday period deleted successfully'
    };
    
  } catch (error) {
    console.error('Error deleting weekly holiday period:', error);
    return {
      success: false,
      error: 'Error deleting period: ' + error.toString()
    };
  }
}

/**
 * Validate weekly holiday period for conflicts
 */
function validateWeeklyHolidayPeriod(employeeId, startDate, endDate, periodType, excludePeriodId) {
  try {
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return { valid: true }; // If sheet doesn't exist, no conflicts
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    const startDateObj = new Date(startDate);
    const endDateObj = endDate ? new Date(endDate) : null;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const existingPeriodId = row[0] || '';
      const existingEmployeeId = row[1] || '';
      const existingStartDate = row[3] ? new Date(row[3]) : null;
      const existingEndDate = row[4] ? new Date(row[4]) : null;
      const existingPeriodType = row[6] || '';
      
      // Skip if different employee or same period (for updates)
      if (existingEmployeeId.toString() !== employeeId.toString() || 
          existingPeriodId.toString() === excludePeriodId.toString()) {
        continue;
      }
      
      // Check for conflicts based on period types
      if (periodType === 'standard') {
        // Standard periods: Allow replacement - just warn but don't block
        if (existingPeriodType === 'standard' && existingStartDate) {
          // If existing has no end date, it will be replaced
          if (!existingEndDate) {
            console.log('Standard period replacement will occur for employee:', employeeId);
            // This is OK - will be handled in createWeeklyHolidayPeriod
            continue;
          }
          
          // Check date overlap only if existing period has an end date
          if (startDateObj <= existingEndDate) {
            return {
              valid: false,
              error: `Standard period conflicts with existing standard period (${Utilities.formatDate(existingStartDate, Session.getScriptTimeZone(), 'MMM dd, yyyy')} - ${Utilities.formatDate(existingEndDate, Session.getScriptTimeZone(), 'MMM dd, yyyy')})`
            };
          }
        }
      } else if (periodType === 'temp') {
        // Temp periods cannot overlap with other temp periods
        if (existingPeriodType === 'temp' && existingStartDate && existingEndDate && endDateObj) {
          if (!(endDateObj < existingStartDate || startDateObj > existingEndDate)) {
            return {
              valid: false,
              error: `Temporary period conflicts with existing temporary period (${Utilities.formatDate(existingStartDate, Session.getScriptTimeZone(), 'MMM dd, yyyy')} - ${Utilities.formatDate(existingEndDate, Session.getScriptTimeZone(), 'MMM dd, yyyy')})`
            };
          }
        }
      }
    }
    
    return { valid: true };
    
  } catch (error) {
    console.error('Error validating weekly holiday period:', error);
    return {
      valid: false,
      error: 'Error validating period: ' + error.toString()
    };
  }
}

/**
 * Get active weekly holiday for specific employee on target date
 * UPDATED: Returns all values as strings
 */
function getActiveWeeklyHolidayForEmployee(employeeId, targetDateStr) {
  try {
    // Convert parameters to strings
    const employeeIdStr = String(employeeId || '');
    const targetDateString = String(targetDateStr || '');
    
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      return {
        success: false,
        error: 'Weekly Holidays sheet not found'
      };
    }
    
    const data = weeklyHolidaysSheet.getDataRange().getValues();
    const targetDate = new Date(targetDateString);
    targetDate.setHours(0, 0, 0, 0);
    
    let validStandardPeriod = null;
    let validTempPeriod = null;
    
    // Scan all rows to find valid periods for this employee on target date
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const empId = String(row[1] || '');
      const startDate = row[3] ? new Date(row[3]) : null;
      const endDate = row[4] ? new Date(row[4]) : null;
      const status = String(row[5] || '');
      const periodType = String(row[6] || '');
      
      // Only check for this employee
      if (empId !== employeeIdStr) continue;
      
      // Only check Active periods
      if (status !== 'Active') continue;
      
      // Set hours to compare dates properly
      if (startDate) startDate.setHours(0, 0, 0, 0);
      if (endDate) endDate.setHours(0, 0, 0, 0);
      
      // Check if target date falls within this period
      const isDateValid = startDate && 
                          targetDate >= startDate && 
                          (!endDate || targetDate <= endDate);
      
      if (!isDateValid) continue;
      
      // Collect holiday days - columns H-N (indices 7-13)
      const holidayDays = [];
      const dayNames = ['Friday', 'Saturday', 'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday'];
      
      for (let j = 0; j < dayNames.length; j++) {
        const cellValue = row[7 + j];
        if (cellValue === 'True' || cellValue === true || String(cellValue).toLowerCase() === 'true') {
          holidayDays.push(dayNames[j]);
        }
      }
      
      // Store the valid period based on type
      if (periodType === 'temp') {
        validTempPeriod = {
          periodId: String(row[0] || ''),
          periodType: 'temp',
          holidayDays: holidayDays
        };
      } else if (periodType === 'standard') {
        validStandardPeriod = {
          periodId: String(row[0] || ''),
          periodType: 'standard',
          holidayDays: holidayDays
        };
      }
    }
    
    // PRIORITY LOGIC: Temp overrides Standard
    let activePeriod = null;
    if (validTempPeriod) {
      activePeriod = validTempPeriod; // Temp takes priority
    } else if (validStandardPeriod) {
      activePeriod = validStandardPeriod; // Use standard if no temp
    }
    
    if (!activePeriod) {
      return {
        success: true,
        holidayDays: '' // Empty string if no period found
      };
    }
    
    // Return with all string values
    return {
      success: true,
      periodId: String(activePeriod.periodId),
      periodType: String(activePeriod.periodType),
      holidayDays: String(activePeriod.holidayDays.join(',')) // Convert array to comma-separated string
    };
    
  } catch (error) {
    console.error('Error getting active weekly holiday:', error);
    return {
      success: false,
      error: String(error.toString())
    };
  }
}

/**
 * Get count of active employees missing ongoing standard weekly holidays
 * Returns simple string format: "SUCCESS:count|totalEmployees|missingNames" or "ERROR:message"
 */
function getMissingStandardHolidaysCount() {
  try {
    console.log('=== GET MISSING STANDARD HOLIDAYS COUNT START ===');
    
    // Get all active employees using existing function
    const activeEmployeesResult = getActiveEmployees();
    console.log('📊 Active employees result:', activeEmployeesResult);
    
    if (!activeEmployeesResult || !activeEmployeesResult.startsWith('SUCCESS:')) {
      console.error('❌ Failed to get active employees:', activeEmployeesResult);
      return "ERROR:Failed to retrieve active employees";
    }
    
    // Parse active employees data
    const employeesData = activeEmployeesResult.replace('SUCCESS:', '');
    const activeEmployees = [];
    
    if (employeesData && employeesData.trim() !== '') {
      const employeeStrings = employeesData.split(';');
      employeeStrings.forEach(empStr => {
        if (empStr && empStr.trim() !== '') {
          const parts = empStr.split('|');
          if (parts.length >= 3) {
            activeEmployees.push({
              name: parts[0] || 'Unknown',
              email: parts[1] || '',
              id: parts[2] || ''
            });
          }
        }
      });
    }
    
    console.log(`📊 Found ${activeEmployees.length} active employees`);
    
    if (activeEmployees.length === 0) {
      return "SUCCESS:0|0|No active employees found";
    }
    
    // Get weekly holidays sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let weeklyHolidaysSheet = ss.getSheetByName('Weekly Holidays');
    if (!weeklyHolidaysSheet) {
      weeklyHolidaysSheet = ss.getSheetByName('weekly holidays');
    }
    if (!weeklyHolidaysSheet) {
      weeklyHolidaysSheet = ss.getSheetByName('WeeklyHolidays');
    }
    
    if (!weeklyHolidaysSheet) {
      console.error('❌ Weekly Holidays sheet not found');
      return "ERROR:Weekly Holidays sheet not found";
    }
    
    const holidayData = weeklyHolidaysSheet.getDataRange().getValues();
    console.log(`📊 Found ${holidayData.length - 1} holiday period records`);
    
    // Track employees with ongoing standard periods
    const employeesWithStandardHolidays = new Set();
    
    // Check each holiday period (skip header row)
    for (let i = 1; i < holidayData.length; i++) {
      const row = holidayData[i];
      
      const employeeId = row[1] ? row[1].toString().trim() : ''; // Column B: Employee ID
      const endDate = row[4]; // Column E: End Date
      const periodType = row[6] ? row[6].toString().trim().toLowerCase() : ''; // Column G: Period Type
      
      console.log(`🔍 Row ${i}: Employee ID: "${employeeId}", Period Type: "${periodType}", End Date: ${endDate}`);
      
      // Check if this is an ongoing standard period
      const isStandardPeriod = periodType === 'standard';
      const isOngoing = !endDate || endDate === '' || endDate === null;
      
      if (isStandardPeriod && isOngoing && employeeId) {
        employeesWithStandardHolidays.add(employeeId);
        console.log(`✅ Employee ${employeeId} has ongoing standard holiday`);
      }
    }
    
    console.log(`📊 ${employeesWithStandardHolidays.size} employees have ongoing standard holidays`);
    
    // Find employees missing ongoing standard holidays
    const missingEmployees = [];
    
    activeEmployees.forEach(employee => {
      const employeeId = employee.id.toString().trim();
      if (!employeesWithStandardHolidays.has(employeeId)) {
        missingEmployees.push(employee.name);
        console.log(`❌ Employee ${employeeId} (${employee.name}) missing ongoing standard holiday`);
      }
    });
    
    // Create simple string response: "SUCCESS:missingCount|totalEmployees|missingNames"
    const missingCount = missingEmployees.length;
    const totalEmployees = activeEmployees.length;
    const missingNames = missingEmployees.join(', ');
    
    const result = `SUCCESS:${missingCount}|${totalEmployees}|${missingNames}`;
    
    console.log(`📊 Final result: ${missingCount} missing out of ${totalEmployees} active employees`);
    console.log(`✅ Returning string result: ${result}`);
    console.log('=== GET MISSING STANDARD HOLIDAYS COUNT END ===');
    
    return result;
    
  } catch (error) {
    console.error('❌ ERROR in getMissingStandardHolidaysCount:', error);
    return "ERROR:Error analyzing standard holiday coverage - " + error.toString();
  }
}

/* ============================================================================
   LEAVE REQUESTS MANAGEMENT FUNCTIONS - ADD TO admin.gs
   Backend functions for handling leave request responses in Admin Portal
   ============================================================================ */

/**
 * Get all leave requests from Annual, Sick, and Emergency leave sheets
 * Returns combined data with filtering options
 */
function getAllLeaveRequests(filterType) {
  try {
    console.log('getAllLeaveRequests called with filter:', filterType);
    
    const sheets = [
      { name: 'Annual leaves', type: 'Annual', prefix: 'A', icon: 'fas fa-calendar-alt' },
      { name: 'Sick leaves', type: 'Sick', prefix: 'S', icon: 'fas fa-user-md' },
      { name: 'Emergency leaves', type: 'Emergency', prefix: 'E', icon: 'fas fa-exclamation-triangle' }
    ];
    
    const allRequests = [];
    let totalCount = 0;
    let respondedCount = 0;
    const typeCounts = {
      all: { total: 0, responded: 0 },
      annual: { total: 0, responded: 0 },
      sick: { total: 0, responded: 0 },
      emergency: { total: 0, responded: 0 }
    };
    
    for (const sheetInfo of sheets) {
      const sheet = getSheet(sheetInfo.name);
      if (!sheet) {
        console.warn(`Sheet '${sheetInfo.name}' not found`);
        continue;
      }
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue; // Skip if only headers or empty
      
      // Process each row (skip header row)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip empty rows
        if (!row[0] || !row[1]) continue;
        
        const responseStatus = row[10] || 'Pending'; // Column K
        const isResponded = responseStatus !== 'Pending';
        const startDate = row[3]; // Column D
        const endDate = row[4]; // Column E
        
        // Skip past/used requests (where end date has passed and status is not pending)
        const today = new Date();
        const requestEndDate = new Date(endDate);
        const isPastRequest = requestEndDate < today && isResponded;
        
        if (isPastRequest) continue;
        
        const request = {
          id: row[0] || '', // Column A - Request ID
          employeeId: row[1] || '', // Column B - Employee ID
          employeeName: row[2] || '', // Column C - Employee Name
          startDate: formatLongDateForDisplay(startDate), // Column D
          endDate: formatLongDateForDisplay(endDate), // Column E
          reason: row[5] || '', // Column F - Reason
          duration: row[6] || 0, // Column G - Duration in days
          weeklyHolidays: row[7] || 0, // Column H - Weekly holidays in period
          netDays: row[8] || 0, // Column I - Net leave days
          requestDate: formatDateTimeForDisplay(row[9]), // Column J - Request timestamp
          responseStatus: responseStatus, // Column K
          
          // Response details (Columns O-V)
          responseTimestamp: row[14] || '', // Column O
          respondedBy: row[15] || '', // Column P
          approvedDates: row[16] || '', // Column Q
          rejectedDates: row[17] || '', // Column R
          approvedHolidays: row[18] || 0, // Column S
          approvedDuration: row[19] || 0, // Column T
          rejectedDuration: row[20] || 0, // Column U
          durationUsed: row[21] || 0, // Column V
          
          // Additional info
          leaveType: sheetInfo.type,
          leaveTypeIcon: sheetInfo.icon,
          sheetName: sheetInfo.name,
          rowIndex: i + 1, // 1-based row index for updates
          isPending: responseStatus === 'Pending',
          isApproved: responseStatus.toLowerCase().includes('approved'),
          isRejected: responseStatus.toLowerCase().includes('rejected'),
          isPartial: responseStatus.toLowerCase().includes('partial')
        };
        
        // Count statistics
        totalCount++;
        typeCounts.all.total++;
        typeCounts[sheetInfo.type.toLowerCase()].total++;
        
        if (isResponded) {
          respondedCount++;
          typeCounts.all.responded++;
          typeCounts[sheetInfo.type.toLowerCase()].responded++;
        }
        
        // Apply filter
        if (!filterType || filterType === 'all' || filterType === sheetInfo.type.toLowerCase()) {
          allRequests.push(request);
        }
      }
    }
    
    // Sort by request date (newest first)
    allRequests.sort((a, b) => {
      const dateA = new Date(a.requestDate || 0);
      const dateB = new Date(b.requestDate || 0);
      return dateB.getTime() - dateA.getTime();
    });
    
    return {
      success: true,
      data: allRequests,
      counts: typeCounts,
      message: `Retrieved ${allRequests.length} leave requests`
    };
    
  } catch (error) {
    console.error('Error getting all leave requests:', error);
    return {
      success: false,
      error: 'Error retrieving leave requests: ' + error.toString(),
      data: [],
      counts: {}
    };
  }
}

/**
 * Get specific leave request by ID - ALL DATA AS STRINGS
 */
function getLeaveRequestById(requestId) {
  try {
    console.log('getLeaveRequestById called with ID:', requestId);
    
    const reqId = String(requestId || '').trim();
    if (!reqId) {
      return {
        success: false,
        error: 'Request ID is required'
      };
    }
    
    // Determine sheet based on request ID prefix
    let sheetName, leaveType, icon;
    const prefix = reqId.charAt(0).toUpperCase();
    
    switch (prefix) {
      case 'A':
        sheetName = 'Annual leaves';
        leaveType = 'Annual';
        icon = 'fas fa-calendar-alt';
        break;
      case 'S':
        sheetName = 'Sick leaves';
        leaveType = 'Sick';
        icon = 'fas fa-user-md';
        break;
      case 'E':
        sheetName = 'Emergency leaves';
        leaveType = 'Emergency';
        icon = 'fas fa-exclamation-triangle';
        break;
      default:
        return {
          success: false,
          error: 'Invalid request ID format'
        };
    }
    
    const sheet = getSheet(sheetName);
    if (!sheet) {
      return {
        success: false,
        error: `Sheet '${sheetName}' not found`
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find the request by ID
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === reqId) {
        const row = data[i];
        
        const request = {
          id: String(row[0] || ''),
          employeeId: String(row[1] || ''),
          employeeName: String(row[2] || ''),
          startDate: String(row[3] || ''),
          endDate: String(row[4] || ''),
          reason: String(row[5] || 'No reason provided'),
          duration: String(row[6] || '0'),
          weeklyHolidays: String(row[7] || '0'),
          netDays: String(row[8] || '0'),
          requestDate: String(row[9] || ''),
          responseStatus: String(row[10] || 'Pending'),
          responseTimestamp: String(row[14] || ''),
          respondedBy: String(row[15] || ''),
          approvedDates: String(row[16] || ''),
          rejectedDates: String(row[17] || ''),
          approvedHolidays: String(row[18] || '0'),
          approvedDuration: String(row[19] || '0'),
          rejectedDuration: String(row[20] || '0'),
          durationUsed: String(row[21] || '0'),
          leaveType: String(leaveType),
          leaveTypeIcon: String(icon),
          sheetName: String(sheetName),
          rowIndex: String(i + 1),
          startDateFormatted: formatLongDateForDisplay(row[3]),
          endDateFormatted: formatLongDateForDisplay(row[4]),
          requestDateFormatted: formatDateTimeForDisplay(row[9]),
          responseTimestampFormatted: row[14] ? formatDateTimeForDisplay(row[14]) : ''
        };
        
        console.log('Found request:', request);
        
        return {
          success: true,
          data: request
        };
      }
    }
    
    return {
      success: false,
      error: 'Request not found'
    };
    
  } catch (error) {
    console.error('Error getting leave request by ID:', error);
    return {
      success: false,
      error: 'Error retrieving request: ' + error.toString()
    };
  }
}


/**
 * Send leave response notification to employee
 */
function sendLeaveResponseNotification(requestId, adminEmail, sendNotification) {
  try {
    console.log('sendLeaveResponseNotification called:', { requestId, adminEmail, sendNotification });
    
    if (!sendNotification || sendNotification === 'false') {
      return {
        success: true,
        message: 'Response saved without notification'
      };
    }
    
    // Get updated request details
    const requestResult = getLeaveRequestById(requestId);
    if (!requestResult.success) {
      return requestResult;
    }
    
    const request = requestResult.data;
    
    // Get employee email
    const employeeEmail = getEmployeeEmailById(request.employeeId);
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email not found'
      };
    }
    
    // Get admin name
    const adminName = getAdminNameByEmail(adminEmail) || adminEmail;
    
    // Prepare notification data
    const statusData = {
      employeeEmail: employeeEmail,
      employeeName: request.employeeName,
      leaveType: request.leaveType,
      startDate: request.startDateFormatted,
      endDate: request.endDateFormatted,
      duration: request.duration,
      requestDate: request.requestDateFormatted,
      status: request.responseStatus,
      approvedDates: request.approvedDates,
      rejectedDates: request.rejectedDates,
      approvedDuration: request.approvedDuration,
      rejectedDuration: request.rejectedDuration,
      adminName: adminName,
      adminResponse: '', // Will be populated if admin comments exist
      actionDate: formatDateTimeForDisplay(new Date())
    };
    
    // Send notification using existing email function
    return sendLeaveStatusNotification(statusData);
    
  } catch (error) {
    console.error('Error sending leave response notification:', error);
    return {
      success: false,
      error: 'Error sending notification: ' + error.toString()
    };
  }
}

/**
 * Get dashboard counts for leave requests
 */
function getLeaveRequestCounts() {
  try {
    console.log('getLeaveRequestCounts called');
    
    const result = getAllLeaveRequests('all');
    if (!result.success) {
      return result;
    }
    
    return {
      success: true,
      counts: result.counts
    };
    
  } catch (error) {
    console.error('Error getting leave request counts:', error);
    return {
      success: false,
      error: 'Error retrieving counts: ' + error.toString(),
      counts: {}
    };
  }
}


/* ============================================================================
   LEAVE REQUESTS MANAGEMENT FUNCTIONS - COMPLETE RESTRUCTURE
   Backend functions for new day-by-day leave request system
   ============================================================================ */

/**
 * Get all leave requests with new day-by-day structure
 * Returns grouped requests with filtering options
 */
function getAllLeaveRequestsNew(filterType) {
  try {
    console.log('getAllLeaveRequestsNew called with filter:', filterType);
    
    const sheets = [
      { name: 'Annual leaves', type: 'Annual', prefix: 'A', icon: 'fas fa-calendar-alt' },
      { name: 'Sick leaves', type: 'Sick', prefix: 'S', icon: 'fas fa-user-md' },
      { name: 'Emergency leaves', type: 'Emergency', prefix: 'E', icon: 'fas fa-exclamation-triangle' }
    ];
    
    const allRequests = [];
    const requestsMap = new Map();
    let totalCount = 0;
    let respondedCount = 0;
    const typeCounts = {
      all: { total: 0, responded: 0 },
      annual: { total: 0, responded: 0 },
      sick: { total: 0, responded: 0 },
      emergency: { total: 0, responded: 0 }
    };
    
    for (const sheetInfo of sheets) {
      const sheet = getSheet(sheetInfo.name);
      if (!sheet) {
        console.warn(`Sheet '${sheetInfo.name}' not found`);
        continue;
      }
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue; // Skip if only headers or empty
      
      // Process each row (skip header row)
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        
        // Skip empty rows
        if (!row[0] || !row[1]) continue;
        
        const requestId = String(row[0] || '');
        const employeeId = String(row[1] || '');
        const employeeName = String(row[2] || '');
        const leaveDate = row[3]; // Column D - single date
        const weekDay = String(row[4] || ''); // Column E
        const reason = String(row[5] || '');
        const isWeeklyHoliday = String(row[6] || '').toLowerCase() === 'true';
        const requestTimestamp = row[7]; // Column H
        const dayResponseStatus = String(row[8] || 'Pending'); // Column I - Individual day response
        const responseTimestamp = row[9] || ''; // Column J
        const respondedBy = String(row[10] || ''); // Column K
        const notificationStatus = String(row[11] || 'Pending'); // Column L
        const adminComment = String(row[12] || ''); // Column M
        // Update used status in sheet if different
        const usedStatus = calculateUsedStatus(leaveDate, dayResponseStatus); // Pass response status
        if (String(row[13] || '') !== usedStatus) {
          sheet.getRange(i + 1, 14).setValue(usedStatus); // Column N
        }
        
        const dayData = {
          date: formatLongDateForDisplay(leaveDate),
          weekDay: weekDay,
          isWeeklyHoliday: isWeeklyHoliday,
          responseStatus: dayResponseStatus, // Use individual day response status
          rowIndex: i + 1,
          usedStatus: usedStatus
        };
        
        // Group by request ID
        if (!requestsMap.has(requestId)) {
          requestsMap.set(requestId, {
            id: requestId,
            employeeId: employeeId,
            employeeName: employeeName,
            reason: reason,
            requestTimestamp: formatDateTimeForDisplay(requestTimestamp),
            responseStatus: 'Pending', // Will be calculated after all days are processed
            responseTimestamp: responseTimestamp ? formatDateTimeForDisplay(responseTimestamp) : '',
            respondedBy: respondedBy,
            notificationStatus: notificationStatus,
            adminComment: adminComment,
            leaveType: sheetInfo.type,
            leaveTypeIcon: sheetInfo.icon,
            sheetName: sheetInfo.name,
            days: [],
            isPending: true, // Will be calculated after all days are processed
            isResponded: false, // Will be calculated after all days are processed
            employeeBalance: '' // Will be calculated after days are processed
          });
        }
        
        // Add day to request
        requestsMap.get(requestId).days.push(dayData);
      }
    }
    
    // Convert map to array and calculate overall status for each request
    for (const request of requestsMap.values()) {
      // Sort days by date
      request.days.sort((a, b) => new Date(a.date) - new Date(b.date));
      
      // Calculate available balance after days are processed
      if (request.days.length > 0) {
      const balanceData = getEmployeeAvailableBalance(
        request.employeeId, 
        request.leaveType, 
        request.days[0].date,
        request.id // Pass current request ID to exclude it
      );
              // Store detailed balance information
        request.employeeBalance = balanceData.available; // For backward compatibility
        request.employeeBalanceBase = balanceData.base;
        request.employeeBalanceOnHold = balanceData.onHold;
        request.employeeBalanceAvailable = balanceData.available;
      } else {
        request.employeeBalance = '0';
        request.employeeBalanceBase = '0';
        request.employeeBalanceOnHold = '0';
        request.employeeBalanceAvailable = '0';
      }
      
      // Calculate overall response status based on individual day responses
      const pendingDays = request.days.filter(day => day.responseStatus === 'Pending').length;
      const approvedDays = request.days.filter(day => day.responseStatus === 'Approved').length;
      const rejectedDays = request.days.filter(day => day.responseStatus === 'Rejected').length;
      const totalDays = request.days.length;
      
      if (pendingDays === totalDays) {
        request.responseStatus = 'Pending';
        request.isPending = true;
        request.isResponded = false;
      } else if (approvedDays === totalDays) {
        request.responseStatus = 'Approved';
        request.isPending = false;
        request.isResponded = true;
      } else if (rejectedDays === totalDays) {
        request.responseStatus = 'Rejected';
        request.isPending = false;
        request.isResponded = true;
      } else if (approvedDays > 0 || rejectedDays > 0) {
        request.responseStatus = 'Partially approved';
        request.isPending = false;
        request.isResponded = true;
      }
      
      // Count statistics
      totalCount++;
      typeCounts.all.total++;
      typeCounts[request.leaveType.toLowerCase()].total++;
      
      if (request.isResponded) {
        respondedCount++;
        typeCounts.all.responded++;
        typeCounts[request.leaveType.toLowerCase()].responded++;
      }
      
      // Calculate derived fields
      request.totalDuration = String(request.days.length);
      request.weeklyHolidaysCount = String(request.days.filter(day => day.isWeeklyHoliday).length);
      request.netDuration = String(request.days.length - parseInt(request.weeklyHolidaysCount));
      request.startDate = request.days[0]?.date || '';
      request.endDate = request.days[request.days.length - 1]?.date || '';
      
      // Calculate approved/rejected counts based on actual day responses
      const approvedDaysList = request.days.filter(day => day.responseStatus === 'Approved');
      const rejectedDaysList = request.days.filter(day => day.responseStatus === 'Rejected');
      
      request.approvedDuration = String(approvedDaysList.length);
      request.approvedWeeklyHolidays = String(approvedDaysList.filter(day => day.isWeeklyHoliday).length);
      request.netApprovedDuration = String(approvedDaysList.length - parseInt(request.approvedWeeklyHolidays));
      request.rejectedDuration = String(rejectedDaysList.length);
      
      // Calculate used duration
      const today = new Date();
      const usedDays = approvedDaysList.filter(day => {
        const dayDate = new Date(day.date);
        return dayDate <= today;
      });
      request.durationUsed = String(usedDays.length);
      
      // Apply filter
      if (!filterType || filterType === 'all' || filterType === request.leaveType.toLowerCase()) {
        allRequests.push(request);
      }
    }
    
    // Sort by request timestamp (newest first)
    // Sort requests with custom logic: Pending first, then responded, by start date within each group
    allRequests.sort((a, b) => {
      // Group 1: Pending requests (isPending = true)
      // Group 2: Responded requests (isPending = false)
      
      // First, sort by pending status (pending requests first)
      if (a.isPending !== b.isPending) {
        return a.isPending ? -1 : 1; // Pending (-1) comes before responded (1)
      }
      
      // Within the same group, sort by start date (ascending)
      const startDateA = new Date(a.startDate || 0);
      const startDateB = new Date(b.startDate || 0);
      
      return startDateA.getTime() - startDateB.getTime(); // Ascending order
    });
    
    return {
      success: true,
      data: allRequests,
      counts: typeCounts,
      message: `Retrieved ${allRequests.length} leave requests`
    };
    
  } catch (error) {
    console.error('Error getting all leave requests:', error);
    return {
      success: false,
      error: 'Error retrieving leave requests: ' + error.toString(),
      data: [],
      counts: {}
    };
  }
}

/**
 * Calculate overall status from individual day responses
 */
function calculateOverallStatusFromDays(requestId, allDayData) {
    const dayStatuses = allDayData.filter(day => day.requestId === requestId);
    
    if (dayStatuses.length === 0) return 'Pending';
    
    const approvedCount = dayStatuses.filter(day => day.responseStatus === 'Approved').length;
    const rejectedCount = dayStatuses.filter(day => day.responseStatus === 'Rejected').length;
    const pendingCount = dayStatuses.filter(day => day.responseStatus === 'Pending').length;
    const totalDays = dayStatuses.length;
    
    if (pendingCount === totalDays) {
        return 'Pending';
    } else if (approvedCount === totalDays) {
        return 'Approved';
    } else if (rejectedCount === totalDays) {
        return 'Rejected';
    } else if (approvedCount > 0) {
        return 'Partially approved';
    } else {
        return 'Pending';
    }
}


/**
 * Get specific leave request by ID with all days
 */
function getLeaveRequestByIdNew(requestId) {
  try {
    console.log('getLeaveRequestByIdNew called with ID:', requestId);
    
    const reqId = String(requestId || '').trim();
    if (!reqId) {
      return {
        success: false,
        error: 'Request ID is required'
      };
    }
    
    // Get all requests and find the specific one
    const allRequestsResult = getAllLeaveRequestsNew('all');
    if (!allRequestsResult.success) {
      return allRequestsResult;
    }
    
    const request = allRequestsResult.data.find(req => req.id === reqId);
    if (!request) {
      return {
        success: false,
        error: 'Request not found'
      };
    }
    
    // Get employee weekly holiday info
    const weeklyHolidayInfo = getEmployeeWeeklyHolidayInfo(request.employeeId);
    request.employeeWeeklyHoliday = weeklyHolidayInfo;
    
    return {
      success: true,
      data: request
    };
    
  } catch (error) {
    console.error('Error getting leave request by ID:', error);
    return {
      success: false,
      error: 'Error retrieving request: ' + error.toString()
    };
  }
}

/**
 * Update leave request response with new day-by-day approach
 */
function updateLeaveRequestResponseNew(requestId, dayResponses, adminComment, saveType, adminEmail) {
  try {
    console.log('updateLeaveRequestResponseNew called:', { requestId, saveType, adminEmail });
    
    if (!requestId || !dayResponses || !adminEmail) {
      return {
        success: false,
        error: 'Request ID, day responses, and admin email are required'
      };
    }
    
    // Get request details
    const requestResult = getLeaveRequestByIdNew(requestId);
    if (!requestResult.success) {
      return requestResult;
    }
    
    const request = requestResult.data;
    const sheet = getSheet(request.sheetName);
    if (!sheet) {
      return {
        success: false,
        error: `Sheet '${request.sheetName}' not found`
      };
    }
    
    const currentTimestamp = new Date();
    const adminCommentStr = String(adminComment || '');
    const notificationStatus = saveType === 'notify' ? 'Sent' : 'Pending';
    
    // Calculate overall response status
    const approvedCount = Object.values(dayResponses).filter(status => status === 'Approved').length;
    const rejectedCount = Object.values(dayResponses).filter(status => status === 'Rejected').length;
    const totalDays = Object.keys(dayResponses).length;
    
    let overallStatus = 'Pending';
    if (approvedCount === totalDays) {
      overallStatus = 'Approved';
    } else if (rejectedCount === totalDays) {
      overallStatus = 'Rejected';
    } else if (approvedCount > 0) {
      overallStatus = 'Partially approved';
    }
    
    // Update each day's response
    for (const day of request.days) {
      const dayDate = day.date;
      const dayResponse = dayResponses[dayDate];
      
      if (dayResponse && day.rowIndex) {
        const rowIndex = parseInt(day.rowIndex);
        
        // Update columns I, J, K, L, M
        sheet.getRange(rowIndex, 9).setValue(dayResponse); // Column I - Response status
        sheet.getRange(rowIndex, 10).setValue(currentTimestamp); // Column J - Response timestamp
        sheet.getRange(rowIndex, 11).setValue(adminEmail); // Column K - Responded by
        sheet.getRange(rowIndex, 12).setValue(notificationStatus); // Column L - Notification status
        sheet.getRange(rowIndex, 13).setValue(adminCommentStr); // Column M - Admin comment
        
        // Update Column N - Used status (auto-calculated)
        const usedStatus = calculateUsedStatus(new Date(dayDate));
        sheet.getRange(rowIndex, 14).setValue(usedStatus);
      }
    }
    
    // Log admin action
    logAdminAction('Leave Request Response', `${requestId}: ${overallStatus} by ${adminEmail}`);
    
    return {
      success: true,
      message: 'Response saved successfully',
      overallStatus: overallStatus,
      notificationStatus: notificationStatus
    };
    
  } catch (error) {
    console.error('Error updating leave request response:', error);
    return {
      success: false,
      error: 'Error saving response: ' + error.toString()
    };
  }
}

/**
 * Get employee weekly holiday information
 */
function getEmployeeWeeklyHolidayInfo(employeeId) {
  try {
    // Get employee's weekly holiday from Employees sheet
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) return { day: 'Friday', shortName: 'Fri' };
    
    const data = employeesSheet.getDataRange().getValues();
    let employeeWeeklyHoliday = 'Friday'; // Default
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(employeeId)) {
        employeeWeeklyHoliday = data[i][14] || 'Friday'; // Column O
        break;
      }
    }
    
    // Get short name for the day
    const dayMap = {
      'Sunday': 'Sun', 'Monday': 'Mon', 'Tuesday': 'Tue', 
      'Wednesday': 'Wed', 'Thursday': 'Thu', 'Friday': 'Fri', 'Saturday': 'Sat'
    };
    
    return {
      day: employeeWeeklyHoliday,
      shortName: dayMap[employeeWeeklyHoliday] || 'Fri'
    };
    
  } catch (error) {
    console.error('Error getting employee weekly holiday info:', error);
    return { day: 'Friday', shortName: 'Fri' };
  }
}

/**
 * Get dashboard counts for new leave requests system
 */
function getLeaveRequestCountsNew() {
  try {
    console.log('getLeaveRequestCountsNew called');
    
    const result = getAllLeaveRequestsNew('all');
    if (!result.success) {
      return result;
    }
    
    return {
      success: true,
      counts: result.counts
    };
    
  } catch (error) {
    console.error('Error getting leave request counts:', error);
    return {
      success: false,
      error: 'Error retrieving counts: ' + error.toString(),
      counts: {}
    };
  }
}

/**
 * Send leave response notification (NEW - ENHANCED)
 * REPLACE in admin.gs
 */
function sendLeaveResponseNotificationNew(requestId, adminEmail, shouldNotify) {
  try {
    console.log('sendLeaveResponseNotificationNew called:', { requestId, adminEmail, shouldNotify });
    
    if (shouldNotify !== 'true') {
      return { success: true, message: 'Notification skipped as requested' };
    }
    
    // Get request details
    const requestResult = getLeaveRequestByIdNew(requestId);
    if (!requestResult.success) {
      return requestResult;
    }
    
    const request = requestResult.data;
    
    // Get employee email
    const employeeEmail = getEmployeeEmailById(request.employeeId);
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email not found'
      };
    }
    
    // Get admin name
    const adminName = getAdminNameByEmail(adminEmail) || adminEmail;
    
    // Calculate approved and rejected dates strings
    const allApprovedDays = request.days.filter(d => d.responseStatus === 'Approved');
    const allRejectedDays = request.days.filter(d => d.responseStatus === 'Rejected');
    
    // Get initial balance (before response) - already available in request data
    const initialCurrent = String(request.employeeBalanceBase || '0');
    const initialOnHold = String(request.employeeBalanceOnHold || '0');
    const initialNet = String(request.employeeBalanceAvailable || '0');
    
    // Calculate approved net days for balance calculation
    const approvedDaysForBalance = request.days.filter(d => d.responseStatus === 'Approved');
    const approvedNetDaysCount = approvedDaysForBalance.filter(d => !d.isWeeklyHoliday).length;
    
    // Calculate final balance (after approval)
    let finalCurrent = initialCurrent;
    let finalOnHold = initialOnHold;
    let finalNet = initialNet;
    
    if (approvedDaysForBalance.length > 0) {
      // Balance After Approval Logic:
      // Current After = Initial Current (unchanged - no days used yet)
      // On-Hold After = Initial On-Hold + Approved Net Days
      // Net After = Current After - On-Hold After
      
      finalCurrent = initialCurrent; // Unchanged
      const finalOnHoldNum = parseInt(initialOnHold) + approvedNetDaysCount;
      finalOnHold = String(finalOnHoldNum);
      
      const finalNetNum = parseInt(finalCurrent) - finalOnHoldNum;
      finalNet = String(Math.max(0, finalNetNum));
    }
    
    // Convert all to strings
    const employeeEmailStr = String(employeeEmail);
    const employeeNameStr = String(request.employeeName || '');
    const leaveTypeStr = String(request.leaveType || '');
    const startDateStr = String(request.startDate || '');
    const endDateStr = String(request.endDate || '');
    const durationStr = String(request.duration || '0');
    const adminNameStr = String(adminName);
    const adminCommentsStr = String(request.adminComment || '');
    const actionDateStr = new Date().toISOString();
    const statusStr = String(request.responseStatus || 'Pending');
    const requestIdStr = String(requestId || '');
    
    const approvedDurationStr = String(allApprovedDays);
    const rejectedDurationStr = String(allRejectedDays);
    
    // Prepare days data array for email
    const daysDataForEmail = request.days.map(day => ({
      date: day.date,
      status: day.responseStatus || 'Pending',
      isWeeklyHoliday: day.isWeeklyHoliday || false
    }));
    
    // Call unified email function
    return sendLeaveResponse(
      employeeEmailStr,
      employeeNameStr,
      leaveTypeStr,
      requestIdStr,
      daysDataForEmail,
      adminNameStr,
      adminCommentsStr,
      actionDateStr,
      initialCurrent,
      initialOnHold,
      initialNet,
      finalCurrent,
      finalOnHold,
      finalNet
    );
    
  } catch (error) {
    console.error('Error sending leave response notification:', error);
    return {
      success: false,
      error: 'Error sending notification: ' + error.toString()
    };
  }
}


function validateLeaveRequestData() {
  try {
    const sheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    const validationResults = [];
    
    for (const sheetName of sheets) {
      const sheet = getSheet(sheetName);
      if (!sheet) {
        validationResults.push(`❌ Sheet '${sheetName}' not found`);
        continue;
      }
      
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        validationResults.push(`✅ Sheet '${sheetName}' - No data (OK)`);
        continue;
      }
      
      // Check column count (should be 14 columns A-N)
      const expectedColumns = 14;
      if (data[0].length < expectedColumns) {
        validationResults.push(`⚠️ Sheet '${sheetName}' - Missing columns (has ${data[0].length}, needs ${expectedColumns})`);
      } else {
        validationResults.push(`✅ Sheet '${sheetName}' - Column structure OK`);
      }
      
      // Check for required columns data
      let validRows = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[0] && row[1] && row[2] && row[3]) { // ID, Employee ID, Name, Date
          validRows++;
        }
      }
      
      validationResults.push(`📊 Sheet '${sheetName}' - ${validRows} valid data rows`);
    }
    
    return {
      success: true,
      validation: validationResults,
      message: 'Data validation completed'
    };
    
  } catch (error) {
    console.error('Error validating leave request data:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get employee's available balance considering on-hold approvals
 */
function getEmployeeAvailableBalance(employeeId, leaveType, requestStartDate, currentRequestId) {
  try {
    console.log(`Getting balance for employee ${employeeId}, type ${leaveType}, date ${requestStartDate}`);
    
    // Step 1: Get base balance from Balance-Calendar sheet
    const baseBalance = getBaseBalanceFromCalendar(employeeId, leaveType, requestStartDate);
    if (baseBalance === null) {
      console.warn(`Could not get base balance for employee ${employeeId}`);
      return {
        base: '5',
        onHold: '0', 
        available: '5'
      };
    }
    
    // Step 2: Calculate on-hold days from OTHER approved requests (exclude current request)
    const onHoldDays = calculateOnHoldDays(employeeId, leaveType, requestStartDate, currentRequestId);
    
    // Step 3: Calculate available balance
    const availableBalance = Math.max(0, baseBalance - onHoldDays);
    
    console.log(`Employee ${employeeId}: Base=${baseBalance}, OnHold=${onHoldDays}, Available=${availableBalance}`);
    
    return {
      base: String(baseBalance),
      onHold: String(onHoldDays),
      available: String(availableBalance)
    };
    
  } catch (error) {
    console.error('Error getting employee available balance:', error);
    return {
      base: '5',
      onHold: '0',
      available: '5'
    };
  }
}

/**
 * Get base balance from Balance-Calendar sheet
 */
function getBaseBalanceFromCalendar(employeeId, leaveType, requestStartDate) {
  try {
    const balanceSheet = getSheet('Balance - Calendar');
    if (!balanceSheet) {
      console.warn('Balance - Calendar sheet not found');
      return null;
    }
    
    // Get all data
    const data = balanceSheet.getDataRange().getValues();
    if (data.length < 5) return null; // Need at least headers + some data
    
    // Find employee columns (starting from column C = index 2)
    let employeeStartCol = -1;
    for (let col = 2; col < data[0].length; col += 3) { // Every 3 columns per employee
      if (String(data[0][col]) === String(employeeId)) {
        employeeStartCol = col;
        break;
      }
    }
    
    if (employeeStartCol === -1) {
      console.warn(`Employee ${employeeId} not found in Balance-Calendar`);
      return null;
    }
    
    // Determine leave type column offset
    let leaveTypeOffset = 0;
    if (leaveType === 'Sick') {
      leaveTypeOffset = 1;
    } else if (leaveType === 'Emergency') {
      leaveTypeOffset = 2;
    }
    
    const balanceCol = employeeStartCol + leaveTypeOffset;
    
    // Find date row (dates start from row 5 = index 4)
    const targetDate = new Date(requestStartDate);
    let dateRow = -1;
    
    for (let row = 4; row < data.length; row++) {
      const cellDate = new Date(data[row][1]); // Column B
      if (cellDate.getTime() === targetDate.getTime()) {
        dateRow = row;
        break;
      }
    }
    
    if (dateRow === -1) {
      console.warn(`Date ${requestStartDate} not found in Balance-Calendar`);
      return null;
    }
    
    // Get balance value
    const balance = data[dateRow][balanceCol];
    return typeof balance === 'number' ? balance : parseInt(balance) || 0;
    
  } catch (error) {
    console.error('Error getting base balance from calendar:', error);
    return null;
  }
}

/**
 * Calculate on-hold days from OTHER approved requests (exclude current request)
 * Only count net approved days (non-weekly holiday days) that are still future
 */
function calculateOnHoldDays(employeeId, leaveType, requestStartDate, excludeRequestId) {
  try {
    let totalOnHold = 0;
    const today = new Date();
    
    // Set time to midnight for accurate comparison
    today.setHours(0, 0, 0, 0);
    
    console.log(`Calculating on-hold for employee ${employeeId}, type ${leaveType}, excludeRequestId: ${excludeRequestId}`);
    console.log(`Today: ${today}`);
    
    // Get sheet name based on leave type
    let sheetName;
    switch (leaveType) {
      case 'Annual': sheetName = 'Annual leaves'; break;
      case 'Sick': sheetName = 'Sick leaves'; break;
      case 'Emergency': sheetName = 'Emergency leaves'; break;
      default: return 0;
    }
    
    const sheet = getSheet(sheetName);
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
    // Process each row to find approved days for this employee
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[0] || !row[1]) continue;
      
      const requestId = String(row[0] || ''); // Column A - Request ID
      const rowEmployeeId = String(row[1]); // Column B - Employee ID
      const leaveDate = row[3]; // Column D - Leave Date
      const isWeeklyHoliday = String(row[6] || '').toLowerCase() === 'true'; // Column G - Weekly Holiday
      const dayResponseStatus = String(row[8] || 'Pending'); // Column I - Response Status
      const currentUsedStatus = String(row[13] || ''); // Column N - Used status
      
      // Update Column N if needed (now considering response status)
      const calculatedUsedStatus = calculateUsedStatus(leaveDate, dayResponseStatus);
      if (currentUsedStatus !== calculatedUsedStatus) {
        sheet.getRange(i + 1, 14).setValue(calculatedUsedStatus); // Column N
        console.log(`Updated Column N for row ${i + 1}: ${calculatedUsedStatus} (status: ${dayResponseStatus})`);
      }
      
      // Skip current request - only count OTHER requests
      if (requestId === String(excludeRequestId)) {
        console.log(`Skipping current request: ${requestId}`);
        continue;
      }
      
      // Only count approved days for the same employee
      if (rowEmployeeId === String(employeeId) && dayResponseStatus === 'Approved') {
        const dayDate = new Date(leaveDate);
        dayDate.setHours(0, 0, 0, 0);
        
        console.log(`Found approved day: ${dayDate} for request ${requestId}, isWeeklyHoliday: ${isWeeklyHoliday}, usedStatus: ${calculatedUsedStatus}`);
        
        // Count only net approved days that are still future (NOT used)
        if (calculatedUsedStatus === 'Not yet' && !isWeeklyHoliday) {
          totalOnHold++;
          console.log(`Added to on-hold (net day): ${dayDate} (total now: ${totalOnHold})`);
        } else if (calculatedUsedStatus === 'Used') {
          console.log(`Day already used: ${dayDate}`);
        } else if (isWeeklyHoliday) {
          console.log(`Skipped weekly holiday: ${dayDate}`);
        }
      }
    }
    
    console.log(`Total on-hold NET days: ${totalOnHold}`);
    return totalOnHold;
    
  } catch (error) {
    console.error('Error calculating on-hold days:', error);
    return 0;
  }
}

/* ============================================================================
   PASSWORD MANAGEMENT - RESTRUCTURED STANDARDIZED IMPLEMENTATION
   ============================================================================ */

/* ---------------------------------------------------------------------------
   SECTION 1: SELF-PASSWORD CHANGE (LOGGED IN - ADMIN & EMPLOYEE)
   --------------------------------------------------------------------------- */

/**
 * Self-password change for logged-in admin (from admin portal)
 * No PIN needed, requires current password
 * @param {string} adminEmail - Admin email
 * @param {string} currentPassword - Current password
 * @param {string} newPassword - New password
 * @return {object} Result with success status
 */
function changeAdminPasswordLoggedIn(adminEmail, currentPassword, newPassword) {
  try {
    console.log('🔐 Admin changing own password (logged in):', adminEmail);
    
    if (!adminEmail || !currentPassword || !newPassword) {
      return {
        success: false,
        error: 'All fields are required'
      };
    }
    
    if (newPassword.length < 4) {
      return {
        success: false,
        error: 'New password must be at least 4 characters long'
      };
    }
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      return {
        success: false,
        error: 'Admins sheet not found'
      };
    }
    
    const data = adminsSheet.getDataRange().getValues();
    let targetRow = -1;
    let adminName = 'Administrator';
    
    // Find admin and verify current password
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      
      if (rowEmail === adminEmail.toLowerCase()) {
        const storedPassword = data[i][ADMIN_COLUMNS.PASSWORD] ? data[i][ADMIN_COLUMNS.PASSWORD].toString() : '';
        
        // Verify current password
        if (storedPassword !== currentPassword) {
          return {
            success: false,
            error: 'Current password is incorrect'
          };
        }
        
        targetRow = i + 1;
        adminName = data[i][ADMIN_COLUMNS.NAME] || 'Administrator';
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Administrator not found'
      };
    }
    
    // Update password
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.PASSWORD + 1).setValue(newPassword);
    console.log('✅ Password updated in sheet');
    
    // Send confirmation email
    const emailResult = sendPasswordChangeConfirmationEmail(adminEmail, adminName);
    console.log('📧 Confirmation email result:', emailResult);
    
    console.log('✅ Password changed successfully for:', adminEmail);
    return {
      success: true,
      message: 'Password changed successfully'
    };
    
  } catch (error) {
    console.error('❌ Error changing admin password:', error);
    return {
      success: false,
      error: 'Error changing password: ' + error.toString()
    };
  }
}

/**
 * Self-password change for logged-in employee (from employee portal)
 * No PIN needed, requires current password
 * @param {string} employeeEmail - Employee email
 * @param {string} currentPassword - Current password
 * @param {string} newPassword - New password
 * @return {object} Result with success status
 */
function changeEmployeePasswordLoggedIn(employeeEmail, currentPassword, newPassword) {
  try {
    console.log('🔐 Employee changing own password (logged in):', employeeEmail);
    
    if (!employeeEmail || !currentPassword || !newPassword) {
      return {
        success: false,
        error: 'All fields are required'
      };
    }

    if (newPassword.length < 4) {
      return {
        success: false,
        error: 'New password must be at least 4 characters long'
      };
    }

    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    let targetRow = -1;
    let employeeName = 'Employee';
    
    // Find employee and verify current password
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === employeeEmail.toLowerCase()) {
        
        // Check current password (Column G - index 6)
        const storedPassword = data[i][6] ? data[i][6].toString() : '';
        
        // Verify current password
        if (storedPassword !== currentPassword) {
          return {
            success: false,
            error: 'Current password is incorrect'
          };
        }
        
        targetRow = i + 1;
        employeeName = data[i][1] || 'Employee';
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }
    
    // Update password (Column G)
    employeesSheet.getRange(targetRow, 7).setValue(newPassword);
    console.log('✅ Password updated in sheet');
    
    // Send confirmation email
    const emailResult = sendPasswordChangeConfirmationEmail(employeeEmail, employeeName);
    console.log('📧 Confirmation email result:', emailResult);
    
    console.log('✅ Password changed successfully for:', employeeEmail);
    return {
      success: true,
      message: 'Password changed successfully'
    };
    
  } catch (error) {
    console.error('❌ Error changing employee password:', error);
    return {
      success: false,
      error: 'Error changing password: ' + error.toString()
    };
  }
}

/* ---------------------------------------------------------------------------
   SECTION 2: FORGOT PASSWORD (FROM HOMEPAGE WITH PIN)
   --------------------------------------------------------------------------- */

/**
 * Send PIN for forgot password (from homepage)
 * @param {string} email - User email
 * @param {string} userType - 'admin' or 'employee'
 * @return {object} Result with token
 */
function sendForgotPasswordPIN(email, userType) {
  try {
    console.log('📧 Sending forgot password PIN to:', email, 'Type:', userType);
    
    // Validate user exists and is active
    let userData = null;
    let userFound = false;
    
    if (userType === 'employee') {
      const employeesSheet = getSheet('Employees');
      if (employeesSheet) {
        const data = employeesSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const rowEmail = data[i][2]; // Column C (email)
          const removedDate = data[i][9]; // Column J (removed on)
          if (rowEmail && rowEmail.toString().toLowerCase() === email.toLowerCase() && !removedDate) {
            userData = {
              name: data[i][1] || 'Employee', // Column B (name)
              email: rowEmail
            };
            userFound = true;
            break;
          }
        }
      }
    } else if (userType === 'admin') {
      const adminsSheet = getSheet('Admins');
      if (adminsSheet) {
        const data = adminsSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const rowEmail = data[i][ADMIN_COLUMNS.EMAIL];
          const status = data[i][ADMIN_COLUMNS.STATUS];
          if (rowEmail && rowEmail.toString().toLowerCase() === email.toLowerCase() && status === 'active') {
            userData = {
              name: data[i][ADMIN_COLUMNS.NAME] || 'Administrator',
              email: rowEmail
            };
            userFound = true;
            break;
          }
        }
      }
    }
    
    if (!userFound) {
      return {
        success: false,
        error: 'User not found or inactive'
      };
    }
    
    // Generate 6-digit PIN
    const pin = Math.floor(100000 + Math.random() * 900000).toString();
    const token = Utilities.getUuid();
    const expirationTime = new Date().getTime() + (10 * 60 * 1000); // 10 minutes
    
    // Store PIN in Properties Service
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty(`forgot_pin_${token}`, pin);
    properties.setProperty(`forgot_exp_${token}`, expirationTime.toString());
    properties.setProperty(`forgot_email_${token}`, email.toLowerCase());
    properties.setProperty(`forgot_type_${token}`, userType);
    
    console.log('✅ PIN stored with token:', token);
    
    // Send PIN email
    const emailResult = sendPasswordResetPINEmail(userData.email, userData.name, pin);
    
    if (!emailResult.success) {
      // Clean up stored PIN if email fails
      properties.deleteProperty(`forgot_pin_${token}`);
      properties.deleteProperty(`forgot_exp_${token}`);
      properties.deleteProperty(`forgot_email_${token}`);
      properties.deleteProperty(`forgot_type_${token}`);
      
      return {
        success: false,
        error: 'Failed to send PIN email'
      };
    }
    
    return {
      success: true,
      token: token,
      message: 'PIN sent successfully'
    };
    
  } catch (error) {
    console.error('❌ Error sending forgot password PIN:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Reset password using PIN (from homepage modal)
 * @param {string} pin - 6-digit PIN
 * @param {string} token - Session token
 * @param {string} email - User email
 * @param {string} newPassword - New password
 * @param {string} userType - 'admin' or 'employee'
 * @return {object} Result
 */
function resetPasswordWithPIN(pin, token, email, newPassword, userType) {
  try {
    console.log('🔐 Resetting password with PIN for:', email);
    
    if (!pin || !token || !email || !newPassword || !userType) {
      return {
        success: false,
        error: 'Missing required parameters'
      };
    }
    
    const properties = PropertiesService.getScriptProperties();
    const storedPin = properties.getProperty(`forgot_pin_${token}`);
    const expirationTime = properties.getProperty(`forgot_exp_${token}`);
    const storedEmail = properties.getProperty(`forgot_email_${token}`);
    const storedType = properties.getProperty(`forgot_type_${token}`);
    
    // Verify PIN session exists
    if (!storedPin || !expirationTime || !storedEmail || !storedType) {
      return {
        success: false,
        error: 'PIN session not found or expired'
      };
    }
    
    // Check expiration
    if (new Date().getTime() > parseInt(expirationTime)) {
      // Clean up expired PIN
      properties.deleteProperty(`forgot_pin_${token}`);
      properties.deleteProperty(`forgot_exp_${token}`);
      properties.deleteProperty(`forgot_email_${token}`);
      properties.deleteProperty(`forgot_type_${token}`);
      
      return {
        success: false,
        error: 'PIN has expired'
      };
    }
    
    // Verify PIN and details match
    if (pin.toString() !== storedPin.toString() ||
        email.toLowerCase() !== storedEmail.toLowerCase() ||
        userType !== storedType) {
      return {
        success: false,
        error: 'Invalid PIN or session mismatch'
      };
    }
    
    // Update password in appropriate sheet
    let updateResult;
    
    if (userType === 'employee') {
      updateResult = updateEmployeePasswordDirect(email, newPassword);
    } else if (userType === 'admin') {
      updateResult = updateAdminPasswordDirect(email, newPassword);
    } else {
      return {
        success: false,
        error: 'Invalid user type'
      };
    }
    
    if (!updateResult.success) {
      return updateResult;
    }
    
    // Clean up PIN data
    properties.deleteProperty(`forgot_pin_${token}`);
    properties.deleteProperty(`forgot_exp_${token}`);
    properties.deleteProperty(`forgot_email_${token}`);
    properties.deleteProperty(`forgot_type_${token}`);
    
    console.log('✅ Password reset successfully');
    
    return {
      success: true,
      message: 'Password reset successfully'
    };
    
  } catch (error) {
    console.error('❌ Error resetting password with PIN:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ---------------------------------------------------------------------------
   SECTION 3: ADMIN-TRIGGERED PASSWORD RESET (WITH EMAIL)
   --------------------------------------------------------------------------- */

/**
 * Admin resets employee password (generates new password and emails it)
 * @param {string} employeeEmail - Employee email
 * @param {string} adminEmail - Admin performing the action
 * @return {object} Result with new password
 */
function adminResetEmployeePassword(employeeEmail, adminEmail) {
  try {
    console.log('🔐 Admin resetting employee password:', employeeEmail);
    
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    let targetRow = -1;
    let employeeName = 'Employee';
    let employeeId = '';
    
    // Find employee
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === employeeEmail.toLowerCase()) {
        targetRow = i + 1;
        employeeId = data[i][0] || '';
        employeeName = data[i][1] || 'Employee';
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Employee not found'
      };
    }
    
    // Generate new password using standardized format
    const newPassword = generateStandardPassword(employeeName);
    
    // Update password (Column G)
    employeesSheet.getRange(targetRow, 7).setValue(newPassword);
    console.log('✅ Password updated in sheet');
    
    // Get admin name for email
    const adminInfo = getAdminByEmail(adminEmail);
    const adminName = adminInfo ? adminInfo.name : 'Administrator';
    
    // Send password reset email
    const emailResult = sendEmployeePasswordResetEmail(
      employeeName,
      employeeEmail,
      employeeId,
      newPassword,
      adminName
    );
    
    console.log('📧 Password reset email result:', emailResult);
    
    return {
      success: true,
      message: 'Password reset successfully',
      newPassword: newPassword
    };
    
  } catch (error) {
    console.error('❌ Error resetting employee password:', error);
    return {
      success: false,
      error: 'Error resetting password: ' + error.toString()
    };
  }
}

/**
 * Admin resets another admin's password (generates new password and emails it)
 * @param {string} targetAdminEmail - Target admin email
 * @param {string} actingAdminEmail - Admin performing the action
 * @return {object} Result with new password
 */
function adminResetAdminPassword(targetAdminEmail, actingAdminEmail) {
  try {
    console.log('🔐 Admin resetting admin password:', targetAdminEmail);
    
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      return {
        success: false,
        error: 'Admins sheet not found'
      };
    }
    
    const data = adminsSheet.getDataRange().getValues();
    let targetRow = -1;
    let targetAdminName = 'Administrator';
    
    // Find target admin
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      
      if (rowEmail === targetAdminEmail.toLowerCase()) {
        targetRow = i + 1;
        targetAdminName = data[i][ADMIN_COLUMNS.NAME] || 'Administrator';
        break;
      }
    }
    
    if (targetRow === -1) {
      return {
        success: false,
        error: 'Administrator not found'
      };
    }
    
    // Generate new password using standardized format
    const newPassword = generateStandardPassword(targetAdminName);
    
    // Update password
    adminsSheet.getRange(targetRow, ADMIN_COLUMNS.PASSWORD + 1).setValue(newPassword);
    console.log('✅ Password updated in sheet');
    
    // Get acting admin name for email
    const actingAdminInfo = getAdminByEmail(actingAdminEmail);
    const actingAdminName = actingAdminInfo ? actingAdminInfo.name : 'Administrator';
    
    // Send password reset email
    const emailResult = sendAdminPasswordResetEmail(
      targetAdminEmail,
      targetAdminName,
      newPassword,
      actingAdminName
    );
    
    console.log('📧 Password reset email result:', emailResult);
    
    return {
      success: true,
      message: 'Password reset successfully',
      newPassword: newPassword
    };
    
  } catch (error) {
    console.error('❌ Error resetting admin password:', error);
    return {
      success: false,
      error: 'Error resetting password: ' + error.toString()
    };
  }
}

/* ---------------------------------------------------------------------------
   HELPER FUNCTIONS
   --------------------------------------------------------------------------- */

/**
 * STANDARDIZED PASSWORD GENERATION
 * Format: 2 lowercase letters (first letters of name OR initials) + 4 random digits
 * @param {string} name - User full name
 * @return {string} Generated password
 */
function generateStandardPassword(name) {
  try {
    // Remove spaces and special characters
    const cleanName = name.replace(/[^a-zA-Z\s]/g, '').trim();
    
    // Split into words
    const words = cleanName.split(/\s+/);
    
    let letters = '';
    
    if (words.length === 1) {
      // Single name: take first 2 letters
      letters = cleanName.substring(0, 2).toLowerCase();
    } else {
      // Multiple names: take initials (first letter of first 2 words)
      letters = (words[0].charAt(0) + words[1].charAt(0)).toLowerCase();
    }
    
    // Ensure we have 2 letters
    if (letters.length < 2) {
      letters = (letters + 'x').substring(0, 2);
    }
    
    // Generate 4 random digits
    const randomDigits = Math.floor(1000 + Math.random() * 9000).toString();
    
    // Combine
    const password = letters + randomDigits;
    
    console.log('🔑 Generated password for:', name);
    return password;
    
  } catch (error) {
    console.error('❌ Error generating password:', error);
    // Fallback to simple random password
    return 'pw' + Math.floor(1000 + Math.random() * 9000).toString();
  }
}

/**
 * Update employee password directly (used by forgot password flow)
 * @param {string} email - Employee email
 * @param {string} newPassword - New password
 * @return {object} Result
 */
function updateEmployeePasswordDirect(email, newPassword) {
  try {
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === email.toLowerCase()) {
        // Update password (Column G)
        employeesSheet.getRange(i + 1, 7).setValue(newPassword);
        console.log('✅ Employee password updated');
        return {
          success: true
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('❌ Error updating employee password:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Update admin password directly (used by forgot password flow)
 * @param {string} email - Admin email
 * @param {string} newPassword - New password
 * @return {object} Result
 */
function updateAdminPasswordDirect(email, newPassword) {
  try {
    const adminsSheet = getSheet('Admins');
    if (!adminsSheet) {
      return {
        success: false,
        error: 'Admins sheet not found'
      };
    }
    
    const data = adminsSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const rowEmail = data[i][ADMIN_COLUMNS.EMAIL] ? data[i][ADMIN_COLUMNS.EMAIL].toString().toLowerCase() : '';
      
      if (rowEmail === email.toLowerCase()) {
        // Update password
        adminsSheet.getRange(i + 1, ADMIN_COLUMNS.PASSWORD + 1).setValue(newPassword);
        console.log('✅ Admin password updated');
        return {
          success: true
        };
      }
    }
    
    return {
      success: false,
      error: 'Administrator not found'
    };
    
  } catch (error) {
    console.error('❌ Error updating admin password:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get calendar data for admin portal - ENHANCED VERSION
 * Returns calendar data for specified month/year
 * Only considers ACTIVE employees
 * Uses Cairo timezone for all dates
 * @param {string} monthStr - Month (1-12)
 * @param {string} yearStr - Year (e.g., "2025")
 * @return {string} SUCCESS:data OR ERROR:message
 */
function getAdminCalendarData(monthStr, yearStr) {
  try {
    console.log('=== GET ADMIN CALENDAR DATA START ===');
    console.log('Month:', monthStr, 'Year:', yearStr);
    
    const month = parseInt(monthStr);
    const year = parseInt(yearStr);
    
    if (isNaN(month) || isNaN(year) || month < 1 || month > 12) {
      return "ERROR:Invalid month or year";
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cairoTimeZone = 'Africa/Cairo';
    
    // Get system boundaries from Config sheet
    const configSheet = ss.getSheetByName('Config');
    if (!configSheet) {
      return "ERROR:Config sheet not found";
    }
    
    const systemStartDate = new Date(configSheet.getRange('B3').getValue());
    const systemEndDate = new Date(configSheet.getRange('B4').getValue());
    
    systemStartDate.setHours(0, 0, 0, 0);
    systemEndDate.setHours(0, 0, 0, 0);
    
    // Get first and last day of requested month
    const firstDay = new Date(year, month - 1, 1);
    const lastDay = new Date(year, month, 0);
    
    firstDay.setHours(0, 0, 0, 0);
    lastDay.setHours(0, 0, 0, 0);
    
    // Ensure dates are within system boundaries
    if (lastDay < systemStartDate || firstDay > systemEndDate) {
      return "ERROR:Requested month is outside system boundaries";
    }
    
    // Get ACTIVE employees list only
    const employeesSheet = ss.getSheetByName('Employees');
    if (!employeesSheet) {
      return "ERROR:Employees sheet not found";
    }
    
    const employeesData = employeesSheet.getDataRange().getValues();
    const activeEmployees = [];
    
    for (let i = 1; i < employeesData.length; i++) {
      const empId = String(employeesData[i][0] || '');
      const empName = String(employeesData[i][1] || '');
      
      if (!empId || !empName) continue;
      
      // Calculate status using enhanced logic - ENSURE STRING PARAMETERS
      const deactivatedOn = employeesData[i][9] ? String(employeesData[i][9]) : '';
      const reactivatedOn = employeesData[i][17] ? String(employeesData[i][17]) : '';
      const explicitStatus = employeesData[i][11] ? String(employeesData[i][11]) : 'Active';
      
      const calculatedStatus = calculateEmployeeStatusEnhanced(
        deactivatedOn,
        reactivatedOn,
        explicitStatus
      );
      
      if (calculatedStatus === 'Active') {
        activeEmployees.push({
          id: empId,
          name: empName
        });
      }
    }
    
    console.log('Total active employees:', activeEmployees.length);
    
    if (activeEmployees.length === 0) {
      return "ERROR:No active employees found";
    }
    
    // Get official holidays periods
    const officialHolidaysSheet = ss.getSheetByName('Official holidays');
    const holidayPeriods = {};
    
    if (officialHolidaysSheet) {
      const holidayData = officialHolidaysSheet.getRange(4, 1, officialHolidaysSheet.getLastRow() - 3, 2).getValues();
      
      holidayData.forEach(row => {
        if (row[0] && row[1]) {
          const holidayName = String(row[0]);
          const holidayDate = new Date(row[1]);
          holidayDate.setHours(0, 0, 0, 0);
          
          const dateKey = Utilities.formatDate(holidayDate, cairoTimeZone, 'yyyy-MM-dd');
          
          if (!holidayPeriods[holidayName]) {
            holidayPeriods[holidayName] = [];
          }
          
          holidayPeriods[holidayName].push(dateKey);
        }
      });
    }
    
    // Get Full-Calendar sheet data
    const fullCalendarSheet = ss.getSheetByName('Full-calendar');
    if (!fullCalendarSheet) {
      return "ERROR:Full-calendar sheet not found";
    }
    
    const lastRow = fullCalendarSheet.getLastRow();
    const lastCol = fullCalendarSheet.getLastColumn();
    
    if (lastRow < 5) {
      return "ERROR:Full-calendar sheet has insufficient data";
    }
    
    // Get employee headers (rows 1-3)
    const employeeIds = fullCalendarSheet.getRange(1, 3, 1, lastCol - 2).getValues()[0];
    const employeeNames = fullCalendarSheet.getRange(2, 3, 1, lastCol - 2).getValues()[0];
    const leaveTypes = fullCalendarSheet.getRange(3, 3, 1, lastCol - 2).getValues()[0];
    
    // Get all date data
    const dateData = fullCalendarSheet.getRange(5, 2, lastRow - 4, lastCol - 1).getValues();
    
    // Build calendar data structure
    const calendarDays = [];
    
    // Loop through each day of the month
    const currentDate = new Date(firstDay);
    
    while (currentDate <= lastDay) {
      const dateKey = Utilities.formatDate(currentDate, cairoTimeZone, 'yyyy-MM-dd');
      
      // Check if date is within system boundaries
      if (currentDate >= systemStartDate && currentDate <= systemEndDate) {
        
        // Find matching row in Full-Calendar
        let matchingRowIndex = -1;
        for (let i = 0; i < dateData.length; i++) {
          const rowDate = new Date(dateData[i][0]);
          rowDate.setHours(0, 0, 0, 0);
          
          if (rowDate.getTime() === currentDate.getTime()) {
            matchingRowIndex = i;
            break;
          }
        }
        
        // Build day data with detailed counts
        const dayData = {
          date: dateKey,
          dayNumber: currentDate.getDate(),
          dayName: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][currentDate.getDay()],
          isOfficialHoliday: false,
          officialHolidayName: '',
          employeesData: [],
          // Detailed counts for unclicked view
          counts: {
            totalActive: activeEmployees.length,
            working: 0,
            approvedLeaves: 0,
            pendingRequests: 0,
            weeklyHolidays: 0,
            officialHolidayOff: 0
          }
        };
        
        // Check if this date is part of an official holiday
        for (const [holidayName, dates] of Object.entries(holidayPeriods)) {
          if (dates.includes(dateKey)) {
            dayData.isOfficialHoliday = true;
            dayData.officialHolidayName = holidayName;
            break;
          }
        }
        
        // Process employee data if row exists
        if (matchingRowIndex >= 0) {
          const rowData = dateData[matchingRowIndex];
          
          for (let colIndex = 1; colIndex < rowData.length; colIndex++) {
            const empIndex = colIndex - 1;
            const empId = employeeIds[empIndex];
            const empName = employeeNames[empIndex];
            const leaveType = leaveTypes[empIndex];
            const value = rowData[colIndex];
            
            if (empId && empName) {
              // Only include ACTIVE employees
              const isActiveEmployee = activeEmployees.some(ae => ae.id === String(empId));
              
              if (isActiveEmployee) {
                const empData = {
                  id: String(empId),
                  name: String(empName),
                  type: String(leaveType),
                  value: String(value || '')
                };
                
                dayData.employeesData.push(empData);
                
                // Count for detailed statistics
                if (leaveType === 'Weekly Holidays' && (value === true || value === 'true')) {
                  dayData.counts.weeklyHolidays++;
                } else if (leaveType === 'Official Holidays' && (value === 'Off' || value === '')) {
                  // Only count if it's official holiday day
                  if (dayData.isOfficialHoliday) {
                    dayData.counts.officialHolidayOff++;
                  }
                } else if (['Annual Leaves', 'Sick Leaves', 'Emergency Leaves'].includes(leaveType) && value) {
                  const valueStr = String(value).toLowerCase();
                  if (valueStr.includes('approved')) {
                    dayData.counts.approvedLeaves++;
                  } else if (valueStr.includes('pending')) {
                    dayData.counts.pendingRequests++;
                  }
                }
              }
            }
          }
        }
        
        // Calculate working employees
        // Group by employee to avoid double counting
        const employeeStatusMap = {};
        
        activeEmployees.forEach(emp => {
          employeeStatusMap[emp.id] = { isOff: false };
        });
        
        dayData.employeesData.forEach(empData => {
          if (!employeeStatusMap[empData.id]) return;
          
          if (empData.type === 'Weekly Holidays' && (empData.value === 'true' || empData.value === true)) {
            employeeStatusMap[empData.id].isOff = true;
          } else if (empData.type === 'Official Holidays' && dayData.isOfficialHoliday && (empData.value === 'Off' || empData.value === '')) {
            employeeStatusMap[empData.id].isOff = true;
          } else if (['Annual Leaves', 'Sick Leaves', 'Emergency Leaves'].includes(empData.type)) {
            const valueStr = String(empData.value).toLowerCase();
            if (valueStr.includes('approved')) {
              employeeStatusMap[empData.id].isOff = true;
            }
          }
        });
        
        // Count working employees
        dayData.counts.working = Object.values(employeeStatusMap).filter(status => !status.isOff).length;
        
        // If official holiday, adjust officialHolidayOff count to exclude those already off for other reasons
        if (dayData.isOfficialHoliday) {
          let officialOnlyOffCount = 0;
          
          activeEmployees.forEach(emp => {
            const empDataList = dayData.employeesData.filter(ed => ed.id === emp.id);
            
            let hasWeeklyHoliday = false;
            let hasApprovedLeave = false;
            let hasOfficialHolidayOff = false;
            
            empDataList.forEach(ed => {
              if (ed.type === 'Weekly Holidays' && (ed.value === 'true' || ed.value === true)) {
                hasWeeklyHoliday = true;
              } else if (['Annual Leaves', 'Sick Leaves', 'Emergency Leaves'].includes(ed.type) && String(ed.value).toLowerCase().includes('approved')) {
                hasApprovedLeave = true;
              } else if (ed.type === 'Official Holidays' && (ed.value === 'Off' || ed.value === '')) {
                hasOfficialHolidayOff = true;
              }
            });
            
            // Only count if off for official holiday ONLY (not for other reasons)
            if (hasOfficialHolidayOff && !hasWeeklyHoliday && !hasApprovedLeave) {
              officialOnlyOffCount++;
            }
          });
          
          dayData.counts.officialHolidayOff = officialOnlyOffCount;
        }
        
        calendarDays.push(dayData);
      }
      
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    // Build response string
    const responseData = {
      month: month,
      year: year,
      systemStartDate: Utilities.formatDate(systemStartDate, cairoTimeZone, 'yyyy-MM-dd'),
      systemEndDate: Utilities.formatDate(systemEndDate, cairoTimeZone, 'yyyy-MM-dd'),
      totalActiveEmployees: activeEmployees.length,
      days: calendarDays
    };
    
    console.log('Calendar data built successfully');
    return "SUCCESS:" + JSON.stringify(responseData);
    
  } catch (error) {
    console.error('Error getting admin calendar data:', error);
    return "ERROR:" + error.toString();
  }
}

function calculateEmployeeStatusEnhanced(deactivatedOnStr, reactivatedOnStr, explicitStatusStr) {
  try {
    // Force convert to strings and handle all edge cases
    const deactivatedOn = deactivatedOnStr ? String(deactivatedOnStr) : '';
    const reactivatedOn = reactivatedOnStr ? String(reactivatedOnStr) : '';
    const explicitStatus = explicitStatusStr ? String(explicitStatusStr) : 'Active';
    
    // Default fallback
    const fallbackStatus = explicitStatus || 'Active';
    
    // If never deactivated, use explicit status or active
    if (!deactivatedOn || deactivatedOn === '' || deactivatedOn === 'null' || deactivatedOn === 'undefined') {
      return fallbackStatus;
    }
    
    // If no reactivation, employee is inactive
    if (!reactivatedOn || reactivatedOn === '' || reactivatedOn === 'null' || reactivatedOn === 'undefined') {
      return 'Inactive';
    }
    
    // Try to compare dates
    try {
      const deactivatedDate = new Date(deactivatedOn);
      const reactivatedDate = new Date(reactivatedOn);
      
      // Validate dates
      if (isNaN(deactivatedDate.getTime()) || isNaN(reactivatedDate.getTime())) {
        console.warn('Invalid date format, using fallback status');
        return fallbackStatus;
      }
      
      // Compare timestamps - latest action wins (>= means reactivation wins on same timestamp)
      return (reactivatedDate >= deactivatedDate) ? 'Active' : 'Inactive';
      
    } catch (dateError) {
      console.warn('Date parsing error, using fallback:', dateError);
      return fallbackStatus;
    }
    
  } catch (error) {
    console.error('Error calculating employee status:', error);
    return explicitStatusStr ? String(explicitStatusStr) : 'Active';
  }
}

/* ============================================================================
   BATCH INITIALIZATION - Add to admin.gs
   ============================================================================ */
function getAdminPortalInitData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    
    // 1. Get System Dates
    const startDate = configSheet.getRange('B3').getValue();
    const endDate = configSheet.getRange('B4').getValue();
    const dates = {
      startDate: startDate ? Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      endDate: endDate ? Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
    };

    // 2. Get Notification Preferences
    const daysValue = configSheet.getRange('B8').getValue();
    const prefs = {
      notifyEnabled: !!(daysValue && daysValue > 0),
      reminderDays: (daysValue && daysValue > 0) ? parseInt(daysValue) : 7
    };

    // 3. Get Badges (Pending Counts)
    // You can call your existing internal logic here or reimplement simplified version
    const employeesSheet = ss.getSheetByName('Employees');
    let pendingCount = 0;
    if (employeesSheet) {
        // Simple logic to count 'Pending' in Column I (index 8)
        const data = employeesSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][8] === 'Pending') pendingCount++;
        }
    }

    return {
      success: true,
      dates: dates,
      prefs: prefs,
      badges: {
        employees: pendingCount,
        requests: 0 // Add your logic for leave requests count if needed
      }
    };

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}
