
/* ============================================================================
   HOLIDAY MANAGEMENT SYSTEM - MAIN GOOGLE APPS SCRIPT
   Core setup, authentication, and data helper functions
   ============================================================================ */

/* ============================================================================
   CORE SETUP FUNCTIONS
   ============================================================================ */

/**
 * Update all used status columns before any user interaction
 * This runs automatically when the webapp is accessed
 */
function onWebAppLoad() {
  try {
    console.log('WebApp accessed - updating used status columns...');
    
    // Update all Column N values across leave sheets
    const result = updateAllUsedStatusColumns();
    
    if (result.success) {
      console.log(`Used status update completed: ${result.message}`);
    } else {
      console.error(`Used status update failed: ${result.error}`);
    }
    
    return result;
    
  } catch (error) {
    console.error('Error in onWebAppLoad:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Enhanced doGet function that updates used status before serving the page
 */
function doGet(e) {
  try {
    // Update used status columns before anything else
    onWebAppLoad();
    
    // Return the main HTML page - FIXED VERSION
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Leave Management System')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);  // ✅ ADD THIS LINE
    
  } catch (error) {
    console.error('Error in doGet:', error);
    
    // Return error page if something goes wrong
    const errorHtml = HtmlService.createHtmlOutput(`
      <html>
        <body>
          <h2>System Error</h2>
          <p>Unable to load the application. Please try again later.</p>
          <p>Error: ${error.toString()}</p>
        </body>
      </html>
    `);
    
    return errorHtml;
  }
}

/**
 * Include external files (CSS/JS)
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    console.error(`Error including file ${filename}:`, error);
    return `/* Error loading ${filename} */`;
  }
}

/* ============================================================================
   SHEET ACCESS HELPER
   ============================================================================ */

/**
 * Get sheet with error handling
 */
function getSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`Sheet '${sheetName}' not found`);
      return null;
    }
    
    return sheet;
  } catch (error) {
    console.error(`Error accessing sheet '${sheetName}':`, error);
    return null;
  }
}

/**
 * Validate that all required sheets exist
 */
function validateRequiredSheets() {
  try {
    const requiredSheets = ['Config', 'Admins', 'Employees', 'Annual leaves','official holidays', 'System versions'];
    const missingSheets = [];
    
    for (const sheetName of requiredSheets) {
      if (!getSheet(sheetName)) {
        missingSheets.push(sheetName);
      }
    }
    
    if (missingSheets.length > 0) {
      return {
        success: false,
        error: `Required sheets not found: ${missingSheets.join(', ')}. Please ensure all sheets exist.`,
        missingSheets: missingSheets
      };
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error validating required sheets:', error);
    return {
      success: false,
      error: 'Error validating system sheets: ' + error.toString()
    };
  }
}

/* ============================================================================
   STANDARDIZED ADMIN FUNCTIONS - REPLACE EXISTING FUNCTIONS
   ============================================================================ */

// **IMPORTANT: Add this constant at the top of both MAIN.GS and ADMIN.GS**
const ADMIN_COLUMNS = {
  NAME: 0,           // A - Admin Name [editable]
  EMAIL: 1,          // B - Admin Email [read-only in table]  
  PASSWORD: 2,       // C - Admin Password [invisible]
  FULL_PERMISSION: 3,// D - Full Permission (boolean) [editable]
  RECEIVE_REQUESTS: 4,// E - Receive Requests (boolean) [editable]
  DATE_CREATED: 5,   // F - Date Created (timestamp)
  CREATED_BY: 6,     // G - Created By [invisible]
  LAST_LOGIN: 7,     // H - Last Login [read-only]
  STATUS: 8,         // I - Status (active/inactive) [editable]
  STATUS_CHANGED_ON: 9 // J - Status Changed On [invisible]
};

/* ============================================================================
   SYSTEM VALIDATION
   ============================================================================ */

/**
 * Check if system is properly initialized
 */
function checkSystemInitialization() {
  return validateRequiredSheets();
}

function logAdminAction(action, details, adminEmail) {
  try {
    console.log('📝 Logging admin action:', action, 'by', adminEmail);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let adminLogSheet = ss.getSheetByName('Admin Actions');
    
    // Create Admin Actions sheet if it doesn't exist
    if (!adminLogSheet) {
      adminLogSheet = ss.insertSheet('Admin Actions');
      
      const headers = ['Timestamp', 'Admin Email', 'Action', 'Details'];
      adminLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      const headerRange = adminLogSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#2c3e50');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
    }
    
    // Get admin name from email
    let adminName = adminEmail;
    try {
      const adminInfo = getCurrentAdminInfo(adminEmail);
      if (adminInfo && adminInfo.name) {
        adminName = adminInfo.name;
      }
    } catch (e) {
      console.warn('Could not get admin name:', e);
    }
    
    // Add log entry
    const logData = [
      new Date(),
      adminEmail,  
      action,
      details
    ];
    
    adminLogSheet.appendRow(logData);
    
    // Format timestamp
    const lastRow = adminLogSheet.getLastRow();
    adminLogSheet.getRange(lastRow, 1).setNumberFormat('MM/dd/yyyy HH:mm:ss');
    
    console.log('✅ Admin action logged');
    
  } catch (error) {
    console.error('Error logging admin action:', error);
    // Don't fail the main operation for logging issues
  }
}

/* ============================================================================
   EMPLOYEE ID GENERATION SYSTEM
   ============================================================================ */

/**
 * Generate next sequential Employee ID starting from 101
 */
function generateNextEmployeeId() {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const lastRow = employeesSheet.getLastRow();
    
    if (lastRow <= 1) {
      return { success: true, id: 101 }; // Start with 101 if only header row exists
    }
    
    // Get all employee IDs from column A (excluding header)
    const idsRange = employeesSheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues().flat();
    
    // Filter out empty values and convert to integers
    const validIds = ids
      .filter(id => id !== '' && id !== null && id !== undefined)
      .map(id => parseInt(id))
      .filter(id => !isNaN(id) && id >= 101);
    
    if (validIds.length === 0) {
      return { success: true, id: 101 }; // Start with 101 if no valid IDs found
    }
    
    // Return the highest ID + 1
    return { success: true, id: Math.max(...validIds) + 1 };
    
  } catch (error) {
    console.error('Error generating employee ID:', error);
    return {
      success: false,
      error: 'Error generating employee ID: ' + error.toString()
    };
  }
}

/**
 * Validate Employee ID uniqueness
 */
function validateEmployeeIdUniqueness(employeeId, excludeRowIndex = null) {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        valid: false,
        error: 'Employees sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const lastRow = employeesSheet.getLastRow();
    if (lastRow <= 1) {
      return { valid: true };
    }
    
    const idsRange = employeesSheet.getRange(2, 1, lastRow - 1, 1);
    const ids = idsRange.getValues().flat();
    
    for (let i = 0; i < ids.length; i++) {
      const rowIndex = i + 2; // Account for 0-based index and header row
      
      // Skip the row we're excluding (for updates)
      if (excludeRowIndex && rowIndex === excludeRowIndex) {
        continue;
      }
      
      if (parseInt(ids[i]) === parseInt(employeeId)) {
        return {
          valid: false,
          error: `Employee ID ${employeeId} already exists`
        };
      }
    }
    
    return { valid: true };
    
  } catch (error) {
    console.error('Error validating employee ID uniqueness:', error);
    return {
      valid: false,
      error: 'Error validating employee ID uniqueness: ' + error.toString()
    };
  }
}


/* ============================================================================
   SYSTEM DATE MANAGEMENT
   ============================================================================ */

/**
 * Get current system dates
 */
function getCurrentSystemDates() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Get raw values from B3 and B4
    const startDateValue = configSheet.getRange('B3').getValue();
    const endDateValue = configSheet.getRange('B4').getValue();
    
    // Convert to simple strings for frontend
    let startDateString = 'Not set';
    let endDateString = 'Not set';
    
    if (startDateValue) {
      if (startDateValue instanceof Date) {
        // Format date as simple string
        startDateString = Utilities.formatDate(startDateValue, Session.getScriptTimeZone(), 'MMM dd, yyyy');
      } else {
        // If it's already a string/number, convert to string
        startDateString = startDateValue.toString();
      }
    }
    
    if (endDateValue) {
      if (endDateValue instanceof Date) {
        // Format date as simple string
        endDateString = Utilities.formatDate(endDateValue, Session.getScriptTimeZone(), 'MMM dd, yyyy');
      } else {
        // If it's already a string/number, convert to string
        endDateString = endDateValue.toString();
      }
    }
    
    // Return simple object with string values
    return {
      success: true,
      data: {
        startDate: startDateString,
        endDate: endDateString
      }
    };
    
  } catch (error) {
    return {
      success: false,
      error: 'Error retrieving system dates: ' + error.toString()
    };
  }
}

/**
 * Get system date constraints for form validation
 */
function getSystemDateConstraints() {
  try {
    const result = getCurrentSystemDates();
    
    if (!result.success) {
      return { startDate: null, endDate: null };
    }
    
    return {
      startDate: result.data.startDate ? result.data.startDate.toISOString().split('T')[0] : null,
      endDate: result.data.endDate ? result.data.endDate.toISOString().split('T')[0] : null
    };
    
  } catch (error) {
    console.error('Error getting system date constraints:', error);
    return { startDate: null, endDate: null };
  }
}

/**
 * Update system dates
 */
function updateSystemDates(startDateString, endDateString) {
  try {
    console.log('Updating system dates:', startDateString, endDateString);
    
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Simple validation - just check if strings are not empty
    if (!startDateString || !endDateString) {
      return {
        success: false,
        error: 'Start date and end date are required'
      };
    }
    
    // Try to parse dates
    let startDate, endDate;
    
    try {
      startDate = new Date(startDateString);
      endDate = new Date(endDateString);
    } catch (parseError) {
      return {
        success: false,
        error: 'Invalid date format: ' + parseError.toString()
      };
    }
    
    // Basic validation
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return {
        success: false,
        error: 'Invalid date values'
      };
    }
    
    if (startDate >= endDate) {
      return {
        success: false,
        error: 'End date must be after start date'
      };
    }
    
    // Update the cells - simple approach
    configSheet.getRange('B3').setValue(startDate);
    configSheet.getRange('B4').setValue(endDate);
    
    console.log('System dates updated successfully');
    
    return {
      success: true,
      message: 'System dates updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating system dates:', error);
    return {
      success: false,
      error: 'Error updating system dates: ' + error.toString()
    };
  }
}

/* ============================================================================
   AUTHENTICATION AND ADMIN MANAGEMENT
   ============================================================================ */

/**
 * Get current admin permissions
 */
function getCurrentAdminPermissions(adminEmail) {
  try {
    // Use provided admin email instead of Session.getActiveUser().getEmail()
    if (!adminEmail) {
      return {
        success: false,
        error: 'Admin email is required'
      };
    }
    
    console.log('🔍 Getting permissions for admin:', adminEmail);
    
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
      
      if (rowEmail === adminEmail.toLowerCase()) {
        const admin = {
          name: data[i][ADMIN_COLUMNS.NAME] || 'Administrator',
          email: data[i][ADMIN_COLUMNS.EMAIL],
          fullPermission: data[i][ADMIN_COLUMNS.FULL_PERMISSION] === 'Yes' || data[i][ADMIN_COLUMNS.FULL_PERMISSION] === true,
          receiveRequests: data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'Yes' || data[i][ADMIN_COLUMNS.RECEIVE_REQUESTS] === true,
          status: data[i][ADMIN_COLUMNS.STATUS] || 'active'
        };
        
        console.log('✅ Found admin permissions:', admin);
        
        return {
          success: true,
          admin: admin
        };
      }
    }
    
    return {
      success: false,
      error: 'Administrator not found'
    };
    
  } catch (error) {
    console.error('❌ Error getting current admin permissions:', error);
    return {
      success: false,
      error: 'Error retrieving admin permissions: ' + error.toString()
    };
  }
}


/**
 * Get all admins with full permission
 */
function getFullPermissionAdmins() {
  try {
    const adminsSheet = getSheet('Admins');
    
    if (!adminsSheet) {
      console.error('Admins sheet not found');
      return [];
    }
    
    const data = adminsSheet.getDataRange().getValues();
    const fullPermissionAdmins = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === 'Yes' || data[i][4] === true) {
        fullPermissionAdmins.push({
          email: data[i][1],
          name: data[i][2],
          receiveRequests: data[i][0] === 'Yes' || data[i][0] === true
        });
      }
    }
    
    return fullPermissionAdmins;
    
  } catch (error) {
    console.error('Error getting full permission admins:', error);
    return [];
  }
}

/**
 * Get all admins who should receive notifications
 */
function getNotificationAdmins() {
  try {
    const adminsSheet = getSheet('Admins');
    
    if (!adminsSheet) {
      console.error('Admins sheet not found');
      return [];
    }
    
    const data = adminsSheet.getDataRange().getValues();
    const notificationAdmins = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Yes' || data[i][0] === true) {
        notificationAdmins.push({
          email: data[i][1],
          name: data[i][2],
          fullPermission: data[i][4] === 'Yes' || data[i][4] === true
        });
      }
    }
    
    return notificationAdmins;
    
  } catch (error) {
    console.error('Error getting notification admins:', error);
    return [];
  }
}

/* ============================================================================
   UTILITY FUNCTIONS
   ============================================================================ */

/**
 * Generate secure password
 */
function generateSecurePassword(length = 8) {
  const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let password = '';
  
  for (let i = 0; i < length; i++) {
    password += charset.charAt(Math.floor(Math.random() * charset.length));
  }
  
  return password;
}

/**
 * Generate random token
 */
function generateToken(length = 32) {
  const charset = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let token = '';
  
  for (let i = 0; i < length; i++) {
    token += charset.charAt(Math.floor(Math.random() * charset.length));
  }
  
  return token;
}

/**
 * Validate email format
 */
function validateEmailFormat(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Safe sheet operation wrapper
 */
function safeSheetOperation(operation, errorMessage = 'Sheet operation failed') {
  try {
    return operation();
  } catch (error) {
    console.error(`${errorMessage}:`, error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   EMAIL SETTINGS BACKEND FUNCTIONS - ADD TO END OF NEW MAIN GS.txt
   ============================================================================ */

/**
 * Get email settings from Config sheet - FIXED version
 */
function getEmailSettings() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Read sender email from B2
    const senderEmail = configSheet.getRange('B2').getValue() || '';
    
    // Read other email settings from B13-B19
    const emailSettings = configSheet.getRange('B13:B19').getValues();
    
    return {
      success: true,
      data: {
        senderEmail: senderEmail, // B2
        autoReminderTime: formatTimeForInput(emailSettings[0][0]) || '', // B13
        holidayAnnouncementDays: emailSettings[1][0] || null, // B14
        autoCompensationDays: emailSettings[2][0] || null, // B15
        holidayWishesEnabled: emailSettings[3][0] === 'Yes' || emailSettings[3][0] === true, // B16
        holidayWishesDays: emailSettings[4][0] || null, // B17
        holidayEndReminderEnabled: emailSettings[5][0] === 'Yes' || emailSettings[5][0] === true, // B18
        holidayEndDays: emailSettings[6][0] || null // B19
      }
    };
    
  } catch (error) {
    console.error('Error getting email settings:', error);
    return {
      success: false,
      error: 'Error retrieving email settings: ' + error.toString()
    };
  }
}

/**
 * Format time value for HTML input (HH:MM format)
 */
function formatTimeForInput(timeValue) {
  try {
    if (!timeValue) return '';
    
    // If it's already a string in HH:MM format, return as-is
    if (typeof timeValue === 'string' && timeValue.match(/^\d{1,2}:\d{2}$/)) {
      return timeValue;
    }
    
    // If it's a Date object, extract time
    if (timeValue instanceof Date) {
      const hours = timeValue.getHours().toString().padStart(2, '0');
      const minutes = timeValue.getMinutes().toString().padStart(2, '0');
      return `${hours}:${minutes}`;
    }
    
    // If it's a decimal (0.375 = 9:00 AM), convert to time
    if (typeof timeValue === 'number' && timeValue < 1) {
      const totalMinutes = Math.round(timeValue * 24 * 60);
      const hours = Math.floor(totalMinutes / 60);
      const minutes = totalMinutes % 60;
      return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
    }
    
    return '';
  } catch (error) {
    console.error('Error formatting time:', error);
    return '';
  }
}

/**
 * Parse time input for storage in Google Sheets
 */
function parseTimeForStorage(timeInput) {
  try {
    if (!timeInput) return '';
    
    // If it's in HH:MM format, create a proper time
    if (typeof timeInput === 'string' && timeInput.match(/^\d{1,2}:\d{2}$/)) {
      const [hours, minutes] = timeInput.split(':');
      const timeDate = new Date();
      timeDate.setHours(parseInt(hours), parseInt(minutes), 0, 0);
      return timeDate;
    }
    
    return timeInput;
  } catch (error) {
    console.error('Error parsing time for storage:', error);
    return '';
  }
}
/**
 * Save email settings to Config sheet B13-B19
 */
function saveEmailSettings(settings) {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Save sender email to B2 - ADD THIS LINE
    configSheet.getRange('B2').setValue(settings.senderEmail || '');
    
    // Prepare values for B13-B19 (existing code)
    const values = [
      [parseTimeForStorage(settings.autoReminderTime) || ''], // B13
      [settings.holidayAnnouncementDays || ''], // B14
      [settings.autoCompensationDays || ''], // B15
      [settings.holidayWishesEnabled ? 'Yes' : 'No'], // B16
      [settings.holidayWishesDays || ''], // B17
      [settings.holidayEndReminderEnabled ? 'Yes' : 'No'], // B18
      [settings.holidayEndDays || ''] // B19
    ];
    
    // Write to Config sheet
    configSheet.getRange('B13:B19').setValues(values);
    
    // Log action
    // logAdminActionWithString('Email Settings Updated');
    
    return {
      success: true,
      message: 'Email settings saved successfully'
    };
    
  } catch (error) {
    console.error('Error saving email settings:', error);
    return {
      success: false,
      error: 'Error saving email settings: ' + error.toString()
    };
  }
}

/**
 * Check if compensation default is set (B12 dependency)
 */
function checkCompensationDefault() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: true,
        hasCompensationDefault: false
      };
    }
    
    // Check B12 value (Official holidays Compensation default)
    const compensationDefault = configSheet.getRange('B12').getValue();
    const hasValue = compensationDefault && compensationDefault > 0;
    
    return {
      success: true,
      hasCompensationDefault: hasValue,
      compensationValue: compensationDefault || 0
    };
    
  } catch (error) {
    console.error('Error checking compensation default:', error);
    return {
      success: true,
      hasCompensationDefault: false
    };
  }
}

/* ============================================================================
   LEAVES & HOLIDAYS BACKEND FUNCTIONS - ADD TO END OF NEW MAIN GS.txt
   ============================================================================ */

/**
 * Get leave default balances from Config sheet B9-B11
 */
function getLeaveDefaults() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Read values from B9-B11
    const annualDefault = configSheet.getRange('B9').getValue() || 0;
    const sickDefault = configSheet.getRange('B10').getValue() || 0;
    const emergencyDefault = configSheet.getRange('B11').getValue() || 0;
    
    return {
      success: true,
      data: {
        annual: annualDefault,
        sick: sickDefault,
        emergency: emergencyDefault
      }
    };
    
  } catch (error) {
    console.error('Error getting leave defaults:', error);
    return {
      success: false,
      error: 'Error retrieving leave defaults: ' + error.toString()
    };
  }
}

/**
 * Save leave default balances to Config sheet B9-B11
 */
function saveLeaveDefaults(defaults) {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Save values to B9-B11
    configSheet.getRange('B9').setValue(defaults.annual || 0);
    configSheet.getRange('B10').setValue(defaults.sick || 0);
    configSheet.getRange('B11').setValue(defaults.emergency || 0);
    
    // Log action
    // logAdminAction('Leave Defaults Updated');
    
    return {
      success: true,
      message: 'Leave defaults saved successfully'
    };
    
  } catch (error) {
    console.error('Error saving leave defaults:', error);
    return {
      success: false,
      error: 'Error saving leave defaults: ' + error.toString()
    };
  }
}

/**
 * Get holiday compensation setting from Config sheet B12
 */
function getHolidayCompensation() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const compensation = configSheet.getRange('B12').getValue() || 0;
    
    return {
      success: true,
      compensation: compensation
    };
    
  } catch (error) {
    console.error('Error getting holiday compensation:', error);
    return {
      success: false,
      error: 'Error retrieving holiday compensation: ' + error.toString()
    };
  }
}

/**
 * Save holiday compensation to Config sheet B12
 */
function saveHolidayCompensation(compensation) {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    configSheet.getRange('B12').setValue(compensation || 0);
    
    // Log action
    // logAdminAction('Holiday Compensation Updated');
    
    return {
      success: true,
      message: 'Holiday compensation saved successfully'
    };
    
  } catch (error) {
    console.error('Error saving holiday compensation:', error);
    return {
      success: false,
      error: 'Error saving holiday compensation: ' + error.toString()
    };
  }
}

/**
 * Add new official holiday
 */
/**
 * Add new official holiday with new sheet structure
 * @param {string} holidayName - Unique holiday name
 * @param {string} startDateString - Start date as string (YYYY-MM-DD)
 * @param {string} endDateString - End date as string (YYYY-MM-DD)
 * @param {string} adminEmail - Admin email for logging
 */
function addOfficialHoliday(holidayName, startDateString, endDateString, adminEmail) {
  try {
    console.log('Adding official holiday:', holidayName, startDateString, endDateString);
    
    if (!holidayName || !startDateString || !endDateString || !adminEmail) {
      return {
        success: false,
        error: 'Missing required parameters'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let holidaysSheet = ss.getSheetByName('Official holidays');
    
    // Create sheet if it doesn't exist
    if (!holidaysSheet) {
      holidaysSheet = ss.insertSheet('Official holidays');
      
      // Set up basic structure (headers will be in row 3)
      holidaysSheet.getRange('A3').setValue('Holiday Name');
      holidaysSheet.getRange('B3').setValue('Date');
    }
    
    // Check if holiday name already exists (must be unique)
    const existingData = holidaysSheet.getDataRange().getValues();
    for (let i = 3; i < existingData.length; i++) { // Start from row 4 (index 3)
      if (existingData[i][0] && existingData[i][0].toString().toLowerCase() === holidayName.toLowerCase()) {
        return {
          success: false,
          error: 'Holiday name already exists. Holiday names must be unique.'
        };
      }
    }
    
    // Convert string dates to Date objects
    const startDate = new Date(startDateString);
    const endDate = new Date(endDateString);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return {
        success: false,
        error: 'Invalid date format'
      };
    }
    
    if (startDate > endDate) {
      return {
        success: false,
        error: 'Start date cannot be after end date'
      };
    }
    
    // Generate rows for each day in the holiday period
    const currentDate = new Date(startDate);
    const rows = [];
    
    while (currentDate <= endDate) {
      rows.push([
        holidayName,
        new Date(currentDate)
      ]);
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    // Add all rows to the sheet
    const startRow = holidaysSheet.getLastRow() + 1;
    if (rows.length > 0) {
      holidaysSheet.getRange(startRow, 1, rows.length, 2).setValues(rows);
      
      // Format date column
      holidaysSheet.getRange(startRow, 2, rows.length, 1).setNumberFormat('MM/dd/yyyy');
    }
    
    console.log('Holiday added successfully:', holidayName);
    
    return {
      success: true,
      message: 'Holiday added successfully'
    };
    
  } catch (error) {
    console.error('Error adding official holiday:', error);
    return {
      success: false,
      error: 'Error adding holiday: ' + error.toString()
    };
  }
}

/**
 * Get all official holidays - SIMPLIFIED with string dates
 */
function getOfficialHolidays() {
  try {
    console.log('Getting official holidays with new structure...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    if (!holidaysSheet) {
      return {
        success: true,
        data: []
      };
    }
    
    const lastRow = holidaysSheet.getLastRow();
    if (lastRow <= 3) { // No data rows (header is row 3)
      return {
        success: true,
        data: []
      };
    }
    
    // Get data starting from row 4
    const data = holidaysSheet.getRange(4, 1, lastRow - 3, 2).getValues();
    
    // Group holidays by name
    const holidayGroups = {};
    
    for (let i = 0; i < data.length; i++) {
      const holidayName = data[i][0];
      const holidayDate = data[i][1];
      
      if (!holidayName || !holidayDate) continue;
      
      if (!holidayGroups[holidayName]) {
        holidayGroups[holidayName] = [];
      }
      
      holidayGroups[holidayName].push(new Date(holidayDate));
    }
    
    // Convert groups to consolidated format
    const holidays = [];
    
    for (const [name, dates] of Object.entries(holidayGroups)) {
      if (dates.length === 0) continue;
      
      // Sort dates
      dates.sort((a, b) => a.getTime() - b.getTime());
      
      const startDate = dates[0];
      const endDate = dates[dates.length - 1];
      
      // Convert to string format for frontend
      const startDateStr = startDate.getFullYear() + '-' + 
                          String(startDate.getMonth() + 1).padStart(2, '0') + '-' + 
                          String(startDate.getDate()).padStart(2, '0');
                          
      const endDateStr = endDate.getFullYear() + '-' + 
                        String(endDate.getMonth() + 1).padStart(2, '0') + '-' + 
                        String(endDate.getDate()).padStart(2, '0');
      
      holidays.push({
        name: name,
        startDate: startDateStr,
        endDate: endDateStr,
        days: dates.length
      });
    }
    
    console.log('Retrieved', holidays.length, 'holidays');
    
    return {
      success: true,
      data: holidays
    };
    
  } catch (error) {
    console.error('Error getting official holidays:', error);
    return {
      success: false,
      error: 'Error retrieving holidays: ' + error.toString()
    };
  }
}

/**
 * Delete official holiday with new sheet structure
 * Deletes all rows with matching holiday name
 * @param {string} holidayName - Holiday name to delete
 * @param {string} adminEmail - Admin email for logging
 */
function deleteOfficialHoliday(holidayName, adminEmail) {
  try {
    console.log('Deleting official holiday:', holidayName);
    
    if (!holidayName || !adminEmail) {
      return {
        success: false,
        error: 'Missing required parameters'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    if (!holidaysSheet) {
      return {
        success: false,
        error: 'Official holidays sheet not found'
      };
    }
    
    const lastRow = holidaysSheet.getLastRow();
    if (lastRow <= 3) {
      return {
        success: false,
        error: 'No holidays found'
      };
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    // Find all rows with matching holiday name (starting from row 4, index 3)
    for (let i = 3; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === holidayName.toLowerCase()) {
        rowsToDelete.push(i + 1); // Convert to 1-based row numbers
      }
    }
    
    if (rowsToDelete.length === 0) {
      return {
        success: false,
        error: 'Holiday not found'
      };
    }
    
    // Delete rows in reverse order to maintain row numbers
    rowsToDelete.reverse();
    for (const rowNum of rowsToDelete) {
      holidaysSheet.deleteRow(rowNum);
    }
    
    console.log('Holiday deleted successfully:', holidayName, '(' + rowsToDelete.length + ' rows)');
    
    return {
      success: true,
      message: 'Holiday deleted successfully'
    };
    
  } catch (error) {
    console.error('Error deleting official holiday:', error);
    return {
      success: false,
      error: 'Error deleting holiday: ' + error.toString()
    };
  }
}

/**
 * Update official holiday with new sheet structure
 * Updates all rows with matching holiday name
 * @param {string} oldHolidayName - Current holiday name to find
 * @param {string} newHolidayName - New unique holiday name
 * @param {string} startDateString - New start date as string
 * @param {string} endDateString - New end date as string
 * @param {string} adminEmail - Admin email for logging
 */
function updateOfficialHoliday(oldHolidayName, newHolidayName, startDateString, endDateString, adminEmail) {
  try {
    console.log('Updating official holiday:', oldHolidayName, 'to', newHolidayName);
    
    if (!oldHolidayName || !newHolidayName || !startDateString || !endDateString || !adminEmail) {
      return {
        success: false,
        error: 'Missing required parameters'
      };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    if (!holidaysSheet) {
      return {
        success: false,
        error: 'Official holidays sheet not found'
      };
    }
    
    // Check if new holiday name already exists (if different from old name)
    if (oldHolidayName.toLowerCase() !== newHolidayName.toLowerCase()) {
      const existingData = holidaysSheet.getDataRange().getValues();
      for (let i = 3; i < existingData.length; i++) {
        if (existingData[i][0] && existingData[i][0].toString().toLowerCase() === newHolidayName.toLowerCase()) {
          return {
            success: false,
            error: 'New holiday name already exists. Holiday names must be unique.'
          };
        }
      }
    }
    
    // Convert string dates to Date objects
    const startDate = new Date(startDateString);
    const endDate = new Date(endDateString);
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return {
        success: false,
        error: 'Invalid date format'
      };
    }
    
    if (startDate > endDate) {
      return {
        success: false,
        error: 'Start date cannot be after end date'
      };
    }
    
    // First, delete all rows with old holiday name
    const deleteResult = deleteOfficialHoliday(oldHolidayName, adminEmail);
    if (!deleteResult.success) {
      return deleteResult;
    }
    
    // Then, add the updated holiday
    const addResult = addOfficialHoliday(newHolidayName, startDateString, endDateString, adminEmail);
    if (!addResult.success) {
      return addResult;
    }
    
    console.log('Holiday updated successfully:', newHolidayName);
    
    return {
      success: true,
      message: 'Holiday updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating official holiday:', error);
    return {
      success: false,
      error: 'Error updating holiday: ' + error.toString()
    };
  }
}

/**
 * Check if a date falls within any official holiday
 */
function isOfficialHoliday(checkDate) {
  try {
    const holidays = getOfficialHolidays();
    
    if (!holidays.success || holidays.data.length === 0) {
      return false;
    }
    
    const targetDate = new Date(checkDate);
    
    for (const holiday of holidays.data) {
      const startDate = new Date(holiday.startDate);
      const endDate = new Date(holiday.endDate);
      
      if (targetDate >= startDate && targetDate <= endDate) {
        return {
          isHoliday: true,
          holidayName: holiday.name,
          startDate: holiday.startDate,
          endDate: holiday.endDate
        };
      }
    }
    
    return false;
    
  } catch (error) {
    console.error('Error checking if date is official holiday:', error);
    return false;
  }
}

/* ============================================================================
   SENDER EMAIL BACKEND FUNCTIONS - ADD TO END OF NEW MAIN GS.txt
   ============================================================================ */

/**
 * Get Gmail aliases for sender selection
 */
function getGmailAliases() {
  try {
    console.log('Getting Gmail aliases...');
    
    const aliases = GmailApp.getAliases();
    const currentUser = Session.getActiveUser().getEmail();
    
    console.log('Gmail aliases found:', aliases.length);
    console.log('Current user email:', currentUser);
    
    const aliasData = [];
    
    // Add primary email first
    aliasData.push({
      email: currentUser,
      name: 'Primary Account',
      isPrimary: true
    });
    
    // Add other aliases
    aliases.forEach(alias => {
      if (alias !== currentUser) { // Avoid duplicates
        aliasData.push({
          email: alias,
          name: alias,
          isPrimary: false
        });
      }
    });
    
    console.log('Processed aliases:', aliasData.length);
    
    return {
      success: true,
      data: aliasData
    };
    
  } catch (error) {
    console.error('Error getting Gmail aliases:', error);
    return {
      success: false,
      error: 'Error retrieving Gmail aliases: ' + error.toString(),
      data: []
    };
  }
}

/**
 * Get current sender email from Config B2
 */
/**
 * Get current sender email from Config B2 - ADD this function if missing
 */
function getSenderEmail() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const senderEmail = configSheet.getRange('B2').getValue() || '';
    
    return {
      success: true,
      data: {
        senderEmail: senderEmail
      }
    };
    
  } catch (error) {
    console.error('Error getting sender email:', error);
    return {
      success: false,
      error: 'Error retrieving sender email: ' + error.toString()
    };
  }
}

/**
 * Set sender email in Config B2
 */
function setSenderEmail(email) {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found. Please ensure the sheet exists.'
      };
    }
    
    // Validate email if provided
    if (email && !validateEmailFormat(email)) {
      return {
        success: false,
        error: 'Invalid email format'
      };
    }
    
    // Save to Config B2
    configSheet.getRange('B2').setValue(email || '');
    
    // Log action
    logAdminAction('Sender Email Updated', email || 'Reset to default');
    
    return {
      success: true,
      message: 'Sender email saved successfully'
    };
    
  } catch (error) {
    console.error('Error setting sender email:', error);
    return {
      success: false,
      error: 'Error saving sender email: ' + error.toString()
    };
  }
}

/* ============================================================================
   AUTHENTICATION FUNCTIONS FOR LOGIN SYSTEM - ADD TO END OF NEW MAIN GS.txt
   ============================================================================ */

/**
 * Get active employees for login dropdown
 */
function getActiveEmployees() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return "ERROR:No spreadsheet found";
    }
    
    let employeesSheet = ss.getSheetByName('Employees');
    if (!employeesSheet) {
      employeesSheet = ss.getSheetByName('employees');
    }
    if (!employeesSheet) {
      return "ERROR:Employees sheet not found";
    }
    
    const data = employeesSheet.getDataRange().getValues();
    const activeEmployees = [];
    
    for (let i = 1; i < data.length; i++) {
      const deactivatedOn = data[i][9];  // Column J - Deactivated on
      const reactivatedOn = data[i][17]; // Column R - Reactivated on
      const name = data[i][1] || 'Unknown';
      const email = data[i][2] || '';
      const id = data[i][0] || '';
      
      if (!email) continue; // Skip if no email
      
      let isActive = true;
      
      // If never deactivated, employee is active
      if (!deactivatedOn) {
        isActive = true;
      }
      // If deactivated but never reactivated, employee is inactive
      else if (deactivatedOn && !reactivatedOn) {
        isActive = false;
      }
      // If both dates exist, compare timestamps
      else if (deactivatedOn && reactivatedOn) {
        isActive = new Date(reactivatedOn) > new Date(deactivatedOn);
      }
      
      if (isActive) {
        activeEmployees.push(`${name}|${email}|${id}`);
      }
    }
    
    if (activeEmployees.length === 0) {
      return "EMPTY:No active employees";
    }
    
    return "SUCCESS:" + activeEmployees.join(';');
    
  } catch (error) {
    return "ERROR:Error retrieving employees";
  }
}

/**
 * Get active admins - RETURNS STRING  
 */
function getActiveAdmins() {
  try {
    console.log('=== GET ACTIVE ADMINS START ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error('❌ No spreadsheet found');
      return "ERROR:No spreadsheet found";
    }
    
    let adminsSheet = ss.getSheetByName('Admins');
    if (!adminsSheet) {
      adminsSheet = ss.getSheetByName('admins');
    }
    if (!adminsSheet) {
      adminsSheet = ss.getSheetByName('Admin');
    }
    if (!adminsSheet) {
      console.error('❌ No admins sheet found');
      return "ERROR:Admins sheet not found";
    }
    
    console.log('✅ Admins sheet found:', adminsSheet.getName());
    
    const data = adminsSheet.getDataRange().getValues();
    console.log('📊 Total rows in sheet:', data.length);
    
    if (data.length <= 1) {
      console.log('⚠️ No admin data found');
      return "EMPTY:No active administrators";
    }
    
    const activeAdmins = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      console.log(`🔍 Processing admin row ${i}:`, row);
      
      // Use standardized column mapping
      const name = row[ADMIN_COLUMNS.NAME] ? row[ADMIN_COLUMNS.NAME].toString().trim() : 'Unknown';
      const email = row[ADMIN_COLUMNS.EMAIL] ? row[ADMIN_COLUMNS.EMAIL].toString().trim() : '';
      const status = row[ADMIN_COLUMNS.STATUS] ? row[ADMIN_COLUMNS.STATUS].toString().trim() : '';
      
      console.log(`   - Name: "${name}" (Column ${ADMIN_COLUMNS.NAME})`);
      console.log(`   - Email: "${email}" (Column ${ADMIN_COLUMNS.EMAIL})`);
      console.log(`   - Status: "${status}" (Column ${ADMIN_COLUMNS.STATUS})`);
      
      const isActive = status && status.toLowerCase().trim() === 'active';
      console.log(`   - Is Active: ${isActive}`);
      
      if (isActive && email && email.includes('@')) {
        activeAdmins.push(`${name}|${email}`);
        console.log(`   ✅ Added active admin: ${name} (${email})`);
      } else {
        console.log(`   ❌ Skipped admin - Active: ${isActive}, Valid Email: ${email && email.includes('@')}`);
      }
    }
    
    console.log('📋 Total active admins found:', activeAdmins.length);
    
    if (activeAdmins.length === 0) {
      return "EMPTY:No active administrators";
    }
    
    const result = "SUCCESS:" + activeAdmins.join(';');
    console.log('✅ Returning result:', result);
    console.log('=== GET ACTIVE ADMINS END ===');
    
    return result;
    
  } catch (error) {
    console.error('❌ ERROR in getActiveAdmins:', error);
    return "ERROR:Error retrieving administrators - " + error.toString();
  }
}

/**
 * Validate admin login credentials - FIXED  
 */
function validateAdminCredentials(email, password) {
  try {
    console.log('=== VALIDATE ADMIN CREDENTIALS START ===');
    console.log('🔐 Login attempt for email:', email);
    
    if (!email || !password) {
      console.error('❌ Missing email or password');
      return "ERROR:Email and password are required";
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      console.error('❌ No spreadsheet found');
      return "ERROR:No spreadsheet found";
    }
    
    let adminsSheet = ss.getSheetByName('Admins');
    if (!adminsSheet) {
      adminsSheet = ss.getSheetByName('admins');
    }
    if (!adminsSheet) {
      adminsSheet = ss.getSheetByName('Admin');
    }
    
    if (!adminsSheet) {
      console.error('❌ No admins sheet found');
      return "ERROR:Admins sheet not found";
    }
    
    console.log('✅ Admins sheet found:', adminsSheet.getName());
    
    const data = adminsSheet.getDataRange().getValues();
    console.log('📊 Total rows in sheet:', data.length);
    
    if (data.length <= 1) {
      console.error('❌ No administrators found');
      return "ERROR:No administrators found";
    }
    
    // Find admin by email using standardized mapping
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      console.log(`🔍 Checking admin row ${i}:`, row);
      
      // Use standardized column mapping
      const rowEmail = row[ADMIN_COLUMNS.EMAIL] ? row[ADMIN_COLUMNS.EMAIL].toString().toLowerCase().trim() : '';
      const adminName = row[ADMIN_COLUMNS.NAME] ? row[ADMIN_COLUMNS.NAME].toString().trim() : 'Administrator';
      const storedPassword = row[ADMIN_COLUMNS.PASSWORD] ? row[ADMIN_COLUMNS.PASSWORD].toString().trim() : '';
      const status = row[ADMIN_COLUMNS.STATUS] ? row[ADMIN_COLUMNS.STATUS].toString().toLowerCase().trim() : '';
      
      console.log(`   - Email: "${rowEmail}" (Column ${ADMIN_COLUMNS.EMAIL})`);
      console.log(`   - Name: "${adminName}" (Column ${ADMIN_COLUMNS.NAME})`);
      console.log(`   - Has Password: ${storedPassword ? 'YES' : 'NO'} (Column ${ADMIN_COLUMNS.PASSWORD})`);
      console.log(`   - Status: "${status}" (Column ${ADMIN_COLUMNS.STATUS})`);
      
      if (rowEmail === email.toLowerCase().trim()) {
        console.log('✅ Email match found!');
        
        const isActive = !status || status.includes('active') || status === '';
        console.log(`   - Is Active: ${isActive}`);
        
        if (!isActive && status) {
          console.error('❌ Administrator account is deactivated');
          return "ERROR:Administrator account has been deactivated";
        }
        
        if (!storedPassword) {
          console.error('❌ No password set');
          return "ERROR:No password set for this administrator";
        }
        
        console.log('🔐 Verifying password...');
        if (storedPassword === password.toString().trim()) {
          console.log('✅ Password match! Login successful');
          
          // Update last login using standardized mapping
          try {
            adminsSheet.getRange(i + 1, ADMIN_COLUMNS.LAST_LOGIN + 1).setValue(new Date());
            console.log('✅ Last login updated');
          } catch (e) {
            console.warn('⚠️ Could not update last login:', e.toString());
          }
          
          // Extract permissions using standardized mapping
          const fullPermission = row[ADMIN_COLUMNS.FULL_PERMISSION] === 'Yes' || 
                                row[ADMIN_COLUMNS.FULL_PERMISSION] === true || 
                                row[ADMIN_COLUMNS.FULL_PERMISSION] === 'TRUE';
          const receiveRequests = row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'Yes' || 
                                row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === true || 
                                row[ADMIN_COLUMNS.RECEIVE_REQUESTS] === 'TRUE';
          
          console.log(`   - Full Permission: ${fullPermission} (Column ${ADMIN_COLUMNS.FULL_PERMISSION})`);
          console.log(`   - Receive Requests: ${receiveRequests} (Column ${ADMIN_COLUMNS.RECEIVE_REQUESTS})`);
          
          const result = `SUCCESS:${adminName}|${email}|${fullPermission}|${receiveRequests}`;
          console.log('✅ Returning successful result:', result);
          console.log('=== VALIDATE ADMIN CREDENTIALS END ===');
          
          return result;
        } else {
          console.error('❌ Password does not match');
          return "ERROR:Invalid password";
        }
      }
    }
    
    console.error('❌ Administrator not found');
    return "ERROR:Administrator not found";
    
  } catch (error) {
    console.error('❌ ERROR in validateAdminCredentials:', error);
    return "ERROR:System error occurred - " + error.toString();
  }
}


/* ============================================================================
   EMPLOYEE DATA MANAGEMENT - UPDATED FOR 19-COLUMN STRUCTURE (A-S)
   Column Mapping:
   A=ID, B=Name, C=E-mail, D=Opening Balance, E=Usable from, F=Usable to, 
   G=Password, H=Added on, I=Email status, J=Deactivated on, K=Last Login, 
   L=Active/Inactive, M=Opening emergency leaves, N=Opening Sick Leaves, 
   O=Weekly Holiday, P=Added by, Q=Deactivated BY, R=Reactivated on, S=Reactivated by
   ============================================================================ */

/**
 * SIMPLE STRING VERSION - REPLACE existing getAllEmployees function in main.gs
 * Returns simple string format instead of complex objects
 */
function getAllEmployees() {
  try {
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return "ERROR:Employees sheet not found";
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return "SUCCESS:EMPTY"; // No employees found
    }
    
    const employeeStrings = [];
    
    // Process each employee row (starting from row 2)
    for (let i = 1; i < data.length; i++) {
      try {
        // Safely extract each field with defaults and convert to strings
        const id = String(data[i][0] || '');
        const name = String(data[i][1] || '');
        const email = String(data[i][2] || '');
        const openingBalance = String(data[i][3] || 0);
        const usableFrom = String(data[i][4] || '');
        const usableTo = String(data[i][5] || '');
        const addedOn = String(data[i][7] || '');
        const emailStatus = String(data[i][8] || 'Pending');
        const deactivatedOn = String(data[i][9] || '');
        const lastLogin = String(data[i][10] || '');
        const emergencyLeaves = String(data[i][12] || 0);
        const sickLeaves = String(data[i][13] || 0);
        // Column O (data[i][14]) is IGNORED - we fetch from Weekly Holidays sheet
        const addedBy = String(data[i][15] || '');
        const reactivatedOn = String(data[i][17] || '');
        
        // Fetch weekly holidays from "Weekly Holidays" sheet
        let weeklyHolidayDisplay = '';
        const idStr = String(id);
        const todayStr = String(new Date().toISOString());
        const weeklyHolidayResult = getActiveWeeklyHolidayForEmployee(idStr, todayStr);
        
        if (weeklyHolidayResult.success && weeklyHolidayResult.holidayDays) {
          weeklyHolidayDisplay = String(weeklyHolidayResult.holidayDays); // Comma-separated string like "Friday,Saturday"
        }
        
        // Create simple pipe-separated string for each employee (all strings)
        const employeeString = `${id}|${name}|${email}|${openingBalance}|${usableFrom}|${usableTo}|${addedOn}|${emailStatus}|${deactivatedOn}|${lastLogin}|${emergencyLeaves}|${sickLeaves}|${weeklyHolidayDisplay}|${addedBy}|${reactivatedOn}`;
        
        employeeStrings.push(employeeString);
        
      } catch (rowError) {
        console.error('Error processing employee row', i, ':', rowError);
        // Skip this row and continue
      }
    }
    
    if (employeeStrings.length === 0) {
      return "SUCCESS:EMPTY";
    }
    
    // Return success with employee data separated by semicolons
    return "SUCCESS:" + employeeStrings.join(';');
    
  } catch (error) {
    console.error('Error in getAllEmployees:', error);
    return "ERROR:System error - " + error.toString();
  }
}

/**
 * Get current employee data for logged-in employee - Updated for 19-column structure
 */
function getCurrentEmployeeData() {
  try {
    const currentEmail = employeeEmail;
    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    // Updated column mapping for 19-column structure
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === currentEmail.toLowerCase()) {
        // Check if employee is deactivated
        if (data[i][9]) { // Deactivated on column (J)
          return {
            success: false,
            error: 'Employee account has been deactivated.'
          };
        }
        
        // Update last login
        employeesSheet.getRange(i + 1, 11).setValue(new Date()); // Column K
        
        return {
          success: true,
          employee: {
            id: data[i][0],                    // A
            name: data[i][1],                  // B
            email: data[i][2],                 // C
            openingBalance: data[i][3],        // D
            usableFrom: data[i][4],           // E
            usableTo: data[i][5],             // F
            addedOn: data[i][7],              // H
            lastLogin: new Date(),            // K (just updated)
            emergencyLeaves: data[i][12],     // M
            sickLeaves: data[i][13],         // N
            weeklyHoliday: data[i][14]       // O
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found with current email address.'
    };
    
  } catch (error) {
    console.error('Error getting current employee data:', error);
    return {
      success: false,
      error: 'Error retrieving employee data: ' + error.toString()
    };
  }
}


/**
 * Deactivate employee (replaces removeEmployee) - Updated for 19-column structure
 */
function deactivateEmployee(employeeId, employeeName) {
    const currentAdmin = getCurrentLoggedInAdmin();
    if (!currentAdmin || !currentAdmin.email) {
        showToast('Admin session expired. Please refresh and login again.', 'error');
        return;
    }
    
    showConfirmationDialog(
        'Deactivate Employee',
        `Are you sure you want to deactivate ${employeeName}? They will lose access to the employee portal.`,
        'deactivate',
        () => {
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showToast('Employee deactivated successfully', 'success');
                        refreshEmployees();
                    } else {
                        showToast('Failed to deactivate employee: ' + result.error, 'error');
                    }
                })
                .withFailureHandler(function(error) {
                    showToast('Error deactivating employee: ' + error.message, 'error');
                })
                .toggleEmployeeStatus(employeeId, currentAdmin.email); // Pass admin email
        }
    );
}

/**
 * Reactivate employee - New function for 19-column structure
 */
function reactivateEmployee(employeeId, employeeName) {
    const currentAdmin = getCurrentLoggedInAdmin();
    if (!currentAdmin || !currentAdmin.email) {
        showToast('Admin session expired. Please refresh and login again.', 'error');
        return;
    }
    
    showConfirmationDialog(
        'Reactivate Employee',
        `Are you sure you want to reactivate ${employeeName}? They will regain access to the employee portal.`,
        'activate',
        () => {
            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        showToast('Employee reactivated successfully', 'success');
                        refreshEmployees();
                    } else {
                        showToast('Failed to reactivate employee: ' + result.error, 'error');
                    }
                })
                .withFailureHandler(function(error) {
                    showToast('Error reactivating employee: ' + error.message, 'error');
                })
                .toggleEmployeeStatus(employeeId, currentAdmin.email); // Pass admin email
        }
    );
}

/**
 * Validate employee login credentials - Updated for 19-column structure
 */
function validateEmployeeCredentials(email, password) {
  try {
    if (!email || !password) {
      return "ERROR:Email and password are required";
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return "ERROR:No spreadsheet found";
    }
    
    // Try different sheet names
    let employeesSheet = ss.getSheetByName('Employees');
    if (!employeesSheet) {
      employeesSheet = ss.getSheetByName('employees');
    }
    if (!employeesSheet) {
      employeesSheet = ss.getSheetByName('Employee');
    }
    
    if (!employeesSheet) {
      return "ERROR:Employees sheet not found";
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return "ERROR:No employees found";
    }
    
    // Find employee by email
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Email should be in Column C (index 2)
      const rowEmail = row[2] ? row[2].toString().toLowerCase() : '';
      
      if (rowEmail === email.toLowerCase()) {
        // Check if deactivated (Column J)
        const deactivatedDate = row[9];
        if (deactivatedDate) {
          return "ERROR:Employee account has been deactivated";
        }
        
        // Check password (Column G)
        const storedPassword = row[6] || '';
        
        if (!storedPassword) {
          return "ERROR:No password set for this employee";
        }
        
        if (storedPassword.toString() === password.toString()) {
          // Update last login (Column K)
          try {
            employeesSheet.getRange(i + 1, 11).setValue(new Date());
          } catch (e) {
            // Ignore update errors
          }
          
          // Return success with employee data as string
          const id = row[0] || '';
          const name = row[1] || 'Employee';
          const openingBalance = row[3] || 0;
          const emergencyLeaves = row[12] || 0;
          const sickLeaves = row[13] || 0;
          
          return `SUCCESS:${id}|${name}|${email}|${openingBalance}|${emergencyLeaves}|${sickLeaves}`;
        } else {
          return "ERROR:Invalid password";
        }
      }
    }
    
    return "ERROR:Employee not found";
    
  } catch (error) {
    return "ERROR:System error occurred";
  }
}

/**
 * Get current admin email for tracking purposes
 */
function getCurrentAdminEmail() {
  try {
    // Get the currently logged in admin email
    // This should be set when admin logs in
    const properties = PropertiesService.getScriptProperties();
    const currentAdminEmail = properties.getProperty('currentAdminEmail');
    return currentAdminEmail || 'Unknown Admin';
  } catch (error) {
    console.error('Error getting current admin email:', error);
    return 'Unknown Admin';
  }
}

function getCurrentSystemConfig() {
  try {
    const configSheet = getSheet('Config');
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    return {
      success: true,
      startDate: configSheet.getRange('B3').getValue(),
      endDate: configSheet.getRange('B4').getValue(),
      annualDefault: configSheet.getRange('B9').getValue() || 21,
      sickDefault: configSheet.getRange('B10').getValue() || 7,
      emergencyDefault: configSheet.getRange('B11').getValue() || 3
    };
    
  } catch (error) {
    console.error('Error getting system config:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   OFFICIAL HOLIDAYS CARD-BASED INTERFACE - BACKEND FUNCTIONS
   Add these functions to admin.gs
   ============================================================================ */

/**
 * Get official holidays data for card-based display
 */
function getOfficialHolidaysForCards() {
  try {
    console.log('🔄 Getting official holidays for card display...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let holidaysSheet = ss.getSheetByName('Official holidays');
    
    // Create sheet if it doesn't exist
    if (!holidaysSheet) {
      console.log('📋 Official holidays sheet does not exist, creating...');
      holidaysSheet = ss.insertSheet('Official holidays');
      const headers = ['Holiday Name', 'Date'];
      holidaysSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      const headerRange = holidaysSheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#2c3e50');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      
      console.log('✅ Official holidays sheet created successfully');
      return "SUCCESS:[]"; // Empty array for new sheet
    }
    
    const lastRow = holidaysSheet.getLastRow();
    if (lastRow <= 1) {
      console.log('📋 No holidays found (only header or empty sheet)');
      return "SUCCESS:[]";
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const holidays = [];
    const holidayGroups = new Map(); // Group by holiday name
    
    // Process holidays and group by name
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[0] || !row[1]) {
        console.warn(`⚠️ Skipping incomplete row ${i}:`, row);
        continue;
      }
      
      const holidayName = row[0].toString().trim();
      let holidayDate;
      
      try {
        holidayDate = new Date(row[1]);
        if (isNaN(holidayDate.getTime())) {
          console.warn(`⚠️ Invalid date in row ${i}:`, row[1]);
          continue;
        }
      } catch (dateError) {
        console.warn(`⚠️ Error parsing date in row ${i}:`, dateError);
        continue;
      }
      
      // Convert to string format
      const dateStr = holidayDate.getFullYear() + '-' + 
                     String(holidayDate.getMonth() + 1).padStart(2, '0') + '-' + 
                     String(holidayDate.getDate()).padStart(2, '0');
      
      // Group dates by holiday name
      if (!holidayGroups.has(holidayName)) {
        holidayGroups.set(holidayName, []);
      }
      holidayGroups.get(holidayName).push(dateStr);
    }
    
    // Convert groups to holiday objects
    let holidayIndex = 0;
    holidayGroups.forEach((dates, name) => {
      dates.sort(); // Sort dates chronologically
      const startDate = dates[0];
      const endDate = dates[dates.length - 1];
      const startDateObj = new Date(startDate);
      const isPast = startDateObj < today;
      
      holidays.push({
        index: holidayIndex++,
        name: name,
        startDate: startDate,
        endDate: endDate,
        dates: dates,
        duration: dates.length,
        isPast: isPast
      });
    });
    
    // Sort holidays by start date
    holidays.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
    
    console.log(`✅ Processed ${holidays.length} holiday groups`);
    return "SUCCESS:" + JSON.stringify(holidays);
    
  } catch (error) {
    console.error('💥 Error getting holidays for cards:', error);
    return "ERROR:" + error.toString();
  }
}

/**
 * Get assignments for specific holiday (for card display) - FIXED TO READ ROW 2 EMPLOYEE IDS
 */
function getHolidayAssignmentsForCard(holidayName) {
  try {
    console.log('📄 Getting assignments for holiday:', holidayName);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    if (!holidaysSheet) {
      return "ERROR:Official holidays sheet not found";
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const headerRow = data[0]; // Row 1 headers
    const idRow = data[1];     // Row 2 employee IDs
    const assignments = {};
    
    // Get active employees
    const employeesResult = getActiveEmployeesForAssignment();
    if (employeesResult.startsWith('ERROR:')) {
      return employeesResult;
    }
    
    let employees = [];
    if (!employeesResult.startsWith('SUCCESS:EMPTY')) {
      const employeesString = employeesResult.replace('SUCCESS:', '');
      const employeeArray = employeesString.split(';');
      
      employeeArray.forEach(empStr => {
        const parts = empStr.split('|');
        if (parts.length >= 4) {
          employees.push({
            id: parts[2],
            name: parts[0],
            email: parts[1],
            weeklyHoliday: parts[3]
          });
        }
      });
    }
    
    // FIXED: Check employee columns using ROW 2 employee IDs (same as saving logic)
    employees.forEach(employee => {
      let employeeColumnIndex = -1;
      
      // Find employee column by checking ROW 2 for employee IDs
      for (let col = 3; col < headerRow.length; col++) {
        const employeeIdInSheet = idRow[col] ? idRow[col].toString().trim() : '';
        if (employeeIdInSheet === employee.id.toString().trim()) {
          employeeColumnIndex = col;
          console.log(`✅ Found column for employee ${employee.id} at column ${col + 1}`);
          break;
        }
      }
      
      // Default assignment for employee
      assignments[employee.id] = {
        status: 'Off',
        notificationStatus: 'Pending',
        workingDays: [],
        shouldNotify: true
      };
      
      // If employee has a column, check assignments for this holiday
      if (employeeColumnIndex >= 0) {
        const workingDays = [];
        let hasWorkAssignment = false;
        let hasNotifiedStatus = false;
        
        // Check all rows for this holiday name (start from row 3, index 2)
        for (let i = 2; i < data.length; i++) {
          const row = data[i];
          if (row[0] && row[0].toString().trim() === holidayName) {
            if (row[employeeColumnIndex]) {
              const cellValue = row[employeeColumnIndex].toString().trim();
              
              // Check if employee is working this day
              if (cellValue.toLowerCase().includes('work')) {
                hasWorkAssignment = true;
                
                // Add this date to working days
                if (row[1]) {
                  try {
                    const holidayDate = new Date(row[1]);
                    const dateStr = holidayDate.getFullYear() + '-' + 
                                   String(holidayDate.getMonth() + 1).padStart(2, '0') + '-' + 
                                   String(holidayDate.getDate()).padStart(2, '0');
                    workingDays.push(dateStr);
                  } catch (dateError) {
                    console.warn('⚠️ Date parsing error:', dateError);
                  }
                }
              }
              
              // Check notification status
              if (cellValue.toLowerCase().includes('notified')) {
                hasNotifiedStatus = true;
              }
            }
          }
        }
        
        // Set final assignment status
        assignments[employee.id] = {
          status: hasWorkAssignment ? 'Work' : 'Off',
          notificationStatus: hasNotifiedStatus ? 'Notified' : 'Pending', 
          workingDays: workingDays,
          shouldNotify: !hasNotifiedStatus // If already notified, don't notify again by default
        };
        
        console.log(`📋 Employee ${employee.id} assignment:`, assignments[employee.id]);
      }
    });
    
    const result = {
      employees: employees,
      assignments: assignments
    };
    
    console.log(`✅ Retrieved assignments for ${employees.length} employees`);
    return "SUCCESS:" + JSON.stringify(result);
    
  } catch (error) {
    console.error('💥 Error getting holiday assignments:', error);
    return "ERROR:" + error.toString();
  }
}

function saveHolidayAssignmentsFromCard(holidayName, holidayDates, employeeAssignments, adminEmail) {
  try {
    console.log('💾 Saving assignments for holiday:', holidayName);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    if (!holidaysSheet) {
      return "ERROR:Official holidays sheet not found";
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const headerRow = data[0];
    const idRow = data[1];
    const dates = holidayDates.split(',').map(d => d.trim());
    
    let updatesMade = 0;
    const employeeChanges = {};
    
    const adminResult = getAdminByEmail(adminEmail);
    const adminName = adminResult.success ? adminResult.data.name : 'Administrator';
    
    // Get employees data
    const employeesResult = getActiveEmployeesForAssignment();
    const employeesMap = {};
    
    if (!employeesResult.startsWith('ERROR:') && !employeesResult.startsWith('SUCCESS:EMPTY')) {
      const employeesString = employeesResult.replace('SUCCESS:', '');
      const employeeArray = employeesString.split(';');
      
      employeeArray.forEach(empStr => {
        const parts = empStr.split('|');
        if (parts.length >= 4) {
          employeesMap[parts[2]] = {
            name: parts[0],
            email: parts[1],
            id: parts[2]
          };
        }
      });
    }
    
    // Process each employee
    Object.keys(employeeAssignments).forEach(employeeId => {
      const assignment = employeeAssignments[employeeId];
      let employeeColumnIndex = -1;
      
      for (let col = 3; col < headerRow.length; col++) {
        const employeeIdInSheet = idRow[col] ? String(idRow[col]).trim() : '';
        if (employeeIdInSheet === String(employeeId).trim()) {
          employeeColumnIndex = col;
          break;
        }
      }
      
      if (employeeColumnIndex === -1) return;
      
      const empData = employeesMap[employeeId];
      if (!empData) return;
      
      if (!employeeChanges[employeeId]) {
        employeeChanges[employeeId] = {
          employeeEmail: String(empData.email),
          employeeName: String(empData.name),
          shouldNotify: assignment.shouldNotify === true,
          allDates: [], // All dates with their assignments
          hasChanges: false
        };
      }
      
      dates.forEach(dateStr => {
        let dateRowIndex = -1;
        for (let row = 2; row < data.length; row++) {
          const rowHolidayName = data[row][0];
          const rowDate = data[row][1] ? formatDateForSheet(new Date(data[row][1])) : '';
          
          if (String(rowHolidayName) === String(holidayName) && rowDate === dateStr) {
            dateRowIndex = row;
            break;
          }
        }
        
        if (dateRowIndex === -1) return;
        
        const isWorkingThisDate = assignment.workingDays && assignment.workingDays.includes(dateStr);
        const willNotify = assignment.shouldNotify === true;
        
        let newValue;
        if (isWorkingThisDate) {
          newValue = willNotify ? 'Work [Notified]' : 'Work [Pending]';
        } else {
          newValue = willNotify ? 'Off [Notified]' : 'Off [Pending]';
        }
        
        const currentValue = String(data[dateRowIndex][employeeColumnIndex] || '').trim();
        
        holidaysSheet.getRange(dateRowIndex + 1, employeeColumnIndex + 1).setValue(newValue);
        updatesMade++;
        
        // Check if this is a change (including blank to work/off)
        const normalizedCurrent = currentValue.toLowerCase().replace(/\s*\[.*?\]\s*/g, '').trim();
        const normalizedNew = newValue.toLowerCase().replace(/\s*\[.*?\]\s*/g, '').trim();
        const isBlankToAssignment = currentValue === '' && normalizedNew !== '';
        
        if (normalizedCurrent !== normalizedNew || isBlankToAssignment) {
          employeeChanges[employeeId].hasChanges = true;
        }
        
        // Add ALL dates to the list (for email content)
        employeeChanges[employeeId].allDates.push({
          date: dateStr,
          status: isWorkingThisDate ? 'Working' : 'Off'
        });
      });
    });
    
    // Send emails
    let emailResults = { success: true, sent: 0, failed: 0, errors: [] };
    
    Object.keys(employeeChanges).forEach(employeeId => {
      const empChange = employeeChanges[employeeId];
      
      // Send email ONLY if: shouldNotify is true AND there are actual changes
      if (empChange.shouldNotify && empChange.hasChanges && empChange.allDates.length > 0) {
        try {
          const workingDays = empChange.allDates.filter(d => d.status === 'Working');
          const offDays = empChange.allDates.filter(d => d.status === 'Off');
          
          let assignmentType, cardType, introMessage;
          if (workingDays.length > 0 && offDays.length > 0) {
            assignmentType = 'Mixed Assignment';
            cardType = 'info';
            introMessage = `You have a mixed assignment for this holiday: ${workingDays.length} working day(s) and ${offDays.length} day(s) off.`;
          } else if (workingDays.length > 0) {
            assignmentType = 'Working Days';
            cardType = 'warning';
            introMessage = 'You are scheduled to work during this official holiday.';
          } else {
            assignmentType = 'Days Off';
            cardType = 'success';
            introMessage = 'You are off during this official holiday.';
          }
          
          const datesList = empChange.allDates.map(d => {
            const formattedDate = formatDateForDisplay(d.date);
            const statusIcon = d.status === 'Working' ? '⛑' : '⛱';
            const statusColor = d.status === 'Working' ? '#e67e22' : '#27ae60';
            const bgColor = d.status === 'Working' ? '#fff3e0' : '#e8f5e9';
            return `<div style="padding: 8px; margin: 4px 0; border-left: 4px solid ${statusColor}; background: ${bgColor}; border-radius: 4px;">
              ${statusIcon} <strong>${formattedDate}</strong> - ${d.status}
            </div>`;
          }).join('');
          
          console.log(`📧 Sending ${assignmentType} email to ${empChange.employeeEmail}`);
          
          const result = sendHolidayAssignment(
            String(empChange.employeeEmail),
            String(empChange.employeeName),
            String(holidayName),
            String(datesList),
            String(assignmentType),
            String(adminName),
            String(cardType),
            String(introMessage)
          );
          
          if (result && result.success) {
            emailResults.sent++;
          } else {
            emailResults.failed++;
            emailResults.errors.push({ email: empChange.employeeEmail, error: result ? result.error : 'Unknown error' });
          }
        } catch (error) {
          emailResults.failed++;
          emailResults.errors.push({ email: empChange.employeeEmail, error: error.toString() });
        }
      }
    });
    
    if (emailResults.failed > 0) emailResults.success = false;
    
    console.log(`📊 Results: ${emailResults.sent} sent, ${emailResults.failed} failed`);
    
    return "SUCCESS:" + JSON.stringify({
      message: `Saved ${updatesMade} assignments`,
      emailResults: emailResults
    });
    
  } catch (error) {
    console.error('💥 Error saving assignments:', error);
    return "ERROR:" + error.toString();
  }
}

// Helper function to format date for sheet comparison
function formatDateForSheet(date) {
  if (!date) return '';
  const d = new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Helper function to format date for display in emails
function formatDateForDisplay(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  return date.toLocaleDateString('en-US', { 
    weekday: 'short', 
    year: 'numeric', 
    month: 'short', 
    day: 'numeric' 
  });
}

/* ============================================================================
   CRITICAL FIX: DATE STATUS UPDATER
   Handles string dates, DD-MM-YYYY formats, and time normalization.
   ============================================================================ */

/**
 * 1. ROBUST DATE PARSER
 * Ensures input is converted to a valid JS Date object, handling "DD-MM-YYYY" strings.
 */
function parseDateForStatusCalc(dateVal) {
  if (!dateVal) return null;

  // Case A: Already a Date object
  if (dateVal instanceof Date) {
    // Check if it's valid
    return isNaN(dateVal.getTime()) ? null : dateVal;
  }

  // Case B: String Date
  if (typeof dateVal === 'string') {
    const cleanStr = dateVal.trim();
    
    // Check for DD-MM-YYYY or DD/MM/YYYY (e.g. 16-10-2025)
    // Regex: Start with 1-2 digits, separator, 1-2 digits, separator, 4 digits
    const dmyMatch = cleanStr.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    
    if (dmyMatch) {
      const day = parseInt(dmyMatch[1], 10);
      const month = parseInt(dmyMatch[2], 10) - 1; // JS months are 0-11
      const year = parseInt(dmyMatch[3], 10);
      return new Date(year, month, day);
    }
    
    // Fallback for standard ISO or MM/DD/YYYY strings
    const standardDate = new Date(cleanStr);
    return isNaN(standardDate.getTime()) ? null : standardDate;
  }

  return null;
}

/**
 * 2. CALCULATE USED STATUS
 * Determines if a leave is "Used" or "Not yet".
 * Logic: Must be 'Approved' AND Date must be strictly in the past (before today).
 */
function calculateUsedStatus(leaveDate, responseStatus) {
  // 1. Check Response Status first (Case insensitive)
  const statusStr = String(responseStatus || '').toLowerCase();
  
  // If not approved, it cannot be "Used"
  if (statusStr !== 'approved') {
    return 'Not yet';
  }

  // 2. Parse the Date
  const dateObj = parseDateForStatusCalc(leaveDate);
  
  // If date is invalid, we cannot calculate, return safely
  if (!dateObj) { 
    console.warn('calculateUsedStatus: Invalid Date found', leaveDate);
    return 'Error'; 
  }

  // 3. Time Normalization (Compare DATES only, ignore TIME)
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set today to 00:00:00
  
  const checkDate = new Date(dateObj);
  checkDate.setHours(0, 0, 0, 0); // Set leave date to 00:00:00

  // 4. Comparison
  // If the leave date is strictly less than today, it is "Used"
  // (i.e., yesterday or before).
  if (checkDate < today) {
    return 'Used';
  }

  // If checkDate >= today, it's not used yet
  return 'Not yet';
}

/**
 * 3. UPDATE ALL SHEETS
 * Batches the updates for Annual, Sick, and Emergency sheets.
 */
function updateAllUsedStatusColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToProcess = ['Annual Leaves', 'Sick Leaves', 'Emergency Leaves'];
  let totalUpdated = 0;

  sheetsToProcess.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Get all data (assuming headers in row 1)
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // No data

    // Fetch Columns: D (Date, Index 3) to N (Used Status, Index 13)
    // We fetch the whole block to be safe and efficient
    // Range: Row 2, Col 1 (A) to LastRow, Col 14 (N)
    const range = sheet.getRange(2, 1, lastRow - 1, 14);
    const data = range.getValues();
    
    // Array to hold only the updated Column N values
    const updates = [];
    let sheetHasChanges = false;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const leaveDate = row[3];      // Column D
      const responseStatus = row[8]; // Column I
      const currentUsedStatus = row[13]; // Column N

      // Calculate correct status
      const newStatus = calculateUsedStatus(leaveDate, responseStatus);

      // Only mark for update if value is different
      if (newStatus !== currentUsedStatus && newStatus !== 'Error') {
        updates.push([newStatus]);
        sheetHasChanges = true;
        totalUpdated++;
      } else {
        // Keep existing value
        updates.push([currentUsedStatus]); 
      }
    }

    // Write back ONLY Column N (efficient batch write)
    if (sheetHasChanges) {
      sheet.getRange(2, 14, updates.length, 1).setValues(updates);
      console.log(`Updated ${sheetName}: Refreshed used status columns.`);
    }
  });

  return {
    success: true,
    message: `Updated status for ${totalUpdated} records across all sheets.`
  };
}

/**
 * Get working and off employee counts with detailed breakdown for tooltips
 * FULL FUNCTION - ADD TO admin.gs
 */
function getWorkingOffCounts(targetDate) {
  try {
    console.log('Getting working/off counts for date:', targetDate);
    
    const fullCalendarSheet = getSheet('Full-calendar');
    if (!fullCalendarSheet) {
      console.warn('Full-calendar sheet not found');
      return {
        workingCount: 0,
        offCount: 0,
        workingEmployees: [],
        offEmployees: []
      };
    }
    
    // Get all active employees
    const activeEmployees = getActiveEmployeesList();
    if (activeEmployees.length === 0) {
      return {
        workingCount: 0,
        offCount: 0,
        workingEmployees: [],
        offEmployees: []
      };
    }
    
    // Get full calendar data
    const data = fullCalendarSheet.getDataRange().getValues();
    if (data.length < 5) {
      return {
        workingCount: 0,
        offCount: 0,
        workingEmployees: [],
        offEmployees: []
      };
    }
    
    // Find target date row (dates start from row 5 = index 4)
    const targetDateObj = new Date(targetDate);
    targetDateObj.setHours(0, 0, 0, 0);
    let dateRow = -1;
    
    for (let row = 4; row < data.length; row++) {
      const cellDate = new Date(data[row][1]); // Column B
      cellDate.setHours(0, 0, 0, 0);
      if (cellDate.getTime() === targetDateObj.getTime()) {
        dateRow = row;
        break;
      }
    }
    
    if (dateRow === -1) {
      console.warn(`Date ${targetDate} not found in Full-calendar`);
      return {
        workingCount: 0,
        offCount: 0,
        workingEmployees: [],
        offEmployees: []
      };
    }
    
    const workingEmployees = [];
    const offEmployees = [];
    
    // Process each active employee
    activeEmployees.forEach(employee => {
      const employeeStatus = getEmployeeStatusOnDate(data, dateRow, employee.id);
      
      if (employeeStatus.isOff) {
        offEmployees.push({
          name: employee.name,
          reason: employeeStatus.reason
        });
      } else {
        workingEmployees.push({
          name: employee.name,
          pendingRequest: employeeStatus.pendingRequest
        });
      }
    });
    
    console.log(`Date ${targetDate}: Working=${workingEmployees.length}, Off=${offEmployees.length}`);
    
    return {
      workingCount: workingEmployees.length,
      offCount: offEmployees.length,
      workingEmployees: workingEmployees,
      offEmployees: offEmployees
    };
    
  } catch (error) {
    console.error('Error getting working/off counts:', error);
    return {
      workingCount: 0,
      offCount: 0,
      workingEmployees: [],
      offEmployees: []
    };
  }
}

/**
 * Get active employees list from Employees sheet with proper activation logic
 * FULL FUNCTION REPLACEMENT - UPDATE IN admin.gs
 */
function getActiveEmployeesList() {
  try {
    const employeesSheet = getSheet('Employees');
    if (!employeesSheet) return [];
    
    const data = employeesSheet.getDataRange().getValues();
    const activeEmployees = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const employeeId = String(row[0] || '');
      const employeeName = String(row[1] || '');
      const addedDate = row[7]; // Column H - Added on
      const deactivatedDate = row[9]; // Column J - Deactivated on
      const reactivatedDate = row[17]; // Column R - Reactivated on
      
      // Skip if missing essential data
      if (!employeeId || !employeeName || !addedDate) {
        continue;
      }
      
      // Determine if employee is currently active
      let isActive = false;
      
      if (!deactivatedDate) {
        // Never deactivated = active
        isActive = true;
      } else if (reactivatedDate) {
        // Has both deactivation and reactivation dates
        // Active only if reactivation is after deactivation
        const deactivatedDateObj = new Date(deactivatedDate);
        const reactivatedDateObj = new Date(reactivatedDate);
        
        isActive = reactivatedDateObj > deactivatedDateObj;
        
        console.log(`Employee ${employeeId}: Deactivated=${deactivatedDate}, Reactivated=${reactivatedDate}, Active=${isActive}`);
      } else {
        // Has deactivation but no reactivation = inactive
        isActive = false;
      }
      
      // Only include active employees
      if (isActive) {
        activeEmployees.push({
          id: employeeId,
          name: employeeName
        });
      }
    }
    
    console.log(`Found ${activeEmployees.length} active employees`);
    return activeEmployees;
    
  } catch (error) {
    console.error('Error getting active employees:', error);
    return [];
  }
}

/**
 * Get employee status on specific date from Full-calendar data
 * FULL FUNCTION - ADD TO admin.gs
 */
function getEmployeeStatusOnDate(fullCalendarData, dateRow, employeeId) {
  try {
    // Find employee columns (starting from column C = index 2)
    let employeeStartCol = -1;
    for (let col = 2; col < fullCalendarData[0].length; col += 5) { // Every 5 columns per employee
      if (String(fullCalendarData[0][col]) === String(employeeId)) {
        employeeStartCol = col;
        break;
      }
    }
    
    if (employeeStartCol === -1) {
      return { isOff: false, reason: '', pendingRequest: '' };
    }
    
    // Get values for all 5 columns for this employee on this date
    const annualLeave = String(fullCalendarData[dateRow][employeeStartCol] || ''); // Annual Leaves
    const weeklyHoliday = String(fullCalendarData[dateRow][employeeStartCol + 1] || ''); // Weekly Holidays
    const officialHoliday = String(fullCalendarData[dateRow][employeeStartCol + 2] || ''); // Official Holidays
    const sickLeave = String(fullCalendarData[dateRow][employeeStartCol + 3] || ''); // Sick Leaves
    const emergencyLeave = String(fullCalendarData[dateRow][employeeStartCol + 4] || ''); // Emergency Leaves
    
    // Check for off status (approved leaves or holidays)
    
    // Holiday checks (TRUE = off)
    if (weeklyHoliday.toLowerCase() === 'true') {
      return { isOff: true, reason: 'Weekly Holiday', pendingRequest: '' };
    }
    
    if (officialHoliday.toLowerCase() === 'true') {
      return { isOff: true, reason: 'Official Holiday', pendingRequest: '' };
    }
    
    // Leave checks (contains "Approved" = off)
    if (annualLeave.toLowerCase().includes('approved')) {
      const requestId = extractRequestId(annualLeave);
      return { isOff: true, reason: `${requestId} : Annual Leave`, pendingRequest: '' };
    }
    
    if (sickLeave.toLowerCase().includes('approved')) {
      const requestId = extractRequestId(sickLeave);
      return { isOff: true, reason: `${requestId} : Sick Leave`, pendingRequest: '' };
    }
    
    if (emergencyLeave.toLowerCase().includes('approved')) {
      const requestId = extractRequestId(emergencyLeave);
      return { isOff: true, reason: `${requestId} : Emergency Leave`, pendingRequest: '' };
    }
    
    // Check for pending requests (working but with pending request)
    let pendingRequest = '';
    
    if (annualLeave.toLowerCase().includes('pending')) {
      pendingRequest = extractRequestId(annualLeave) + ' : Pending';
    } else if (sickLeave.toLowerCase().includes('pending')) {
      pendingRequest = extractRequestId(sickLeave) + ' : Pending';
    } else if (emergencyLeave.toLowerCase().includes('pending')) {
      pendingRequest = extractRequestId(emergencyLeave) + ' : Pending';
    }
    
    // Employee is working
    return { isOff: false, reason: '', pendingRequest: pendingRequest };
    
  } catch (error) {
    console.error('Error getting employee status on date:', error);
    return { isOff: false, reason: '', pendingRequest: '' };
  }
}

/**
 * Extract request ID from Full-calendar entry (e.g., "A01 : Approved" -> "A01")
 * FULL FUNCTION - ADD TO admin.gs
 */
function extractRequestId(entry) {
  try {
    if (!entry || typeof entry !== 'string') return '';
    
    // Split by colon and get the first part, trim whitespace
    const parts = entry.split(':');
    if (parts.length > 0) {
      return parts[0].trim();
    }
    
    return '';
  } catch (error) {
    console.error('Error extracting request ID:', error);
    return '';
  }
}

/* ============================================================================
   AUTOMATED EMAIL SYSTEM - DAILY TRIGGER FUNCTIONS
   ADD THESE FUNCTIONS TO main.gs
   
   These functions handle all automated email notifications:
   - System End Date Reminders
   - Holiday Announcements
   - Auto Compensation
   - Holiday Wishes
   - Holiday End Reminders
   ============================================================================ */

/**
 * Setup daily email automation trigger
 * Call this once to install the trigger
 */
function setupDailyEmailAutomation() {
  try {
    // Delete existing triggers first
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runDailyEmailAutomation') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Get automation time from Config B13
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    const automationTime = configSheet.getRange('B13').getValue();
    if (!automationTime) {
      return {
        success: false,
        error: 'Automation time not set in Config B13'
      };
    }
    
    // Parse time (expected format: "HH:MM" or time value)
    let hour = 9; // Default 9 AM
    let minute = 0;
    
    if (typeof automationTime === 'string') {
      const parts = automationTime.split(':');
      if (parts.length === 2) {
        hour = parseInt(parts[0]) || 9;
        minute = parseInt(parts[1]) || 0;
      }
    } else if (automationTime instanceof Date) {
      hour = automationTime.getHours();
      minute = automationTime.getMinutes();
    }
    
    // Create daily trigger at specified time
    ScriptApp.newTrigger('runDailyEmailAutomation')
      .timeBased()
      .atHour(hour)
      .everyDays(1)
      .create();
    
    console.log(`Daily automation trigger created for ${hour}:${String(minute).padStart(2, '0')}`);
    
    return {
      success: true,
      message: `Automation trigger set for ${hour}:${String(minute).padStart(2, '0')} daily`
    };
    
  } catch (error) {
    console.error('Error setting up daily automation:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Main daily automation function (called by trigger)
 * Checks all automation conditions and sends appropriate emails
 */
function runDailyEmailAutomation() {
  try {
    console.log('=== DAILY EMAIL AUTOMATION START ===');
    console.log('Time:', new Date().toLocaleString());
    
    const results = {
      systemEndReminders: { sent: 0, errors: [] },
      holidayAnnouncements: { sent: 0, errors: [] },
      autoCompensations: { sent: 0, errors: [] },
      holidayWishes: { sent: 0, errors: [] },
      holidayEndReminders: { sent: 0, errors: [] }
    };
    
    // 1. Check and send System End Date Reminders
    const endReminders = checkAndSendSystemEndReminders();
    results.systemEndReminders = endReminders;
    
    // 2. Check and send Holiday Announcements
    const announcements = checkAndSendHolidayAnnouncements();
    results.holidayAnnouncements = announcements;
    
    // 3. Check and send Auto Compensation
    const compensations = checkAndSendAutoCompensations();
    results.autoCompensations = compensations;
    
    // 4. Check and send Holiday Wishes
    const wishes = checkAndSendHolidayWishes();
    results.holidayWishes = wishes;
    
    // 5. Check and send Holiday End Reminders
    const endRemindersLeave = checkAndSendHolidayEndReminders();
    results.holidayEndReminders = endRemindersLeave;
    
    console.log('=== DAILY EMAIL AUTOMATION COMPLETE ===');
    console.log('Results:', JSON.stringify(results));
    
    return results;
    
  } catch (error) {
    console.error('Error in daily automation:', error);
    
    // Notify admins of automation failure
    try {
      const admins = getFullPermissionAdmins();
      for (const admin of admins) {
        sendScheduledResetError(error.toString(), new Date());
      }
    } catch (notifyError) {
      console.error('Failed to notify admins of automation error:', notifyError);
    }
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   AUTOMATION CHECK FUNCTIONS
   ============================================================================ */

/**
 * Check and send System End Date Reminders (Scenario #9)
 */
function checkAndSendSystemEndReminders() {
  try {
    console.log('Checking System End Date Reminders...');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    const adminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins');
    
    if (!configSheet || !adminsSheet) {
      return { sent: 0, errors: ['Required sheets not found'] };
    }
    
    // Get system end date from Config B4
    const systemEndDate = new Date(configSheet.getRange('B4').getValue());
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    systemEndDate.setHours(0, 0, 0, 0);
    
    // Calculate days remaining
    const daysRemaining = Math.ceil((systemEndDate - today) / (1000 * 60 * 60 * 24));
    
    if (daysRemaining < 0) {
      return { sent: 0, errors: ['System period already ended'] };
    }
    
    // Get admins data
    const adminsData = adminsSheet.getDataRange().getValues();
    const results = { sent: 0, errors: [] };
    
    for (let i = 1; i < adminsData.length; i++) {
      const adminName = adminsData[i][0];
      const adminEmail = adminsData[i][1];
      const deactivatedOn = adminsData[i][5];
      
      // Skip deactivated admins
      if (deactivatedOn) continue;
      
      // Check if admin has notification preference enabled
      // This would be stored in a preferences column (adjust as needed)
      // For now, we'll check a hypothetical column or send to all active admins
      
      // Get admin's preference for reminder days (if stored)
      // For this implementation, we'll use a default or check a specific column
      const adminReminderDays = 7; // Default value, adjust based on your preference storage
      
      // Send reminder if days remaining matches admin's preference
      if (daysRemaining === adminReminderDays) {
        const result = sendSystemEndDateReminder(adminEmail, adminName, systemEndDate, daysRemaining);
        
        if (result.success) {
          results.sent++;
          console.log(`System end reminder sent to ${adminName}`);
        } else {
          results.errors.push(`Failed to send to ${adminEmail}: ${result.error}`);
        }
      }
    }
    
    return results;
    
  } catch (error) {
    console.error('Error in checkAndSendSystemEndReminders:', error);
    return { sent: 0, errors: [error.toString()] };
  }
}

/**
 * Check and send Holiday Announcements (Scenario #16) - UPDATED
 */
function checkAndSendHolidayAnnouncements() {
  try {
    console.log('Checking Holiday Announcements...');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    const holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Official Holidays');
    
    if (!configSheet || !holidaysSheet) {
      return { sent: 0, errors: ['Required sheets not found'] };
    }
    
    const announcementDays = parseInt(configSheet.getRange('B14').getValue()) || 0;
    
    if (announcementDays === 0) {
      return { sent: 0, errors: ['Holiday announcements disabled (B14 = 0)'] };
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const holidaysData = holidaysSheet.getDataRange().getValues();
    const results = { sent: 0, errors: [] };
    
    for (let i = 1; i < holidaysData.length; i++) {
      const holidayName = holidaysData[i][0];
      const holidayDate = new Date(holidaysData[i][1]);
      const holidayDescription = holidaysData[i][2] || '';
      
      holidayDate.setHours(0, 0, 0, 0);
      
      const daysUntilHoliday = Math.ceil((holidayDate - today) / (1000 * 60 * 60 * 24));
      
      if (daysUntilHoliday === announcementDays) {
        const recipients = getAllActiveEmployees();
        
        if (recipients.length === 0) {
          continue;
        }
        
        // UPDATED: Call sendHolidayAnnouncement with string parameters
        // Note: This function uses batch sending, so we need to call it once
        // and let sendToMultipleRecipients handle the distribution
        const emailPartsTemplate = sendHolidayAnnouncement(
          String(holidayName),                // holidayName
          holidayDate.toISOString(),         // holidayDate
          String(holidayDescription)          // holidayDescription
        );
        
        // The function returns a template with receiver: null
        // Now send to all recipients
        const sendResult = sendToMultipleRecipients(emailPartsTemplate, recipients);
        results.sent += sendResult.sent || 0;
        if (sendResult.errors && sendResult.errors.length > 0) {
          results.errors.push(...sendResult.errors.map(e => e.error));
        }
        
        console.log(`Holiday announcement sent for ${holidayName} to ${sendResult.sent} recipients`);
      }
    }
    
    return results;
    
  } catch (error) {
    console.error('Error in checkAndSendHolidayAnnouncements:', error);
    return { sent: 0, errors: [error.toString()] };
  }
}

/**
 * Check and send Auto Compensation (Scenario #18) - UPDATED
 */
function checkAndSendAutoCompensations() {
  try {
    console.log('Checking Auto Compensations...');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    const holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Official Holidays');
    const employeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
    
    if (!configSheet || !holidaysSheet || !employeesSheet) {
      return { sent: 0, errors: ['Required sheets not found'] };
    }
    
    const compensationDays = parseInt(configSheet.getRange('B15').getValue()) || 0;
    
    if (compensationDays === 0) {
      return { sent: 0, errors: ['Auto compensation disabled (B15 = 0)'] };
    }
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const holidaysData = holidaysSheet.getDataRange().getValues();
    const results = { sent: 0, errors: [] };
    
    for (let i = 1; i < holidaysData.length; i++) {
      const holidayName = holidaysData[i][0];
      const holidayDate = new Date(holidaysData[i][1]);
      holidayDate.setHours(0, 0, 0, 0);
      
      const daysSinceHoliday = Math.ceil((today - holidayDate) / (1000 * 60 * 60 * 24));
      
      if (daysSinceHoliday === compensationDays) {
        const workingEmployees = getEmployeesWhoWorkedOnHoliday(holidayDate);
        
        for (const emp of workingEmployees) {
          const compensationAmount = 1;
          const updateResult = addCompensationToEmployee(emp.id, compensationAmount, holidayName);
          
          if (updateResult.success) {
            // UPDATED: Pass string parameters instead of object
            const sendResult = sendAutoCompensation(
              emp.email,                          // employeeEmail
              String(emp.name),                   // employeeName
              String(holidayName),                // holidayName
              holidayDate.toISOString(),         // holidayDate
              String(compensationAmount),         // compensationDays
              String(updateResult.newBalance)     // newBalance
            );
            
            if (sendResult.success) {
              results.sent++;
              console.log(`Compensation email sent to ${emp.name}`);
            } else {
              results.errors.push(`Failed to send to ${emp.email}: ${sendResult.error}`);
            }
          } else {
            results.errors.push(`Failed to update balance for ${emp.name}: ${updateResult.error}`);
          }
        }
      }
    }
    
    return results;
    
  } catch (error) {
    console.error('Error in checkAndSendAutoCompensations:', error);
    return { sent: 0, errors: [error.toString()] };
  }
}

/**
 * Check and send Holiday Wishes (Scenario #19) - UPDATED
 */
function checkAndSendHolidayWishes() {
  try {
    console.log('Checking Holiday Wishes...');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    
    if (!configSheet) {
      return { sent: 0, errors: ['Config sheet not found'] };
    }
    
    const wishesEnabled = configSheet.getRange('B16').getValue();
    
    if (!wishesEnabled) {
      return { sent: 0, errors: ['Holiday wishes disabled (B16 = FALSE)'] };
    }
    
    const wishesDays = parseInt(configSheet.getRange('B17').getValue()) || 0;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const leaveSheets = ['Annual leaves'];
    const results = { sent: 0, errors: [] };
    
    for (const sheetName of leaveSheets) {
      const leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!leaveSheet) continue;
      
      const leaveData = leaveSheet.getDataRange().getValues();
      
      for (let i = 1; i < leaveData.length; i++) {
        const employeeId = leaveData[i][0];
        const employeeName = leaveData[i][1];
        const startDate = new Date(leaveData[i][3]);
        const endDate = new Date(leaveData[i][4]);
        const responseStatus = leaveData[i][10];
        
        startDate.setHours(0, 0, 0, 0);
        endDate.setHours(0, 0, 0, 0);
        
        if ((responseStatus === 'Approved' || responseStatus === 'Partially Approved') && startDate > today) {
          const daysUntilStart = Math.ceil((startDate - today) / (1000 * 60 * 60 * 24));
          
          if (daysUntilStart === wishesDays) {
            const empEmail = getEmployeeEmailById(employeeId);
            
            if (empEmail) {
              const duration = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
              
              // UPDATED: Pass string parameters instead of object
              const sendResult = sendHolidayWishes(
                empEmail,                           // employeeEmail
                String(employeeName),               // employeeName
                startDate.toISOString(),           // startDate
                endDate.toISOString(),             // endDate
                String(duration)                    // duration
              );
              
              if (sendResult.success) {
                results.sent++;
                console.log(`Holiday wishes sent to ${employeeName}`);
              } else {
                results.errors.push(`Failed to send to ${empEmail}: ${sendResult.error}`);
              }
            }
          }
        }
      }
    }
    
    return results;
    
  } catch (error) {
    console.error('Error in checkAndSendHolidayWishes:', error);
    return { sent: 0, errors: [error.toString()] };
  }
}

/**
 * Check and send Holiday End Reminders (Scenario #20) - UPDATED
 */
function checkAndSendHolidayEndReminders() {
  try {
    console.log('Checking Holiday End Reminders...');
    
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    
    if (!configSheet) {
      return { sent: 0, errors: ['Config sheet not found'] };
    }
    
    const remindersEnabled = configSheet.getRange('B18').getValue();
    
    if (!remindersEnabled) {
      return { sent: 0, errors: ['Holiday end reminders disabled (B18 = FALSE)'] };
    }
    
    const reminderDays = parseInt(configSheet.getRange('B19').getValue()) || 0;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const leaveSheets = ['Annual leaves'];
    const results = { sent: 0, errors: [] };
    
    for (const sheetName of leaveSheets) {
      const leaveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!leaveSheet) continue;
      
      const leaveData = leaveSheet.getDataRange().getValues();
      
      for (let i = 1; i < leaveData.length; i++) {
        const employeeId = leaveData[i][0];
        const employeeName = leaveData[i][1];
        const startDate = new Date(leaveData[i][3]);
        const endDate = new Date(leaveData[i][4]);
        const responseStatus = leaveData[i][10];
        
        startDate.setHours(0, 0, 0, 0);
        endDate.setHours(0, 0, 0, 0);
        
        if ((responseStatus === 'Approved' || responseStatus === 'Partially Approved') && startDate <= today && endDate > today) {
          const daysUntilEnd = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
          
          if (daysUntilEnd === reminderDays) {
            const empEmail = getEmployeeEmailById(employeeId);
            
            if (empEmail) {
              const returnDate = new Date(endDate);
              returnDate.setDate(returnDate.getDate() + 1);
              
              // UPDATED: Pass string parameters instead of object
              const sendResult = sendHolidayEndReminder(
                empEmail,                           // employeeEmail
                String(employeeName),               // employeeName
                endDate.toISOString(),             // endDate
                String(daysUntilEnd),              // daysRemaining
                returnDate.toISOString()           // returnDate
              );
              
              if (sendResult.success) {
                results.sent++;
                console.log(`Holiday end reminder sent to ${employeeName}`);
              } else {
                results.errors.push(`Failed to send to ${empEmail}: ${sendResult.error}`);
              }
            }
          }
        }
      }
    }
    
    return results;
    
  } catch (error) {
    console.error('Error in checkAndSendHolidayEndReminders:', error);
    return { sent: 0, errors: [error.toString()] };
  }
}

/* ============================================================================
   HELPER FUNCTIONS FOR AUTOMATION
   ============================================================================ */

/**
 * Get employee email by ID
 */
function getEmployeeEmailById(employeeId) {
  try {
    const employeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
    if (!employeesSheet) return null;
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(employeeId)) {
        return data[i][2]; // Column C - Email
      }
    }
    
    return null;
    
  } catch (error) {
    console.error('Error getting employee email:', error);
    return null;
  }
}

/**
 * Get employees who worked on a specific holiday
 * Adjust this based on your actual holiday assignment tracking structure
 */
function getEmployeesWhoWorkedOnHoliday(holidayDate) {
  try {
    // This is a placeholder implementation
    // Adjust based on how you track holiday assignments
    
    // If you have a "Holiday Assignments" sheet, query it
    // Otherwise, implement your own logic
    
    const assignmentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Holiday Assignments');
    if (!assignmentsSheet) return [];
    
    const data = assignmentsSheet.getDataRange().getValues();
    const workingEmployees = [];
    
    for (let i = 1; i < data.length; i++) {
      const assignmentDate = new Date(data[i][2]); // Adjust column index
      const status = data[i][3]; // Adjust column index
      
      assignmentDate.setHours(0, 0, 0, 0);
      
      if (assignmentDate.getTime() === holidayDate.getTime() && status === 'Working') {
        workingEmployees.push({
          id: data[i][0],
          name: data[i][1],
          email: data[i][4] // Adjust based on your structure
        });
      }
    }
    
    return workingEmployees;
    
  } catch (error) {
    console.error('Error getting working employees:', error);
    return [];
  }
}

/**
 * Add compensation days to employee balance
 */
function addCompensationToEmployee(employeeId, compensationDays, reason) {
  try {
    const employeesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees');
    if (!employeesSheet) {
      return { success: false, error: 'Employees sheet not found' };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(employeeId)) {
        const currentBalance = parseFloat(data[i][6]) || 0; // Column G - Current Balance
        const newBalance = currentBalance + compensationDays;
        
        // Update balance
        employeesSheet.getRange(i + 1, 7).setValue(newBalance); // Column G
        
        return {
          success: true,
          newBalance: newBalance
        };
      }
    }
    
    return { success: false, error: 'Employee not found' };
    
  } catch (error) {
    console.error('Error adding compensation:', error);
    return { success: false, error: error.toString() };
  }
}

/* ============================================================================
   EMAIL TESTING BACKEND FUNCTIONS
   ADD THESE FUNCTIONS TO main.gs
   
   These functions handle test email sending with sample data
   ============================================================================ */

/**
 * Show email test interface
 */
function showEmailTestInterface() {
  const html = HtmlService.createHtmlOutputFromFile('test-emails')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Email System Testing');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Email System Testing');
}

/**
 * Send test email based on type
 * @param {string} emailType - Type of email to test
 * @param {string} testEmailOverride - Optional email override for testing
 * @returns {Object} Result object
 */
function sendTestEmail(emailType, testEmailOverride) {
  try {
    console.log('Sending test email:', emailType);
    
    // Get current user email (logged in admin)
    const currentAdminEmail = Session.getActiveUser().getEmail();
    let currentAdminName = 'Test Admin';
    
    // Try to get admin name from Admins sheet
    try {
      const adminsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admins');
      if (adminsSheet) {
        const data = adminsSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][1] === currentAdminEmail) {
            currentAdminName = data[i][0] || 'Test Admin';
            break;
          }
        }
      }
    } catch (e) {
      console.log('Could not fetch admin name:', e);
    }
    
    // Determine recipient
    const recipient = testEmailOverride || currentAdminEmail;
    
    let result;
    
    switch(emailType) {
      
      // SYSTEM NOTIFICATIONS
      case 'otp':
        result = sendSystemResetOTP(recipient, currentAdminName, '123456');
        break;
        
      case 'adminWelcome':
        result = sendAdminWelcomeEmail(
          recipient,
          'Test New Admin',
          'TempPass123!',
          { fullPermission: true, receiveRequests: true },
          currentAdminName
        );
        break;
        
      case 'adminReactivation':
        result = sendAdminReactivationEmail(
          recipient,
          'Test Admin',
          { fullPermission: true, receiveRequests: true },
          currentAdminName
        );
        break;
        
      case 'passwordPin':
        result = sendPasswordChangePIN(recipient, currentAdminName, '789456');
        break;
        
      case 'passwordConfirm':
        result = sendPasswordChangeConfirmation(recipient, currentAdminName);
        break;
        
      case 'systemReset':
        const resetDetails = {
          type: 'Immediate',
          newStartDate: new Date(),
          newEndDate: new Date(Date.now() + 365 * 24 * 60 * 60 * 1000)
        };
        // Send to single recipient for test
        const emailTemplate = {
          receiver: recipient,
          subject: 'System Reset Executed',
          subtitle: 'System Reset Notification',
          emailType: 'System Reset',
          actualSender: currentAdminName,
          body_greeting: 'Hello Administrator,',
          body_intro: 'The Holiday Management System has been reset.',
          body_content: buildEmailCard(
            'System Reset Details',
            `Reset Type: Immediate\nNew Start Date: ${formatDateForDisplay(resetDetails.newStartDate)}\nNew End Date: ${formatDateForDisplay(resetDetails.newEndDate)}\nExecuted At: ${formatDateTimeForDisplay(new Date())}`,
            'warning'
          ) + buildEmailAlert(
            'All employee balances and leave request statuses have been reset to their opening values. This action cannot be undone.',
            'warning'
          ),
          body_footer: 'Best regards,\nHoliday Management System',
          logDescription: 'System reset notification - Test'
        };
        result = standardSendEmail(emailTemplate);
        break;
        
      case 'resetCancel':
        const cancelTemplate = {
          receiver: recipient,
          subject: 'Scheduled System Reset Cancelled',
          subtitle: 'System Reset Cancellation',
          emailType: 'System Reset Cancellation',
          actualSender: currentAdminName,
          body_greeting: 'Hello Administrator,',
          body_intro: 'The scheduled system reset has been cancelled.',
          body_content: buildEmailAlert(
            'The scheduled system reset has been cancelled. No changes have been made to employee balances or leave requests.',
            'info'
          ),
          body_footer: 'Best regards,\nHoliday Management System',
          logDescription: 'System reset cancellation - Test'
        };
        result = standardSendEmail(cancelTemplate);
        break;
        
      case 'resetError':
        result = sendScheduledResetError('Test error: Scheduled execution failed due to permission issues', new Date());
        result.receiver = recipient; // Override for test
        break;
        
      case 'endReminder':
        const endDate = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000);
        result = sendSystemEndDateReminder(recipient, currentAdminName, endDate, 7);
        break;
        
      // EMPLOYEE NOTIFICATIONS
      case 'employeeWelcome':
        result = sendEmployeeWelcomeEmail(
          'Test Employee',
          recipient,
          'EMP001',
          'Welcome123!',
          currentAdminName
        );
        break;
        
      case 'balanceUpdate':
        result = sendBalanceUpdate(
          recipient,              // employeeEmail
          'Test Employee',        // employeeName
          '21',                   // previousBalance
          '5',                    // changeAmount
          '26',                   // newBalance
          'Annual bonus leave days',  // reason
          currentAdminName        // updatedBy
        );
        break;
        
      // LEAVE & HOLIDAY NOTIFICATIONS
      case 'leaveRequest':
        result = sendLeaveRequestNotification(
          'Test Employee',                    // employeeName
          'EMP001',                           // employeeId
          'Annual',                           // leaveType
          new Date().toISOString(),          // startDate
          new Date(Date.now() + 5 * 24 * 60 * 60 * 1000).toISOString(),  // endDate
          '5',                               // duration
          'Family vacation',                 // reason
          new Date().toISOString()          // requestDate
        );
        break;
        
      case 'leaveApproved':
        result = sendLeaveApproved({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          leaveType: 'Annual',
          startDate: new Date(),
          endDate: new Date(Date.now() + 5 * 24 * 60 * 60 * 1000),
          duration: 5,
          adminName: currentAdminName,
          adminComments: 'Approved as requested. Enjoy your time off!',
          actionDate: new Date()
        });
        break;
        
      case 'leaveRejected':
        result = sendLeaveRejected({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          leaveType: 'Annual',
          startDate: new Date(),
          endDate: new Date(Date.now() + 5 * 24 * 60 * 60 * 1000),
          duration: 5,
          adminName: currentAdminName,
          adminComments: 'Unfortunately, we cannot approve leave during this busy period. Please consider alternative dates.',
          actionDate: new Date()
        });
        break;
        
      case 'leavePartial':
        result = sendLeavePartiallyApproved({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          leaveType: 'Annual',
          startDate: new Date(),
          endDate: new Date(Date.now() + 5 * 24 * 60 * 60 * 1000),
          duration: 5,
          approvedDates: '01/01/2025, 02/01/2025, 03/01/2025',
          rejectedDates: '04/01/2025, 05/01/2025',
          approvedDuration: 3,
          rejectedDuration: 2,
          adminName: currentAdminName,
          adminComments: 'Approved first 3 days. Last 2 days conflict with project deadline.',
          actionDate: new Date()
        });
        break;
        
      case 'holidayAnnouncement':
        const announceTemplate = {
          receiver: recipient,
          subject: 'Holiday Announcement - New Year Day',
          subtitle: 'Official Holiday Announcement',
          emailType: 'Holiday Announcement',
          actualSender: 'System',
          body_greeting: 'Hello Team Member,',
          body_intro: 'We would like to inform you about an upcoming official holiday:',
          body_content: buildEmailCard(
            'New Year Day',
            `Date: ${formatDateForDisplay(new Date('2025-01-01'))}\n\nOfficial public holiday - offices will be closed.`,
            'success'
          ) + '<h4 style="color: #2c3e50; margin: 25px 0 15px 0;">Important Notes:</h4>' + buildEmailList([
            'This is an official holiday - offices will be closed',
            'If you are scheduled to work on this holiday, compensation will be added to your leave balance',
            'Please plan your work schedule accordingly',
            'For urgent matters, contact your supervisor'
          ], false),
          body_footer: 'Best regards,\nHoliday Management System',
          logDescription: 'Holiday announcement test'
        };
        result = standardSendEmail(announceTemplate);
        break;
        
      case 'holidayAssignment':
        result = sendHolidayAssignment({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          holidayName: 'New Year Day',
          holidayDate: new Date('2025-01-01'),
          status: 'Working',
          adminName: currentAdminName
        });
        break;
        
      case 'autoCompensation':
        result = sendAutoCompensation({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          holidayName: 'New Year Day',
          holidayDate: new Date('2025-01-01'),
          compensationDays: 1,
          newBalance: 22
        });
        break;
        
      case 'holidayWishes':
        result = sendHolidayWishes({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          startDate: new Date(),
          endDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000),
          duration: 7
        });
        break;
        
      case 'holidayEndReminder':
        result = sendHolidayEndReminder({
          employeeEmail: recipient,
          employeeName: 'Test Employee',
          endDate: new Date(Date.now() + 2 * 24 * 60 * 60 * 1000),
          daysRemaining: 2,
          returnDate: new Date(Date.now() + 3 * 24 * 60 * 60 * 1000)
        });
        break;
        
      default:
        return {
          success: false,
          error: 'Unknown email type: ' + emailType
        };
    }
    
    if (result && result.success) {
      return {
        success: true,
        message: 'Test email sent successfully',
        recipient: recipient
      };
    } else {
      return {
        success: false,
        error: result ? result.error : 'Unknown error occurred'
      };
    }
    
  } catch (error) {
    console.error('Error in sendTestEmail:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Remove daily email automation trigger
 */
function removeDailyEmailAutomation() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let removed = 0;
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'runDailyEmailAutomation') {
        ScriptApp.deleteTrigger(trigger);
        removed++;
      }
    });
    
    if (removed > 0) {
      return {
        success: true,
        message: `Removed ${removed} automation trigger(s)`
      };
    } else {
      return {
        success: false,
        error: 'No automation triggers found'
      };
    }
    
  } catch (error) {
    console.error('Error removing automation:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function formatLongDateForDisplay(date) {
  try {
    return new Date(date).toLocaleDateString('en-GB', {
      day: '2-digit',
      month: 'short',
      year: 'numeric'
    });
  } catch (error) {
    return date.toString();
  }
}
