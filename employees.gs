/* ============================================================================
   EMPLOYEE PORTAL - BACKEND FUNCTIONS
   Backend functionality for employee portal operations
   ============================================================================ */

/* ============================================================================
   EMPLOYEE AUTHENTICATION AND SESSION MANAGEMENT
   ============================================================================ */

/**
 * Get current employee data for logged-in employee - CORRECT 19-COLUMN STRUCTURE
 * Column Mapping (A-S):
 */
function getCurrentEmployeeData(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
      };
    }

    const employeesSheet = getSheet('Employees');
    
    if (!employeesSheet) {
      return {
        success: false,
        error: 'Employees sheet not found. Please ensure the sheet exists.'
      };
    }
    
    const data = employeesSheet.getDataRange().getValues();
    
    // Find employee by email (Column C)
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === employeeEmail.toLowerCase()) {
        
        // Check if employee is currently active
        const deactivatedOn = data[i][9];  // Column J - Deactivated on
        const reactivatedOn = data[i][17]; // Column R - Reactivated on
        
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
        
        if (!isActive) {
          return {
            success: false,
            error: 'Employee account has been deactivated'
          };
        }
        
        // Return employee data using correct 19-column structure
        return {
          success: true,
          data: {
            id: data[i][0] || '',                    // A - ID
            name: data[i][1] || 'Unknown',           // B - Name
            email: data[i][2] || '',                 // C - E-mail
            openingBalance: data[i][3] || 0,         // D - Opening Balance
            usableFrom: data[i][4] || '',            // E - Usable from
            usableTo: data[i][5] || '',              // F - Usable to
            addedOn: data[i][7] || '',               // H - Added on
            emailStatus: data[i][8] || 'Pending',    // I - Email status
            lastLogin: data[i][10] || '',            // K - Last login
            activeInactive: data[i][11] || 'Active', // L - Active/Inactive
            emergencyLeaves: data[i][12] || 0,       // M - Opening emergency leaves
            sickLeaves: data[i][13] || 0,            // N - Opening Sick Leaves
            weeklyHoliday: data[i][14] || '',        // O - Weekly Holiday
            addedBy: data[i][15] || ''               // P - Added by
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error in getCurrentEmployeeData:', error);
    return {
      success: false,
      error: 'System error: ' + error.toString()
    };
  }
}

/**
 * Update employee last login timestamp
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Success status
 */
function updateEmployeeLastLogin(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
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
    
    // Find employee by email and update last login
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === employeeEmail.toLowerCase()) {
        // Update Last Login (Column K)
        employeesSheet.getRange(i + 1, 11).setValue(new Date());
        
        return {
          success: true,
          message: 'Last login updated successfully'
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error updating employee last login:', error);
    return {
      success: false,
      error: 'Error updating last login: ' + error.toString()
    };
  }
}


/* ============================================================================
   EMPLOYEE LEAVE BALANCE MANAGEMENT
   ============================================================================ */

/**
 * Get employee leave balance summary
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Leave balance data or error
 */
function getEmployeeLeaveBalance(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
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
    
    // Find employee by email
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString().toLowerCase() === employeeEmail.toLowerCase()) {
        
        return {
          success: true,
          data: {
            openingBalance: data[i][3] || 0,        // Column D - Annual Opening Balance
            emergencyLeaves: data[i][12] || 0,      // Column M - Opening emergency leaves
            sickLeaves: data[i][13] || 0,           // Column N - Opening Sick Leaves
            usableFrom: data[i][4] || '',           // Column E - Usable from
            usableTo: data[i][5] || '',             // Column F - Usable to
            weeklyHoliday: data[i][14] || ''        // Column O - Weekly Holiday
          }
        };
      }
    }
    
    return {
      success: false,
      error: 'Employee not found'
    };
    
  } catch (error) {
    console.error('Error getting employee leave balance:', error);
    return {
      success: false,
      error: 'Error retrieving leave balance: ' + error.toString()
    };
  }
}

/* ============================================================================
   EMPLOYEE WEEKLY HOLIDAYS MANAGEMENT
   ============================================================================ */

/**
 * Get employee weekly holiday preferences
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Weekly holiday data or error
 */
function getEmployeeWeeklyHolidays(employeeEmail) {
  try {
    console.log('=== GET WEEKLY HOLIDAYS START ===');
    console.log('Employee email:', employeeEmail);
    
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return "ERROR:" + employeeResult.error;
    }
    
    const employeeId = employeeResult.data.id;
    console.log('Employee ID:', employeeId);
    
    const weeklySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Weekly Holidays');
    
    if (!weeklySheet) {
      console.log('Weekly Holidays sheet not found');
      return "SUCCESS:EMPTY";
    }
    
    const data = weeklySheet.getDataRange().getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    console.log('Scanning weekly holidays, today:', today);
    
    let activeHolidays = null;
    
    for (let i = 1; i < data.length; i++) {
      const empId = String(data[i][1]); // Column B - Employee ID
      const startDate = data[i][3] ? new Date(data[i][3]) : null; // Column D
      const endDate = data[i][4] ? new Date(data[i][4]) : null; // Column E
      const status = String(data[i][5]); // Column F - Status
      const type = String(data[i][6]); // Column G - Period Type
      
      // Only check for matching employee with Active status
      if (empId === String(employeeId) && status === 'Active') {
        
        if (startDate) startDate.setHours(0, 0, 0, 0);
        if (endDate) endDate.setHours(0, 0, 0, 0);
        
        // Check if today falls within this period
        const isValid = today >= startDate && (!endDate || today <= endDate);
        
        console.log(`Row ${i}: EmpID=${empId}, Type=${type}, Start=${startDate}, End=${endDate}, Valid=${isValid}`);
        
        if (isValid) {
          // Extract holiday days (columns H-N)
          // CRITICAL: Check for both boolean true AND string 'True'
          const friday = data[i][7] === true || String(data[i][7]).toLowerCase() === 'true';
          const saturday = data[i][8] === true || String(data[i][8]).toLowerCase() === 'true';
          const sunday = data[i][9] === true || String(data[i][9]).toLowerCase() === 'true';
          const monday = data[i][10] === true || String(data[i][10]).toLowerCase() === 'true';
          const tuesday = data[i][11] === true || String(data[i][11]).toLowerCase() === 'true';
          const wednesday = data[i][12] === true || String(data[i][12]).toLowerCase() === 'true';
          const thursday = data[i][13] === true || String(data[i][13]).toLowerCase() === 'true';
          
          console.log('Holiday days:', { friday, saturday, sunday, monday, tuesday, wednesday, thursday });
          
          // Temp period takes priority over standard
          if (type === 'temp' || !activeHolidays) {
            activeHolidays = {
              type: type,
              friday: friday,
              saturday: saturday,
              sunday: sunday,
              monday: monday,
              tuesday: tuesday,
              wednesday: wednesday,
              thursday: thursday
            };
            
            console.log('Active holidays set:', activeHolidays);
          }
        }
      }
    }
    
    if (!activeHolidays) {
      console.log('No active holidays found');
      return "SUCCESS:EMPTY";
    }
    
    // Return as string: type|friday|saturday|sunday|monday|tuesday|wednesday|thursday
    const holidayString = `${activeHolidays.type}|${activeHolidays.friday}|${activeHolidays.saturday}|${activeHolidays.sunday}|${activeHolidays.monday}|${activeHolidays.tuesday}|${activeHolidays.wednesday}|${activeHolidays.thursday}`;
    
    console.log('Returning holiday string:', holidayString);
    console.log('=== GET WEEKLY HOLIDAYS END ===');
    
    return "SUCCESS:" + holidayString;
    
  } catch (error) {
    console.error('Error getting weekly holidays:', error);
    return "ERROR:System error - " + error.toString();
  }
}

/* ============================================================================
   EMPLOYEE CALENDAR FUNCTIONS
   ============================================================================ */

/**
 * Get employee calendar data (leaves, holidays, etc.)
 * @param {string} employeeEmail - Email of the employee
 * @param {string} year - Year to get calendar for (optional, defaults to current year)
 * @return {Object} Calendar data or error
 */
function getEmployeeCalendarData(employeeEmail, year) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
      };
    }

    const currentYear = year || new Date().getFullYear().toString();
    
    // Get employee basic data
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return employeeResult;
    }
    
    // Get leave balance
    const balanceResult = getEmployeeLeaveBalance(employeeEmail);
    if (!balanceResult.success) {
      return balanceResult;
    }
    
    // Get weekly holidays
    const weeklyResult = getEmployeeWeeklyHolidays(employeeEmail);
    if (!weeklyResult.success) {
      return weeklyResult;
    }
    
    // TODO: Get official holidays, leave requests, etc.
    // For now, return basic calendar structure
    
    return {
      success: true,
      data: {
        employee: employeeResult.data,
        balance: balanceResult.data,
        weeklyHolidays: weeklyResult.data,
        year: currentYear,
        // TODO: Add leave requests, official holidays, etc.
        leaveRequests: [],
        officialHolidays: []
      }
    };
    
  } catch (error) {
    console.error('Error getting employee calendar data:', error);
    return {
      success: false,
      error: 'Error retrieving calendar data: ' + error.toString()
    };
  }
}

/* ============================================================================
   UTILITY FUNCTIONS FOR EMPLOYEE PORTAL
   ============================================================================ */

/**
 * Format employee data as simple string for frontend - CORRECT 19-COLUMN STRUCTURE
 * Updated mapping based on 19-column structure (A-S)
 * @param {Object} employeeData - Employee data object
 * @return {string} Formatted string
 */
function formatEmployeeDataAsString(employeeData) {
  try {
    const parts = [
      employeeData.id || '',               // A - ID
      employeeData.name || '',             // B - Name
      employeeData.email || '',            // C - E-mail
      employeeData.openingBalance || '0',  // D - Opening Balance
      employeeData.usableFrom || '',       // E - Usable from
      employeeData.usableTo || '',         // F - Usable to
      employeeData.addedOn || '',          // H - Added on
      employeeData.emailStatus || 'Pending', // I - Email status
      employeeData.deactivatedOn || '',    // J - Deactivated on
      employeeData.lastLogin || '',        // K - Last login
      employeeData.emergencyLeaves || '0', // M - Opening emergency leaves
      employeeData.sickLeaves || '0',      // N - Opening Sick Leaves
      employeeData.weeklyHoliday || '',    // O - Weekly Holiday
      employeeData.addedBy || '',          // P - Added by
      employeeData.reactivatedOn || ''     // R - Reactivated on
    ];
    
    return parts.join('|');
    
  } catch (error) {
    console.error('Error formatting employee data as string:', error);
    return '';
  }
}

/**
 * Validate employee session and permissions
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Validation result
 */
function validateEmployeeSession(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
      };
    }

    const employeeResult = getCurrentEmployeeData(employeeEmail);
    
    if (!employeeResult.success) {
      return {
        success: false,
        error: 'Invalid employee session: ' + employeeResult.error
      };
    }
    
    return {
      success: true,
      data: employeeResult.data
    };
    
  } catch (error) {
    console.error('Error validating employee session:', error);
    return {
      success: false,
      error: 'Session validation error: ' + error.toString()
    };
  }
}

/**
 * Log employee activity
 * @param {string} employeeEmail - Email of the employee
 * @param {string} action - Action performed
 * @param {string} details - Additional details (optional)
 * @return {void}
 */
function logEmployeeActivity(employeeEmail, action, details) {
  try {
    const timestamp = new Date();
    const logMessage = `[EMPLOYEE] ${timestamp.toISOString()} - ${employeeEmail} - ${action}`;
    
    if (details) {
      console.log(`${logMessage} - ${details}`);
    } else {
      console.log(logMessage);
    }
    
    // TODO: Implement activity logging to sheet if required
    
  } catch (error) {
    console.error('Error logging employee activity:', error);
  }
}

/* ============================================================================
   EMPLOYEE PORTAL DATA RETRIEVAL FUNCTIONS
   ============================================================================ */

/**
 * Get system configuration relevant to employees
 * @return {Object} System configuration data
 */
function getEmployeeSystemConfig() {
  try {
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    // Get basic system configuration
    const systemDateStart = configSheet.getRange('B3').getValue();
    const systemDateEnd = configSheet.getRange('B4').getValue();
    const senderEmail = configSheet.getRange('B2').getValue();
    
    return {
      success: true,
      data: {
        systemDateStart: systemDateStart || '',
        systemDateEnd: systemDateEnd || '',
        senderEmail: senderEmail || '',
        currentDate: new Date()
      }
    };
    
  } catch (error) {
    console.error('Error getting employee system config:', error);
    return {
      success: false,
      error: 'Error retrieving system configuration: ' + error.toString()
    };
  }
}

function getSystemConfiguration() {
  try {
    console.log('getSystemConfiguration called - reading from Config sheet');
    
    // Get the Config sheet
    const configSheet = getSheet('Config');
    
    if (!configSheet) {
      console.error('Config sheet not found');
      return getDefaultSystemConfig();
    }
    
    // Read START DATE from B3 and END DATE from B4
    const startDateValue = configSheet.getRange('B3').getValue();
    const endDateValue = configSheet.getRange('B4').getValue();
    
    console.log('Raw values from Config sheet:', { startDateValue, endDateValue });
    
    // Parse dates
    let startDate = null;
    let endDate = null;
    
    // Handle START DATE (B3)
    if (startDateValue) {
      if (startDateValue instanceof Date) {
        startDate = new Date(startDateValue);
      } else {
        startDate = new Date(startDateValue);
      }
    }
    
    // Handle END DATE (B4)
    if (endDateValue) {
      if (endDateValue instanceof Date) {
        endDate = new Date(endDateValue);
      } else {
        endDate = new Date(endDateValue);
      }
    }
    
    // Validate dates
    if (!startDate || isNaN(startDate.getTime())) {
      console.warn('Invalid start date in Config B3, using default');
      startDate = new Date(new Date().getFullYear(), 0, 1);
    }
    
    if (!endDate || isNaN(endDate.getTime())) {
      console.warn('Invalid end date in Config B4, using default');
      endDate = new Date(new Date().getFullYear(), 11, 31);
    }
    
    // Ensure end date is after start date
    if (endDate <= startDate) {
      console.warn('End date is not after start date, adjusting');
      endDate = new Date(startDate);
      endDate.setFullYear(startDate.getFullYear() + 1);
    }
    
    console.log('Parsed system dates:', { startDate, endDate });
    
    return {
      success: true,
      data: {
        startDate: startDate,
        endDate: endDate,
        allowWeekends: true,
        allowPastDates: false,
        source: 'Config sheet B3, B4'
      }
    };
    
  } catch (error) {
    console.error('Error reading from Config sheet:', error);
    return getDefaultSystemConfig();
  }
}

/**
 * Check if current date is within employee's usable period
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Validation result
 */
function validateEmployeeUsablePeriod(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
      };
    }

    const employeeResult = getCurrentEmployeeData(employeeEmail);
    
    if (!employeeResult.success) {
      return employeeResult;
    }
    
    const employee = employeeResult.data;
    const currentDate = new Date();
    const usableFrom = new Date(employee.usableFrom);
    const usableTo = new Date(employee.usableTo);
    
    // Check if current date is within usable period
    const isWithinPeriod = currentDate >= usableFrom && currentDate <= usableTo;
    
    return {
      success: true,
      data: {
        isWithinPeriod: isWithinPeriod,
        usableFrom: employee.usableFrom,
        usableTo: employee.usableTo,
        currentDate: currentDate.toISOString(),
        daysUntilStart: isWithinPeriod ? 0 : Math.ceil((usableFrom - currentDate) / (1000 * 60 * 60 * 24)),
        daysUntilEnd: isWithinPeriod ? Math.ceil((usableTo - currentDate) / (1000 * 60 * 60 * 24)) : 0
      }
    };
    
  } catch (error) {
    console.error('Error validating employee usable period:', error);
    return {
      success: false,
      error: 'Error validating usable period: ' + error.toString()
    };
  }
}

/* ============================================================================
   EMPLOYEE PORTAL MAIN FUNCTIONS (PLACEHOLDERS FOR FUTURE IMPLEMENTATION)
   ============================================================================ */

/**
 * Get employee leave requests (placeholder)
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Leave requests data
 */

/**
 * COMPLETE REPLACEMENT for getEmployeeLeaveRequests function in employees.gs
 * This fixes the past days data retrieval issue
 */
function getEmployeeLeaveRequests(employeeEmail) {
  try {
    if (!employeeEmail) {
      return "ERROR:Employee email is required";
    }

    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return "ERROR:" + employeeResult.error;
    }
    
    const employeeId = employeeResult.data.id;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const requests = [];
    
    const sheets = [
      { name: 'Annual leaves', type: 'annual', prefix: 'A' },
      { name: 'Sick leaves', type: 'sick', prefix: 'S' },
      { name: 'Emergency leaves', type: 'emergency', prefix: 'E' }
    ];
    
    sheets.forEach(sheetInfo => {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetInfo.name);
        if (!sheet) return;
        
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return; // No data rows
        
        const processedRequests = new Map();
        
        // Step 1: Find all request IDs that have at least one future/ongoing day
        const eligibleRequestIds = new Set();
        
        for (let i = 1; i < data.length; i++) {
          const requestId = data[i][0];   // Column A
          const empId = data[i][1];       // Column B  
          const dateValue = data[i][3];   // Column D
          
          if (requestId && empId == employeeId && dateValue) {
            const rowDate = new Date(dateValue);
            if (!isNaN(rowDate.getTime()) && rowDate >= today) {
              eligibleRequestIds.add(requestId);
            }
          }
        }
        
        // Step 2: Get ALL days for eligible requests (including past days)
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const requestId = row[0];       // Column A
          const empId = row[1];           // Column B
          const empName = row[2];         // Column C
          const dateValue = row[3];       // Column D
          const weekDay = row[4];         // Column E
          const reason = row[5];          // Column F
          const weeklyHoliday = row[6];   // Column G
          const timestamp = row[7];       // Column H
          const responseStatus = row[8];  // Column I
          const responseTime = row[9];    // Column J
          const respondedBy = row[10];    // Column K
          const notification = row[11];   // Column L
          const adminComment = row[12];   // Column M
          const used = row[13];           // Column N
          
          // Process ALL days for eligible requests
          if (requestId && empId == employeeId && eligibleRequestIds.has(requestId) && dateValue) {
            
            // Initialize request if not exists
            if (!processedRequests.has(requestId)) {
              processedRequests.set(requestId, {
                id: requestId,
                type: sheetInfo.type,
                reason: reason || '',
                requestTimestamp: timestamp || new Date(),
                days: [],
                weeklyHolidayDays: [],
                isPending: false,
                hasResponse: false,
                summary: { approved: 0, rejected: 0, pending: 0 }
              });
            }
            
            const request = processedRequests.get(requestId);
            
            // Track weekly holiday days for badges
            if (weeklyHoliday === 'True' || weeklyHoliday === true) {
              const dayName = weekDay;
              if (dayName && !request.weeklyHolidayDays.includes(dayName)) {
                request.weeklyHolidayDays.push(dayName);
              }
            }
            
            // Add day details (ALL days - past, present, future)
            request.days.push({
              date: dateValue,
              weekDay: weekDay || '',
              weeklyHoliday: weeklyHoliday === 'True' || weeklyHoliday === true,
              responseStatus: responseStatus || 'Pending',
              responseTimestamp: responseTime || '',
              respondedBy: respondedBy || '',
              adminComment: adminComment || '',
              used: used || 'Not yet'
            });
            
            // Track pending and response status
            if (responseStatus === 'Pending') request.isPending = true;
            if (responseStatus && responseStatus !== 'Pending') request.hasResponse = true;
          }
        }
        
        // Step 3: Process each request - calculate summaries and dates
        processedRequests.forEach((request, requestId) => {
          if (request.days.length === 0) return;
          
          // Sort days by date
          request.days.sort((a, b) => new Date(a.date) - new Date(b.date));
          
          // Calculate summary counts
          const approved = request.days.filter(d => d.responseStatus === 'Approved').length;
          const rejected = request.days.filter(d => d.responseStatus === 'Rejected').length;
          const pending = request.days.filter(d => d.responseStatus === 'Pending').length;
          
          request.summary = {
            approved: approved,
            rejected: rejected,
            pending: pending,
            status: calculateRequestStatus(approved, rejected, pending)
          };
          
          // Set actual start and end dates from data
          request.startDate = request.days[0].date;
          request.endDate = request.days[request.days.length - 1].date;
          request.totalDays = request.days.length;
          request.netDays = request.days.filter(d => !d.weeklyHoliday).length;
          
          // Determine overall request status
          request.isPending = pending > 0;
          request.hasResponse = (approved + rejected) > 0;
          
          requests.push(request);
        });
        
      } catch (sheetError) {
        console.warn('Error processing ' + sheetInfo.name + ':', sheetError);
      }
    });
    
    // Sort requests: Pending first, then by type, then by start date
    requests.sort((a, b) => {
      if (a.isPending && !b.isPending) return -1;
      if (!a.isPending && b.isPending) return 1;
      
      const typeOrder = { annual: 0, sick: 1, emergency: 2 };
      if (typeOrder[a.type] !== typeOrder[b.type]) {
        return typeOrder[a.type] - typeOrder[b.type];
      }
      
      return new Date(a.startDate) - new Date(b.startDate);
    });
    
    // Convert to string format for frontend
    const requestStrings = requests.map(req => {
      const daysString = req.days.map(day => 
        `${new Date(day.date).toISOString()}~${day.weekDay}~${day.weeklyHoliday}~${day.responseStatus}~${day.responseTimestamp}~${day.respondedBy}~${day.adminComment}~${day.used}`
      ).join('#');
      
      const weeklyHolidayBadges = req.weeklyHolidayDays.join(',');
      
      return `${req.id}|${req.type}|${req.reason}|${req.requestTimestamp}|${new Date(req.startDate).toISOString()}|${new Date(req.endDate).toISOString()}|${req.totalDays}|${req.netDays}|${req.summary.approved}|${req.summary.rejected}|${req.summary.pending}|${req.summary.status}|${req.isPending}|${req.hasResponse}|${daysString}|${weeklyHolidayBadges}`;
    });
    
    console.log('✅ getEmployeeLeaveRequests - Found requests:', requests.length);
    
    return "SUCCESS:" + requestStrings.join(';');
    
  } catch (error) {
    console.error('Error getting employee leave requests:', error);
    return "ERROR:System error - " + error.toString();
  }
}

/**
 * NEW: Format date for display
 */
function formatDateForDisplay(date) {
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

/**
 * Get official holidays (placeholder)
 * @param {string} year - Year to get holidays for
 * @return {Object} Official holidays data
 */
function getOfficialHolidaysEp() {
  try {
    const holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('official holidays');
    if (!holidaysSheet) {
      return "SUCCESS:EMPTY";
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const holidays = [];
    
    for (let i = 3; i < data.length; i++) { // Data starts at row 4
      if (data[i][1]) { // Column B has date
        holidays.push(`${data[i][0]}|${new Date(data[i][1]).toISOString()}`);
      }
    }
    
    return "SUCCESS:" + holidays.join(';');
    
  } catch (error) {
    console.error('Error getting official holidays:', error);
    return "ERROR:System error - " + error.toString();
  }
}

/* ============================================================================
   HISTORY TAB BACKEND - FULL REPLACEMENT
   ============================================================================ */

/**
 * HELPER: Robust Date Parser
 * Handles Excel serial dates, standard dates, and DD-MM-YYYY strings
 */
function parseDateSafe(dateVal) {
  if (!dateVal) return null;

  // 1. If it's already a Date object
  if (dateVal instanceof Date) {
    return isNaN(dateVal.getTime()) ? null : dateVal;
  }

  // 2. If it's a string, try to detect format
  if (typeof dateVal === 'string') {
    dateVal = dateVal.trim();
    
    // DETECT DD-MM-YYYY or DD/MM/YYYY format (e.g. "16-10-2025")
    // Regex matches starts with 1-2 digits, separator, 1-2 digits, separator, 4 digits
    const dmyMatch = dateVal.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    
    if (dmyMatch) {
      // Create date: Year, Month (0-indexed), Day
      // dmyMatch[1] = Day, dmyMatch[2] = Month, dmyMatch[3] = Year
      return new Date(dmyMatch[3], parseInt(dmyMatch[2]) - 1, dmyMatch[1]);
    }
  }

  // 3. Fallback to standard parser
  const d = new Date(dateVal);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * HELPER: Format DateTime string for display
 */
function formatDateTime(dateVal) {
  const d = parseDateSafe(dateVal);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM dd, yyyy HH:mm");
}

/**
 * MAIN FUNCTION: Fetch and aggregate leave history
 */
function getEmployeeLeaveHistory(employeeEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Get Employee ID based on Email
    const empSheet = ss.getSheetByName('Employees');
    if (!empSheet) {
      return { success: false, error: 'Employees sheet missing' };
    }
    
    const empData = empSheet.getDataRange().getValues();
    let employeeId = null;
    
    // Find employee ID (Column A is ID, Column C is Email)
    for (let i = 1; i < empData.length; i++) {
      if (String(empData[i][2]).toLowerCase() === String(employeeEmail).toLowerCase()) {
        employeeId = empData[i][0];
        break;
      }
    }
    
    if (!employeeId) {
      return { success: false, error: 'Employee profile not found' };
    }

    // 2. Define sources with SPECIFIC ICONS
    const sources = [
      { sheet: 'Annual Leaves', type: 'Annual', icon: 'fa-calendar-day' },        // Green
      { sheet: 'Sick Leaves', type: 'Sick', icon: 'fa-user-md' },                 // Orange
      { sheet: 'Emergency Leaves', type: 'Emergency', icon: 'fa-exclamation-triangle' } // Red
    ];

    let history = [];

    // 3. Iterate sources and aggregate data
    sources.forEach(source => {
      const sheet = ss.getSheetByName(source.sheet);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      let requestsMap = {};

      // Parse Sheet Data
      for (let i = 1; i < data.length; i++) {
        // Filter by Employee ID (Column B, Index 1)
        if (String(data[i][1]) === String(employeeId)) {
          const reqNo = data[i][0]; // Col A: Request No
          
          if (!requestsMap[reqNo]) {
            // Initialize request object if first time seeing this ID
            requestsMap[reqNo] = {
              id: reqNo,
              type: source.type,
              typeIcon: source.icon,
              reason: data[i][5] || 'No reason provided', // Col F
              status: data[i][8] || 'Pending',            // Col I
              
              // Meta Data
              requestedOn: formatDateTime(data[i][7]),    // Col H
              respondOn: formatDateTime(data[i][9]),      // Col J
              respondBy: data[i][10] || '',               // Col K
              adminComment: data[i][12] || '',            // Col M
              
              dates: [] // Array to store Date objects
            };
          }
          
          // Parse and collect the specific date (Col D, Index 3)
          const validDate = parseDateSafe(data[i][3]);
          if (validDate) {
            requestsMap[reqNo].dates.push(validDate);
          }
        }
      }

      // Finalize request objects (Calculate ranges, durations, etc.)
      Object.values(requestsMap).forEach(req => {
        if (req.dates.length > 0) {
          // Sort dates to find start and end
          req.dates.sort((a, b) => a - b);
          
          const startDate = req.dates[0];
          const endDate = req.dates[req.dates.length - 1];
          req.duration = req.dates.length;
          
          // Format date range string
          const startStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMM dd, yyyy");
          const endStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMM dd, yyyy");
          req.dateRange = (startStr === endStr) ? startStr : `${startStr} - ${endStr}`;
          
          // Store raw start date for sorting the final list
          req.startDate = startDate.toISOString();
          
          // Determine status key for CSS styling
          let statusKey = 'default';
          const s = String(req.status).toLowerCase();
          if (s.includes('approve')) statusKey = 'approved';
          else if (s.includes('reject')) statusKey = 'rejected';
          else if (s.includes('pending')) statusKey = 'pending';
          
          req.statusKey = statusKey;
          
          // Remove raw date objects before sending to frontend (optimization)
          delete req.dates;
          
          history.push(req);
        }
      });
    });

    // 4. Sort all history by Start Date descending (Newest first)
    history.sort((a, b) => new Date(b.startDate) - new Date(a.startDate));

    return {
      success: true,
      data: history
    };

  } catch (error) {
    console.error('Error in getEmployeeLeaveHistory:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/* ============================================================================
   EMPLOYEE PORTAL INTEGRATION FUNCTIONS
   ============================================================================ */

/**
 * Get employee portal dashboard data
 * This function combines all necessary data for the employee portal
 * @param {string} employeeEmail - Email of the employee
 * @return {Object} Complete dashboard data
 */
function getEmployeePortalDashboard(employeeEmail) {
  try {
    if (!employeeEmail) {
      return {
        success: false,
        error: 'Employee email is required'
      };
    }

    // Get employee data
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return employeeResult;
    }

    // Get leave balance
    const balanceResult = getEmployeeLeaveBalance(employeeEmail);
    
    // Get weekly holidays
    const weeklyResult = getEmployeeWeeklyHolidays(employeeEmail);
    
    // Get system config
    const configResult = getEmployeeSystemConfig();
    
    // Get usable period validation
    const periodResult = validateEmployeeUsablePeriod(employeeEmail);
    
    // Update last login
    updateEmployeeLastLogin(employeeEmail);
    
    // Log activity
    logEmployeeActivity(employeeEmail, 'Portal Access', 'Employee portal dashboard loaded');
    
    return {
      success: true,
      data: {
        employee: employeeResult.data,
        balance: balanceResult.success ? balanceResult.data : null,
        weeklyHolidays: weeklyResult.success ? weeklyResult.data : null,
        systemConfig: configResult.success ? configResult.data : null,
        usablePeriod: periodResult.success ? periodResult.data : null,
        lastAccessed: new Date().toISOString()
      }
    };
    
  } catch (error) {
    console.error('Error getting employee portal dashboard:', error);
    return {
      success: false,
      error: 'Error loading dashboard: ' + error.toString()
    };
  }
}

/* ============================================================================
   EMPLOYEE PORTAL HELPER FUNCTIONS FOR FRONTEND INTEGRATION
   ============================================================================ */

/**
 * Format employee data for frontend display (simple string format)
 * This function returns employee data in the format expected by frontend
 * @param {string} employeeEmail - Email of the employee
 * @return {string} Formatted employee data string or error
 */
function getEmployeeDataForFrontend(employeeEmail) {
  try {
    const result = getCurrentEmployeeData(employeeEmail);
    
    if (!result.success) {
      return "ERROR:" + result.error;
    }
    
    const formattedData = formatEmployeeDataAsString(result.data);
    return "SUCCESS:" + formattedData;
    
  } catch (error) {
    console.error('Error getting employee data for frontend:', error);
    return "ERROR:System error - " + error.toString();
  }
}

/* ============================================================================
   LEAVE REQUESTS TAB - BACKEND FUNCTIONS
   Functions for employee leave requests management
   ============================================================================ */

/**
 * Get employee leave dashboard data
 * @param {string} employeeEmail - Employee's email
 * @return {Object} Dashboard data with balances and counts
 */
function getEmployeeLeaveDashboard(employeeEmail) {
  try {
    console.log('Getting dashboard for employee:', employeeEmail);
    
    if (!employeeEmail) {
      return "ERROR:Employee email is required";
    }

    // Get employee data first
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return "ERROR:" + employeeResult.error;
    }
    
    const employeeId = employeeResult.data.id;
    console.log('Employee ID found:', employeeId);
    
    // Get current balances
    const currentBalances = getSimpleEmployeeBalances(employeeId);
    console.log('Current balances:', currentBalances);
    
    // Get on-hold counts
    const onHoldCounts = getSimpleOnHoldCounts(employeeId);
    console.log('On-hold counts:', onHoldCounts);
    
    // Get awaiting response counts
    const awaitingCounts = getSimpleAwaitingCounts(employeeId);
    console.log('Awaiting counts:', awaitingCounts);
    
    // Format as simple string
    const dashboardString = [
      currentBalances.annual || 0, onHoldCounts.annual || 0, 
      (currentBalances.annual || 0) - (onHoldCounts.annual || 0), awaitingCounts.annual || 0,
      currentBalances.sick || 0, onHoldCounts.sick || 0,
      (currentBalances.sick || 0) - (onHoldCounts.sick || 0), awaitingCounts.sick || 0,
      currentBalances.emergency || 0, onHoldCounts.emergency || 0,
      (currentBalances.emergency || 0) - (onHoldCounts.emergency || 0), awaitingCounts.emergency || 0
    ].join('|');
    
    console.log('Dashboard string:', dashboardString);
    return "SUCCESS:" + dashboardString;
    
  } catch (error) {
    console.error('Error getting dashboard:', error);
    return "ERROR:System error - " + error.toString();
  }
}

function getSimpleEmployeeBalances(employeeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const balanceSheet = ss.getSheetByName('Balance - Calendar');
    
    if (!balanceSheet) {
      console.error('Balance - Calendar sheet not found');
      return { annual: 0, sick: 0, emergency: 0 };
    }
    
    console.log('Reading from Balance - Calendar sheet for employee ID:', employeeId);
    
    // Get employee IDs from row 1 (starting from column C)
    const lastCol = balanceSheet.getLastColumn();
    const employeeIdsRow = balanceSheet.getRange(1, 3, 1, lastCol - 2).getValues()[0];
    
    // Get leave types from row 3 (starting from column C)
    const leaveTypesRow = balanceSheet.getRange(3, 3, 1, lastCol - 2).getValues()[0];
    
    console.log('Employee IDs found:', employeeIdsRow);
    console.log('Leave types found:', leaveTypesRow);
    
    // Find columns for this employee
    let annualCol = -1, sickCol = -1, emergencyCol = -1;
    
    for (let i = 0; i < employeeIdsRow.length; i++) {
      if (employeeIdsRow[i] == employeeId) {
        const leaveType = leaveTypesRow[i];
        if (leaveType === 'Annual Leaves') {
          annualCol = i + 3; // Convert to actual column number
        } else if (leaveType === 'Sick Leaves') {
          sickCol = i + 3;
        } else if (leaveType === 'Emergency Leaves') {
          emergencyCol = i + 3;
        }
      }
    }
    
    console.log('Found columns - Annual:', annualCol, 'Sick:', sickCol, 'Emergency:', emergencyCol);
    
    if (annualCol === -1 && sickCol === -1 && emergencyCol === -1) {
      console.warn('No columns found for employee ID:', employeeId);
      return { annual: 0, sick: 0, emergency: 0 };
    }
    
    // Find today's date row (dates start from row 5 in column B)
    const today = new Date();
    const dateColumn = balanceSheet.getRange(5, 2, balanceSheet.getLastRow() - 4, 1).getValues();
    let todayRow = -1;
    
    for (let i = 0; i < dateColumn.length; i++) {
      if (dateColumn[i][0]) {
        const sheetDate = new Date(dateColumn[i][0]);
        if (sheetDate.toDateString() === today.toDateString()) {
          todayRow = i + 5; // Convert to actual row number
          break;
        }
      }
    }
    
    // If today's date not found, get the latest available date
    if (todayRow === -1 && dateColumn.length > 0) {
      todayRow = dateColumn.length + 4; // Last row with data
      console.log('Today not found, using latest date row:', todayRow);
    }
    
    if (todayRow === -1) {
      console.warn('No date rows found in Balance - Calendar sheet');
      return { annual: 0, sick: 0, emergency: 0 };
    }
    
    console.log('Using date row:', todayRow);
    
    // Get the balances
    const balances = {
      annual: annualCol > -1 ? (balanceSheet.getRange(todayRow, annualCol).getValue() || 0) : 0,
      sick: sickCol > -1 ? (balanceSheet.getRange(todayRow, sickCol).getValue() || 0) : 0,
      emergency: emergencyCol > -1 ? (balanceSheet.getRange(todayRow, emergencyCol).getValue() || 0) : 0
    };
    
    console.log('Retrieved balances:', balances);
    return balances;
    
  } catch (error) {
    console.error('Error reading from Balance - Calendar sheet:', error);
    return { annual: 0, sick: 0, emergency: 0 };
  }
}

/**
 * FIXES:
 * 1. Counts individual dates instead of request IDs
 * 2. Excludes weekly holidays from count
 * 3. Only counts future dates (date > today)
 */
function getSimpleOnHoldCounts(employeeId) {
  try {
    const counts = { annual: 0, sick: 0, emergency: 0 };
    const sheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Set to midnight for accurate comparison
    
    sheets.forEach((sheetName, index) => {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return;
        
        const data = sheet.getDataRange().getValues();
        const processedDays = new Set(); // Track unique dates (not request IDs)
        
        for (let i = 1; i < data.length; i++) {
          const empId = data[i][1];           // Column B - Employee ID
          const requestId = data[i][0];       // Column A - Request ID
          const dateValue = data[i][3];       // Column D - Date
          const weeklyHoliday = data[i][6];   // Column G - Weekly Holiday (True/False)
          const responseStatus = data[i][8];  // Column I - Response status
          const used = data[i][13];           // Column N - Used
          
          // Check if this is a weekly holiday
          const isWeeklyHoliday = weeklyHoliday && 
                                  (weeklyHoliday.toString().toLowerCase() === 'true' || 
                                   weeklyHoliday === true);
          
          // All conditions must be met:
          // 1. Same employee
          // 2. Approved
          // 3. Not yet used
          // 4. NOT a weekly holiday (this is NET days)
          // 5. Date exists
          if (empId == employeeId && 
              responseStatus === 'Approved' && 
              used === 'Not yet' && 
              !isWeeklyHoliday &&  // <-- EXCLUDE WEEKLY HOLIDAYS
              dateValue) {
            
            const leaveDate = new Date(dateValue);
            leaveDate.setHours(0, 0, 0, 0);
            
            // Only count future dates (exclude today and past)
            if (leaveDate > today) {
              // Create unique key: requestId + date to count each day separately
              const uniqueKey = requestId + '-' + leaveDate.toISOString();
              
              if (!processedDays.has(uniqueKey)) {
                processedDays.add(uniqueKey);
                
                // Increment appropriate counter based on sheet
                if (index === 0) counts.annual++;
                else if (index === 1) counts.sick++;
                else counts.emergency++;
              }
            }
          }
        }
        
      } catch (sheetError) {
        console.warn('Error processing sheet ' + sheetName + ':', sheetError);
      }
    });
    
    return counts;
    
  } catch (error) {
    console.error('Error getting on-hold counts:', error);
    return { annual: 0, sick: 0, emergency: 0 };
  }
}

function getSimpleAwaitingCounts(employeeId) {
  try {
    const counts = { annual: 0, sick: 0, emergency: 0 };
    const sheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    
    sheets.forEach((sheetName, index) => {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) return;
        
        const data = sheet.getDataRange().getValues();
        const pendingRequests = new Set();
        
        for (let i = 1; i < data.length; i++) {
          const empId = data[i][1];           // Column B - Employee ID
          const requestId = data[i][0];       // Column A - Request ID
          const responseStatus = data[i][8];  // Column I - Response status
          
          if (empId == employeeId && 
              responseStatus === 'Pending' && 
              !pendingRequests.has(requestId)) {
            
            pendingRequests.add(requestId);
          }
        }
        
        if (index === 0) counts.annual = pendingRequests.size;
        else if (index === 1) counts.sick = pendingRequests.size;
        else counts.emergency = pendingRequests.size;
        
      } catch (sheetError) {
        console.warn('Error processing sheet ' + sheetName + ':', sheetError);
      }
    });
    
    return counts;
    
  } catch (error) {
    console.error('Error getting awaiting counts:', error);
    return { annual: 0, sick: 0, emergency: 0 };
  }
}

/**
 * Submit new leave request - COMPLETE IMPLEMENTATION
 * @param {string} employeeEmail - Employee's email
 * @param {string} leaveType - Type of leave (annual, sick, emergency)
 * @param {string} startDate - Start date (YYYY-MM-DD format)
 * @param {string} endDate - End date (YYYY-MM-DD format)
 * @param {string} reason - Reason for leave
 * @return {string} Success/Error response
 */
function submitLeaveRequest(employeeEmail, leaveType, startDate, endDate, reason) {
  try {
    console.log('='.repeat(70));
    console.log('=== SUBMIT LEAVE REQUEST START ===');
    console.log('='.repeat(70));
    console.log('Employee:', employeeEmail);
    console.log('Leave Type:', leaveType);
    console.log('Start Date:', startDate);
    console.log('End Date:', endDate);
    console.log('Reason:', reason);
    
    // Validate required fields
    if (!employeeEmail || !leaveType || !startDate || !endDate || !reason) {
      console.log('❌ VALIDATION FAILED: Missing required fields');
      return "ERROR:All fields are required";
    }
    
    // Get employee data
    console.log('\n--- Getting Employee Data ---');
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      console.log('❌ Failed to get employee data:', employeeResult.error);
      return "ERROR:" + employeeResult.error;
    }
    
    const employeeId = employeeResult.data.id;
    const employeeName = employeeResult.data.name;
    
    console.log('✅ Employee found:');
    console.log('  - ID:', employeeId);
    console.log('  - Name:', employeeName);
    
    // Validate dates
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    if (start > end) {
      console.log('❌ VALIDATION FAILED: Start date after end date');
      return "ERROR:Start date must be before end date";
    }
    
    // Get sheet info
    const sheetInfo = {
      annual: { name: 'Annual leaves', prefix: 'A' },
      sick: { name: 'Sick leaves', prefix: 'S' },
      emergency: { name: 'Emergency leaves', prefix: 'E' }
    };
    
    const info = sheetInfo[leaveType];
    if (!info) {
      console.log('❌ Invalid leave type:', leaveType);
      return "ERROR:Invalid leave type";
    }
    
    console.log('\n--- Getting Leave Sheet ---');
    console.log('Sheet name:', info.name);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info.name);
    if (!sheet) {
      console.log('❌ Sheet not found:', info.name);
      return "ERROR:" + info.name + " sheet not found";
    }
    console.log('✅ Sheet found');
    
    // Generate date range
    console.log('\n--- Generating Date Range ---');
    const dates = [];
    const currentDate = new Date(start);
    
    while (currentDate <= end) {
      dates.push(new Date(currentDate));
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    console.log('Total dates in range:', dates.length);
    console.log('Date range:', dates[0].toISOString().split('T')[0], 'to', dates[dates.length - 1].toISOString().split('T')[0]);
    
    // Check for official holidays if annual leave
    if (leaveType === 'annual') {
      console.log('\n--- Checking Official Holidays ---');
      const holidayCheck = checkOfficialHolidaysInRange(dates);
      if (holidayCheck && holidayCheck.hasHolidays) {
        console.log('❌ Official holidays found in range:', holidayCheck.holidays.join(', '));
        return "ERROR:Cannot request annual leave on official holidays: " + holidayCheck.holidays.join(', ');
      }
      console.log('✅ No official holidays in range');
    }
    
    // Map short day name to full lowercase name
    const dayMapping = {
      'Sun': 'sunday',
      'Mon': 'monday',
      'Tue': 'tuesday',
      'Wed': 'wednesday',
      'Thu': 'thursday',
      'Fri': 'friday',
      'Sat': 'saturday'
    };
    
    // Calculate net days requested (excluding weekly holidays)
    console.log('\n' + '='.repeat(70));
    console.log('=== BALANCE VALIDATION START ===');
    console.log('='.repeat(70));
    console.log('Validating balance for:', leaveType, 'leave');
    
    let totalDays = dates.length;
    let weeklyHolidayDays = 0;
    
    console.log('\n--- Checking Weekly Holidays for Each Date ---');
    dates.forEach((date, index) => {
      const dateStr = date.toISOString().split('T')[0];
      const holidayInfo = getWeeklyHolidayForSpecificDate(employeeId, date);
      const weekDay = date.toLocaleDateString('en-US', { weekday: 'short' });
      const fullDayName = dayMapping[weekDay];
      
      const isWeeklyHoliday = holidayInfo.success && holidayInfo.hasHoliday && holidayInfo.holidayDays[fullDayName] === true;
      
      if (isWeeklyHoliday) {
        weeklyHolidayDays++;
        console.log(`  ${index + 1}. ${dateStr} (${weekDay}) → ✅ WEEKLY HOLIDAY [Period: ${holidayInfo.periodType}]`);
      } else {
        console.log(`  ${index + 1}. ${dateStr} (${weekDay}) → Working day [Period: ${holidayInfo.periodType || 'None'}]`);
      }
    });
    
    const netDaysRequested = totalDays - weeklyHolidayDays;
    
    console.log('\n--- Request Summary ---');
    console.log('Total days in request:', totalDays);
    console.log('Weekly holiday days:', weeklyHolidayDays);
    console.log('NET DAYS REQUESTED:', netDaysRequested);
    
    // Get current balance from Balance-Calendar sheet
    console.log('\n--- Getting Current Balance from Balance-Calendar Sheet ---');
    const balanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balance - Calendar');
    if (!balanceSheet) {
      console.log('❌ Balance - Calendar sheet not found');
      return "ERROR:Balance - Calendar sheet not found";
    }
    console.log('✅ Balance - Calendar sheet found');
    
    // OPTIMIZATION: Find last column with data in Row 1
    console.log('\n--- OPTIMIZATION: Finding Last Column with Data ---');
    const fullRow1 = balanceSheet.getRange(1, 1, 1, balanceSheet.getLastColumn()).getValues()[0];
    let lastColWithData = balanceSheet.getLastColumn();
    
    // Find first empty column starting from the end
    for (let i = fullRow1.length - 1; i >= 2; i--) {
      if (fullRow1[i] && fullRow1[i].toString().trim() !== '') {
        lastColWithData = i + 1;
        break;
      }
    }
    
    console.log('Sheet last column:', balanceSheet.getLastColumn());
    console.log('Last column with data in Row 1:', lastColWithData);
    console.log('Columns to scan reduced from', balanceSheet.getLastColumn() - 2, 'to', lastColWithData - 2);
    
    // OPTIMIZATION: Find last row with data in Column B (dates)
    console.log('\n--- OPTIMIZATION: Finding Last Row with Dates ---');
    const fullColB = balanceSheet.getRange(5, 2, balanceSheet.getLastRow() - 4, 1).getValues();
    let lastRowWithData = balanceSheet.getLastRow();
    
    // Find first empty row starting from the end
    for (let i = fullColB.length - 1; i >= 0; i--) {
      if (fullColB[i][0] && fullColB[i][0] !== '') {
        lastRowWithData = i + 5; // +5 because data starts at row 5
        break;
      }
    }
    
    console.log('Sheet last row:', balanceSheet.getLastRow());
    console.log('Last row with date in Column B:', lastRowWithData);
    console.log('Date rows to scan reduced from', balanceSheet.getLastRow() - 4, 'to', lastRowWithData - 4);
    
    // Find employee columns in Balance-Calendar (OPTIMIZED)
    console.log('\n--- Finding Employee and Leave Type Column (Optimized Scan) ---');
    console.log('Looking for Employee ID:', employeeId);
    
    const headerRow1 = balanceSheet.getRange(1, 3, 1, lastColWithData - 2).getValues()[0];
    const headerRow3 = balanceSheet.getRange(3, 3, 1, lastColWithData - 2).getValues()[0];
    
    console.log('Scanning columns from C to', String.fromCharCode(65 + lastColWithData - 1));
    console.log('Total columns to scan:', headerRow1.length);
    
    let balanceColumn = -1;
    const leaveTypeMap = {
      'annual': 'Annual Leaves',
      'sick': 'Sick Leaves',
      'emergency': 'Emergency Leaves'
    };
    
    const targetLeaveType = leaveTypeMap[leaveType];
    console.log('Looking for Leave Type:', targetLeaveType);
    
    // Find the column for this employee and leave type
    for (let i = 0; i < headerRow1.length; i++) {
      const cellEmpId = String(headerRow1[i]);
      const cellLeaveType = String(headerRow3[i]);
      
      if (cellEmpId === String(employeeId)) {
        console.log(`  Column ${String.fromCharCode(65 + i + 2)} (index ${i}): Employee ID match → ${cellEmpId}, Leave Type: ${cellLeaveType}`);
        
        if (cellLeaveType === targetLeaveType) {
          balanceColumn = i + 3; // +3 because we started from column C (3)
          console.log(`  ✅ MATCH FOUND at Column ${String.fromCharCode(65 + balanceColumn - 1)} (${balanceColumn})`);
          break;
        }
      }
    }
    
    if (balanceColumn === -1) {
      console.log('❌ Could not find balance column for employee');
      console.log('Employee ID searched:', employeeId);
      console.log('Leave type searched:', targetLeaveType);
      return "ERROR:Could not find balance information for employee";
    }
    
    // Find the row for start date (OPTIMIZED)
    console.log('\n--- Finding Start Date Row (Optimized Scan) ---');
    console.log('Looking for date:', start.toISOString().split('T')[0]);
    
    const dateColumn = balanceSheet.getRange(5, 2, lastRowWithData - 4, 1).getValues();
    let balanceRow = -1;
    
    console.log('Scanning dates from row 5 to row', lastRowWithData);
    console.log('Total date rows to scan:', dateColumn.length);
    
    for (let i = 0; i < dateColumn.length; i++) {
      const cellDate = new Date(dateColumn[i][0]);
      cellDate.setHours(0, 0, 0, 0);
      const startDateCheck = new Date(start);
      startDateCheck.setHours(0, 0, 0, 0);
      
      if (cellDate.getTime() === startDateCheck.getTime()) {
        balanceRow = i + 5; // +5 because data starts at row 5
        console.log(`✅ Date match found at Row ${balanceRow}: ${cellDate.toISOString().split('T')[0]}`);
        break;
      }
    }
    
    if (balanceRow === -1) {
      console.log('❌ Could not find start date in Balance-Calendar sheet');
      console.log('Start date searched:', start.toISOString().split('T')[0]);
      return "ERROR:Could not find balance for start date";
    }
    
    // Get balance value
    console.log('\n--- Reading Balance Value ---');
    console.log('Reading from Cell:', String.fromCharCode(65 + balanceColumn - 1) + balanceRow);
    console.log('  Row:', balanceRow);
    console.log('  Column:', balanceColumn, '(' + String.fromCharCode(65 + balanceColumn - 1) + ')');
    
    const currentBalance = balanceSheet.getRange(balanceRow, balanceColumn).getValue() || 0;
    console.log('✅ Current balance retrieved:', currentBalance);
    
    // Get on-hold balance
    console.log('\n--- Getting On-Hold Balance ---');
    console.log('Calling getSimpleOnHoldCounts() for Employee ID:', employeeId);
    
    const onHoldCounts = getSimpleOnHoldCounts(employeeId);
    console.log('On-hold counts returned:', JSON.stringify(onHoldCounts));
    
    let onHoldBalance = 0;
    
    if (leaveType === 'annual') {
      onHoldBalance = onHoldCounts.annual || 0;
      console.log('Using annual on-hold:', onHoldBalance);
    } else if (leaveType === 'sick') {
      onHoldBalance = onHoldCounts.sick || 0;
      console.log('Using sick on-hold:', onHoldBalance);
    } else if (leaveType === 'emergency') {
      onHoldBalance = onHoldCounts.emergency || 0;
      console.log('Using emergency on-hold:', onHoldBalance);
    }
    
    console.log('✅ On-hold balance for', leaveType + ':', onHoldBalance);
    
    // Calculate net available balance
    console.log('\n--- Net Balance Calculation ---');
    console.log('Formula: Net Available = Current Balance - On-Hold Balance');
    console.log('  Current Balance:', currentBalance);
    console.log('  On-Hold Balance:', onHoldBalance);
    
    const netAvailableBalance = currentBalance - onHoldBalance;
    console.log('  NET AVAILABLE BALANCE:', netAvailableBalance);
    
    // VALIDATION: Check if net days requested <= net available balance
    console.log('\n--- Validation Check ---');
    console.log('Comparing:');
    console.log('  Net Days Requested:', netDaysRequested);
    console.log('  Net Available Balance:', netAvailableBalance);
    console.log('  Condition:', netDaysRequested, '<=', netAvailableBalance, '?');
    
    if (netDaysRequested > netAvailableBalance) {
      console.log('❌ VALIDATION FAILED: Insufficient balance');
      console.log('\nValidation Details:');
      console.log('  Requested:', netDaysRequested, 'net days');
      console.log('  Available:', netAvailableBalance, 'days');
      console.log('  Shortage:', (netDaysRequested - netAvailableBalance), 'days');
      console.log('  Current Balance:', currentBalance);
      console.log('  On-Hold:', onHoldBalance);
      
      const errorMsg = `Insufficient balance. You are requesting ${netDaysRequested} net days but only have ${netAvailableBalance} days available (Current: ${currentBalance}, On-Hold: ${onHoldBalance})`;
      console.log('\n' + '='.repeat(70));
      console.log('=== BALANCE VALIDATION END (FAILED) ===');
      console.log('='.repeat(70));
      return "ERROR:" + errorMsg;
    }
    
    console.log('✅ VALIDATION PASSED');
    console.log('  Request is within available balance');
    console.log('  Remaining after approval:', (netAvailableBalance - netDaysRequested), 'days');
    console.log('\n' + '='.repeat(70));
    console.log('=== BALANCE VALIDATION END (SUCCESS) ===');
    console.log('='.repeat(70));
    
    // Generate new request ID
    console.log('\n--- Generating Request ID ---');
    const requestId = generateNextRequestId(sheet, info.prefix);
    if (!requestId) {
      console.log('❌ Failed to generate request ID');
      return "ERROR:Failed to generate request ID";
    }
    
    console.log('✅ Generated Request ID:', requestId);
    
    const requestTimestamp = new Date();
    console.log('Request timestamp:', requestTimestamp.toISOString());
    
    // Prepare rows for insertion with date-by-date weekly holiday check
    console.log('\n--- Preparing Data Rows ---');
    const rows = [];
    
    dates.forEach((date, index) => {
      const weekDay = date.toLocaleDateString('en-US', { weekday: 'short' });
      const fullDayName = dayMapping[weekDay];
      
      const holidayInfo = getWeeklyHolidayForSpecificDate(employeeId, date);
      
      let isWeeklyHoliday = false;

      if (holidayInfo.success && holidayInfo.hasHoliday) {
        isWeeklyHoliday = holidayInfo.holidayDays[fullDayName] === true;
      }
      
      rows.push([
        requestId,                               // A - Request ID
        employeeId,                              // B - Employee ID
        employeeName,                            // C - Employee Name
        date,                                    // D - Date
        weekDay,                                 // E - Week day (Ddd format)
        reason,                                  // F - Reason
        isWeeklyHoliday ? 'True' : '',          // G - Weekly holiday
        requestTimestamp,                        // H - Request timestamp
        'Pending'                                // I - Response status
      ]);
      
      console.log(`  Row ${index + 1}: ${date.toISOString().split('T')[0]} (${weekDay}) - Weekly Holiday: ${isWeeklyHoliday ? 'Yes' : 'No'}`);
    });
    
    console.log('✅ Prepared', rows.length, 'data rows');
    
    // Insert rows (9 columns: A to I)
    console.log('\n--- Inserting Data to Sheet ---');
    const lastRow = sheet.getLastRow();
    console.log('Last row in sheet:', lastRow);
    console.log('Inserting at row:', lastRow + 1);
    console.log('Number of rows:', rows.length);
    console.log('Number of columns:', 9);
    
    sheet.getRange(lastRow + 1, 1, rows.length, 9).setValues(rows);
    
    console.log('✅ Data inserted successfully');
    console.log('Row range:', (lastRow + 1), 'to', (lastRow + rows.length));
    
    console.log('\n' + '='.repeat(70));
    console.log('=== SUBMIT LEAVE REQUEST END (SUCCESS) ===');
    console.log('='.repeat(70));

    // Send notification to admins
    console.log('\n--- Sending Email Notifications ---');
    try {
      const notificationAdmins = getNotificationAdmins();
      
      if (notificationAdmins && notificationAdmins.length > 0) {
        console.log('Found', notificationAdmins.length, 'notification admins');
        
        // Map leave type to display name
        const leaveTypeDisplay = leaveType === 'annual' ? 'Annual' : 
                                leaveType === 'sick' ? 'Sick' : 'Emergency';
        
        // Prepare email template - USE STRING PARAMETERS
        const emailTemplate = sendLeaveRequestNotification(
          String(employeeResult.data.name),        // employeeName
          String(employeeResult.data.id),          // employeeId
          String(leaveTypeDisplay),                // leaveType
          String(startDate),                       // startDate
          String(endDate),                         // endDate
          String(dates.length),                    // duration (total days)
          String(reason),                          // reason
          String(requestTimestamp.toISOString())   // requestDate
        );
        
        // Send to all notification admins
        const sendResult = sendToMultipleAdmins(emailTemplate, notificationAdmins);
        console.log('Email send result:', sendResult.sent, 'sent,', sendResult.failed, 'failed');
        
      } else {
        console.log('⚠️ No notification admins found');
      }
    } catch (emailError) {
      console.error('❌ Email notification error:', emailError);
      // Don't fail the request if email fails
    }

    console.log('\n' + '='.repeat(70));
    console.log('=== SUBMIT LEAVE REQUEST END (SUCCESS) ===');
    console.log('='.repeat(70));

    return "SUCCESS:Request submitted successfully. Request ID: " + requestId;
    
  } catch (error) {
    console.log('\n' + '='.repeat(70));
    console.log('=== SUBMIT LEAVE REQUEST END (ERROR) ===');
    console.log('='.repeat(70));
    console.error('❌ EXCEPTION:', error);
    console.error('Error name:', error.name);
    console.error('Error message:', error.message);
    console.error('Stack trace:', error.stack);
    return "ERROR:System error - " + error.toString();
  }
}

function getWeeklyHolidayForSpecificDate(employeeId, targetDate) {
  try {
    const weeklySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Weekly Holidays');
    
    if (!weeklySheet) {
      return {
        success: true,
        hasHoliday: false,
        holidayDays: {}
      };
    }
    
    // Get system end date from Config sheet
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    let systemEndDate = new Date('2099-12-31'); // Default far future
    
    if (configSheet) {
      const configEndDate = configSheet.getRange('B4').getValue();
      if (configEndDate) {
        systemEndDate = new Date(configEndDate);
        systemEndDate.setHours(0, 0, 0, 0);
      }
    }
    
    const data = weeklySheet.getDataRange().getValues();
    const checkDate = new Date(targetDate);
    checkDate.setHours(0, 0, 0, 0);
    
    let tempPeriod = null;
    let standardPeriod = null;
    
    // Scan all periods for this employee
    for (let i = 1; i < data.length; i++) {
      const empId = String(data[i][1]); // Column B - Employee ID
      const startDate = data[i][3] ? new Date(data[i][3]) : null; // Column D
      let endDate = data[i][4] ? new Date(data[i][4]) : null; // Column E
      const status = String(data[i][5]); // Column F - Status
      const type = String(data[i][6]); // Column G - Period Type
      
      // Only check Active periods for this employee
      if (empId !== String(employeeId) || status !== 'Active' || !startDate) {
        continue;
      }
      
      startDate.setHours(0, 0, 0, 0);
      
      // If end date is blank (for standard periods), use system end date
      if (!endDate && type === 'standard') {
        endDate = new Date(systemEndDate);
      }
      
      if (endDate) {
        endDate.setHours(0, 0, 0, 0);
      }
      
      // Check if target date falls within this period
      const isInPeriod = checkDate >= startDate && (!endDate || checkDate <= endDate);
      
      if (isInPeriod) {
        // Extract holiday days from columns H-N
        const holidayDays = {
          friday: data[i][7] === true || String(data[i][7]).toLowerCase() === 'true',
          saturday: data[i][8] === true || String(data[i][8]).toLowerCase() === 'true',
          sunday: data[i][9] === true || String(data[i][9]).toLowerCase() === 'true',
          monday: data[i][10] === true || String(data[i][10]).toLowerCase() === 'true',
          tuesday: data[i][11] === true || String(data[i][11]).toLowerCase() === 'true',
          wednesday: data[i][12] === true || String(data[i][12]).toLowerCase() === 'true',
          thursday: data[i][13] === true || String(data[i][13]).toLowerCase() === 'true'
        };
        
        // Store temp or standard period (temp has priority)
        if (type === 'temp') {
          tempPeriod = {
            type: type,
            holidayDays: holidayDays,
            periodId: data[i][0]
          };
        } else if (type === 'standard') {
          standardPeriod = {
            type: type,
            holidayDays: holidayDays,
            periodId: data[i][0]
          };
        }
      }
    }
    
    // Priority: Temp period > Standard period
    const activePeriod = tempPeriod || standardPeriod;
    
    if (!activePeriod) {
      return {
        success: true,
        hasHoliday: false,
        holidayDays: {}
      };
    }
    
    return {
      success: true,
      hasHoliday: true,
      holidayDays: activePeriod.holidayDays,
      periodType: activePeriod.type,
      periodId: activePeriod.periodId
    };
    
  } catch (error) {
    console.error('Error getting weekly holiday for date:', error);
    return {
      success: false,
      hasHoliday: false,
      holidayDays: {},
      error: error.toString()
    };
  }
}

function generateNextRequestId(sheet, prefix) {
  try {
    const data = sheet.getDataRange().getValues();
    let maxId = 0;
    
    // Scan all rows starting from row 2 (skip header)
    for (let i = 1; i < data.length; i++) {
      const requestId = data[i][0]; // Column A
      
      if (requestId && requestId.toString().startsWith(prefix)) {
        // Extract numeric part after prefix
        const numPart = parseInt(requestId.toString().substring(1));
        
        if (!isNaN(numPart) && numPart > maxId) {
          maxId = numPart;
        }
      }
    }
    
    // Generate next ID with 2-digit zero-padding
    const nextId = maxId + 1;
    return prefix + String(nextId).padStart(2, '0');
    
  } catch (error) {
    console.error('Error generating request ID:', error);
    return null; // Return null instead of fallback
  }
}

function validateLeaveRequestDates(employee, leaveType, startDate, endDate) {
  try {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Check if dates are in the past (for annual leave)
    if (leaveType === 'annual' && startDate <= today) {
      return {
        success: false,
        error: 'Annual leave cannot be requested for past or current dates'
      };
    }
    
    // Validate against employee's usable period
    const usablePeriodResult = validateEmployeeUsablePeriod(employee.email);
    if (usablePeriodResult.success && usablePeriodResult.data) {
      const usablePeriod = usablePeriodResult.data;
      
      if (!usablePeriod.isWithinPeriod) {
        return {
          success: false,
          error: `Leave can only be requested between ${formatDateForDisplay(usablePeriod.usableFrom)} and ${formatDateForDisplay(usablePeriod.usableTo)}`
        };
      }
      
      // Check if dates fall within usable period
      if (startDate < usablePeriod.usableFrom || endDate > usablePeriod.usableTo) {
        return {
          success: false,
          error: `Dates must be within your usable period: ${formatDateForDisplay(usablePeriod.usableFrom)} to ${formatDateForDisplay(usablePeriod.usableTo)}`
        };
      }
    }
    
    // Check for overlapping requests
    const overlapResult = checkForOverlappingRequests(employee.id, startDate, endDate, leaveType);
    if (!overlapResult.success) {
      return overlapResult;
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error validating dates:', error);
    return {
      success: false,
      error: 'Date validation failed'
    };
  }
}

/**
 * Generate unique leave request ID
 */
function generateLeaveRequestId(leaveType) {
  try {
    const prefix = leaveType.charAt(0).toUpperCase(); // A, S, E
    const timestamp = new Date().getTime().toString().slice(-8); // Last 8 digits
    const random = Math.floor(Math.random() * 100).toString().padStart(2, '0');
    
    return `${prefix}${timestamp}${random}`;
    
  } catch (error) {
    console.error('Error generating request ID:', error);
    // Fallback ID generation
    return leaveType.charAt(0).toUpperCase() + Date.now().toString().slice(-6);
  }
}

/**
 * Check for official holidays in date range
 */
function checkOfficialHolidaysInRange(dates) {
  try {
    const holidaysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('official holidays');
    if (!holidaysSheet) {
      return { hasHolidays: false, holidays: [] };
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const holidayDates = [];
    const holidayNames = [];
    
    for (let i = 3; i < data.length; i++) { // Data starts at row 4
      if (data[i][1]) {
        holidayDates.push(new Date(data[i][1]).toDateString());
        holidayNames.push(data[i][0]);
      }
    }
    
    const conflictingHolidays = [];
    dates.forEach(date => {
      const dateString = date.toDateString();
      const index = holidayDates.indexOf(dateString);
      if (index !== -1) {
        conflictingHolidays.push(holidayNames[index]);
      }
    });
    
    return {
      hasHolidays: conflictingHolidays.length > 0,
      holidays: conflictingHolidays
    };
    
  } catch (error) {
    console.error('Error checking official holidays:', error);
    return { hasHolidays: false, holidays: [] };
  }
}

/**
 * Get calendar data for date picker (other employees' data)
 */
function getCalendarData(employeeEmail, startDate, endDate) {
  try {
    const currentEmployeeResult = getCurrentEmployeeData(employeeEmail);
    if (!currentEmployeeResult.success) {
      return "ERROR:" + currentEmployeeResult.error;
    }
    
    const currentEmployeeId = currentEmployeeResult.data.id;
    const fullCalendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Full-calendar');
    
    if (!fullCalendarSheet) {
      return "SUCCESS:EMPTY";
    }
    
    const start = new Date(startDate);
    const end = new Date(endDate);
    const otherEmployeeData = [];
    
    // Get headers (employee IDs in row 1, leave types in row 3)
    const row1 = fullCalendarSheet.getRange(1, 3, 1, fullCalendarSheet.getLastColumn() - 2).getValues()[0];
    const row3 = fullCalendarSheet.getRange(3, 3, 1, fullCalendarSheet.getLastColumn() - 2).getValues()[0];
    
    // Get date data (starts from row 5)
    const dateData = fullCalendarSheet.getRange(5, 2, fullCalendarSheet.getLastRow() - 4, fullCalendarSheet.getLastColumn() - 1).getValues();
    
    dateData.forEach((row, dateIndex) => {
      const date = new Date(row[0]);
      if (date >= start && date <= end) {
        
        for (let colIndex = 1; colIndex < row.length; colIndex++) {
          const employeeId = row1[colIndex - 1];
          const leaveType = row3[colIndex - 1];
          const value = row[colIndex];
          
          if (employeeId && employeeId != currentEmployeeId && value) {
            let isDayOff = false;
            
            if (leaveType === 'Weekly Holidays' || leaveType === 'Official Holidays') {
              isDayOff = value === true || value === 'true';
            } else if (['Annual Leaves', 'Sick Leaves', 'Emergency Leaves'].includes(leaveType)) {
              isDayOff = value.toString().includes('Approved');
            }
            
            if (isDayOff) {
              otherEmployeeData.push(`${date.toISOString()}|${employeeId}|${leaveType}`);
            }
          }
        }
      }
    });
    
    return "SUCCESS:" + otherEmployeeData.join(';');
    
  } catch (error) {
    console.error('Error getting calendar data:', error);
    return "ERROR:System error - " + error.toString();
  }
}

function getSystemDateLimits() {
  try {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!configSheet) {
      return {
        success: false,
        error: 'Config sheet not found'
      };
    }
    
    const startDate = new Date(configSheet.getRange('B3').getValue());
    const endDate = new Date(configSheet.getRange('B4').getValue());
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      return {
        success: false,
        error: 'Invalid system dates in Config sheet'
      };
    }
    
    return {
      success: true,
      startDate: startDate,
      endDate: endDate
    };
    
  } catch (error) {
    console.error('Error getting system date limits:', error);
    return {
      success: false,
      error: 'Error reading system dates: ' + error.toString()
    };
  }
}

/**
 * NEW: Check available balance for leave request
 */
function checkAvailableBalance(employeeId, leaveType, startDate, endDate) {
  try {
    // Get balance from Balance-Calendar sheet on start date
    const balance = getBalanceFromCalendarSheet(employeeId, leaveType, startDate);
    if (balance === null) {
      return {
        success: false,
        error: 'Could not retrieve balance information'
      };
    }
    
    // Calculate on-hold balance
    const onHoldBalance = calculateOnHoldBalance(employeeId, leaveType);
    
    // Calculate net available balance
    const netAvailable = balance - onHoldBalance;
    
    // Calculate requested duration
    const requestedDuration = calculateRequestDuration(startDate, endDate);
    
    if (requestedDuration > netAvailable) {
      return {
        success: false,
        error: `Insufficient balance. Available: ${netAvailable} days, Requested: ${requestedDuration} days`
      };
    }
    
    return {
      success: true,
      balance: balance,
      onHold: onHoldBalance,
      netAvailable: netAvailable,
      requested: requestedDuration
    };
    
  } catch (error) {
    console.error('Error checking available balance:', error);
    return {
      success: false,
      error: 'Error checking balance: ' + error.toString()
    };
  }
}

/**
 * NEW: Get balance from Balance-Calendar sheet on specific date
 */
function getBalanceFromCalendarSheet(employeeId, leaveType, date) {
  try {
    const balanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Balance - Calendar');
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
    if (leaveType === 'sick') {
      leaveTypeOffset = 1;
    } else if (leaveType === 'emergency') {
      leaveTypeOffset = 2;
    }
    
    const balanceCol = employeeStartCol + leaveTypeOffset;
    
    // Find date row (dates start from row 5 = index 4)
    const targetDate = new Date(date);
    targetDate.setHours(0, 0, 0, 0);
    let dateRow = -1;
    
    for (let row = 4; row < data.length; row++) {
      const cellDate = new Date(data[row][1]); // Column B
      cellDate.setHours(0, 0, 0, 0);
      if (cellDate.getTime() === targetDate.getTime()) {
        dateRow = row;
        break;
      }
    }
    
    if (dateRow === -1) {
      console.warn(`Date ${date} not found in Balance-Calendar`);
      return null;
    }
    
    // Get balance value
    const balance = data[dateRow][balanceCol];
    return typeof balance === 'number' ? balance : parseInt(balance) || 0;
    
  } catch (error) {
    console.error('Error getting balance from calendar sheet:', error);
    return null;
  }
}

/**
 * NEW: Calculate on-hold balance (approved future requests of same type)
 */
function calculateOnHoldBalance(employeeId, leaveType) {
  try {
    const sheetNames = {
      annual: 'Annual leaves',
      sick: 'Sick leaves', 
      emergency: 'Emergency leaves'
    };
    
    const sheetName = sheetNames[leaveType];
    if (!sheetName) return 0;
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) return 0;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let onHoldDays = 0;
    const processedRequests = new Set();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Skip empty rows
      if (!row[0] || !row[1]) continue;
      
      const requestId = String(row[0]);
      const rowEmployeeId = String(row[1]);
      const leaveDate = new Date(row[3]); // Column D
      const responseStatus = String(row[8] || 'Pending'); // Column I
      const isWeeklyHoliday = String(row[6] || '').toLowerCase() === 'true'; // Column G
      
      // Only count for same employee
      if (rowEmployeeId !== String(employeeId)) continue;
      
      // Only count approved requests
      if (responseStatus !== 'Approved') continue;
      
      // Only count future dates
      leaveDate.setHours(0, 0, 0, 0);
      if (leaveDate <= today) continue;
      
      // Only count non-weekly holiday days (net days)
      if (isWeeklyHoliday) continue;
      
      // Count each day once per request
      if (!processedRequests.has(requestId + '-' + leaveDate.toISOString())) {
        onHoldDays++;
        processedRequests.add(requestId + '-' + leaveDate.toISOString());
      }
    }
    
    return onHoldDays;
    
  } catch (error) {
    console.error('Error calculating on-hold balance:', error);
    return 0;
  }
}

/**
 * NEW: Calculate request duration (end - start + 1)
 */
function calculateRequestDuration(startDate, endDate) {
  try {
    const start = new Date(startDate);
    const end = new Date(endDate);
    
    start.setHours(0, 0, 0, 0);
    end.setHours(0, 0, 0, 0);
    
    const timeDiff = end.getTime() - start.getTime();
    const daysDiff = Math.ceil(timeDiff / (1000 * 60 * 60 * 24));
    
    return daysDiff + 1; // Include both start and end dates
    
  } catch (error) {
    console.error('Error calculating request duration:', error);
    return 0;
  }
}

/**
 * FIXED: Get employee weekly holidays for a specific period
 */
function getEmployeeWeeklyHolidaysForPeriod(employeeId, startDate, endDate) {
  try {
    const weeklySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Weekly Holidays');
    
    if (!weeklySheet) {
      return {
        success: true,
        isWeeklyHoliday: function() { return false; }
      };
    }
    
    const data = weeklySheet.getDataRange().getValues();
    const holidayPeriods = [];
    
    // Process all periods for this employee
    for (let i = 1; i < data.length; i++) {
      const empId = data[i][1]; // Column B
      const periodStartDate = data[i][3] ? new Date(data[i][3]) : null; // Column D
      const periodEndDate = data[i][4] ? new Date(data[i][4]) : null; // Column E
      const status = data[i][5]; // Column F
      const type = data[i][6]; // Column G
      
      if (String(empId) === String(employeeId) && status === 'Active') {
        // Extract holiday days (columns H-N: Fri, Sat, Sun, Mon, Tue, Wed, Thu)
        const holidayDays = {
          friday: data[i][7] === 'true' || data[i][7] === true,
          saturday: data[i][8] === 'true' || data[i][8] === true,
          sunday: data[i][9] === 'true' || data[i][9] === true,
          monday: data[i][10] === 'true' || data[i][10] === true,
          tuesday: data[i][11] === 'true' || data[i][11] === true,
          wednesday: data[i][12] === 'true' || data[i][12] === true,
          thursday: data[i][13] === 'true' || data[i][13] === true
        };
        
        holidayPeriods.push({
          startDate: periodStartDate,
          endDate: periodEndDate,
          type: type,
          holidayDays: holidayDays
        });
      }
    }
    
    // Function to check if a specific date is a weekly holiday
    const isWeeklyHolidayFunction = function(date) {
      const checkDate = new Date(date);
      checkDate.setHours(0, 0, 0, 0);
      
      // Find the active period for this date
      let activePeriod = null;
      
      for (const period of holidayPeriods) {
        const periodStart = new Date(period.startDate);
        const periodEnd = period.endDate ? new Date(period.endDate) : null;
        
        periodStart.setHours(0, 0, 0, 0);
        if (periodEnd) periodEnd.setHours(0, 0, 0, 0);
        
        const isDateInPeriod = checkDate >= periodStart && (!periodEnd || checkDate <= periodEnd);
        
        if (isDateInPeriod) {
          // Temp periods override standard periods
          if (period.type === 'temp' || !activePeriod) {
            activePeriod = period;
          }
        }
      }
      
      if (!activePeriod) return false;
      
      // Check if the day of week is a holiday
      const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
      const dayOfWeek = dayNames[checkDate.getDay()];
      
      return activePeriod.holidayDays[dayOfWeek] || false;
    };
    
    return {
      success: true,
      isWeeklyHoliday: isWeeklyHolidayFunction,
      periods: holidayPeriods
    };
    
  } catch (error) {
    console.error('Error getting weekly holidays for period:', error);
    return {
      success: false,
      error: 'Error retrieving weekly holidays: ' + error.toString(),
      isWeeklyHoliday: function() { return false; }
    };
  }
}

function calculateRequestStatus(approved, rejected, pending) {
  const total = approved + rejected + pending;
  if (total === 0) return 'Pending';
  
  const uniqueStatuses = [];
  if (pending > 0) uniqueStatuses.push('Pending');
  if (approved > 0) uniqueStatuses.push('Approved');
  if (rejected > 0) uniqueStatuses.push('Rejected');
  
  if (pending === total) return 'Pending';
  else if (approved === total) return 'Approved';
  else if (rejected === total) return 'Rejected';
  else if (uniqueStatuses.length >= 2) return 'Partial';
  else return 'Pending';
}

/**
 * Calculate date range details including weekends and holidays
 */
function calculateDateRangeDetails(startDate, endDate, employeeId) {
  try {
    const details = {
      totalDays: 0,
      weeklyHolidayDays: 0,
      netDays: 0,
      dateList: []
    };
    
    const currentDate = new Date(startDate);
    
    // Get employee weekly holidays
    const weeklyHolidays = getEmployeeWeeklyHolidayDays(employeeId);
    
    while (currentDate <= endDate) {
      const dayName = currentDate.toLocaleDateString('en-US', { weekday: 'long' }).toLowerCase();
      const isWeeklyHoliday = weeklyHolidays.includes(dayName);
      
      details.totalDays++;
      details.dateList.push({
        date: new Date(currentDate),
        dayName: dayName,
        isWeeklyHoliday: isWeeklyHoliday
      });
      
      if (isWeeklyHoliday) {
        details.weeklyHolidayDays++;
      } else {
        details.netDays++;
      }
      
      currentDate.setDate(currentDate.getDate() + 1);
    }
    
    return details;
    
  } catch (error) {
    console.error('Error calculating date details:', error);
    // Return basic calculation as fallback
    const diffTime = Math.abs(endDate - startDate);
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;
    
    return {
      totalDays: diffDays,
      weeklyHolidayDays: 0,
      netDays: diffDays,
      dateList: []
    };
  }
}

/**
 * Get leave sheet name based on leave type
 */
function getLeaveSheetName(leaveType) {
  const sheetNames = {
    'annual': 'Annual leaves',
    'sick': 'Sick leaves',
    'emergency': 'Emergency leaves'
  };
  
  return sheetNames[leaveType] || 'Annual leaves';
}

/**
 * Save leave request data to appropriate sheet
 */
function saveLeaveRequestToSheet(sheet, requestData) {
  try {
    const { requestId, employee, startDate, endDate, reason, requestTimestamp, dateDetails } = requestData;
    
    // Prepare rows for each date in the range
    const rowsToAdd = [];
    
    dateDetails.dateList.forEach(dateInfo => {
      const row = [
        requestId,                                    // Column A - Request ID
        employee.id,                                  // Column B - Employee ID
        employee.name,                                // Column C - Employee Name
        dateInfo.date,                                // Column D - Leave Date
        dateInfo.dayName,                             // Column E - Week Day
        reason,                                       // Column F - Reason
        dateInfo.isWeeklyHoliday.toString(),          // Column G - Weekly Holiday
        requestTimestamp,                             // Column H - Request Timestamp
        'Pending',                                    // Column I - Response Status
        '',                                           // Column J - Response Timestamp
        '',                                           // Column K - Responded By
        '',                                           // Column L - Notification Status
        '',                                           // Column M - Admin Comment
        'Not yet'                                     // Column N - Used Status
      ];
      
      rowsToAdd.push(row);
    });
    
    // Add all rows at once for better performance
    if (rowsToAdd.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      const range = sheet.getRange(startRow, 1, rowsToAdd.length, rowsToAdd[0].length);
      range.setValues(rowsToAdd);
      
      // Format date columns
      const dateRange = sheet.getRange(startRow, 4, rowsToAdd.length, 1); // Column D
      dateRange.setNumberFormat('MM/dd/yyyy');
      
      const timestampRange = sheet.getRange(startRow, 8, rowsToAdd.length, 1); // Column H
      timestampRange.setNumberFormat('MM/dd/yyyy HH:mm:ss');
    }
    
    return {
      success: true,
      rowsAdded: rowsToAdd.length,
      requestId: requestId
    };
    
  } catch (error) {
    console.error('Error saving to sheet:', error);
    return {
      success: false,
      error: 'Failed to save request data'
    };
  }
}

/**
 * Check for overlapping leave requests
 */
function checkForOverlappingRequests(employeeId, startDate, endDate, leaveType) {
  try {
    const sheets = ['Annual leaves', 'Sick leaves', 'Emergency leaves'];
    
    for (const sheetName of sheets) {
      const sheet = getSheet(sheetName);
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowEmployeeId = String(row[1]); // Column B
        const rowDate = new Date(row[3]); // Column D
        const responseStatus = String(row[8] || 'Pending'); // Column I
        
        // Skip if different employee or rejected request
        if (rowEmployeeId !== String(employeeId) || responseStatus === 'Rejected') {
          continue;
        }
        
        // Check for date overlap
        if (rowDate >= startDate && rowDate <= endDate) {
          return {
            success: false,
            error: `You already have a ${sheetName.replace(' leaves', '')} leave request on ${formatDateForDisplay(rowDate)}`
          };
        }
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error checking overlapping requests:', error);
    return { success: true }; // Allow request if check fails
  }
}

/**
 * Get employee weekly holiday days
 */
function getEmployeeWeeklyHolidayDays(employeeId) {
  try {
    // This function should return array of weekly holiday day names
    // Implementation depends on your weekly holiday storage structure
    // For now, return empty array as fallback
    return [];
    
  } catch (error) {
    console.error('Error getting weekly holidays:', error);
    return [];
  }
}

/**
 * Send leave request notification to notification admins
 */
function sendLeaveRequestNotificationToAdmins(requestData) {
  try {
    // Use existing email notification function if available
    if (typeof sendLeaveRequestNotification === 'function') {
      return sendLeaveRequestNotification(requestData);
    }
    
    console.log('Leave request notification would be sent:', requestData);
    return { success: true };
    
  } catch (error) {
    console.error('Error sending notifications:', error);
    throw error;
  }
}

/* ============================================================================
   WEEKLY HOLIDAY CALENDAR FUNCTIONS FOR EMPLOYEE PORTAL
   UPDATED: Fixed priority logic and added date information
   ============================================================================ */

/**
 * Get weekly holiday calendar data for all active employees
 * Returns: SUCCESS:currentEmpData;otherEmployeesData OR ERROR:message
 * ALL STRING FORMAT - NO OBJECTS
 */
function getWeeklyHolidayCalendarData(employeeEmail) {
  try {
    console.log('=== GET WEEKLY HOLIDAY CALENDAR START ===');
    console.log('Employee Email:', employeeEmail);
    
    if (!employeeEmail) {
      return "ERROR:Employee email is required";
    }
    
    employeeEmail = String(employeeEmail).trim();
    
    const employeesSheet = getSheet('Employees');
    const weeklyHolidaysSheet = getSheet('Weekly Holidays');
    
    if (!employeesSheet || !weeklyHolidaysSheet) {
      return "ERROR:Required sheets not found";
    }
    
    const employeesData = employeesSheet.getDataRange().getValues();
    const holidaysData = weeklyHolidaysSheet.getDataRange().getValues();
    const currentDate = new Date();
    
    const activeEmployees = [];
    let currentEmployeeId = null;
    
    for (let i = 1; i < employeesData.length; i++) {
      const row = employeesData[i];
      const empId = String(row[0] || '');
      const empName = String(row[1] || '');
      const empEmail = String(row[2] || '').trim();
      const deactivatedOn = row[9];
      const reactivatedOn = row[17];
      
      if (!empId || !empName) continue;
      
      let isActive = false;
      if (!deactivatedOn) {
        isActive = true;
      } else if (reactivatedOn && new Date(reactivatedOn) > new Date(deactivatedOn)) {
        isActive = true;
      }
      
      if (isActive) {
        activeEmployees.push({
          id: empId,
          name: empName,
          email: empEmail
        });
        
        if (empEmail.toLowerCase() === employeeEmail.toLowerCase()) {
          currentEmployeeId = empId;
        }
      }
    }
    
    if (!currentEmployeeId) {
      return "ERROR:Current employee not found";
    }
    
    console.log('Found ' + activeEmployees.length + ' active employees');
    console.log('Current employee ID:', currentEmployeeId);
    
    const employeeHolidayStrings = [];
    
    for (let emp of activeEmployees) {
      const holidays = getEmployeeWeeklyHolidayStatus(emp.id, holidaysData, currentDate);
      const empString = formatEmployeeHolidayString(emp.id, emp.name, holidays);
      employeeHolidayStrings.push({
        id: emp.id,
        string: empString,
        isCurrent: emp.id === currentEmployeeId
      });
    }
    
    let currentEmpStr = '';
    const otherEmpStrings = [];
    
    for (let empData of employeeHolidayStrings) {
      if (empData.isCurrent) {
        currentEmpStr = empData.string;
      } else {
        otherEmpStrings.push({
          id: empData.id,
          string: empData.string
        });
      }
    }
    
    otherEmpStrings.sort(function(a, b) {
      return parseInt(a.id) - parseInt(b.id);
    });
    
    const otherEmpStr = otherEmpStrings.map(function(e) { return e.string; }).join(';');
    
    console.log('Current employee string:', currentEmpStr);
    console.log('Other employees count:', otherEmpStrings.length);
    console.log('=== GET WEEKLY HOLIDAY CALENDAR END ===');
    
    return 'SUCCESS:' + currentEmpStr + '|' + otherEmpStr;
    
  } catch (error) {
    console.error('Error getting weekly holiday calendar:', error);
    return 'ERROR:' + error.toString();
  }
}

/**
 * Format employee holiday data as string with dates and paused periods
 * Format: id~name~fri~sat~sun~mon~tue~wed~thu~fri_paused~sat_paused~sun_paused~mon_paused~tue_paused~wed_paused~thu_paused
 * Day format: status:type:startDate:endDate:reactivationDate OR "none"
 */
function formatEmployeeHolidayString(empId, empName, holidays) {
  const days = ['friday', 'saturday', 'sunday', 'monday', 'tuesday', 'wednesday', 'thursday'];
  
  const dayStrings = [];
  const pausedStrings = [];
  
  // Format regular day data
  for (let day of days) {
    const holiday = holidays[day];
    if (!holiday || holiday.status === 'none') {
      dayStrings.push('none');
    } else {
      const startDate = holiday.startDate || '';
      const endDate = holiday.endDate || '';
      const reactivationDate = holiday.reactivationDate || '';
      dayStrings.push(String(holiday.status) + ':' + String(holiday.type) + ':' + startDate + ':' + endDate + ':' + reactivationDate);
    }
  }
  
  // Format paused day data
  for (let day of days) {
    const pausedKey = day + '_paused';
    const pausedHoliday = holidays[pausedKey];
    if (!pausedHoliday || pausedHoliday.status === 'none') {
      pausedStrings.push('none');
    } else {
      const startDate = pausedHoliday.startDate || '';
      const endDate = pausedHoliday.endDate || '';
      const reactivationDate = pausedHoliday.reactivationDate || '';
      pausedStrings.push(String(pausedHoliday.status) + ':' + String(pausedHoliday.type) + ':' + startDate + ':' + endDate + ':' + reactivationDate);
    }
  }
  
  // Combine: id~name~day1~day2~...~day7~day1_paused~day2_paused~...~day7_paused
  return String(empId) + '~' + String(empName) + '~' + dayStrings.join('~') + '~' + pausedStrings.join('~');
}
/**
 * Get weekly holiday status for a specific employee with priority logic
 * Returns object with status for each day of week
 * UPDATED: Temp periods override standard periods
 */
function getEmployeeWeeklyHolidayStatus(employeeId, holidaysData, currentDate) {
  try {
    const days = ['friday', 'saturday', 'sunday', 'monday', 'tuesday', 'wednesday', 'thursday'];
    const dayColumns = [7, 8, 9, 10, 11, 12, 13];
    
    const result = {};
    for (let day of days) {
      result[day] = { status: 'none', type: '', startDate: '', endDate: '', reactivationDate: '' };
    }
    
    employeeId = String(employeeId);
    
    const employeePeriods = [];
    
    for (let i = 1; i < holidaysData.length; i++) {
      const row = holidaysData[i];
      const empId = String(row[1] || '');
      
      if (empId !== employeeId) continue;
      
      const periodType = String(row[6] || '').toLowerCase();
      const startDate = row[3] ? new Date(row[3]) : null;
      const endDate = row[4] ? new Date(row[4]) : null;
      
      if (!startDate) continue;
      
      let status = 'none';
      
      if (periodType === 'temp') {
        if (!endDate) continue;
        
        if (currentDate >= startDate && currentDate <= endDate) {
          status = 'active';
        } else if (currentDate < startDate) {
          status = 'inactive';
        } else {
          continue;
        }
      } else if (periodType === 'standard') {
        if (currentDate >= startDate) {
          if (!endDate || currentDate <= endDate) {
            status = 'active';
          } else {
            continue;
          }
        } else {
          status = 'inactive';
        }
      }
      
      const holidayDays = {};
      for (let j = 0; j < dayColumns.length; j++) {
        const col = dayColumns[j];
        const dayName = days[j];
        const isHoliday = row[col] === 'True' || row[col] === true || String(row[col]).toLowerCase() === 'true';
        if (isHoliday) {
          holidayDays[dayName] = true;
        }
      }
      
      employeePeriods.push({
        type: periodType,
        status: status,
        days: holidayDays,
        startDate: startDate,
        endDate: endDate
      });
    }
    
    // CRITICAL: Apply priority with temp override logic
    const activeTempPeriods = employeePeriods.filter(function(p) { return p.status === 'active' && p.type === 'temp'; });
    const activeStandardPeriods = employeePeriods.filter(function(p) { return p.status === 'active' && p.type === 'standard'; });
    const inactiveTempPeriods = employeePeriods.filter(function(p) { return p.status === 'inactive' && p.type === 'temp'; });
    const inactiveStandardPeriods = employeePeriods.filter(function(p) { return p.status === 'inactive' && p.type === 'standard'; });
    
    // Process each day
    for (let day of days) {
      // Check if there's an active temp period for this day
      const activeTempDay = activeTempPeriods.find(function(p) { return p.days[day]; });
      
      if (activeTempDay) {
        // Active temp exists - show it as active
        const startDateStr = formatDateForCalendar(activeTempDay.startDate);
        const endDateStr = formatDateForCalendar(activeTempDay.endDate);
        
        result[day] = {
          status: 'active',
          type: 'temp',
          startDate: startDateStr,
          endDate: endDateStr,
          reactivationDate: ''
        };
        
        // Check if there's an active standard that's being overridden
        const activeStandardDay = activeStandardPeriods.find(function(p) { return p.days[day]; });
        
        if (activeStandardDay) {
          // Standard exists but is paused by temp - store as secondary
          const reactivationDate = new Date(activeTempDay.endDate);
          reactivationDate.setDate(reactivationDate.getDate() + 1);
          const reactivationDateStr = formatDateForCalendar(reactivationDate);
          
          result[day + '_paused'] = {
            status: 'inactive',
            type: 'standard',
            startDate: formatDateForCalendar(activeStandardDay.startDate),
            endDate: formatDateForCalendar(activeStandardDay.endDate),
            reactivationDate: reactivationDateStr
          };
        }
        
        continue;
      }
      
      // No active temp - check active standard
      const activeStandardDay = activeStandardPeriods.find(function(p) { return p.days[day]; });
      if (activeStandardDay) {
        result[day] = {
          status: 'active',
          type: 'standard',
          startDate: formatDateForCalendar(activeStandardDay.startDate),
          endDate: formatDateForCalendar(activeStandardDay.endDate),
          reactivationDate: ''
        };
        continue;
      }
      
      // Check inactive temp
      const inactiveTempDay = inactiveTempPeriods.find(function(p) { return p.days[day]; });
      if (inactiveTempDay) {
        result[day] = {
          status: 'inactive',
          type: 'temp',
          startDate: formatDateForCalendar(inactiveTempDay.startDate),
          endDate: formatDateForCalendar(inactiveTempDay.endDate),
          reactivationDate: ''
        };
        continue;
      }
      
      // Check inactive standard
      const inactiveStandardDay = inactiveStandardPeriods.find(function(p) { return p.days[day]; });
      if (inactiveStandardDay) {
        result[day] = {
          status: 'inactive',
          type: 'standard',
          startDate: formatDateForCalendar(inactiveStandardDay.startDate),
          endDate: formatDateForCalendar(inactiveStandardDay.endDate),
          reactivationDate: ''
        };
      }
    }
    
    return result;
    
  } catch (error) {
    console.error('Error getting employee weekly holiday status:', error);
    const days = ['friday', 'saturday', 'sunday', 'monday', 'tuesday', 'wednesday', 'thursday'];
    const result = {};
    for (let day of days) {
      result[day] = { status: 'none', type: '', startDate: '', endDate: '', reactivationDate: '' };
    }
    return result;
  }
}

/**
 * Format date for calendar display
 * Format: "03 Sep 2025"
 */
function formatDateForCalendar(date) {
  if (!date) return '';
  
  try {
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const month = months[d.getMonth()];
    const year = d.getFullYear();
    
    return day + ' ' + month + ' ' + year;
  } catch (error) {
    return '';
  }
}

/* ============================================================================
   EMPLOYEE PORTAL - OFFICIAL HOLIDAYS TAB
   Backend function to retrieve official holidays with employee-specific data
   ============================================================================ */

/**
 * Get official holidays for employee portal with employee-specific assignments
 * Returns holidays grouped by name with detailed daily breakdown
 * @param {string} employeeEmail - Current employee's email
 * @return {string} SUCCESS:JSON or ERROR:message
 */
function getEmployeeOfficialHolidays(employeeEmail) {
  try {
    console.log('📅 Getting official holidays for employee:', employeeEmail);
    
    // Get current employee data
    const employeeResult = getCurrentEmployeeData(employeeEmail);
    if (!employeeResult.success) {
      return "ERROR:" + employeeResult.error;
    }
    
    const currentEmployeeId = employeeResult.data.id.toString().trim();
    const currentEmployeeName = employeeResult.data.name;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const holidaysSheet = ss.getSheetByName('Official holidays');
    
    // If sheet doesn't exist or is empty, return empty array
    if (!holidaysSheet) {
      console.log('📋 Official holidays sheet not found');
      return "SUCCESS:" + JSON.stringify({ 
        holidays: [], 
        currentEmployeeId: currentEmployeeId,
        currentEmployeeName: currentEmployeeName 
      });
    }
    
    const lastRow = holidaysSheet.getLastRow();
    if (lastRow <= 2) {
      console.log('📋 No holidays found (empty sheet)');
      return "SUCCESS:" + JSON.stringify({ 
        holidays: [], 
        currentEmployeeId: currentEmployeeId,
        currentEmployeeName: currentEmployeeName 
      });
    }
    
    const data = holidaysSheet.getDataRange().getValues();
    const headerRow = data[0]; // Row 1: Employee_{ID} headers
    const idRow = data[1];     // Row 2: Employee IDs
    
    // Get all active employees for reference
    const employeesResult = getActiveEmployeesForAssignment();
    const employeesMap = {};
    
    if (!employeesResult.startsWith('ERROR:') && !employeesResult.startsWith('SUCCESS:EMPTY')) {
      const employeesString = employeesResult.replace('SUCCESS:', '');
      const employeeArray = employeesString.split(';');
      
      employeeArray.forEach(empStr => {
        const parts = empStr.split('|');
        if (parts.length >= 4) {
          employeesMap[parts[2]] = {
            id: parts[2],
            name: parts[0],
            email: parts[1],
            weeklyHoliday: parts[3]
          };
        }
      });
    }
    
    // Group holidays by name
    const holidaysMap = {};
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Process each data row (starting from row 4, index 3)
    for (let i = 3; i < data.length; i++) {
      const holidayName = data[i][0];
      const dateValue = data[i][1];
      
      if (!holidayName || !dateValue) continue;
      
      const holidayDate = new Date(dateValue);
      const dateStr = holidayDate.getFullYear() + '-' + 
                     String(holidayDate.getMonth() + 1).padStart(2, '0') + '-' + 
                     String(holidayDate.getDate()).padStart(2, '0');
      
      // Initialize holiday group if not exists
      if (!holidaysMap[holidayName]) {
        holidaysMap[holidayName] = {
          name: holidayName,
          dates: [],
          startDate: dateStr,
          endDate: dateStr,
          isPast: false,
          employeeAssignments: {} // Will store assignments per employee per date
        };
      }
      
      // Add date to holiday
      if (!holidaysMap[holidayName].dates.includes(dateStr)) {
        holidaysMap[holidayName].dates.push(dateStr);
      }
      
      // Update date range
      if (dateStr < holidaysMap[holidayName].startDate) {
        holidaysMap[holidayName].startDate = dateStr;
      }
      if (dateStr > holidaysMap[holidayName].endDate) {
        holidaysMap[holidayName].endDate = dateStr;
      }
      
      // Check employee assignments for this date (columns C onwards)
      for (let col = 2; col < headerRow.length; col++) {
        const employeeIdInSheet = idRow[col] ? idRow[col].toString().trim() : '';
        
        if (!employeeIdInSheet || !employeesMap[employeeIdInSheet]) continue;
        
        const assignmentValue = data[i][col];
        const employee = employeesMap[employeeIdInSheet];
        
        // Initialize employee assignments if not exists
        if (!holidaysMap[holidayName].employeeAssignments[employeeIdInSheet]) {
          holidaysMap[holidayName].employeeAssignments[employeeIdInSheet] = {
            employeeId: employeeIdInSheet,
            employeeName: employee.name,
            weeklyHoliday: employee.weeklyHoliday,
            dailyAssignments: {} // dateStr: {status, notified}
          };
        }
        
        // Parse assignment value
        let status = 'Off';
        let notified = false;
        
        if (assignmentValue) {
          const valueStr = assignmentValue.toString().trim();
          if (valueStr.includes('Work')) {
            status = 'Work';
            notified = valueStr.includes('[Notified]');
          }
        }
        
        holidaysMap[holidayName].employeeAssignments[employeeIdInSheet].dailyAssignments[dateStr] = {
          status: status,
          notified: notified
        };
      }
    }
    
    // Convert to array and determine past/upcoming status
    const holidays = Object.values(holidaysMap).map(holiday => {
      // Sort dates
      holiday.dates.sort();
      
      // Determine if holiday is past
      const endDate = new Date(holiday.endDate);
      endDate.setHours(23, 59, 59, 999);
      holiday.isPast = endDate < today;
      
      // Convert employeeAssignments to array and sort
      const employeeAssignmentsArray = Object.values(holiday.employeeAssignments);
      
      // Sort: Current employee first, then by ID
      employeeAssignmentsArray.sort((a, b) => {
        if (a.employeeId === currentEmployeeId) return -1;
        if (b.employeeId === currentEmployeeId) return 1;
        return a.employeeId.localeCompare(b.employeeId);
      });
      
      holiday.employeeAssignments = employeeAssignmentsArray;
      
      // Count working vs off for current employee
      let currentEmployeeWorkingCount = 0;
      let currentEmployeeOffCount = 0;
      
      const currentEmpAssignment = holiday.employeeAssignments.find(
        emp => emp.employeeId === currentEmployeeId
      );
      
      if (currentEmpAssignment) {
        holiday.dates.forEach(dateStr => {
          const assignment = currentEmpAssignment.dailyAssignments[dateStr];
          if (assignment && assignment.status === 'Work') {
            currentEmployeeWorkingCount++;
          } else {
            currentEmployeeOffCount++;
          }
        });
      } else {
        currentEmployeeOffCount = holiday.dates.length;
      }
      
      holiday.currentEmployeeWorkingCount = currentEmployeeWorkingCount;
      holiday.currentEmployeeOffCount = currentEmployeeOffCount;
      
      // Count total working vs off (all employees)
      let totalWorkingCount = 0;
      let totalOffCount = 0;
      
      holiday.employeeAssignments.forEach(empAssignment => {
        holiday.dates.forEach(dateStr => {
          const assignment = empAssignment.dailyAssignments[dateStr];
          if (assignment && assignment.status === 'Work') {
            totalWorkingCount++;
          } else {
            totalOffCount++;
          }
        });
      });
      
      holiday.totalWorkingCount = totalWorkingCount;
      holiday.totalOffCount = totalOffCount;
      
      return holiday;
    });
    
    // Sort holidays by start date
    holidays.sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
    
    console.log(`✅ Processed ${holidays.length} holiday groups for employee portal`);
    
    return "SUCCESS:" + JSON.stringify({
      holidays: holidays,
      currentEmployeeId: currentEmployeeId,
      currentEmployeeName: currentEmployeeName
    });
    
  } catch (error) {
    console.error('💥 Error getting employee official holidays:', error);
    return "ERROR:" + error.toString();
  }
}
