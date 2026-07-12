
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Admin Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Authenticate admin user
function authenticateUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    var data = usersSheet.getDataRange().getValues();
    
    // 🔑 Extract REC_ID from password prefix
    var recIdFromPassword = '';
    var passwordParts = password.split('-');
    if (passwordParts.length > 1 && /^\d+$/.test(passwordParts[0])) {
      recIdFromPassword = passwordParts[0];
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username && 
          data[i][1] === password && 
          (data[i][2] === 'admin' || data[i][2] === 'superadmin')) {
        
        var userRecId = data[i][15] ? data[i][15].toString() : '';
        
        // ✅ Verify REC_ID matches password prefix
        if (recIdFromPassword && userRecId !== recIdFromPassword) {
          continue; // Skip wrong school
        }
        
        return {
          success: true,
          user: {
            Username: data[i][0],
            UserType: data[i][2],
            Teacher_Name: data[i][3],
            student_delete: data[i][8] || 'No',
            canQueryAttendance: data[i][9] || 'No',
            canDownload: data[i][10] || 'No',
            canAmendSIRN: data[i][16] || 'No',
            REC_ID: userRecId,  // ← Correct REC_ID
            allowedTabs: data[i][7] || 'users,user_permissions,class_sections,previous_date_permissions,teacher_assignments,student_data,student_attendance,qr_management,reserve_codes'
          }
        };
      }
    }
    return { success: false, message: 'Invalid credentials or insufficient permissions' };
  } catch (e) {
    Logger.log('Authentication error: ' + e.message);
    return { success: false, message: 'Authentication error: ' + e.message };
  }
}

// Get data from a specific sheet
// Get data from a specific sheet
function getSheetData(sheetName, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    // Find REC_ID column index
    const recIdCol = headers.indexOf('REC_ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        // If REC_ID column exists and we have a REC_ID filter
        if (recIdCol !== -1 && recId) {
          const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
          // ONLY show exact matches - hide blank and different schools
          if (rowRecId !== recId) {
            continue; // Skip this row - different school or blank
          }
        }
        
        const row = {};
        for (let j = 0; j < headers.length; j++) {
          row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j].toString() : '';
        }
        row._rowNumber = i + 1;
        result.push(row);
      }
    }
    return result;
  } catch (e) {
    Logger.log('Failed to fetch data from ' + sheetName + ': ' + e.message);
    throw new Error('Failed to fetch data from ' + sheetName + ': ' + e.message);
  }
}

// Get Class_Section data (column A)
function getClassSectionData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Class_Sections');
    if (!sheet) throw new Error('Class_Sections sheet not found');
    
    const data = sheet.getDataRange().getValues();
    
    return data.slice(1) // Skip header row
      .filter(row => row[0]) // Filter empty rows
      .map((row, index) => ({
        Class_Section: row[0] ? row[0].toString() : '',
        _rowNumber: index + 2
      }));
  } catch (e) {
    Logger.log('Failed to fetch class section data: ' + e.message);
    throw new Error('Failed to fetch class section data: ' + e.message);
  }
}

// Get class and section mappings from Class_Sections sheet
function getClassMappingData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Class_Sections');
    if (!sheet) throw new Error('Class_Sections sheet not found');
    
    var data = sheet.getDataRange().getValues();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] || data[i][3]) {
        result.push({
          Class: data[i][2] !== null && data[i][2] !== undefined ? data[i][2].toString() : '',
          Sections: data[i][3] !== null && data[i][3] !== undefined ? data[i][3].toString() : '',
          _rowNumber: i + 1
        });
      }
    }
    return result;
  } catch (e) {
    Logger.log('Failed to fetch class mapping data: ' + e.message);
    throw new Error('Failed to fetch class mapping data: ' + e.message);
  }
}

// Add a new row to a sheet
// Add a new row to a sheet
function addSheetRow(sheetName, rowData, currentUsername) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');
    
    // Handle special cases first
    if (sheetName === 'Student_Data') {
      // Pass REC_ID from rowData if available
      const recId = rowData.REC_ID || '';
      return addStudent(rowData, currentUsername, recId);
    }
    
    if (sheetName === 'Teacher_Assignments') {
      const existingData = sheet.getDataRange().getValues();
      const exists = existingData.some((row, i) => i > 0 && row[0] === rowData.Username);
      if (exists) {
        return { 
          success: false, 
          message: `Username ${rowData.Username} already has assigned classes. Please edit the existing record.` 
        };
      }
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Get REC_ID from the current user
    let recId = '';
    if (currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        const userRow = userData.find(row => row[0] === currentUsername);
        if (userRow) {
          recId = userRow[15] ? userRow[15].toString() : ''; // Column P is index 15
        }
      }
    }

    // 🔧 FIX: For Users sheet, set REC_ID in rowData BEFORE building the row
    if (sheetName === 'Users' && recId) {
      rowData.REC_ID = recId;
    }

    // Build new row
    const newRow = headers.map(header => {
      const value = rowData[header];
      
      // Auto-fill REC_ID if not provided
      if (header === 'REC_ID' && !value && recId) {
        return recId;
      }
      
      // Handle special formatting cases
      if (sheetName === 'User_Permissions' && ['canMarkManual', 'canMarkBarcode', 'canSeeSummary'].includes(header)) {
        return value === 'true' || value === true;
      }
      if (sheetName === 'Users' && ['StartTime', 'EndTime'].includes(header)) {
        return value ? value.toString() : '';
      }
      if (sheetName === 'Users' && ['canQueryAttendance', 'canDownload'].includes(header)) {
        return value === 'Yes' ? 'Yes' : 'No';
      }
      
      return value !== undefined && value !== null ? value : '';
    });
    
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    // Handle time format for Users sheet
    if (sheetName === 'Users') {
      const startTimeIndex = headers.indexOf('StartTime');
      const endTimeIndex = headers.indexOf('EndTime');
      const lastRow = sheet.getLastRow();
      
      if (startTimeIndex > -1) {
        sheet.getRange(lastRow, startTimeIndex + 1).setNumberFormat('@');
      }
      if (endTimeIndex > -1) {
        sheet.getRange(lastRow, endTimeIndex + 1).setNumberFormat('@');
      }
    }
    
    return { 
      success: true, 
      message: 'Record added successfully', 
      data: { ...rowData, _rowNumber: sheet.getLastRow() } 
    };
  } catch (e) {
    Logger.log('Error in addSheetRow: ' + e.message);
    return { success: false, message: 'Failed to add row to ' + sheetName + ': ' + e.message };
  }
}


function updateSheetRow(sheetName, rowIndex, rowData, currentUsername) {
  try {
    // ✅ Get username from multiple sources
    let username = currentUsername;
    
    if (!username) {
      username = rowData.Username || rowData.currentUsername || '';
    }
    
    if (!username) {
      const sessionUser = Session.getActiveUser().getEmail();
      if (sessionUser) {
        username = sessionUser.split('@')[0];
      }
    }
    
    if (!username) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (let i = 1; i < userData.length; i++) {
          if (userData[i][15] && userData[i][15].toString() === rowData.REC_ID) {
            username = userData[i][0];
            break;
          }
        }
      }
    }
    
    if (!username) {
      return { 
        success: false, 
        message: 'Unable to identify user. Please re-login.' 
      };
    }
    
    Logger.log(`updateSheetRow called for sheet: ${sheetName}, rowIndex: ${rowIndex}, username: ${username}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No active spreadsheet found');
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`${sheetName} sheet not found`);
    
    // Block superadmin modifications
    if (sheetName === 'Users') {
      const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (rowData[0] === 'superadmin') {
        return { success: false, message: 'Superadmin account cannot be modified' };
      }
    }
    
    if (!Number.isInteger(rowIndex) || rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      Logger.log(`Invalid row index: ${rowIndex}. Last row: ${sheet.getLastRow()}`);
      return { success: false, message: `Invalid row index: ${rowIndex}. Must be between 2 and ${sheet.getLastRow()}` };
    }
    
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    const rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (protections.length > 0 || rangeProtections.length > 0) {
      const userEmail = Session.getActiveUser().getEmail();
      let canEdit = true;
      protections.forEach(protection => {
        if (!protection.getEditors().includes(userEmail)) {
          canEdit = false;
        }
      });
      rangeProtections.forEach(protection => {
        const protectedRange = protection.getRange();
        if (protectedRange.getRow() <= rowIndex && rowIndex <= protectedRange.getLastRow()) {
          if (!protection.getEditors().includes(userEmail)) {
            canEdit = false;
          }
        }
      });
      if (!canEdit) {
        Logger.log(`User ${userEmail} lacks permission to edit protected ranges in ${sheetName}`);
        return { success: false, message: 'Permission denied: Sheet or range is protected' };
      }
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers || headers.length === 0) {
      throw new Error('No headers found in the sheet');
    }
    
    // ✅ STUDENT DATA SECTION
    if (sheetName === 'Student_Data') {
      const rowRange = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn());
      const rowValues = rowRange.getValues()[0];
      const originalStdId = rowValues[0];
      
      const currentUser = username;
      const currentDate = new Date();
      const formattedDate = Utilities.formatDate(currentDate, 'GMT+0500', 'dd-MMM-yyyy');
      
      // ✅ Check if SIRN is actually changing
      const newStdId = rowData.Std_ID ? rowData.Std_ID.toString() : '';
      const oldStdId = originalStdId ? originalStdId.toString() : '';
      
      if (newStdId && newStdId !== oldStdId) {
        // 🔒 CHECK PERMISSION TO AMEND SIRN
        const usersSheet = ss.getSheetByName('Users');
        if (usersSheet) {
          const userHeaders = usersSheet.getRange(1, 1, 1, usersSheet.getLastColumn()).getValues()[0];
          const canAmendSIRNCol = userHeaders.indexOf('canAmendSIRN');
          
          if (canAmendSIRNCol === -1) {
            Logger.log('WARNING: canAmendSIRN column not found in Users sheet');
            return { 
              success: false, 
              message: 'System configuration error: canAmendSIRN permission not configured. Please contact administrator.' 
            };
          }
          
          const userData = usersSheet.getDataRange().getValues();
          const userRow = userData.find(row => row[0] === currentUser);
          
          if (!userRow) {
            Logger.log(`User ${currentUser} not found in Users sheet`);
            return { 
              success: false, 
              message: 'User not found. Please re-login.' 
            };
          }
          
          // ✅ IMPORTANT: Blank = "No" (default)
          const canAmendSIRN = userRow[canAmendSIRNCol] || 'No';
          
          if (canAmendSIRN !== 'Yes') {
            Logger.log(`User ${currentUser} does not have permission to amend SIRN. canAmendSIRN: "${canAmendSIRN}"`);
            return { 
              success: false, 
              message: 'You do not have permission to amend SIRN (Student ID)' 
            };
          }
          
          Logger.log(`User ${currentUser} has permission to amend SIRN: ${canAmendSIRN}`);
        } else {
          Logger.log('ERROR: Users sheet not found');
          return { 
            success: false, 
            message: 'System error: Users sheet not found' 
          };
        }
        
        // ✅ Validate the new SIRN
        if (!/^\d+$/.test(newStdId)) {
          Logger.log(`Std_ID validation failed: ${newStdId} is not numeric`);
          return { success: false, message: 'Std_ID must be numeric' };
        }
        
        // ✅ Check duplicate ONLY within the same REC_ID
        const stdIdCol = headers.indexOf('Std_ID');
        const recIdCol = headers.indexOf('REC_ID');
        const currentRowRecId = rowValues[recIdCol] ? rowValues[recIdCol].toString() : '';
        
        if (stdIdCol !== -1 && recIdCol !== -1) {
          const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
          
          for (let i = 0; i < existingData.length; i++) {
            const rowNum = i + 2;
            if (rowNum !== rowIndex) {
              const existingStdId = existingData[i][stdIdCol] ? existingData[i][stdIdCol].toString() : '';
              const existingRecId = existingData[i][recIdCol] ? existingData[i][recIdCol].toString() : '';
              
              if (existingStdId === newStdId && existingRecId === currentRowRecId) {
                Logger.log(`Duplicate Std_ID found: ${newStdId} at row ${rowNum}`);
                return { success: false, message: 'SIRN already exists in your school' };
              }
            }
          }
        }
      } else {
        Logger.log(`SIRN not changed. Original: "${oldStdId}", New: "${newStdId}"`);
      }
      
      // ✅ Format Creation Date
      let creationDate = rowValues[9] || '';
      if (creationDate instanceof Date) {
        creationDate = Utilities.formatDate(creationDate, 'GMT+0500', 'dd-MMM-yyyy');
      }
      
      // ✅ Handle Inactive Date logic
      let inactiveDate = rowValues[8] || '';
      if (rowData.Status === 'Inactive' && rowValues[5] !== 'Inactive') {
        inactiveDate = formattedDate;
      } else if (rowData.Status === 'Active' && rowValues[5] === 'Inactive') {
        inactiveDate = '';
      }
      
      // ✅ BUILD THE UPDATED ROW - Dynamically match sheet columns
      const updatedRow = [];
      
      for (let i = 0; i < headers.length; i++) {
        const header = headers[i];
        
        switch(header) {
          case 'Std_ID':
            updatedRow.push(rowData.Std_ID || originalStdId);
            break;
          case 'Barcode_ID':
            updatedRow.push(rowData.Barcode_ID || rowValues[1] || '');
            break;
          case 'Student_Name':
            updatedRow.push(rowData.Student_Name || rowValues[2] || '');
            break;
          case 'Student_Class':
            updatedRow.push(rowData.Student_Class || rowValues[3] || '');
            break;
          case 'Student_Section':
            updatedRow.push(rowData.Student_Section || rowValues[4] || '');
            break;
          case 'Status':
            updatedRow.push(rowData.Status || rowValues[5] || 'Active');
            break;
          case 'Gender':
            updatedRow.push(rowData.Gender || rowValues[6] || '');
            break;
          case 'Date_of_Joining':
            updatedRow.push(rowData.Date_of_Joining || rowValues[7] || '');
            break;
          case 'Inactive_Date':
            updatedRow.push(inactiveDate);
            break;
          case 'Creation_Date':
            updatedRow.push(creationDate);
            break;
          case 'Created_By':
            updatedRow.push(rowValues[10] || '');
            break;
          case 'Last_Modified_Date':
            updatedRow.push(formattedDate);
            break;
          case 'Last_Modified_By':
            updatedRow.push(currentUser);
            break;
          case 'REC_ID':
            updatedRow.push(rowValues[13] || '');
            break;
          case 'QR_Printed':
            updatedRow.push(rowValues[14] || '');
            break;
          case 'Unique_ID':
            updatedRow.push(rowData.Unique_ID || rowValues[15] || '');
            break;
          default:
            // For any other columns, preserve the original value
            updatedRow.push(rowValues[i] !== undefined && rowValues[i] !== null ? rowValues[i] : '');
            break;
        }
      }
      
      // ✅ Log the column count for debugging
      Logger.log(`Updated row has ${updatedRow.length} columns, sheet has ${headers.length} columns`);
      
      // ✅ Update the row
      let attempts = 3;
      while (attempts > 0) {
        try {
          rowRange.setValues([updatedRow]);
          SpreadsheetApp.flush();
          Logger.log(`Row ${rowIndex} updated successfully for Std_ID: ${updatedRow[0]}`);
          break;
        } catch (e) {
          Logger.log(`Attempt ${4 - attempts} failed: ${e.message}`);
          attempts--;
          if (attempts === 0) {
            return { success: false, message: `Failed to update row after retries: ${e.message}` };
          }
          Utilities.sleep(1000);
        }
      }
      
      // ✅ Return success with all data
return {
  success: true,
  message: 'Student updated successfully: ' + (rowData.Student_Name || rowValues[2] || 'Unknown') + ' (SIRN: ' + (rowData.Std_ID || rowValues[0] || 'Unknown') + ')',
  data: {
    Std_ID: updatedRow[0],
    Barcode_ID: updatedRow[1],
    Student_Name: updatedRow[2],
    Student_Class: updatedRow[3],
    Student_Section: updatedRow[4],
    Status: updatedRow[5],
    Gender: updatedRow[6],
    Date_of_Joining: updatedRow[7],
    Inactive_Date: updatedRow[8],
    Creation_Date: updatedRow[9],
    Created_By: updatedRow[10],
    Last_Modified_Date: updatedRow[11],
    Last_Modified_By: updatedRow[12],
    REC_ID: updatedRow[13] || '',
    QR_Printed: updatedRow[14] || '',
    Unique_ID: updatedRow[15] || '',
    _rowNumber: rowIndex
  }
};
    }
    
    // ✅ OTHER SHEETS (Users, User_Permissions, etc.)
    const newRow = headers.map(function(header) {
      const value = rowData[header];
      if (sheetName === 'User_Permissions' && ['canMarkManual', 'canMarkBarcode', 'canSeeSummary'].includes(header)) {
        return value === 'true' || value === true;
      }
      if (sheetName === 'Users' && ['StartTime', 'EndTime'].includes(header)) {
        return value ? value.toString() : '';
      }
      if (sheetName === 'Users' && ['canQueryAttendance', 'canDownload'].includes(header)) {
        return value === 'Yes' ? 'Yes' : 'No';
      }
      return value !== undefined && value !== null ? value : '';
    });
    
    let attempts = 3;
    while (attempts > 0) {
      try {
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
        SpreadsheetApp.flush();
        Logger.log(`Row ${rowIndex} updated successfully in ${sheetName}`);
        break;
      } catch (e) {
        Logger.log(`Attempt ${4 - attempts} failed: ${e.message}`);
        attempts--;
        if (attempts === 0) {
          return { success: false, message: `Failed to update row after retries: ${e.message}` };
        }
        Utilities.sleep(1000);
      }
    }
    
    if (sheetName === 'Users') {
      const startTimeIndex = headers.indexOf('StartTime') + 1;
      const endTimeIndex = headers.indexOf('EndTime') + 1;
      if (startTimeIndex > 0) {
        sheet.getRange(rowIndex, startTimeIndex).setNumberFormat('@');
      }
      if (endTimeIndex > 0) {
        sheet.getRange(rowIndex, endTimeIndex).setNumberFormat('@');
      }
    }
    
    return {
      success: true,
      message: 'Record updated successfully',
      data: { ...rowData, _rowNumber: rowIndex }
    };
    
  } catch (e) {
    Logger.log(`Error in updateSheetRow for ${sheetName}: ${e.message}`);
    return {
      success: false,
      message: `Failed to update record: ${e.message}`
    };
  }
}

// Delete a row from a sheet
function deleteSheetRow(sheetName, rowNumber, username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + ' sheet not found');

    // Block superadmin deletion
if (sheetName === 'Users') {
  const rowData = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (rowData[0] === 'superadmin') {
    return { success: false, message: 'Superadmin account cannot be deleted' };
  }
}
    
    if (sheetName === 'Student_Data') {
      var usersSheet = ss.getSheetByName('Users');
      if (!usersSheet) throw new Error('Users sheet not found');
      
      var userData = usersSheet.getDataRange().getValues();
      var userRow = userData.find(row => row[0] === username);
      if (!userRow || userRow[8] !== 'Yes') {
        return { success: false, message: 'You do not have permission to delete student records' };
      }
    }
    
    sheet.deleteRow(rowNumber);
    SpreadsheetApp.flush();
    return { success: true, message: 'Record deleted successfully' };
  } catch (e) {
    Logger.log('Error in deleteSheetRow: ' + e.message);
    return { success: false, message: 'Failed to delete record: ' + e.message };
  }
}

// Get all class sections for teacher assignments
function getClassSections() {
  try {
    var data = getClassSectionData();
    return data.map(row => row.Class_Section);
  } catch (e) {
    Logger.log('Failed to fetch class sections: ' + e.message);
    throw new Error('Failed to fetch class sections: ' + e.message);
  }
}

// Add a student with barcode generation and audit fields
// Add a student with barcode generation and audit fields
function addStudent(studentData, currentUsername, recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    // ✅ VALIDATE Std_ID
    if (!studentData.Std_ID || !/^\d+$/.test(studentData.Std_ID)) {
      return { success: false, message: 'Std_ID must be numeric and required' };
    }
    
    // ✅ VALIDATE Unique_ID
    if (!studentData.Unique_ID || !/^\d{6}$/.test(studentData.Unique_ID)) {
      return { success: false, message: 'Unique_ID must be exactly 6 digits' };
    }
    
    const lastRow = sheet.getLastRow();
    
    // Get headers to find column indices
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const stdIdCol = headers.indexOf('Std_ID') + 1; // Column A
    const recIdCol = headers.indexOf('REC_ID') + 1; // Column N
    
    // ✅ FIX: Check duplicate Std_ID ONLY within the same REC_ID (school)
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      
      // Check if any row has the same Std_ID AND same REC_ID
      let duplicateFound = false;
      let existingRow = 0;
      
      for (let i = 0; i < dataRange.length; i++) {
        const rowStdId = dataRange[i][stdIdCol - 1] ? dataRange[i][stdIdCol - 1].toString() : '';
        const rowRecId = dataRange[i][recIdCol - 1] ? dataRange[i][recIdCol - 1].toString() : '';
        
        // Same Std_ID AND same REC_ID => duplicate
        if (rowStdId === studentData.Std_ID.toString() && rowRecId === recId) {
          duplicateFound = true;
          existingRow = i + 2; // +2 because dataRange starts at row 2
          break;
        }
      }
      
      if (duplicateFound) {
        return { 
          success: false, 
          message: `SIRN ${studentData.Std_ID} already exists in your school (Row ${existingRow}). Please use a different SIRN.` 
        };
      }
    }
    
    // Check duplicate Unique_ID (global - across all schools)
    const uniqueIdCol = headers.indexOf('Unique_ID') + 1;
    if (uniqueIdCol > 0 && lastRow > 1) {
      const existingUniqueIds = sheet.getRange(2, uniqueIdCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingUniqueIds.length; i++) {
        if (existingUniqueIds[i][0] && existingUniqueIds[i][0].toString() === studentData.Unique_ID.toString()) {
          return { success: false, message: 'Unique_ID already exists' };
        }
      }
    }
    
    // Generate Barcode
    let nextBarcode = 1000001;
    if (lastRow > 1) {
      const barcodes = sheet.getRange('B2:B' + lastRow).getValues();
      const numbers = barcodes
        .filter(row => row[0] && typeof row[0] === 'string' && row[0].startsWith('BC'))
        .map(row => parseInt(row[0].replace('BC', '')))
        .filter(num => !isNaN(num));
      if (numbers.length > 0) {
        nextBarcode = Math.max(...numbers) + 1;
      }
    }
    
    studentData.Barcode_ID = 'BC' + nextBarcode;
    
    // Get current user and timestamp
    const currentUser = currentUsername;
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, 'GMT+0500', 'dd-MMM-yyyy');
    
    // ✅ USE PASSED REC_ID - Backend verification
    let finalRecId = recId || '';
    
    // 🔒 BACKEND VERIFICATION: Verify recId matches the user's REC_ID
    let isValidRecId = false;
    if (finalRecId && currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (var i = 0; i < userData.length; i++) {
          if (userData[i][0] === currentUsername) {
            var userRecId = userData[i][15] ? userData[i][15].toString() : '';
            if (userRecId === finalRecId) {
              isValidRecId = true;
              break;
            }
          }
        }
      }
    }
    
    // If no valid REC_ID found, try to look it up (fallback)
    if (!isValidRecId && currentUsername) {
      const usersSheet = ss.getSheetByName('Users');
      if (usersSheet) {
        const userData = usersSheet.getDataRange().getValues();
        for (var i = 0; i < userData.length; i++) {
          if (userData[i][0] === currentUsername) {
            finalRecId = userData[i][15] ? userData[i][15].toString() : '';
            if (finalRecId) {
              isValidRecId = true;
            }
            break;
          }
        }
      }
    }
    
    // If still no valid REC_ID, return error
    if (!isValidRecId || !finalRecId) {
      return { success: false, message: 'Unable to determine school (REC_ID). Please contact administrator.' };
    }
    
    // If we have a REC_ID and it's not in studentData, add it
    if (finalRecId && !studentData.REC_ID) {
      studentData.REC_ID = finalRecId;
    }
    
    // Ensure Unique_ID column exists
    if (uniqueIdCol === 0) {
      sheet.getRange(1, headers.length + 1).setValue('Unique_ID');
    }
    
    // NEW ROW with REC_ID
    const newRow = [
      studentData.Std_ID,                // A - 0
      studentData.Barcode_ID,            // B - 1
      studentData.Student_Name || '',    // C - 2
      studentData.Student_Class || '',   // D - 3
      studentData.Student_Section || '', // E - 4
      studentData.Status || 'Active',    // F - 5
      studentData.Gender || '',          // G - 6
      studentData.Date_of_Joining || '', // H - 7
      '',                                 // I - 8 Inactive Date
      formattedDate,                     // J - 9 Creation Date
      currentUser,                       // K - 10 Created By
      formattedDate,                     // L - 11 Last Modified Date
      currentUser,                       // M - 12 Last Modified By
      finalRecId,                        // N - 13 REC_ID (passed value)
      '',                                 // O - 14 QR_Printed
      studentData.Unique_ID              // P - 15 Unique_ID
    ];
    
    sheet.getRange(lastRow + 1, 1, 1, 16).setValues([newRow]);
    SpreadsheetApp.flush();
    
    return {
      success: true,
      message: 'Student added successfully: ' + studentData.Student_Name + ' (SIRN: ' + studentData.Std_ID + ', Barcode: ' + studentData.Barcode_ID + ')',
      data: { 
        ...studentData, 
        Creation_Date: formattedDate,
        Created_By: currentUser,
        Last_Modified_Date: formattedDate,
        Last_Modified_By: currentUser,
        _rowNumber: lastRow + 1 
      }
    };
  } catch (e) {
    Logger.log('Error in addStudent: ' + e.message);
    return { success: false, message: 'Error adding student: ' + e.message };
  }
}

// Get student data
function getStudentData(recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    // Find REC_ID column index
    const recIdCol = headers.indexOf('REC_ID');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        // Filter by REC_ID
        if (recIdCol !== -1 && recId) {
          const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
          if (rowRecId !== recId) {
            continue; // Skip - different school
          }
        }
        
        let row = {};
        for (let j = 0; j < headers.length; j++) {
          if (headers[j] === 'Creation_Date' && data[i][j]) {
            try {
              const dateObj = new Date(data[i][j]);
              row[headers[j]] = Utilities.formatDate(dateObj, 'GMT+0500', 'dd-MMM-yyyy');
            } catch (e) {
              row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j].toString() : '';
            }
          } else if (['Date_of_Joining', 'Last_Modified_Date', 'Inactive_Date'].includes(headers[j]) && data[i][j] instanceof Date) {
            row[headers[j]] = Utilities.formatDate(data[i][j], 'GMT+0500', 'dd-MMM-yyyy');
          } else {
            row[headers[j]] = data[i][j] !== null && data[i][j] !== undefined ? data[i][j].toString() : '';
          }
        }
        row._rowNumber = i + 1;
        result.push(row);
      }
    }
    return result;
  } catch (e) {
    Logger.log('Error in getStudentData: ' + e.message);
    throw new Error('Failed to fetch student data: ' + e.message);
  }
}

// Get gender options for dropdown
function getGenderOptions() {
  return ['Male', 'Female'];
}

// Get student attendance based on search criteria
function getStudentAttendance(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Permission check
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === formData.username);
    if (!userRow || userRow[9] !== 'Yes') {
      throw new Error('You do not have permission to query attendance');
    }
    
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // No data
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Pre-process search criteria
    const searchValue = formData.searchValue ? formData.searchValue.toString().toLowerCase() : '';
    const searchType = formData.searchType;
    const statusFilter = formData.statusFilter;
    
    // Date range setup
    let startDate, endDate;
    if (formData.startDate && formData.endDate) {
      startDate = new Date(formData.startDate);
      endDate = new Date(formData.endDate);
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);
    }
    
    // Cache column indices
    const columnIndices = {
      date: headers.indexOf('Date'),
      stdId: headers.indexOf('Std_ID'),
      barcodeId: headers.indexOf('Barcode_ID'),
      studentName: headers.indexOf('Student_Name'),
      status: headers.indexOf('Status')
    };
    
    return data
      .filter(row => {
        if (!row[columnIndices.date]) return false; // Skip empty rows
        
        // Date filtering
        if (startDate && endDate) {
          const attendanceDate = new Date(row[columnIndices.date]);
          if (isNaN(attendanceDate)) return false;
          attendanceDate.setHours(0, 0, 0, 0);
          if (attendanceDate < startDate || attendanceDate > endDate) return false;
        }
        
        // Search value filtering
        if (searchValue) {
          let matches = false;
          switch (searchType) {
            case 'Std_ID':
              matches = row[columnIndices.stdId].toString().toLowerCase() === searchValue;
              break;
            case 'Barcode_ID':
              matches = row[columnIndices.barcodeId].toString().toLowerCase() === searchValue;
              break;
            case 'Student_Name':
              matches = row[columnIndices.studentName].toString().toLowerCase().includes(searchValue);
              break;
          }
          if (!matches) return false;
        }
        
        // Status filtering
        if (statusFilter !== 'All' && row[columnIndices.status].toString() !== statusFilter) {
          return false;
        }
        
        return true;
      })
      .map((row, index) => {
        const resultRow = {};
        headers.forEach((header, j) => {
          resultRow[header] = header === 'Date' && row[j] instanceof Date 
            ? Utilities.formatDate(row[j], 'GMT+0500', 'yyyy-MMM-dd')
            : (row[j] !== null && row[j] !== undefined ? row[j].toString() : '');
        });
        resultRow._rowNumber = index + 2;
        return resultRow;
      });
  } catch (e) {
    Logger.log('Error in getStudentAttendance: ' + e.message);
    throw new Error('Failed to fetch attendance data: ' + e.message);
  }
}

// Download attendance data in CSV or PDF format
function downloadAttendance(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) throw new Error('Users sheet not found');
    
    const userData = usersSheet.getDataRange().getValues();
    const userRow = userData.find(row => row[0] === formData.username);
    if (!userRow || userRow[10] !== 'Yes') {
      throw new Error('You do not have permission to download attendance reports');
    }
    
    const sheet = ss.getSheetByName('Student_Attendance');
    if (!sheet) throw new Error('Student_Attendance sheet not found');
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const requiredColumns = ['Attendance_ID', 'Date', 'Std_ID', 'Barcode_ID', 'Student_Name', 'Student_Class', 'Student_Section', 'Status', 'Timestamp', 'Teacher_ID', 'Class_Section'];
    const columnIndices = requiredColumns.map(col => headers.indexOf(col) + 1).filter(idx => idx > 0);
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    let filteredData = [headers];
    
    // Pre-parse search criteria
    const searchValue = formData.searchValue ? formData.searchValue.toString().toLowerCase() : '';
    const searchType = formData.searchType;
    const statusFilter = formData.statusFilter;
    const dateIndex = headers.indexOf('Date');

    // Normalize startDate and endDate once
    let startDate, endDate;
    if (formData.startDate && formData.endDate) {
      startDate = new Date(formData.startDate);
      endDate = new Date(formData.endDate);
      startDate.setHours(0, 0, 0, 0); // Start of day
      endDate.setHours(23, 59, 59, 999); // End of day
    }

    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      let matches = true;

      // Date range filtering first
      if (startDate && endDate) {
        const attendanceDate = new Date(data[i][dateIndex]);
        if (isNaN(attendanceDate)) {
          Logger.log(`Invalid date at row ${i + 2}: ${data[i][dateIndex]}`);
          continue;
        }
        attendanceDate.setHours(0, 0, 0, 0); // Normalize to start of day
        if (attendanceDate < startDate || attendanceDate > endDate) {
          matches = false;
          continue;
        }
      }

      // Search value filtering
      if (searchValue) {
        if (searchType === 'Std_ID' && data[i][2].toString().toLowerCase() !== searchValue) {
          matches = false;
        } else if (searchType === 'Barcode_ID' && data[i][3].toString().toLowerCase() !== searchValue) {
          matches = false;
        } else if (searchType === 'Student_Name' && !data[i][4].toString().toLowerCase().includes(searchValue)) {
          matches = false;
        }
      }

      // Status filtering
      if (statusFilter !== 'All' && data[i][7].toString() !== statusFilter) {
        matches = false;
      }

      if (matches) {
        let row = data[i].slice();
        if (dateIndex !== -1 && row[dateIndex] instanceof Date) {
          row[dateIndex] = Utilities.formatDate(row[dateIndex], 'GMT+0500', 'yyyy-MMM-dd');
        }
        filteredData.push(row);
      }
    }
    
    if (formData.format === 'CSV') {
      let csv = headers.join(',') + '\n';
      for (let i = 1; i < filteredData.length; i++) {
        csv += filteredData[i].map(val => `"${val !== null && val !== undefined ? val.toString().replace(/"/g, '""') : ''}"`).join(',') + '\n';
      }
      return { success: true, data: csv };
    } else if (formData.format === 'PDF') {
      const doc = DocumentApp.create('Attendance_Report_' + new Date().toISOString());
      const body = doc.getBody();
      body.appendParagraph('Attendance Report').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      const table = body.appendTable(filteredData);
      table.setBorderWidth(1);
      table.getRow(0).setAttributes({ BOLD: true });
      doc.saveAndClose();
      const pdf = doc.getAs('application/pdf');
      const pdfBase64 = Utilities.base64Encode(pdf.getBytes());
      DriveApp.getFileById(doc.getId()).setTrashed(true);
      return { success: true, data: pdfBase64 };
    }
    return { success: false, message: 'Invalid format specified' };
  } catch (e) {
    Logger.log('Error in downloadAttendance: ' + e.message);
    return { success: false, message: 'Failed to download attendance: ' + e.message };
  }
}

// Generate QR code data URL for a student
function generateStudentQRCode(stdId, barcodeId, studentName) {
  try {
    // Create QR code data with student information
    const qrData = {
      type: 'student',
      stdId: stdId,
      barcodeId: barcodeId,
      name: studentName,
      timestamp: new Date().toISOString()
    };
    
    // Convert to JSON string for QR code
    const qrText = JSON.stringify(qrData);
    
    // Use Google Charts API to generate QR code
    const qrCodeUrl = `https://chart.googleapis.com/chart?chs=150x150&cht=qr&chl=${encodeURIComponent(qrText)}&choe=UTF-8`;
    
    return qrCodeUrl;
  } catch (e) {
    Logger.log('Error generating QR code: ' + e.message);
    throw e;
  }
}

// Get unprinted students for QR code generation
function getUnprintedStudents(recId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Check if QR_Printed column exists, if not create it
    if (!headers.includes('QR_Printed')) {
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1).setValue('QR_Printed');
      headers.push('QR_Printed');
    }
    
    const stdIdCol = headers.indexOf('Std_ID');
    const barcodeCol = headers.indexOf('Barcode_ID');
    const nameCol = headers.indexOf('Student_Name');
    const classCol = headers.indexOf('Student_Class');
    const sectionCol = headers.indexOf('Student_Section');
    const qrPrintedCol = headers.indexOf('QR_Printed');
    const recIdCol = headers.indexOf('REC_ID');  // ← ADD THIS
    
    const unprintedStudents = [];
    
    for (let i = 1; i < data.length; i++) {
      const qrPrinted = data[i][qrPrintedCol] || '';
      
      // 🔒 FILTER BY REC_ID
      if (recIdCol !== -1 && recId) {
        const rowRecId = data[i][recIdCol] ? data[i][recIdCol].toString() : '';
        if (rowRecId !== recId) {
          continue; // Skip - different school
        }
      }
      
      // Only include students whose QR hasn't been printed
      if (qrPrinted.toString().toLowerCase() !== 'yes') {
        unprintedStudents.push({
          Std_ID: data[i][stdIdCol],
          Barcode_ID: data[i][barcodeCol],
          Student_Name: data[i][nameCol],
          Student_Class: data[i][classCol],
          Student_Section: data[i][sectionCol],
          _rowNumber: i + 1
        });
      }
    }
    
    return unprintedStudents;
    
  } catch (e) {
    Logger.log('Error getting unprinted students: ' + e.message);
    throw e;
  }
}

// Mark students as printed
function markStudentsAsPrinted(studentIds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let qrPrintedCol = headers.indexOf('QR_Printed');
    
    // If column doesn't exist, create it
    if (qrPrintedCol === -1) {
      qrPrintedCol = headers.length;
      sheet.getRange(1, qrPrintedCol + 1).setValue('QR_Printed');
    }
    
    const data = sheet.getDataRange().getValues();
    const stdIdCol = headers.indexOf('Std_ID');
    
    let markedCount = 0;
    
    studentIds.forEach(stdId => {
      for (let i = 1; i < data.length; i++) {
        if (data[i][stdIdCol].toString() === stdId.toString()) {
          sheet.getRange(i + 1, qrPrintedCol + 1).setValue('Yes');
          markedCount++;
          break;
        }
      }
    });
    
    SpreadsheetApp.flush();
    return { success: true, count: markedCount };
    
  } catch (e) {
    Logger.log('Error marking students as printed: ' + e.message);
    return { success: false, message: e.message };
  }
}

// Generate reserve codes (bulk barcode generation)
function generateReserveCodes(count) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Student_Data');
    if (!sheet) throw new Error('Student_Data sheet not found');
    
    // Get the last used barcode number
    const lastRow = sheet.getLastRow();
    let lastBarcodeNumber = 1000000; // Start from 1000001
    
    if (lastRow > 1) {
      const barcodes = sheet.getRange('B2:B' + lastRow).getValues();
      const numbers = barcodes
        .filter(row => row[0] && typeof row[0] === 'string' && row[0].startsWith('BC'))
        .map(row => parseInt(row[0].replace('BC', '')))
        .filter(num => !isNaN(num));
      
      if (numbers.length > 0) {
        lastBarcodeNumber = Math.max(...numbers);
      }
    }
    
    // Generate reserve codes
    const reserveCodes = [];
    for (let i = 1; i <= count; i++) {
      const barcodeNumber = lastBarcodeNumber + i;
      reserveCodes.push('BC' + barcodeNumber);
    }
    
    return reserveCodes;
    
  } catch (e) {
    Logger.log('Error generating reserve codes: ' + e.message);
    throw e;
  }
}

// Generate PDF with multiple QR codes
function generateQRCodePDF(selectedStudents) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create HTML template for QR code printing
    const html = HtmlService.createHtmlOutputFromFile('QRPrintTemplate');
    html.setTitle('Student QR Codes');
    
    // Generate QR codes for selected students
    const studentsWithQR = selectedStudents.map(student => ({
      ...student,
      qrCodeUrl: generateStudentQRCode(student.Std_ID, student.Barcode_ID, student.Student_Name)
    }));
    
    // Pass data to template
    const output = HtmlService.createTemplate(html.getContent());
    output.students = studentsWithQR;
    output.pageTitle = 'Student QR Codes';
    
    return output.evaluate().getContent();
    
  } catch (e) {
    Logger.log('Error generating QR PDF: ' + e.message);
    throw e;
  }
}