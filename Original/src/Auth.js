/**
 * AUTHENTICATION MODULE & ADMIN ACTIONS
 * Handles Google OAuth, Domain validation, and administrative data management.
 */

// --- CONFIGURATION ---
// REQUIRED: Change this to your school's actual Google Workspace domain
const ALLOWED_DOMAIN = "centralreg.k12.nj.us"; 

// List specific administrative emails that should always have access regardless of domain.
const ALLOWED_ADMINS = [
  "admin@centralreg.k12.nj.us,ckleier@centralreg.k12.nj.us"
];

/**
 * Checks if the current user is authorized to access the system.
 * @param {string} email - Optional email parameter from Code.gs
 * @return {boolean}
 */
function checkUserAuth(email) {
  const userEmail = email || Session.getActiveUser().getEmail();
  
  if (!userEmail) return false;

  // 1. Check if user is in the explicit Admin list
  if (ALLOWED_ADMINS.map(e => e.toLowerCase()).includes(userEmail.toLowerCase())) {
    return true;
  }

  // 2. Check if the user belongs to the primary organization domain
  const domain = userEmail.split('@')[1];
  if (domain && domain.toLowerCase() === ALLOWED_DOMAIN.toLowerCase()) {
    return true;
  }

  return false;
}

/**
 * Helper to get the display domain for the UI.
 */
function getAuthorizedDomain() {
  return ALLOWED_DOMAIN;
}

/**
 * ADMIN ACTION: Reset points for a specific student.
 */
function adminResetStudentPoints(studentId, studentName) {
  if (!checkUserAuth()) throw new Error("Unauthorized access attempt.");
  
  // findStudent is located in Code.gs
  const student = findStudent(studentId);
  if (!student.found) return "Student not found in database.";
  
  const adjustment = -student.currentPoints;
  if (adjustment === 0) return "Student balance is already zero.";
  
  // submitPoints is located in Code.gs
  submitPoints(studentId, studentName, adjustment);
  return "Successfully reset " + studentName + " to 0 points.";
}

/**
 * ADMIN ACTION: Set a specific point score for a student.
 */
function adminSetStudentPoints(studentId, studentName, targetScore) {
  if (!checkUserAuth()) throw new Error("Unauthorized access attempt.");
  
  const student = findStudent(studentId);
  if (!student.found) return "Student not found in database.";
  
  const adjustment = targetScore - student.currentPoints;
  if (adjustment === 0) return "Student already has that score.";
  
  submitPoints(studentId, studentName, adjustment);
  return "Successfully adjusted " + studentName + " to " + targetScore + " points.";
}

/**
 * ADMIN ACTION: Archive points for the next year.
 */
function adminArchiveYearlyPoints() {
  if (!checkUserAuth()) throw new Error("Unauthorized access attempt.");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("PointsLog");
  
  // Create a descriptive name for the archive tab
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HHmm");
  const archiveName = "Archive_" + timestamp;
  
  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return "No data found in PointsLog to archive.";
  
  // Create the new archive sheet
  const archiveSheet = ss.insertSheet(archiveName);
  archiveSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  
  // Clear the active logs (keep the header row)
  if (logSheet.getLastRow() > 1) {
    logSheet.deleteRows(2, logSheet.getLastRow() - 1);
  }
  
  // Trigger a refresh of the leaderboard export if the function exists
  if (typeof updateSlideExportSheet === 'function') {
    updateSlideExportSheet();
  }
  
  return "All points archived to " + archiveName + ". Current log cleared.";
}