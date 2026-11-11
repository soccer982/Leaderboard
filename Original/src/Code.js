/**
 * EAGLE POINTS SYSTEM - MAIN BACKEND (Code.gs)
 * Optimized for 3-Column Leaderboard & Kiosk Modes
 */

function doGet(e) {
  const params = e.parameter || {};
  
  // 1. CAPTURE & CONVERT PARAMETERS
  // Using explicit comparison to ensure we have true Booleans
  const isvivi = params.vivi === 'true';
  const isteacher = params.teacher === 'true';
  const isMS = params.isMS === 'true'; // NEW: Capture MS parameter
  const adminPassword = params.p; 
  
  // 2. AUTHENTICATION
  const storedPass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASS') || "eagles1956";
  const isPasswordValid = adminPassword === storedPass; 

  // Security Gate: Redirect to Access Denied if not a public Kiosk and not authorized
  if (!isvivi && !isPasswordValid) {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail || !checkUserAuth(userEmail)) {
      const accessTemplate = HtmlService.createTemplateFromFile('AccessDenied');
      accessTemplate.ALLOWED_DOMAIN = typeof ALLOWED_DOMAIN !== 'undefined' ? ALLOWED_DOMAIN : "centralreg.k12.nj.us";
      accessTemplate.currentUserEmail = userEmail || "Not Signed In";
      accessTemplate.scriptUrl = ScriptApp.getService().getUrl();
      return accessTemplate.evaluate().setTitle("Access Denied");
    }
  }

  // 3. BRANDING & CONFIG
  const eagleBrand = {
    colors: {
      background: "#0f172a",
      accent: "#fbbf24",
      accentText: "#1e3a8a",
      gradStart: "#6366f1",
      gradEnd: "#a855f7",
      fontMain: "'Atkinson Hyperlegible', sans-serif"
    },
    messaging: {
      rookie: ["Welcome To The Flock", "First Points Secured", "Great Start Today"],
      regular: ["Consistency Is Key", "Reliable As Gravity", "The Streak Continues"],
      highFlyer: ["Outstanding Performance", "Setting The Standard", "Top Tier Effort"],
      grinder: ["Every Point Counts", "Keep Pushing Forward", "Building Momentum"]
    }
  };

  // 4. INJECT DATA INTO HTML
  const template = HtmlService.createTemplateFromFile('Index');
  template.brand = eagleBrand;
  template.isvivi = isvivi; 
  template.isTeacher = isteacher; 
  template.isAdmin = isPasswordValid; // If true, buttons will DEFINITELY show
  
  let pageTitle = isvivi ? "Eagle Leaderboard" : "Eagle Points";

  return template.evaluate()
    .setTitle(pageTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

/**
 * STUDENT LOOKUP
 */
function findStudent(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName("Students");
    const logSheet = ss.getSheetByName("PointsLog");
    const idStr = id.toString().trim(); 

    const studentData = studentSheet.getDataRange().getValues();
    const logData = logSheet.getDataRange().getValues();

    const studentRow = studentData.find(row => row[0].toString().trim() === idStr);
    if (!studentRow) return { found: false };

    const stats = logData.reduce((acc, row) => {
      if (row[1].toString().trim() === idStr) {
        acc.total += Number(row[2]) || 0;
        acc.scans += 1;
      }
      return acc;
    }, { total: 0, scans: 0 });

    return { 
      found: true, 
      name: `${studentRow[2]} ${studentRow[1]}`, 
      id: idStr, 
      currentPoints: stats.total, 
      scanCount: stats.scans 
    };
  } catch (err) { 
    return { found: false }; 
  }
}

/**
 * LOG POINTS
 */
function submitPoints(id, name, pts) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("PointsLog");
    const user = Session.getActiveUser().getEmail() || "Kiosk System";
    
    // Structure: [Date, ID, Points, Name, User]
    sheet.appendRow([new Date(), id, pts, name, user]);
    return true;
  } catch (e) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

/**
 * DASHBOARD DATA GENERATOR
 */
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logData = ss.getSheetByName("PointsLog").getDataRange().getValues();
    const studentData = ss.getSheetByName("Students").getDataRange().getValues();
    
    const now = new Date();
    const oneWeekAgo = new Date(now.getTime() - (7 * 24 * 60 * 60 * 1000));
    const todayMD = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM-dd");
    
    const studentMap = {};
    const birthdaysToday = [];
    
    // Populate Student Map
    for(let i = 1; i < studentData.length; i++) {
      const id = studentData[i][0].toString().trim();
      if (!id) continue;
      studentMap[id] = { 
        name: `${studentData[i][2]} ${studentData[i][1]}`, 
        grade: studentData[i][3] 
      };
      
      try { 
        if (Utilities.formatDate(new Date(studentData[i][4]), Session.getScriptTimeZone(), "MM-dd") === todayMD) {
          birthdaysToday.push({ name: studentData[i][2], type: 'birthday' }); 
        }
      } catch(e) {}
    }

    const stats = {};
    const recentScans = [];
    const seenRecent = new Set();

    // Process Points Log
    for(let i = logData.length - 1; i >= 1; i--) {
      const id = logData[i][1].toString().trim();
      const pts = Number(logData[i][2]) || 0;
      const date = new Date(logData[i][0]);

      if(!stats[id]) stats[id] = { total: 0, recent: 0 };
      stats[id].total += pts;
      if (date > oneWeekAgo) stats[id].recent += pts;

      if (recentScans.length < 15 && !seenRecent.has(id)) {
        recentScans.push({ name: (logData[i][3] || "Eagle").split(' ')[0], type: 'scan' });
        seenRecent.add(id);
      }
    }

    const allProcessed = Object.keys(stats).map(id => {
      const info = studentMap[id] || { name: "Unknown", grade: 0 };
      return { ...info, ...stats[id] };
    });

    const getT3P = (list) => list.sort((a, b) => b.total - a.total).slice(0, 3);
    const getT3G = (list) => list.filter(s => s.recent > 0).sort((a, b) => b.recent - a.recent).slice(0, 3);
    // NEW: Filter the overall list based on the school level
    let filteredProcessed = allProcessed;
    if (isMS) {
      filteredProcessed = allProcessed.filter(s => s.grade == 7 || s.grade == 8);
    } else {
      filteredProcessed = allProcessed.filter(s => s.grade >= 9);
    }
    return {
      overall: getT3P([...allProcessed]),
      overallGrowth: getT3G([...allProcessed]),
      grades: {
        7:  { points: getT3P(allProcessed.filter(s => s.grade == 7)), growth: getT3G(allProcessed.filter(s => s.grade == 7)) },
        8:  { points: getT3P(allProcessed.filter(s => s.grade == 8)), growth: getT3G(allProcessed.filter(s => s.grade == 8)) },
        9:  { points: getT3P(allProcessed.filter(s => s.grade == 9)), growth: getT3G(allProcessed.filter(s => s.grade == 9)) },
        10: { points: getT3P(allProcessed.filter(s => s.grade == 10)), growth: getT3G(allProcessed.filter(s => s.grade == 10)) },
        11: { points: getT3P(allProcessed.filter(s => s.grade == 11)), growth: getT3G(allProcessed.filter(s => s.grade == 11)) },
        12: { points: getT3P(allProcessed.filter(s => s.grade == 12)), growth: getT3G(allProcessed.filter(s => s.grade == 12)) }
      },
      recognition: [...birthdaysToday, ...recentScans].slice(0, 25)
    };
  } catch (e) { 
    return { error: e.toString() }; 
  }
}
/**
 * STUDENT NAME LOOKUP (MULTIPLE RESULTS)
 */
function findStudentsByName(nameQuery) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName("Students");
    const logSheet = ss.getSheetByName("PointsLog");
    
    // Split query into individual words for flexible matching (e.g., "John Jacobs")
    const queryWords = nameQuery.toString().toLowerCase().trim().split(/\s+/);
    const studentData = studentSheet.getDataRange().getValues();
    const logData = logSheet.getDataRange().getValues();
    
    // Filter for ALL matching students
    const matches = studentData.filter(row => {
      if (row[0] === "ID" || !row[0]) return false; // Skip headers/empty rows
      const fName = (row[2] || "").toString().toLowerCase();
      const lName = (row[1] || "").toString().toLowerCase();
      const fullName = `${fName} ${lName}`;
      
      // Ensure EVERY word in the query appears somewhere in the student's name
      return queryWords.every(word => fullName.includes(word));
    });

    if (matches.length === 0) return { found: false, results: [] };
    
    // Limit to top 15 results to prevent Kiosk UI overload
    const topMatches = matches.slice(0, 15);

    const results = topMatches.map(studentRow => {
      const idStr = studentRow[0].toString().trim();
      const stats = logData.reduce((acc, row) => {
        if (row[1].toString().trim() === idStr) {
          acc.total += Number(row[2]) || 0;
          acc.scans += 1;
        }
        return acc;
      }, { total: 0, scans: 0 });
      
      return { 
        found: true, 
        name: `${studentRow[2]} ${studentRow[1]}`, 
        id: idStr, 
        grade: studentRow[3],
        currentPoints: stats.total, 
        scanCount: stats.scans 
      };
    });
    
    return { found: true, results: results };
  } catch (err) { 
    return { found: false, results: [] }; 
  }
}
