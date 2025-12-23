// ⚙️ CONFIGURATION SETTINGS - Change these values to customize the system

// Check interval - Choose ONE option below:
// Option 1: Check every N hours (set CHECK_INTERVAL_HOURS, leave CHECK_INTERVAL_MINUTES as null)
const CHECK_INTERVAL_HOURS = 1; // Options: 1, 2, 3, 4, 6, 8, 12, or 24 hours
const CHECK_INTERVAL_MINUTES = null; // Set to null when using hours

// Option 2: Check every N minutes (set CHECK_INTERVAL_MINUTES, leave CHECK_INTERVAL_HOURS as null)
// const CHECK_INTERVAL_HOURS = null; // Set to null when using minutes
// const CHECK_INTERVAL_MINUTES = 1; // Options: 1, 5, 10, 15, or 30 minutes

// 🎯 Single File Mode - Process only ONE specific file
// Set to null to process all files in folder (normal mode)
// Set to filename (with extension) to process only that file
// const TARGET_SINGLE_FILE = null; // Example: "SwingVision-match-2025-11-09 at 15.59.30.xlsx"
 const TARGET_SINGLE_FILE = "SwingVision-match-2025-12-03 at 22.59.47"; // Uncomment to use

const SWINGVISION_FOLDER_NAME = "SwingVision"; // Name of the folder in Google Drive containing XLSX files
const ENABLE_DIAGNOSTIC_LOGGING = true; // Set to false to disable diagnostic logging to "Diagnostic" sheet
const STATS_SHEET_NAME = "Stats"; // Name of the sheet within each XLSX file to extract data from
const DATA_SHEET_NAME = "Match Data"; // Sheet where all match data will be stored
const CHARTS_SHEET_NAME = "Performance Charts"; // Sheet where charts will be displayed

/**
 * 📊 DIAGNOSTIC LOGGING SYSTEM
 * Logs system events to "Diagnostic" sheet for troubleshooting
 */

/**
 * Get or create the Diagnostic sheet
 */
function getDiagnosticSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let diagSheet = spreadsheet.getSheetByName("Diagnostic");
  
  if (!diagSheet) {
    diagSheet = spreadsheet.insertSheet("Diagnostic");
    
    // Set up headers
    const headers = [
      "Timestamp",
      "Event Type",
      "Details",
      "File Name",
      "Status",
      "Error"
    ];
    
    diagSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    diagSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    diagSheet.getRange(1, 1, 1, headers.length).setBackground("#4285F4");
    diagSheet.getRange(1, 1, 1, headers.length).setFontColor("#FFFFFF");
    
    // Set column widths
    diagSheet.setColumnWidth(1, 180); // Timestamp
    diagSheet.setColumnWidth(2, 150); // Event Type
    diagSheet.setColumnWidth(3, 400); // Details
    diagSheet.setColumnWidth(4, 250); // File Name
    diagSheet.setColumnWidth(5, 100); // Status
    diagSheet.setColumnWidth(6, 400); // Error
    
    diagSheet.setFrozenRows(1);
    
    console.log("✅ Created 'Diagnostic' sheet for logging");
  }
  
  return diagSheet;
}

/**
 * Log an event to the Diagnostic sheet
 */
function logDiagnostic(eventType, details, fileName = "", status = "SUCCESS", error = "") {
  if (!ENABLE_DIAGNOSTIC_LOGGING) {
    return; // Logging disabled
  }
  
  try {
    const diagSheet = getDiagnosticSheet();
    const timestamp = new Date();
    
    const logEntry = [
      timestamp,
      eventType,
      details,
      fileName,
      status,
      error
    ];
    
    // Insert at row 2 (after header) to keep newest entries at top
    diagSheet.insertRowAfter(1);
    diagSheet.getRange(2, 1, 1, logEntry.length).setValues([logEntry]);
    
    // Color code by status
    if (status === "SUCCESS") {
      diagSheet.getRange(2, 5).setBackground("#D5E8D4"); // Light green
    } else if (status === "WARNING") {
      diagSheet.getRange(2, 5).setBackground("#FFF4CE"); // Light yellow
    } else if (status === "ERROR") {
      diagSheet.getRange(2, 5).setBackground("#F4CCCC"); // Light red
    }
    
    // Keep only last 1000 entries to prevent sheet from growing too large
    const lastRow = diagSheet.getLastRow();
    if (lastRow > 1001) {
      diagSheet.deleteRows(1002, lastRow - 1001);
    }
    
  } catch (err) {
    // Don't let logging errors break the main functionality
    console.error(`⚠️ Diagnostic logging failed: ${err.message}`);
  }
}

/**
 * Clear the diagnostic log
 */
function clearDiagnosticLog() {
  try {
    const diagSheet = getDiagnosticSheet();
    const lastRow = diagSheet.getLastRow();
    
    if (lastRow > 1) {
      diagSheet.deleteRows(2, lastRow - 1);
      console.log(`✅ Cleared ${lastRow - 1} diagnostic log entries`);
    } else {
      console.log("📋 Diagnostic log is already empty");
    }
    
    logDiagnostic("SYSTEM", "Diagnostic log cleared", "", "SUCCESS");
    
  } catch (error) {
    console.error(`❌ Error clearing diagnostic log: ${error.message}`);
  }
}

/**
 * 🎾 MAIN TENNIS STATS ANALYZER FUNCTIONS
 */

/**
 * Install the time-driven trigger to check for new files periodically
 * Supports both minute-based and hour-based intervals
 */
function installTrigger() {
  try {
    // Remove any existing triggers for this function to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'checkForNewMatches') {
        ScriptApp.deleteTrigger(trigger);
      }
    }
    
    // Create new trigger based on configuration
    let triggerBuilder = ScriptApp.newTrigger('checkForNewMatches').timeBased();
    let intervalDescription = "";
    
    if (CHECK_INTERVAL_MINUTES !== null && CHECK_INTERVAL_MINUTES > 0) {
      // Use minute-based interval
      const validMinutes = [1, 5, 10, 15, 30];
      if (!validMinutes.includes(CHECK_INTERVAL_MINUTES)) {
        throw new Error(`Invalid CHECK_INTERVAL_MINUTES: ${CHECK_INTERVAL_MINUTES}. Must be one of: 1, 5, 10, 15, or 30`);
      }
      triggerBuilder.everyMinutes(CHECK_INTERVAL_MINUTES).create();
      intervalDescription = `${CHECK_INTERVAL_MINUTES} minute(s)`;
      
    } else if (CHECK_INTERVAL_HOURS !== null && CHECK_INTERVAL_HOURS > 0) {
      // Use hour-based interval
      const validHours = [1, 2, 3, 4, 6, 8, 12, 24];
      if (!validHours.includes(CHECK_INTERVAL_HOURS)) {
        throw new Error(`Invalid CHECK_INTERVAL_HOURS: ${CHECK_INTERVAL_HOURS}. Must be one of: 1, 2, 3, 4, 6, 8, 12, or 24`);
      }
      triggerBuilder.everyHours(CHECK_INTERVAL_HOURS).create();
      intervalDescription = `${CHECK_INTERVAL_HOURS} hour(s)`;
      
    } else {
      throw new Error("Either CHECK_INTERVAL_MINUTES or CHECK_INTERVAL_HOURS must be set (not both, not neither)");
    }
    
    logDiagnostic("SYSTEM", `Installed trigger to check for new matches every ${intervalDescription}`, "", "SUCCESS");
    console.log(`✅ Trigger installed: checking every ${intervalDescription}`);
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Failed to install trigger", "", "ERROR", error.message);
    console.error(`❌ Error installing trigger: ${error.message}`);
    throw error;
  }
}

/**
 * Uninstall all triggers for this script
 */
function uninstallTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let count = 0;
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'checkForNewMatches') {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
    }
    
    logDiagnostic("SYSTEM", `Uninstalled ${count} trigger(s)`, "", "SUCCESS");
    console.log(`✅ Uninstalled ${count} trigger(s)`);
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Failed to uninstall triggers", "", "ERROR", error.message);
    console.error(`❌ Error uninstalling triggers: ${error.message}`);
  }
}

/**
 * Find the SwingVision folder in Google Drive
 */
function getSwingVisionFolder() {
  try {
    const folders = DriveApp.getFoldersByName(SWINGVISION_FOLDER_NAME);
    
    if (!folders.hasNext()) {
      throw new Error(`Folder "${SWINGVISION_FOLDER_NAME}" not found in Google Drive`);
    }
    
    const folder = folders.next();
    
    // If multiple folders with same name, warn but use the first one
    if (folders.hasNext()) {
      logDiagnostic("FOLDER", `Multiple folders named "${SWINGVISION_FOLDER_NAME}" found. Using the first one.`, "", "WARNING");
    }
    
    return folder;
    
  } catch (error) {
    logDiagnostic("FOLDER", "Failed to find SwingVision folder", "", "ERROR", error.message);
    throw error;
  }
}

/**
 * Get or create the Match Data sheet
 */
function getMatchDataSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
  
  if (!dataSheet) {
    dataSheet = spreadsheet.insertSheet(DATA_SHEET_NAME);
    
    // Set up headers - based on actual SwingVision Stats sheet format
    const headers = [
      "Match Date",
      "File Name",
      "Opponent",
      "Result",
      "Score",
      // Serve Statistics
      "First Serve %",
      "First Serve Points Won %",
      "Second Serve Points Won %",
      "Aces",
      "Double Faults",
      // Break Points
      "Break Points Won",
      "Break Points Total",
      "Break Point Conversion %",
      // Winners & Errors
      "Total Winners",
      "Service Winners",
      "Forehand Winners",
      "Backhand Winners",
      "Total Unforced Errors",
      "Forehand UE",
      "Backhand UE",
      "Winners/UE Ratio",
      // Match Totals
      "Total Points Won",
      "Total Points",
      "Points Won %",
      "Games Won",
      "Games Total",
      // Physical Stats
      "Calories Burned",
      "Distance Run (mi)",
      "Avg Heart Rate (BPM)",
      // Speed Statistics (from Shots sheet)
      "Avg 1st Serve Speed (mph)",
      "Avg 2nd Serve Speed (mph)",
      "Avg Forehand Speed (mph)",
      "Avg Backhand Speed (mph)",
      // Serve Spin Distribution
      "Serve Flat Count",
      "Serve Kick Count",
      "Serve Slice Count",
      // Forehand Spin Distribution
      "FH Topspin Count",
      "FH Flat Count",
      "FH Slice Count",
      // Backhand Spin Distribution
      "BH Topspin Count",
      "BH Flat Count",
      "BH Slice Count",
      // 🆕 Unforced Error Analysis (Net vs Out)
      "FH Errors Net",
      "FH Errors Out",
      "BH Errors Net",
      "BH Errors Out",
      // 🆕 Error Spin Analysis - Forehand Net
      "FH Net Topspin",
      "FH Net Flat",
      "FH Net Slice",
      // 🆕 Error Spin Analysis - Forehand Out
      "FH Out Topspin",
      "FH Out Flat",
      "FH Out Slice",
      // 🆕 Error Spin Analysis - Backhand Net
      "BH Net Topspin",
      "BH Net Flat",
      "BH Net Slice",
      // 🆕 Error Spin Analysis - Backhand Out
      "BH Out Topspin",
      "BH Out Flat",
      "BH Out Slice",
      // Metadata
      "File ID",
      "Processed Date"
    ];
    
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dataSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    dataSheet.getRange(1, 1, 1, headers.length).setBackground("#34A853");
    dataSheet.getRange(1, 1, 1, headers.length).setFontColor("#FFFFFF");
    dataSheet.setFrozenRows(1);
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      dataSheet.autoResizeColumn(i);
    }
    
    logDiagnostic("SYSTEM", "Created Match Data sheet", "", "SUCCESS");
  }
  
  return dataSheet;
}

/**
 * Extract date from filename
 * Supports multiple formats:
 * - SwingVision format: "SwingVision-match-2025-11-09 at 15.59.30"
 * - Standard formats: "2024-03-15" or "20240315"
 * - US format: "03-15-2024" or "03/15/2024"
 */
function extractDateFromFilename(filename) {
  try {
    // Try to match SwingVision format: YYYY-MM-DD at HH.MM.SS
    let match = filename.match(/(\d{4})-(\d{2})-(\d{2})\s+at\s+(\d{1,2})\.(\d{2})\.(\d{2})/);
    if (match) {
      // Extract date and time
      const year = parseInt(match[1]);
      const month = parseInt(match[2]) - 1; // JavaScript months are 0-indexed
      const day = parseInt(match[3]);
      const hour = parseInt(match[4]);
      const minute = parseInt(match[5]);
      const second = parseInt(match[6]);
      
      return new Date(year, month, day, hour, minute, second);
    }
    
    // Try to match standard YYYY-MM-DD format
    match = filename.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      return new Date(match[1], parseInt(match[2]) - 1, match[3]);
    }
    
    // Try to match YYYYMMDD format
    match = filename.match(/(\d{4})(\d{2})(\d{2})/);
    if (match) {
      return new Date(match[1], parseInt(match[2]) - 1, match[3]);
    }
    
    // Try to match MM-DD-YYYY or MM/DD/YYYY format
    match = filename.match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
    if (match) {
      return new Date(match[3], parseInt(match[1]) - 1, match[2]);
    }
    
    // If no date found, return null
    return null;
    
  } catch (error) {
    console.error(`Error extracting date from filename "${filename}": ${error.message}`);
    return null;
  }
}

/**
 * Check if a file has already been processed
 */
function isFileProcessed(fileId) {
  const dataSheet = getMatchDataSheet();
  const data = dataSheet.getDataRange().getValues();
  
  // Check if file ID exists in the File ID column (column 22)
  for (let i = 1; i < data.length; i++) {
    if (data[i][21] === fileId) {
      return true;
    }
  }
  
  return false;
}

/**
 * Extract stats from the Stats sheet of an XLSX file
 * 
 * ⭐ IMPORTANT: This function extracts ONLY YOUR (Host) statistics!
 * - Stats sheet: Sums only Host columns (B, D, F, H, J)
 * - Shots sheet: Filters only Host player shots
 * - Opponent data is completely ignored
 */
function extractStatsFromFile(file) {
  try {
    const fileId = file.getId();
    const fileName = file.getName();
    
    logDiagnostic("PROCESSING", `Opening file to extract HOST (your) stats only`, fileName, "SUCCESS");
    
    // Open the spreadsheet file
    const spreadsheet = SpreadsheetApp.openById(fileId);
    const statsSheet = spreadsheet.getSheetByName(STATS_SHEET_NAME);
    
    if (!statsSheet) {
      throw new Error(`Sheet "${STATS_SHEET_NAME}" not found in file`);
    }
    
    // Get all data from the Stats sheet
    const statsData = statsSheet.getDataRange().getValues();
    
    // Extract the match date from filename
    const matchDate = extractDateFromFilename(fileName);
    
    // Parse the stats from Stats sheet
    const stats = parseSwingVisionStats(statsData, fileName);
    
    // Parse detailed shot data from Shots sheet for speed and spin statistics
    const shotsSheet = spreadsheet.getSheetByName("Shots");
    if (shotsSheet) {
      const shotsData = shotsSheet.getDataRange().getValues();
      const shotStats = parseShotsData(shotsData, fileName);
      Object.assign(stats, shotStats);
    } else {
      logDiagnostic("PROCESSING", "Shots sheet not found - speed/spin stats unavailable", fileName, "WARNING");
    }
    
    return {
      matchDate: matchDate,
      fileName: fileName,
      fileId: fileId,
      ...stats
    };
    
  } catch (error) {
    logDiagnostic("PROCESSING", "Failed to extract stats from file", file.getName(), "ERROR", error.message);
    throw error;
  }
}

/**
 * Parse SwingVision stats data
 * SwingVision format: Column A = Stat Name, Columns B-K = values for Host/Guest by set
 * 
 * IMPORTANT: We ONLY track YOUR (Host) stats, not opponent stats!
 * 
 * Host columns (YOU): B, D, F, H, J (indices 1, 3, 5, 7, 9) ⭐
 * Guest columns (opponent): C, E, G, I, K (indices 2, 4, 6, 8, 10) - IGNORED
 * 
 * The parser sums your stats across all sets to get match totals.
 */
function parseSwingVisionStats(data, fileName) {
  try {
    const stats = {
      opponent: "Unknown",
      result: "",
      score: "",
      firstServePercent: 0,
      firstServePointsWonPercent: 0,
      secondServePointsWonPercent: 0,
      breakPointsWon: 0,
      breakPointsTotal: 0,
      breakPointConversionPercent: 0,
      winners: 0,
      unforcedErrors: 0,
      winnersUERatio: 0,
      aces: 0,
      doubleFaults: 0,
      totalPointsWon: 0,
      totalPoints: 0,
      pointsWonPercent: 0,
      gamesWon: 0,
      gamesTotal: 0
    };
    
    // Helper function to sum Host columns (B, D, F, H, J = indices 1, 3, 5, 7, 9)
    // This extracts ONLY YOUR stats across all sets
    function sumHostColumns(row) {
      let sum = 0;
      for (let col = 1; col < row.length; col += 2) { // Host columns: 1, 3, 5, 7, 9
        const val = parseFloat(row[col]);
        if (!isNaN(val)) {
          sum += val;
        }
      }
      return sum;
    }
    
    // Helper function to sum Guest columns (C, E, G, I, K = indices 2, 4, 6, 8, 10)
    // NOTE: Only used to calculate total points in match (Host + Guest points)
    // We don't track or store opponent stats otherwise
    function sumGuestColumns(row) {
      let sum = 0;
      for (let col = 2; col < row.length; col += 2) { // Guest columns: 2, 4, 6, 8, 10
        const val = parseFloat(row[col]);
        if (!isNaN(val)) {
          sum += val;
        }
      }
      return sum;
    }
    
    // Parse the data array to extract statistics
    // SwingVision format has stat names in column A, values in columns B-K
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[0]) continue;
      
      const label = String(row[0]).toLowerCase().trim();
      
      // Match stat labels and sum values across all sets for Host (you)
      if (label === "1st serves") {
        stats.firstServes = sumHostColumns(row);
      } else if (label === "1st serves in") {
        stats.firstServesIn = sumHostColumns(row);
      } else if (label === "1st serves won") {
        stats.firstServesWon = sumHostColumns(row);
      } else if (label === "2nd serves") {
        stats.secondServes = sumHostColumns(row);
      } else if (label === "2nd serves in") {
        stats.secondServesIn = sumHostColumns(row);
      } else if (label === "2nd serves won") {
        stats.secondServesWon = sumHostColumns(row);
      } else if (label === "break points") {
        stats.breakPointsAgainst = sumHostColumns(row);
      } else if (label === "break points saved") {
        stats.breakPointsSaved = sumHostColumns(row);
      } else if (label === "break point opportunities") {
        stats.breakPointsTotal = sumHostColumns(row);
      } else if (label === "break points won") {
        stats.breakPointsWon = sumHostColumns(row);
      } else if (label === "total points") {
        stats.totalPoints = sumHostColumns(row) + sumGuestColumns(row);
      } else if (label === "total points won") {
        stats.totalPointsWon = sumHostColumns(row);
      } else if (label === "aces") {
        stats.aces = sumHostColumns(row);
      } else if (label === "service winners") {
        stats.serviceWinners = sumHostColumns(row);
      } else if (label === "forehand winners") {
        stats.forehandWinners = sumHostColumns(row);
      } else if (label === "backhand winners") {
        stats.backhandWinners = sumHostColumns(row);
      } else if (label === "forehand unforced errors") {
        stats.forehandUE = sumHostColumns(row);
      } else if (label === "backhand unforced errors") {
        stats.backhandUE = sumHostColumns(row);
      } else if (label === "forehand forced errors") {
        stats.forehandFE = sumHostColumns(row);
      } else if (label === "backhand forced errors") {
        stats.backhandFE = sumHostColumns(row);
      } else if (label === "calories burned (cal)") {
        stats.caloriesBurned = sumHostColumns(row);
      } else if (label === "distance run (mi)") {
        stats.distanceRun = sumHostColumns(row);
      } else if (label === "average heart rate (bpm)") {
        const totalBPM = sumHostColumns(row);
        const numSets = row.slice(1).filter(v => v && parseFloat(v) > 0).length / 2;
        stats.avgHeartRate = numSets > 0 ? totalBPM / numSets : 0;
      }
    }
    
    // Calculate aggregate stats
    stats.winners = (stats.serviceWinners || 0) + (stats.forehandWinners || 0) + (stats.backhandWinners || 0);
    stats.unforcedErrors = (stats.forehandUE || 0) + (stats.backhandUE || 0);
    
    // Calculate percentages
    if (stats.firstServes > 0) {
      stats.firstServePercent = (stats.firstServesIn / stats.firstServes) * 100;
    }
    
    if (stats.firstServesIn > 0) {
      stats.firstServePointsWonPercent = (stats.firstServesWon / stats.firstServesIn) * 100;
    }
    
    if (stats.secondServesIn > 0) {
      stats.secondServePointsWonPercent = (stats.secondServesWon / stats.secondServesIn) * 100;
    }
    
    if (stats.breakPointsTotal > 0) {
      stats.breakPointConversionPercent = (stats.breakPointsWon / stats.breakPointsTotal) * 100;
    }
    
    if (stats.unforcedErrors > 0) {
      stats.winnersUERatio = stats.winners / stats.unforcedErrors;
    }
    
    if (stats.totalPoints > 0) {
      stats.pointsWonPercent = (stats.totalPointsWon / stats.totalPoints) * 100;
    }
    
    // Try to determine result from points won
    if (stats.pointsWonPercent > 50) {
      stats.result = "Win";
    } else if (stats.pointsWonPercent > 0) {
      stats.result = "Loss";
    }
    
    return stats;
    
  } catch (error) {
    logDiagnostic("PARSING", "Failed to parse stats data", fileName, "ERROR", error.message);
    throw error;
  }
}

/**
 * Parse SwingVision Shots sheet for detailed speed and spin statistics
 * Shots sheet columns: Player, Shot, Type, Stroke, Spin, Speed (MPH), ...
 * 
 * ⭐ CRITICAL: This function processes ONLY YOUR (Host) shots!
 * - Filters out "Opponent" shots (line 661-663)
 * - Only counts YOUR serves, forehands, backhands
 * - Speed averages are for YOUR shots only
 * - Spin distributions are for YOUR shots only
 */
function parseShotsData(data, fileName) {
  try {
    const shotStats = {
      // Serve speeds (YOUR serves only)
      avg1stServeSpeed: 0,
      avg2ndServeSpeed: 0,
      // Shot speeds (YOUR shots only)
      avgForehandSpeed: 0,
      avgBackhandSpeed: 0,
      // Serve spin distribution (YOUR serves only)
      serveFlat: 0,
      serveKick: 0,
      serveSlice: 0,
      // Forehand spin distribution (YOUR forehands only)
      forehandTopspin: 0,
      forehandFlat: 0,
      forehandSlice: 0,
      // Backhand spin distribution (YOUR backhands only)
      backhandTopspin: 0,
      backhandFlat: 0,
      backhandSlice: 0,
      // 🆕 Unforced Error Analysis (Net vs Out)
      forehandErrorsNet: 0,
      forehandErrorsOut: 0,
      backhandErrorsNet: 0,
      backhandErrorsOut: 0,
      // 🆕 Spin used on unforced errors
      forehandErrorsNetTopspin: 0,
      forehandErrorsNetFlat: 0,
      forehandErrorsNetSlice: 0,
      forehandErrorsOutTopspin: 0,
      forehandErrorsOutFlat: 0,
      forehandErrorsOutSlice: 0,
      backhandErrorsNetTopspin: 0,
      backhandErrorsNetFlat: 0,
      backhandErrorsNetSlice: 0,
      backhandErrorsOutTopspin: 0,
      backhandErrorsOutFlat: 0,
      backhandErrorsOutSlice: 0
    };
    
    // Skip header row (row 0)
    // Columns: A=Player, B=Shot, C=Type, D=Stroke, E=Spin, F=Speed(MPH), ... V=Result
    
    // Collections for speed calculations
    const serve1stSpeeds = [];
    const serve2ndSpeeds = [];
    const forehandSpeeds = [];
    const backhandSpeeds = [];
    
    // Determine player name (host) from first data row
    let playerName = null;
    for (let i = 1; i < data.length && i < 10; i++) {
      const player = data[i][0];
      if (player && player !== "Opponent" && String(player).trim() !== "") {
        playerName = String(player).trim();
        break;
      }
    }
    
    // Process each shot - BUT ONLY YOURS (Host shots)!
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const player = String(row[0] || "").trim();
      const shotNumber = row[1];
      const type = String(row[2] || "").toLowerCase().trim();
      const stroke = String(row[3] || "").toLowerCase().trim();
      const spin = String(row[4] || "").toLowerCase().trim();
      const speed = parseFloat(row[5]);
      const result = String(row[21] || "").trim();  // Column V (index 21) = Result
      
      // ⭐ CRITICAL FILTER: Only process HOST (YOUR) shots - skip opponent shots!
      // Skip if: empty player, explicitly "Opponent", or doesn't match your player name
      if (!player || player === "Opponent" || (playerName && player !== playerName)) {
        continue;  // Skip this shot - it's not yours!
      }
      
      // Process serves
      if (stroke.includes("serve")) {
        // Determine if 1st or 2nd serve based on shot number pattern
        // In SwingVision, shot 0 or even shots are typically 1st serves
        const isFirstServe = (shotNumber === 0 || shotNumber % 2 === 0);
        
        if (!isNaN(speed) && speed > 0) {
          if (isFirstServe) {
            serve1stSpeeds.push(speed);
          } else {
            serve2ndSpeeds.push(speed);
          }
        }
        
        // Count serve spin types
        if (spin.includes("flat")) {
          shotStats.serveFlat++;
        } else if (spin.includes("kick") || spin.includes("topspin")) {
          shotStats.serveKick++;
        } else if (spin.includes("slice")) {
          shotStats.serveSlice++;
        }
      }
      
      // Process forehands
      else if (stroke.includes("forehand")) {
        if (!isNaN(speed) && speed > 0) {
          forehandSpeeds.push(speed);
        }
        
        if (spin.includes("topspin")) {
          shotStats.forehandTopspin++;
        } else if (spin.includes("flat")) {
          shotStats.forehandFlat++;
        } else if (spin.includes("slice") || spin.includes("backspin")) {
          shotStats.forehandSlice++;
        }
      }
      
      // Process backhands
      else if (stroke.includes("backhand")) {
        if (!isNaN(speed) && speed > 0) {
          backhandSpeeds.push(speed);
        }
        
        if (spin.includes("topspin")) {
          shotStats.backhandTopspin++;
        } else if (spin.includes("flat")) {
          shotStats.backhandFlat++;
        } else if (spin.includes("slice") || spin.includes("backspin")) {
          shotStats.backhandSlice++;
        }
      }
      
      // 🆕 Process UNFORCED ERRORS - Net vs Out analysis
      // Filter: Type != "none" and Type != "in_play" and Result != "In"
      const isUnforcedError = type !== "none" && type !== "in_play" && result !== "In";
      
      if (isUnforcedError && (stroke.includes("forehand") || stroke.includes("backhand"))) {
        const isNet = result === "Net";
        const isOut = result === "Out" || result === "Long" || result === "Wide";
        
        // Forehand unforced errors
        if (stroke.includes("forehand")) {
          if (isNet) {
            shotStats.forehandErrorsNet++;
            // Track spin used on net errors
            if (spin.includes("topspin")) {
              shotStats.forehandErrorsNetTopspin++;
            } else if (spin.includes("flat")) {
              shotStats.forehandErrorsNetFlat++;
            } else if (spin.includes("slice")) {
              shotStats.forehandErrorsNetSlice++;
            }
          } else if (isOut) {
            shotStats.forehandErrorsOut++;
            // Track spin used on out errors
            if (spin.includes("topspin")) {
              shotStats.forehandErrorsOutTopspin++;
            } else if (spin.includes("flat")) {
              shotStats.forehandErrorsOutFlat++;
            } else if (spin.includes("slice")) {
              shotStats.forehandErrorsOutSlice++;
            }
          }
        }
        
        // Backhand unforced errors
        if (stroke.includes("backhand")) {
          if (isNet) {
            shotStats.backhandErrorsNet++;
            // Track spin used on net errors
            if (spin.includes("topspin")) {
              shotStats.backhandErrorsNetTopspin++;
            } else if (spin.includes("flat")) {
              shotStats.backhandErrorsNetFlat++;
            } else if (spin.includes("slice")) {
              shotStats.backhandErrorsNetSlice++;
            }
          } else if (isOut) {
            shotStats.backhandErrorsOut++;
            // Track spin used on out errors
            if (spin.includes("topspin")) {
              shotStats.backhandErrorsOutTopspin++;
            } else if (spin.includes("flat")) {
              shotStats.backhandErrorsOutFlat++;
            } else if (spin.includes("slice")) {
              shotStats.backhandErrorsOutSlice++;
            }
          }
        }
      }
    }
    
    // Calculate averages
    if (serve1stSpeeds.length > 0) {
      shotStats.avg1stServeSpeed = serve1stSpeeds.reduce((a, b) => a + b, 0) / serve1stSpeeds.length;
    }
    
    if (serve2ndSpeeds.length > 0) {
      shotStats.avg2ndServeSpeed = serve2ndSpeeds.reduce((a, b) => a + b, 0) / serve2ndSpeeds.length;
    }
    
    if (forehandSpeeds.length > 0) {
      shotStats.avgForehandSpeed = forehandSpeeds.reduce((a, b) => a + b, 0) / forehandSpeeds.length;
    }
    
    if (backhandSpeeds.length > 0) {
      shotStats.avgBackhandSpeed = backhandSpeeds.reduce((a, b) => a + b, 0) / backhandSpeeds.length;
    }
    
    return shotStats;
    
  } catch (error) {
    logDiagnostic("PARSING", "Failed to parse shots data", fileName, "ERROR", error.message);
    console.error(`Error parsing shots data: ${error.message}`);
    // Return empty stats on error - don't fail the whole file processing
    return {};
  }
}

/**
 * Add match data to the Match Data sheet
 */
function addMatchData(matchData) {
  try {
    const dataSheet = getMatchDataSheet();
    
    const rowData = [
      matchData.matchDate,
      matchData.fileName,
      matchData.opponent,
      matchData.result,
      matchData.score,
      // Serve Statistics
      matchData.firstServePercent || 0,
      matchData.firstServePointsWonPercent || 0,
      matchData.secondServePointsWonPercent || 0,
      matchData.aces || 0,
      matchData.doubleFaults || 0,
      // Break Points
      matchData.breakPointsWon || 0,
      matchData.breakPointsTotal || 0,
      matchData.breakPointConversionPercent || 0,
      // Winners & Errors
      matchData.winners || 0,
      matchData.serviceWinners || 0,
      matchData.forehandWinners || 0,
      matchData.backhandWinners || 0,
      matchData.unforcedErrors || 0,
      matchData.forehandUE || 0,
      matchData.backhandUE || 0,
      matchData.winnersUERatio || 0,
      // Match Totals
      matchData.totalPointsWon || 0,
      matchData.totalPoints || 0,
      matchData.pointsWonPercent || 0,
      matchData.gamesWon || 0,
      matchData.gamesTotal || 0,
      // Physical Stats
      matchData.caloriesBurned || 0,
      matchData.distanceRun || 0,
      matchData.avgHeartRate || 0,
      // Speed Statistics
      matchData.avg1stServeSpeed || 0,
      matchData.avg2ndServeSpeed || 0,
      matchData.avgForehandSpeed || 0,
      matchData.avgBackhandSpeed || 0,
      // Serve Spin Distribution
      matchData.serveFlat || 0,
      matchData.serveKick || 0,
      matchData.serveSlice || 0,
      // Forehand Spin Distribution
      matchData.forehandTopspin || 0,
      matchData.forehandFlat || 0,
      matchData.forehandSlice || 0,
      // Backhand Spin Distribution
      matchData.backhandTopspin || 0,
      matchData.backhandFlat || 0,
      matchData.backhandSlice || 0,
      // 🆕 Unforced Error Analysis
      matchData.forehandErrorsNet || 0,
      matchData.forehandErrorsOut || 0,
      matchData.backhandErrorsNet || 0,
      matchData.backhandErrorsOut || 0,
      // 🆕 Error Spin Analysis - Forehand Net
      matchData.forehandErrorsNetTopspin || 0,
      matchData.forehandErrorsNetFlat || 0,
      matchData.forehandErrorsNetSlice || 0,
      // 🆕 Error Spin Analysis - Forehand Out
      matchData.forehandErrorsOutTopspin || 0,
      matchData.forehandErrorsOutFlat || 0,
      matchData.forehandErrorsOutSlice || 0,
      // 🆕 Error Spin Analysis - Backhand Net
      matchData.backhandErrorsNetTopspin || 0,
      matchData.backhandErrorsNetFlat || 0,
      matchData.backhandErrorsNetSlice || 0,
      // 🆕 Error Spin Analysis - Backhand Out
      matchData.backhandErrorsOutTopspin || 0,
      matchData.backhandErrorsOutFlat || 0,
      matchData.backhandErrorsOutSlice || 0,
      // Metadata
      matchData.fileId,
      new Date()
    ];
    
    dataSheet.appendRow(rowData);
    
    // Sort by match date (column 1)
    const lastRow = dataSheet.getLastRow();
    if (lastRow > 2) {
      const dataRange = dataSheet.getRange(2, 1, lastRow - 1, rowData.length);
      dataRange.sort({column: 1, ascending: true});
    }
    
    logDiagnostic("DATA", "Added match data to sheet", matchData.fileName, "SUCCESS");
    
  } catch (error) {
    logDiagnostic("DATA", "Failed to add match data", matchData.fileName, "ERROR", error.message);
    throw error;
  }
}

/**
 * Main function to check for new matches and process them
 */
function checkForNewMatches() {
  try {
    const mode = TARGET_SINGLE_FILE ? `Single File Mode: ${TARGET_SINGLE_FILE}` : "All Files Mode";
    logDiagnostic("SYSTEM", `Starting check for new matches (${mode})`, "", "SUCCESS");
    
    const folder = getSwingVisionFolder();
    
    let newFilesCount = 0;
    let processedFilesCount = 0;
    let errorCount = 0;
    
    // 🎯 Single File Mode: Process only one specific file
    if (TARGET_SINGLE_FILE) {
      try {
        const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
        let targetFile = null;
        
        // Search for the specific file in the folder
        while (files.hasNext()) {
          const file = files.next();
          if (file.getName() === TARGET_SINGLE_FILE) {
            targetFile = file;
            break;
          }
        }
        
        if (!targetFile) {
          const message = `Target file "${TARGET_SINGLE_FILE}" not found in ${SWINGVISION_FOLDER_NAME} folder`;
          logDiagnostic("SYSTEM", message, "", "ERROR");
          console.error(`❌ ${message}`);
          return;
        }
        
        const fileId = targetFile.getId();
        const fileName = targetFile.getName();
        
        // Check if already processed
        if (isFileProcessed(fileId)) {
          const message = `Target file "${TARGET_SINGLE_FILE}" has already been processed`;
          logDiagnostic("SYSTEM", message, "", "WARNING");
          console.log(`⚠️ ${message}`);
          return;
        }
        
        newFilesCount = 1;
        
        try {
          logDiagnostic("PROCESSING", "Processing target file", fileName, "SUCCESS");
          
          // Extract stats from the file
          const matchData = extractStatsFromFile(targetFile);
          
          // Add to Match Data sheet
          addMatchData(matchData);
          
          processedFilesCount++;
          logDiagnostic("PROCESSING", "Successfully processed file", fileName, "SUCCESS");
          
        } catch (error) {
          errorCount++;
          logDiagnostic("PROCESSING", "Failed to process file", fileName, "ERROR", error.message);
          console.error(`❌ Error processing file "${fileName}": ${error.message}`);
        }
        
      } catch (error) {
        logDiagnostic("SYSTEM", "Single file mode failed", TARGET_SINGLE_FILE, "ERROR", error.message);
        console.error(`❌ Error in single file mode: ${error.message}`);
      }
      
    } else {
      // 📁 All Files Mode: Process all unprocessed files
      const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
      
      while (files.hasNext()) {
        const file = files.next();
        const fileId = file.getId();
        const fileName = file.getName();
        
        // Check if file has already been processed
        if (isFileProcessed(fileId)) {
          continue;
        }
        
        newFilesCount++;
        
        try {
          logDiagnostic("PROCESSING", "Found new file", fileName, "SUCCESS");
          
          // Extract stats from the file
          const matchData = extractStatsFromFile(file);
          
          // Add to Match Data sheet
          addMatchData(matchData);
          
          processedFilesCount++;
          logDiagnostic("PROCESSING", "Successfully processed file", fileName, "SUCCESS");
          
        } catch (error) {
          errorCount++;
          logDiagnostic("PROCESSING", "Failed to process file", fileName, "ERROR", error.message);
          console.error(`❌ Error processing file "${fileName}": ${error.message}`);
        }
      }
    }
    
    // Update charts after processing new files
    if (processedFilesCount > 0) {
      updateCharts();
    }
    
    const summary = `Found ${newFilesCount} new file(s), processed ${processedFilesCount}, errors: ${errorCount}`;
    logDiagnostic("SYSTEM", summary, "", processedFilesCount > 0 ? "SUCCESS" : "WARNING");
    console.log(`✅ ${summary}`);
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Check for new matches failed", "", "ERROR", error.message);
    console.error(`❌ Error checking for new matches: ${error.message}`);
  }
}

/**
 * 📊 CHART CREATION FUNCTIONS
 */

/**
 * Get or create the Performance Charts sheet
 */
function getChartsSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let chartsSheet = spreadsheet.getSheetByName(CHARTS_SHEET_NAME);
  
  if (!chartsSheet) {
    chartsSheet = spreadsheet.insertSheet(CHARTS_SHEET_NAME);
    logDiagnostic("SYSTEM", "Created Performance Charts sheet", "", "SUCCESS");
  }
  
  return chartsSheet;
}

/**
 * Create or update all performance charts
 */
function updateCharts() {
  try {
    logDiagnostic("CHARTS", "Starting chart update", "", "SUCCESS");
    
    const dataSheet = getMatchDataSheet();
    const chartsSheet = getChartsSheet();
    
    // Clear existing charts
    const charts = chartsSheet.getCharts();
    charts.forEach(chart => chartsSheet.removeChart(chart));
    
    // Clear the sheet
    chartsSheet.clear();
    
    // Add title
    chartsSheet.getRange("A1").setValue("🎾 Tennis Performance Analysis");
    chartsSheet.getRange("A1").setFontSize(18).setFontWeight("bold");
    
    const lastRow = dataSheet.getLastRow();
    
    if (lastRow <= 1) {
      chartsSheet.getRange("A3").setValue("No match data available yet. Process some match files to see charts.");
      logDiagnostic("CHARTS", "No data available for charts", "", "WARNING");
      return;
    }
    
    // Create multiple charts
    let chartRow = 3;
    
    // 1. Winners vs Unforced Errors Over Time
    createWinnersUEChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 2. First Serve % and Points Won % Over Time
    createServeStatsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 3. Match Results Timeline
    createResultsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 4. Serve Speed Trends (NEW!)
    createServeSpeedChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 5. Shot Speed Comparison (NEW!)
    createShotSpeedChart(dataSheet, chartsSheet, chartRow, lastRow);
    
    logDiagnostic("CHARTS", "Successfully updated all charts", "", "SUCCESS");
    console.log("✅ Charts updated successfully");
    
  } catch (error) {
    logDiagnostic("CHARTS", "Failed to update charts", "", "ERROR", error.message);
    console.error(`❌ Error updating charts: ${error.message}`);
  }
}

/**
 * Create Winners vs Unforced Errors chart
 */
function createWinnersUEChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 14, lastDataRow, 1)) // Total Winners (column 14)
      .addRange(dataSheet.getRange(1, 18, lastDataRow, 1)) // Total Unforced Errors (column 18)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', 'Winners vs Unforced Errors Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count'})
      .setOption('series', {
        0: {color: '#34A853', lineWidth: 3},
        1: {color: '#EA4335', lineWidth: 3}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Winners/UE chart: ${error.message}`);
  }
}

/**
 * Create Serve Statistics chart
 */
function createServeStatsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 6, lastDataRow, 1)) // First Serve % (column 6)
      .addRange(dataSheet.getRange(1, 7, lastDataRow, 1)) // First Serve Points Won % (column 7)
      .addRange(dataSheet.getRange(1, 24, lastDataRow, 1)) // Points Won % (column 24)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', 'Serve Statistics Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('series', {
        0: {color: '#4285F4', lineWidth: 3},
        1: {color: '#FBBC04', lineWidth: 3},
        2: {color: '#34A853', lineWidth: 3}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Serve Stats chart: ${error.message}`);
  }
}

/**
 * Create Match Results chart
 */
function createResultsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // This creates a simple visualization of match results
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 21, lastDataRow, 1)) // Winners/UE Ratio (column 21)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', 'Winners/Unforced Errors Ratio Over Time (Higher is Better)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'none'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Winners/UE Ratio', minValue: 0})
      .setOption('colors', ['#34A853'])
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Results chart: ${error.message}`);
  }
}

/**
 * Create Serve Speed Trends chart (NEW!)
 */
function createServeSpeedChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 30, lastDataRow, 1)) // Avg 1st Serve Speed (column 30)
      .addRange(dataSheet.getRange(1, 31, lastDataRow, 1)) // Avg 2nd Serve Speed (column 31)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', 'Serve Speed Trends Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Speed (mph)', minValue: 0})
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3}, // 1st Serve (red)
        1: {color: '#FBBC04', lineWidth: 3}  // 2nd Serve (yellow)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Serve Speed chart: ${error.message}`);
  }
}

/**
 * Create Shot Speed Comparison chart (NEW!)
 */
function createShotSpeedChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 32, lastDataRow, 1)) // Avg Forehand Speed (column 32)
      .addRange(dataSheet.getRange(1, 33, lastDataRow, 1)) // Avg Backhand Speed (column 33)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', 'Forehand vs Backhand Speed Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Speed (mph)', minValue: 0})
      .setOption('series', {
        0: {color: '#4285F4', lineWidth: 3}, // Forehand (blue)
        1: {color: '#9C27B0', lineWidth: 3}  // Backhand (purple)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Shot Speed chart: ${error.message}`);
  }
}

/**
 * 🔧 UTILITY FUNCTIONS
 */

/**
 * Manual trigger to process all files (including already processed ones)
 */
function reprocessAllFiles() {
  try {
    logDiagnostic("SYSTEM", "Starting full reprocessing of all files", "", "SUCCESS");
    
    // Clear existing data
    const dataSheet = getMatchDataSheet();
    const lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      dataSheet.deleteRows(2, lastRow - 1);
    }
    
    // Process all files
    const folder = getSwingVisionFolder();
    const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    
    let processedCount = 0;
    let errorCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      
      try {
        const matchData = extractStatsFromFile(file);
        addMatchData(matchData);
        processedCount++;
      } catch (error) {
        errorCount++;
        logDiagnostic("PROCESSING", "Failed to process file", file.getName(), "ERROR", error.message);
      }
    }
    
    // Update charts
    updateCharts();
    
    const summary = `Reprocessed ${processedCount} file(s), errors: ${errorCount}`;
    logDiagnostic("SYSTEM", summary, "", "SUCCESS");
    console.log(`✅ ${summary}`);
    
    SpreadsheetApp.getUi().alert(`Reprocessing complete!\n\n${summary}`);
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Reprocessing failed", "", "ERROR", error.message);
    console.error(`❌ Error reprocessing files: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/**
 * Create a menu in the spreadsheet UI
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Determine interval description for menu
  let intervalText = "";
  if (CHECK_INTERVAL_MINUTES !== null && CHECK_INTERVAL_MINUTES > 0) {
    intervalText = `Every ${CHECK_INTERVAL_MINUTES}m`;
  } else if (CHECK_INTERVAL_HOURS !== null && CHECK_INTERVAL_HOURS > 0) {
    intervalText = `Every ${CHECK_INTERVAL_HOURS}h`;
  } else {
    intervalText = "Configure First";
  }
  
  // Show if single file mode is active
  const modeText = TARGET_SINGLE_FILE ? ` [🎯 Single File]` : '';
  
  ui.createMenu('🎾 Tennis Stats')
    .addItem('🔍 Check for New Matches Now' + modeText, 'checkForNewMatches')
    .addItem('🔄 Reprocess All Files', 'reprocessAllFiles')
    .addSeparator()
    .addItem('📊 Update Charts', 'updateCharts')
    .addSeparator()
    .addItem('⚙️ Install Automatic Check (' + intervalText + ')', 'installTrigger')
    .addItem('🛑 Uninstall Automatic Check', 'uninstallTriggers')
    .addSeparator()
    .addItem('🗑️ Clear Diagnostic Log', 'clearDiagnosticLog')
    .addToUi();
}

/**
 * Initialize the system
 */
function initialize() {
  try {
    console.log("🎾 Initializing Tennis Stats Analyzer...");
    
    // Create all necessary sheets
    getDiagnosticSheet();
    getMatchDataSheet();
    getChartsSheet();
    
    // Install the trigger
    installTrigger();
    
    // Run initial check
    checkForNewMatches();
    
    // Determine interval for display
    let intervalText = "";
    if (CHECK_INTERVAL_MINUTES !== null && CHECK_INTERVAL_MINUTES > 0) {
      intervalText = `every ${CHECK_INTERVAL_MINUTES} minute(s)`;
    } else if (CHECK_INTERVAL_HOURS !== null && CHECK_INTERVAL_HOURS > 0) {
      intervalText = `every ${CHECK_INTERVAL_HOURS} hour(s)`;
    }
    
    logDiagnostic("SYSTEM", "System initialized successfully", "", "SUCCESS");
    console.log("✅ System initialized successfully!");
    
    SpreadsheetApp.getUi().alert(
      '🎾 Tennis Stats Analyzer Initialized!\n\n' +
      `✅ Sheets created\n` +
      `✅ Automatic check installed (${intervalText})\n` +
      `✅ Initial file check completed\n\n` +
      `The system will now automatically check for new match files.\n` +
      `Check the "Diagnostic" sheet for detailed logs.`
    );
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Initialization failed", "", "ERROR", error.message);
    console.error(`❌ Initialization error: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Initialization Error: ${error.message}`);
  }
}

