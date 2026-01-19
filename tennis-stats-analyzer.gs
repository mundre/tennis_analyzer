// ⚙️ CONFIGURATION SETTINGS - Change these values to customize the system

// Check interval - Choose ONE option below:
// Option 1: Check every N hours (set CHECK_INTERVAL_HOURS, leave CHECK_INTERVAL_MINUTES as null)
//const CHECK_INTERVAL_HOURS = 1; // Options: 1, 2, 3, 4, 6, 8, 12, or 24 hours
//const CHECK_INTERVAL_MINUTES = null; // Set to null when using hours

// Option 2: Check every N minutes (set CHECK_INTERVAL_MINUTES, leave CHECK_INTERVAL_HOURS as null)
// const CHECK_INTERVAL_HOURS = null; // Set to null when using minutes
 const CHECK_INTERVAL_MINUTES = 1; // Options: 1, 5, 10, 15, or 30 minutes

// 🎯 Single File Mode - Process only ONE specific file
// Set to null to process all files in folder (normal mode)
// Set to filename (with extension) to process only that file
// const TARGET_SINGLE_FILE = null; // Example: "SwingVision-match-2025-11-09 at 15.59.30.xlsx"
 const TARGET_SINGLE_FILE = "SwingVision-match-2025-12-03 at 22.59.47.xlsx"; // Uncomment to use

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
      // Speed Statistics (from Shots sheet)
      "Avg 1st Serve Speed (mph)",
      "Avg 2nd Serve Speed (mph)",
      "Avg Forehand Speed (mph)",
      "Avg Backhand Speed (mph)",
      // Serve Spin Distribution
      "Serve Flat Count",
      "Serve Kick Count",
      "Serve Slice Count",
      // 🆕 Serve Error Spin Distribution
      "Serve Error Flat",
      "Serve Error Kick",
      "Serve Error Slice",
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
      "Is Practice Match",
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
 * Checks both FileID and FileName to prevent duplicates
 */
function isFileProcessed(fileId, fileName) {
  const dataSheet = getMatchDataSheet();
  const lastRow = dataSheet.getLastRow();
  
  if (lastRow <= 1) {
    return false; // No data yet
  }
  
  // Get only the necessary columns (FileID and FileName)
  const fileIdColumn = dataSheet.getRange(2, 57, lastRow - 1, 1).getValues(); // Column 57: File ID
  const fileNameColumn = dataSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column 2: File Name
  
  // Check if file ID or file name already exists
  for (let i = 0; i < fileIdColumn.length; i++) {
    const rowFileId = fileIdColumn[i][0];
    const rowFileName = fileNameColumn[i][0];
    
    if (rowFileId === fileId || rowFileName === fileName) {
      console.log(`File already processed: ${fileName} (ID: ${fileId})`);
      logDiagnostic("DUPLICATE", `Skipping already processed file: ${fileName}`, fileName, "WARNING");
      return true;
    }
  }
  
  console.log(`File NOT yet processed: ${fileName}`);
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
    
    // Convert XLSX to Google Sheets if needed
    let spreadsheetId = fileId;
    const mimeType = file.getMimeType();
    
    if (mimeType === MimeType.MICROSOFT_EXCEL || mimeType === MimeType.MICROSOFT_EXCEL_LEGACY) {
      // File is XLSX - need to convert to Google Sheets
      logDiagnostic("PROCESSING", `Converting XLSX to Google Sheets`, fileName, "SUCCESS");
      
      const blob = file.getBlob();
      const folder = file.getParents().next(); // Get parent folder
      const resource = {
        title: fileName.replace('.xlsx', '') + ' (Google Sheets)',
        mimeType: MimeType.GOOGLE_SHEETS
      };
      
      // Convert XLSX to Google Sheets using Drive API
      const convertedFile = Drive.Files.insert(resource, blob, {
        convert: true
      });
      
      spreadsheetId = convertedFile.id;
      logDiagnostic("PROCESSING", `Converted to Google Sheets (ID: ${spreadsheetId})`, fileName, "SUCCESS");
    }
    
    // Open the spreadsheet file
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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
    
    // Cleanup: Delete temporary converted Google Sheets file if we created one
    if (spreadsheetId !== fileId) {
      try {
        Drive.Files.remove(spreadsheetId);
        logDiagnostic("PROCESSING", "Deleted temporary converted file", fileName, "SUCCESS");
      } catch (cleanupError) {
        logDiagnostic("PROCESSING", "Failed to delete temporary file (non-critical)", fileName, "WARNING", cleanupError.message);
      }
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
      pointsWonPercent: 0
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
      } else if (label === "double faults") {
        stats.doubleFaults = sumHostColumns(row);
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
    
    // Determine if this is a practice match (all stats are 0)
    const isPracticeMatch = (
      (stats.totalPointsWon || 0) === 0 &&
      (stats.totalPoints || 0) === 0 &&
      (stats.winners || 0) === 0 &&
      (stats.unforcedErrors || 0) === 0 &&
      (stats.firstServes || 0) === 0 &&
      (stats.aces || 0) === 0
    );
    stats.isPracticeMatch = isPracticeMatch;
    
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
      // 🆕 Double faults (from Shots sheet)
      doubleFaults: 0,
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
      backhandErrorsOutSlice: 0,
      // 🆕 Serve Error Spin Distribution (YOUR serve errors: Stroke=serve, Result=Out)
      serveErrorFlat: 0,
      serveErrorKick: 0,
      serveErrorSlice: 0
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
      
      // ⭐ CRITICAL FILTER 1: Skip shots where Type=none
      if (type === "none") {
        continue;  // Skip this shot completely
      }
      
      // ⭐ CRITICAL FILTER 2: Only process HOST (YOUR) shots - skip opponent shots!
      // Skip if: empty player, explicitly "Opponent", or doesn't match your player name
      if (!player || player === "Opponent" || (playerName && player !== playerName)) {
        continue;  // Skip this shot - it's not yours!
      }
      
      // Process serves - Use TYPE column to distinguish 1st vs 2nd serve
      if (stroke.includes("serve")) {
        // Determine if 1st or 2nd serve based on TYPE column
        const isFirstServe = type.includes("first_serve") || type.includes("first serve");
        const isSecondServe = type.includes("second_serve") || type.includes("second serve");
        
        // Collect serve speeds based on type
        if (!isNaN(speed) && speed > 0) {
          if (isFirstServe) {
            serve1stSpeeds.push(speed);
          } else if (isSecondServe) {
            serve2ndSpeeds.push(speed);
          }
        }
        
        // Count serve spin types (for successful serves)
        if (result === "In") {
          if (spin.includes("flat")) {
            shotStats.serveFlat++;
          } else if (spin.includes("kick") || spin.includes("topspin")) {
            shotStats.serveKick++;
          } else if (spin.includes("slice")) {
            shotStats.serveSlice++;
          }
        }
        
        // 🆕 Track serve ERROR spin composition (Stroke=serve, Result=Out)
        const isServeError = result === "Out" || result === "Net" || result === "Long" || result === "Wide";
        if (isServeError) {
          if (spin.includes("flat")) {
            shotStats.serveErrorFlat++;
          } else if (spin.includes("kick") || spin.includes("topspin")) {
            shotStats.serveErrorKick++;
          } else if (spin.includes("slice")) {
            shotStats.serveErrorSlice++;
          }
        }
        
        // 🆕 Count DOUBLE FAULTS: Type=second_serve AND Result=Out or Net
        const isDoubleFault = isSecondServe && (result === "Out" || result === "Net");
        if (isDoubleFault) {
          shotStats.doubleFaults++;
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
      // Speed Statistics
      matchData.avg1stServeSpeed || 0,
      matchData.avg2ndServeSpeed || 0,
      matchData.avgForehandSpeed || 0,
      matchData.avgBackhandSpeed || 0,
      // Serve Spin Distribution
      matchData.serveFlat || 0,
      matchData.serveKick || 0,
      matchData.serveSlice || 0,
      // Serve Error Spin Distribution
      matchData.serveErrorFlat || 0,
      matchData.serveErrorKick || 0,
      matchData.serveErrorSlice || 0,
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
      matchData.isPracticeMatch || false,  // Flag for practice/warmup matches
      new Date()
    ];
    
    dataSheet.appendRow(rowData);
    
    // Sort by match date (column 1) - CHRONOLOGICAL ORDER
    const lastRow = dataSheet.getLastRow();
    if (lastRow > 2) {
      // Get all data rows
      const dataRange = dataSheet.getRange(2, 1, lastRow - 1, rowData.length);
      const allData = dataRange.getValues();
      
      // Sort chronologically by date (column 0 in the array, column 1 in sheet)
      allData.sort((a, b) => {
        const dateA = a[0] instanceof Date ? a[0] : new Date(a[0]);
        const dateB = b[0] instanceof Date ? b[0] : new Date(b[0]);
        return dateA.getTime() - dateB.getTime(); // Ascending (oldest first)
      });
      
      // Write sorted data back
      dataRange.setValues(allData);
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
        if (isFileProcessed(fileId, fileName)) {
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
        if (isFileProcessed(fileId, fileName)) {
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
    
    // Create multiple charts - Organized by category
    let chartRow = 3;
    
    // ============ MATCH PERFORMANCE ============
    // 1. Winners vs Unforced Errors Over Time
    createWinnersUEChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 2. Winners/UE Ratio Trend
    createWinnersUERatioChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 3. Groundstroke Winners (FH vs BH)
    createGroundstrokeWinnersChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 4. Service Winners
    createServiceWinnersChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 5. Points Won Analysis
    createPointsWonChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // ============ SERVE PERFORMANCE ============
    // 6. First Serve % and Points Won %
    createServeStatsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 7. Second Serve Points Won % Trend
    createSecondServePointsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 8. Aces & Double Faults
    createAcesDoubleFaultsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 9. Serve Speed Trends
    createServeSpeedChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 10. Serve Spin Distribution (%)
    createServeSpinTrendsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 11. Serve Error Spin Distribution (%)
    createServeErrorSpinChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // ============ BREAK POINTS ============
    // 12. Break Point Conversion
    createBreakPointChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // ============ UNFORCED ERRORS ============
    // 13. Unforced Errors Breakdown (FH/BH) - Line Chart
    createUEBreakdownChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 14. Unforced Errors by Location (Net vs Out)
    createUnforcedErrorLocationChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 15. Error Location Totals (FH/BH Net/Out)
    createErrorLocationTotalsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 16. Error Spin Composition (FH Net)
    createFHNetSpinChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 17. Error Spin Composition (FH Out)
    createFHOutSpinChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 18. Error Spin Composition (BH Net)
    createBHNetSpinChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 19. Error Spin Composition (BH Out)
    createBHOutSpinChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // ============ SHOT ANALYSIS ============
    // 20. Shot Speed Comparison
    createShotSpeedChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 21. Forehand Spin Distribution
    createForehandSpinTrendsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 22. Backhand Spin Distribution
    createBackhandSpinTrendsChart(dataSheet, chartsSheet, chartRow, lastRow);
    chartRow += 25;
    
    // 23. Match Results Timeline
    createResultsChart(dataSheet, chartsSheet, chartRow, lastRow);
    
    logDiagnostic("CHARTS", "Successfully updated all charts", "", "SUCCESS");
    console.log("✅ Charts updated successfully");
    
  } catch (error) {
    logDiagnostic("CHARTS", "Failed to update charts", "", "ERROR", error.message);
    console.error(`❌ Error updating charts: ${error.message}`);
  }
}

/**
 * Helper function to get filtered data (exclude practice matches)
 * Returns filtered data array with specified columns, sorted by date
 */
function getFilteredRealMatchData(dataSheet, lastDataRow, columns) {
  const isPracticeColumn = 58; // "Is Practice Match" column
  
  // Read ALL columns including isPracticeMatch column
  const allData = dataSheet.getRange(1, 1, lastDataRow, Math.max(isPracticeColumn, ...columns)).getValues();
  
  // Header row
  const filteredData = [columns.map(col => allData[0][col - 1])];
  
  // Collect non-practice match rows with their original data
  const realMatches = [];
  for (let i = 1; i < allData.length; i++) {
    const isPractice = allData[i][isPracticeColumn - 1];
    
    if (!isPractice) { // If NOT practice match, include it
      const date = allData[i][0];
      realMatches.push({
        date: date,
        dateTime: date instanceof Date ? date.getTime() : new Date(date).getTime(),
        row: columns.map(col => allData[i][col - 1])
      });
    }
  }
  
  // Sort by date chronologically using timestamp for accuracy
  realMatches.sort((a, b) => a.dateTime - b.dateTime);
  
  // Add sorted rows to filtered data
  for (const match of realMatches) {
    filteredData.push(match.row);
  }
  
  console.log(`Filtered ${realMatches.length} real matches from ${allData.length - 1} total matches`);
  
  return filteredData;
}

/**
 * Create Winners vs Unforced Errors chart (REAL MATCHES ONLY)
 */
function createWinnersUEChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 14, 18]);
    
    // Write filtered data to temporary range on charts sheet
    const tempStartRow = lastDataRow + 50;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 3);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Total Winners
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // Total Unforced Errors
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Winners vs Unforced Errors Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#34A853', lineWidth: 3, labelInLegend: 'Winners'},
        1: {color: '#EA4335', lineWidth: 3, labelInLegend: 'Unforced Errors'}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Winners/UE chart: ${error.message}`);
  }
}

/**
 * Create Serve Statistics chart (REAL MATCHES ONLY)
 */
function createServeStatsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 6, 7, 24]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 80;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 4);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // First Serve %
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // First Serve Points Won %
      .addRange(chartsSheet.getRange(tempStartRow, 4, filteredData.length, 1)) // Points Won %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎯 Serve Statistics Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#4285F4', lineWidth: 3, labelInLegend: '1st Serve %'},
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: '1st Serve Points Won %'},
        2: {color: '#34A853', lineWidth: 3, labelInLegend: 'Total Points Won %'}
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
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('colors', ['#34A853'])
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Results chart: ${error.message}`);
  }
}

/**
 * Create Serve Speed Trends chart (ALL MATCHES - includes practice)
 */
function createServeSpeedChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 27, lastDataRow, 1)) // Avg 1st Serve Speed (column 27)
      .addRange(dataSheet.getRange(1, 28, lastDataRow, 1)) // Avg 2nd Serve Speed (column 28)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '⚡ Serve Speed Trends Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Speed (mph)', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: '1st Serve Speed'}, // 1st Serve (red)
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: '2nd Serve Speed'}  // 2nd Serve (yellow)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Serve Speed chart: ${error.message}`);
  }
}

/**
 * Create Shot Speed Comparison chart (ALL MATCHES - includes practice)
 */
function createShotSpeedChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 29, lastDataRow, 1)) // Avg Forehand Speed (column 29)
      .addRange(dataSheet.getRange(1, 30, lastDataRow, 1)) // Avg Backhand Speed (column 30)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '⚡ Forehand vs Backhand Speed Over Time')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Speed (mph)', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#4285F4', lineWidth: 3, labelInLegend: 'Forehand Speed'}, // Forehand (blue)
        1: {color: '#9C27B0', lineWidth: 3, labelInLegend: 'Backhand Speed'}  // Backhand (purple)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Shot Speed chart: ${error.message}`);
  }
}

/**
 * Create Unforced Errors by Location chart (ALL MATCHES - includes practice)
 */
function createUnforcedErrorLocationChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 40, lastDataRow, 1)) // FH Errors Net
      .addRange(dataSheet.getRange(1, 41, lastDataRow, 1)) // FH Errors Out
      .addRange(dataSheet.getRange(1, 42, lastDataRow, 1)) // BH Errors Net
      .addRange(dataSheet.getRange(1, 43, lastDataRow, 1)) // BH Errors Out
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎯 Unforced Errors: Net vs Out (Forehand & Backhand)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('isStacked', false)
      .setOption('dataLabelsPlacement', 'outsideEnd')
      .setOption('series', {
        0: {color: '#EA4335', targetAxisIndex: 0, labelInLegend: 'FH Net'}, // FH Net (red)
        1: {color: '#FBBC04', targetAxisIndex: 0, labelInLegend: 'FH Out'}, // FH Out (yellow)
        2: {color: '#4285F4', targetAxisIndex: 0, labelInLegend: 'BH Net'}, // BH Net (blue)
        3: {color: '#34A853', targetAxisIndex: 0, labelInLegend: 'BH Out'}  // BH Out (green)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Unforced Error Location chart: ${error.message}`);
  }
}

/**
 * Create Error Location Totals Chart (FH/BH Net/Out)
 */
function createErrorLocationTotalsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 40, lastDataRow, 1)) // FH Errors Net
      .addRange(dataSheet.getRange(1, 41, lastDataRow, 1)) // FH Errors Out
      .addRange(dataSheet.getRange(1, 42, lastDataRow, 1)) // BH Errors Net
      .addRange(dataSheet.getRange(1, 43, lastDataRow, 1)) // BH Errors Out
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '❌ Error Location Totals (FH/BH Net/Out)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)  // Smooth lines
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: 'FH Net'},   // Red
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: 'FH Out'},   // Yellow
        2: {color: '#4285F4', lineWidth: 3, labelInLegend: 'BH Net'},   // Blue
        3: {color: '#34A853', lineWidth: 3, labelInLegend: 'BH Out'}    // Green
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Error Location Totals chart: ${error.message}`);
  }
}

/**
 * Create FH Net Spin Composition Chart
 */
function createFHNetSpinChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 44, lastDataRow, 1)) // FH Net Topspin
      .addRange(dataSheet.getRange(1, 45, lastDataRow, 1)) // FH Net Flat
      .addRange(dataSheet.getRange(1, 46, lastDataRow, 1)) // FH Net Slice
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 FH Net Errors - Spin Composition')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#D32F2F', lineWidth: 3, labelInLegend: 'Topspin'},  // Dark Red
        1: {color: '#FFA000', lineWidth: 3, labelInLegend: 'Flat'},     // Orange
        2: {color: '#1976D2', lineWidth: 3, labelInLegend: 'Slice'}     // Dark Blue
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating FH Net Spin chart: ${error.message}`);
  }
}

/**
 * Create FH Out Spin Composition Chart
 */
function createFHOutSpinChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 47, lastDataRow, 1)) // FH Out Topspin
      .addRange(dataSheet.getRange(1, 48, lastDataRow, 1)) // FH Out Flat
      .addRange(dataSheet.getRange(1, 49, lastDataRow, 1)) // FH Out Slice
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 FH Out Errors - Spin Composition')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#C2185B', lineWidth: 3, labelInLegend: 'Topspin'},  // Pink
        1: {color: '#00897B', lineWidth: 3, labelInLegend: 'Flat'},     // Teal
        2: {color: '#5E35B1', lineWidth: 3, labelInLegend: 'Slice'}     // Purple
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating FH Out Spin chart: ${error.message}`);
  }
}

/**
 * Create BH Net Spin Composition Chart
 */
function createBHNetSpinChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 50, lastDataRow, 1)) // BH Net Topspin
      .addRange(dataSheet.getRange(1, 51, lastDataRow, 1)) // BH Net Flat
      .addRange(dataSheet.getRange(1, 52, lastDataRow, 1)) // BH Net Slice
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 BH Net Errors - Spin Composition')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#F44336', lineWidth: 3, labelInLegend: 'Topspin'},  // Red
        1: {color: '#FF9800', lineWidth: 3, labelInLegend: 'Flat'},     // Orange
        2: {color: '#2196F3', lineWidth: 3, labelInLegend: 'Slice'}     // Blue
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating BH Net Spin chart: ${error.message}`);
  }
}

/**
 * Create BH Out Spin Composition Chart
 */
function createBHOutSpinChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 53, lastDataRow, 1)) // BH Out Topspin
      .addRange(dataSheet.getRange(1, 54, lastDataRow, 1)) // BH Out Flat
      .addRange(dataSheet.getRange(1, 55, lastDataRow, 1)) // BH Out Slice
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 BH Out Errors - Spin Composition')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#9C27B0', lineWidth: 3, labelInLegend: 'Topspin'},  // Purple
        1: {color: '#4CAF50', lineWidth: 3, labelInLegend: 'Flat'},     // Green
        2: {color: '#00BCD4', lineWidth: 3, labelInLegend: 'Slice'}     // Cyan
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating BH Out Spin chart: ${error.message}`);
  }
}

/**
 * Create Serve Error Spin Distribution chart
 */
function createServeErrorSpinChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Calculate percentages for each row dynamically
    const data = dataSheet.getRange(1, 1, lastDataRow, 33).getValues();
    const percentageData = [['Match Date', 'Flat %', 'Kick %', 'Slice %']];
    
    // Collect all matches with their data - use filename as key to avoid duplicates
    const matchesMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const matchDate = data[i][0];
      const fileName = String(data[i][1] || ''); // Column 2 = File Name
      const flat = data[i][30] || 0;  // Column 31
      const kick = data[i][31] || 0;  // Column 32
      const slice = data[i][32] || 0; // Column 33
      const total = flat + kick + slice;
      
      // Skip if no filename (invalid row)
      if (!fileName) continue;
      
      // Use filename as key to prevent duplicates
      const matchData = {
        date: matchDate,
        dateTime: matchDate instanceof Date ? matchDate.getTime() : new Date(matchDate).getTime(),
        flat: total > 0 ? (flat / total) * 100 : 0,
        kick: total > 0 ? (kick / total) * 100 : 0,
        slice: total > 0 ? (slice / total) * 100 : 0
      };
      
      // Only keep the first occurrence of each filename
      if (!matchesMap.has(fileName)) {
        matchesMap.set(fileName, matchData);
      }
    }
    
    // Convert map to array and sort by date
    const matches = Array.from(matchesMap.values());
    matches.sort((a, b) => a.dateTime - b.dateTime);
    
    console.log(`Serve Error Spin chart: Processing ${matches.length} unique matches from ${data.length - 1} total rows`);
    
    // Build sorted data array
    for (const match of matches) {
      percentageData.push([match.date, match.flat, match.kick, match.slice]);
    }
    
    // Write percentage data to temporary range
    const tempStartRow = lastDataRow + 5;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, percentageData.length, 4);
    tempRange.setValues(percentageData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, percentageData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, percentageData.length, 1)) // Flat %
      .addRange(chartsSheet.getRange(tempStartRow, 3, percentageData.length, 1)) // Kick %
      .addRange(chartsSheet.getRange(tempStartRow, 4, percentageData.length, 1)) // Slice %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Serve Error Spin Distribution (%)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: 'Flat %'},   // Flat (red)
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: 'Kick %'},   // Kick (yellow)
        2: {color: '#4285F4', lineWidth: 3, labelInLegend: 'Slice %'}   // Slice (blue)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Serve Error Spin chart: ${error.message}`);
  }
}

/**
 * Create Aces & Double Faults chart
 */
function createAcesDoubleFaultsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataSheet.getRange(1, 1, lastDataRow, 1)) // Match Date
      .addRange(dataSheet.getRange(1, 9, lastDataRow, 1)) // Aces (column 9)
      .addRange(dataSheet.getRange(1, 10, lastDataRow, 1)) // Double Faults (column 10)
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Aces vs Double Faults')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#34A853', labelInLegend: 'Aces'}, // Aces (green)
        1: {color: '#EA4335', labelInLegend: 'Double Faults'}  // Double Faults (red)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Aces/Double Faults chart: ${error.message}`);
  }
}

/**
 * Create Break Point Conversion chart
 */
function createBreakPointChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 11, 12, 13]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 110;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 4);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Break Points Won
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // Break Points Total
      .addRange(chartsSheet.getRange(tempStartRow, 4, filteredData.length, 1)) // Break Point Conversion %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎯 Break Point Performance')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxes', {
        0: {title: 'Count'},
        1: {title: 'Percentage (%)'}
      })
      .setOption('series', {
        0: {color: '#34A853', targetAxisIndex: 0, labelInLegend: 'BP Won'}, // BP Won (green)
        1: {color: '#FBBC04', targetAxisIndex: 0, labelInLegend: 'BP Total'}, // BP Total (yellow)
        2: {color: '#4285F4', targetAxisIndex: 1, lineWidth: 3, labelInLegend: 'BP Conversion %'} // BP % (blue)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Break Point chart: ${error.message}`);
  }
}

/**
 * Create Winners Breakdown chart (Service/FH/BH)
 */
function createGroundstrokeWinnersChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only) - FH and BH winners
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 16, 17]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 140;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 3);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Forehand Winners
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // Backhand Winners
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🏆 Groundstroke Winners (FH vs BH)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#34A853', lineWidth: 3, labelInLegend: 'Forehand Winners'}, // Forehand (green)
        1: {color: '#4285F4', lineWidth: 3, labelInLegend: 'Backhand Winners'}  // Backhand (blue)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Groundstroke Winners chart: ${error.message}`);
  }
}

/**
 * Create Service Winners chart
 */
function createServiceWinnersChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only) - Service winners
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 15]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 170;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 2);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Service Winners
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🏆 Service Winners Trend')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#FBBC04', lineWidth: 3, labelInLegend: 'Service Winners'} // Service (yellow)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Service Winners chart: ${error.message}`);
  }
}

/**
 * Create Unforced Errors Breakdown chart (FH/BH)
 */
function createUEBreakdownChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 19, 20]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 200;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 3);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Forehand UE
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // Backhand UE
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '❌ Unforced Errors Breakdown (FH vs BH)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Count', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: 'Forehand UE'}, // Forehand (red)
        1: {color: '#9C27B0', lineWidth: 3, labelInLegend: 'Backhand UE'}  // Backhand (purple)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating UE Breakdown chart: ${error.message}`);
  }
}

/**
 * Create Points Won Analysis chart
 */
function createPointsWonChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 22, 23, 24]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 230;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 4);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Total Points Won
      .addRange(chartsSheet.getRange(tempStartRow, 3, filteredData.length, 1)) // Total Points
      .addRange(chartsSheet.getRange(tempStartRow, 4, filteredData.length, 1)) // Points Won %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '📊 Points Won Analysis')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxes', {
        0: {title: 'Count'},
        1: {title: 'Percentage (%)'}
      })
      .setOption('series', {
        0: {color: '#34A853', targetAxisIndex: 0, labelInLegend: 'Points Won'}, // Points Won (green)
        1: {color: '#FBBC04', targetAxisIndex: 0, labelInLegend: 'Total Points'}, // Total Points (yellow)
        2: {color: '#4285F4', targetAxisIndex: 1, lineWidth: 3, labelInLegend: 'Points Won %'} // Points Won % (blue)
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Points Won chart: ${error.message}`);
  }
}

/**
 * Create Serve Spin Trends (Per Match)
 */
function createServeSpinTrendsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Calculate percentages for each row dynamically
    const data = dataSheet.getRange(1, 1, lastDataRow, 29).getValues();
    const percentageData = [['Match Date', 'Flat %', 'Kick %', 'Slice %']];
    
    // Collect all matches with their data - use filename as key to avoid duplicates
    const matchesMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const matchDate = data[i][0];
      const fileName = String(data[i][1] || ''); // Column 2 = File Name
      const flat = data[i][26] || 0;  // Column 27
      const kick = data[i][27] || 0;  // Column 28
      const slice = data[i][28] || 0; // Column 29
      const total = flat + kick + slice;
      
      // Skip if no filename (invalid row)
      if (!fileName) continue;
      
      // Use filename as key to prevent duplicates
      const matchData = {
        date: matchDate,
        dateTime: matchDate instanceof Date ? matchDate.getTime() : new Date(matchDate).getTime(),
        flat: total > 0 ? (flat / total) * 100 : 0,
        kick: total > 0 ? (kick / total) * 100 : 0,
        slice: total > 0 ? (slice / total) * 100 : 0
      };
      
      // Only keep the first occurrence of each filename
      if (!matchesMap.has(fileName)) {
        matchesMap.set(fileName, matchData);
      }
    }
    
    // Convert map to array and sort by date
    const matches = Array.from(matchesMap.values());
    matches.sort((a, b) => a.dateTime - b.dateTime);
    
    console.log(`Serve Spin chart: Processing ${matches.length} unique matches from ${data.length - 1} total rows`);
    
    // Build sorted data array
    for (const match of matches) {
      percentageData.push([match.date, match.flat, match.kick, match.slice]);
    }
    
    // Write percentage data to temporary range
    const tempStartRow = lastDataRow + 10;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, percentageData.length, 4);
    tempRange.setValues(percentageData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, percentageData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, percentageData.length, 1)) // Flat %
      .addRange(chartsSheet.getRange(tempStartRow, 3, percentageData.length, 1)) // Kick %
      .addRange(chartsSheet.getRange(tempStartRow, 4, percentageData.length, 1)) // Slice %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Serve Spin Distribution (%)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: 'Flat %'},
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: 'Kick %'},
        2: {color: '#4285F4', lineWidth: 3, labelInLegend: 'Slice %'}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Serve Spin Trends chart: ${error.message}`);
  }
}

/**
 * Create Forehand Spin Trends (Per Match)
 */
function createForehandSpinTrendsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Calculate percentages for each row dynamically
    const data = dataSheet.getRange(1, 1, lastDataRow, 36).getValues();
    const percentageData = [['Match Date', 'Topspin %', 'Flat %', 'Slice %']];
    
    // Collect all matches with their data - use filename as key to avoid duplicates
    const matchesMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const matchDate = data[i][0];
      const fileName = String(data[i][1] || ''); // Column 2 = File Name
      const topspin = data[i][33] || 0;  // Column 34
      const flat = data[i][34] || 0;     // Column 35
      const slice = data[i][35] || 0;    // Column 36
      const total = topspin + flat + slice;
      
      // Skip if no filename (invalid row)
      if (!fileName) continue;
      
      // Use filename as key to prevent duplicates
      const matchData = {
        date: matchDate,
        dateTime: matchDate instanceof Date ? matchDate.getTime() : new Date(matchDate).getTime(),
        topspin: total > 0 ? (topspin / total) * 100 : 0,
        flat: total > 0 ? (flat / total) * 100 : 0,
        slice: total > 0 ? (slice / total) * 100 : 0
      };
      
      // Only keep the first occurrence of each filename
      if (!matchesMap.has(fileName)) {
        matchesMap.set(fileName, matchData);
      }
    }
    
    // Convert map to array and sort by date
    const matches = Array.from(matchesMap.values());
    matches.sort((a, b) => a.dateTime - b.dateTime);
    
    console.log(`Forehand Spin chart: Processing ${matches.length} unique matches from ${data.length - 1} total rows`);
    
    // Build sorted data array
    for (const match of matches) {
      percentageData.push([match.date, match.topspin, match.flat, match.slice]);
    }
    
    // Write percentage data to temporary range
    const tempStartRow = lastDataRow + 15;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, percentageData.length, 4);
    tempRange.setValues(percentageData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, percentageData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, percentageData.length, 1)) // Topspin %
      .addRange(chartsSheet.getRange(tempStartRow, 3, percentageData.length, 1)) // Flat %
      .addRange(chartsSheet.getRange(tempStartRow, 4, percentageData.length, 1)) // Slice %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Forehand Spin Distribution (%)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#EA4335', lineWidth: 3, labelInLegend: 'Topspin %'},
        1: {color: '#FBBC04', lineWidth: 3, labelInLegend: 'Flat %'},
        2: {color: '#4285F4', lineWidth: 3, labelInLegend: 'Slice %'}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Forehand Spin Trends chart: ${error.message}`);
  }
}

/**
 * Create Backhand Spin Trends (Per Match)
 */
function createBackhandSpinTrendsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Calculate percentages for each row dynamically
    const data = dataSheet.getRange(1, 1, lastDataRow, 39).getValues();
    const percentageData = [['Match Date', 'Topspin %', 'Flat %', 'Slice %']];
    
    // Collect all matches with their data - use filename as key to avoid duplicates
    const matchesMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const matchDate = data[i][0];
      const fileName = String(data[i][1] || ''); // Column 2 = File Name
      const topspin = data[i][36] || 0;  // Column 37
      const flat = data[i][37] || 0;     // Column 38
      const slice = data[i][38] || 0;    // Column 39
      const total = topspin + flat + slice;
      
      // Skip if no filename (invalid row)
      if (!fileName) continue;
      
      // Use filename as key to prevent duplicates
      const matchData = {
        date: matchDate,
        dateTime: matchDate instanceof Date ? matchDate.getTime() : new Date(matchDate).getTime(),
        topspin: total > 0 ? (topspin / total) * 100 : 0,
        flat: total > 0 ? (flat / total) * 100 : 0,
        slice: total > 0 ? (slice / total) * 100 : 0
      };
      
      // Only keep the first occurrence of each filename
      if (!matchesMap.has(fileName)) {
        matchesMap.set(fileName, matchData);
      }
    }
    
    // Convert map to array and sort by date
    const matches = Array.from(matchesMap.values());
    matches.sort((a, b) => a.dateTime - b.dateTime);
    
    console.log(`Backhand Spin chart: Processing ${matches.length} unique matches from ${data.length - 1} total rows`);
    
    // Build sorted data array
    for (const match of matches) {
      percentageData.push([match.date, match.topspin, match.flat, match.slice]);
    }
    
    // Write percentage data to temporary range
    const tempStartRow = lastDataRow + 20;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, percentageData.length, 4);
    tempRange.setValues(percentageData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, percentageData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, percentageData.length, 1)) // Topspin %
      .addRange(chartsSheet.getRange(tempStartRow, 3, percentageData.length, 1)) // Flat %
      .addRange(chartsSheet.getRange(tempStartRow, 4, percentageData.length, 1)) // Slice %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎾 Backhand Spin Distribution (%)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#FF1744', lineWidth: 3, labelInLegend: 'Topspin %'},  // Bright Red
        1: {color: '#00E676', lineWidth: 3, labelInLegend: 'Flat %'},     // Bright Green
        2: {color: '#2979FF', lineWidth: 3, labelInLegend: 'Slice %'}     // Bright Blue
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Backhand Spin Trends chart: ${error.message}`);
  }
}

/**
 * Create Second Serve Points Won % Trend
 */
function createSecondServePointsChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 8]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 260;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 2);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Second Serve Points Won %
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '🎯 Second Serve Points Won % Trend')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Percentage (%)', minValue: 0, maxValue: 100})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#FF5722', lineWidth: 3, labelInLegend: '2nd Serve Points Won %'}
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Second Serve Points chart: ${error.message}`);
  }
}

/**
 * Create Winners/UE Ratio Trend
 */
function createWinnersUERatioChart(dataSheet, chartsSheet, startRow, lastDataRow) {
  try {
    // Get filtered data (real matches only)
    const filteredData = getFilteredRealMatchData(dataSheet, lastDataRow, [1, 21]);
    
    // Write filtered data to temporary range
    const tempStartRow = lastDataRow + 290;
    const tempRange = chartsSheet.getRange(tempStartRow, 1, filteredData.length, 2);
    tempRange.setValues(filteredData);
    
    const chart = chartsSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(chartsSheet.getRange(tempStartRow, 1, filteredData.length, 1)) // Match Date
      .addRange(chartsSheet.getRange(tempStartRow, 2, filteredData.length, 1)) // Winners/UE Ratio
      .setPosition(startRow, 1, 0, 0)
      .setOption('title', '📊 Winners/UE Ratio Trend (Higher is Better)')
      .setOption('width', 800)
      .setOption('height', 400)
      .setOption('legend', {position: 'bottom'})
      .setOption('hAxis', {title: 'Match Date', slantedText: true, slantedTextAngle: 45})
      .setOption('vAxis', {title: 'Ratio', minValue: 0})
      .setOption('curveType', 'function')
      .setOption('pointSize', 5)
      .setOption('series', {
        0: {color: '#34A853', lineWidth: 3, labelInLegend: 'Winners/UE Ratio'}
      })
      .setOption('trendlines', {
        0: {
          type: 'linear',
          color: '#EA4335',
          lineWidth: 2,
          opacity: 0.5,
          showR2: true,
          visibleInLegend: true
        }
      })
      .build();
    
    chartsSheet.insertChart(chart);
    
  } catch (error) {
    console.error(`Error creating Winners/UE Ratio chart: ${error.message}`);
  }
}

/**
 * 🔧 UTILITY FUNCTIONS
 */

/**
 * Delete all contents from all sheets (fresh start)
 * Completely deletes and recreates all sheets with fresh headers
 * Note: Google Sheets requires at least one sheet to exist at all times
 */
function clearAllSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Strategy: Recreate each sheet one at a time to maintain at least one sheet
    console.log("🔄 Starting fresh sheet recreation...");
    
    // 1. Delete and recreate Match Data sheet
    const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
    if (dataSheet) {
      spreadsheet.deleteSheet(dataSheet);
      console.log("✅ Deleted Match Data sheet");
    }
    getMatchDataSheet(); // Recreate with headers
    console.log("✅ Recreated Match Data sheet with headers");
    
    // 2. Delete and recreate Diagnostic sheet
    const diagSheet = spreadsheet.getSheetByName("Diagnostic");
    if (diagSheet) {
      spreadsheet.deleteSheet(diagSheet);
      console.log("✅ Deleted Diagnostic sheet");
    }
    getDiagnosticSheet(); // Recreate with headers
    console.log("✅ Recreated Diagnostic sheet with headers");
    
    // 3. Delete and recreate Performance Charts sheet
    const chartsSheet = spreadsheet.getSheetByName(CHARTS_SHEET_NAME);
    if (chartsSheet) {
      spreadsheet.deleteSheet(chartsSheet);
      console.log("✅ Deleted Performance Charts sheet");
    }
    getChartsSheet(); // Recreate empty
    console.log("✅ Recreated Performance Charts sheet");
    
    // 4. Delete any default sheets (like "Sheet1") if they exist
    const allSheets = spreadsheet.getSheets();
    for (let sheet of allSheets) {
      const sheetName = sheet.getName();
      // Only delete if it's not one of our three main sheets
      if (sheetName !== DATA_SHEET_NAME && 
          sheetName !== CHARTS_SHEET_NAME && 
          sheetName !== "Diagnostic") {
        try {
          spreadsheet.deleteSheet(sheet);
          console.log(`✅ Deleted extra sheet: ${sheetName}`);
        } catch (e) {
          // Can't delete if it's the last sheet - that's fine
          console.log(`⚠️ Kept sheet: ${sheetName} (last sheet protection)`);
        }
      }
    }
    
    const message = "🗑️ All sheets deleted and recreated with fresh headers! Ready for fresh start.";
    
    // Log to the newly recreated diagnostic sheet
    logDiagnostic("SYSTEM", "All sheets cleared and reinitialized", "", "SUCCESS");
    console.log(`✅ ${message}`);
    
  } catch (error) {
    const errorMsg = `Failed to clear sheets: ${error.message}`;
    console.error(`❌ ${errorMsg}`);
  }
}

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
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Reprocessing failed", "", "ERROR", error.message);
    console.error(`❌ Error reprocessing files: ${error.message}`);
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
    .addItem('🗑️ Clear All Sheets (Fresh Start)', 'clearAllSheets')
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
    console.log(`✅ Sheets created`);
    console.log(`✅ Automatic check installed (${intervalText})`);
    console.log(`✅ Initial file check completed`);
    
  } catch (error) {
    logDiagnostic("SYSTEM", "Initialization failed", "", "ERROR", error.message);
    console.error(`❌ Initialization error: ${error.message}`);
  }
}


