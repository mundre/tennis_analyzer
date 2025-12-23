# 🎾 Tennis Stats Analyzer - Complete Documentation

## 📖 **OVERVIEW**

The Tennis Stats Analyzer is a comprehensive Google Apps Script system for automatically tracking and analyzing tennis performance from SwingVision match data. It monitors your Google Drive, extracts detailed statistics, and generates visual charts to track improvement over time.

### 🎯 **Core Purpose**
- **Automatic monitoring** of SwingVision folder for new match files
- **Comprehensive stats extraction** from both Stats and Shots sheets
- **Visual performance tracking** with automatically generated charts
- **Detailed analytics** including speed, spin, winners, errors, and physical metrics
- **Host-only tracking** - focuses on YOUR performance, not opponent data

---

## ⭐ **RECENT UPDATES (v1.2)**

### 🆕 **NEW: Enhanced Error & Serve Analysis**
- **✅ Fixed duplicate match entries** - Proper file tracking prevents duplicates
- **✅ Accurate speed extraction** - Uses Type column (first_serve/second_serve) from Shots sheet
- **✅ Serve error spin composition** - Track which spin types cause serve errors
- **✅ Removed calories & heart rate** - Cleaner focus on tennis-specific metrics
- **✅ Enhanced UE analysis** - 3 new comprehensive error charts

### 🎾 **Unforced Error Deep Dive** (16 metrics)
- **Location Analysis**: Net vs Out breakdown for FH/BH
- **Spin Analysis**: Topspin/Flat/Slice for each error type
- **Serve Errors**: Spin composition of missed serves
- **Visual Charts**: 3 dedicated error analysis charts

### 📊 **Chart Updates**
Now featuring **22 comprehensive charts** organized by category:

**Match Performance (4 charts)**
1. Winners vs Unforced Errors Over Time
2. Winners/UE Ratio Trend
3. Winners Breakdown (Service/FH/BH)
4. Points Won Analysis

**Serve Performance (6 charts)**
5. First Serve % and Points Won %
6. Second Serve Points Won % Trend
7. Aces & Double Faults
8. Serve Speed Trends (1st vs 2nd)
9. Serve Spin Distribution (%)
10. Serve Error Spin Distribution (%)

**Break Points (1 chart)**
11. Break Point Conversion

**Unforced Errors (7 charts)**
12. UE Breakdown (FH vs BH) - Line Chart
13. UE by Location (Net vs Out)
14. Error Location Totals (FH/BH Net/Out)
15. FH Net Errors - Spin Composition
16. FH Out Errors - Spin Composition
17. BH Net Errors - Spin Composition
18. BH Out Errors - Spin Composition

**Shot Analysis (4 charts)**
19. Shot Speed Comparison (FH vs BH)
20. Forehand Spin Distribution
21. Backhand Spin Distribution
22. Match Results Timeline

### 📈 **Data Expansion**
- **60+ columns** in Match Data sheet
- **19 serve & shot metrics** per match
- **19 unforced error analysis metrics** per match
- **22 organized charts** (Match Performance, Serve, Break Points, UE Analysis, Shot Analysis)
- **Single file mode** for testing/debugging
- **Percentage-based serve spin charts** for better comparison

### ⏱️ **Flexible Check Intervals**
- **Minute-based**: 1, 5, 10, 15, or 30 minutes
- **Hour-based**: 1, 2, 3, 4, 6, 8, 12, or 24 hours
- **Configurable**: Single point of control for interval settings

---

## 🚀 **KEY FEATURES**

### 📊 **1. Comprehensive Stats Extraction**
**From Stats Sheet** (aggregated match statistics):
- Serve statistics (first/second serve %, points won %)
- Break point conversion
- Winners breakdown (service, forehand, backhand)
- Unforced errors breakdown (forehand, backhand)
- Match totals (points, games, distance)

**From Shots Sheet** (shot-by-shot analysis):
- **Speed metrics**: 1st/2nd serve, forehand, backhand (uses Type column)
- **Spin distribution**: Flat/Kick/Slice for serves and groundstrokes
- **🆕 Serve error analysis**: Spin composition of missed serves (Out/Net)
- **🆕 UE deep dive**: Net vs Out breakdown with spin analysis
- **Detailed tactical patterns**: Shot-by-shot insights

### 📈 **2. Automatic Performance Charts**
Twenty-two visual charts organized by category track your improvement:

**🏆 Match Performance (4 charts)**
1. **Winners vs Unforced Errors** - Shot quality trends
2. **Winners/UE Ratio** - Efficiency metric over time
3. **Winners Breakdown** - Service/Forehand/Backhand winners
4. **Points Won Analysis** - Total points and win %

**🎾 Serve Performance (6 charts)**
5. **Serve Statistics** - First serve %, points won %
6. **Second Serve Points Won %** - Effectiveness trend
7. **Aces & Double Faults** - Service extremes
8. **Serve Speed Trends** - 1st vs 2nd serve power
9. **Serve Spin Distribution (%)** - Flat/Kick/Slice usage
10. **Serve Error Spin (%)** - Which serve types miss

**💪 Break Points (1 chart)**
11. **Break Point Conversion** - Clutch performance

**❌ Unforced Errors (7 charts)**
12. **UE Breakdown** - Forehand vs Backhand (line chart)
13. **UE by Location** - Net vs Out patterns
14. **Error Location Totals** - FH/BH Net/Out trends
15. **FH Net Spin** - Topspin/Flat/Slice composition
16. **FH Out Spin** - Topspin/Flat/Slice composition
17. **BH Net Spin** - Topspin/Flat/Slice composition
18. **BH Out Spin** - Topspin/Flat/Slice composition

**🎯 Shot Analysis (4 charts)**
19. **Shot Speed Comparison** - Forehand vs backhand speed
20. **Forehand Spin Distribution** - Shot variety
21. **Backhand Spin Distribution** - Shot variety
22. **Match Results Timeline** - Win/loss patterns

### 🔄 **3. Intelligent Monitoring System**
- **Time-based triggers**: Configurable check frequency
- **Smart processing**: Skips already-processed files automatically
- **Robust error handling**: Continues processing even if one file fails
- **Diagnostic logging**: Comprehensive troubleshooting logs

### 🎯 **4. Player-Focused Analysis**
- **Host-only tracking**: Extracts YOUR stats, ignores opponent
- **Match-by-match trends**: See improvement over time
- **Detailed breakdowns**: Understand strengths and weaknesses
- **Actionable insights**: Identify areas for improvement

---

## 🔧 **SETUP & INSTALLATION**

### ✅ **Pre-Setup Checklist**
Before you begin, make sure you have:
- [ ] Google account
- [ ] SwingVision match files (XLSX format)
- [ ] Files uploaded to Google Drive

### 📋 **Quick Setup (10 Minutes)**

#### **Step 1: Organize Your Files (2 minutes)**
1. Go to [Google Drive](https://drive.google.com)
2. Create a folder named **"SwingVision"** (or use existing)
3. Upload your SwingVision XLSX files to this folder
4. ✨ **Pro Tip**: SwingVision automatically dates files - no renaming needed!

#### **Step 2: Create Tracking Spreadsheet (1 minute)**
1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it **"Tennis Performance Tracker"**
4. Keep this tab open

#### **Step 3: Install the Script (3 minutes)**
1. In your spreadsheet: **Extensions** → **Apps Script**
2. Delete any default code in the editor
3. Open `tennis-stats-analyzer.gs` from this repository
4. **Copy all the code** (Ctrl/Cmd + A, then Ctrl/Cmd + C)
5. **Paste** into Apps Script editor (Ctrl/Cmd + V)
6. Click **save icon** (💾)
7. Name your project: **"Tennis Stats Analyzer"**

#### **Step 4: Enable Drive API (1 minute)**
🔑 **REQUIRED**: The script needs Drive API to convert XLSX files to Google Sheets format.

1. In the Apps Script editor, click on **"Services"** (➕ icon in left sidebar)
2. Find **"Drive API"** in the list
3. Click **"Add"**
4. Leave version as "v2" and identifier as "Drive"
5. Click **"Add"**

> ✅ This allows the script to read your XLSX files and convert them automatically.

#### **Step 5: Authorize the Script (2 minutes)**
1. Select **`initialize`** from the function dropdown
2. Click the **Run button** (▶️)
3. When prompted for permissions:
   - Click **"Review permissions"**
   - Select your Google account
   - Click **"Advanced"** → **"Go to Tennis Stats Analyzer (unsafe)"**
   - Click **"Allow"**

> ⚠️ The "unsafe" warning is normal - this is your own script, not a third-party app. Your data stays private.

#### **Step 6: Wait for Initialization (1 minute)**
1. Script will run and set everything up
2. You'll see a completion dialog
3. Click **"OK"**
4. Close the Apps Script editor tab

#### **Step 7: Verify Installation (1 minute)**
Go back to your spreadsheet. You should see:

**Three new sheets:**
- Match Data (all statistics)
- Performance Charts (visual trends)
- Diagnostic (system logs)

**New menu:**
- 🎾 Tennis Stats (in top menu bar)

## 🎉 **You're Done!**

Your system is now:
- ✅ Automatically checking for new files every hour
- ✅ Processing statistics from each match
- ✅ Creating performance charts
- ✅ Logging all activities

---

## ⚙️ **CONFIGURATION**

### 🎮 **Configuration Settings**

All settings are at the top of `tennis-stats-analyzer.gs`:

```javascript
// Check interval - Choose ONE option:

// Option 1: Hour-based (default)
const CHECK_INTERVAL_HOURS = 1;      // Options: 1, 2, 3, 4, 6, 8, 12, 24
const CHECK_INTERVAL_MINUTES = null;

// Option 2: Minute-based (faster, uses more quota)
// const CHECK_INTERVAL_HOURS = null;
// const CHECK_INTERVAL_MINUTES = 5;    // Options: 1, 5, 10, 15, 30

// 🎯 Single File Mode - Process only ONE specific file
const TARGET_SINGLE_FILE = null;     // null = process all files (normal)
// const TARGET_SINGLE_FILE = "SwingVision-match-2025-11-09 at 15.59.30.xlsx";  // Uncomment to use

const SWINGVISION_FOLDER_NAME = "SwingVision";  // Google Drive folder name
const ENABLE_DIAGNOSTIC_LOGGING = true;          // Enable/disable logging
const STATS_SHEET_NAME = "Stats";                // Stats sheet name in XLSX
const DATA_SHEET_NAME = "Match Data";            // Output sheet name
const CHARTS_SHEET_NAME = "Performance Charts";  // Charts sheet name
```

### ⏱️ **Check Interval Options**

**Hour-Based** (default - recommended):
- 1 hour (default) - Balanced frequency
- 2, 3, 4, 6, 8, 12, 24 hours - Less frequent

**Minute-Based** (faster - for active players):
- 5 minutes ⭐ **RECOMMENDED** - Very fast, quota-safe
- 1 minute - Fastest (⚠️ may hit quota limits)
- 10, 15, 30 minutes - Fast alternatives

**Quota Considerations:**
- Google Apps Script: 90 minutes total runtime per day
- 1-minute checking: ~120 minutes/day ❌ EXCEEDS QUOTA
- 5-minute checking: ~24 minutes/day ✅ SAFE
- Hourly checking: ~2-3 minutes/day ✅ VERY SAFE

**After changing**: Run "🎾 Tennis Stats" → "Uninstall Automatic Check" → "Install Automatic Check"

### 🎯 **Single File Mode**

Process **one specific file** instead of all files in the folder.

**Use Cases:**
- Testing the script with one match
- Reprocessing a single file with errors
- Adding one file without scanning entire folder
- Debugging data extraction issues

**How to Use:**

1. **Set the target filename:**
```javascript
const TARGET_SINGLE_FILE = "SwingVision-match-2025-11-09 at 15.59.30.xlsx";
```

2. **Run the script:**
- Use "🎾 Tennis Stats" → "Check for New Matches"
- Only the specified file will be processed

3. **Return to normal mode:**
```javascript
const TARGET_SINGLE_FILE = null;  // Process all files
```

**Important Notes:**
- ✅ File must exist in your SwingVision folder
- ✅ Filename must match **exactly** (including `.xlsx` extension)
- ⚠️ **Common mistake**: Don't forget `.xlsx` at the end!
  - ❌ Wrong: `"SwingVision-match-2025-12-03 at 22.59.47"`
  - ✅ Correct: `"SwingVision-match-2025-12-03 at 22.59.47.xlsx"`
- ✅ Already-processed files will be skipped (same as normal mode)
- ⚠️ Automatic triggers will respect this setting
- 💡 Useful for testing before enabling auto-check

**Example Workflow:**
```javascript
// 1. Test with one file first
const TARGET_SINGLE_FILE = "SwingVision-match-2025-11-09 at 15.59.30.xlsx";
// Run: Check for New Matches

// 2. Once confirmed working, switch to all files
const TARGET_SINGLE_FILE = null;
// Run: Install Automatic Check
```

### 🗂️ **Folder Configuration**

**Change folder name:**
```javascript
const SWINGVISION_FOLDER_NAME = "Tennis Matches 2024";
```

**Requirements:**
- Exact name match (case-sensitive)
- Must be in "My Drive" (not Shared Drives)
- Folder must exist before running script

**After changing**: Run "🎾 Tennis Stats" → "Reprocess All Files"

### 📝 **Logging Configuration**

**Disable diagnostic logging:**
```javascript
const ENABLE_DIAGNOSTIC_LOGGING = false;
```

**What you lose when disabled:**
- No troubleshooting logs
- Can't track when files were processed
- Harder to debug errors

**Storage impact**: Minimal (last 1,000 entries kept)

---

## 📊 **UNDERSTANDING YOUR DATA**

### 📋 **Match Data Sheet (58 Columns)**

**Basic Info:**
- Match Date, File Name, Opponent, Result, Score

**Serve Statistics (8 columns):**
- First Serve %, First Serve Points Won %, Second Serve Points Won %
- Aces, Double Faults
- 🆕 Avg 1st Serve Speed (mph), Avg 2nd Serve Speed (mph)

**Break Points (3 columns):**
- Break Points Won, Break Points Total, Break Point Conversion %

**Winners & Errors (8 columns):**
- Total Winners, Service Winners, Forehand Winners, Backhand Winners
- Total Unforced Errors, Forehand UE, Backhand UE, Winners/UE Ratio

**Match Totals (5 columns):**
- Total Points Won, Total Points, Points Won %
- Games Won, Games Total

**Physical Stats (3 columns):**
- Calories Burned, Distance Run (mi), Avg Heart Rate (BPM)

**🆕 Speed Statistics (4 columns):**
- Avg 1st Serve Speed (mph), Avg 2nd Serve Speed (mph)
- Avg Forehand Speed (mph), Avg Backhand Speed (mph)

**🆕 Serve Spin Distribution (3 columns):**
- Serve Flat Count, Serve Kick Count, Serve Slice Count

**🆕 Forehand Spin Distribution (3 columns):**
- FH Topspin Count, FH Flat Count, FH Slice Count

**🆕 Backhand Spin Distribution (3 columns):**
- BH Topspin Count, BH Flat Count, BH Slice Count

**🆕 Unforced Error Analysis (16 columns):**

*Error Type Breakdown (4 columns):*
- FH Errors Net - Forehand errors into the net
- FH Errors Out - Forehand errors hit long/wide
- BH Errors Net - Backhand errors into the net
- BH Errors Out - Backhand errors hit long/wide

*Forehand Net Error Spin (3 columns):*
- FH Net Topspin, FH Net Flat, FH Net Slice

*Forehand Out Error Spin (3 columns):*
- FH Out Topspin, FH Out Flat, FH Out Slice

*Backhand Net Error Spin (3 columns):*
- BH Net Topspin, BH Net Flat, BH Net Slice

*Backhand Out Error Spin (3 columns):*
- BH Out Topspin, BH Out Flat, BH Out Slice

**Metadata (2 columns):**
- File ID, Processed Date

### 📈 **Performance Charts Sheet**

**Chart 1: Winners vs Unforced Errors**
- **Green line**: Winners (want trending up ↗️)
- **Red line**: Unforced errors (want trending down ↘️)
- **Goal**: Increasing gap between lines

**Chart 2: Serve Statistics**
- **Blue line**: First serve %
- **Yellow line**: First serve points won %
- **Green line**: Overall points won %
- **Goal**: All lines trending up

**Chart 3: Winners/UE Ratio**
- **Green bars**: Ratio value
- **Interpretation**: 
  - < 1.0 = Too many errors
  - 1.0-1.5 = Solid play
  - > 1.5 = Excellent shot quality
- **Goal**: Bars trending higher

**Chart 4: 🆕 Serve Speed Trends**
- **Red line**: 1st serve speed
- **Yellow line**: 2nd serve speed
- **Typical ranges**:
  - Recreational: 70-90 / 50-70 mph
  - Competitive: 110-130 / 85-100 mph
- **Goal**: Both lines trending up

**Chart 5: 🆕 Shot Speed Comparison**
- **Blue line**: Forehand speed
- **Purple line**: Backhand speed
- **Typical gap**: 5-15 mph (forehand faster)
- **Goals**: 
  - Both trending up (improving power)
  - Gap closing (developing backhand)

### 🔍 **Diagnostic Sheet**

**Color-Coded Logs:**
- 🟢 **Green** = Success
- 🟡 **Yellow** = Warning
- 🔴 **Red** = Error

**Information Tracked:**
- Timestamp of each event
- Event type (PROCESSING, DATA, CHARTS, etc.)
- File names processed
- Error messages (if any)
- Success confirmations

**Auto-Maintenance:**
- Keeps last 1,000 entries
- Oldest automatically deleted
- Can be cleared manually: "🎾 Tennis Stats" → "Clear Diagnostic Log"

---

## 📄 **SWINGVISION DATA FORMAT**

### 🎾 **File Structure**

**SwingVision XLSX files contain 6 sheets:**
1. Settings - Match configuration
2. **Shots** 🆕 - Shot-by-shot data (used for speed/spin)
3. Points - Point-by-point breakdown
4. Games - Game-by-game stats
5. Sets - Set-by-set summary
6. **Stats** - Aggregated statistics (primary data source)

### 📊 **Stats Sheet Format**

**Layout**: Multi-column organized by sets

| Column A | B | C | D | E | ... |
|----------|---|---|---|---|-----|
| Stat Name | Host Set 1 | Guest Set 1 | Host Set 2 | Guest Set 2 | ... |
| 1st Serves | 15 | 12 | 14 | 13 | ... |
| Aces | 2 | 1 | 1 | 0 | ... |

**Key Points:**
- Column A: Stat names
- Columns B, D, F, H, J: **YOUR stats** (Host) ⭐
- Columns C, E, G, I, K: Opponent stats (ignored)
- Script sums YOUR stats across all sets

**Extracted Stats:**
- Serve statistics (serves, serves in, points won)
- Break points (opportunities, converted, saved)
- Winners & errors (service, forehand, backhand)
- Match totals (points, games)
- Physical metrics (calories, distance, heart rate)

### 🎯 **Shots Sheet Format** 🆕

**Layout**: One row per shot

| Column | Data | Example |
|--------|------|---------|
| A | Player | "Host" or "Opponent" |
| B | Shot # | 0, 1, 2, ... |
| C | Type | "none" |
| D | Stroke | "Forehand", "Backhand", "Serve" |
| E | Spin | "Topspin", "Flat", "Slice", "Kick" |
| F | Speed (MPH) | 68.5 |

**Processing:**
- Filters shots by "Host" (YOU only)
- Groups by stroke type
- Calculates average speeds
- Counts spin distributions
- Detects 1st vs 2nd serves by pattern

### 📝 **File Naming**

**SwingVision Default (Best):**
```
✅ SwingVision-match-2025-11-09 at 15.59.30.xlsx
✅ SwingVision-match-2024-03-15 at 14.30.45.xlsx
```
- Includes date AND time
- No renaming needed!

**Alternative Formats (Supported):**
```
✅ 2024-03-15-vs-opponent.xlsx (YYYY-MM-DD)
✅ match-20240315.xlsx (YYYYMMDD)
✅ tennis-03-15-2024.xlsx (MM-DD-YYYY)
```

**Problematic:**
```
❌ match1.xlsx (no date)
❌ latest.xlsx (no date)
```

---

## 💻 **MENU FUNCTIONS**

Access via **"🎾 Tennis Stats"** menu:

### 🔍 **Check for New Matches Now**
- Immediately scans SwingVision folder
- Processes any new files
- Updates charts
- **Use for**: Instant processing without waiting

### 🔄 **Reprocess All Files**
- Clears all existing data
- Reprocesses every file in folder
- Regenerates all charts
- **Use for**: 
  - After script updates
  - To fix data issues
  - When changing parsing logic

### 📊 **Update Charts**
- Regenerates all 5 performance charts
- Uses existing Match Data
- **Use for**: Chart display issues

### ⚙️ **Install Automatic Check**
- Enables time-based trigger
- Uses configured interval (hour or minute)
- **Use for**: Initial setup or reactivation

### 🛑 **Uninstall Automatic Check**
- Disables time-based trigger
- Stops automatic checking
- **Use for**: 
  - Temporary pause
  - Before configuration changes
  - Quota management

### 🗑️ **Clear Diagnostic Log**
- Clears all entries from Diagnostic sheet
- Keeps header row
- **Use for**: Clean slate after troubleshooting

---

## 🔧 **TROUBLESHOOTING**

### ❌ **Problem: No files being processed**

**Solutions:**
1. Check Diagnostic sheet for error messages
2. Verify folder name matches exactly (case-sensitive)
3. Ensure XLSX files contain "Stats" sheet
4. Check file permissions (script needs Drive access)
5. Run "Check for New Matches Now" manually
6. View Apps Script execution log for detailed errors

### ❌ **Problem: Date extraction not working**

**Solutions:**
1. Check filename includes date in supported format
2. SwingVision default format works best
3. View Diagnostic sheet for parsing warnings
4. Manually test with: `extractDateFromFilename("your-filename.xlsx")`

### ❌ **Problem: Charts empty or incorrect**

**Solutions:**
1. Ensure Match Data sheet has data (at least 2 rows)
2. Click "🎾 Tennis Stats" → "Update Charts"
3. Check numeric columns contain numbers (not text)
4. Verify column references in chart functions
5. Run "Reprocess All Files" to rebuild from scratch

### ❌ **Problem: Script stops running automatically**

**Solutions:**
1. Check trigger status in Apps Script:
   - Click "Triggers" (left sidebar)
   - Verify trigger is active
2. Reinstall trigger:
   - "🎾 Tennis Stats" → "Uninstall Automatic Check"
   - Wait 10 seconds
   - "🎾 Tennis Stats" → "Install Automatic Check"
3. Check quota usage:
   - Apps Script Dashboard → View executions
   - If quota exceeded, increase interval

### ❌ **Problem: Speed/spin data showing as 0**

**Solutions:**
1. Verify XLSX file contains "Shots" sheet
2. Check Shots sheet has data (not empty)
3. View Diagnostic sheet for "Shots sheet not found" warning
4. Older SwingVision files may not have Shots data
5. Process newer files (after SwingVision added shot tracking)

### ❌ **Problem: Permission errors**

**Solutions:**
1. Reauthorize script in Apps Script editor
2. Remove and reinstall authorization
3. Check Google Drive permissions
4. Ensure account has access to folder

---

## 🎯 **ADVANCED USAGE**

### 🎯 **Unforced Error Analysis** 🆕

**Understanding Your Errors:**

The system tracks exactly HOW you're making unforced errors:
- **Net errors**: Ball doesn't clear the net
- **Out errors**: Ball lands long or wide

**For each error type, you know:**
- Which stroke (forehand vs backhand)
- What spin you used (topspin, flat, slice)

**Example Analysis:**
```
Forehand Errors:
- Net: 8 (Topspin: 6, Flat: 2, Slice: 0)
- Out: 4 (Topspin: 1, Flat: 3, Slice: 0)

Insight: You're netting topspin forehands too much!
Action: Work on net clearance with topspin
```

```
Backhand Errors:
- Net: 3 (Topspin: 2, Slice: 1, Flat: 0)
- Out: 9 (Flat: 7, Topspin: 2, Slice: 0)

Insight: Flat backhands going long consistently
Action: Add more topspin or improve depth control
```

**Key Questions Answered:**
1. **Am I netting or hitting out more?**
   - Compare Net vs Out totals for each wing
   
2. **Which spin causes more errors?**
   - See which spin type appears most in errors
   
3. **Is one wing more prone to net errors?**
   - Compare FH Net vs BH Net
   
4. **Are flat shots causing problems?**
   - Check flat counts in error categories

**Actionable Insights:**

| Error Pattern | Likely Cause | Solution |
|---------------|-------------|----------|
| High Net + Topspin | Too much spin, not enough height | Increase net clearance |
| High Net + Flat | Ball trajectory too low | Lift the ball more |
| High Out + Flat | Overhitting | Add topspin or reduce power |
| High Out + Topspin | Wrong brush angle | Check swing path |
| High Net + Slice | Not lifting enough | Open racket face earlier |
| High Out + Slice | Too flat slice | Increase downward angle |

### 📊 **Speed & Spin Insights**

**Speed Benchmarks by Level:**

| Level | 1st Serve | 2nd Serve | Forehand | Backhand |
|-------|-----------|-----------|----------|----------|
| Recreational | 70-90 | 50-70 | 45-60 | 40-55 |
| Club | 90-110 | 70-85 | 60-75 | 55-70 |
| Competitive | 110-130 | 85-100 | 75-90 | 70-85 |
| Professional | 120+ | 90+ | 85+ | 80+ |

*All speeds in MPH*

**Typical Spin Distributions:**

**Serves:**
- Flat: 50-70% (power, first serves)
- Kick: 20-40% (high bounce, second serves)
- Slice: 5-15% (wide angles)

**Forehands:**
- Topspin: 70-85% (control, consistency)
- Flat: 10-25% (winners, pressure)
- Slice: 5-10% (variety, defense)

**Backhands:**
- Topspin: 60-75% (rallying)
- Slice: 20-35% (variety, approach)
- Flat: 5-15% (countering)

### 🎯 **Setting Performance Goals**

**Power Development:**
```
Current: 1st Serve = 95 mph
Target: 1st Serve = 105 mph
Timeline: 3 months
Action: Strength training, serve mechanics
```

**Shot Variety:**
```
Current: 95% topspin forehands
Issue: Too predictable
Target: 15-20% flat drives
Action: Practice flat penetration drills
```

**Tactical Balance:**
```
Current: FH 72 mph, BH 58 mph (14 mph gap)
Issue: Exploitable backhand weakness
Target: <10 mph gap
Action: Backhand power development
```

### 🔄 **Custom Metrics**

**To add custom statistics:**

1. **Modify parsing function** (`parseSwingVisionStats` or `parseShotsData`)
2. **Add column** to Match Data sheet headers
3. **Update data insertion** (`addMatchData` function)
4. **Create custom chart** (optional)
5. **Run "Reprocess All Files"**

**Example - Add net points:**
```javascript
// In parseSwingVisionStats():
else if (label === "net points won") {
  stats.netPointsWon = sumHostColumns(row);
}

// In getMatchDataSheet() headers array:
"Net Points Won",

// In addMatchData() rowData array:
matchData.netPointsWon || 0,
```

---

## 📈 **PERFORMANCE & OPTIMIZATION**

### ⚡ **Resource Usage**

**Typical Execution Times:**
- Match file processing: 10-15 seconds
- Chart generation: 2-3 seconds
- Diagnostic logging: <1 second
- Full reprocess (10 files): 2-3 minutes

**Google Apps Script Limits:**
- Script runtime: 6 minutes per execution
- Daily runtime: 90 minutes total
- Trigger frequency: 1 minute minimum
- Maximum triggers: 20 per user

### 🔋 **Quota Management**

**Recommended Intervals by Usage:**

| Match Frequency | Recommended Interval | Daily Runtime |
|----------------|----------------------|---------------|
| Multiple/day | 5 minutes | ~24 minutes ✅ |
| Daily | 1 hour | ~2 minutes ✅ |
| 3-4/week | 2-3 hours | ~1 minute ✅ |
| Weekly | 6-12 hours | <1 minute ✅ |

**If You Hit Quota Limits:**
1. Increase check interval (e.g., 1 min → 5 min)
2. Use manual checking instead of automatic
3. Wait 24 hours for quota reset
4. Check Apps Script dashboard for usage

---

## 📝 **DATA MANAGEMENT**

### 💾 **Backup Strategy**

**Regular Backups:**
- Download spreadsheet: File → Download → Excel (.xlsx)
- Frequency: Weekly or after significant matches
- Store locally or cloud backup service

**Script Backup:**
- Copy script code to local file
- Save after any customizations
- Version control recommended

### 📦 **Data Export**

**Export Match Data:**
1. Select Match Data sheet
2. File → Download → CSV
3. Import into analysis tools (Excel, Python, R)

**Export Charts:**
1. Right-click chart → Download
2. Save as PNG or PDF
3. Include in reports or presentations

### 🗄️ **Storage Limits**

**Google Sheets Limits:**
- 5 million cells total
- 18,278 columns max (you use 42)
- Unlimited rows (practically)

**With Current Setup:**
- 42 columns × 100,000 matches = 4.2M cells ✅
- Can track thousands of matches
- Very unlikely to hit limits

---

## 🔄 **VERSION HISTORY**

### **v1.2 - Unforced Error Analysis** (Latest)
- **NEW**: 🎯 Unforced error breakdown (Net vs Out)
- **NEW**: Error analysis by stroke (Forehand/Backhand)
- **NEW**: Spin used on each error type
- **ADDED**: 16 new error analysis columns (58 total)
- **ENHANCED**: Understand exactly how you're making errors
- **ACTIONABLE**: Clear insights for improvement

### **v1.1 - Speed & Spin Analytics**
- **NEW**: Shot-by-shot analysis from Shots sheet
- **NEW**: Speed tracking (serves, forehand, backhand)
- **NEW**: Spin distribution (all shot types)
- **NEW**: 2 additional charts (serve speed, shot speed)
- **ADDED**: 15 new data columns (42 total)
- **ENHANCED**: Comprehensive speed/spin insights

### **v1.0 - Initial Release**
- Automatic SwingVision folder monitoring
- Stats sheet parsing
- Match Data sheet with 30 columns
- 3 performance charts
- Diagnostic logging system
- Time-based triggers
- Manual menu functions

---

## 🎓 **BEST PRACTICES**

### ✅ **Do:**
- Upload matches regularly (weekly or after each session)
- Review charts to track trends
- Check Diagnostic sheet if issues arise
- Back up spreadsheet monthly
- Use manual check for testing
- Keep diagnostic logging enabled
- Start with default settings

### ❌ **Don't:**
- Use 1-minute checking unless necessary
- Modify core parsing functions without testing
- Delete Diagnostic sheet (helpful for troubleshooting)
- Process non-SwingVision files
- Change folder name without reprocessing
- Skip authorization steps

### 💡 **Tips:**
- SwingVision filenames include dates - no renaming needed
- 5-minute interval is best balance of speed and quota
- Manual check is instant and quota-free
- Charts update automatically after processing
- Diagnostic sheet is your troubleshooting friend
- Speed/spin data requires newer SwingVision files

---

## 🆘 **SUPPORT**

### 📚 **Resources**
- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [SpreadsheetApp Reference](https://developers.google.com/apps-script/reference/spreadsheet)
- [SwingVision Support](https://swingvision.io/support)

### 🔍 **Debugging Checklist**
1. ✅ Check Diagnostic sheet for errors
2. ✅ View Apps Script execution log
3. ✅ Run "Check for New Matches Now" manually
4. ✅ Verify folder name and permissions
5. ✅ Ensure XLSX files have required sheets
6. ✅ Check quota usage in Apps Script dashboard
7. ✅ Test with sample file first

### 💬 **Common Questions**

**Q: Can I track opponent stats too?**
A: Currently no - script focuses on YOUR performance only.

**Q: Works with other tennis tracking apps?**
A: Only SwingVision format currently supported.

**Q: Can I track multiple players?**
A: Create separate spreadsheets for each player.

**Q: What if I don't have Shots sheet data?**
A: Older files work fine - speed/spin will be 0, other stats still tracked.

**Q: How far back can I go?**
A: Process all historical SwingVision files at once.

---

## 🔧 **TROUBLESHOOTING**

### Common Errors and Solutions

#### ❌ **Error: "Service Spreadsheets failed while accessing document"**

**Problem**: Script can't read XLSX files directly.

**Solution**: Enable Drive API (required for XLSX conversion)

1. Open Apps Script editor
2. Click **"Services"** (➕ icon in left sidebar)
3. Find **"Drive API"** → Click **"Add"**
4. Use version "v2" and identifier "Drive"
5. Save and run script again

> ✅ The script now automatically converts XLSX → Google Sheets → reads data → deletes temp file

---

#### ❌ **Error: Target file not found (Single File Mode)**

**Problem**: Filename doesn't match exactly.

**Common mistakes:**
- Missing `.xlsx` extension
- Typo in filename
- Wrong case (filenames are case-sensitive)

**Solution**: Check exact filename in Google Drive

```javascript
// ❌ Wrong - missing extension
const TARGET_SINGLE_FILE = "SwingVision-match-2025-12-03 at 22.59.47";

// ✅ Correct - includes .xlsx
const TARGET_SINGLE_FILE = "SwingVision-match-2025-12-03 at 22.59.47.xlsx";
```

---

#### ❌ **Error: "Stats" sheet not found**

**Problem**: XLSX file doesn't have a "Stats" sheet.

**Solution**: Verify your SwingVision export includes "Stats" sheet

1. Open XLSX in Excel/Google Sheets
2. Check sheet tabs at bottom
3. Must have: "Stats" (and optionally "Shots")

---

#### ⚠️ **Warning: Shots sheet not found**

**Not an error** - Speed and spin stats will be unavailable.

**Solution**: Use newer SwingVision exports that include "Shots" sheet

- Script continues without error
- Basic stats still work
- Missing: serve speed, shot speed, spin distribution, detailed UE analysis

---

#### ❌ **Error: Trigger quota exceeded**

**Problem**: Running checks too frequently (usually 1-minute intervals).

**Solution**: Change to less frequent interval

```javascript
// Switch from 1 minute to 5 minutes or hourly
const CHECK_INTERVAL_HOURS = 1;  // Safer option
const CHECK_INTERVAL_MINUTES = null;
```

**Quota limits:**
- Total runtime: 90 minutes/day
- 1-minute checks = ~24 min/day ✅
- 5-minute checks = ~5 min/day ✅✅

---

#### 🔍 **Diagnostic Logs**

Check the **"Diagnostic"** sheet for detailed error messages:
- Timestamp of errors
- Specific file causing issues
- Full error messages
- Processing steps completed

---

## 🏆 **CONCLUSION**

The Tennis Stats Analyzer provides professional-grade performance tracking with minimal setup and zero ongoing maintenance. With comprehensive statistics, visual charts, and detailed analytics, it's everything you need to track and improve your tennis game.

The system is designed for reliability, ease of use, and actionable insights - from initial setup through years of tennis tracking.

**Happy analyzing!** 🎾📊✨

---

*Last Updated: December 2024*  
*Version: 1.3.0*  
*Features: 60+ metrics, 22 charts organized by category, shot-by-shot analysis, comprehensive error analysis*  
*Total Functions: 40+*  
*Lines of Code: 2,000+*  
*Sheets Created: 3 (Match Data, Performance Charts, Diagnostic)*  
*Latest: 🆕 Reorganized charts by category, serve spin %, detailed error spin composition (4 charts)*  
*Data Sources: Stats sheet (aggregated) + Shots sheet (detailed)*  
*Error Tracking: Net/Out breakdown + spin analysis for FH/BH Net/Out + serve error composition*
