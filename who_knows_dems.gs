// ============================================================
// WOODSTOCK DTC NEIGHBOR NETWORK
// Google Apps Script
//
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Click Extensions → Apps Script
// 3. Delete any existing code and paste this entire script
// 4. Click Save (floppy disk icon)
// 5. Run setupSheet() once to add the Road column + member headers:
//      - Click the function dropdown (top bar), select "setupSheet"
//      - Click Run ▶
//      - Accept permissions when prompted
// 6. Deploy as web app:
//      - Click Deploy → New deployment
//      - Type: Web app
//      - Execute as: Me
//      - Who has access: Anyone
//      - Click Deploy and copy the web app URL
// 7. Paste that URL into the HTML game (APPS_SCRIPT_URL constant)
// ============================================================

// --- CONFIGURATION -------------------------------------------
const SHEET_NAME = "Sheet1";   // Change if your tab is named differently

const MEMBERS = [
  "Aino Kardestuncer",
  "Alana Frenkel",
  "Bob Freudenberger",
  "Charlene Perkins Cutler",
  "Charles Super",
  "Chelsea Weiss Baum",
  "Christine Lessig",
  "Crystal Adams",
  "Emily G. Hayden",
  "Ethan Werstler",
  "Gail Dickenson",
  "Glen Lessig",
  "Greg Kline",
  "Ivar McDonald",
  "Ives McDonald",
  "Jacob Tworzydlo",
  "Jessica Weaver Boose",
  "Jim Kaeding",
  "Joan Polulech",
  "Joe Polulech",
  "John Day",
  "Karen Ryker",
  "Kathleen Barach",
  "Kevin Bernier",
  "Laura Graves",
  "Margie Wholean",
  "Matthew Anderson",
  "Matthew Bennett",
  "Michael Shepherd",
  "Peg Wilson",
  "Polly Hayden",
  "R. Sky Bridgman",
  "Rebecca Hyde",
  "Rick Canavan",
  "Sara Harkness",
  "Sarah Bridgman",
  "Sean Connor",
  "Stacy Bennett",
  "Su Connor",
  "Suzanne Kline",
  "Suzanne Woodward",
  "Vincent Tocci"
];

// Column indices (1-based) of fixed voter data columns
// LastName=1, FirstName=2, MiddleName=3, Suffix=4, Sex=5,
// Address=6, City=7, State=8, HD=9, PrefPhone=10, PreferredEmail=11
const ADDRESS_COL   = 8;
const ROAD_COL      = 12;  // New column inserted after PreferredEmail
const MEMBERS_START = 14;  // Member columns start here

// ============================================================
// SETUP — Formula-Free Engine with Targeted Trailing Unit Scrub
// ============================================================
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${SHEET_NAME}" not found. Check SHEET_NAME in the script.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  
  // --- 1. PULL & SCRUB RAW ROAD STRINGS DIRECTLY FROM COLUMN L (12) ---
  const rawRoadValues = sheet.getRange(2, ROAD_COL, lastRow - 1, 1).getValues();
  const cleanedRoadsForSheet = [];
  
  for (let i = 0; i < rawRoadValues.length; i++) {
    let road = String(rawRoadValues[i][0]).trim();
    if (!road || road === "null" || road === "Road" || road === "Street Not Listed") {
      cleanedRoadsForSheet.push(["Street Not Listed"]);
      continue;
    }
    
    // Targeted cleaning pass: Strip ONLY trailing #, Apt, Unit, or Fl designators
    // This regular expression safely leaves thoroughfares like "Sherman Lane" completely untouched!
    let polishedRoad = road.replace(/\s+(?:#|Apt\.?|Unit|Fl\.?|Floor)\s*.*$/i, '');
    cleanedRoadsForSheet.push([polishedRoad.trim()]);
  }
  
  // Overwrite Column L (12) on Sheet1 with your beautifully polished, standardized street names
  sheet.getRange(2, ROAD_COL, lastRow - 1, 1).setValues(cleanedRoadsForSheet);

  // Read all member reviews starting from Column N (14) to the end of your sheet data
  const rawMemberValues = sheet.getRange(2, MEMBERS_START, lastRow - 1, MEMBERS.length).getValues();
  
  // --- 2. CLEAN REBUILD OF THE ROADPROGRESS TAB ---
  let progressTab = ss.getSheetByName("RoadProgress");
  if (progressTab) {
    ss.deleteSheet(progressTab); // Completely erase old tab to clear cache errors
  }
  progressTab = ss.insertSheet("RoadProgress");

  // Deduplicate and pull out our clean array list of unique roads using the polished list
  const rawRoadList = cleanedRoadsForSheet.map(r => r[0]);
  const uniqueRoads = [...new Set(rawRoadList)]
    .filter(r => r && r !== "" && r !== "Street Not Listed")
    .sort();
    
  // Force our fallback category row right to the very bottom
  uniqueRoads.push("Street Not Listed");

  // Build the horizontal tracking header column layout
  const progressHeaders = ["Road Name", "Total Voters", "Global Progress Status"];
  MEMBERS.forEach(member => { progressHeaders.push(member); });
  progressTab.getRange(1, 1, 1, progressHeaders.length).setValues([progressHeaders]);

  // --- 3. RUN THE CALCULATIONS NATIVELY IN JAVASCRIPT ---
  const matrixRows = [];
  
  // Outer loop: Evaluate one unique street row at a time
  for (let k = 0; k < uniqueRoads.length; k++) {
    const roadName = uniqueRoads[k];
    
    let totalVotersOnRoad = 0;
    let memberDoneCounts = Array(MEMBERS.length).fill(0);
    
    // Inner loop: Scan through all voter records on Sheet1
    for (let r = 0; r < rawRoadList.length; r++) {
      if (rawRoadList[r] === roadName) {
        totalVotersOnRoad++; // We found a voter living on this street!
        
        // Check what answers each individual member logged for this specific voter row
        for (let m = 0; m < MEMBERS.length; m++) {
          let memberAnswer = String(rawMemberValues[r][m]).trim();
          if (memberAnswer !== "" && memberAnswer !== "null" && memberAnswer !== "undefined") {
            memberDoneCounts[m]++;
          }
        }
      }
    }
    
    // --- 4. ASSEMBLE VALUES FOR THIS ROW ---
    const rowCells = [];
    rowCells.push(roadName);        // Column A: Clean text road name
    rowCells.push(totalVotersOnRoad); // Column B: Raw number of total voters
    
    // Column C: Global Progress Status determination
    let totalReviewsOnRoad = memberDoneCounts.reduce((sum, count) => sum + count, 0);
    let totalPossibleReviews = totalVotersOnRoad * MEMBERS.length;
    
    let globalStatus = "🔴 Unstarted";
    if (totalReviewsOnRoad === totalPossibleReviews && totalPossibleReviews > 0) {
      globalStatus = "🟢 Completed";
    } else if (totalReviewsOnRoad > 0) {
      globalStatus = "🟡 In Progress";
    }
    rowCells.push(globalStatus);
    
    // Columns D onward: Format plain completion fraction strings for each member
    for (let m = 0; m < MEMBERS.length; m++) {
      rowCells.push(`${memberDoneCounts[m]} / ${totalVotersOnRoad}`);
    }
    
    matrixRows.push(rowCells);
  }

  // --- 5. WRITE PLAIN DATA TO THE SHEET ---
  progressTab.getRange(2, 1, matrixRows.length, progressHeaders.length).setValues(matrixRows);
  
  // Style and format properties
  progressTab.setFrozenRows(1);
  progressTab.setFrozenColumns(1);
  progressTab.getRange(2, 1, matrixRows.length, 1).setFontWeight("bold");
  progressTab.getRange(1, 1, 1, 3).setBackground("#2a4494").setFontColor("white").setFontWeight("bold");
  progressTab.getRange(1, 4, 1, MEMBERS.length).setBackground("#c8a84b").setFontColor("#1a2744").setFontWeight("bold");
  
  // Center numbers and statuses for readability
  progressTab.getRange(2, 2, matrixRows.length, progressHeaders.length).setHorizontalAlignment("center");
  progressTab.getRange(2, 1, matrixRows.length, 1).setHorizontalAlignment("left");
  
  progressTab.setColumnWidth(1, 220);
  progressTab.setColumnWidth(2, 110);
  progressTab.setColumnWidth(3, 160);
  for (let c = 4; c <= progressHeaders.length; c++) {
    progressTab.setColumnWidth(c, 140);
  }
  buildSummaryStats(sheet, uniqueRoads, rawRoadList, rawMemberValues, lastRow);
  SpreadsheetApp.getUi().alert("✓ Step 1 Complete! Trailing unit anomalies scrubbed and list compiled perfectly with zero formulas.");
  }

// Global utility helper to map column indexes cleanly to spreadsheet text letters (e.g. 1 -> A, 27 -> AA)
function getColumnLetter(col) {
  let letter = "";
  while (col > 0) {
    let t = (col - 1) % 26;
    letter = String.fromCharCode(65 + t) + letter;
    col = (col - t - 1) / 26;
  }
  return letter;
}

// ============================================================
// WEB API FEED — Handles incoming browser GET requests
// ============================================================
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const member = e && e.parameter && e.parameter.member;

  // 1. Base API Status Check / Ping Route
  if (!action || action === "ping") {
    return jsonResponse({ status: "ok", message: "DTC Network API is running perfectly." });
  }
  
  // 1a. Serve the raw data rows mapped into clean objects for the client browser engine
  if (action === "voters") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    const dataRange = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = dataRange[0];
    const voters = [];
    
    for (let i = 1; i < dataRange.length; i++) {
      let rowObj = {};
      for (let j = 0; j < headers.length; j++) {
        // Map headers to key values (e.g. voter.LastName, voter.VANID, voter.road)
        let keyName = headers[j].trim();
        rowObj[keyName] = dataRange[i][j];
      }
      
      // Legacy layout backward-compatibility mapping fallbacks:
      rowObj.id = rowObj.VANID || ("v" + i); // VANID becomes your anchor ID token!
      rowObj.name = `${rowObj.FirstName || ""} ${rowObj.LastName || ""}`.trim();
      rowObj.address = rowObj.Address || "";
      rowObj.road = rowObj.Road || "";
      
      voters.push(rowObj);
    }
    
    return jsonResponse({ voters });
  }

  // 2. action=ages — Returns raw voter age data arrays for game initialization
  if (action === "ages") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    // Column G (7) = Age, rows 2..lastRow
    const ageValues = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
    const ages = ageValues.map(r => r[0] || null);
    return jsonResponse({ ages });
  }

// ─── NEW: SUMMARY STATS API ENDPOINT ──────────────────────────────────────
  if (action === "summary") {
    if (!member) return jsonResponse({ success: false, error: "Missing required member name parameter." });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summaryTab = ss.getSheetByName("SummaryStats");
    if (!summaryTab) return jsonResponse({ success: false, error: "SummaryStats tab not found. Re-run setupSheet." });

    const lastCol = summaryTab.getLastColumn();
    const lastRow = summaryTab.getLastRow();
    
    // Fetch raw values for header text matching, and display values to read evaluated formula metrics
    const summaryData = summaryTab.getRange(1, 1, lastRow, lastCol).getValues();
    const displayData = summaryTab.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    const headers = summaryData[0];
    const cleanTarget = member.replace(/\s*\*$/, '').replace(/[-\s]/g, '').trim().toLowerCase();

    // Find our member's column index dynamically by stripping spaces/hyphens
    let memberColIdx = -1;
    for (let c = 1; c < headers.length; c++) {
      const cleanHeader = String(headers[c]).replace(/\s*\*$/, '').replace(/[-\s]/g, '').trim().toLowerCase();
      if (cleanHeader === cleanTarget) {
        memberColIdx = c;
        break;
      }
    }

    if (memberColIdx === -1) {
      return jsonResponse({ success: true, yes: 0, ko: 0, no: 0, reviewed: 0 });
    }

    let yesVal = 0, koVal = 0, noVal = 0, reviewedVal = 0;

    // Scan column A text labels, but extract text from displayData to read calculated formula states!
    for (let r = 0; r < summaryData.length; r++) {
      const label = String(summaryData[r][0]).trim().toLowerCase();
      const rawTextDisplay = String(displayData[r][memberColIdx]).trim();
      
      if (label.indexOf("yes") !== -1) {
        yesVal = parseInt(rawTextDisplay, 10) || 0;
      } else if (label.indexOf("know of") !== -1 || label.indexOf("ko") !== -1) {
        koVal = parseInt(rawTextDisplay, 10) || 0;
      } else if (label.indexOf("don't know") !== -1 || label.indexOf("(n)") !== -1) {
        noVal = parseInt(rawTextDisplay, 10) || 0;
      } else if (label.indexOf("reviewed") !== -1) {
        reviewedVal = parseInt(rawTextDisplay, 10) || 0;
      }
    }

    return jsonResponse({
      success: true,
      member: member,
      yes: yesVal,
      ko: koVal,
      no: noVal,
      reviewed: reviewedVal
    });
  }
  // 3. NEW ACTION: action=roads — Fetches available streets and individual member progress
  if (action === "roads") {
    if (!member) {
      return jsonResponse({ success: false, error: "Missing required member parameter query." });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const progressTab = ss.getSheetByName("RoadProgress");
    if (!progressTab) {
      return jsonResponse({ success: false, error: "RoadProgress data tracking tab not found. Run setupSheet first." });
    }

    const lastRow = progressTab.getLastRow();
    const lastCol = progressTab.getLastColumn();
    // Fetch the entire Tab 2 matrix grid into memory
    const progressData = progressTab.getRange(1, 1, lastRow, lastCol).getValues();

    const headers = progressData[0];
    
    // Scan headers to find exactly what column index matches the member logging in right now
    let memberColIdx = -1;
    for (let c = 3; c < headers.length; c++) {
      // Strip out the trailing " *" marker we append for custom guests so names match precisely
      if (headers[c].replace(/\s*\*$/, '').trim().toLowerCase() === member.trim().toLowerCase()) {
        memberColIdx = c;
        break;
      }
    }

    const roadList = [];
    
    // Loop through all generated street rows on Tab 2 (skipping Row 1 header)
    for (let r = 1; r < progressData.length; r++) {
      const roadName = progressData[r][0];
      const totalVoters = parseInt(progressData[r][1], 10) || 0;
      
      let reviewCount = 0;
      
      // If the member column was located, parse their fraction string (e.g. "12 / 42")
      if (memberColIdx !== -1) {
        const fractionString = String(progressData[r][memberColIdx]);
        // Regex extract the number BEFORE the slash line symbol
        const completedMatch = fractionString.match(/^(\d+)\s*\//);
        if (completedMatch) {
          reviewCount = parseInt(completedMatch[1], 10);
        }
      }

      // Package each street up into a structured object format for the HTML app frontend
      roadList.push({
        name: roadName,
        total: totalVoters,
        reviewed: reviewCount,
        isDone: (reviewCount >= totalVoters && totalVoters > 0)
      });
    }

    return jsonResponse({ success: true, member: member, roads: roadList });
  }

  return jsonResponse({ status: "error", message: "Unknown API route invocation parameter context." });
}

// ============================================================
// WEB API WRITE — Receives data updates and searches by VANID
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { voterId, memberName, answer } = data; // voterId here represents the unique VANID string

    // 1. Validate incoming answer values (allowing "" for Undo operations)
    const validAnswers = ["Y", "N", "KO", ""];
    if (!validAnswers.includes(answer)) {
      return jsonResponse({ success: false, error: "Invalid answer value structure submitted." });
    }

    // 2. Validate that we actually received a VANID anchor key
    if (!voterId) {
      return jsonResponse({ success: false, error: "Missing required unique identifier anchor (voterId/VANID)." });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ success: false, error: `Sheet "${SHEET_NAME}" not found.` });
    }

    // 3. DYNAMIC LOOKUP PHASE: Find the voter's row matching their unique VANID
    const lastRow = sheet.getLastRow();
    // VANID is locked in Column F (Column 6). Let's fetch the entire active column into cache memory
    const vanIdColumnValues = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
    
    let targetRowIndex = -1;
    for (let i = 0; i < vanIdColumnValues.length; i++) {
      if (String(vanIdColumnValues[i][0]).trim() === String(voterId).trim()) {
        targetRowIndex = i + 2; // Offset by 2 (row 1 is header, index 0 is row 2)
        break;
      }
    }

    // Error exit route if the tracking code cannot locate that specific individual
    if (targetRowIndex === -1) {
      return jsonResponse({ success: false, error: `Voter identifier "${voterId}" could not be located in Column F.` });
    }

    // 4. Find or auto-generate the current member's data column
    const memberCol = findOrCreateMemberColumn(sheet, memberName);
    if (!memberCol) {
      return jsonResponse({ success: false, error: `Could not reconcile active data column tracking for "${memberName}".` });
    }

    const cell = sheet.getRange(targetRowIndex, memberCol);

    // 5. Execute core cell write manipulation behaviors
    if (answer === "") {
      cell.clearContent();
      cell.setBackground(null); // Wipe styling cleanly on Undo requests
      cell.setFontColor(null);
      return jsonResponse({ success: true, row: targetRowIndex, col: memberCol, status: "cleared", voterId });
    }

    // Write normal score evaluation inputs
    cell.setValue(answer);

    // Apply clean conditional coloring matching standard project style tokens
    if (answer === "Y") {
      cell.setBackground("#d4edda").setFontColor("#155724"); // Green highlight
    } else if (answer === "KO") {
      cell.setBackground("#fff3cd").setFontColor("#856404"); // Gold highlight
    } else {
      cell.setBackground("#f8d7da").setFontColor("#721c24"); // Red highlight
    }

    return jsonResponse({ success: true, row: targetRowIndex, col: memberCol, answer, voterId });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

function findOrCreateMemberColumn(sheet, memberName) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Search existing headers — if found, return immediately
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]).trim() === memberName.trim()) {
      return i + 1;
    }
  }

  // ── Not found: create new column on Sheet1 ──
  const newCol = lastCol + 1;
  const headerCell = sheet.getRange(1, newCol);
  headerCell.setValue(memberName);
  headerCell
    .setBackground("#e8d5a3")
    .setFontColor("#1a2744")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setColumnWidth(newCol, 140);
  Logger.log(`Created new Sheet1 column for "${memberName}" at column ${newCol}`);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const colLetter = columnToLetter(newCol);
  const lastRow = sheet.getLastRow();

  // ── Add member to RoadProgress ──
  const progressTab = ss.getSheetByName("RoadProgress");
  if (progressTab) {
    const progressLastCol = progressTab.getLastColumn();
    const progressLastRow = progressTab.getLastRow();
    const progressHeaders = progressTab.getRange(1, 1, 1, progressLastCol).getValues()[0];

    // Only add if not already present
    const alreadyInProgress = progressHeaders.some(h => String(h).trim() === memberName.trim());
    if (!alreadyInProgress) {
      const newProgressCol = progressLastCol + 1;
      const headerCell2 = progressTab.getRange(1, newProgressCol);
      headerCell2.setValue(memberName);
      headerCell2
        .setBackground("#c8a84b")
        .setFontColor("#1a2744")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      progressTab.setColumnWidth(newProgressCol, 140);

      // Read road names and Sheet1 data to compute 0/total fractions for this new member
      const roadNames = progressTab.getRange(2, 1, progressLastRow - 1, 1).getValues();
      const sheet1Roads = sheet.getRange(2, ROAD_COL, lastRow - 1, 1).getValues();
      const sheet1Answers = sheet.getRange(2, newCol, lastRow - 1, 1).getValues();

      const fractions = roadNames.map(([roadName]) => {
        let total = 0, reviewed = 0;
        for (let r = 0; r < sheet1Roads.length; r++) {
          if (String(sheet1Roads[r][0]).trim() === String(roadName).trim()) {
            total++;
            const ans = String(sheet1Answers[r][0]).trim();
            if (ans !== "" && ans !== "null" && ans !== "undefined") reviewed++;
          }
        }
        return [`${reviewed} / ${total}`];
      });

      if (fractions.length > 0) {
        progressTab.getRange(2, newProgressCol, fractions.length, 1)
          .setValues(fractions)
          .setHorizontalAlignment("center");
      }
    }
  }

  // ── Add member to SummaryStats ──
  const summaryTab = ss.getSheetByName("SummaryStats");
  if (summaryTab) {
    const summaryLastCol = summaryTab.getLastColumn();
    const summaryHeaders = summaryTab.getRange(1, 1, 1, summaryLastCol).getValues()[0];

    const alreadyInSummary = summaryHeaders.some(h => String(h).trim() === memberName.trim());
    if (!alreadyInSummary) {
      const newSummaryCol = summaryLastCol + 1;
      const summaryColLetter = columnToLetter(newSummaryCol);

      summaryTab.getRange(1, newSummaryCol).setValue(memberName)
        .setBackground("#1a2744").setFontColor("white").setFontWeight("bold");
      summaryTab.setColumnWidth(newSummaryCol, 140);

      // Write the same live COUNTIF formulas as buildSummaryStats does
      const formulas = [
        [`=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRow}, "Y")`],
        [`=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRow}, "KO")`],
        [`=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRow}, "N")`],
        [`=SUM(${summaryColLetter}2:${summaryColLetter}4)`],
        [`=(COUNTA(Sheet1!$A$2:$A$${lastRow})) - ${summaryColLetter}5`]
      ];

      summaryTab.getRange(2, newSummaryCol, formulas.length, 1)
        .setFormulas(formulas)
        .setHorizontalAlignment("center");
    }
  }

  return newCol;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// CORS HANDLER — needed when called from a browser
// ============================================================
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}


// ============================================================
// UTILITY — run to see a summary of current coverage
// ============================================================
function coverageSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  // Find member columns
  const memberCols = [];
  for (let c = MEMBERS_START - 1; c < headers.length; c++) {
    if (headers[c]) memberCols.push(c);
  }

  let totalVoters = lastRow - 1;
  let coveredY = 0;      // at least one Y
  let coveredKO = 0;     // at least one KO (no Y)
  let uncovered = 0;     // all blank or N

  for (let r = 1; r < data.length; r++) {
    let hasY = false, hasKO = false;
    memberCols.forEach(c => {
      if (data[r][c] === "Y") hasY = true;
      if (data[r][c] === "KO") hasKO = true;
    });
    if (hasY) coveredY++;
    else if (hasKO) coveredKO++;
    else uncovered++;
  }

  // Member completion rates
  let memberSummary = "Member completion:\n";
  memberCols.forEach(c => {
    const name = headers[c];
    const reviewed = data.slice(1).filter(row => row[c] !== "").length;
    const pct = Math.round((reviewed / totalVoters) * 100);
    memberSummary += `  ${name}: ${reviewed}/${totalVoters} (${pct}%)\n`;
  });

  SpreadsheetApp.getUi().alert(
    `COVERAGE SUMMARY\n` +
    `─────────────────────\n` +
    `Total voters: ${totalVoters}\n` +
    `Known personally (Y): ${coveredY} (${Math.round(coveredY/totalVoters*100)}%)\n` +
    `Known of (KO): ${coveredKO} (${Math.round(coveredKO/totalVoters*100)}%)\n` +
    `Uncovered: ${uncovered} (${Math.round(uncovered/totalVoters*100)}%)\n\n` +
    memberSummary
  );
}

// ============================================================
// TEST — run this manually to verify sheet writing works
// ============================================================
function testWrite() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const memberCol = findOrCreateMemberColumn(sheet, "Ivar McDonald");
  sheet.getRange(13, memberCol).setValue("Y");
  Logger.log("Test write complete — check row 13, column " + memberCol);
}

// ============================================================
// SUMMARY LOGIC — Writes LIVE Formulas to the SummaryStats Tab
// ============================================================
function buildSummaryStats(sheet, uniqueRoads, rawRoadList, rawMemberValues, lastRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Clean Rebuild of the Summary Stats Tab
  let summaryTab = ss.getSheetByName("SummaryStats");
  if (summaryTab) {
    ss.deleteSheet(summaryTab);
  }
  summaryTab = ss.insertSheet("SummaryStats");

  // 2. Build Horizontal Headers (Member Names)
  const headers = ["Metric"];
  MEMBERS.forEach(member => { headers.push(member); });
  summaryTab.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 3. Set up rows for our Formula Matrix
  const formulaMatrix = [
    ["Total Yes (Y)"],
    ["Total Know Of (KO)"],
    ["Total Don't Know (N)"],
    ["Grand Total Reviewed"],
    ["Remaining Voters in Town"]
  ];

  // 4. Generate dynamic =COUNTIF formulas for every single member column
  for (let m = 0; m < MEMBERS.length; m++) {
    // Convert column index to Google Sheet column letters (Column B = Aino, Column C = Alana, etc.)
    // Sheet1 column 1 is VANID, column 2 is FirstName... your member columns start around column 8+
    // Let's find the exact column letter on Sheet1 matching this member name:
    const sheet1Headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const memberColumnIndex = sheet1Headers.findIndex(h => h.trim() === MEMBERS[m].trim()) + 1;
    
    if (memberColumnIndex > 0) {
      const colLetter = columnToLetter(memberColumnIndex);
      const lastRowSheet1 = sheet.getLastRow();
      
      // Build the live formula strings pointing to Sheet1!
      const yesFormula = `=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRowSheet1}, "Y")`;
      const koFormula  = `=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRowSheet1}, "KO")`;
      const noFormula  = `=COUNTIF(Sheet1!$${colLetter}$2:$${colLetter}$${lastRowSheet1}, "N")`;
      
      // The current column letter on the SummaryStats sheet itself (B, C, D...)
      const summaryColLetter = columnToLetter(m + 2); 
      const totalReviewedFormula = `=SUM(${summaryColLetter}2:${summaryColLetter}4)`;
      const remainingFormula = `=(COUNTA(Sheet1!$A$2:$A$${lastRowSheet1})) - ${summaryColLetter}5`;

      // Push formulas into our data matrix
      formulaMatrix[0].push(yesFormula);
      formulaMatrix[1].push(koFormula);
      formulaMatrix[2].push(noFormula);
      formulaMatrix[3].push(totalReviewedFormula);
      formulaMatrix[4].push(remainingFormula);
    } else {
      // Fallback placeholders if a column header isn't found
      formulaMatrix[0].push(0); formulaMatrix[1].push(0); formulaMatrix[2].push(0); formulaMatrix[3].push(0); formulaMatrix[4].push(0);
    }
  }

 // 5. Write the Formula strings directly to the cells!
  // Extract just the text labels for Column A
  const labelsOnly = [
    [formulaMatrix[0][0]],
    [formulaMatrix[1][0]],
    [formulaMatrix[2][0]],
    [formulaMatrix[3][0]],
    [formulaMatrix[4][0]]
  ];

  // Extract ONLY the formula strings for Columns B and beyond
  const formulasOnly = [];
  for (let r = 0; r < formulaMatrix.length; r++) {
    formulasOnly.push(formulaMatrix[r].slice(1));
  }

  // 1. Write plain text labels safely to Column A
  summaryTab.getRange(2, 1, labelsOnly.length, 1).setValues(labelsOnly);

  // 2. Write the live formulas seamlessly to Columns B through the end
  summaryTab.getRange(2, 2, formulasOnly.length, MEMBERS.length).setFormulas(formulasOnly);

  // Apply clean styling themes
  summaryTab.setFrozenColumns(1);
  summaryTab.getRange(1, 1, 1, headers.length).setBackground("#1a2744").setFontColor("white").setFontWeight("bold");
  summaryTab.getRange(2, 1, formulaMatrix.length, 1).setFontWeight("bold").setBackground("#faf8f4");
  summaryTab.getRange(5, 1, 2, headers.length).setFontWeight("bold"); // Bold the totals rows
  summaryTab.getRange(5, 1, 1, headers.length).setBackground("#e8f0fe"); // Light blue highlight for Total Reviewed
  summaryTab.getRange(6, 1, 1, headers.length).setBackground("#fdf2e9"); // Light orange highlight for Remaining
  summaryTab.getRange(2, 2, formulaMatrix.length, MEMBERS.length).setHorizontalAlignment("center");
  summaryTab.setColumnWidth(1, 200);
  for (let c = 2; c <= headers.length; c++) {
    summaryTab.setColumnWidth(c, 140);
  }
}

// Helper utility function to turn numbers into column letters (e.g. 2 -> "B", 28 -> "AB")
function columnToLetter(column) {
  let temp, letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
