// ============================================================
// WOODSTOCK DTC NEIGHBOR NETWORK — v2 APPS SCRIPT
// Separate spreadsheet, wide storage, HEADER-MAP architecture.
//
// CORE PRINCIPLE: no column is ever referenced by hardcoded
// position. getHeaderMap() reads the header row once and every
// access goes through it. Adding a volunteer (new column on the
// right) or a VAN column reorder can never silently break this.
//
// SETUP:
// 1. Import v2_voter_master.csv as a tab named exactly "Sheet1".
//    On import, leave VANID/VHHID as text (don't convert to numbers).
// 2. Extensions -> Apps Script -> paste this file -> Save.
// 3. Deploy -> New deployment -> Web app ->
//      Execute as: Me | Who has access: Anyone -> Deploy.
// 4. Copy the web app URL into the v2 HTML (APPS_SCRIPT_URL).
// ============================================================

const SHEET_NAME = "Sheet1";

// Known fixed voter-field headers (everything NOT in this set, to the
// right of them, is treated as a member/volunteer column).
const VOTER_FIELDS = [
  "Voter File VANID","VHHID","LastName","FirstName","MiddleName","Suffix",
  "Sex","Age","Party","StreetNo","StreetNoHalf","StreetPrefix","StreetName",
  "StreetType","StreetSuffix","AptType","AptNo","City","State","Zip5","HD",
  "PreferredEmail","Preferred Phone"
];

// Fields the /voters endpoint is ALLOWED to emit to the browser.
// Anything not listed here (e.g. any future support-score column) is
// withheld server-side no matter what gets added to the sheet later.
const CLIENT_SAFE_FIELDS = [
  "Voter File VANID","VHHID","LastName","FirstName","Age","Party",
  "StreetNo","StreetPrefix","StreetName","StreetType","City","HD"
];

const HOUSEHOLD_CONFIRM_THRESHOLD = 8; // batch-N confirm above this size

// ------------------------------------------------------------
// HEADER MAP — the heart of v2. Built once per request.
// ------------------------------------------------------------
function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                       .map(h => String(h).trim());
  const map = {};                 // header name -> 1-based column index
  headers.forEach((h, i) => { if (h) map[h] = i + 1; });

  // Member columns = any header to the right of the voter fields that
  // isn't itself a known voter field.
  const voterSet = {};
  VOTER_FIELDS.forEach(f => voterSet[f] = true);
  const members = [];             // [{name, col}]
  headers.forEach((h, i) => {
    if (h && !voterSet[h]) members.push({ name: h, col: i + 1 });
  });

  return { headers, map, members, lastCol };
}

function col(hm, name) {
  const c = hm.map[name];
  if (!c) throw new Error("Column not found in header map: " + name);
  return c;
}

// Read whole sheet once, return {headers, rows, hm}. rows are raw arrays.
function readAll(sheet) {
  const hm = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, hm.lastCol).getValues()
    : [];
  return { hm, rows: values, lastRow };
}

// Compose a full road name from its VAN components, in order:
//   StreetPrefix + StreetName + StreetType + StreetSuffix
// e.g. "W" + "Quasset" + "Rd" -> "W Quasset Rd". This keeps W/E/N
// Quasset Rd distinct instead of collapsing them to "Quasset Rd".
function composeRoad(get) {
  const parts = [
    String(get("StreetPrefix") || "").trim(),
    String(get("StreetName")   || "").trim(),
    String(get("StreetType")   || "").trim(),
    String(get("StreetSuffix") || "").trim(),
  ].filter(p => p !== "");
  return parts.join(" ").trim();
}

// Build a voter object from a raw row using the header map.
function rowToVoter(row, hm) {
  const get = (name) => {
    const c = hm.map[name];
    return c ? row[c - 1] : "";
  };
  const road = composeRoad(get) || "Street Not Listed";
  const num = String(get("StreetNo") || "").trim();
  const address = (num + " " + road).trim();
  return {
    vanid:   String(get("Voter File VANID") || "").trim(),
    vhhid:   String(get("VHHID") || "").trim(),
    last:    String(get("LastName") || "").trim(),
    first:   String(get("FirstName") || "").trim(),
    age:     get("Age") === "" ? null : parseInt(get("Age"), 10),
    party:   String(get("Party") || "").trim(),
    road:    road,
    address: address,
    city:    String(get("City") || "").trim(),
  };
}

function answerFor(row, hm, memberName) {
  const c = hm.map[memberName];
  if (!c) return "";
  return String(row[c - 1] || "").trim();
}

function isAnswered(v) {
  return v !== "" && v !== "null" && v !== "undefined";
}

// ============================================================
// GET ROUTER
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action;
  const member = p.member;

  try {
    if (!action || action === "ping")
      return json({ status: "ok", message: "WDTC v2 API running." });

    if (action === "members")        return apiMembers();
    if (action === "voters")         return apiVoters();
    if (action === "progress")       return apiProgress(member);
    if (action === "summary")        return apiSummary(member);
    if (action === "dashboard")      return apiDashboard();
    if (action === "lookup")         return apiLookup(p.vanids);
    if (action === "households")     return apiHouseholdsCoverage();
    if (action === "householdQueue") return apiHouseholdQueue(member);
    if (action === "search")         return apiSearch(p.q);
    if (action === "leaderboard")    return apiLeaderboard();

    return json({ success: false, error: "Unknown action: " + action });
  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

function sheet_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!s) throw new Error('Sheet "' + SHEET_NAME + '" not found.');
  return s;
}

// ---- members ----
function apiMembers() {
  const hm = getHeaderMap(sheet_());
  return json({ success: true, members: hm.members.map(m => m.name) });
}

// ---- voters (ALLOWLISTED — scores can never leak) ----
function apiVoters() {
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const voters = rows.map(r => {
    const o = {};
    CLIENT_SAFE_FIELDS.forEach(f => {
      const c = hm.map[f];
      if (c) o[f] = r[c - 1];
    });
    // convenience fields the old frontend expects
    o.id = String(o["Voter File VANID"] || "").trim();
    o.name = ((o.FirstName || "") + " " + (o.LastName || "")).trim();
    o.road = composeRoad((name) => o[name] !== undefined ? o[name] : "");
    o.address = ((String(o.StreetNo||"").trim()) + " " + o.road).trim();
    o.Age = o.Age;
    return o;
  });
  return json({ voters });
}

// ---- progress (per member, by road) — v1-compatible shape ----
function apiProgress(member) {
  if (!member) return json({ success: false, error: "Missing member." });
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const memberCol = hm.map[member.trim()];
  const roadMap = {};
  rows.forEach(r => {
    const v = rowToVoter(r, hm);
    const road = v.road;
    if (!road) return;
    if (!roadMap[road]) roadMap[road] = { total: 0, reviewed: 0 };
    roadMap[road].total++;
    if (memberCol && isAnswered(String(r[memberCol - 1] || "").trim()))
      roadMap[road].reviewed++;
  });
  const roads = Object.keys(roadMap).sort().map(name => ({
    name,
    total: roadMap[name].total,
    reviewed: roadMap[name].reviewed,
    isDone: roadMap[name].reviewed >= roadMap[name].total && roadMap[name].total > 0
  }));
  return json({ success: true, member, roads });
}

// ---- summary (per member counts) — computed live, v1-compatible ----
function apiSummary(member) {
  if (!member) return json({ success: false, error: "Missing member." });
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const memberCol = hm.map[member.trim()];
  let yes = 0, ko = 0, no = 0;
  if (memberCol) {
    rows.forEach(r => {
      const a = String(r[memberCol - 1] || "").trim();
      if (a === "Y") yes++;
      else if (a === "KO") ko++;
      else if (a === "N") no++;
    });
  }
  return json({ success: true, member, yes, ko, no, reviewed: yes + ko + no });
}

// ---- dashboard — v1-compatible shape ----
function apiDashboard() {
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const members = hm.members;
  const totalVoters = rows.length;
  let totalResponses = 0;
  const voterKnown = new Array(totalVoters).fill(false);
  const voterKnownOf = new Array(totalVoters).fill(false);
  const perMember = {};
  members.forEach(m => perMember[m.name] = 0);

  rows.forEach((r, idx) => {
    members.forEach(m => {
      const a = String(r[m.col - 1] || "").trim();
      if (isAnswered(a)) {
        totalResponses++; perMember[m.name]++;
        if (a === "Y") voterKnown[idx] = true;
        if (a === "KO") voterKnownOf[idx] = true;
      }
    });
  });

  let membersStarted = 0, membersCompleted = 0;
  members.forEach(m => {
    if (perMember[m.name] > 0) membersStarted++;
    if (perMember[m.name] === totalVoters) membersCompleted++;
  });

  const knownCount = voterKnown.filter(Boolean).length;
  const knownOfCount = voterKnownOf.filter((v, i) => v && !voterKnown[i]).length;
  const uniqueConnections = knownCount + knownOfCount;
  const coveragePct = totalVoters > 0
    ? Math.round((uniqueConnections / totalVoters) * 1000) / 10 : 0;

// Per-party coverage (D / R / U). Non-D/non-R (minor parties, blank)
  // all fold into U. Uses the same voterKnown / voterKnownOf arrays,
  // so the scale matches the overall bar exactly.
  const partyRollup = {
    D: { known: 0, knownOf: 0, total: 0 },
    R: { known: 0, knownOf: 0, total: 0 },
    U: { known: 0, knownOf: 0, total: 0 }
  };
  rows.forEach((r, idx) => {
    const v = rowToVoter(r, hm);
    const raw = String(v.party || "").trim().toUpperCase();
    const bucket = (raw === "D") ? "D" : (raw === "R") ? "R" : "U";
    partyRollup[bucket].total++;
    if (voterKnown[idx]) partyRollup[bucket].known++;
    else if (voterKnownOf[idx]) partyRollup[bucket].knownOf++;
  });
  return json({
    success: true, totalVoters, totalResponses, uniqueConnections,
    coveragePct, knownCount, knownOfCount, membersStarted, membersCompleted,
    partyRollup
  });
}

// ---- lookup (comma-separated VANIDs) — v1-compatible shape ----
function apiLookup(vanidsParam) {
  if (!vanidsParam) return json({ success: false, error: "Missing vanids." });
  const requested = String(vanidsParam).split(",").map(v => v.trim()).filter(Boolean);
  if (!requested.length) return json({ success: false, error: "No valid VANIDs." });

  const s = sheet_();
  const { hm, rows } = readAll(s);
  const members = hm.members;
  const vIdx = {};
  rows.forEach((r, i) => {
    const v = String(r[col(hm, "Voter File VANID") - 1] || "").trim();
    if (v) vIdx[v] = i;
  });

  const results = requested.map(vanid => {
    const i = vIdx[vanid];
    if (i === undefined) return { vanid, found: false };
    const r = rows[i];
    const v = rowToVoter(r, hm);
    const connections = [];
    members.forEach(m => {
      const a = String(r[m.col - 1] || "").trim();
      if (a === "Y" || a === "KO") connections.push({ member: m.name, answer: a });
    });
    return {
      vanid, found: true,
      name: (v.first + " " + v.last).trim(),
      age: v.age,
      address: [v.address, v.city].filter(Boolean).join(", "),
      connections,
      isKnown: connections.some(c => c.answer === "Y"),
      isKnownOf: connections.some(c => c.answer === "KO") && !connections.some(c => c.answer === "Y"),
      hasAnyConnection: connections.length > 0
    };
  });

  return json({
    success: true,
    queried: requested.length,
    found: results.filter(r => r.found).length,
    notFound: results.filter(r => !r.found).length,
    withKnown: results.filter(r => r.found && r.isKnown).length,
    withKnownOf: results.filter(r => r.found && r.isKnownOf).length,
    unconnected: results.filter(r => r.found && !r.hasAnyConnection).length,
    results
  });
}

// ---- households coverage (road rollup) — v1-compatible-ish ----
function apiHouseholdsCoverage() {
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const members = hm.members;
  // group by VHHID
  const hh = {};
  rows.forEach(r => {
    const v = rowToVoter(r, hm);
    if (!v.vhhid) return;
    if (!hh[v.vhhid]) hh[v.vhhid] = { road: v.road, hasY: false, hasKO: false };
    members.forEach(m => {
      const a = String(r[m.col - 1] || "").trim();
      if (a === "Y") hh[v.vhhid].hasY = true;
      if (a === "KO") hh[v.vhhid].hasKO = true;
    });
  });
  const roadMap = {};
  Object.keys(hh).forEach(id => {
    const { road, hasY, hasKO } = hh[id];
    if (!roadMap[road]) roadMap[road] = { total: 0, known: 0, knownOf: 0, unknown: 0 };
    roadMap[road].total++;
    if (hasY) roadMap[road].known++;
    else if (hasKO) roadMap[road].knownOf++;
    else roadMap[road].unknown++;
  });
  const roads = Object.keys(roadMap).sort().map(name => {
    const r = roadMap[name];
    const connected = r.known + r.knownOf;
    return { name, total: r.total, known: r.known, knownOf: r.knownOf,
             unknown: r.unknown, connected,
             pct: r.total ? Math.round(connected / r.total * 100) : 0 };
  });
  const totalHouseholds = roads.reduce((s, r) => s + r.total, 0);
  const totalConnected  = roads.reduce((s, r) => s + r.connected, 0);
  const totalKnown      = roads.reduce((s, r) => s + r.known, 0);
  const totalKnownOf    = roads.reduce((s, r) => s + r.knownOf, 0);
  const totalUnknown    = roads.reduce((s, r) => s + r.unknown, 0);
  return json({
    success: true, totalHouseholds, totalConnected,
    totalKnown, totalKnownOf, totalUnknown,
    coveragePct: totalHouseholds ? Math.round(totalConnected / totalHouseholds * 1000) / 10 : 0,
    roads
  });
}

// ============================================================
// HOUSEHOLD QUEUE — the new v2 review engine (Goal 4)
// Returns this member's unreviewed households, grouped by VHHID,
// ordered by: own road, then surname matches to their Y connections,
// then remainder. Members who finished a prior pass see only
// households containing >=1 voter they haven't answered (delta).
// ============================================================
function apiHouseholdQueue(member) {
  if (!member) return json({ success: false, error: "Missing member." });
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const memberCol = hm.map[member.trim()];

  // Group rows by VHHID
  const households = {};
  rows.forEach(r => {
    const v = rowToVoter(r, hm);
    if (!v.vhhid) return;
    if (!households[v.vhhid]) {
      households[v.vhhid] = {
        vhhid: v.vhhid, road: v.road, address: v.address,
        surnames: {}, members: []
      };
    }
    const ans = memberCol ? String(r[memberCol - 1] || "").trim() : "";
    households[v.vhhid].surnames[v.last] = true;
    households[v.vhhid].members.push({
      vanid: v.vanid, first: v.first, last: v.last,
      age: v.age, party: v.party,   // party kept server-side; UI reveals post-answer only
      answered: isAnswered(ans),
      priorAnswer: isAnswered(ans) ? ans : null
    });
  });

  // Surnames this member personally knows (Y) — for affinity ordering
  const knownSurnames = {};
  if (memberCol) {
    rows.forEach(r => {
      if (String(r[memberCol - 1] || "").trim() === "Y") {
        const v = rowToVoter(r, hm);
        if (v.last) knownSurnames[v.last] = true;
      }
    });
  }

  // Keep only households with >=1 unanswered member (delta queue)
  let queue = Object.values(households).filter(
    h => h.members.some(m => !m.answered)
  );

  // Order: surname-affinity first, then alphabetical by road then address.
  // (Own-road / adjacency needs the member's address; can be layered in
  //  later if you store member home roads. For now: affinity + road sort.)
  queue.forEach(h => {
    h.affinity = Object.keys(h.surnames).some(sn => knownSurnames[sn]) ? 1 : 0;
    h.size = h.members.length;
    h.unreviewedCount = h.members.filter(m => !m.answered).length;
    h.needsConfirm = h.size > HOUSEHOLD_CONFIRM_THRESHOLD;
    // surname label for the card
    
    const sns = Object.keys(h.surnames);
    h.surnameLabel = sns.length <= 3
    ? sns.join(" / ")
    : (sns.length + " surnames");
  });
 
 // Group queue households by road, then order roads RANDOMLY, but float
  // roads with >=10 reviewable households to the front (also shuffled
  // among themselves). This spreads volunteers across different roads so
  // we don't all converge on the same streets, while still front-loading
  // the high-yield roads. Households within a road stay address-sorted.
  const BIG_ROAD = 10;
  const byRoad = {};
  queue.forEach(h => {
    if (!byRoad[h.road]) byRoad[h.road] = [];
    byRoad[h.road].push(h);
  });
  function shuffle(arr) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [arr[i], arr[j]] = [arr[j], arr[i]];
    }
    return arr;
  }
  const roadNames = Object.keys(byRoad);
  const bigRoads   = shuffle(roadNames.filter(r => byRoad[r].length >= BIG_ROAD));
  const smallRoads = shuffle(roadNames.filter(r => byRoad[r].length <  BIG_ROAD));
  const orderedRoadNames = bigRoads.concat(smallRoads);

  queue = [];
  orderedRoadNames.forEach(rn => {
    byRoad[rn].sort((a, b) => a.address.localeCompare(b.address));
    byRoad[rn].forEach(h => queue.push(h));
  });

  return json({
    success: true, member,
    householdCount: queue.length,
    households: queue
  });
}

// ---- search (for "I just met someone") ----
function apiSearch(q) {
  if (!q || String(q).trim().length < 2)
    return json({ success: false, error: "Query too short." });
  const needle = String(q).trim().toLowerCase();
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const out = [];
  for (let i = 0; i < rows.length && out.length < 50; i++) {
    const v = rowToVoter(rows[i], hm);
    const hay = (v.first + " " + v.last + " " + v.address + " " + v.road).toLowerCase();
    if (hay.indexOf(needle) !== -1) {
      out.push({ vanid: v.vanid, name: (v.first + " " + v.last).trim(),
                 age: v.age, address: v.address, road: v.road });
    }
  }
  return json({ success: true, count: out.length, results: out });
}

// ============================================================
// POST — writes. Single voter, household batch, add member.
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const s = sheet_();

    if (data.action === "addMember") {
      const name = String(data.name || "").trim();
      if (!name) return json({ success: false, error: "Name required." });
      const c = findOrCreateMemberColumn(s, name);
      return json({ success: !!c, name });
    }

    if (data.action === "householdBatchN") {
      return writeHouseholdN(s, data);
    }
  // In doPost, add this branch before the single-write fallback:
    if (data.action === "batchWrite") {
      return writeBatch(s, data);
    }
    // Single-voter write (also used by "I just met someone" overwrite & undo)
    return writeSingle(s, data);

  } catch (err) {
    return json({ success: false, error: String(err) });
  }
}

// New function: writes an array of {voterId, answer} for one member in a
// single request. answer in Y/KO/N/"" ("" clears). Mirrors writeSingle's
// validation and styling, but loops over many rows with one header-map read.
function writeBatch(s, data) {
  const { memberName, answers } = data; // answers: [{voterId, answer}, ...]
  if (!memberName) return json({ success: false, error: "Missing member." });
  if (!Array.isArray(answers) || !answers.length)
    return json({ success: false, error: "No answers supplied." });

  const hm = getHeaderMap(s);
  const vanCol = col(hm, "Voter File VANID");
  const lastRow = s.getLastRow();
  const colValues = s.getRange(2, vanCol, lastRow - 1, 1).getValues();
  const rowOf = {};
  for (let i = 0; i < colValues.length; i++)
    rowOf[String(colValues[i][0]).trim()] = i + 2;

  const memberCol = findOrCreateMemberColumn(s, memberName);
  const valid = ["Y", "N", "KO", ""];
  let written = 0, skipped = 0;
  answers.forEach(a => {
    const ans = String(a.answer);
    if (!valid.includes(ans)) { skipped++; return; }
    const rr = rowOf[String(a.voterId).trim()];
    if (!rr) { skipped++; return; }
    const cell = s.getRange(rr, memberCol);
    if (ans === "") { cell.clearContent(); cell.setBackground(null); cell.setFontColor(null); }
    else { cell.setValue(ans); styleAnswer(cell, ans); }
    written++;
  });
  return json({ success: true, written, skipped });
}

function writeSingle(s, data) {
  const { voterId, memberName, answer } = data;
  const valid = ["Y", "N", "KO", ""];
  if (!valid.includes(answer))
    return json({ success: false, error: "Invalid answer." });
  if (!voterId) return json({ success: false, error: "Missing voterId." });

  const hm = getHeaderMap(s);
  const vanCol = col(hm, "Voter File VANID");
  const lastRow = s.getLastRow();
  const vanids = s.getRange(2, vanCol, lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < vanids.length; i++) {
    if (String(vanids[i][0]).trim() === String(voterId).trim()) {
      targetRow = i + 2; break;
    }
  }
  if (targetRow === -1)
    return json({ success: false, error: "Voter not found: " + voterId });

  const memberCol = findOrCreateMemberColumn(s, memberName);
  const cell = s.getRange(targetRow, memberCol);
  if (answer === "") { cell.clearContent(); cell.setBackground(null); cell.setFontColor(null);
    return json({ success: true, status: "cleared", voterId }); }
  cell.setValue(answer);
  styleAnswer(cell, answer);
  return json({ success: true, voterId, answer });
}

function writeHouseholdN(s, data) {
  const { vhhid, memberName, vanids } = data; // vanids: array to mark N
  if (!memberName) return json({ success: false, error: "Missing member." });
  if (!Array.isArray(vanids) || !vanids.length)
    return json({ success: false, error: "No vanids supplied." });

  const hm = getHeaderMap(s);
  const vanCol = col(hm, "Voter File VANID");
  const lastRow = s.getLastRow();
  const colValues = s.getRange(2, vanCol, lastRow - 1, 1).getValues();
  const rowOf = {};
  for (let i = 0; i < colValues.length; i++)
    rowOf[String(colValues[i][0]).trim()] = i + 2;

  const memberCol = findOrCreateMemberColumn(s, memberName);
  let written = 0;
  vanids.forEach(v => {
    const rr = rowOf[String(v).trim()];
    if (rr) {
      const cell = s.getRange(rr, memberCol);
      cell.setValue("N");
      styleAnswer(cell, "N");
      written++;
    }
  });
  return json({ success: true, vhhid, written });
}

function styleAnswer(cell, answer) {
  if (answer === "Y") cell.setBackground("#d4edda").setFontColor("#155724");
  else if (answer === "KO") cell.setBackground("#fff3cd").setFontColor("#856404");
  else cell.setBackground("#f8d7da").setFontColor("#721c24");
}

// Creates a new member column on the RIGHT (the growing-roster design).
function findOrCreateMemberColumn(sheet, memberName) {
  const name = String(memberName || "").trim();
  if (!name) return null;
  const hm = getHeaderMap(sheet);
  if (hm.map[name]) return hm.map[name];
  const newCol = hm.lastCol + 1;
  sheet.getRange(1, newCol).setValue(name)
    .setBackground("#e8d5a3").setFontColor("#1a2744")
    .setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setColumnWidth(newCol, 140);
  return newCol;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ---- leaderboard (per-member contribution + roads completed) ----
function apiLeaderboard() {
  const s = sheet_();
  const { hm, rows } = readAll(s);
  const members = hm.members;

  // Per-member tallies
  const stats = {};
  members.forEach(m => stats[m.name] = { name: m.name, yes: 0, ko: 0, no: 0 });

  // Unique knowns: voters this member marked Y where NO OTHER member did.
  // First, for each row, count how many members marked Y.
  const yCountPerRow = rows.map(r => {
    let c = 0;
    members.forEach(m => { if (String(r[m.col - 1] || "").trim() === "Y") c++; });
    return c;
  });
  // A member's unique knowns = rows where they are Y and the row's total Y count is exactly 1.
  members.forEach(m => {
    let uniq = 0;
    rows.forEach((r, idx) => {
      if (String(r[m.col - 1] || "").trim() === "Y" && yCountPerRow[idx] === 1) uniq++;
    });
    stats[m.name].uniqueKnown = uniq;
  });

  // Road completion needs, per member: count of roads where the member has
  // answered every voter on that road. Build road -> [rowIdx] once.
  const roadRows = {};
  rows.forEach((r, idx) => {
    const v = rowToVoter(r, hm);
    const road = v.road;
    if (!road) return;
    (roadRows[road] = roadRows[road] || []).push(idx);
  });

  // Tally answers
  rows.forEach(r => {
    members.forEach(m => {
      const a = String(r[m.col - 1] || "").trim();
      if (a === "Y") stats[m.name].yes++;
      else if (a === "KO") stats[m.name].ko++;
      else if (a === "N") stats[m.name].no++;
    });
  });

  // Roads completed per member
  members.forEach(m => {
    let done = 0;
    Object.keys(roadRows).forEach(road => {
      const idxs = roadRows[road];
      const allAnswered = idxs.every(i => isAnswered(String(rows[i][m.col - 1] || "").trim()));
      if (allAnswered && idxs.length > 0) done++;
    });
    stats[m.name].roadsCompleted = done;
  });

  // Shape rows + derived fields
  const out = members.map(m => {
    const st = stats[m.name];
    const completed = st.yes + st.ko + st.no;     // total reviewed
    const known = st.yes;                          // most known
    const knownPlusOf = st.yes + st.ko;            // known + known of
    return {
      name: m.name, completed, known, knownPlusOf,
      knownOf: st.ko, no: st.no,
      roadsCompleted: st.roadsCompleted,
      uniqueKnown: st.uniqueKnown
    };
  }).filter(r => r.completed > 0); // only participants

  // Town-wide rollup for the banner
  const totals = out.reduce((t, r) => {
    t.completed += r.completed; t.known += r.known; t.knownOf += r.knownOf;
    return t;
  }, { completed: 0, known: 0, knownOf: 0 });

  return json({
    success: true,
    participantCount: out.length,
    totalVolunteers: members.length,
    totals,
    members: out
  });
}
