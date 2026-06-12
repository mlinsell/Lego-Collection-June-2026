// ─────────────────────────────────────────────────────────────────────────────
//  Mark's Lego Collection Tracker — Code.gs
// ─────────────────────────────────────────────────────────────────────────────

const SPREADSHEET_ID   = '13F9LqpkDi5iw_kUnnWrYy5nFizeGmM3svln-S77m0EI';
const COLLECTION_SHEET = 'Collection';
const DATA_START_ROW   = 2;
const SETS_CACHE_KEY = 'sets_cache_v1';
const SHORTLIST_CACHE_KEY = 'shortlist_cache_v1';
const SETS_CACHE_TTL = 60; // seconds

// ── Collection column positions (1-indexed) ───────────────────────────────────
const COL = {
  SET_ID:        1,   // A  Set_ID            - editable
  NAME:          2,   // B  Set Name          - editable
  STATUS:        3,   // C  Built Status       - editable (dropdown)
  LOCATION:      4,   // D  Location          - editable
  BUILT_DATE:    5,   // E  Built Date        - editable
  RETAILER:      6,   // F  Retailer          - editable
  PURCHASE_DATE: 7,   // G  Purchase Date     - editable
  PURCHASE_YEAR: 8,   // H  Purchase Year     - formula (auto)
  PRICE:         9,   // I  Purchase Price    - editable
  OWNERSHIP:     10,  // J  Ownership         - editable (dropdown)
  KEEP_SELL:     11,  // K  Keep / Sell       - editable
  PURPOSE:       12,  // L  Purpose           - editable
  THEME:         13,  // M  Theme             - STATIC VALUE (Rebrickable API)
  NAME_LOOKUP:   14,  // N  Set_Name_Lookup   - STATIC VALUE (Rebrickable API)
  YEAR:          15,  // O  Year              - STATIC VALUE (Rebrickable API)
  RRP:           16,  // P  UK RRP            - STATIC VALUE (Brickset API)
  PIECES:        17,  // Q  Pieces            - STATIC VALUE (Rebrickable API)
  AVAILABILITY:  18,  // R  Availability      - VLOOKUP (Dynamic)
  BL_NEW:        19,  // S  BrickLink New     - VLOOKUP (Dynamic)
  BL_USED:       20,  // T  BrickLink Used    - VLOOKUP (Dynamic)
  BACKLOG:       21,  // U  Backlog           - formula (read-only)
  CONDITION:     22,  // V  Condition         - formula (read-only)
  SHORTLIST:     23,  // W  Shortlist         - user flag (TRUE/FALSE)
  SHORTLIST_PRIORITY: 24, // X Shortlist priority - numeric rank
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle("Mark's Lego Collection Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getName() !== "") return ss;
  } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🧱 Lego Tools')
      .addItem('Sync Data Tools', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Lego Data Sync').setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function ping() {
  try {
    const ss = getSpreadsheet();
    const sheetNames = ss.getSheets().map(s => s.getName());
    return { ok: true, spreadsheetName: ss.getName(), sheets: sheetNames, hasCollection: sheetNames.includes(COLLECTION_SHEET) };
  } catch(e) { return { ok: false, error: e.message }; }
}

function calcBuildHours(pieces, theme) {
  if (!pieces || pieces === '') return '';
  const p = parseFloat(pieces);
  if (isNaN(p) || p <= 0) return '';
  const mults = { 'Technic':2.0, 'Creator Expert':1.5,'Icons':1.5,'Architecture':1.5,'Modular Buildings':1.5, 'Ideas':1.2,'Disney':1.2,'Star Wars':1.2,'Harry Potter':1.2,'Winter Village':1.2 };
  return Math.round(p / 10 * (mults[theme] || 1.0) / 60 * 10) / 10;
}

function calcBuildCategory(hours) {
  if (hours === '' || hours === null || hours === undefined) return '';
  const h = parseFloat(hours);
  if (isNaN(h)) return '';
  if (h < 1)  return '⚡ Quick Build';
  if (h < 3)  return '☀️ Afternoon Project';
  if (h < 8)  return '🏠 Weekend Build';
  if (h < 20) return '📅 Multi-Day Build';
  return '🏔️ Epic Build';
}

function getSets() {
  try {
    const cache = CacheService.getUserCache();
    const cached = cache.get(SETS_CACHE_KEY);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) { /* fall through to refresh */ }
    }
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    if (!sheet) return { error: 'Tab not found.' };
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return { sets: [] };
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();

    const sets = data.map((row, i) => {
      if (!row[0] && !row[1]) return null;
      const pieces = toInt(row[COL.PIECES - 1]);
      const theme  = String(row[COL.THEME - 1] || '');
      const hrs    = calcBuildHours(pieces, theme);
      const shortlistRaw = data[i][COL.SHORTLIST - 1];
      const shortlist = (shortlistRaw === true) || (String(shortlistRaw||'').toLowerCase() === 'true') || (shortlistRaw === 1);
      const priorityRaw = data[i][COL.SHORTLIST_PRIORITY - 1];
      const shortlistPriority = (priorityRaw === undefined || priorityRaw === '') ? '' : Number(priorityRaw);
      return {
        rowIndex:     DATA_START_ROW + i,
        setNo:        String(row[COL.SET_ID - 1] || '').trim(),
        name:         String(row[COL.NAME - 1] || ''),
        status:       String(row[COL.STATUS - 1] || ''),
        location:     String(row[COL.LOCATION - 1] || ''),
        builtDate:    formatDate(row[COL.BUILT_DATE - 1]),
        retailer:     String(row[COL.RETAILER - 1] || ''),
        purchaseDate: formatDate(row[COL.PURCHASE_DATE - 1]),
        price:        toNum(row[COL.PRICE - 1]),
        ownership:    String(row[COL.OWNERSHIP - 1] || ''),
        keepSell:     String(row[COL.KEEP_SELL - 1] || ''),
        purpose:      String(row[COL.PURPOSE - 1] || ''),
        theme:        theme,
        year:         toInt(row[COL.YEAR - 1]),
        rrp:          toNum(row[COL.RRP - 1]),
        pieces:       pieces,
        availability: String(row[COL.AVAILABILITY - 1] || ''),
        blNew:        toNum(row[COL.BL_NEW - 1]),
        blUsed:       toNum(row[COL.BL_USED - 1]),
        condition:    String(row[COL.CONDITION - 1] || ''),
        buildHrs:     hrs,
        buildCat:     calcBuildCategory(hrs),
        shortlist:    shortlist,
        shortlistPriority: shortlistPriority,
      };
    }).filter(Boolean);
    try{ cache.put(SETS_CACHE_KEY, JSON.stringify({ sets }), SETS_CACHE_TTL); }catch(e){}
    return { sets };
  } catch(e) { return { error: e.message }; }
}

function clearSetsCache(){
  try{ CacheService.getUserCache().remove(SETS_CACHE_KEY); }catch(e){}
  try{ CacheService.getUserCache().remove(SHORTLIST_CACHE_KEY); }catch(e){}
}

function getShortlist(){
  try {
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    if (!sheet) return { error: 'Tab not found.' };
    const lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return { sets: [] };
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 24).getValues();
    const sets = data.map((row, i) => {
      if (!row[0] && !row[1]) return null;
      const shortlistRaw = data[i][COL.SHORTLIST - 1];
      const shortlist = (shortlistRaw === true) || (String(shortlistRaw||'').toLowerCase() === 'true') || (shortlistRaw === 1);
      if (!shortlist) return null;
      const pieces = toInt(row[COL.PIECES - 1]);
      const theme  = String(row[COL.THEME - 1] || '');
      const hrs    = calcBuildHours(pieces, theme);
      const priorityRaw = data[i][COL.SHORTLIST_PRIORITY - 1];
      const shortlistPriority = (priorityRaw === undefined || priorityRaw === '') ? '' : Number(priorityRaw);
      return {
        rowIndex: DATA_START_ROW + i,
        setNo: String(row[COL.SET_ID - 1] || '').trim(),
        name: String(row[COL.NAME - 1] || ''),
        status: String(row[COL.STATUS - 1] || ''),
        theme: theme,
        pieces: pieces,
        buildHrs: hrs,
        buildCat: calcBuildCategory(hrs),
        shortlist: shortlist,
        shortlistPriority: shortlistPriority,
      };
    }).filter(Boolean);
    return { sets };
  } catch(e) { return { error: e.message }; }
}

function applyRowUpdates(sheet, rowNumber, updates){
  try{
    const cols = Object.keys(updates).map(c => parseInt(c,10)).sort((a,b)=>a-b);
    if (cols.length === 0) return;
    const minCol = cols[0];
    const maxCol = cols[cols.length - 1];
    const width = maxCol - minCol + 1;
    const rowVals = new Array(width).fill('');
    cols.forEach(col => { rowVals[col - minCol] = updates[col]; });
    sheet.getRange(rowNumber, minCol, 1, width).setValues([rowVals]);
  }catch(e){
    // best-effort: fall back to individual writes if batch fails
    try{
      cols.forEach(col => { sheet.getRange(rowNumber, col).setValue(updates[col]); });
    }catch(e2){}
  }
}

function addSet(data) {
  try {
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    const nextRow = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
    sheet.getRange(nextRow, 1, 1, 24).setValues([buildRow(data, nextRow)]);
    clearSetsCache();
    return { success: true, rowIndex: nextRow };
  } catch(e) { return { error: e.message }; }
}

function updateSet(rowIndex, data) {
  try {
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    sheet.getRange(rowIndex, 1, 1, 24).setValues([buildRow(data, rowIndex)]);
    clearSetsCache();
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function deleteSet(rowIndex) {
  try {
    getSpreadsheet().getSheetByName(COLLECTION_SHEET).deleteRow(rowIndex);
    clearSetsCache();
    return { success: true };
  } catch(e) { return { error: e.message }; }
}

function buildRow(data, r) {
  return [
    data.setNo || '',
    data.name  || '',
    data.status || '',
    data.location || '',
    data.builtDate || '',
    data.retailer || '',
    data.purchaseDate || '',
    `=IF(G${r}="","",YEAR(G${r}))`, // H: Purchase Year
    data.price !== undefined && data.price !== '' ? parseFloat(data.price) : '',
    data.ownership || '',
    data.keepSell || '',
    data.purpose || '',
    data.theme || '',
    data.name || '',
    data.year || '',
    data.rrp !== undefined && data.rrp !== '' ? parseFloat(data.rrp) : '',
    data.pieces !== undefined && data.pieces !== '' ? parseInt(data.pieces) : '',
    `=VLOOKUP(A${r},'Brickset-List Export'!B:V,21,FALSE)`, // R: Availability
    `=VLOOKUP(A${r},'Brickset-List Export'!B:AQ,41,FALSE)`, // S: BL New
    `=VLOOKUP(A${r},'Brickset-List Export'!B:AQ,42,FALSE)`, // T: BL Used
    `=VLOOKUP(C${r},Inputs!A:C,2,FALSE)`, // U: Backlog 
    `=VLOOKUP(C${r},Inputs!A:C,3,FALSE)`  // V: Condition 
    , data.shortlist ? true : ''
    , data.shortlistPriority !== undefined && data.shortlistPriority !== '' ? parseInt(data.shortlistPriority) : ''
  ];
}

function setShortlist(rowIndex, value){
  try{
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    sheet.getRange(rowIndex, COL.SHORTLIST).setValue(value?true:'');
    // If removing from shortlist, also clear the priority cell
    if(!value){
      sheet.getRange(rowIndex, COL.SHORTLIST_PRIORITY).setValue('');
    }
    clearSetsCache();
    return { success: true };
  }catch(e){ return { error: e.message }; }
}

function setShortlistPriority(rowIndex, priority){
  try{
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    sheet.getRange(rowIndex, COL.SHORTLIST_PRIORITY).setValue(priority!==''?parseInt(priority):'');
    clearSetsCache();
    return { success: true };
  }catch(e){ return { error: e.message }; }
}

function updateShortlistPriorities(updates){
  try{
    const sheet = getSpreadsheet().getSheetByName(COLLECTION_SHEET);
    updates.forEach(u=>{
      const r = u.rowIndex;
      const p = (u.priority === undefined || u.priority === '') ? '' : parseInt(u.priority);
      sheet.getRange(r, COL.SHORTLIST_PRIORITY).setValue(p);
    });
    clearSetsCache();
    return { success: true };
  }catch(e){ return { error: e.message }; }
}

// ── API Logic ─────────────────────────────────────────────────────────────────
function getRebrickableKey() { return '36e0f7e0ccd74c532308f2854a7408c0'; }
function getBricksetKey() { return '3-iuiZ-56qL-yTuAU'; }

function lookupSet(setNo) {
  try {
    const rebKey = getRebrickableKey();
    const bsKey  = getBricksetKey();
    const clean = String(setNo).trim().replace(/-1$/, '');
    const setFull = clean + '-1';

    const rebResp = UrlFetchApp.fetch('https://rebrickable.com/api/v3/lego/sets/' + setFull + '/?key=' + rebKey, { muteHttpExceptions: true });
    if (rebResp.getResponseCode() !== 200) return { error: 'Set not found on Rebrickable' };
    const d = JSON.parse(rebResp.getContentText());

    let rawTheme = '';
    if (d.theme_id) {
      try {
        const tr = UrlFetchApp.fetch('https://rebrickable.com/api/v3/lego/themes/' + d.theme_id + '/?key=' + rebKey, { muteHttpExceptions: true });
        if (tr.getResponseCode() === 200) rawTheme = JSON.parse(tr.getContentText()).name || '';
      } catch(e) {}
    }

    let rrp = '';
    try {
      const bsParams = JSON.stringify({ setNumber: setFull });
      const bsResp = UrlFetchApp.fetch(`https://brickset.com/api/v3.asmx/getSets?apiKey=${bsKey}&userHash=&params=${encodeURIComponent(bsParams)}`, { muteHttpExceptions: true });
      if (bsResp.getResponseCode() === 200) {
        const bs = JSON.parse(bsResp.getContentText());
        if (bs.sets && bs.sets.length > 0) {
          const setObj = bs.sets[0];
          rrp = (setObj.retailPrice && setObj.retailPrice.UK !== undefined) ? setObj.retailPrice.UK :
                (setObj.LEGOCom && setObj.LEGOCom.UK && setObj.LEGOCom.UK.retailPrice !== undefined ? setObj.LEGOCom.UK.retailPrice : '');
        }
      }
    } catch(e) {}

    return { name: d.name || '', pieces: d.num_parts || '', year: d.year || '', rawTheme: rawTheme, rrp: rrp, ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── UI Progress Tracker ───────────────────────────────────────────────────────
function getSyncProgress() {
  return CacheService.getUserCache().get('syncProgress') || JSON.stringify({text: 'Preparing scan...', logs: []});
}

// ── Fill Missing Basic Data (Theme, Name, Year, Pieces, RRP) ──────────────────
function fillMissingData() {
  const cache = CacheService.getUserCache();
  try {
    cache.put('syncProgress', JSON.stringify({text: 'Initializing scan...', logs: []}), 60);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(COLLECTION_SHEET);
    const data = sheet.getDataRange().getValues();
    const rebKey = getRebrickableKey();
    const bsKey = getBricksetKey();
    
    let count = 0;
    const total = data.length - 1;
    let logs = [];

    // Correctly identifies missing data but safely ignores "0" as intentional data
    const isMissing = (v) => {
      if (v === 0 || v === '0') return false; 
      return !v || String(v).trim() === '' || String(v).trim().startsWith('#');
    };

    for (let i = 1; i < data.length; i++) {
      const setID = String(data[i][COL.SET_ID - 1]).trim();
      if (!setID) continue;

      cache.put('syncProgress', JSON.stringify({
        text: `Scanning row ${i} of ${total} (Updates made: ${count})`, 
        logs: logs
      }), 60);

      const theme = data[i][COL.THEME - 1];
      const nameLookup = data[i][COL.NAME_LOOKUP - 1];
      const year = data[i][COL.YEAR - 1];
      const pieces = data[i][COL.PIECES - 1];
      const rrp = data[i][COL.RRP - 1];

      const missingReb = isMissing(theme) || isMissing(nameLookup) || isMissing(year) || isMissing(pieces);
      const missingBs = isMissing(rrp);

      if (!missingReb && !missingBs) continue;

      const clean = setID.replace(/-1$/, '');
      const setFull = clean + '-1';

      let newTheme = theme, newName = nameLookup, newYear = year, newPieces = pieces, newRrp = rrp;
      let actuallyUpdated = [];

      // Rebrickable API calls for Name, Theme, Year, Pieces
      if (missingReb) {
        try {
          const rebResp = UrlFetchApp.fetch('https://rebrickable.com/api/v3/lego/sets/' + setFull + '/?key=' + rebKey, { muteHttpExceptions: true });
          if (rebResp.getResponseCode() === 200) {
            const d = JSON.parse(rebResp.getContentText());
            if (isMissing(newName)) newName = d.name || '';
            if (isMissing(newYear)) newYear = d.year || '';
            if (isMissing(newPieces) && d.num_parts !== undefined) newPieces = d.num_parts;
            if (isMissing(newTheme) && d.theme_id) {
              const tr = UrlFetchApp.fetch('https://rebrickable.com/api/v3/lego/themes/' + d.theme_id + '/?key=' + rebKey, { muteHttpExceptions: true });
              if (tr.getResponseCode() === 200) newTheme = JSON.parse(tr.getContentText()).name || '';
            }
          } else {
            logs.unshift(`⚠️ Rebrickable error for ${setID}: HTTP ${rebResp.getResponseCode()}`);
          }
        } catch(e) {
          logs.unshift(`⚠️ Rebrickable error for ${setID}: ${e.message}`);
        }
        Utilities.sleep(150); // Rebrickable allows multiple reqs per second
      }

      // Brickset API calls for UK RRP
      if (missingBs) {
        try {
          const bsParams = JSON.stringify({ setNumber: setFull });
          const bsResp = UrlFetchApp.fetch(`https://brickset.com/api/v3.asmx/getSets?apiKey=${bsKey}&userHash=&params=${encodeURIComponent(bsParams)}`, { muteHttpExceptions: true });
          if (bsResp.getResponseCode() === 200) {
            const bs = JSON.parse(bsResp.getContentText());
            if (bs.sets && bs.sets.length > 0) {
              const setObj = bs.sets[0];
              // Check standard retailPrice path AND LEGO.com specific path, handling 0 correctly
              newRrp = (setObj.retailPrice && setObj.retailPrice.UK !== undefined) ? setObj.retailPrice.UK :
                       (setObj.LEGOCom && setObj.LEGOCom.UK && setObj.LEGOCom.UK.retailPrice !== undefined ? setObj.LEGOCom.UK.retailPrice : '');
              
              if (newRrp === '') logs.unshift(`ℹ️ Set ${setID}: No UK RRP found in Brickset database.`);
            }
          } else {
            logs.unshift(`⚠️ Brickset Rate Limit hit for ${setID}. Trying to continue...`);
          }
        } catch(e) {
          logs.unshift(`⚠️ Brickset error for ${setID}: ${e.message}`);
        }
        // STRICT BRICKSET RATE LIMIT: Must be >1000ms to avoid IP ban
        Utilities.sleep(1200); 
      }

      // Prepare batched updates for this row to minimise calls to the Sheet service
      const rowUpdates = {};
      if (isMissing(theme) && !isMissing(newTheme)) { rowUpdates[COL.THEME] = newTheme; actuallyUpdated.push('Theme'); }
      if (isMissing(nameLookup) && !isMissing(newName)) { rowUpdates[COL.NAME_LOOKUP] = newName; actuallyUpdated.push('Name'); }
      if (isMissing(year) && !isMissing(newYear)) { rowUpdates[COL.YEAR] = newYear; actuallyUpdated.push('Year'); }
      if (isMissing(pieces) && !isMissing(newPieces)) { rowUpdates[COL.PIECES] = newPieces; actuallyUpdated.push('Pieces'); }
      if (isMissing(rrp) && !isMissing(newRrp)) { rowUpdates[COL.RRP] = newRrp; actuallyUpdated.push('RRP'); }

      if (actuallyUpdated.length > 0) {
        applyRowUpdates(sheet, i + 1, rowUpdates);
        logs.unshift(`✅ Set <b>${setID}</b>: Fetched ${actuallyUpdated.join(', ')}`);
        if (logs.length > 8) logs.pop();
        count++;
      }
    }
    
    cache.remove('syncProgress');
    return JSON.stringify({
      text: `Success! Scanned collection and filled missing data for ${count} sets.`,
      logs: logs
    });
  } catch (e) {
    CacheService.getUserCache().remove('syncProgress');
    return JSON.stringify({ text: "Error: " + e.message, logs: [] });
  }
}

// ── Standard Bulk API Updaters (for P and Q only) ─────────────────────────────
function bulkUpdateFromAPI(type) {
  const cache = CacheService.getUserCache();
  try {
    cache.put('syncProgress', JSON.stringify({text: 'Initializing scan...', logs: []}), 60);
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(COLLECTION_SHEET);
    const data = sheet.getDataRange().getValues();
    const bsKey = getBricksetKey();
    
    let count = 0;
    const total = data.length - 1;
    let logs = [];

    for (let i = 1; i < data.length; i++) {
      const setID = String(data[i][COL.SET_ID - 1]).trim();
      if (!setID) continue;

      cache.put('syncProgress', JSON.stringify({
        text: `Scanning row ${i} of ${total} (Updates made: ${count})`, 
        logs: logs
      }), 60);

      const cell = sheet.getRange(i + 1, type === 'RRP' ? COL.RRP : COL.PIECES);
      if (String(cell.getFormula()) !== '' || cell.getValue() === '') {
        const setFull = setID.includes('-') ? setID : setID + '-1';
        const bsParams = JSON.stringify({ setNumber: setFull });
        
        try {
          const res = JSON.parse(UrlFetchApp.fetch(`https://brickset.com/api/v3.asmx/getSets?apiKey=${bsKey}&userHash=&params=${encodeURIComponent(bsParams)}`, { muteHttpExceptions: true }).getContentText());
          if (res.sets && res.sets.length > 0) {
            let val = type === 'RRP' ? (res.sets[0].retailPrice ? res.sets[0].retailPrice.UK : '') : res.sets[0].pieces;
            if (val !== '' && val !== undefined) { 
              cell.setValue(val); 
              logs.unshift(`✅ Set <b>${setID}</b>: ${type} = ${val}`);
              if (logs.length > 8) logs.pop();
              count++; 
            }
          }
        } catch(e) {}
        Utilities.sleep(150);
      }
    }
    cache.remove('syncProgress');
    return JSON.stringify({
      text: `Success! Updated ${count} sets with static ${type} data via API.`,
      logs: logs
    });
  } catch (e) {
    CacheService.getUserCache().remove('syncProgress');
    return JSON.stringify({ text: "Error: " + e.message, logs: [] });
  }
}

function bulkUpdateRRP() { return bulkUpdateFromAPI('RRP'); }
function bulkUpdatePieces() { return bulkUpdateFromAPI('PIECES'); }

function toNum(val) {
  if (val === '' || val === null || val === undefined) return '';
  const n = parseFloat(val);
  return isNaN(n) ? '' : n;
}

function toInt(val) {
  if (val === '' || val === null || val === undefined) return '';
  const n = parseInt(val);
  return isNaN(n) ? '' : n;
}

function formatDate(val) {
  if (!val || val === '') return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val);
}