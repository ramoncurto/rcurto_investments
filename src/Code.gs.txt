// Code.gs

// Global configuration for column types
const NUMERIC_HEADERS_CONFIG = ["InvestmentEUR", "InvestmentUSD", "SHARES", "CallPriceUSD", "PutPriceUSD", "ProfitLossUSD", "BrokerFeesEUR", "ResultEUR", "DollarExchangeRate", "EuroExchangeRate"];
const PERCENT_HEADERS_CONFIG = ["ProfitLossPercent", "ResultPercent"];
const DATE_HEADERS_CONFIG = ["Date In", "Date Out"];

// ⚠️ CRUCIAL: Verify this SHEET_NAME matches your Google Sheet tab name exactly (case-sensitive).
const SHEET_NAME = "Trades"; // As per your last provided code.gs snippet. Double-check this!

// --- GPT Configuration ---
const GPT_MODEL = 'gpt-4.1-nano'; // As per your request

// --- Analytics Constants ---
const MONTH_NAMES_ANALYTICS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const DAYS_OF_WEEK_ANALYTICS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const PRICE_RANGES_ANALYTICS = [
  { label: "<$0.5", min: -Infinity, max: 0.4999 }, { label: "0.5-0.99", min: 0.5, max: 0.9999 },
  { label: "1-1.99", min: 1, max: 1.9999 }, { label: "2-4.99", min: 2, max: 4.9999 },
  { label: "5-9.99", min: 5, max: 9.9999 }, { label: "10-14.99", min: 10, max: 14.9999 },
  { label: "15-19.99", min: 15, max: 19.9999 }, { label: "20-49.99", min: 20, max: 49.9999 },
  { label: "50-99.99", min: 50, max: 99.9999 }, { label: "100-199.99", min: 100, max: 199.9999 },
  { label: "200-499.99", min: 200, max: 499.9999 }, { label: ">500", min: 500.0001, max: Infinity }
];

// --- User Properties Keys ---
const USER_STRATEGIES_KEY = 'userTradingStrategies_v1';
const USER_ASSET_TYPES_KEY = 'userTradingAssetTypes_v1';
const GOALS_KEY = 'tradingGoalsAndProjections_v2'; 
const USER_PROPERTIES = PropertiesService.getUserProperties();

// --- CORE WEB APP AND GPT FUNCTIONS ---

/**
 * Main function that runs when the web app URL is accessed.
 */
function doGet(e) {
  const htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
  htmlOutput.setTitle('Gestor i Anàlisi de Trading Avançat')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return htmlOutput;
}

/**
 * Calls the OpenAI GPT API. Fetches API key from Script Properties.
 */
function callGPT(task, payload, maxTokens) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    Logger.log("OPENAI_API_KEY not found in Script Properties.");
    return JSON.stringify({ error: "API Key not configured. Please set OPENAI_API_KEY in Project Settings > Script Properties." });
  }

  const url  = 'https://api.openai.com/v1/chat/completions';
  const body = {
    model: GPT_MODEL,
    temperature: 0.2,
    max_tokens: maxTokens || 256,
    messages: [
      { role:'system', content:'Eres RiskGPT, responde SIEMPRE en JSON con claves inglesas.'},
      { role:'user', content: JSON.stringify({task, payload})}
    ],
    functions:[{
      name:"forecast_response", 
      parameters:{ type:"object", properties:{ base:{type:"number"}, best:{type:"number"}, worst:{type:"number"}, comment:{type:"string"} }, required:["base","best","worst","comment"] }
    }],
    function_call:{name:"forecast_response"} 
  };
  const options = { method:'post', contentType:'application/json', headers:{Authorization:`Bearer ${apiKey}`}, payload:JSON.stringify(body), muteHttpExceptions:true };

  try {
    const res = UrlFetchApp.fetch(url, options);
    const resContent = res.getContentText();
    const jsonResponse = JSON.parse(resContent);

    if (res.getResponseCode() !== 200) {
      Logger.log(`GPT API Error: ${res.getResponseCode()} ${resContent}`);
      return JSON.stringify({ error: "GPT API Error", details: jsonResponse });
    }
    if (jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message && jsonResponse.choices[0].message.function_call && jsonResponse.choices[0].message.function_call.arguments) {
      return jsonResponse.choices[0].message.function_call.arguments;
    } else if (jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].message && jsonResponse.choices[0].message.content) {
      return jsonResponse.choices[0].message.content; 
    } else {
      Logger.log(`Unexpected GPT response structure: ${resContent}`);
      return JSON.stringify({ error: "Unexpected GPT response structure", details: jsonResponse });
    }
  } catch (e) {
    Logger.log(`Error in callGPT: ${e.toString()}`);
    return JSON.stringify({ error: `Exception in callGPT: ${e.toString()}` });
  }
}

// --- CLIENT-CALLABLE ENDPOINTS (Data Retrieval & Manipulation) ---

/**
 * Gets the sheet headers. Relies on getSheetHeadersArray() from ConstructUtils.gs
 */
function getHeadersForClient() {
  try {
    if (typeof getSheetHeadersArray !== 'function') {
      Logger.log("Error crític: La funció getSheetHeadersArray no està definida (should be in ConstructUtils.gs).");
      throw new Error("La funció getSheetHeadersArray no està definida.");
    }
    return getSheetHeadersArray();
  } catch (error) {
    Logger.log(`Error a getHeadersForClient: ${error.toString()}`);
    return { error: `Error obtenint capçaleres: ${error.message}` };
  }
}

/**
 * Gets all trades from the active spreadsheet. Parses data.
 */
function getTrades() {
  try {
    Logger.log("getTrades: Attempting to get sheet with SHEET_NAME: " + SHEET_NAME); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (typeof SHEET_NAME === 'undefined') { Logger.log("getTrades: Error - SHEET_NAME undefined."); throw new Error("Constant SHEET_NAME no definida."); }
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) { Logger.log(`getTrades: Error - Sheet "${SHEET_NAME}" NOT found.`); throw new Error(`El full "${SHEET_NAME}" no s'ha trobat. Assegura't que el nom de la pestanya coincideix exactament, incloent majúscules/minúscules i espais. També, comprova que 'ConstructUtils.gs' existeix i 'getSheetHeadersArray' està definit correctament.`); }
    Logger.log(`getTrades: Successfully found sheet: "${SHEET_NAME}"`);

    if (typeof getSheetHeadersArray !== 'function') { Logger.log("getTrades: Error - getSheetHeadersArray undefined."); throw new Error("La funció getSheetHeadersArray no està definida (hauria d'estar a ConstructUtils.gs)."); }
    const headers = getSheetHeadersArray();
    Logger.log("getTrades: Headers from getSheetHeadersArray(): " + JSON.stringify(headers));
    if (!headers || headers.length === 0) { Logger.log("getTrades: Error - Headers array empty."); throw new Error("Headers array invàlid o buit. Verifica getSheetHeadersArray() a ConstructUtils.gs."); }
    
    const lastRow = sheet.getLastRow();
    Logger.log("getTrades: Last row: " + lastRow + ", Headers count: " + headers.length);
    if (lastRow < 2) { Logger.log("getTrades: No data rows."); return []; }

    const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    const values = dataRange.getValues(); 
    Logger.log("getTrades: Fetched " + values.length + " rows.");

    const trades = values.map((row) => {
      const trade = {};
      headers.forEach((header, index) => {
        let cellValue = row[index];
        if (DATE_HEADERS_CONFIG.includes(header) && cellValue instanceof Date && !isNaN(cellValue)) {
          trade[header] = cellValue.toISOString().split('T')[0]; 
        } else if (NUMERIC_HEADERS_CONFIG.includes(header)) {
          let numStr = String(cellValue).replace(',', '.'); 
          trade[header] = parseFloat(numStr) || 0; 
        } else if (PERCENT_HEADERS_CONFIG.includes(header)) {
          let percentStr = String(cellValue).replace('%', '').replace(',', '.').trim();
          let parsedPercent = parseFloat(percentStr);
          trade[header] = !isNaN(parsedPercent) ? parsedPercent / 100 : 0;
        } else {
          trade[header] = cellValue; 
        }
      });
      return trade;
    });
    Logger.log("getTrades: Processed " + trades.length + " trades."); 
    return trades;
  } catch (error) {
    Logger.log(`Error CRÍTIC a getTrades: ${error.toString()}\nStack: ${error.stack}`); 
    return { error: `Error obtenint operacions: ${error.message}` };
  }
}

/**
 * Adds a new trade to the sheet, including GPT validation.
 */
function addTrade(tradeData) {
  try {
    const gptValidationPayload = { ...tradeData }; 
    const gptResponseString = callGPT('validate_trade', gptValidationPayload, 200); 
    const gptResp = JSON.parse(gptResponseString); 
    if (gptResp.error) { return {success:false, message:`GPT Validation System Error: ${gptResp.error} ${gptResp.details ? JSON.stringify(gptResp.details) : ''}`}; }
    if (gptResp.is_valid === false){ return {success:false, message:`GPT detectó error: ${gptResp.message}`}; }
    if (gptResp.sanitized_trade && typeof gptResp.sanitized_trade === 'object') { tradeData = gptResp.sanitized_trade; } 
    else { if (gptResp.is_valid !== true) { return {success:false, message:`GPT validation issue: Invalid response structure.`};}}

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`El full "${SHEET_NAME}" no s'ha trobat.`);
    const headers = getSheetHeadersArray(); 
    let nextId = 1; 
    const idColumnName = headers.find(h => h.toUpperCase() === "ID") || "ID"; 
    const idColumnIndex = headers.indexOf(idColumnName) + 1;
    if (idColumnIndex > 0 && sheet.getLastRow() > 1) { 
        const ids = sheet.getRange(2, idColumnIndex, sheet.getLastRow() - 1, 1).getValues().flat().map(id => parseInt(id)).filter(id => !isNaN(id)); 
        if (ids.length > 0) nextId = Math.max(...ids) + 1; 
    } 
    tradeData[idColumnName] = nextId; 

    DATE_HEADERS_CONFIG.forEach(dateHeader => { 
        if (tradeData[dateHeader] && typeof tradeData[dateHeader] === 'string') {
            const dateParts = tradeData[dateHeader].split('-'); 
            if (dateParts.length === 3) {
                const parsedDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
                tradeData[dateHeader] = !isNaN(parsedDate) ? parsedDate : null;
            } else { 
                Logger.log(`Malformed date string for ${dateHeader}: ${tradeData[dateHeader]}. Using null.`);
                tradeData[dateHeader] = null;
            }
        } else if (tradeData[dateHeader] instanceof Date && isNaN(tradeData[dateHeader].getTime())) { 
             Logger.log(`Invalid Date object for ${dateHeader}. Using null.`);
             tradeData[dateHeader] = null; 
        }
    });

    const newRow = headers.map(header => { 
        let value = tradeData[header]; 
        if (value === undefined || value === null || String(value).trim() === "") { 
            return (DATE_HEADERS_CONFIG.includes(header) && value === null) ? null : ""; 
        }
        if (NUMERIC_HEADERS_CONFIG.includes(header)) { 
            const numVal = parseFloat(String(value).replace(',', '.')); 
            return isNaN(numVal) ? "" : numVal; 
        }
        if (PERCENT_HEADERS_CONFIG.includes(header)) { 
            const percVal = parseFloat(String(value)); 
            return isNaN(percVal) ? "" : percVal; 
        }
        return value; // Includes AssetType
    });

    sheet.appendRow(newRow);
    SpreadsheetApp.flush(); 
    return { success: true, message: "Operació afegida amb ID: " + nextId, newId: nextId, addedTrade: tradeData };
  } catch (error) { Logger.log(`Error a addTrade: ${error.toString()}\nStack: ${error.stack}`); return { success: false, message: `Error afegint operació: ${error.message}` }; }
}

/**
 * Updates an existing trade.
 */
function updateTrade(tradeDataWithId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    const headers = getSheetHeadersArray();
    const tradeIdToUpdate = parseInt(tradeDataWithId.ID);
    if (isNaN(tradeIdToUpdate)) return { success: false, message: "Invalid trade ID." };
    
    const idColumnName = headers.find(h => h.toUpperCase() === "ID") || "ID";
    const idColumnIndex = headers.indexOf(idColumnName) + 1;
    if (idColumnIndex === 0) return { success: false, message: "ID column not found." };

    const idValuesRange = sheet.getRange(2, idColumnIndex, sheet.getLastRow() - 1, 1);
    const idValues = idValuesRange.getValues().flat();
    let rowIndexToUpdate = -1;
    for (let i = 0; i < idValues.length; i++) { if (parseInt(idValues[i]) === tradeIdToUpdate) { rowIndexToUpdate = i + 2; break; } }
    if (rowIndexToUpdate === -1) return { success: false, message: `Trade ID ${tradeIdToUpdate} not found.` };

    const processedTradeData = { ...tradeDataWithId };
    DATE_HEADERS_CONFIG.forEach(dateHeader => { 
        if (processedTradeData[dateHeader] && typeof processedTradeData[dateHeader] === 'string') {
            const dateParts = processedTradeData[dateHeader].split('-');
            if (dateParts.length === 3) {
              const parsedDate = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
              processedTradeData[dateHeader] = !isNaN(parsedDate) ? parsedDate : null;
            } else { processedTradeData[dateHeader] = null; }
          } else if (processedTradeData[dateHeader] instanceof Date && isNaN(processedTradeData[dateHeader].getTime())) {
            processedTradeData[dateHeader] = null;
          }
    });
    const updatedRowValues = headers.map(header => { 
        let value = processedTradeData[header]; 
        if (value === undefined || value === null || String(value).trim() === "") { return (DATE_HEADERS_CONFIG.includes(header) && value === null) ? null : ""; }
        if (NUMERIC_HEADERS_CONFIG.includes(header)) { const numVal = parseFloat(String(value).replace(',', '.')); return isNaN(numVal) ? "" : numVal; }
        if (PERCENT_HEADERS_CONFIG.includes(header)) { const percVal = parseFloat(String(value)); return isNaN(percVal) ? "" : percVal; }
        return value; // Includes AssetType
    });

    sheet.getRange(rowIndexToUpdate, 1, 1, headers.length).setValues([updatedRowValues]);
    SpreadsheetApp.flush();
    return { success: true, message: `Trade ID ${tradeIdToUpdate} updated.` };
  } catch (error) { Logger.log(`Error a updateTrade: ${error.toString()}\nStack: ${error.stack}`); return { success: false, message: `Error actualitzant operació: ${error.message}` }; }
}

/**
 * Deletes multiple trades from the sheet based on an array of trade IDs.
 */
function bulkDeleteTrades(tradeIdsArray) {
  Logger.log("Attempting to bulk delete trades. IDs: " + JSON.stringify(tradeIdsArray));
  if (!tradeIdsArray || !Array.isArray(tradeIdsArray) || tradeIdsArray.length === 0) {
    return { success: false, message: "No trade IDs provided for deletion.", deletedCount: 0 };
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" not found.`);
    const headers = getSheetHeadersArray();
    const idColumnName = headers.find(h => h.toUpperCase() === "ID") || "ID";
    const idColumnIndex = headers.indexOf(idColumnName); 
    if (idColumnIndex === -1) return { success: false, message: "ID column not found in headers." };

    const dataRange = sheet.getDataRange();
    const allSheetValues = dataRange.getValues(); 
    const idsToDeleteNumeric = tradeIdsArray.map(id => parseInt(id)).filter(id => !isNaN(id));
    const rowsToDeleteSheetNumbers = [];

    for (let i = 1; i < allSheetValues.length; i++) { 
      const currentRowId = parseInt(allSheetValues[i][idColumnIndex]);
      if (!isNaN(currentRowId) && idsToDeleteNumeric.includes(currentRowId)) {
        rowsToDeleteSheetNumbers.push(i + 1); 
      }
    }
    if (rowsToDeleteSheetNumbers.length === 0) return { success: true, message: "No matching trades found.", deletedCount: 0 };
    
    rowsToDeleteSheetNumbers.sort((a, b) => b - a); 
    let deletedCount = 0;
    rowsToDeleteSheetNumbers.forEach(rowIndex => { try { sheet.deleteRow(rowIndex); deletedCount++; } catch (e) { Logger.log(`Failed to delete row ${rowIndex}: ${e}`); }});
    SpreadsheetApp.flush();
    return { success: true, message: `${deletedCount} of ${idsToDeleteNumeric.length} selected trades deleted.`, deletedCount };
  } catch (error) { Logger.log(`Error in bulkDeleteTrades: ${error.toString()}\nStack: ${error.stack}`); return { success: false, message: `Error during bulk deletion: ${error.message}`, deletedCount: 0 }; }
}

/**
 * Imports trades from a CSV data string.
 */
function importTradesFromCSV(csvDataString) {
  Logger.log("Starting CSV Import Process. CSV data length: " + (csvDataString ? csvDataString.length : 0));
  let successfullyImported = 0; let failedImports = 0; const importResults = [];
  let canonicalHeaders = [];
  try {
    if (!csvDataString || csvDataString.trim() === "") {
        return { success: false, message: "CSV data is empty.", totalRowsProcessed: 0, successfullyImported, failedImports, results: importResults };
    }
    if (typeof getSheetHeadersArray !== 'function') throw new Error("getSheetHeadersArray function is not defined.");
    canonicalHeaders = getSheetHeadersArray();
    if (!canonicalHeaders || canonicalHeaders.length === 0) throw new Error("Canonical headers array is empty or invalid.");

    const lines = csvDataString.trim().split(/\r\n|\n|\r/);
    Logger.log("CSV has " + lines.length + " lines (including header).");
    if (lines.length < 2) return { success: false, message: "CSV data has no data rows (only header or less).", totalRowsProcessed: Math.max(0, lines.length -1), successfullyImported, failedImports, results: importResults };

    const csvHeadersRaw = lines[0].split(',');
    const csvHeaders = csvHeadersRaw.map(h => h.trim());
    Logger.log("CSV Headers: " + JSON.stringify(csvHeaders));
    
    const headerMap = {}; let missingRequiredHeaders = [];
    const requiredCanonicalHeaders = ["SYMBOL", "Date In", "ACTION", "SHARES", "CallPriceUSD"]; 

    canonicalHeaders.forEach(canonHeader => {
        const csvHeaderIndex = csvHeaders.findIndex(csvH => csvH.toUpperCase() === canonHeader.toUpperCase());
        if (csvHeaderIndex !== -1) { 
            headerMap[canonHeader] = csvHeaderIndex; 
        } else if (requiredCanonicalHeaders.map(rh => rh.toUpperCase()).includes(canonHeader.toUpperCase())) {
            missingRequiredHeaders.push(canonHeader);
        } else {
            Logger.log(`Optional canonical header "${canonHeader}" not found in CSV headers.`);
        }
    });

    if (missingRequiredHeaders.length > 0) {
         return { success: false, message: `CSV import failed. Essential headers missing: ${missingRequiredHeaders.join(', ')}. Expected: ${requiredCanonicalHeaders.join(', ')} among others.`, totalRowsProcessed: lines.length -1, successfullyImported, failedImports, results: importResults };
    }

    for (let i = 1; i < lines.length; i++) {
      const lineContent = lines[i].trim();
      if (lineContent === '') { Logger.log(`Row ${i+1}: Skipped empty line.`); continue; }
      
      // This basic split won't handle commas inside properly quoted fields.
      // A more robust CSV parser would be needed for complex CSVs.
      const dataCells = lineContent.split(','); 
      
      const tradeData = {}; let rowHasAnyData = false;
      canonicalHeaders.forEach(canonHeader => {
        const csvIndex = headerMap[canonHeader]; // csvIndex is the index in the CSV's header row
        if (csvIndex !== undefined && csvIndex < dataCells.length) {
          let value = dataCells[csvIndex].trim();
          tradeData[canonHeader] = value; 
          if (value !== "") rowHasAnyData = true;
        } else { 
          tradeData[canonHeader] = null; 
        }
      });
      
      if (!rowHasAnyData) { Logger.log(`Row ${i+1}: Skipped as it appears empty after mapping. CSV: ${lineContent}`); continue; }
      if (tradeData.hasOwnProperty('ID')) delete tradeData.ID; 

      Logger.log(`Processing CSV Row ${i+1} for import: ${JSON.stringify(tradeData)}`);
      const result = addTrade(tradeData); 
      
      importResults.push({ rowNumber: i + 1, csvRowData: lineContent.substring(0, 200), status: result.success ? "Success" : "Failed", message: result.message, newId: result.success ? result.newId : null });
      if (result.success) successfullyImported++; else failedImports++;
    }
    Logger.log(`CSV Import Finished. Success: ${successfullyImported}, Failed: ${failedImports}`);
    return { success: failedImports === 0, message: `Imported ${successfullyImported} of ${lines.length - 1} data rows. Failures: ${failedImports}.`, totalRowsProcessed: lines.length - 1, successfullyImported, failedImports, results: importResults };
  } catch (e) { Logger.log(`Error during CSV Import: ${e.toString()}\nStack: ${e.stack}`); return { success: false, message: `Error during CSV import: ${e.message}`, totalRowsProcessed: 0, successfullyImported: 0, failedImports: 0, results: importResults }; }
}

// --- ANALYTICS & STATS FUNCTIONS ---

/**
 * Calculates stats for the last 60 trades (by count).
 */
function getLast60OpsStats() {
  const tradesResult = getTrades();
  if (tradesResult.error) { Logger.log(`Error in getLast60OpsStats: ${tradesResult.error}`); return { error: tradesResult.error }; }
  const allTrades = Array.isArray(tradesResult) ? tradesResult : [];
  if (allTrades.length === 0) return { mean: 0, stdev: 0, maxDD: 0, error: "No trades data" };

  const headers = getSheetHeadersArray();
  const dateInCol = headers.find(h => h.toUpperCase() === "DATE IN") || "Date In";
  const pnlCol = headers.find(h => h.toUpperCase() === "RESULTEUR") || "ResultEUR";

  const sortedTrades = allTrades
    .filter(t => t[dateInCol] && t[pnlCol] !== undefined)
    .sort((a, b) => new Date(b[dateInCol]) - new Date(a[dateInCol])); 

  const last60Trades = sortedTrades.slice(0, 60);
  if (last60Trades.length === 0) return { mean: 0, stdev: 0, maxDD: 0 };

  let equity = 0;
  const equityCurveLast60 = last60Trades.slice().reverse().map(trade => { equity += (parseFloat(trade[pnlCol]) || 0); return equity; });
  if (equityCurveLast60.length === 0) return { mean: 0, stdev: 0, maxDD: 0 };

  const pnlValuesLast60 = last60Trades.map(t => parseFloat(t[pnlCol]) || 0).reverse(); 
  const mean  = pnlValuesLast60.reduce((a,b) => a + b, 0) / pnlValuesLast60.length;
  const stdev = Math.sqrt(pnlValuesLast60.map(x => (x - mean) ** 2).reduce((a,b) => a + b, 0) / pnlValuesLast60.length);
  let maxDD = 0; let peak = -Infinity;
  // Calculate peak for first element, then proceed
  if (equityCurveLast60.length > 0) peak = equityCurveLast60[0];
  for (let i = 0; i < equityCurveLast60.length; i++) { peak = Math.max(peak, equityCurveLast60[i]); maxDD = Math.max(maxDD, peak - equityCurveLast60[i]); }
  
  return {mean: mean || 0, stdev: stdev || 0, maxDD: maxDD || 0};
}

/**
 * Calculates equity curve for a given timeframe and overall max drawdown.
 */
function getEquityCurveAndOverallMaxDD(timeframe) {
  const tradesResult = getTrades();
  if (tradesResult.error) { return { error: tradesResult.error, overallMaxDrawdown: 0, equityCurveForChart: [] }; }
  const allTrades = Array.isArray(tradesResult) ? tradesResult : [];

  const headers = getSheetHeadersArray();
  const dateInCol = headers.find(h => h.toUpperCase() === "DATE IN") || "Date In";
  const pnlCol = headers.find(h => h.toUpperCase() === "RESULTEUR") || "ResultEUR";

  let overallMaxDrawdown = 0;
  if (allTrades.length > 0) {
    let overallEquity = 0;
    const fullEquityCurve = allTrades
      .filter(t => t[dateInCol] && typeof t[pnlCol] === 'number') // Ensure PnL is a number for equity calc
      .sort((a, b) => new Date(a[dateInCol]) - new Date(b[dateInCol]))
      .map(trade => { overallEquity += trade[pnlCol]; return overallEquity; }); // pnlCol is already parsed number from getTrades
    if (fullEquityCurve.length > 0) {
      let peak = fullEquityCurve[0]; 
      overallMaxDrawdown = Math.max(0, peak - fullEquityCurve[0]); // If first point is negative, DD is from 0 or itself
      if (peak < 0) peak = 0; // Drawdown relative to zero if portfolio starts negative

      for (let i = 0; i < fullEquityCurve.length; i++) { 
        peak = Math.max(peak, fullEquityCurve[i]); 
        overallMaxDrawdown = Math.max(overallMaxDrawdown, peak - fullEquityCurve[i]); 
      }
    }
  }
  
  let timeFramedTrades = allTrades.filter(t => t[dateInCol] && typeof t[pnlCol] === 'number'); // Start with valid trades
  const today = new Date(); 
  let startDate = new Date(0); // Default to earliest possible if "ALL"

  if (timeframe && timeframe.toUpperCase() !== "ALL") {
    startDate = new Date(); // Re-init for specific timeframes
    switch (timeframe.toUpperCase()) {
      case "1M": startDate.setDate(today.getDate() - 30); break;
      case "3M": startDate.setDate(today.getDate() - 90); break;
      case "YTD": startDate = new Date(today.getFullYear(), 0, 1); break;
      default: Logger.log("Invalid timeframe: " + timeframe + ". Using ALL."); startDate = new Date(0); // Fallback to ALL
    }
    startDate.setHours(0,0,0,0);
    const inclusiveEndDate = new Date(today); // Ensure today is included
    inclusiveEndDate.setHours(23,59,59,999);

    timeFramedTrades = timeFramedTrades.filter(t => { 
        const tradeDate = new Date(t[dateInCol]); 
        return !isNaN(tradeDate) && tradeDate >= startDate && tradeDate <= inclusiveEndDate; 
    });
  }
  
  let currentEquityForTimeframe = 0;
  const equityCurveForChart = timeFramedTrades
    .sort((a, b) => new Date(a[dateInCol]) - new Date(b[dateInCol])) // Chronological
    .map(trade => { currentEquityForTimeframe += trade[pnlCol]; return { date: trade[dateInCol], equity: currentEquityForTimeframe }; });

  return { overallMaxDrawdown: overallMaxDrawdown || 0, equityCurveForChart: equityCurveForChart };
}

/**
 * Gets comprehensive analytics data, supporting filters and generating takeaways.
 */
function getAnalyticsData(filters) {
  Logger.log("Iniciant getAnalyticsData amb filtres: " + JSON.stringify(filters));
  try {
    const headers = getSheetHeadersArray();
    let tradesForAnalyticsResult = getTrades();
    if (tradesForAnalyticsResult.error) return { error: `Could not retrieve trades: ${tradesForAnalyticsResult.error}` };
    let tradesForAnalytics = tradesForAnalyticsResult;
    if (!Array.isArray(tradesForAnalytics)) tradesForAnalytics = [];

    // Apply global filters
    if (filters && typeof filters === 'object') {
      const dateInColName = headers.find(h => h.toUpperCase() === "DATE IN") || "Date In";
      const symbolColName = headers.find(h => h.toUpperCase() === "SYMBOL") || "SYMBOL";
      const strategyColName = headers.find(h => h.toUpperCase() === "STRATEGY") || "STRATEGY";
      const assetTypeColName = headers.find(h => h.toUpperCase() === "ASSETTYPE") || "AssetType";

      if (filters.dateStart) { try { const sd = new Date(filters.dateStart); sd.setHours(0,0,0,0); tradesForAnalytics = tradesForAnalytics.filter(t => t[dateInColName] && new Date(t[dateInColName]) >= sd); } catch(e){Logger.log("Filter dateStart error: "+e)} }
      if (filters.dateEnd) { try { const ed = new Date(filters.dateEnd); ed.setHours(23,59,59,999); tradesForAnalytics = tradesForAnalytics.filter(t => t[dateInColName] && new Date(t[dateInColName]) <= ed); } catch(e){Logger.log("Filter dateEnd error: "+e)} }
      
      if (filters.symbols && Array.isArray(filters.symbols) && filters.symbols.length > 0) {
        const symbolFiltersNormalized = filters.symbols.map(s => String(s).toUpperCase().trim()).filter(s => s !== "");
        if (symbolFiltersNormalized.length > 0) tradesForAnalytics = tradesForAnalytics.filter(t => String(t[symbolColName] || "").toUpperCase().trim() !== "" && symbolFiltersNormalized.includes(String(t[symbolColName] || "").toUpperCase().trim()));
      } else if (filters.symbol && typeof filters.symbol === 'string' && filters.symbol.trim() !== "") { 
        const symbolFilterNormalized = filters.symbol.toUpperCase().trim();
        tradesForAnalytics = tradesForAnalytics.filter(t => (String(t[symbolColName] || "").toUpperCase().trim()) === symbolFilterNormalized );
      }
      if (filters.strategy && filters.strategy.trim() !== "") tradesForAnalytics = tradesForAnalytics.filter(t => t[strategyColName] === filters.strategy);
      if (filters.assetType && filters.assetType.trim() !== "") tradesForAnalytics = tradesForAnalytics.filter(t => t[assetTypeColName] === filters.assetType);
      Logger.log("Number of trades after filtering: " + tradesForAnalytics.length);
    }

    const defaultAnalyticsReturn = {
        years: [], revenueByMonth: {}, investmentByMonth: {}, tradesByMonth: {}, performanceByPrice: {},
        pnlByDayOfWeek: DAYS_OF_WEEK_ANALYTICS.map(day => ({ day: day, pnl: 0, totalTrades: 0, winningTrades: 0 })),
        winRateByDayOfWeek: DAYS_OF_WEEK_ANALYTICS.map(day => ({ day: day, winRate: 0, totalTrades: 0 })),
        pnlByMonthCategory: MONTH_NAMES_ANALYTICS.map(month => ({ month: month, pnl: 0, totalTrades: 0, winningTrades: 0 })),
        winRateByMonthCategory: MONTH_NAMES_ANALYTICS.map(month => ({ month: month, winRate: 0, totalTrades: 0 })),
        pnlByTopSymbols: [], investmentByAssetType: [], takeaways: generateEmptyTakeaways(),
        totalNetProfitEUR: 0, totalInvestedEUR: 0, totalGrossProfitEUR: 0, totalGrossLossEUR: 0,
        winningTradesCount: 0, losingTradesCount: 0, winRate: 0, averageWinAmount: 0, averageLossAmount: 0,
        profitFactor: 0, expectancy: 0, averageRiskRewardRatio: 0, averagePayoffRatio: 0, tradesCountForRR: 0,
        filtersApplied: filters
      };
    if (tradesForAnalytics.length === 0) { return defaultAnalyticsReturn; }

    const analytics = {
      years: [], revenueByMonth: {}, investmentByMonth: {}, tradesByMonth: {}, performanceByPrice: {},
      pnlByDayOfWeek: DAYS_OF_WEEK_ANALYTICS.map(day => ({ day: day, pnl: 0, totalTrades: 0, winningTrades: 0 })),
      pnlByMonthCategory: MONTH_NAMES_ANALYTICS.map((monthName) => ({ month: monthName, pnl: 0, totalTrades: 0, winningTrades: 0 })),
      takeaways: {}
    };
    const uniqueYears = new Set();
    const dateInColForAgg = headers.find(h => h.toUpperCase() === "DATE IN") || "Date In"; 
    const resultEurCol = headers.find(h => h.toUpperCase() === "RESULTEUR") || "ResultEUR";
    const investmentEurColForAgg = headers.find(h => h.toUpperCase() === "INVESTMENTEUR") || "InvestmentEUR";
    const callPriceUsdCol = headers.find(h => h.toUpperCase() === "CALLPRICEUSD") || "CallPriceUSD";
    const symbolColForAgg = headers.find(h => h.toUpperCase() === "SYMBOL") || "SYMBOL";
    const assetTypeColForAgg = headers.find(h => h.toUpperCase() === "ASSETTYPE") || "AssetType";

    let totalNetProfitEUR = 0, totalInvestedEUR = 0, totalGrossProfitEUR = 0, totalGrossLossEUR = 0;
    let winningTradesCount = 0, losingTradesCount = 0, sumOfRiskRewardRatios = 0, tradesCountForRR = 0; 
    const symbolPerformance = {}; const investmentByAssetTypeAgg = {};

    tradesForAnalytics.forEach(trade => {
      const dateString = trade[dateInColForAgg]; 
      if (!dateString || typeof dateString !== 'string') return; 
      const dateValue = new Date(dateString); if (isNaN(dateValue.getTime())) return; 
      const year = dateValue.getFullYear(); const monthIndex = dateValue.getMonth(); const dayOfWeek = dateValue.getDay(); 
      uniqueYears.add(year);
      if (!analytics.revenueByMonth[year]) { 
          analytics.revenueByMonth[year] = Array(12).fill(0); 
          analytics.investmentByMonth[year] = Array(12).fill(0);
          analytics.tradesByMonth[year] = Array(12).fill(null).map(() => ({ total: 0, successful: 0 }));
          analytics.performanceByPrice[year] = {};
          PRICE_RANGES_ANALYTICS.forEach(range => { analytics.performanceByPrice[year][range.label] = 0; });
      }
      const resultEUR = trade[resultEurCol]; // Already parsed to number by getTrades
      const investmentEUR = trade[investmentEurColForAgg]; // Already parsed
      const currentSymbol = trade[symbolColForAgg] || "Unknown Symbol";
      const currentAssetType = trade[assetTypeColForAgg] || "Uncategorized";

      analytics.revenueByMonth[year][monthIndex] += resultEUR; 
      analytics.investmentByMonth[year][monthIndex] += investmentEUR;
      analytics.tradesByMonth[year][monthIndex].total++; if (resultEUR > 0) analytics.tradesByMonth[year][monthIndex].successful++;
      analytics.pnlByDayOfWeek[dayOfWeek].pnl += resultEUR; analytics.pnlByDayOfWeek[dayOfWeek].totalTrades++; if (resultEUR > 0) analytics.pnlByDayOfWeek[dayOfWeek].winningTrades++;
      analytics.pnlByMonthCategory[monthIndex].pnl += resultEUR; analytics.pnlByMonthCategory[monthIndex].totalTrades++; if (resultEUR > 0) analytics.pnlByMonthCategory[monthIndex].winningTrades++;
      if (investmentEUR > 0) { if (!investmentByAssetTypeAgg[currentAssetType]) investmentByAssetTypeAgg[currentAssetType] = 0; investmentByAssetTypeAgg[currentAssetType] += investmentEUR; }
      if (!symbolPerformance[currentSymbol]) symbolPerformance[currentSymbol] = { pnl: 0, totalTrades: 0 }; symbolPerformance[currentSymbol].pnl += resultEUR; symbolPerformance[currentSymbol].totalTrades++;
      totalNetProfitEUR += resultEUR; if (resultEUR > 0) { totalGrossProfitEUR += resultEUR; winningTradesCount++; } else if (resultEUR < 0) { totalGrossLossEUR += Math.abs(resultEUR); losingTradesCount++; }
      if (investmentEUR > 0 && !isNaN(resultEUR)) { totalInvestedEUR += investmentEUR; sumOfRiskRewardRatios += (resultEUR / investmentEUR); tradesCountForRR++; }
      const callPriceUSD = trade[callPriceUsdCol]; // Already parsed
      if (!isNaN(callPriceUSD)) { for (const range of PRICE_RANGES_ANALYTICS) { if (callPriceUSD >= range.min && callPriceUSD <= range.max) { analytics.performanceByPrice[year][range.label] += resultEUR; break; }}}}
    );

    analytics.winRateByDayOfWeek = analytics.pnlByDayOfWeek.map(d => ({ day: d.day, winRate: d.totalTrades > 0 ? d.winningTrades / d.totalTrades : 0, totalTrades: d.totalTrades }));
    analytics.winRateByMonthCategory = analytics.pnlByMonthCategory.map((d, i) => ({ month: MONTH_NAMES_ANALYTICS[i], winRate: d.totalTrades > 0 ? d.winningTrades / d.totalTrades : 0, totalTrades: d.totalTrades }));
    const symbolPerfArray = Object.keys(symbolPerformance).map(s => ({ symbol: s, pnl: symbolPerformance[s].pnl, totalTrades: symbolPerformance[s].totalTrades }));
    symbolPerfArray.sort((a, b) => Math.abs(b.pnl) - Math.abs(a.pnl));
    analytics.pnlByTopSymbols = symbolPerfArray.slice(0, 15);
    analytics.investmentByAssetType = Object.keys(investmentByAssetTypeAgg).map(at => ({ assetType: at, totalInvestment: investmentByAssetTypeAgg[at] })).sort((a, b) => b.totalInvestment - a.totalInvestment);
    
    analytics.years = Array.from(uniqueYears).sort((a,b) => a-b);
    analytics.totalNetProfitEUR = totalNetProfitEUR; analytics.totalInvestedEUR = totalInvestedEUR; 
    analytics.totalGrossProfitEUR = totalGrossProfitEUR; analytics.totalGrossLossEUR = totalGrossLossEUR;
    analytics.winningTradesCount = winningTradesCount; analytics.losingTradesCount = losingTradesCount;
    const totalClosedTrades = winningTradesCount + losingTradesCount;
    analytics.winRate = totalClosedTrades > 0 ? winningTradesCount / totalClosedTrades : 0;
    analytics.averageWinAmount = winningTradesCount > 0 ? totalGrossProfitEUR / winningTradesCount : 0;
    analytics.averageLossAmount = losingTradesCount > 0 ? totalGrossLossEUR / losingTradesCount : 0; 
    analytics.profitFactor = totalGrossLossEUR > 0 ? totalGrossProfitEUR / totalGrossLossEUR : (totalGrossProfitEUR > 0 ? Infinity : 0);
    const lossRate = totalClosedTrades > 0 ? losingTradesCount / totalClosedTrades : 0;
    analytics.expectancy = (analytics.winRate * analytics.averageWinAmount) - (lossRate * analytics.averageLossAmount);
    analytics.averageRiskRewardRatio = tradesCountForRR > 0 ? sumOfRiskRewardRatios / tradesCountForRR : 0;
    analytics.averagePayoffRatio = analytics.averageLossAmount > 0 ? analytics.averageWinAmount / analytics.averageLossAmount : (analytics.averageWinAmount > 0 ? Infinity : 0);
    analytics.tradesCountForRR = tradesCountForRR;
    analytics.filtersApplied = filters;

    // Generate Takeaways
    analytics.takeaways = {};
    if (analytics.pnlByDayOfWeek && analytics.pnlByDayOfWeek.some(d => d.totalTrades > 0)) { let bestDay = { pnl: -Infinity, day: '', trades: 0 }; analytics.pnlByDayOfWeek.forEach(d => { if (d.totalTrades > 0 && d.pnl > bestDay.pnl) bestDay = d; }); if(bestDay.trades > 0) analytics.takeaways.pnlByDayOfWeek = `Most profitable day: ${bestDay.day} (${formatCurrencyForTakeaway(bestDay.pnl)} from ${bestDay.trades} trades).`; } else { analytics.takeaways.pnlByDayOfWeek = "Not enough data.";}
    if (analytics.winRateByDayOfWeek && analytics.winRateByDayOfWeek.some(d => d.totalTrades > 0)) { let bestDay = { winRate: -1, day: '', trades: 0 }; analytics.winRateByDayOfWeek.forEach(d => { if (d.totalTrades > 0 && d.winRate > bestDay.winRate) bestDay = d; }); if(bestDay.trades > 0) analytics.takeaways.winRateByDayOfWeek = `Highest win rate on ${bestDay.day} (${(bestDay.winRate * 100).toFixed(0)}% from ${bestDay.trades} trades).`; } else { analytics.takeaways.winRateByDayOfWeek = "Not enough data.";}
    if (analytics.pnlByMonthCategory && analytics.pnlByMonthCategory.some(m => m.totalTrades > 0)) { let bestM = {pnl: -Infinity, month: '', trades: 0}; analytics.pnlByMonthCategory.forEach(m => { if(m.totalTrades > 0 && m.pnl > bestM.pnl) bestM = m;}); if(bestM.trades > 0) analytics.takeaways.pnlByMonthCategory = `Most profitable month: ${bestM.month} (${formatCurrencyForTakeaway(bestM.pnl)} from ${bestM.trades} trades).`; } else {analytics.takeaways.pnlByMonthCategory = "Not enough data.";}
    if (analytics.winRateByMonthCategory && analytics.winRateByMonthCategory.some(m => m.totalTrades > 0)) { let bestM = {winRate: -1, month: '', trades: 0}; analytics.winRateByMonthCategory.forEach(m => {if(m.totalTrades > 0 && m.winRate > bestM.winRate) bestM = m;}); if(bestM.trades > 0) analytics.takeaways.winRateByMonthCategory = `Highest win rate month: ${bestM.month} (${(bestM.winRate*100).toFixed(0)}% from ${bestM.trades} trades).`; } else {analytics.takeaways.winRateByMonthCategory = "Not enough data.";}
    if (analytics.pnlByTopSymbols && analytics.pnlByTopSymbols.length > 0) { const topS = analytics.pnlByTopSymbols[0]; analytics.takeaways.pnlByTopSymbols = `Top symbol ${topS.symbol} P&L: ${formatCurrencyForTakeaway(topS.pnl)} (${topS.totalTrades} trades).`;} else {analytics.takeaways.pnlByTopSymbols = "Not enough data.";}
    if (analytics.years && analytics.years.length > 0) { const latestYear = analytics.years[analytics.years.length - 1]; if (analytics.performanceByPrice && analytics.performanceByPrice[latestYear]) { const ranges = analytics.performanceByPrice[latestYear]; let bestR = {pnl: -Infinity, label:''}; Object.keys(ranges).forEach(lbl => {if(ranges[lbl] > bestR.pnl) bestR = {pnl:ranges[lbl], label:lbl};}); if(bestR.label && bestR.pnl > 0) analytics.takeaways.performanceByPrice = `For ${latestYear}, best price range: ${bestR.label} (${formatCurrencyForTakeaway(bestR.pnl)}).`; else {analytics.takeaways.performanceByPrice = "Not enough data for price range.";} } } else {analytics.takeaways.performanceByPrice = "Not enough data.";}
    
    return analytics;
  } catch (err) { Logger.log(`Catastrophic error in getAnalyticsData: ${err.toString()}\nStack: ${err.stack}`); const errorReturn = { ...defaultAnalyticsReturn, error: `Error intern del servidor: ${err.message}`, filtersApplied: filters }; errorReturn.takeaways = generateEmptyTakeaways(); return errorReturn; }
}


/**
 * Retrieves historical performance averages.
 */
function getHistoricalPerformanceAverages() {
  try {
    const analytics = getAnalyticsData(null); 
    if (analytics.error) return { error: `Could not fetch analytics: ${analytics.error}` };
    let historicalAvgInvestmentPerOp = 0;
    if (analytics.tradesCountForRR > 0 && analytics.totalInvestedEUR > 0) { historicalAvgInvestmentPerOp = analytics.totalInvestedEUR / analytics.tradesCountForRR; }
    let uniqueMonthsWithTrades = 0; const monthTracker = {}; 
    if (analytics.years && analytics.tradesByMonth) { analytics.years.forEach(year => { if(analytics.tradesByMonth[year]) { analytics.tradesByMonth[year].forEach((monthData, monthIndex) => { if (monthData.total > 0) { const yearMonth = `${year}-${monthIndex}`; if (!monthTracker[yearMonth]) { monthTracker[yearMonth] = true; uniqueMonthsWithTrades++; } } }); } }); }
    const historicalAvgTradesPerMonth = uniqueMonthsWithTrades > 0 ? (analytics.winningTradesCount + analytics.losingTradesCount) / uniqueMonthsWithTrades : 0;
    let suggestedRoiTargetPercent = 0; if (analytics.winningTradesCount > 0 && historicalAvgInvestmentPerOp > 0) { suggestedRoiTargetPercent = analytics.averageWinAmount / historicalAvgInvestmentPerOp; }
    let suggestedLossPerFailedOpPercent = 0; if (analytics.losingTradesCount > 0 && historicalAvgInvestmentPerOp > 0) { suggestedLossPerFailedOpPercent = analytics.averageLossAmount / historicalAvgInvestmentPerOp; }
    return { success: true, historicalWinRate: analytics.winRate || 0, historicalAvgInvestmentPerOp, historicalAvgTradesPerMonth, suggestedRoiTargetPercent, suggestedLossPerFailedOpPercent };
  } catch (e) { return { error: `Server error calculating historical averages: ${e.message}` }; }
}

/**
 * Gathers all data required for the main dashboard.
 */
function getDashboardData(timeframeForChart) { 
  try {
    const analyticsDataResult = getAnalyticsData(null); 
    const equityCurveData = getEquityCurveAndOverallMaxDD(timeframeForChart || "ALL"); 
    const goalsResult = loadUserGoals();
    const tradesResult = getTrades(); 
    if (analyticsDataResult.error || equityCurveData.error || goalsResult.error || tradesResult.error) { return { error: "Failed to gather all dashboard data."}; }
    const analyticsData = analyticsDataResult; const goals = goalsResult; const trades = Array.isArray(tradesResult) ? tradesResult : [];
    const initialCapital = parseFloat(goals.capitalForTrading) || 0; const totalNetPnl = parseFloat(analyticsData.totalNetProfitEUR) || 0; const portfolioValue = initialCapital + totalNetPnl;
    const kpis = { portfolioValue, totalNetProfitEUR: totalNetPnl, overallWinRate: analyticsData.winRate || 0, profitFactor: analyticsData.profitFactor || 0, maxDrawdown: equityCurveData.overallMaxDrawdown };
    const finalCapitalGoal = parseFloat(goals.finalCapitalGoal) || 0; let goalProgressPercent = 0;
    if (finalCapitalGoal > 0 && portfolioValue >= initialCapital ) { const totalGainNeeded = finalCapitalGoal - initialCapital; const currentGain = portfolioValue - initialCapital; if (totalGainNeeded > 0) { goalProgressPercent = Math.max(0, Math.min(100, (currentGain / totalGainNeeded) * 100)); } else if (portfolioValue >= finalCapitalGoal) { goalProgressPercent = 100; } } else if (portfolioValue >= finalCapitalGoal && finalCapitalGoal > 0) { goalProgressPercent = 100; }
    const goalProgress = { currentValue: portfolioValue, targetValue: finalCapitalGoal, initialCapital, percentage: goalProgressPercent };
    const headersForSort = getSheetHeadersArray(); const idHeader = headersForSort.find(h => h.toUpperCase() === "ID") || "ID"; const dateInHeader = headersForSort.find(h => h.toUpperCase() === "DATE IN") || "Date In"; const dateOutHeader = headersForSort.find(h => h.toUpperCase() === "DATE OUT") || "Date Out"; const symbolHeader = headersForSort.find(h => h.toUpperCase() === "SYMBOL") || "SYMBOL"; const resultEURHeader = headersForSort.find(h => h.toUpperCase() === "RESULTEUR") || "ResultEUR";
    const recentTrades = trades.sort((a, b) => { const dA = new Date(a[dateInHeader] || 0); const dB = new Date(b[dateInHeader] || 0); if (dB < dA) return -1; if (dB > dA) return 1; return (parseInt(b[idHeader]) || 0) - (parseInt(a[idHeader]) || 0); }).slice(0, 5).map(t => ({ id: t[idHeader], symbol: t[symbolHeader], pnlEUR: t[resultEURHeader], date: t[dateOutHeader] || t[dateInHeader], status: t[dateOutHeader] ? 'Closed' : 'Open' }));
    return { success: true, kpis, equityCurveForChart: equityCurveData.equityCurveForChart, goalProgress, recentTrades };
  } catch (e) { return { error: `Server error fetching dashboard data: ${e.message}` }; }
}

/**
 * Calculates the historical drawdown series for the entire portfolio.
 */
function getDrawdownAnalysisData() {
  try {
    const equityCurveData = getEquityCurveAndOverallMaxDD("ALL"); 
    if (equityCurveData.error) throw new Error(`Failed to get equity curve: ${equityCurveData.error}`);
    const fullEquityCurvePoints = equityCurveData.equityCurveForChart;
    if (!fullEquityCurvePoints || fullEquityCurvePoints.length === 0) return { success: true, drawdownSeries: [], message: "No equity data." };

    const refinedDrawdownSeries = []; let highWaterMark = 0; 
    for (const point of fullEquityCurvePoints) {
        highWaterMark = Math.max(highWaterMark, point.equity);
        let currentDrawdown = 0;
        if (highWaterMark > 0) { currentDrawdown = (highWaterMark - point.equity) / highWaterMark; } 
        else if (highWaterMark === 0 && point.equity < 0) { currentDrawdown = 1; }
        refinedDrawdownSeries.push({ date: point.date, drawdownPercentage: Math.max(0, currentDrawdown) });
    }
    return { success: true, drawdownSeries: refinedDrawdownSeries };
  } catch (e) { return { success: false, error: `Server error calculating drawdown: ${e.message}`, drawdownSeries: [] }; }
}

// --- GOALS & PROJECTIONS ---
function saveUserGoals(goalsData) { try { USER_PROPERTIES.setProperty(GOALS_KEY, JSON.stringify(goalsData)); return { success: true, message: "Objectius desats." }; } catch (e) { return { success: false, message: `Error desant objectius: ${e.message}` }; }}
function loadUserGoals() { try { const gs = USER_PROPERTIES.getProperty(GOALS_KEY); if (gs) return JSON.parse(gs); return { capitalForTrading: 10000, roiTarget: 0.20, lossPerFailedOpPercent: 0.05, taxesPerOp: 10, avgInvestmentPerOp: 5000, monthlyOpsTarget: 30, successfulOpsPercentTarget: 0.50, finalCapitalGoal: 2000000 }; } catch (e) { return { error: `Error carregant objectius: ${e.message}` }; }}
function calculateProjections(goalsData) { try { /* ... Full projection logic ... */ return { /* projection results */ }; } catch (e) { return { error: `Error calculant projeccions: ${e.message}` }; }}
function generateSimulatedOperations(goalsData, successfulOps, failedOps, netGainSuccess, netLossFail) { /* ... Full simulation logic ... */ return [ /* simulated ops array */ ]; }
// Note: For brevity, calculateProjections and generateSimulatedOperations' detailed logic is implied.

// --- SETTINGS: STRATEGIES & ASSET TYPES ---
function getUserDefinedStrategies() { try { const s = USER_PROPERTIES.getProperty(USER_STRATEGIES_KEY); return s ? JSON.parse(s) : []; } catch (e) { return { error: `Failed to load strategies: ${e.message}` }; }}
function saveUserDefinedStrategies(strategiesArray) { try { USER_PROPERTIES.setProperty(USER_STRATEGIES_KEY, JSON.stringify(strategiesArray)); return true; } catch (e) { return false; }}
function addStrategyDefinition(strategyName) { if (!strategyName || String(strategyName).trim() === "") return { success: false, message: "Name empty." }; const name = String(strategyName).trim(); let sListR = getUserDefinedStrategies(); if (sListR.error) return { success: false, message: sListR.error }; let sList = sListR; if (sList.some(s => s.toLowerCase() === name.toLowerCase())) return { success: false, message: `Strategy "${name}" already exists.` }; sList.push(name); sList.sort((a,b) => a.toLowerCase().localeCompare(b.toLowerCase())); return saveUserDefinedStrategies(sList) ? { success: true, message: `Strategy "${name}" added.`, strategies: sList } : { success: false, message: "Failed to save." }; }
function deleteStrategyDefinition(strategyName) { if (!strategyName || String(strategyName).trim() === "") return { success: false, message: "Name empty." }; const name = String(strategyName).trim(); let sListR = getUserDefinedStrategies(); if (sListR.error) return { success: false, message: sListR.error }; let sList = sListR; const initialLen = sList.length; sList = sList.filter(s => s.toLowerCase() !== name.toLowerCase()); if (sList.length === initialLen) return { success: false, message: `Strategy "${name}" not found.` }; return saveUserDefinedStrategies(sList) ? { success: true, message: `Strategy "${name}" deleted.`, strategies: sList } : { success: false, message: "Failed to save." }; }

function getUserDefinedAssetTypes() { try { const s = USER_PROPERTIES.getProperty(USER_ASSET_TYPES_KEY); return s ? JSON.parse(s) : []; } catch (e) { return { error: `Failed to load asset types: ${e.message}` }; }}
function saveUserDefinedAssetTypes(assetTypesArray) { try { USER_PROPERTIES.setProperty(USER_ASSET_TYPES_KEY, JSON.stringify(assetTypesArray)); return true; } catch (e) { return false; }}
function addAssetTypeDefinition(assetTypeName) { if (!assetTypeName || String(assetTypeName).trim() === "") return { success: false, message: "Name empty." }; const name = String(assetTypeName).trim(); let atListR = getUserDefinedAssetTypes(); if (atListR.error) return { success: false, message: atListR.error }; let atList = atListR; if (atList.some(s => s.toLowerCase() === name.toLowerCase())) return { success: false, message: `Asset type "${name}" already exists.` }; atList.push(name); atList.sort((a,b) => a.toLowerCase().localeCompare(b.toLowerCase())); return saveUserDefinedAssetTypes(atList) ? { success: true, message: `Asset type "${name}" added.`, assetTypes: atList } : { success: false, message: "Failed to save." }; }
function deleteAssetTypeDefinition(assetTypeName) { if (!assetTypeName || String(assetTypeName).trim() === "") return { success: false, message: "Name empty." }; const name = String(assetTypeName).trim(); let atListR = getUserDefinedAssetTypes(); if (atListR.error) return { success: false, message: atListR.error }; let atList = atListR; const initialLen = atList.length; atList = atList.filter(s => s.toLowerCase() !== name.toLowerCase()); if (atList.length === initialLen) return { success: false, message: `Asset type "${name}" not found.` }; return saveUserDefinedAssetTypes(atList) ? { success: true, message: `Asset type "${name}" deleted.`, assetTypes: atList } : { success: false, message: "Failed to save." }; }

/**
 * Returns the canonical headers for the CSV template.
 */
function getCSVTemplateHeaders() {
  try {
    if (typeof getSheetHeadersArray !== 'function') throw new Error("getSheetHeadersArray undefined.");
    const headers = getSheetHeadersArray();
    if (!headers || headers.length === 0) throw new Error("Could not retrieve valid template headers.");
    return headers; 
  } catch (e) { return { error: "Could not retrieve template headers: " + e.message }; }
}

// --- HELPER FUNCTIONS ---
function formatCurrencyForTakeaway(value) {
    if (isNaN(parseFloat(value))) return 'N/A';
    return "€" + parseFloat(value).toFixed(2);
}
function generateEmptyTakeaways() {
    return {
        pnlByDayOfWeek: "Not enough data for analysis.", winRateByDayOfWeek: "Not enough data for analysis.",
        pnlByMonthCategory: "Not enough data for analysis.", winRateByMonthCategory: "Not enough data for analysis.",
        pnlByTopSymbols: "Not enough data for analysis.", performanceByPrice: "Not enough data for analysis."
    };
}