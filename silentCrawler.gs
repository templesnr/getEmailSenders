/**
 * SILENT GMAIL CRAWLER - Self-Triggering Background Analysis
 * Runs in short sessions, saves progress to sheets, and reschedules itself until done.
 */

// --------------------------- Globals --------------------------- //
let CRAWLER_RUN_MS = 4.5 * 60 * 1000;              // allowed run time per session (ms)  // global run-time constant

// --------------------------- startSilentCrawler --------------------------- //
function startSilentCrawler() {                                                              // public starter for the crawler
  clearCrawlerTriggers();                                                                    // remove any existing triggers
  initializeCrawlerState(true);                                                              // initialize sheets & props for fresh start
  runCrawlerSession();                                                                       // execute first session immediately
}                                                                                            // end startSilentCrawler

// --------------------------- stopSilentCrawler --------------------------- //
function stopSilentCrawler() {                                                               // public stop for the crawler
  clearCrawlerTriggers();                                                                    // delete scheduled triggers
  updateCrawlerStatus('STOPPED', 'Stopped by user');                                        // mark status stopped
}                                                                                            // end stopSilentCrawler

// --------------------------- initializeCrawlerState --------------------------- //
function initializeCrawlerState() {                                                                     // Initialize or reset crawler progress & status sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();                                            // Get active spreadsheet
  
  // --- Progress sheet setup ---
  let progressSheet = spreadsheet.getSheetByName("Crawler Progress");                                    // Try to get existing "Crawler Progress" sheet
  if (!progressSheet) {                                                                                  // If it doesn't exist
    progressSheet = spreadsheet.insertSheet("Crawler Progress");                                         // Create it
  }                                                                                                      // End if
  progressSheet.clear();                                                                                 // Clear any existing content
  progressSheet.getRange("A1:D1").setValues([["Page Token", "Threads Processed", "Current Batch", "Batch Position"]]); // Set headers
  progressSheet.getRange("A2:D2").setValues([["START", 0, "", 0]]);                                      // Initialize with start values
  
  // --- Status sheet setup ---
  let statusSheet = spreadsheet.getSheetByName("Crawler Status");                                        // Try to get existing "Crawler Status" sheet
  if (!statusSheet) {                                                                                    // If it doesn't exist
    statusSheet = spreadsheet.insertSheet("Crawler Status");                                             // Create it
  }                                                                                                      // End if
  statusSheet.clear();                                                                                   // Clear any existing content
  statusSheet.getRange("A1:B1").setValues([["Crawler Status", ""]]);                                     // Sheet title
  statusSheet.getRange("A2:B12").setValues([                                                             // Fill layout rows
    ["Status", ""],                                                                                      // Row 2
    ["Started", ""],                                                                                     // Row 3
    ["Last Update", ""],                                                                                 // Row 4
    ["Total Threads in Gmail", ""],                                                                      // Row 5
    ["Triggers Active", ""],                                                                             // Row 6
    ["Total Threads Processed", ""],                                                                     // Row 7
    ["Total Emails Found", ""],                                                                          // Row 8
    ["Unique Senders Found", ""],                                                                        // Row 9
    ["Estimated Progress", ""],                                                                          // Row 10
    ["Current Phase", ""],                                                                               // Row 11
    ["Next Run", ""]                                                                                     // Row 12
  ]);                                                                                                    // End setValues
  
  // --- Initial counts ---
  const totalThreads = countAllGmailThreadsFast();                                                       // Get total Gmail threads
  statusSheet.getRange("B5").setValue(totalThreads);                                                      // Save total threads
  const triggerCount = countActiveCrawlerTriggers();                                                     // Get active crawler triggers
  statusSheet.getRange("B6").setValue(triggerCount);                                                      // Save trigger count
  
  updateCrawlerStatus("INITIALIZED", "Crawler state initialized for fresh start.");                       // Set initial status
  console.log("init");                                                                                    // Debug log
}                                                                                                         // End initializeCrawlerState

// --------------------------- runCrawlerSession --------------------------- //
function runCrawlerSession() {                                                                     // main crawler loop to process Gmail threads
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();                                      // get active spreadsheet
  const progressSheet = spreadsheet.getSheetByName("Crawler Progress");                           // open progress sheet
  let [pageTokenRaw, threadsProcessedRaw, currentBatchJSONRaw, batchPositionRaw] =
    progressSheet.getRange("A2:D2").getValues()[0];                                                // read saved state row
  const pageToken = pageTokenRaw == null ? "" : String(pageTokenRaw);                              // normalize pageToken
  let threadsProcessed = parseInt(threadsProcessedRaw, 10) || 0;                                   // parse threads processed
  let currentBatchJSON = currentBatchJSONRaw || "";                                                // normalize stored batch JSON
  let batchPosition = parseInt(batchPositionRaw, 10) || 0;                                         // parse batch position
  const startTime = Date.now();                                                                    // record start time of session
  let sessionThreads = 0;                                                                          // counter for session threads

  if (pageToken === "START") {                                                                     // if fresh start
    console.log("fresh");                                                                          // log fresh start
    threadsProcessed = 0;                                                                          // reset count
    currentBatchJSON = "";                                                                         // clear batch
    batchPosition = 0;                                                                             // reset position
    saveProgress("", 0, [], 0);                                                                    // save reset state
    const totalThreads = countAllGmailThreadsFast();                                               // get total threads once
    updateCrawlerStatus('RUNNING', `Total threads: ${totalThreads}`);                              // update status
  } else {
    console.log("resume");                                                                         // log resume
  }

  let threads = currentBatchJSON ? JSON.parse(currentBatchJSON) : [];                              // load batch from state
  if (threads.length === 0) {                                                                      // if batch empty
    const searchResults = GmailApp.search('', 0, 100);                                             // fetch next 100 threads
    threads = searchResults.map(thread => thread.getId());                                         // store thread IDs
    batchPosition = 0;                                                                             // reset position
  }

  while (batchPosition < threads.length) {                                                         // loop through batch
    const threadId = threads[batchPosition];                                                       // get thread ID
    const thread = GmailApp.getThreadById(threadId);                                               // fetch thread
    processThreadForCrawler(thread);                                                               // process thread data
    batchPosition++;                                                                               // move to next thread
    threadsProcessed++;                                                                            // increment total count
    sessionThreads++;                                                                              // increment session count
    if (timeLimitReached(startTime)) break;                                                        // stop if time limit hit
  }

  saveProgress(pageToken, threadsProcessed, threads.slice(batchPosition), batchPosition);          // save state

  const statusSheet = spreadsheet.getSheetByName("Crawler Status");                                // open status sheet
  const totalThreads = statusSheet.getRange("B5").getValue();                                      // read total threads
  const progressPercent = totalThreads > 0
    ? Math.min((threadsProcessed / totalThreads) * 100, 100).toFixed(1)                            // calc %
    : 0;                                                                                           // default 0
  statusSheet.getRange("B10").setValue(`${progressPercent}%`);                                     // update progress %
  console.log(`progress: ${progressPercent}%`);                                                    // log progress %

  scheduleNextCrawlerRun(2 * 60 * 1000);                                                           // schedule next run in 2 mins
}                                                                                                  // end runCrawlerSession
// --------------------------- saveProgress --------------------------- //
function saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition) {                  // write progress to sheet (A2:D2)
  const ss = SpreadsheetApp.getActiveSpreadsheet();                                                // get spreadsheet
  const ps = ss.getSheetByName('Crawler Progress');                                                // open progress sheet
  const json = (Array.isArray(currentBatch) && currentBatch.length) ? JSON.stringify(currentBatch) : ''; // serialize minimal batch
  ps.getRange('A2').setValue(pageToken || '');                                                     // column A = pageToken
  ps.getRange('B2').setValue(threadsProcessed || 0);                                               // column B = threadsProcessed
  ps.getRange('C2').setValue(json);                                                                // column C = currentBatch JSON
  ps.getRange('D2').setValue(batchPosition || 0);                                                  // column D = batchPosition
  console.log('progress saved');                                                                   // short log
}                                                                                                  // end saveProgress

// --------------------------- loadCrawlerProgress --------------------------- //
function loadCrawlerProgress() {                                                                 // read progress from sheet and return normalized object
  const ss = SpreadsheetApp.getActiveSpreadsheet();                                              // get spreadsheet
  const ps = ss.getSheetByName('Crawler Progress');                                             // open progress sheet
  const row = ps.getRange('A2:D2').getValues()[0];                                               // read A2:D2
  const pageTokenRaw = row[0];                                                                   // raw token value
  const threadsProcessedRaw = row[1];                                                            // raw threads processed
  const batchJsonRaw = row[2];                                                                   // raw batch JSON string
  const batchPositionRaw = row[3];                                                               // raw batch position
  const pageToken = pageTokenRaw == null ? '' : String(pageTokenRaw);                            // normalize token
  const threadsProcessed = parseInt(threadsProcessedRaw, 10) || 0;                               // normalize processed count
  let currentBatch = [];                                                                         // placeholder for parsed batch
  try {                                                                                          // try parse if present
    currentBatch = batchJsonRaw ? JSON.parse(batchJsonRaw) : [];                                 // parse JSON to array or empty
  } catch (e) {                                                                                  // if parse fails
    currentBatch = [];                                                                           // fall back to empty
  }                                                                                              // end try/catch
  const batchPosition = parseInt(batchPositionRaw, 10) || 0;                                     // normalize position
  return { pageToken, threadsProcessed, currentBatch, batchPosition };                            // return state object
}                                                                                               // end loadCrawlerProgress

// --------------------------- processThreadForCrawler --------------------------- //
function processThreadForCrawler(threadStub, keeperEmails, sendersMap) {                          // process a single thread id (uses Gmail API)
  let emailsFound = 0;                                                                           // counter of emails found in this thread
  try {                                                                                          // wrap Gmail call in try
    const thread = Gmail.Users.Threads.get('me', threadStub.id, {                               // get thread metadata/messages
      format: 'metadata',
      metadataHeaders: ['From','Date']
    });                                                                                          // end get
    const messages = thread.messages || [];                                                      // message list
    for (let i = 0; i < messages.length; i++) {                                                  // iterate messages
      const msg = messages[i];                                                                   // this message
      const headers = msg.payload.headers || [];                                                 // headers array
      const fromH = headers.find(h => h.name === 'From');                                       // locate From header
      const dateH = headers.find(h => h.name === 'Date');                                       // locate Date header
      if (!fromH) continue;                                                                      // skip if no From
      const parsed = parseSenderInfo(fromH.value);                                              // parse name/email
      const email = parsed.email ? parsed.email.toLowerCase() : '';                             // normalized email
      if (!email || keeperEmails.indexOf(email) !== -1) continue;                                // skip keepers or invalid
      emailsFound++;                                                                             // count one email
      const baseEmail = email.includes('+') ? (email.split('+')[0] + '@' + email.split('@')[1]) : email; // base email
      const messageDate = dateH ? new Date(dateH.value) : new Date(parseInt(msg.internalDate));   // get date
      if (!sendersMap.has(baseEmail)) {                                                          // if new sender
        sendersMap.set(baseEmail, { primaryName: parsed.name || baseEmail, email: baseEmail, date: messageDate, count: 1 }); // add
      } else {                                                                                   // existing sender
        const ex = sendersMap.get(baseEmail);                                                    // existing record
        ex.count = (ex.count || 0) + 1;                                                          // increment count
        if (messageDate > ex.date) { ex.date = messageDate; ex.primaryName = parsed.name || ex.primaryName; } // update date & name
        sendersMap.set(baseEmail, ex);                                                           // save back
      }                                                                                          // end else
    }                                                                                            // end for messages
  } catch (e) {                                                                                  // on error
    console.warn('thread get err');                                                              // short warn
  }                                                                                              // end try/catch
  return { emailsFound };                                                                        // return summary
}                                                                                               // end processThreadForCrawler

// --------------------------- parseSenderInfo --------------------------- //
function parseSenderInfo(fromValue) {                                                            // simple parser for "Name <email>" style From header
  if (!fromValue) return { name: '', email: '' };                                                // guard empty
  const m = fromValue.match(/^(.*?)[\s]*<(.+?)>$/);                                              // regex match name+email
  if (m) return { name: m[1].trim(), email: m[2].trim() };                                        // return parsed parts
  return { name: fromValue.trim(), email: fromValue.trim() };                                     // fallback single token
}                                                                                               // end parseSenderInfo

// --------------------------- getKeeperEmails --------------------------- //
function getKeeperEmails() {                                                                     // load keeper emails from "Keepers" sheet
  const out = [];                                                                                // array to return
  try {                                                                                          // try reading sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const sheet = ss.getSheetByName('Keepers');                                                  // open keepers sheet
    if (!sheet) return out;                                                                      // return empty if none
    const data = sheet.getRange(2,1, sheet.getLastRow() - 1, 1).getValues() || [];               // read keeper rows
    for (let i = 0; i < data.length; i++) {                                                      // loop rows
      const v = data[i][0];                                                                      // cell value
      if (v && typeof v === 'string') out.push(v.trim().toLowerCase());                          // push normalized
    }                                                                                            // end loop
  } catch (e) {                                                                                  // on error
    console.warn('keepers err');                                                                  // short warn
  }                                                                                              // end try/catch
  return out;                                                                                    // return keeper emails
}                                                                                               // end getKeeperEmails

// --------------------------- loadSendersMapFromSheet --------------------------- //
function loadSendersMapFromSheet() {                                                             // helper to load senders into a Map
  const map = new Map();                                                                         // map to return
  try {                                                                                          // try block
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const sheet = ss.getSheetByName('Senders');                                                  // open senders sheet
    if (!sheet || sheet.getLastRow() < 2) return map;                                            // nothing to load
    const rows = sheet.getRange(2,1, sheet.getLastRow() - 1, 4).getValues();                     // read all senders rows
    for (let i = 0; i < rows.length; i++) {                                                       // iterate rows
      const r = rows[i];                                                                         // row
      const email = (r[1] || '').toString().toLowerCase();                                       // normalize email
      if (!email) continue;                                                                      // skip blanks
      map.set(email, { primaryName: r[0], email: email, date: new Date(r[2]), count: r[3] || 0 }); // set map entry
    }                                                                                            // end for
  } catch (e) {                                                                                  // on error
    console.warn('load senders err');                                                            // short warn
  }                                                                                              // end try/catch
  return map;                                                                                    // return map
}                                                                                               // end loadSendersMapFromSheet

// --------------------------- persistSendersMapToSheet --------------------------- //
function persistSendersMapToSheet(map) {                                                         // write Map contents to Senders sheet (replace)
  try {                                                                                          // try write
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const sheet = ss.getSheetByName('Senders');                                                  // get sheet
    if (!sheet) return;                                                                          // nothing to do without sheet
    const arr = Array.from(map.values()).sort((a,b) => b.date - a.date).map(s => [s.primaryName, s.email, s.date, s.count]); // prepare rows
    sheet.clearContents();                                                                        // clear old content
    sheet.getRange(1,1,1,4).setValues([['Name','Email','Most Recent Email Date','Count of Emails']]); // header
    if (arr.length > 0) sheet.getRange(2,1,arr.length,4).setValues(arr);                         // write rows
  } catch (e) {                                                                                  // on error
    console.warn('persist senders err');                                                         // warn
  }                                                                                              // end try/catch
}                                                                                               // end persistSendersMapToSheet

// --------------------------- getTotalEmailsCount --------------------------- //
function getTotalEmailsCount() {                                                                 // compute total emails from Senders sheet
  try {                                                                                          // try read
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const sheet = ss.getSheetByName('Senders');                                                  // open senders
    if (!sheet || sheet.getLastRow() < 2) return 0;                                              // none present
    const data = sheet.getRange(2,4,sheet.getLastRow()-1,1).getValues();                         // read counts column
    return data.reduce((s,r) => s + (parseInt(r[0],10) || 0), 0);                                 // sum counts
  } catch (e) {                                                                                  // on error
    return 0;                                                                                    // fallback 0
  }                                                                                              // end try/catch
}                                                                                               // end getTotalEmailsCount

// --------------------------- getUniqueSendersCount --------------------------- //
function getUniqueSendersCount() {                                                               // get unique senders count from sheet
  try {                                                                                          // try read
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const sheet = ss.getSheetByName('Senders');                                                  // open sheet
    if (!sheet) return 0;                                                                        // none present
    return Math.max(0, sheet.getLastRow() - 1);                                                  // rows minus header
  } catch (e) {                                                                                  // on error
    return 0;                                                                                    // fallback 0
  }                                                                                              // end try/catch
}                                                                                               // end getUniqueSendersCount

// --------------------------- timeLimitReached --------------------------- //
function timeLimitReached(startTime) {                                                           // returns true if runtime is nearly exhausted
  const BUFFER = 10000;                                                                          // leave buffer (ms) to save & schedule
  return (Date.now() - startTime) >= (CRAWLER_RUN_MS - BUFFER);                                   // compare elapsed to allowed time minus buffer
}                                                                                               // end timeLimitReached

// --------------------------- updateCrawlerProgress --------------------------- //
function updateCrawlerProgress(threadsProcessed, emailsFound, sendersCount) {                     // update counters in status sheet
  try {                                                                                          // try update
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const s = ss.getSheetByName('Crawler Status');                                               // open status sheet
    if (!s) return;                                                                              // nothing to update
    s.getRange('B4').setValue(new Date());                                                       // Last Update cell
    if (threadsProcessed != null) s.getRange('B5').setValue(threadsProcessed);                   // Total Threads Processed
    if (emailsFound != null) s.getRange('B6').setValue(emailsFound);                             // Total Emails Found
    if (sendersCount != null) s.getRange('B7').setValue(sendersCount);                           // Unique Senders Found
    PropertiesService.getScriptProperties().setProperty('totalThreadsProcessed', String(threadsProcessed || 0)); // persist total
  } catch (e) {                                                                                  // on error
    console.warn('update progress err');                                                          // warn
  }                                                                                              // end try/catch
}                                                                                               // end updateCrawlerProgress

// --------------------------- updateCrawlerStatus --------------------------- //
function updateCrawlerStatus(status, message) {                                                  // write status & message to sheet & props
  try {                                                                                          // try update
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const s = ss.getSheetByName('Crawler Status');                                               // open status
    if (s) {                                                                                     // if sheet exists
      s.getRange('B2').setValue(status);                                                         // set Status cell
      s.getRange('B4').setValue(new Date());                                                     // set Last Update
      s.getRange('B8').setValue(message || '');                                                  // set Current Phase
    }                                                                                            // end if
    PropertiesService.getScriptProperties().setProperty('crawlerStatus', status || '');          // persist to script properties
  } catch (e) {                                                                                  // on error
    console.warn('status err');                                                                  // warn
  }                                                                                              // end try/catch
}                                                                                               // end updateCrawlerStatus

// --------------------------- scheduleNextCrawlerRun --------------------------- //
function scheduleNextCrawlerRun(delayMs = 3 * 60 * 1000) {                                       // schedule next run with optional delay
  try {                                                                                          // try scheduling
    clearCrawlerTriggers();                                                                      // clear existing crawler triggers
    const nextTime = new Date(Date.now() + delayMs);                                             // compute next run time
    ScriptApp.newTrigger('runCrawlerSession')                                                    // create new trigger for runCrawlerSession
      .timeBased()                                                                                // make it a time-based trigger
      .at(nextTime)                                                                               // set trigger to run at computed time
      .create();                                                                                  // create the trigger
    updateCrawlerStatus('SCHEDULED', `Next run: ${nextTime.toLocaleString()}`);                   // update status with planned run
  } catch (e) {                                                                                  // on error
    console.warn(`schedule err: ${e.message}`);                                                   // log detailed warning message
  }                                                                                              // end try/catch
}                                                                                                // end scheduleNextCrawlerRun

// --------------------------- clearCrawlerTriggers --------------------------- //
function clearCrawlerTriggers() {                                                                // remove all crawler-related triggers
  const triggers = ScriptApp.getProjectTriggers();                                               // get all project triggers
  for (const trigger of triggers) {                                                              // loop through each trigger
    if (trigger.getHandlerFunction() === 'runCrawlerSession') {                                  // if trigger is for runCrawlerSession
      ScriptApp.deleteTrigger(trigger);                                                          // delete this trigger
    }                                                                                            // end if
  }                                                                                              // end loop
}                                                                                                // end clearCrawlerTriggers


// --------------------------- completeCrawlerAnalysis --------------------------- //
function completeCrawlerAnalysis() {                                                             // finalize run and clear triggers
  updateCrawlerStatus('COMPLETE', 'Analysis complete');                                          // set status to COMPLETE
  clearCrawlerTriggers();                                                                        // remove scheduled triggers
  console.log('complete');                                                                       // short complete log
}                                                                                               // end completeCrawlerAnalysis

// --------------------------- checkCrawlerStatus --------------------------- //
function checkCrawlerStatus() {                                                                  // quick helper to show status in logs
  try {                                                                                          // try show
    const ss = SpreadsheetApp.getActiveSpreadsheet();                                            // get spreadsheet
    const s = ss.getSheetByName('Crawler Status');                                               // open status sheet
    const status = s ? s.getRange('B2').getValue() : 'NOT FOUND';                                // read status cell
    console.log(`status: ${status}`);                                                            // log status
  } catch (e) {                                                                                  // on error
    console.warn('check status err');                                                             // warn
  }                                                                                              // end try/catch
}                                                                                               // end checkCrawlerStatus

// --------------------------- Helpers / Backwards-compat --------------------------- //
// These helpers are thin wrappers to match previous API expectations and to keep small-surface changes.
// persistSendersMapToSheet and collect/load helpers above are used by runCrawlerSession.

// If you want me to change any naming, formatting, or make comments aligned to a fixed column width,
// tell me the column number and I'll regenerate with perfect alignment.
