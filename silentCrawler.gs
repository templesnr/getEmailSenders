/**
 * SILENT GMAIL CRAWLER - Master Script
 * ------------------------------------
 * Runs in repeated short sessions, saving state to sheets and rescheduling itself
 * until the entire Gmail account has been processed. Tracks senders, name variations,
 * totals, progress %, trigger count, and preserves progress across runs.
 *
 * Version 14: Merged Crawler Progress into the Crawler Status sheet for consolidation.
 */

// --------------------------- Globals --------------------------- //
const CRAWLER_RUN_MS = 4.5 * 60 * 1000;                                      // Define the maximum runtime for a single session (4.5 minutes in milliseconds).

// --------------------------- UI & Menu Functions --------------------------- //
/**
 * @summary Adds a custom menu to the spreadsheet UI when the file is opened.
 * @description This function runs automatically when the spreadsheet is opened and creates
 * the "Email Management" menu with options to start, stop, and check the crawler.
 */
function onOpen() {                                                         // Function that runs automatically on spreadsheet open.
  const ui = SpreadsheetApp.getUi();                                        // Get the UI object to interact with the spreadsheet's interface.
  ui.createMenu('Email Management')                                         // Create a new top-level menu named "Email Management".
    .addItem('🤖 Start Silent Crawler (Fresh Start)', 'startSilentCrawlerFresh') // Add a menu item to start a fresh crawl.
    .addItem('▶ Resume Silent Crawler', 'startSilentCrawlerResume')         // Add a menu item to resume a crawl.
    .addItem('⏹️ Stop Silent Crawler', 'stopSilentCrawler')                  // Add a menu item to manually stop the crawler.
    .addItem('📊 Check Crawler Status', 'checkCrawlerStatus')                // Add a menu item to log the current status.
    .addToUi();                                                             // Add the created menu to the spreadsheet's UI.
}                                                                           // End of onOpen function.

/**
 * @summary Menu helper function to start the crawler with a fresh start.
 */
function startSilentCrawlerFresh() {                                        // Function definition for the "Fresh Start" menu item.
  startSilentCrawler(true);                                                 // Call the main start function, passing 'true' to signal a fresh start.
}                                                                           // End of startSilentCrawlerFresh function.

/**
 * @summary Menu helper function to resume the crawler.
 */
function startSilentCrawlerResume() {                                       // Function definition for the "Resume" menu item.
  startSilentCrawler(false);                                                // Call the main start function, passing 'false' to signal a resume.
}                                                                           // End of startSilentCrawlerResume function.


// --------------------------- Main Controller Functions --------------------------- //
/**
 * @summary Main entry point to start or resume the crawler.
 * @description This function orchestrates the start of the process, whether it's a
 * fresh start or a resume. It clears old triggers and prepares the sheets.
 * @param {boolean} fresh - If true, starts a fresh crawl. If false, resumes.
 */
function startSilentCrawler(fresh) {                                        // Function definition for the main start process.
  console.info(`Starting silent crawler. Fresh start: ${fresh}`);          // Log whether this is a fresh start or a resume.
  clearCrawlerTriggers();                                                   // Clear any existing triggers to ensure a clean run.
  prepareCrawlerSheets(fresh);                                              // Prepare the spreadsheet, passing the 'fresh' flag.
}                                                                           // End of startSilentCrawler function.

/**
 * @summary Stops crawler operation by clearing triggers and updating status.
 */
function stopSilentCrawler() {                                              // Function definition to manually stop the crawl.
  clearCrawlerTriggers();                                                   // Call the function to remove any scheduled triggers.
  updateCrawlerStatus('STOPPED', 'Stopped by user');                        // Update the status display to show that the process was stopped.
  console.info('Silent crawler stopped');                                   // Log a confirmation message that the crawler has been stopped.
}                                                                           // End of stopSilentCrawler function.

/**
 * @summary Counts total Gmail threads and then starts the first crawler session.
 */
function countThreadsAndStart() {                                           // Function definition to count threads and start the main process.
  console.info('Counting Gmail threads...');                                // Log that the thread counting is starting.
  const totalThreads = countAllGmailThreadsFast();                          // Call the fast counting function and store the result.
  updateCrawlerStatus('INITIALIZED', 'Thread count complete');              // Update the status to show that counting is complete.
  console.info('Thread count complete:', totalThreads);                     // Log the final count of total threads.
  runCrawlerSession();                                                      // Call the main processing function to start the crawl.
}                                                                           // End of countThreadsAndStart function.

/**
 * @summary Executes a single crawling session, acting as the main controller.
 * @description This function coordinates the entire process for a single 4.5-minute run.
 * It calls sheet service functions to load state, API service functions to fetch data,
 * core logic functions to process data, and trigger functions to schedule the next run.
 */
function runCrawlerSession() {                                              // Function definition for a single processing session.
  const sessionStart = Date.now();                                          // Record the start time of the session.
  console.info('Starting crawler session');                                 // Log that a new session is beginning.

  // --- Initialization ---
  const { pageToken: loadedToken, threadsProcessed: loadedThreads, currentBatch: loadedBatch, batchPosition: loadedPos } = loadCrawlerProgress(); // Load the saved progress from the sheet.
  let pageToken = (loadedToken === 'START') ? '' : (loadedToken || '');     // Set the page token, handling the special 'START' case.
  let threadsProcessed = (loadedToken === 'START') ? 0 : (loadedThreads || 0); // Set the count of processed threads, resetting if it's a 'START'.
  let currentBatch = Array.isArray(loadedBatch) ? loadedBatch : [];         // Ensure the current batch is a valid array.
  let batchPosition = (loadedToken === 'START') ? 0 : (loadedPos || 0);     // Set the position within the batch, resetting if it's a 'START'.

  let sendersMap = loadSendersMapFromSheet();                               // Load the map of senders from the "Senders" sheet.
  let nameVarMap = loadNameVariationsMapFromSheet();                        // Load the map of name variations from its sheet.
  
  if (loadedToken === 'START') {                                            // If the loaded token is 'START'...
    saveProgress('', 0, [], 0);                                             // ...persist a clean state to the sheet.
    console.info('Fresh START detected — clearing in-memory maps and starting from first page.'); // ...log that a fresh start is happening.
    sendersMap.clear();                                                     // ...clear the in-memory senders map to prevent double counting.
    nameVarMap.clear();                                                     // ...clear the in-memory name variations map.
  }                                                                         // End of if block.

  console.info('Resuming from processed:', threadsProcessed, 'pageToken:', pageToken || '<none>'); // Log the point from which the session is resuming.

  try {                                                                     // Start a try block to catch any unexpected errors.
    let processedThisSession = 0;                                           // Initialize a counter for threads processed in this session only.
    let morePages = true;                                                   // A flag to indicate if there are more pages of threads to fetch.

    while (!timeLimitReached(sessionStart) && morePages) {                  // Loop as long as the time limit has not been reached and there are more pages.
      let batchToIterate = [];                                              // Initialize an array to hold the batch of threads to be processed.

      if (Array.isArray(currentBatch) && currentBatch.length > 0 && batchPosition < currentBatch.length) { // If there's an unfinished batch from a previous run...
        batchToIterate = currentBatch;                                      // ...use that batch.
      } else {                                                              // Otherwise...
        const response = fetchThreadBatch(pageToken);                       // ...call the API service to fetch a new batch.
        if (!response) {                                                    // If the fetch failed...
          scheduleNextCrawlerRun();                                         // ...schedule a retry and exit.
          return;                                                           // ...stop this session.
        }                                                                   // End of if block.

        const threads = response.threads || [];                             // Get the threads from the response, or an empty array if none.
        const nextToken = response.nextToken || '';                         // Get the next page token, or an empty string if none.

        if (!nextToken) {                                                   // If there is no next token from the API...
          morePages = false;                                                // ...this is the last page, so set the flag to stop the main loop after this batch.
          console.info('Last page of threads received. The script will finalize after this batch.'); // Log this event.
        }                                                                   // End of if block.

        if (threads.length === 0) {                                         // If the API returns an empty batch of threads...
          console.info('No more threads to process — finalizing now.');     // ...log that the process is complete.
          completeCrawlerAnalysis();                                        // ...run the final completion function.
          return;                                                           // ...and exit immediately.
        }                                                                   // End of if block.

        currentBatch = threads.map(t => ({ id: t.id, snippet: t.snippet || '', historyId: t.historyId || '' })); // Create a compact representation of the batch.
        batchPosition = 0;                                                  // Reset the position within the new batch to the beginning.
        pageToken = nextToken;                                              // Store the next page token for the next fetch.
        saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition); // Save the new batch and token information.
        batchToIterate = currentBatch;                                      // Set the new batch as the one to be iterated over.
        console.info('Fetched', batchToIterate.length, 'threads; nextPageToken exists?', !!pageToken); // Log the results of the fetch.
      }                                                                     // End of else block.

      // Process one thread from the batch
      if (batchPosition < batchToIterate.length) {                          // If the current position is valid within the batch...
        const stub = batchToIterate[batchPosition];                         // ...get the thread information at the current position.
        const threadDetails = fetchThreadDetails(stub.id);                  // ...call the API service to get the full thread details.
        if (threadDetails) {                                                // If the details were fetched successfully...
          processThreadForCrawler(threadDetails, sendersMap, nameVarMap);   // ...call the core logic function to process the data.
        }                                                                   // End of if block.
        batchPosition++;                                                    // ...increment the position within the batch.
        threadsProcessed++;                                                 // ...increment the total number of threads processed.
        processedThisSession++;                                             // ...increment the number of threads processed in this session.
      } else {                                                              // If the position is out of range...
        console.warn('Batch position out of range — resetting batch');      // ...log a warning.
        currentBatch = [];                                                  // ...reset the batch.
        batchPosition = 0;                                                  // ...reset the position.
      }                                                                     // End of else block.

      // Periodic persist + dashboard updates
      if (processedThisSession > 0 && processedThisSession % 100 === 0) {   // Every 100 threads processed in this session...
        saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition); // ...save the current progress.
        persistSendersMapToSheet(sendersMap);                               // ...save the senders map to the sheet.
        persistNameVariationsMapToSheet(nameVarMap);                        // ...save the name variations map to the sheet.
        updateCrawlerProgress(threadsProcessed, getTotalEmailsCount(), getUniqueSendersCount()); // ...update the dashboard counters.
        updateTriggerCountInStatus();                                       // ...update the trigger count on the dashboard.
        console.info('BATCH PROGRESS: processed', threadsProcessed, '(session', processedThisSession, ')'); // ...log a progress update.
      }                                                                     // End of if block.
    }                                                                       // End of while loop.

    // --- Session End Logic ---
    saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition); // Save the final state for this run.
    persistSendersMapToSheet(sendersMap);                                   // Persist the senders map.
    persistNameVariationsMapToSheet(nameVarMap);                            // Persist the name variations map.
    updateCrawlerProgress(threadsProcessed, getTotalEmailsCount(), getUniqueSendersCount()); // Update the dashboard counters.
    updateTriggerCountInStatus();                                           // Update the trigger count.

    if (morePages) {                                                        // If the `morePages` flag is still true, we stopped due to time.
      scheduleNextCrawlerRun();                                             // Schedule the next session to continue the work.
      console.info('Session paused — processed so far =', threadsProcessed);// Log that the session has paused.
    } else {                                                                // If `morePages` is false, the last batch was processed.
      console.info('Final batch processed. Completing analysis.');          // Log that we are finalizing.
      completeCrawlerAnalysis();                                            // Run the completion routine.
    }                                                                       // End of if/else block.

  } catch (err) {                                                           // If an unexpected error occurs in the main block...
    console.error('Unexpected error in runCrawlerSession:', err, err.stack); // ...log the full error with its stack trace.
    saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition); // ...attempt to save the current progress.
    scheduleNextCrawlerRun(5 * 60 * 1000);                                  // ...schedule a retry after a 5-minute delay.
  }                                                                         // End of try-catch block.
}                                                                           // End of runCrawlerSession function.

/**
 * @summary Finalizes the crawler run.
 * @description This function is called when the entire inbox has been processed. It ensures
 * all data is saved, moves keepers from the Senders list, updates the status to "COMPLETE",
 * and clears any remaining triggers to stop the process.
 */
function completeCrawlerAnalysis() {                                            // Function definition for completing the analysis.
  console.info('Completing crawler analysis');                                  // Log that the process is complete.
  persistSendersMapToSheet(loadSendersMapFromSheet());                          // Do a final save of the senders map.
  persistNameVariationsMapToSheet(loadNameVariationsMapFromSheet());            // Do a final save of the name variations map.
  
  moveKeepersToEnd();                                                           // Move specified senders to the keepers sheet.

  updateCrawlerProgress(                                                        // Update the dashboard with the final counts.
    parseInt(PropertiesService.getScriptProperties().getProperty('totalThreadsProcessed'), 10) || 0, // Get the final thread count.
    getTotalEmailsCount(),                                                      // Get the final email count.
    getUniqueSendersCount()                                                     // Get the final sender count.
  );                                                                            // End of updateCrawlerProgress call.
  updateCrawlerStatus('COMPLETE', 'Analysis complete');                         // Set the final status to "COMPLETE".
  clearCrawlerTriggers();                                                       // Clear any remaining triggers.
}                                                                               // End of completeCrawlerAnalysis function.


// --------------------------- Core Logic Functions --------------------------- //
/**
 * @summary Processes a single email thread to extract sender information for thread and email counts.
 * @description This function takes a thread object, identifies the unique senders within that
 * thread, and correctly increments both the overall thread count (once per sender per thread)
 * and the email count (for every message from that sender).
 * @param {object} thread - The full thread object from the Gmail API.
 * @param {Map} sendersMap - The in-memory map of sender data to be updated.
 * @param {Map} nameVarMap - The in-memory map of name variations to be updated.
 */
function processThreadForCrawler(thread, sendersMap, nameVarMap) {              // Function definition to process one thread.
  const messages = thread.messages || [];                                       // Get the messages from the thread, or an empty array.
  const sendersInThisThread = new Set();                                        // Create a temporary Set to track senders already counted for this thread.

  for (const msg of messages) {                                                 // Loop through each message in the thread.
    const headers = msg.payload && msg.payload.headers ? msg.payload.headers : []; // Get the headers for the message.
    const fromH = headers.find(h => h.name.toLowerCase() === 'from');           // Find the 'From' header.
    const dateH = headers.find(h => h.name.toLowerCase() === 'date');           // Find the 'Date' header.
    if (!fromH || !fromH.value) continue;                                       // If there's no 'From' header, skip to the next message.

    const parsed = parseSenderInfo(fromH.value);                                // Parse the 'From' header into name and email.
    const emailRaw = (parsed.email || '').toLowerCase().trim();                 // Normalize the email address (lowercase, trimmed).
    if (!emailRaw) continue;                                                    // If the email is empty, skip it.

    const baseEmail = (emailRaw.includes('+')) ? (emailRaw.split('+')[0] + '@' + emailRaw.split('@')[1]) : emailRaw; // Remove plus-addressing to get the base email.
    const messageDate = dateH && dateH.value ? new Date(dateH.value) : new Date(parseInt(msg.internalDate, 10)); // Get the message date.

    const entry = sendersMap.get(baseEmail);                                    // Check if the sender already exists in the map.
    if (!entry) {                                                               // If the sender is new...
      sendersMap.set(baseEmail, {                                               // ...create a new entry.
        primaryName: parsed.name || baseEmail,                                  // Set the primary name.
        email: baseEmail,                                                       // Set the email.
        date: messageDate,                                                      // Set the last seen date.
        threadCount: 1,                                                         // Initialize thread count to 1.
        emailCount: 1                                                           // Initialize email count to 1.
      });                                                                       // End of new entry object.
    } else {                                                                    // If the sender exists...
      entry.emailCount++;                                                       // ...always increment their email count.
      if (!sendersInThisThread.has(baseEmail)) {                                // ...if this is the first time we see them in THIS thread...
        entry.threadCount++;                                                    // ......then increment their thread count.
      }                                                                         // End of thread count check.
      if (messageDate > entry.date) {                                           // ...if the new message is more recent...
        entry.date = messageDate;                                               // ......update the last seen date.
        entry.primaryName = parsed.name || entry.primaryName;                   // ......and update their primary name.
      }                                                                         // ...end of date check.
    }                                                                           // End of if/else block.
    
    sendersInThisThread.add(baseEmail);                                         // Add the sender to the set for this thread to prevent re-counting threads.

    if (!nameVarMap.has(baseEmail)) nameVarMap.set(baseEmail, new Set());       // Ensure a Set exists for this email in the name variations map.
    const nm = (parsed.name || '').trim();                                      // Get the name variant from this email.
    if (nm && !nm.includes('@')) nameVarMap.get(baseEmail).add(nm);             // If the name is valid, add it to the set of variations.
  }                                                                             // End of message loop.
}                                                                               // End of processThreadForCrawler function.


// --------------------------- Gmail API Service --------------------------- //
/**
 * @summary Quickly counts all Gmail threads using GmailApp.search in batches.
 */
function countAllGmailThreadsFast() {                                       // Function definition for fast thread counting.
  const BATCH_SIZE = 500;                                                   // Set the number of threads to fetch in each batch.
  let totalThreads = 0;                                                     // Initialize a counter for the total number of threads.
  let startIndex = 0;                                                       // Initialize the starting index for the search.
  console.info('Counting all Gmail threads in batches of', BATCH_SIZE);     // Log the start of the counting process.

  while (true) {                                                            // Start an infinite loop (will be broken manually).
    const batch = GmailApp.search('', startIndex, BATCH_SIZE);              // Search for threads, getting a batch of 500.
    totalThreads += batch.length;                                           // Add the number of threads in the current batch to the total.
    startIndex += BATCH_SIZE;                                               // Increment the starting index for the next batch.
    console.info('Batch count:', batch.length, 'Total so far:', totalThreads); // Log the progress for the current batch.
    if (batch.length < BATCH_SIZE) break;                                   // If a batch has fewer than 500 threads, it's the last one, so exit the loop.
  }                                                                         // End of while loop.

  console.info('Total Gmail threads found:', totalThreads);                 // Log the final total.
  const ss = SpreadsheetApp.getActiveSpreadsheet();                         // Get the active spreadsheet.
  const statusSheet = ss.getSheetByName('Crawler Status');                  // Get the status sheet.
  statusSheet.getRange('B6').setValue(totalThreads);                        // Write the final total to the status sheet.
  return totalThreads;                                                      // Return the total count.
}                                                                           // End of countAllGmailThreadsFast function.

/**
 * @summary Fetches a batch of thread IDs from the Gmail API.
 * @param {string} pageToken - The token for the page to fetch.
 * @returns {object|null} The API response object or null on failure.
 */
function fetchThreadBatch(pageToken) {                                      // Function definition for fetching a batch of threads.
  const listOptions = { maxResults: 50 };                                   // Prepare options for the API call, requesting 50 threads.
  if (pageToken) listOptions.pageToken = pageToken;                         // If a page token is provided, add it to the options.
  try {                                                                     // Start a try block for the API call.
    const response = Gmail.Users.Threads.list('me', listOptions);           // Call the Gmail API to get a list of threads.
    return {                                                                // Return a structured response object.
      threads: response.threads || [],                                      // The list of threads.
      nextToken: response.nextPageToken || ''                               // The token for the next page.
    };                                                                      // End of return object.
  } catch (e) {                                                             // If the API call fails...
    console.error('Gmail API error while fetching thread list:', e);        // ...log the error.
    return null;                                                            // ...and return null.
  }                                                                         // End of try-catch block.
}                                                                           // End of fetchThreadBatch function.

/**
 * @summary Fetches the detailed metadata for a single thread, with a retry mechanism.
 * @param {string} threadId - The ID of the thread to fetch.
 * @returns {object|null} The thread object from the API, or null on failure.
 */
function fetchThreadDetails(threadId) {                                     // Function definition for fetching thread details.
  const MAX_RETRIES = 3;                                                    // Define the maximum number of retry attempts.
  for (let i = 0; i < MAX_RETRIES; i++) {                                   // Start a loop for retry attempts.
    try {                                                                   // Start a try block for the API call.
      return Gmail.Users.Threads.get('me', threadId, { format: 'metadata', metadataHeaders: ['From', 'Date'] }); // Get thread metadata and return it.
    } catch (e) {                                                           // If the API call fails...
      if (i < MAX_RETRIES - 1) {                                            // ...and it's not the last attempt...
        console.log(`API error for thread ${threadId}, retrying... (${i + 1}/${MAX_RETRIES})`); // ...log a retry message.
        Utilities.sleep(1000);                                              // ...wait for 1 second before the next attempt.
      } else {                                                              // If it is the last attempt...
        console.warn('Could not process thread ID:', threadId, e);          // ...log the final warning.
      }                                                                     // End of if/else block.
    }                                                                       // End of try-catch block.
  }                                                                         // End of retry loop.
  return null;                                                              // If all retries fail, return null.
}                                                                           // End of fetchThreadDetails function.


// --------------------------- Sheet Service Functions --------------------------- //
/**
 * @summary Ensure required sheets exist, create headers, and optionally schedule the first count.
 */
function prepareCrawlerSheets(fresh = false) {                              // Function definition to set up the spreadsheet.
  const ss = SpreadsheetApp.getActiveSpreadsheet();                         // Get the currently active spreadsheet object.
  console.info('Preparing crawler sheets. Fresh start:', fresh);            // Log that the sheet preparation is starting.

  // ---------- Crawler Status sheet ----------
  let statusSheet = ss.getSheetByName('Crawler Status');                    // Get the sheet named "Crawler Status".
  if (!statusSheet) statusSheet = ss.insertSheet('Crawler Status');         // If the sheet doesn't exist, create it.
  statusSheet.clear();                                                      // Clear all content from the status sheet.
  statusSheet.getRange('A1:B15').setValues([                                // Set up the layout and labels for the status dashboard.
    ['Crawler Status', ''],                                                 // Title for the status dashboard.
    ['Status', ''],                                                         // Label for the current status (e.g., RUNNING, STOPPED).
    ['Started', ''],                                                        // Label for the timestamp when the crawl started.
    ['Last Update', ''],                                                    // Label for the timestamp of the last activity.
    ['Trigger Count', ''],                                                  // Label for the number of active triggers.
    ['Total Threads Found', ''],                                            // Label for the total number of email threads found in the inbox.
    ['Total Threads Processed', 0],                                         // Label for the number of threads processed so far.
    ['Estimated Progress', '=IF(B6=0, 0, B7/B6)'],                          // Use a formula for automatic progress calculation.
    ['Total Emails Found', 0],                                              // Label for the total number of individual emails found.
    ['Unique Senders Found', 0],                                            // Label for the number of unique senders found.
    ['Current Phase', ''],                                                  // Label for a descriptive message about the current operation.
    [],                                                                     // Spacer row.
    ['--- Progress State ---', ''],                                         // Sub-header for progress state.
    ['Page Token', 'START'],                                                // Label and initial value for the page token.
    ['Batch JSON', ''], ['Batch Pos', 0]                                    // Labels and initial values for batch data.
  ]);                                                                       // End of status sheet values.
  statusSheet.getRange('B8').setNumberFormat('0.00%');                      // Format the formula cell as a percentage.
  statusSheet.getRange('A1:A').setFontWeight('bold');                       // Make the first column of the status sheet bold for readability.

  // ---------- Senders sheet ----------
  let sendersSheet = ss.getSheetByName('Senders');                          // Get the sheet named "Senders".
  if (!sendersSheet) sendersSheet = ss.insertSheet('Senders');              // If the sheet doesn't exist, create it.
  if (fresh) sendersSheet.clear();                                          // If it's a fresh start, clear the sheet.
  if (sendersSheet.getLastRow() < 1) {                                      // If the sheet is completely empty...
    sendersSheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Last Seen', 'Thread Count', 'Email Count']]); // ...add the NEW 5-column header row.
  }                                                                         // End of if block.
  sendersSheet.setFrozenRows(1);                                            // Freeze the header row so it's always visible.
  if (!sendersSheet.getFilter()) {                                          // If there is no filter on the sheet...
    try { sendersSheet.getRange(1, 1, sendersSheet.getMaxRows(), 5).createFilter(); } catch (e) {} // ...add one to allow sorting.
  }                                                                         // End of if block.

  // ---------- Name Variations sheet ----------
  let nvSheet = ss.getSheetByName('Name Variations');                       // Get the sheet named "Name Variations".
  if (!nvSheet) nvSheet = ss.insertSheet('Name Variations');                // If the sheet doesn't exist, create it.
  if (fresh) nvSheet.clear();                                               // If it's a fresh start, clear the sheet.
  if (nvSheet.getLastRow() < 1) {                                           // If the sheet is completely empty...
    nvSheet.getRange(1, 1, 1, 2).setValues([['Email', 'Name Variations']]);  // ...add the header row.
  }                                                                         // End of if block.
  nvSheet.setFrozenRows(1);                                                 // Freeze the header row.
  if (!nvSheet.getFilter()) {                                               // If there is no filter on the sheet...
    try { nvSheet.getRange(1, 1, nvSheet.getMaxRows(), 2).createFilter(); } catch (e) {} // ...add one.
  }                                                                         // End of if block.

  // ---------- If fresh → start fast count in separate run ----------
  if (fresh) {                                                              // If this is a fresh start...
    updateCrawlerStatus('INITIALIZING', 'Counting threads...');             // ...update the status to show initialization is in progress.
    ScriptApp.newTrigger('countThreadsAndStart').timeBased().after(2000).create(); // ...create a new trigger to run the counting function after 2 seconds.
    console.info('Scheduling thread count');                                // ...log that the counting has been scheduled.
  }                                                                         // End of if block.
}                                                                           // End of prepareCrawlerSheets function.

/**
 * @summary Persist crawler progress into the "Crawler Status" sheet.
 */
function saveProgress(pageToken, threadsProcessed, currentBatch, batchPosition) { // Function definition to save progress.
  const ss = SpreadsheetApp.getActiveSpreadsheet();                             // Get the active spreadsheet.
  const statusSheet = ss.getSheetByName('Crawler Status');                      // Get the status sheet.
  if (!statusSheet) return;                                                     // If the sheet doesn't exist, do nothing.
  const json = (Array.isArray(currentBatch) && currentBatch.length) ? JSON.stringify(currentBatch) : ''; // Convert the current batch array to a JSON string.
  statusSheet.getRange('B14').setValue(pageToken || '');                        // Write the page token to cell B14.
  statusSheet.getRange('B7').setValue(threadsProcessed || 0);                   // Write the processed count to cell B7.
  statusSheet.getRange('B15').setValue(json);                                   // Write the batch JSON to cell B15.
  statusSheet.getRange('B16').setValue(batchPosition || 0);                     // Write the batch position to cell B16.
}                                                                               // End of saveProgress function.

/**
 * @summary Read and normalize the saved progress from the "Crawler Status" sheet.
 */
function loadCrawlerProgress() {                                                // Function definition to load progress.
  const ss = SpreadsheetApp.getActiveSpreadsheet();                             // Get the active spreadsheet.
  const statusSheet = ss.getSheetByName('Crawler Status');                      // Get the status sheet.
  if (!statusSheet || statusSheet.getLastRow() < 14) return { pageToken: 'START', threadsProcessed: 0, currentBatch: [], batchPosition: 0 }; // If no progress is saved, return a 'START' state.
  
  const pageToken = statusSheet.getRange('B14').getValue();                     // Read the page token from cell B14.
  const threadsProcessed = statusSheet.getRange('B7').getValue();               // Read the processed count from cell B7.
  const batchJson = statusSheet.getRange('B15').getValue();                     // Read the batch JSON from cell B15.
  const batchPosition = statusSheet.getRange('B16').getValue();                 // Read the batch position from cell B16.

  let currentBatch = [];                                                        // Initialize an empty array for the batch.
  try {                                                                         // Start a try block for safe JSON parsing.
    currentBatch = batchJson ? JSON.parse(batchJson) : [];                      // Parse the JSON string back into an array.
  } catch (e) { currentBatch = []; }                                            // If parsing fails, use an empty array.
  
  return {                                                                      // Return the loaded progress as an object.
    pageToken: pageToken == null ? '' : String(pageToken),                      // Ensure page token is a string.
    threadsProcessed: parseInt(threadsProcessed, 10) || 0,                      // Ensure processed count is an integer.
    currentBatch: currentBatch,                                                 // The parsed batch array.
    batchPosition: parseInt(batchPosition, 10) || 0                             // Ensure batch position is an integer.
  };                                                                            // End of return object.
}                                                                               // End of loadCrawlerProgress function.

/**
 * @summary Loads the Senders sheet into an in-memory Map for fast access.
 */
function loadSendersMapFromSheet() {                                            // Function definition to load the senders sheet.
  const map = new Map();                                                        // Create a new Map object.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Senders'); // Get the "Senders" sheet.
    if (!sheet || sheet.getLastRow() < 2) return map;                           // If the sheet is empty, return the empty map.
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();   // Get all the data rows from the sheet (now 5 columns).
    for (const r of rows) {                                                     // Loop through each row.
      const email = (r[1] || '').toString().toLowerCase();                      // Get and normalize the email from the second column.
      if (!email) continue;                                                     // If the email is blank, skip to the next row.
      const dateVal = r[2] ? new Date(r[2]) : new Date(0);                      // Get the date from the third column, converting it to a Date object.
      map.set(email, {                                                          // Add the sender's data to the map.
        primaryName: r[0],                                                      // The sender's name.
        email: email,                                                           // The sender's email.
        date: dateVal,                                                          // The last seen date.
        threadCount: parseInt(r[3], 10) || 0,                                   // The thread count.
        emailCount: parseInt(r[4], 10) || 0                                     // The email count.
      });                                                                       // End of map set.
    }                                                                           // End of loop.
  } catch (e) { console.warn('load senders err', e); }                          // If an error occurs, log a warning.
  return map;                                                                   // Return the populated map.
}                                                                               // End of loadSendersMapFromSheet function.

/**
 * @summary Loads the Name Variations sheet into an in-memory Map.
 */
function loadNameVariationsMapFromSheet() {                                     // Function definition to load the name variations sheet.
  const map = new Map();                                                        // Create a new Map object.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Variations'); // Get the "Name Variations" sheet.
    if (!sheet || sheet.getLastRow() < 2) return map;                           // If the sheet is empty, return the empty map.
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();   // Get all the data rows.
    for (const row of rows) {                                                   // Loop through each row.
      const email = (row[0] || '').toString().toLowerCase();                    // Get and normalize the email from the first column.
      if (!email) continue;                                                     // If the email is blank, skip it.
      const set = new Set((row[1] || '').toString().split(' | ').map(s => s.trim()).filter(Boolean)); // Get the variations string, split it, and create a Set of names.
      map.set(email, set);                                                      // Add the email and its set of names to the map.
    }                                                                           // End of loop.
  } catch (e) { console.warn('load name variations err', e); }                  // If an error occurs, log a warning.
  return map;                                                                   // Return the populated map.
}                                                                               // End of loadNameVariationsMapFromSheet function.

/**
 * @summary Writes the in-memory sendersMap to the "Senders" sheet.
 */
function persistSendersMapToSheet(map) {                                        // Function definition to save the senders map.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Senders'); // Get the "Senders" sheet.
    if (!sheet) return;                                                         // If the sheet doesn't exist, do nothing.
    const arr = Array.from(map.values()).sort((a, b) => (b.date || 0) - (a.date || 0)); // Convert the map values to an array and sort by date descending.
    const rows = arr.map(s => [s.primaryName, s.email, s.date, s.threadCount, s.emailCount]); // Convert the array of objects to a 2D array for writing to the sheet.

    sheet.clearContents();                                                      // Clear all existing data from the sheet.
    sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Last Seen', 'Thread Count', 'Email Count']]); // Write the NEW 5-column header row.
    if (rows.length > 0) {                                                      // If there is data to write...
        sheet.getRange(2, 1, rows.length, 5).setValues(rows);                   // ...write all the rows to the sheet.
    }                                                                           // End of if block.
    sheet.setFrozenRows(1);                                                     // Freeze the header row.
    if (sheet.getFilter()) sheet.getFilter().remove();                          // Remove any existing filter to prevent errors.
    sheet.getRange(1, 1, sheet.getMaxRows(), 5).createFilter();                 // Create a new filter on the data range.
    if (rows.length > 0) {                                                      // If there are rows...
        sheet.getRange(2, 1, rows.length, 5).sort({ column: 3, ascending: false }); // ...sort the sheet by the "Last Seen" column (column 3).
    }                                                                           // End of if block.
  } catch (e) { console.warn('persist senders err', e); }                       // If an error occurs, log a warning.
}                                                                               // End of persistSendersMapToSheet function.

/**
 * @summary Writes the in-memory name variations map to its sheet, sorted alphabetically.
 */
function persistNameVariationsMapToSheet(map) {                                 // Function definition to save the name variations map.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Name Variations'); // Get the "Name Variations" sheet.
    if (!sheet) return;                                                         // If the sheet doesn't exist, do nothing.
    const rows = Array.from(map.entries())                                      // Convert the map to an array of [email, set] pairs.
      .filter(([email, set]) => set.size > 1)                                   // Keep only the entries where the Set of names has more than one item.
      .sort((a, b) => a[0].localeCompare(b[0]))                                 // Sort the array alphabetically based on the email address (the first element).
      .map(([email, set]) => [email, Array.from(set || []).join(' | ')]);        // Convert the filtered entries to a 2D array, joining the Set of names with a separator.
    sheet.clearContents();                                                      // Clear all existing data from the sheet.
    sheet.getRange(1, 1, 1, 2).setValues([['Email', 'Name Variations']]);       // Write the header row.
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 2).setValues(rows);   // If there is data, write it to the sheet.
    sheet.setFrozenRows(1);                                                     // Freeze the header row.
    if (sheet.getFilter()) sheet.getFilter().remove();                          // Remove any existing filter.
    sheet.getRange(1, 1, sheet.getMaxRows(), 2).createFilter();                 // Create a new filter.
  } catch (e) { console.warn('persist name variations err', e); }               // If an error occurs, log a warning.
}                                                                               // End of persistNameVariationsMapToSheet function.

/**
 * @summary Returns the total number of emails counted from the Senders sheet.
 */
function getTotalEmailsCount() {                                                // Function definition to get the total email count.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Senders'); // Get the "Senders" sheet.
    if (!sheet || sheet.getLastRow() < 2) return 0;                             // If the sheet is empty, return 0.
    const data = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getValues();   // Get all values from the "Email Count" column (column 5).
    return data.reduce((sum, row) => sum + (parseInt(row[0], 10) || 0), 0);      // Use reduce to sum all the values in the column.
  } catch (e) { console.warn('get total emails err', e); return 0; }            // If an error occurs, log a warning and return 0.
}                                                                               // End of getTotalEmailsCount function.

/**
 * @summary Returns the number of unique senders from the Senders sheet.
 */
function getUniqueSendersCount() {                                              // Function definition to get the unique sender count.
  try {                                                                         // Start a try block for sheet operations.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Senders'); // Get the "Senders" sheet.
    return sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;                     // Return the number of rows minus 1 (for the header).
  } catch (e) { console.warn('unique senders err', e); return 0; }              // If an error occurs, log a warning and return 0.
}                                                                               // End of getUniqueSendersCount function.

/**
 * @summary Updates progress counters in the 'Crawler Status' sheet.
 */
function updateCrawlerProgress(threadsProcessed, emailsFound, sendersCount) {   // Function definition to update the progress dashboard.
  try {                                                                         // Start a try block for sheet operations.
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crawler Status'); // Get the "Crawler Status" sheet.
    if (!s) return;                                                             // If the sheet doesn't exist, do nothing.
    s.getRange('B4').setValue(new Date());                                      // Update the "Last Update" timestamp.
    if (threadsProcessed != null) s.getRange('B7').setValue(threadsProcessed);  // Update the "Total Threads Processed" count.
    if (emailsFound != null) s.getRange('B9').setValue(emailsFound);            // Update the "Total Emails Found" count.
    if (sendersCount != null) s.getRange('B10').setValue(sendersCount);         // Update the "Unique Senders Found" count.
    PropertiesService.getScriptProperties().setProperty('totalThreadsProcessed', String(threadsProcessed || 0)); // Save the thread count to script properties for persistence.
  } catch (e) { console.warn('update progress err', e); }                       // If an error occurs, log a warning.
}                                                                               // End of updateCrawlerProgress function.

/**
 * @summary Updates the status message on the dashboard.
 */
function updateCrawlerStatus(status, message) {                                 // Function definition to update the status.
  const ss = SpreadsheetApp.getActiveSpreadsheet();                             // Get the active spreadsheet.
  const statusSheet = ss.getSheetByName('Crawler Status');                      // Get the "Crawler Status" sheet.
  if (!statusSheet) return;                                                     // If the sheet doesn't exist, do nothing.
  const now = new Date();                                                       // Get the current time.
  statusSheet.getRange('B2').setValue(status);                                  // Set the main status message.
  const startedVal = statusSheet.getRange('B3').getValue();                     // Get the value of the "Started" cell.
  if (!startedVal && (status === 'INITIALIZING' || status === 'INITIALIZED')) { // If the "Started" cell is empty and the status is initializing...
    statusSheet.getRange('B3').setValue(now);                                   // ...set the "Started" timestamp.
  }                                                                             // End of if block.
  statusSheet.getRange('B4').setValue(now);                                     // Set the "Last Update" timestamp.
  if (message) statusSheet.getRange('B11').setValue(message);                   // If a descriptive message was provided, set it.
  console.info('Status updated:', status, '-', message || '');                  // Log the status update.
}                                                                               // End of updateCrawlerStatus function.

/**
 * @summary Updates the Keepers sheet with full sender data.
 */
function moveKeepersToEnd() {                                                   // Function definition for moving keepers.
  try {                                                                         // Start a try block for sheet operations.
    const ss = SpreadsheetApp.getActiveSpreadsheet();                           // Get the active spreadsheet.
    const sendersSheet = ss.getSheetByName('Senders');                          // Get the "Senders" sheet.
    const keepersSheet = ss.getSheetByName('Keepers');                          // Get the "Keepers" sheet.

    if (!sendersSheet || !keepersSheet || sendersSheet.getLastRow() < 2) {      // If any required sheet is missing or Senders is empty...
      console.info('Skipping keeper move: No sender data to process.');         // ...log a message and exit.
      return;                                                                   // ...stop the function.
    }                                                                           // End of if block.

    const sendersData = sendersSheet.getRange(2, 1, sendersSheet.getLastRow() - 1, 5).getValues(); // Get all data from the Senders sheet (now 5 columns).
    const keepersData = keepersSheet.getRange(1, 1, keepersSheet.getLastRow(), 1).getValues(); // Get the list of keeper emails.
    const keepersMap = new Map(keepersData.map((row, i) => [row[0].toLowerCase().trim(), i + 1])); // Create a map of keeper emails to their row number.

    const rowsToKeepInSenders = [];                                             // Initialize an array for senders that are not keepers.
    const keepersToUpdate = [];                                                 // Initialize an array for keepers that need updating.

    for (const senderRow of sendersData) {                                      // Loop through each row of sender data.
      const email = senderRow[1].toLowerCase().trim();                          // Get and normalize the sender's email.
      if (keepersMap.has(email)) {                                              // If the sender's email is in the keepers map...
        const rowIndex = keepersMap.get(email);                                 // ...get the row number from the keepers map.
        keepersToUpdate.push({ row: rowIndex, data: senderRow });               // ...add the row number and data to the update list.
      } else {                                                                  // Otherwise...
        rowsToKeepInSenders.push(senderRow);                                    // ...add the sender row to the list of rows to keep.
      }                                                                         // End of if/else block.
    }                                                                           // End of loop.

    if (keepersToUpdate.length > 0) {                                           // If there are any keepers to update...
      console.info(`Updating ${keepersToUpdate.length} sender(s) in the Keepers sheet.`); // ...log how many are being updated.

      for (const keeper of keepersToUpdate) {                                   // Loop through each keeper to be updated.
        keepersSheet.getRange(keeper.row, 1, 1, 5).setValues([keeper.data]);    // Write the full 5-column sender data to the correct row in the Keepers sheet.
      }                                                                         // End of loop.
      
      keepersSheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Last Seen', 'Thread Count', 'Email Count']]); // Set the 5-column header.

      sendersSheet.clearContents();                                             // Clear the entire Senders sheet.
      sendersSheet.getRange(1, 1, 1, 5).setValues([['Name', 'Email', 'Last Seen', 'Thread Count', 'Email Count']]); // Write the header back to the Senders sheet.
      if (rowsToKeepInSenders.length > 0) {                                     // If there are any senders left...
        sendersSheet.getRange(2, 1, rowsToKeepInSenders.length, 5).setValues(rowsToKeepInSenders); // ...write them back to the Senders sheet.
      }                                                                         // End of if block.

    } else {                                                                    // If no matching keepers were found...
      console.info('No senders found matching the Keepers list.');              // ...log that no action was taken.
    }                                                                           // End of if/else block.
  } catch (e) {                                                                 // If any error occurs during the process...
    console.error('Error while moving keepers:', e);                            // ...log a detailed error message.
  }                                                                             // End of try-catch block.
}                                                                               // End of moveKeepersToEnd function.


// --------------------------- Trigger Management Functions --------------------------- //
/**
 * @summary Clears existing crawler triggers and schedules the next run.
 */
function scheduleNextCrawlerRun(delayMs = 2 * 60 * 1000) {                      // Function definition to schedule the next run.
  try {                                                                         // Start a try block for trigger service operations.
    clearCrawlerTriggers();                                                     // Clear any old recurring triggers.
    const nextTime = new Date(Date.now() + delayMs);                            // Calculate the time for the next run.
    ScriptApp.newTrigger('runCrawlerSession').timeBased().at(nextTime).create(); // Create a new trigger to run at the calculated time.
    updateTriggerCountInStatus();                                               // Update the trigger count on the dashboard.
    updateCrawlerStatus('SCHEDULED', `Next run: ${nextTime.toLocaleString()}`); // Update the status to "SCHEDULED" with the next run time.
    console.info('Next run scheduled for', nextTime);                           // Log the time of the next scheduled run.
  } catch (e) { console.warn('schedule err', e); }                              // If an error occurs, log a warning.
}                                                                               // End of scheduleNextCrawlerRun function.

/**
 * @summary Clears only the recurring 'runCrawlerSession' triggers.
 */
function clearCrawlerTriggers() {                                               // Function definition to clear triggers.
  try {                                                                         // Start a try block for trigger service operations.
    ScriptApp.getProjectTriggers().forEach(t => {                               // Get all project triggers and loop through them.
      if (t.getHandlerFunction() === 'runCrawlerSession') {                     // If the trigger is for the main session function...
        ScriptApp.deleteTrigger(t);                                             // ...delete it.
        console.info('Deleted trigger:', t.getHandlerFunction());               // ...and log that it was deleted.
      }                                                                         // End of if block.
    });                                                                         // End of loop.
    updateTriggerCountInStatus();                                               // Update the trigger count on the dashboard.
  } catch (e) { console.warn('clear triggers err', e); }                        // If an error occurs, log a warning.
}                                                                               // End of clearCrawlerTriggers function.

/**
 * @summary Counts active crawler-related triggers and writes the count to the status sheet.
 */
function updateTriggerCountInStatus() {                                         // Function definition to update the trigger count.
  try {                                                                         // Start a try block for trigger service operations.
    const triggers = ScriptApp.getProjectTriggers();                            // Get all triggers for this script project.
    const count = triggers.filter(t => ['runCrawlerSession', 'countThreadsAndStart'].includes(t.getHandlerFunction())).length; // Count how many triggers match the crawler's function names.
    const statusSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crawler Status'); // Get the "Crawler Status" sheet.
    if (statusSheet) statusSheet.getRange('B5').setValue(count);                // If the sheet exists, write the count to it.
  } catch (e) { console.warn('update trigger count err', e); }                  // If an error occurs, log a warning.
}                                                                               // End of updateTriggerCountInStatus function.


// --------------------------- Utility & Debug Functions --------------------------- //
/**
 * @summary Parse "From" header value into a name and email.
 */
function parseSenderInfo(fromValue) {                                           // Function definition to parse sender info.
  if (!fromValue) return { name: '', email: '' };                               // If the input is empty, return an empty object.
  const m = fromValue.match(/^(.*?)\s*<(.+?)>$/);                               // Try to match the "Name <email>" format using a regular expression.
  if (m && m[2]) return { name: m[1].replace(/"/g, '').trim(), email: m[2].trim() }; // If it matches, return the cleaned name and email.
  return { name: '', email: fromValue.trim() };                                 // Otherwise, assume the whole string is the email.
}                                                                               // End of parseSenderInfo function.

/**
 * @summary Checks if the script is approaching its execution time limit.
 */
function timeLimitReached(startTime) {                                          // Function definition to check the time limit.
  const BUFFER_MS = 15000;                                                      // Define a 15-second safety buffer.
  return (Date.now() - startTime) >= (CRAWLER_RUN_MS - BUFFER_MS);              // Return true if the elapsed time is greater than the limit minus the buffer.
}                                                                               // End of timeLimitReached function.

/**
 * @summary A quick helper function to log the current status for debugging.
 */
function checkCrawlerStatus() {                                                 // Function definition for the debug helper.
  try {                                                                         // Start a try block for sheet operations.
    const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Crawler Status'); // Get the "Crawler Status" sheet.
    const status = s ? s.getRange('B2').getValue() : 'NOT FOUND';               // Get the value from the status cell.
    console.info('Crawler status:', status);                                    // Log the current status.
  } catch (e) { console.warn('check status err', e); }                          // If an error occurs, log a warning.
}                                                                               // End of checkCrawlerStatus function.
