/*
 * GMAIL SENDER ANALYSIS SCRIPT (PART 1)
 * Purpose: Analyze Gmail and create sender lists with email counts
 * Features: Keepers functionality, resume capability, name variations tracking
 * Status: Standalone script for sender analysis only
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Management')
    .addItem('Get Sender List', 'runFilteredSenderList')
    .addItem('Resume Processing', 'resumeProcessing')
    .addSeparator()
    .addItem('Delete Emails from Multiple Senders', 'runDeleteEmailsFromMultipleSenders')
    .addToUi();
}

function runFilteredSenderList() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Start Email Analysis',
      'This will analyze your Gmail to create a sender list. For large accounts, you may need to run this multiple times. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      ui.alert('Analysis started. Large accounts may require multiple runs to complete...');
      const message = filteredSenderList();
      ui.alert('Analysis Complete', message, ui.ButtonSet.OK);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', `Failed to analyze emails: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function filteredSenderList(resumeToken = null) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sendersSheet = spreadsheet.getSheetByName("Senders");
    let progressSheet = spreadsheet.getSheetByName("Progress");
    let keepersSheet = spreadsheet.getSheetByName("Keepers");

    let currentBatchThreads = []; // Store current batch of threads
    let batchPosition = 0; // Position within current batch    

    // Get the user's own email address to exclude
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    console.log(`Excluding emails from user's own address: ${userEmail}`);

    // Get keeper emails to exclude
    const keeperEmails = new Set();
    keeperEmails.add(userEmail); // Always include user's own email
    
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

    // Create or manage the Senders sheet
    if (!sendersSheet) {
      sendersSheet = spreadsheet.insertSheet("Senders");
    } else if (!resumeToken) {
      sendersSheet.clear();
    }

    // Create or manage the Senders sheet
    if (keepersSheet) {
      keepersSheet = spreadsheet.insertSheet("Keepers");
    } else if (!resumeToken) {
      keepersSheet.clear();
    }

    // Create or get progress tracking sheet
    if (!progressSheet) {
      progressSheet = spreadsheet.insertSheet("Progress");
      progressSheet.getRange("A1:B1").setValues([["Last Token", "Processed Count"]]);
    }

    const senders = new Map();
    let processedThreads = 0;
    const maxThreads = 3000; // Conservative limit to prevent timeout
    const timeLimit = 4.5 * 60 * 1000; // 4.5 minutes to leave buffer for name variations sheet
    const startTime = Date.now();

    // Load existing data if resuming
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

    // Fetch threads with pagination
    let pageToken = resumeToken;
    do {
      try {
        // If we're resuming mid-batch, skip the API call
        if (currentBatchThreads.length === 0) {
          const threads = Gmail.Users.Threads.list("me", { 
            maxResults: 50,
            pageToken: pageToken 
          });
          
          if (threads.threads) {
            currentBatchThreads = threads.threads;
            batchPosition = 0; // Start from beginning of new batch
          } else {
            break; // No more threads
          }
          
          pageToken = threads.nextPageToken;
        }
        
        // Process threads from current position in batch
        for (let i = batchPosition; i < currentBatchThreads.length; i++) {
          const thread = currentBatchThreads[i];
          
          // Check limits more frequently
          if (processedThreads >= maxThreads) {
            console.log(`Max threads limit (${maxThreads}) reached`);
            saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, i, true);
            return `MAX_REACHED: Processed ${processedThreads} threads. Run resumeProcessing() to continue.`;
          }
          
          // Time check during processing - save exact position
          if (Date.now() - startTime > timeLimit) {
            console.log(`Time limit reached during processing at thread ${processedThreads}, batch position ${i}`);
            saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, i, true);
            return `TIMEOUT: Processed ${processedThreads} threads. Run resumeProcessing() to continue from batch position ${i}.`;
          }
          
          try {
            const threadDetails = Gmail.Users.Threads.get("me", thread.id, {
              format: 'metadata',
              metadataHeaders: ['From', 'Date']
            });
            
            const message = threadDetails.messages[0];
            const headers = message.payload.headers;
            
            const senderHeader = headers.find(h => h.name === "From");
            const dateHeader = headers.find(h => h.name === "Date");
            
            if (senderHeader) {
              const { name, email } = parseSenderInfo(senderHeader.value);
              if (email && !keeperEmails.has(email)) {
                const messageDate = dateHeader ? new Date(dateHeader.value) : new Date(parseInt(message.internalDate));
                
                if (!senders.has(email)) {
                  senders.set(email, {
                    primaryName: name || email,
                    email: email,
                    date: messageDate,
                    count: 1,
                    nameVariations: new Set([name || email])
                  });
                } else {
                  const existing = senders.get(email);
                  existing.count += 1;
                  
                  if (name && name.trim() !== '') {
                    existing.nameVariations.add(name.trim());
                  }
                  
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
            console.warn(`Error processing thread ${thread.id}:`, threadError.message || threadError);
            processedThreads++; // Still count as processed
            continue;
          }
        }
        
        // Finished processing current batch - clear it for next iteration
        currentBatchThreads = [];
        batchPosition = 0;
        
        // Less frequent saving to reduce overhead
        if (processedThreads % 1500 === 0) {
          console.log(`Processed ${processedThreads} threads...`);
          saveProgressAndResultsQuick(spreadsheet, senders, pageToken, processedThreads, false);
        }
        
      } catch (batchError) {
        console.error('Error fetching thread batch:', batchError);
        saveProgressWithBatch(spreadsheet, senders, pageToken, processedThreads, currentBatchThreads, batchPosition, true);
        throw new Error(`API error after ${processedThreads} threads: ${batchError.message}`);
      }
      
    } while (pageToken && processedThreads < maxThreads && (Date.now() - startTime) < timeLimit);

    // Completed successfully
    const timeRemaining = timeLimit - (Date.now() - startTime);
    const hasTimeForExtras = timeRemaining > 30000; // 30+ seconds left
    
    saveProgressAndResults(spreadsheet, senders, null, processedThreads, false);
    
    // Only create name variations sheet if we have time
    if (hasTimeForExtras) {
      createNameVariationsSheet(spreadsheet, senders);
    } else {
      console.log("Skipping name variations sheet due to time constraints");
    }
    
    const excludedCount = keeperEmails.size;
    return `SUCCESS: Processed ${processedThreads} threads, found ${senders.size} unique senders (excluding ${excludedCount} keeper emails).`;
    
  } catch (error) {
    console.error('Main function error:', error);
    throw new Error(`Failed to get sender list: ${error.message}`);
  }
}
// Save progress with exact batch position
function saveProgressWithBatch(spreadsheet, senders, pageToken, processedCount, currentBatch, batchPosition, isIncomplete) {
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);

    if (sendersArray.length > 0) {
      sendersSheet.clear();
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      sendersSheet.autoResizeColumns(1, 4);
      sendersSheet.setFrozenRows(1);
    }
    
    // Save progress with batch information
    const batchData = currentBatch.length > 0 ? JSON.stringify(currentBatch) : null;
    progressSheet.getRange("A2:D2").setValues([[
      pageToken || "COMPLETE", 
      processedCount,
      batchData,
      batchPosition
    ]]);
    
    const statusRange = sendersSheet.getRange("F1");
    if (isIncomplete) {
      statusRange.setValue(`INCOMPLETE: ${sendersArray.length} senders from ${processedCount} threads (batch pos: ${batchPosition})`);
      statusRange.setBackground("#ffcccc");
    } else {
      statusRange.setValue(`COMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ccffcc");
    }
    
    console.log(`Saved ${sendersArray.length} senders at batch position ${batchPosition}. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    console.error('Error saving progress with batch:', error);
  }
}
// Quick save function without name variations sheet
function saveProgressAndResultsQuick(spreadsheet, senders, pageToken, processedCount, isIncomplete) {
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);

    if (sendersArray.length > 0) {
      sendersSheet.clear();
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      sendersSheet.autoResizeColumns(1, 4);
      sendersSheet.setFrozenRows(1);
    }
    
    progressSheet.getRange("A2:D2").setValues([[pageToken || "COMPLETE", processedCount, null, 0]]);
    
    console.log(`Quick saved ${sendersArray.length} senders. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    console.error('Error in quick save:', error);
  }
}

function saveProgressAndResults(spreadsheet, senders, pageToken, processedCount, isIncomplete) {
  try {
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    const progressSheet = spreadsheet.getSheetByName("Progress");
    
    const sendersArray = Array.from(senders.values())
      .sort((a, b) => b.date - a.date)
      .map(sender => [sender.primaryName, sender.email, sender.date, sender.count]);

    if (sendersArray.length > 0) {
      sendersSheet.clear();
      const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...sendersArray];
      
      sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
      
      const dateRange = sendersSheet.getRange(2, 3, sendersArray.length, 1);
      dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
      
      sendersSheet.autoResizeColumns(1, 4);
      
      try {
        const headerRange = sendersSheet.getRange(1, 1, dataWithHeaders.length, 4);
        if (!sendersSheet.getFilter()) {
          headerRange.createFilter();
        }
      } catch (filterError) {
        console.warn('Filter creation skipped:', filterError);
      }
      
      sendersSheet.setFrozenRows(1);
    }
    
    progressSheet.getRange("A2:B2").setValues([[pageToken || "COMPLETE", processedCount]]);
    
    const statusRange = sendersSheet.getRange("F1");
    if (isIncomplete) {
      statusRange.setValue(`INCOMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ffcccc");
    } else {
      statusRange.setValue(`COMPLETE: ${sendersArray.length} senders from ${processedCount} threads`);
      statusRange.setBackground("#ccffcc");
    }
    
    console.log(`Saved ${sendersArray.length} senders. Status: ${isIncomplete ? 'INCOMPLETE' : 'COMPLETE'}`);
    
  } catch (error) {
    console.error('Error saving progress:', error);
  }
}

function createNameVariationsSheet(spreadsheet, senders) {
  try {
    let nameVariationsSheet = spreadsheet.getSheetByName("Name Variations");
    
    if (!nameVariationsSheet) {
      nameVariationsSheet = spreadsheet.insertSheet("Name Variations");
    } else {
      nameVariationsSheet.clear();
    }
    
    const variationsData = [];
    
    Array.from(senders.values()).forEach(sender => {
      const variations = Array.from(sender.nameVariations);
      if (variations.length > 1) {
        const cleanVariations = [...new Set(variations.filter(name => name && name.trim() !== ''))];
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
    
    if (variationsData.length > 0) {
      variationsData.sort((a, b) => b[1] - a[1]);
      
      const headersAndData = [
        ["Email Address", "Total Emails", "Name Variations", "Variation Count"],
        ...variationsData
      ];
      
      nameVariationsSheet.getRange(1, 1, headersAndData.length, 4).setValues(headersAndData);
      nameVariationsSheet.autoResizeColumns(1, 4);
      
      const headerRange = nameVariationsSheet.getRange(1, 1, 1, 4);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      try {
        const dataRange = nameVariationsSheet.getRange(1, 1, headersAndData.length, 4);
        if (!nameVariationsSheet.getFilter()) {
          dataRange.createFilter();
        }
      } catch (filterError) {
        console.warn('Filter creation skipped for Name Variations:', filterError);
      }
      
      nameVariationsSheet.setFrozenRows(1);
      
      nameVariationsSheet.getRange("F1").setValue(`Found ${variationsData.length} emails with multiple name variations`);
      nameVariationsSheet.getRange("F1").setBackground("#fff2cc");
      
      console.log(`Created Name Variations sheet with ${variationsData.length} entries`);
    } else {
      nameVariationsSheet.getRange(1, 1, 2, 4).setValues([
        ["Email Address", "Total Emails", "Name Variations", "Variation Count"],
        ["No email addresses found with multiple name variations", "", "", ""]
      ]);
      nameVariationsSheet.autoResizeColumns(1, 4);
      
      const headerRange = nameVariationsSheet.getRange(1, 1, 1, 4);
      headerRange.setBackground("#4285f4");
      headerRange.setFontColor("white");
      headerRange.setFontWeight("bold");
      
      nameVariationsSheet.setFrozenRows(1);
      console.log("No name variations found");
    }
    
  } catch (error) {
    console.error('Error creating name variations sheet:', error);
  }
}

function parseSenderInfo(fromHeader) {
  if (!fromHeader) return { name: null, email: null };
  
  let match = fromHeader.match(/^"([^"]+)"\s*<([^>]+)>$/);
  if (match) {
    return { name: match[1].trim(), email: match[2].trim().toLowerCase() };
  }
  
  match = fromHeader.match(/^([^<]+)<([^>]+)>$/);
  if (match) {
    return { name: match[1].trim(), email: match[2].trim().toLowerCase() };
  }
  
  match = fromHeader.match(/^<([^>]+)>$/);
  if (match) {
    return { name: null, email: match[1].trim().toLowerCase() };
  }
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  const cleanEmail = fromHeader.trim().toLowerCase();
  if (emailRegex.test(cleanEmail)) {
    return { name: null, email: cleanEmail };
  }
  
  return { name: null, email: null };
}

function resumeProcessing() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = spreadsheet.getSheetByName("Progress");
  
  if (!progressSheet) {
    console.log("No progress sheet found. Starting fresh...");
    return filteredSenderList();
  }
  
  const progressData = progressSheet.getRange("A2:B2").getValues()[0];
  const lastToken = progressData[0];
  const processedCount = progressData[1];
  
  if (!lastToken || lastToken === "COMPLETE") {
    console.log("Processing already complete or no resume token found.");
    return "Already complete or no progress to resume.";
  }
  
  console.log(`Resuming from token with ${processedCount} threads already processed...`);
  return filteredSenderList(lastToken);
}
