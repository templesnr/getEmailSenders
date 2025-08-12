/**
 * Count all Gmail threads safely in batches before the crawl starts.
 * Stores the total in the "Crawler Status" sheet for progress calculations.
 */
function countAllGmailThreadsFast() {
  const BATCH_SIZE = 500;                                                // Max allowed by GmailApp
  let totalThreads = 0;                                                  // Counter for all threads
  let startIndex = 0;                                                     // Start position in search
  let batch;                                                              // Placeholder for fetched threads
  
  console.log("Starting Gmail thread count...");                          // Debug log
  
  while (true) {                                                          // Loop until no more threads
    batch = GmailApp.search('', startIndex, BATCH_SIZE);                  // Fetch next batch of threads
    totalThreads += batch.length;                                         // Add to counter
    startIndex += BATCH_SIZE;                                             // Move to next set
    
    if (batch.length < BATCH_SIZE) {                                      // If last batch is smaller, stop
      break;
    }
  }
  
  console.log(`Total Gmail threads found: ${totalThreads}`);              // Log total count
  
  // Save to "Crawler Status" sheet for later percentage calculations
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();             
  const statusSheet = spreadsheet.getSheetByName("Crawler Status");      
  statusSheet.getRange("B11").setValue(totalThreads);                      // Assuming B11 is "Total Threads Found"
  
  return totalThreads;                                                    // Return total for immediate use
}
