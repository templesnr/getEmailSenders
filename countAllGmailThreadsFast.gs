// --------------------------- countAllGmailThreadsFast --------------------------- //
function countAllGmailThreadsFast() {                                                              // count all Gmail threads quickly in batches
  const BATCH_SIZE = 500;                                                                          // max threads GmailApp can fetch at once
  let totalThreads = 0;                                                                            // total threads counter
  let startIndex = 0;                                                                              // starting index for search pagination
  let batch;                                                                                       // placeholder for fetched batch

  console.log("count start");                                                                      // log start of count

  while (true) {                                                                                   // loop until no more threads
    batch = GmailApp.search('', startIndex, BATCH_SIZE);                                           // fetch next batch of threads
    totalThreads += batch.length;                                                                  // add to total counter
    startIndex += BATCH_SIZE;                                                                      // move to next batch index
    if (batch.length < BATCH_SIZE) break;                                                          // if batch smaller than max, stop
  }

  console.log(`threads total: ${totalThreads}`);                                                   // log total found

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();                                      // get active spreadsheet
  const statusSheet = spreadsheet.getSheetByName("Crawler Status");                               // open crawler status sheet
  statusSheet.insertRowBefore(5);                                                                  // insert a new row before row 5
  statusSheet.getRange("A5").setValue("Total Gmail Threads");                                      // label the row
  statusSheet.getRange("B5").setValue(totalThreads);                                               // write the total threads count

  return totalThreads;                                                                             // return total for caller
}                                                                                                  // end countAllGmailThreadsFast
