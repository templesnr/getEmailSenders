function onOpen() {
  const ui = SpreadsheetApp.getUi(); // Get the UI object so we can add menu items
  ui.createMenu('Email Management')  // Create a new top-level menu
    .addItem('Get Sender List', 'runFilteredSenderList') // Existing function for getting sender list
    .addItem('Resume Processing', 'resumeProcessing')    // Existing function for resuming other process
    .addSeparator() // Add a divider line in the menu
    .addItem('🤖 Start Silent Crawler (Fresh Start)', 'startSilentCrawlerFresh') // Start from scratch
    .addItem('▶ Resume Silent Crawler', 'startSilentCrawlerResume')              // Continue from saved point
    .addItem('⏹️ Stop Silent Crawler', 'stopSilentCrawler')                       // Stop crawler manually
    .addItem('📊 Check Crawler Status', 'checkCrawlerStatus')                     // View current crawler status
    .addSeparator() // Another divider
    .addItem('Delete Emails from Multiple Senders', 'runDeleteEmailsFromMultipleSenders') // Existing bulk delete
    .addToUi(); // Add the menu to the UI
}

/**
 * Menu helper: starts crawler with fresh start = true.
 */
function startSilentCrawlerFresh() {
  startSilentCrawler(true); // Pass true to signal fresh start
}

/**
 * Menu helper: starts crawler with fresh start = false.
 */
function startSilentCrawlerResume() {
  startSilentCrawler(false); // Pass false to signal resume
}
