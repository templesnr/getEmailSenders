function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Management')
    .addItem('Get Sender List', 'runFilteredSenderList')
    .addItem('Resume Processing', 'resumeProcessing')
    .addSeparator()
    .addItem('Delete Emails from Multiple Senders', 'runDeleteEmailsFromMultipleSenders')
    .addToUi();
}
