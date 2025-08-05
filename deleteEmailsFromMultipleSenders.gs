/*
 * GMAIL EMAIL DELETION SCRIPT (PART 2)
 * Purpose: Delete emails from multiple selected senders
 * Features: Multi-sender selection, trash vs permanent delete, safety confirmations
 * Status: Standalone script for email deletion only
 * Prerequisite: Requires "Senders" sheet created by Part 1 (Sender Analysis Script)
 */

function deleteEmailsFromMultipleSenders() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    
    if (!sendersSheet || sendersSheet.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert('Error', 'No sender data found. Please run the Sender Analysis Script (Part 1) first to create the sender list.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const senderData = sendersSheet.getRange(2, 1, sendersSheet.getLastRow() - 1, 4).getValues();
    
    const sortedSenders = senderData
      .filter(row => row[1]) // Filter out empty email addresses
      .sort((a, b) => (b[3] || 0) - (a[3] || 0)) // Sort by email count descending
      .map((row, index) => ({
        originalIndex: index,
        name: row[0] || 'No Name',
        email: row[1],
        date: row[2],
        count: row[3] || 0,
        displayText: `${row[1]} (${row[0] || 'No Name'}) - ${row[3] || 0} emails`
      }));
    
    if (sortedSenders.length === 0) {
      SpreadsheetApp.getUi().alert('Error', 'No valid sender data found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    const ui = SpreadsheetApp.getUi();
    const htmlOutput = HtmlService.createHtmlOutput(createMultiSenderSelectionHtml(sortedSenders))
      .setWidth(700)
      .setHeight(500);
    
    ui.showModalDialog(htmlOutput, 'Select Multiple Senders to Process');
    
  } catch (error) {
    console.error('Error in deleteEmailsFromMultipleSenders:', error);
    SpreadsheetApp.getUi().alert('Error', `Failed to load sender list: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createMultiSenderSelectionHtml(senders) {
  const senderCheckboxes = senders.map((sender, index) => 
    `<div class="sender-item">
       <input type="checkbox" id="sender_${index}" value="${sender.email}" data-count="${sender.count}">
       <label for="sender_${index}">${sender.displayText}</label>
     </div>`
  ).join('');
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .container { max-width: 100%; }
        .info { background: #e7f3ff; padding: 10px; border-radius: 5px; margin-bottom: 15px; }
        .senders-container { 
          max-height: 200px; overflow-y: auto; border: 1px solid #ddd; 
          padding: 10px; border-radius: 5px; background: #fafafa; margin: 10px 0;
        }
        .sender-item { margin: 8px 0; padding: 5px; border-radius: 3px; transition: background-color 0.2s; }
        .sender-item:hover { background-color: #e9ecef; }
        .sender-item input[type="checkbox"] { margin-right: 8px; }
        .sender-item label { cursor: pointer; font-size: 14px; }
        .selection-controls { margin: 10px 0; padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #dee2e6; }
        .selection-controls button { margin: 0 5px; padding: 5px 10px; font-size: 12px; border: 1px solid #ccc; border-radius: 3px; background: #fff; cursor: pointer; }
        .selection-summary { margin: 10px 0; padding: 10px; background: #d1ecf1; border: 1px solid #bee5eb; border-radius: 5px; font-weight: bold; }
        .action-selection { background: #f8f9fa; padding: 15px; border-radius: 5px; margin: 15px 0; border: 1px solid #dee2e6; }
        .radio-option { margin: 10px 0; padding: 8px; border-radius: 3px; transition: background-color 0.2s; }
        .radio-option:hover { background-color: #e9ecef; }
        .radio-option input[type="radio"] { margin-right: 8px; }
        .radio-option label { cursor: pointer; font-weight: normal; }
        .safe-option { color: #155724; }
        .danger-option { color: #721c24; }
        .button-container { margin-top: 20px; text-align: center; }
        button { padding: 10px 20px; margin: 0 10px; font-size: 14px; border-radius: 5px; border: 1px solid #ccc; cursor: pointer; }
        .action-btn { background: #007bff; color: white; }
        .action-btn:disabled { background: #6c757d; cursor: not-allowed; }
        .cancel-btn { background: #f0f0f0; color: #333; }
        .warning { background: #fff3cd; border: 1px solid #ffeaa7; padding: 10px; border-radius: 5px; margin: 10px 0; }
        .danger-warning { background: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="info">
          <strong>Select multiple senders and choose what to do with ALL their emails.</strong><br>
          Senders are sorted by email count (highest first). Use checkboxes to select multiple senders.
        </div>
        
        <label><strong>Choose senders:</strong></label>
        <div class="selection-controls">
          <button onclick="selectAll()">Select All</button>
          <button onclick="selectNone()">Select None</button>
          <button onclick="selectTop10()">Select Top 10</button>
          <button onclick="selectByCount()">Select 50+ Emails</button>
        </div>
        
        <div class="senders-container">
          ${senderCheckboxes}
        </div>
        
        <div class="selection-summary" id="selectionSummary">
          No senders selected
        </div>
        
        <div class="action-selection">
          <strong>Choose action for selected senders:</strong>
          
          <div class="radio-option safe-option">
            <input type="radio" id="trash" name="action" value="trash" checked>
            <label for="trash">
              <strong>Move to Trash</strong> (Recommended - Recoverable for 30 days)
              <div style="font-size: 12px; margin-top: 3px;">
                ✅ Can be undone • ✅ 30-day recovery window • ✅ Safer option
              </div>
            </label>
          </div>
          
          <div class="radio-option danger-option">
            <input type="radio" id="delete" name="action" value="delete">
            <label for="delete">
              <strong>Delete Permanently</strong> (Cannot be undone!)
              <div style="font-size: 12px; margin-top: 3px;">
                ⚠️ Permanent • ⚠️ No recovery • ⚠️ Use with extreme caution
              </div>
            </label>
          </div>
        </div>
        
        <div class="warning" id="trashWarning">
          📧 Selected emails will be moved to Trash. You can restore them within 30 days if needed.
        </div>
        
        <div class="warning danger-warning" id="deleteWarning" style="display: none;">
          ⚠️ <strong>DANGER:</strong> Selected emails will be permanently deleted immediately with no way to recover them!
        </div>
        
        <div class="button-container">
          <button class="action-btn" onclick="confirmAndProcess()" id="actionBtn" disabled>Move to Trash</button>
          <button class="cancel-btn" onclick="google.script.host.close()">Cancel</button>
        </div>
      </div>
      
      <script>
        const senders = ${JSON.stringify(senders)};
        
        document.querySelectorAll('input[name="action"]').forEach(radio => {
          radio.addEventListener('change', updateActionUI);
        });
        
        document.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
          checkbox.addEventListener('change', updateSelectionSummary);
        });
        
        function selectAll() {
          document.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
          updateSelectionSummary();
        }
        
        function selectNone() {
          document.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
          updateSelectionSummary();
        }
        
        function selectTop10() {
          selectNone();
          document.querySelectorAll('input[type="checkbox"]').forEach((cb, index) => {
            if (index < 10) cb.checked = true;
          });
          updateSelectionSummary();
        }
        
        function selectByCount() {
          selectNone();
          document.querySelectorAll('input[type="checkbox"]').forEach(cb => {
            const count = parseInt(cb.dataset.count) || 0;
            if (count >= 50) cb.checked = true;
          });
          updateSelectionSummary();
        }
        
        function updateSelectionSummary() {
          const selectedCheckboxes = document.querySelectorAll('input[type="checkbox"]:checked');
          const totalEmails = Array.from(selectedCheckboxes).reduce((sum, cb) => {
            return sum + (parseInt(cb.dataset.count) || 0);
          }, 0);
          
          const summaryEl = document.getElementById('selectionSummary');
          const actionBtn = document.getElementById('actionBtn');
          
          if (selectedCheckboxes.length === 0) {
            summaryEl.textContent = 'No senders selected';
            actionBtn.disabled = true;
          } else {
            summaryEl.innerHTML = \`<strong>\${selectedCheckboxes.length} senders selected</strong> - Total emails: <strong>\${totalEmails}</strong>\`;
            actionBtn.disabled = false;
          }
        }
        
        function updateActionUI() {
          const selectedAction = document.querySelector('input[name="action"]:checked').value;
          const actionBtn = document.getElementById('actionBtn');
          const trashWarning = document.getElementById('trashWarning');
          const deleteWarning = document.getElementById('deleteWarning');
          
          if (selectedAction === 'trash') {
            actionBtn.textContent = 'Move to Trash';
            actionBtn.className = 'action-btn';
            trashWarning.style.display = 'block';
            deleteWarning.style.display = 'none';
          } else {
            actionBtn.textContent = 'Delete Permanently';
            actionBtn.className = 'action-btn';
            actionBtn.style.background = '#dc3545';
            trashWarning.style.display = 'none';
            deleteWarning.style.display = 'block';
          }
        }
        
        function confirmAndProcess() {
          const selectedCheckboxes = document.querySelectorAll('input[type="checkbox"]:checked');
          
          if (selectedCheckboxes.length === 0) {
            alert('Please select at least one sender.');
            return;
          }
          
          const selectedAction = document.querySelector('input[name="action"]:checked').value;
          const selectedEmails = Array.from(selectedCheckboxes).map(cb => cb.value);
          const totalEmails = Array.from(selectedCheckboxes).reduce((sum, cb) => {
            return sum + (parseInt(cb.dataset.count) || 0);
          }, 0);
          
          const actionText = selectedAction === 'trash' ? 'move to trash' : 'permanently delete';
          const recoveryText = selectedAction === 'trash' ? '\\n\\n✅ You can recover from Trash within 30 days' : '\\n\\n❌ THIS CANNOT BE UNDONE!\\n❌ NO RECOVERY POSSIBLE!';
          
          const confirmMessage = \`⚠️ GOOGLE SECURITY NOTICE ABOVE IS NORMAL ⚠️

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📋 CONFIRMATION REQUIRED 📋
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

You are about to \${actionText.toUpperCase()}:

📧 \${selectedEmails.length} senders
📊 \${totalEmails} total emails\${recoveryText}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Click OK to proceed or Cancel to abort.\`;

          const isConfirmed = confirm(confirmMessage);
          
          if (isConfirmed && selectedAction === 'delete') {
            const finalConfirmMessage = \`⚠️ GOOGLE SECURITY NOTICE ABOVE IS NORMAL ⚠️

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ FINAL WARNING ⚠️
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

🚨 PERMANENT DELETION 🚨

📧 \${selectedEmails.length} senders
📊 \${totalEmails} emails

❌ THIS CANNOT BE UNDONE!
❌ NO RECOVERY POSSIBLE!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LAST CHANCE - Click OK to DELETE FOREVER
or Cancel to abort safely.\`;

            const doubleConfirm = confirm(finalConfirmMessage);
            if (!doubleConfirm) return;
          }
          
          if (isConfirmed) {
            const actionBtn = document.getElementById('actionBtn');
            actionBtn.disabled = true;
            actionBtn.textContent = selectedAction === 'trash' ? 'Moving to Trash...' : 'Deleting...';
            
            google.script.run
              .withSuccessHandler(onProcessSuccess)
              .withFailureHandler(onProcessFailure)
              .performMultipleEmailActions(selectedEmails, selectedAction);
          }
        }
        
        function onProcessSuccess(result) {
          alert(\`⚠️ GOOGLE SECURITY NOTICE ABOVE IS NORMAL ⚠️\\n\\n✅ SUCCESS! ✅\\n\\n\${result}\\n\\n━━━━━━━━━━━━━━━━━━━━━━━━━\\nOperation completed successfully.\`);
          google.script.host.close();
        }
        
        function onProcessFailure(error) {
          alert(\`⚠️ GOOGLE SECURITY NOTICE ABOVE IS NORMAL ⚠️\\n\\n❌ ERROR ❌\\n\\n\${error.message || error}\\n\\n━━━━━━━━━━━━━━━━━━━━━━━━━\\nPlease try again or contact support.\`);
          const actionBtn = document.getElementById('actionBtn');
          actionBtn.disabled = false;
          updateActionUI();
        }
        
        updateSelectionSummary();
      </script>
    </body>
    </html>
  `;
}

function performMultipleEmailActions(emailAddresses, action) {
  try {
    console.log(`Starting ${action} of emails from ${emailAddresses.length} senders`);
    
    let totalProcessed = 0;
    let totalEmails = 0;
    const results = [];
    
    for (let i = 0; i < emailAddresses.length; i++) {
      const emailAddress = emailAddresses[i];
      console.log(`Processing sender ${i + 1}/${emailAddresses.length}: ${emailAddress}`);
      
      try {
        let processedCount = 0;
        let iterationsForThisSender = 0;
        const maxIterationsPerSender = 200;
        const batchSize = 100;
        
        const searchQuery = `from:${emailAddress}`;
        
        do {
          iterationsForThisSender++;
          if (iterationsForThisSender > maxIterationsPerSender) {
            console.warn(`Reached max iterations for ${emailAddress}. Some emails may remain.`);
            break;
          }
          
          const threads = Gmail.Users.Threads.list("me", {
            q: searchQuery,
            maxResults: batchSize
          });
          
          if (!threads.threads || threads.threads.length === 0) {
            break;
          }
          
          const threadIds = threads.threads.map(thread => thread.id);
          
          for (const threadId of threadIds) {
            try {
              if (action === 'trash') {
                Gmail.Users.Threads.trash("me", threadId);
              } else if (action === 'delete') {
                Gmail.Users.Threads.remove("me", threadId);
              }
              processedCount++;
              totalProcessed++;
            } catch (threadError) {
              console.warn(`Failed to ${action} thread ${threadId} from ${emailAddress}:`, threadError);
            }
          }
          
          // Rate limiting
          if (processedCount % 25 === 0) {
            Utilities.sleep(500);
          }
          
        } while (true);
        
        totalEmails += processedCount;
        results.push(`${emailAddress}: ${processedCount} emails`);
        
        // Update sender list after each sender
        updateSenderListAfterDeletion(emailAddress);
        
        console.log(`Completed ${emailAddress}: ${processedCount} emails processed`);
        
      } catch (senderError) {
        console.error(`Error processing ${emailAddress}:`, senderError);
        results.push(`${emailAddress}: ERROR - ${senderError.message}`);
      }
    }
    
    const actionText = action === 'trash' ? 'moved to trash' : 'permanently deleted';
    const recoveryNote = action === 'trash' ? ' (recoverable from Trash for 30 days)' : '';
    
    const summary = `Successfully ${actionText} ${totalEmails} emails from ${emailAddresses.length} senders${recoveryNote}`;
    const detailedResults = results.join('\n');
    
    console.log(`SUMMARY: ${summary}`);
    console.log(`DETAILS:\n${detailedResults}`);
    
    return `${summary}\n\nDetails:\n${detailedResults}`;
    
  } catch (error) {
    console.error(`Error in performMultipleEmailActions (${action}):`, error);
    throw new Error(`Failed to ${action} emails: ${error.message}`);
  }
}

function updateSenderListAfterDeletion(deletedEmailAddress) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sendersSheet = spreadsheet.getSheetByName("Senders");
    
    if (!sendersSheet || sendersSheet.getLastRow() < 2) {
      return;
    }
    
    const range = sendersSheet.getRange(2, 1, sendersSheet.getLastRow() - 1, 4);
    const senderData = range.getValues();
    
    const filteredData = senderData.filter(row => row[1] !== deletedEmailAddress);
    
    if (filteredData.length < senderData.length) {
      sendersSheet.clear();
      
      if (filteredData.length > 0) {
        const dataWithHeaders = [["Name", "Email", "Most Recent Email Date", "Count of Emails"], ...filteredData];
        sendersSheet.getRange(1, 1, dataWithHeaders.length, 4).setValues(dataWithHeaders);
        
        const dateRange = sendersSheet.getRange(2, 3, filteredData.length, 1);
        dateRange.setNumberFormat("dd/mm/yyyy hh:mm:ss");
        sendersSheet.autoResizeColumns(1, 4);
        
        try {
          const headerRange = sendersSheet.getRange(1, 1, dataWithHeaders.length, 4);
          if (!sendersSheet.getFilter()) {
            headerRange.createFilter();
          }
        } catch (filterError) {
          console.warn('Filter creation skipped after deletion update:', filterError);
        }
        
        sendersSheet.setFrozenRows(1);
        
      } else {
        sendersSheet.getRange(1, 1, 1, 4).setValues([["Name", "Email", "Most Recent Email Date", "Count of Emails"]]);
        sendersSheet.setFrozenRows(1);
      }
      
      console.log(`Updated sender list: removed ${deletedEmailAddress}`);
    }
    
  } catch (error) {
    console.error('Error updating sender list:', error);
  }
}
