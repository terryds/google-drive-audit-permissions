/**
 * Drive Audit - Open Source
 * Audits Google Drive files and their permissions
 * 
 * Version: 2.0.0 (Open Source)
 * License: MIT
 * 
 * Features:
 * - Fast processing with 1-minute continuation intervals
 * - Batch processing to handle large Drive accounts
 * - Real-time status tracking
 * - Scheduled weekly audits
 */

/**
 * Creates a custom menu in Google Sheets when the script is opened.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Drive Audit')
    .addItem('Run Audit Now', 'runDriveAudit')
    .addItem('Check Audit Status', 'showAuditStatus')
    .addItem('Cancel Running Audit', 'cancelRunningAudit')
    .addSeparator()
    .addItem('Setup Weekly Schedule', 'showScheduleDialog')
    .addItem('Remove Schedule', 'removeScheduledAudits')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

/**
 * Runs when the script is installed.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Updates the audit status sheet
 */
function updateAuditStatus(status, message, filesProcessed, totalFiles) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let statusSheet = ss.getSheetByName('Audit Status');
    
    if (!statusSheet) {
      statusSheet = ss.insertSheet('Audit Status', 0);
    }
    
    statusSheet.clear();
    
    const statusData = [
      ['Drive Audit Status', ''],
      ['', ''],
      ['Current Status:', status],
      ['Last Updated:', new Date()],
      ['Message:', message],
      ['', '']
    ];
    
    if (totalFiles > 0) {
      statusData.push(['Files Processed:', filesProcessed + ' / ' + totalFiles]);
      statusData.push(['Progress:', Math.round((filesProcessed / totalFiles) * 100) + '%']);
    }
    
    statusSheet.getRange(1, 1, statusData.length, 2).setValues(statusData);
    
    // Format status
    statusSheet.getRange(1, 1, 1, 2).merge()
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    
    statusSheet.getRange(3, 1, statusData.length - 2, 1).setFontWeight('bold');
    
    // Color code status
    const statusCell = statusSheet.getRange(3, 2);
    if (status === 'RUNNING') {
      statusCell.setBackground('#fff3cd').setFontColor('#856404');
    } else if (status === 'COMPLETED') {
      statusCell.setBackground('#d4edda').setFontColor('#155724');
    } else if (status === 'ERROR') {
      statusCell.setBackground('#f8d7da').setFontColor('#721c24');
    } else if (status === 'CANCELLED') {
      statusCell.setBackground('#e0e0e0').setFontColor('#424242');
    }
    
    statusSheet.setColumnWidth(1, 200);
    statusSheet.setColumnWidth(2, 400);
    
    Logger.log('Status updated: ' + status + ' - ' + message);
  } catch (error) {
    Logger.log('Error updating status sheet: ' + error.toString());
  }
}

/**
 * Main function to audit Google Drive files and permissions
 * This version handles timeouts by processing in batches
 */
function runDriveAudit() {
  // Clear any previous audit state
  PropertiesService.getScriptProperties().deleteProperty('AUDIT_STATE');
  PropertiesService.getScriptProperties().deleteProperty('AUDIT_PAGE_TOKEN');
  
  Logger.log('=== DRIVE AUDIT STARTED (FRESH) ===');
  Logger.log('Start time: ' + new Date().toISOString());
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Update status to RUNNING
  updateAuditStatus('RUNNING', 
    'Audit is in progress. For large Drive accounts, this may take some time. The audit will automatically continue every minute if needed.', 
    0, 0);
  
  // Show progress message
  ui.alert('Drive Audit', 
    'Starting audit... This may take several minutes depending on the number of files.\n\n' +
    '⏳ The audit will run in the background.\n' +
    '📊 Check the "Audit Status" sheet for progress.\n' +
    '⏱️ For large Drive accounts (1000+ files):\n' +
    '   • Processes ~500 files every 4.5 minutes\n' +
    '   • Auto-continues every minute if needed\n' +
    '   • May take some time to complete\n\n' +
    'Click OK to start.',
    ui.ButtonSet.OK);
  
  // Call the batch processor
  processDriveAuditBatch();
}

/**
 * Processes the drive audit in batches to avoid timeouts
 * Can be called multiple times to continue where it left off
 */
function processDriveAuditBatch() {
  Logger.log('=== BATCH PROCESSING STARTED ===');
  Logger.log('Start time: ' + new Date().toISOString());
  
  const startTime = new Date().getTime();
  const MAX_EXECUTION_TIME = 4.5 * 60 * 1000; // 4.5 minutes (leaving 1.5 min buffer)
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();
  
  try {
    // Get or initialize audit state
    let auditState = scriptProps.getProperty('AUDIT_STATE');
    let isFirstRun = !auditState;
    
    if (!auditState) {
      Logger.log('First run - initializing audit');
      auditState = {
        phase: 'SETUP',
        totalFilesFound: 0,
        filesProcessed: 0,
        auditDataCount: 0,
        pageToken: null,
        startTime: new Date().toISOString()
      };
    } else {
      auditState = JSON.parse(auditState);
      Logger.log('Continuing audit from phase: ' + auditState.phase);
      Logger.log('Files processed so far: ' + auditState.filesProcessed);
    }
    
    let auditSheet = ss.getSheetByName('Drive Audit');
    
    // SETUP PHASE
    if (auditState.phase === 'SETUP') {
      Logger.log('Setting up audit sheet...');
      if (auditSheet) {
        Logger.log('Clearing existing "Drive Audit" sheet');
        auditSheet.clear();
      } else {
        Logger.log('Creating new "Drive Audit" sheet');
        auditSheet = ss.insertSheet('Drive Audit');
      }
      
      // Set up headers
      Logger.log('Setting up headers...');
      const headers = [
        'File Name',
        'File ID',
        'Owner',
        'Type',
        'MIME Type',
        'Created Date',
        'Modified Date',
        'Size (bytes)',
        'URL',
        'Permissions Count',
        'Permission Type',
        'Permission Role',
        'Permission Email',
        'Permission Domain',
        'Permission Display Name'
      ];
      
      auditSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
      
      auditSheet.setFrozenRows(1);
      Logger.log('Headers created successfully');
      
      auditState.phase = 'PROCESSING';
      scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
    }
    
    // PROCESSING PHASE
    if (auditState.phase === 'PROCESSING') {
      Logger.log('Processing files...');
      updateAuditStatus('RUNNING', 
        'Processing files and permissions...', 
        auditState.filesProcessed, 0);
      
      // Process files in batches, with automatic continuation
      // This number is tuned to stay under the 4.5 minute limit while processing as much as possible
      const BATCH_SIZE = 500; // Process up to 500 files at a time (or until time limit)
      let filesInThisBatch = 0;
      let continueProcessing = true;
      
      while (continueProcessing && filesInThisBatch < BATCH_SIZE) {
        // Check execution time
        const elapsedTime = new Date().getTime() - startTime;
        if (elapsedTime > MAX_EXECUTION_TIME) {
          Logger.log('Approaching timeout limit. Saving state and scheduling continuation...');
          scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
          scheduleAuditContinuation();
          return;
        }
        
        // Get next batch of files
        const filesBatch = getDriveFilesBatch(auditState.pageToken, 100);
        
        if (!filesBatch || !filesBatch.files || filesBatch.files.length === 0) {
          Logger.log('No more files to process');
          auditState.phase = 'FINALIZING';
          break;
        }
        
        Logger.log('Processing batch of ' + filesBatch.files.length + ' files');
        
        // Process files
        const auditData = [];
        filesBatch.files.forEach(function(file) {
          auditState.filesProcessed++;
          filesInThisBatch++;
          
          const permissions = getFilePermissions(file.id);
          
          if (permissions.length === 0) {
            auditData.push([
              file.name,
              file.id,
              file.owners && file.owners.length > 0 ? file.owners[0].emailAddress : 'Unknown',
              getFileType(file),
              file.mimeType,
              file.createdTime ? new Date(file.createdTime) : '',
              file.modifiedTime ? new Date(file.modifiedTime) : '',
              file.size || '',
              file.webViewLink || '',
              0,
              '',
              '',
              '',
              '',
              ''
            ]);
          } else {
            permissions.forEach(function(permission) {
              auditData.push([
                file.name,
                file.id,
                file.owners && file.owners.length > 0 ? file.owners[0].emailAddress : 'Unknown',
                getFileType(file),
                file.mimeType,
                file.createdTime ? new Date(file.createdTime) : '',
                file.modifiedTime ? new Date(file.modifiedTime) : '',
                file.size || '',
                file.webViewLink || '',
                permissions.length,
                permission.type,
                permission.role,
                permission.emailAddress || '',
                permission.domain || '',
                permission.displayName || ''
              ]);
            });
          }
        });
        
        // Write data to sheet
        if (auditData.length > 0) {
          const lastRow = auditSheet.getLastRow();
          auditSheet.getRange(lastRow + 1, 1, auditData.length, 15).setValues(auditData);
          auditState.auditDataCount += auditData.length;
          Logger.log('Wrote ' + auditData.length + ' rows. Total rows: ' + auditState.auditDataCount);
        }
        
        // Update page token
        auditState.pageToken = filesBatch.nextPageToken;
        
        // Update status every 50 files
        if (auditState.filesProcessed % 50 === 0) {
          updateAuditStatus('RUNNING', 'Processing files... ' + auditState.filesProcessed + ' files processed', auditState.filesProcessed, 0);
        }
        
        // Check if there are more pages
        if (!auditState.pageToken) {
          Logger.log('All files processed');
          auditState.phase = 'FINALIZING';
          continueProcessing = false;
        }
        
        // Save state periodically
        scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
      }
      
      // If still processing, schedule continuation
      if (auditState.phase === 'PROCESSING') {
        Logger.log('Batch complete. Files processed in this run: ' + filesInThisBatch);
        Logger.log('Total files processed: ' + auditState.filesProcessed);
        scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
        scheduleAuditContinuation();
        return;
      }
    }
    
    // FINALIZING PHASE
    if (auditState.phase === 'FINALIZING') {
      Logger.log('Finalizing audit...');
      updateAuditStatus('RUNNING', 'Finalizing audit...', auditState.filesProcessed, auditState.filesProcessed);
      
      // Auto-resize columns and add filter
      auditSheet = ss.getSheetByName('Drive Audit');
      const lastRow = auditSheet.getLastRow();
      
      if (lastRow > 1) {
        Logger.log('Auto-resizing columns...');
        for (let i = 1; i <= 15; i++) {
          auditSheet.autoResizeColumn(i);
        }
        
        Logger.log('Adding filter...');
        auditSheet.getRange(1, 1, lastRow, 15).createFilter();
      }
      
      // Create summary
      Logger.log('Creating summary sheet...');
      let summarySheet = ss.getSheetByName('Audit Summary');
      if (summarySheet) {
        summarySheet.clear();
      } else {
        summarySheet = ss.insertSheet('Audit Summary', 0);
      }
      
      createSummary(summarySheet, auditState.filesProcessed, auditState.auditDataCount);
      
      const endTime = new Date();
      const totalDuration = (new Date(endTime) - new Date(auditState.startTime)) / 1000;
      
      Logger.log('=== AUDIT COMPLETED SUCCESSFULLY ===');
      Logger.log('End time: ' + endTime.toISOString());
      Logger.log('Total duration: ' + totalDuration + ' seconds');
      Logger.log('Files audited: ' + auditState.filesProcessed);
      Logger.log('Permission entries: ' + auditState.auditDataCount);
      
      // Update status to COMPLETED
      updateAuditStatus('COMPLETED', 
        'Audit completed successfully! ' +
        'Files audited: ' + auditState.filesProcessed + ', ' +
        'Permission entries: ' + auditState.auditDataCount + ', ' +
        'Duration: ' + Math.round(totalDuration) + ' seconds',
        auditState.filesProcessed, auditState.filesProcessed);
      
      // Clear audit state
      scriptProps.deleteProperty('AUDIT_STATE');
      scriptProps.deleteProperty('AUDIT_PAGE_TOKEN');
      
      // Delete continuation triggers
      deleteContinuationTriggers();
      
      Logger.log('Audit complete and state cleared');
    }
      
  } catch (error) {
    Logger.log('=== ERROR DURING AUDIT ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    // Update status to ERROR
    updateAuditStatus('ERROR', 'An error occurred during the audit: ' + error.toString(), 0, 0);
    
    // Clear state on error
    scriptProps.deleteProperty('AUDIT_STATE');
    scriptProps.deleteProperty('AUDIT_PAGE_TOKEN');
    deleteContinuationTriggers();
  }
}

/**
 * Schedules the audit to continue
 * Uses 1-minute intervals for fast processing
 */
function scheduleAuditContinuation() {
  Logger.log('Scheduling audit continuation...');
  
  try {
    // Delete any existing continuation triggers
    deleteContinuationTriggers();
    
    // Use 1-minute intervals for fast continuation
    const CONTINUATION_DELAY_MINUTES = 1;
    
    ScriptApp.newTrigger('processDriveAuditBatch')
      .timeBased()
      .after(CONTINUATION_DELAY_MINUTES * 60 * 1000)
      .create();
    
    Logger.log('Audit continuation scheduled for ' + CONTINUATION_DELAY_MINUTES + ' minute from now');
    updateAuditStatus('RUNNING', 
      'Audit in progress. Will continue automatically in ' + CONTINUATION_DELAY_MINUTES + ' minute. ' +
      'Check back later for results.',
      0, 0);
    
  } catch (error) {
    Logger.log('Error scheduling continuation: ' + error.toString());
  }
}

/**
 * Deletes all continuation triggers
 */
function deleteContinuationTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processDriveAuditBatch') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Gets a batch of files from Google Drive
 */
function getDriveFilesBatch(pageToken, pageSize) {
  Logger.log('Fetching batch of files. PageToken: ' + (pageToken || 'null'));
  
  try {
    const response = Drive.Files.list({
      pageSize: pageSize || 100,
      fields: 'nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime, size, webViewLink, permissions)',
      pageToken: pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    
    Logger.log('Retrieved ' + (response.files ? response.files.length : 0) + ' files');
    return response;
    
  } catch (error) {
    Logger.log('ERROR fetching files batch: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return null;
  }
}

/**
 * Gets all files from Google Drive (legacy - kept for compatibility)
 * Note: For large Drive accounts, use processDriveAuditBatch instead
 */
function getAllDriveFiles() {
  Logger.log('Fetching files from Drive API...');
  const files = [];
  let pageToken = null;
  let pageNumber = 0;
  
  do {
    try {
      pageNumber++;
      Logger.log('Fetching page ' + pageNumber + ' (100 files per page)...');
      
      const response = Drive.Files.list({
        pageSize: 100,
        fields: 'nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime, size, webViewLink, permissions)',
        pageToken: pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true
      });
      
      if (response.files && response.files.length > 0) {
        files.push(...response.files);
        Logger.log('Page ' + pageNumber + ' retrieved: ' + response.files.length + ' files (Total so far: ' + files.length + ')');
      } else {
        Logger.log('Page ' + pageNumber + ' contained no files');
      }
      
      pageToken = response.nextPageToken;
      
      if (pageToken) {
        Logger.log('More pages available, continuing...');
      } else {
        Logger.log('No more pages to fetch');
      }
      
    } catch (error) {
      Logger.log('ERROR fetching files on page ' + pageNumber + ': ' + error.toString());
      Logger.log('Stack trace: ' + error.stack);
      break;
    }
  } while (pageToken);
  
  Logger.log('File fetching complete. Total files: ' + files.length);
  return files;
}

/**
 * Gets permissions for a specific file
 */
function getFilePermissions(fileId) {
  const permissions = [];
  
  try {
    const response = Drive.Permissions.list(fileId, {
      fields: 'permissions(id, type, role, emailAddress, domain, displayName)',
      supportsAllDrives: true
    });
    
    if (response.permissions) {
      return response.permissions;
    }
  } catch (error) {
    Logger.log('WARNING: Error fetching permissions for file ' + fileId + ': ' + error.toString());
    // Return null to indicate an error occurred (vs empty array for no permissions)
    return [];
  }
  
  return permissions;
}

/**
 * Determines the file type based on MIME type
 */
function getFileType(file) {
  const mimeType = file.mimeType;
  
  if (mimeType === 'application/vnd.google-apps.folder') {
    return 'Folder';
  } else if (mimeType.startsWith('application/vnd.google-apps.')) {
    return 'Google ' + mimeType.replace('application/vnd.google-apps.', '').replace('-', ' ');
  } else if (mimeType.startsWith('image/')) {
    return 'Image';
  } else if (mimeType.startsWith('video/')) {
    return 'Video';
  } else if (mimeType.startsWith('audio/')) {
    return 'Audio';
  } else if (mimeType.includes('pdf')) {
    return 'PDF';
  } else if (mimeType.includes('document') || mimeType.includes('text')) {
    return 'Document';
  } else if (mimeType.includes('spreadsheet')) {
    return 'Spreadsheet';
  } else if (mimeType.includes('presentation')) {
    return 'Presentation';
  } else {
    return 'File';
  }
}

/**
 * Creates a summary sheet with audit statistics
 */
function createSummary(sheet, totalFiles, totalPermissions) {
  Logger.log('Generating summary data...');
  
  const summaryData = [
    ['Drive Audit Summary', ''],
    ['', ''],
    ['Audit Date:', new Date()],
    ['Total Files Audited:', totalFiles],
    ['Total Permission Entries:', totalPermissions],
    ['', ''],
    ['Next Steps:', ''],
    ['1. Review the "Drive Audit" sheet for detailed permissions', ''],
    ['2. Use filters to find files with specific sharing settings', ''],
    ['3. Look for files shared with "anyone" or external domains', ''],
    ['4. Set up a weekly schedule to run audits automatically', '']
  ];
  
  sheet.getRange(1, 1, summaryData.length, 2).setValues(summaryData);
  Logger.log('Summary data written to sheet');
  
  // Format summary
  Logger.log('Formatting summary sheet...');
  sheet.getRange(1, 1, 1, 2).merge()
    .setFontSize(16)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  sheet.getRange(3, 1, 3, 1).setFontWeight('bold');
  sheet.getRange(7, 1, 1, 1).setFontWeight('bold').setFontSize(12);
  
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 200);
  Logger.log('Summary formatting complete');
}

/**
 * Shows dialog to set up scheduled audit
 */
function showScheduleDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Setup Weekly Scheduled Audit',
    'Do you want to set up a weekly audit that runs automatically?\n\n' +
    'The audit will run every Monday at 6:00 AM and update the sheets.\n\n' +
    'Note: You can remove the schedule anytime from the menu.',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    setupWeeklyTrigger();
  }
}

/**
 * Sets up a weekly trigger to run the audit automatically
 */
function setupWeeklyTrigger() {
  Logger.log('Setting up weekly trigger...');
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Delete existing triggers for this function
    const triggers = ScriptApp.getProjectTriggers();
    Logger.log('Checking for existing triggers. Found ' + triggers.length + ' total triggers');
    
    let deletedCount = 0;
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'runDriveAudit') {
        Logger.log('Deleting existing trigger: ' + trigger.getUniqueId());
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    Logger.log('Deleted ' + deletedCount + ' existing runDriveAudit triggers');
    
    // Create new weekly trigger (every Monday at 6 AM)
    Logger.log('Creating new weekly trigger for Monday at 6:00 AM...');
    const newTrigger = ScriptApp.newTrigger('runDriveAudit')
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(6)
      .create();
    
    Logger.log('Trigger created successfully. Trigger ID: ' + newTrigger.getUniqueId());
    
    ui.alert('Success', 
      'Weekly audit has been scheduled!\n\n' +
      'The audit will run automatically every Monday at 6:00 AM.\n\n' +
      'You can:\n' +
      '• Run it manually anytime from the Add-ons menu\n' +
      '• Remove the schedule from "Remove Schedule" menu option',
      ui.ButtonSet.OK);
      
  } catch (error) {
    Logger.log('ERROR setting up trigger: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    ui.alert('Error', 
      'Failed to set up the scheduled audit:\n' + error.toString() + '\n\n' +
      'Make sure you have authorized the necessary permissions.',
      ui.ButtonSet.OK);
  }
}

/**
 * Removes scheduled audit triggers (not continuation triggers)
 */
function removeScheduledAudits() {
  Logger.log('Removing scheduled audit triggers...');
  const ui = SpreadsheetApp.getUi();
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    triggers.forEach(function(trigger) {
      // Only delete runDriveAudit triggers (not processDriveAuditBatch continuation triggers)
      if (trigger.getHandlerFunction() === 'runDriveAudit') {
        Logger.log('Deleting trigger: ' + trigger.getUniqueId());
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    if (deletedCount > 0) {
      Logger.log('Deleted ' + deletedCount + ' scheduled audit trigger(s)');
      ui.alert('Success', 
        'Scheduled weekly audit has been removed.\n\n' +
        'Note: Any currently running audit will continue to completion.\n' +
        'You can still run audits manually from the Add-ons menu.',
        ui.ButtonSet.OK);
    } else {
      Logger.log('No scheduled audit triggers found');
      ui.alert('No Schedule Found', 
        'There are no scheduled audits to remove.\n\n' +
        'Use "Setup Weekly Schedule" to create one.',
        ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('ERROR removing triggers: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    ui.alert('Error', 
      'Failed to remove scheduled audits:\n' + error.toString(),
      ui.ButtonSet.OK);
  }
}

/**
 * Cancels a currently running audit
 */
function cancelRunningAudit() {
  Logger.log('User requested to cancel running audit');
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();
  
  // Check if there's an audit running
  const auditState = scriptProps.getProperty('AUDIT_STATE');
  
  if (!auditState) {
    ui.alert('No Running Audit', 
      'There is no audit currently running.\n\n' +
      'If you recently started an audit, it may have already completed.',
      ui.ButtonSet.OK);
    return;
  }
  
  // Confirm cancellation
  const result = ui.alert(
    'Cancel Running Audit?',
    'Are you sure you want to cancel the currently running audit?\n\n' +
    '⚠️ This will:\n' +
    '• Stop the audit process\n' +
    '• Clear any partial results\n' +
    '• Remove scheduled continuation triggers\n\n' +
    'Note: Results already written to the sheet will remain.',
    ui.ButtonSet.YES_NO
  );
  
  if (result === ui.Button.YES) {
    try {
      // Delete audit state
      scriptProps.deleteProperty('AUDIT_STATE');
      scriptProps.deleteProperty('AUDIT_PAGE_TOKEN');
      
      // Delete continuation triggers
      deleteContinuationTriggers();
      Logger.log('Deleted continuation triggers');
      
      // Update status
      updateAuditStatus('CANCELLED', 
        'Audit was cancelled by user. You can run a new audit anytime from the Add-ons menu.',
        0, 0);
      
      Logger.log('Audit cancelled successfully');
      
      ui.alert('Audit Cancelled', 
        '✅ The running audit has been cancelled.\n\n' +
        'Any results already written to the sheet will remain.\n' +
        'You can start a new audit anytime from "Run Audit Now".',
        ui.ButtonSet.OK);
        
    } catch (error) {
      Logger.log('ERROR cancelling audit: ' + error.toString());
      Logger.log('Stack trace: ' + error.stack);
      ui.alert('Error', 
        'Failed to cancel the audit:\n' + error.toString(),
        ui.ButtonSet.OK);
    }
  } else {
    Logger.log('User chose not to cancel audit');
  }
}

/**
 * Shows the current audit status
 */
function showAuditStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const statusSheet = ss.getSheetByName('Audit Status');
  
  if (!statusSheet) {
    ui.alert('No Audit Status', 
      'No audit has been run yet.\n\n' +
      'Click "Run Audit Now" to start your first audit.',
      ui.ButtonSet.OK);
    return;
  }
  
  try {
    const statusData = statusSheet.getRange('B3:B8').getValues();
    const status = statusData[0][0] || 'UNKNOWN';
    const lastUpdated = statusData[1][0] || 'Never';
    const message = statusData[2][0] || 'No message';
    
    let icon = '❓';
    if (status === 'RUNNING') {
      icon = '⏳';
    } else if (status === 'COMPLETED') {
      icon = '✅';
    } else if (status === 'ERROR') {
      icon = '❌';
    } else if (status === 'CANCELLED') {
      icon = '🛑';
    }
    
    let displayMessage = icon + ' Status: ' + status + '\n\n' +
                        'Last Updated: ' + lastUpdated + '\n\n' +
                        'Message:\n' + message;
    
    if (status === 'RUNNING') {
      displayMessage += '\n\n⏱️ Still running... Check back in a few hours.\n📊 See the "Audit Status" sheet for real-time progress.\n🛑 Use "Cancel Running Audit" to stop it.';
    }
    
    ui.alert('Audit Status', displayMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('Error reading status: ' + error.toString());
    ui.alert('Error', 'Could not read audit status.', ui.ButtonSet.OK);
  }
}


/**
 * Shows information about the add-on
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Drive Audit - Open Source',
    'Version: 2.0.0 (Open Source)\n\n' +
    'This open-source tool audits your Google Drive files and their permissions.\n\n' +
    'Features:\n' +
    '• Lists all files you have access to\n' +
    '• Shows detailed permission information\n' +
    '• Identifies who has access to each file\n' +
    '• Real-time status tracking\n' +
    '• Scheduled weekly audits\n' +
    '• Fast automatic continuation (1-minute intervals)\n' +
    '• Automatic continuation for large Drive accounts\n\n' +
    'Use the filters in the audit sheet to find:\n' +
    '• Files shared with "anyone"\n' +
    '• Files shared externally\n' +
    '• Files with specific roles (viewer, editor, etc.)\n\n' +
    'This is free and open-source software.',
    ui.ButtonSet.OK
  );
}

