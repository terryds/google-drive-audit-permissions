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

// ============ LICENSING ============

const LICENSE_CHECKOUT_URL = 'https://terrydjony.lemonsqueezy.com/checkout/buy/07b31196-a486-47a7-a527-df65ce09ea8d';
const FREE_AUDIT_LIMIT = 2;
const SUPPORT_EMAIL = 'driveauditr@terrydjony.com';

const LICENSE_BENEFITS =
  '• Unlimited audits (free version: ' + FREE_AUDIT_LIMIT + ' total)\n' +
  '• Weekly scheduled audits\n' +
  '• Prioritized customer support: ' + SUPPORT_EMAIL + '\n' +
  '  (reply within 1-2 business days max)';

/**
 * Verifies a license code.
 * Basic check for now; will be replaced with a proper license API later.
 */
function isValidLicenseCode(code) {
  return code === 'f71e606bf83512828fc58fa35db31c15';
}

/**
 * Whether the current user has activated a valid license.
 */
function isLicensed() {
  const code = PropertiesService.getUserProperties().getProperty('LICENSE_CODE');
  return !!code && isValidLicenseCode(code);
}

/**
 * How many free audits the current user has run.
 */
function getFreeAuditsUsed() {
  return parseInt(PropertiesService.getUserProperties().getProperty('FREE_AUDITS_USED') || '0', 10);
}

/**
 * Menu entry point: opens the license dialog (or shows status if already active).
 */
function activateLicense() {
  if (isLicensed()) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('License Active',
      '✅ Your license is already active. Thank you for your support!\n\n' +
      'You have access to:\n' + LICENSE_BENEFITS,
      ui.ButtonSet.OK);
    return;
  }
  showLicenseDialog('', 'none');
}

/**
 * Opens the license HTML dialog (clickable checkout link + code entry).
 *
 * @param {string} message    - contextual banner explaining why the dialog
 *                              appeared ('' for none)
 * @param {string} nextAction - what to do after successful activation:
 *                              'none' | 'schedule' | 'audit:<scope>'
 */
function showLicenseDialog(message, nextAction) {
  const template = HtmlService.createTemplateFromFile('LicenseDialog');
  template.message = message || '';
  template.nextAction = nextAction || 'none';
  template.checkoutUrl = LICENSE_CHECKOUT_URL;
  template.freeLimit = FREE_AUDIT_LIMIT;
  template.supportEmail = SUPPORT_EMAIL;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(400),
    'Drive Audit License');
}

/**
 * Called from the license dialog (google.script.run) to validate and
 * store a license code for the current user.
 */
function activateLicenseWithCode(code) {
  code = (code || '').trim();

  if (!isValidLicenseCode(code)) {
    return {
      success: false,
      message: '❌ That code doesn\'t look right. Check the code in your ' +
               'purchase receipt email and make sure it was copied completely.'
    };
  }

  PropertiesService.getUserProperties().setProperty('LICENSE_CODE', code);
  Logger.log('License activated');
  return { success: true };
}

/**
 * Called from the license dialog after a successful activation to
 * resume whatever the user was trying to do when they hit the paywall.
 */
function continueAfterActivation(nextAction) {
  if (!isLicensed()) {
    return; // safety: never resume gated actions without a license
  }

  if (nextAction === 'schedule') {
    showScheduleConfirm();
  } else if (nextAction && nextAction.indexOf('audit:') === 0) {
    runDriveAudit(nextAction.split(':')[1]);
  }
}

/**
 * Creates a custom menu in Google Sheets when the script is opened.
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Label the license menu item by state. UserProperties may not be
  // accessible in limited auth modes, so fall back to the default label.
  let licenseLabel = '🔑 Activate License Code';
  try {
    if (isLicensed()) {
      licenseLabel = '✅ License: Active';
    }
  } catch (err) {
    Logger.log('Could not read license state in onOpen: ' + err.toString());
  }

  ui.createMenu('Drive Audit')
    .addSubMenu(ui.createMenu('Run Audit Now')
      .addItem('All Drives (including shared)', 'runDriveAuditAll')
      .addItem('My Drive only (skip shared drives)', 'runDriveAuditMyDrive')
      .addItem('Only files I own', 'runDriveAuditOwned'))
    .addItem('Check Audit Status', 'showAuditStatus')
    .addItem('Cancel Running Audit', 'cancelRunningAudit')
    .addSeparator()
    .addItem('Setup Weekly Schedule', 'showScheduleDialog')
    .addItem('Remove Schedule', 'removeScheduledAudits')
    .addSeparator()
    .addItem(licenseLabel, 'activateLicense')
    .addSeparator()
    .addItem('Tutorial', 'showTutorial')
    // .addItem('Support the Creator', 'showSupportCreator')
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
 * Menu wrappers that start an audit with a specific scope.
 * - 'all'     : every file you can access, including shared/organizational drives (original behavior)
 * - 'myDrive' : only files in your My Drive (skips shared/organizational drives)
 * - 'owned'   : only files you own
 */
function runDriveAuditAll() { startAuditFromMenu('all'); }
function runDriveAuditMyDrive() { startAuditFromMenu('myDrive'); }
function runDriveAuditOwned() { startAuditFromMenu('owned'); }

/**
 * License gate for audits started from the menu.
 * Licensed users run unlimited audits; free users get FREE_AUDIT_LIMIT total,
 * tracked per user in UserProperties. Scheduled triggers call runDriveAudit
 * directly and are not counted (the schedule itself requires a license).
 */
function startAuditFromMenu(scope) {
  if (isLicensed()) {
    runDriveAudit(scope);
    return;
  }

  const used = getFreeAuditsUsed();

  if (used >= FREE_AUDIT_LIMIT) {
    // Free audits exhausted. If they activate a license in the dialog,
    // the audit they asked for starts automatically.
    showLicenseDialog(
      '🔒 You\'ve used your ' + FREE_AUDIT_LIMIT + ' free audits. ' +
      'Activate a license to keep auditing — your audit will start right after activation.',
      'audit:' + scope);
    return;
  }

  PropertiesService.getUserProperties().setProperty('FREE_AUDITS_USED', String(used + 1));
  const remaining = FREE_AUDIT_LIMIT - used - 1;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      remaining > 0
        ? 'You have ' + remaining + ' free audit' + (remaining === 1 ? '' : 's') + ' left after this one.'
        : 'This is your last free audit. Activate a license from the Drive Audit menu for unlimited audits.',
      'Free audit ' + (used + 1) + ' of ' + FREE_AUDIT_LIMIT,
      10
    );
  } catch (err) {
    Logger.log('Could not show free-audit toast: ' + err.toString());
  }

  runDriveAudit(scope);
}

/**
 * Main function to audit Google Drive files and permissions
 * This version handles timeouts by processing in batches
 *
 * @param {string} scope - 'all' | 'myDrive' | 'owned'. Defaults to 'all'
 *   (e.g. when invoked by the scheduled weekly trigger).
 */
function runDriveAudit(scope) {
  if (scope !== 'myDrive' && scope !== 'owned') {
    scope = 'all';
  }

  // Clear any previous audit state and record the requested scope
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.deleteProperty('AUDIT_STATE');
  scriptProps.deleteProperty('AUDIT_PAGE_TOKEN');
  scriptProps.setProperty('AUDIT_SCOPE', scope);

  Logger.log('=== DRIVE AUDIT STARTED (FRESH) ===');
  Logger.log('Scope: ' + scope);
  Logger.log('Start time: ' + new Date().toISOString());
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Update status to RUNNING
  updateAuditStatus('RUNNING', 
    'Audit is in progress. For large Drive accounts, this may take some time. The audit will automatically continue every minute if needed.', 
    0, 0);
  
  // Show progress message
  const scopeLabel = scope === 'myDrive'
    ? 'My Drive only (shared drives skipped)'
    : (scope === 'owned' ? 'Only files you own' : 'All Drives (including shared)');
  ui.alert('Drive Audit',
    'Scope: ' + scopeLabel + '\n\n' +
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
        scope: scriptProps.getProperty('AUDIT_SCOPE') || 'all',
        totalFilesFound: 0,
        filesProcessed: 0,
        auditDataCount: 0,
        pageToken: null,
        startTime: new Date().toISOString()
      };
      Logger.log('Audit scope: ' + auditState.scope);
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
        'Folder Path',
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

      // Cache of folder id -> folder resource, to avoid re-fetching the same
      // ancestor folders when resolving full folder paths within this run.
      const folderCache = {};
      
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
        const filesBatch = getDriveFilesBatch(auditState.pageToken, 100, auditState.scope);
        
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
          const folderPath = getFolderPath(file.parents, folderCache);

          if (permissions.length === 0) {
            auditData.push([
              file.name,
              folderPath,
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
                folderPath,
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
          auditSheet.getRange(lastRow + 1, 1, auditData.length, 16).setValues(auditData);
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
        for (let i = 1; i <= 16; i++) {
          auditSheet.autoResizeColumn(i);
        }

        Logger.log('Adding filter...');
        auditSheet.getRange(1, 1, lastRow, 16).createFilter();
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
      scriptProps.deleteProperty('AUDIT_SCOPE');

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
    scriptProps.deleteProperty('AUDIT_SCOPE');
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
 *
 * @param {string} pageToken - Drive API page token (null for first page)
 * @param {number} pageSize  - files per page
 * @param {string} scope     - 'all' | 'myDrive' | 'owned'
 */
function getDriveFilesBatch(pageToken, pageSize, scope) {
  scope = scope || 'all';
  Logger.log('Fetching batch of files. Scope: ' + scope + ', PageToken: ' + (pageToken || 'null'));

  try {
    const params = {
      pageSize: pageSize || 100,
      fields: 'nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime, size, webViewLink, permissions, parents)',
      pageToken: pageToken
    };

    if (scope === 'all') {
      // Original behavior: include shared/organizational drives
      params.supportsAllDrives = true;
      params.includeItemsFromAllDrives = true;
    } else {
      // 'myDrive' and 'owned': restrict to the user's own corpus, no shared drives
      params.corpora = 'user';
      params.supportsAllDrives = false;
      params.includeItemsFromAllDrives = false;
      if (scope === 'owned') {
        params.q = "'me' in owners";
      }
    }

    const response = Drive.Files.list(params);

    Logger.log('Retrieved ' + (response.files ? response.files.length : 0) + ' files');
    return response;

  } catch (error) {
    Logger.log('ERROR fetching files batch: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return null;
  }
}

/**
 * Resolves the full folder path for a file from its parent id(s),
 * e.g. "My Drive/Projects/2026/Reports".
 *
 * Walks up the parent chain, caching each folder lookup in folderCache
 * (keyed by folder id) so repeated ancestors aren't re-fetched. Returns
 * an empty string for files with no parent (e.g. items shared with you
 * that don't live in a folder you can traverse).
 *
 * @param {string[]} parents     - file.parents from the Drive API
 * @param {Object}   folderCache - id -> folder resource (or null) cache
 */
function getFolderPath(parents, folderCache) {
  if (!parents || parents.length === 0) {
    return '';
  }

  const pathParts = [];
  let currentId = parents[0]; // a file is normally in a single folder
  let depth = 0;
  const MAX_DEPTH = 100; // safety guard against unexpected cycles

  while (currentId && depth < MAX_DEPTH) {
    depth++;

    let folder = folderCache[currentId];
    if (folder === undefined) {
      try {
        folder = Drive.Files.get(currentId, {
          fields: 'id, name, parents',
          supportsAllDrives: true
        });
      } catch (error) {
        Logger.log('WARNING: could not resolve folder ' + currentId + ': ' + error.toString());
        folder = null;
      }
      folderCache[currentId] = folder;
    }

    if (!folder) {
      break;
    }

    pathParts.unshift(folder.name);
    currentId = (folder.parents && folder.parents.length > 0) ? folder.parents[0] : null;
  }

  return pathParts.join('/');
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
        fields: 'nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime, size, webViewLink, permissions, parents)',
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
  if (!isLicensed()) {
    showLicenseDialog(
      '🔒 Weekly scheduled audits require a license. ' +
      'Schedule setup will continue right after activation.',
      'schedule');
    return;
  }
  showScheduleConfirm();
}

/**
 * The actual schedule setup confirmation (license already verified).
 */
function showScheduleConfirm() {
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
      scriptProps.deleteProperty('AUDIT_SCOPE');

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
 * Shows the tutorial link
 */
function showTutorial() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:\'Google Sans\',Roboto,Arial,sans-serif;font-size:14px;color:#202124;padding:4px 4px 0;">' +
    '<p style="margin-top:0;">To learn how to use Drive Audit, visit the getting started guide:</p>' +
    '<p><a href="https://driveauditr.com/docs/getting-started/" target="_blank" rel="noopener">' +
    'https://driveauditr.com/docs/getting-started/</a></p>' +
    '</div>')
    .setWidth(400)
    .setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tutorial');
}

/**
 * Shows support/donation information
 */
// function showSupportCreator() {
//   const ui = SpreadsheetApp.getUi();
//   ui.alert(
//     'Support the Creator',
//     'If you find Drive Audit helpful, please consider supporting the creator ' +
//     'Terry Djony via Ko-fi so I can keep maintaining this script!\n\n' +
//     'https://ko-fi.com/terrydjony\n\n' +
//     'Copy and paste the link above into your browser.\n\n' +
//     'Thank you for your support!',
//     ui.ButtonSet.OK
//   );
// }

/**
 * Shows information about the add-on
 */
function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Drive Audit',
    'Version: 2.0.0\n\n' +
    'This tool audits your Google Drive files and their permissions.\n\n' +
    'Features:\n' +
    '• Lists all files you have access to\n' +
    '• Shows detailed permission information\n' +
    '• Identifies who has access to each file\n' +
    '• Real-time status tracking\n' +
    '• Scheduled weekly audits (license)\n' +
    '• Fast automatic continuation (1-minute intervals)\n' +
    '• Automatic continuation for large Drive accounts\n\n' +
    'Use the filters in the audit sheet to find:\n' +
    '• Files shared with "anyone"\n' +
    '• Files shared externally\n' +
    '• Files with specific roles (viewer, editor, etc.)\n\n' +
    'Free version includes ' + FREE_AUDIT_LIMIT + ' audits. A license unlocks:\n' +
    LICENSE_BENEFITS + '\n\n' +
    'Get a license: use "Activate License Code" in the menu.',
    ui.ButtonSet.OK
  );
}

