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
const FREE_FILE_LIMIT = 15; // free tier: results are capped at this many files
const LICENSE_PRICE = '$12'; // one-time payment for lifetime access (not a subscription)
const SUPPORT_EMAIL = 'driveauditr@terrydjony.com';

const LICENSE_BENEFITS =
  '• Full audit results (free tier shows only the first ' + FREE_FILE_LIMIT + ' files)\n' +
  '• Weekly scheduled audits\n' +
  '• Prioritized customer support: ' + SUPPORT_EMAIL + '\n' +
  '  (reply within 1-2 business days max)\n' +
  '• Lifetime access for a one-time ' + LICENSE_PRICE + ' payment';

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
 *                              'none' | 'schedule'
 */
function showLicenseDialog(message, nextAction) {
  const template = HtmlService.createTemplateFromFile('LicenseDialog');
  template.message = message || '';
  template.nextAction = nextAction || 'none';
  template.checkoutUrl = LICENSE_CHECKOUT_URL;
  template.freeLimit = FREE_FILE_LIMIT;
  template.price = LICENSE_PRICE;
  template.supportEmail = SUPPORT_EMAIL;

  SpreadsheetApp.getUi().showModalDialog(
    template.evaluate().setWidth(420).setHeight(450),
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
    .addItem('Resume Stalled Audit', 'resumeAudit')
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
      ['Results Location:', 'The "Drive Audit" sheet — results appear there as the audit runs'],
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
    } else if (status === 'PAUSED') {
      statusCell.setBackground('#fce8b2').setFontColor('#b45309');
    } else if (status === 'FREE LIMIT HIT') {
      statusCell.setBackground('#fce8b2').setFontColor('#b45309').setFontWeight('bold');
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
function runDriveAuditAll() { runDriveAudit('all'); }
function runDriveAuditMyDrive() { runDriveAudit('myDrive'); }
function runDriveAuditOwned() { runDriveAudit('owned'); }

// ============ AUDIT ENGINE ============
//
// An audit of a large Drive takes far longer than one Apps Script execution
// is allowed to run (6 minutes), so it runs as a series of executions that
// each pick up where the last one stopped. Progress lives in the
// AUDIT_STATE script property. Two triggers keep the series going:
//
// - a one-shot "continuation" trigger, created at the end of each
//   execution, which starts the next one a minute later;
// - a recurring "watchdog" trigger, which restarts the series if it ever
//   breaks: an execution killed by Google, a continuation trigger that
//   could not be created, or the daily trigger-runtime quota running out
//   (90 minutes for consumer accounts) and resetting the next day.

// Each execution stops itself after this long, leaving a buffer before
// Google's hard 6-minute kill.
const MAX_EXECUTION_TIME_MS = 4.5 * 60 * 1000;
const CONTINUATION_DELAY_MINUTES = 1;
const WATCHDOG_INTERVAL_MINUTES = 10;
// Consecutive failed executions before the audit gives up with an ERROR.
const MAX_CONSECUTIVE_ERRORS = 5;
const STATUS_UPDATE_INTERVAL_MS = 30 * 1000;
// Upper bound on the time an execution spends preloading folders.
const FOLDER_PRELOAD_TIME_MS = 60 * 1000;

/**
 * Main function to audit Google Drive files and permissions.
 * Starts a fresh audit; the work itself happens in runAuditBatch.
 *
 * @param {string} scope - 'all' | 'myDrive' | 'owned'. Defaults to 'all'
 *   (e.g. when invoked by the scheduled weekly trigger).
 */
function runDriveAudit(scope) {
  if (scope !== 'myDrive' && scope !== 'owned') {
    scope = 'all';
  }

  // The weekly schedule calls this from a time-based trigger, where there
  // is no UI to show a dialog in (getUi() throws there).
  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (err) {
    Logger.log('No UI available (running from a trigger): ' + err.toString());
  }

  // Stop any previous audit. If one of its executions is still running, it
  // notices its state is gone and stops without saving.
  const scriptProps = PropertiesService.getScriptProperties();
  clearAuditState(scriptProps);
  deleteAuditTriggers();

  Logger.log('=== DRIVE AUDIT STARTED (FRESH) ===');
  Logger.log('Scope: ' + scope);
  Logger.log('Start time: ' + new Date().toISOString());

  // Update status to RUNNING
  updateAuditStatus('RUNNING',
    'Audit is starting. For large Drive accounts this may take some time; it continues automatically in the background.',
    0, 0);

  // Show progress message
  if (ui) {
    const scopeLabel = scope === 'myDrive'
      ? 'My Drive only (shared drives skipped)'
      : (scope === 'owned' ? 'Only files you own' : 'All Drives (including shared)');
    ui.alert('Drive Audit',
      'Scope: ' + scopeLabel + '\n\n' +
      'Starting audit... This may take several minutes depending on the number of files.\n\n' +
      '⏳ The audit will run in the background.\n' +
      '📊 Check the "Audit Status" sheet for progress.\n' +
      '📄 Results will appear in the "Drive Audit" sheet when ready.\n' +
      '⏱️ For large Drive accounts (1000+ files):\n' +
      '   • Works in chunks of about 5 minutes and continues automatically\n' +
      '   • Picks itself back up if a chunk fails or Google\'s daily quota runs out\n' +
      '   • You can close this sheet while it runs\n\n' +
      'Click OK to start.',
      ui.ButtonSet.OK);
  }

  const auditState = {
    runId: Utilities.getUuid(),
    phase: 'SETUP',
    scope: scope,
    filesProcessed: 0,
    auditDataCount: 0,
    pageToken: null,
    pageOffset: 0,
    nextRow: 2,
    startTime: new Date().toISOString()
  };
  scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
  ensureWatchdogTrigger();

  // An execution of the previous audit may still hold the lock for a few
  // seconds. If it doesn't let go in time, start through a trigger instead.
  if (!runAuditBatch(30 * 1000)) {
    scheduleAuditContinuation(auditState);
  }
}

/**
 * Menu entry point: restarts the background processing of an audit that
 * has stopped making progress, continuing from where it left off.
 */
function resumeAudit() {
  const ui = SpreadsheetApp.getUi();
  const scriptProps = PropertiesService.getScriptProperties();

  let auditState = null;
  try {
    auditState = JSON.parse(scriptProps.getProperty('AUDIT_STATE'));
  } catch (error) {
    auditState = null;
  }

  if (!auditState) {
    ui.alert('No Audit To Resume',
      'There is no unfinished audit.\n\n' +
      'Use "Run Audit Now" to start a new one.',
      ui.ButtonSet.OK);
    return;
  }

  if (scheduleAuditContinuation(auditState)) {
    ui.alert('Audit Resumed',
      'The audit will continue in the background within a couple of minutes, ' +
      'from where it stopped.\n\n' +
      'Check the "Audit Status" sheet for progress.',
      ui.ButtonSet.OK);
    return;
  }

  // Background triggers can't be created, so do one chunk of the work
  // right here instead.
  runAuditBatch(1000);
  if (scriptProps.getProperty('AUDIT_STATE')) {
    ui.alert('Audit Partly Resumed',
      'The audit could not be scheduled to run in the background, so one ' +
      'chunk of it was processed just now.\n\n' +
      'Choose "Resume Stalled Audit" again to process the next chunk. ' +
      'Check the "Audit Status" sheet for progress.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('Audit Finished',
      'The audit has ended. Check the "Audit Status" sheet for the result.',
      ui.ButtonSet.OK);
  }
}

/**
 * Handler of the one-shot continuation trigger.
 */
function processDriveAuditBatch() {
  runAuditBatch(1000);
}

/**
 * Handler of the recurring watchdog trigger. Does nothing while another
 * execution is working on the audit.
 */
function auditWatchdog() {
  runAuditBatch(1000);
}

/**
 * Runs one execution's worth of the audit, unless another execution is
 * already doing so.
 *
 * @param {number} lockWaitMs - how long to wait for a running execution
 * @return {boolean} false if another execution held the lock throughout
 */
function runAuditBatch(lockWaitMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(lockWaitMs)) {
    Logger.log('Another execution is processing the audit. Skipping.');
    return false;
  }

  try {
    processAuditBatchLocked();
  } finally {
    lock.releaseLock();
  }
  return true;
}

/**
 * Processes the drive audit in batches to avoid timeouts.
 * Continues from the saved state; only one execution runs this at a time.
 */
function processAuditBatchLocked() {
  Logger.log('=== BATCH PROCESSING STARTED ===');
  Logger.log('Start time: ' + new Date().toISOString());

  const startTime = new Date().getTime();
  const outOfTime = function() {
    return new Date().getTime() - startTime > MAX_EXECUTION_TIME_MS;
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();

  const savedState = scriptProps.getProperty('AUDIT_STATE');
  if (!savedState) {
    // Finished or cancelled: a trigger that outlived its audit.
    Logger.log('No audit in progress. Removing audit triggers.');
    deleteAuditTriggers();
    return;
  }

  let auditState = null;

  try {
    auditState = JSON.parse(savedState);
    Logger.log('Continuing audit from phase: ' + auditState.phase);
    Logger.log('Files processed so far: ' + auditState.filesProcessed);

    let auditSheet = ss.getSheetByName('Drive Audit');

    // State saved by an older version of this script, which didn't track
    // which sheet row each saved position corresponds to. Continuing from
    // it could silently skip or repeat files, so don't.
    if (!auditState.runId) {
      updateAuditStatus('ERROR',
        'This audit was started with an older version of Drive Audit and can\'t be continued. ' +
        'Please start a new one from Drive Audit > Run Audit Now.',
        0, 0);
      clearAuditState(scriptProps);
      deleteAuditTriggers();
      return;
    }

    // Free tier caps results at FREE_FILE_LIMIT files. Re-checked on every
    // execution (continuation triggers run as the user who started the
    // audit, so UserProperties resolves to the right person).
    const licensed = isLicensed();

    // SETUP PHASE
    if (auditState.phase === 'SETUP') {
      Logger.log('Setting up audit sheet...');
      if (auditSheet) {
        Logger.log('Clearing existing "Drive Audit" sheet');
        auditSheet.clear();
        // clear() keeps the previous audit's filter, and a sheet can only
        // have one: finalizing would fail when it creates the new one.
        const previousFilter = auditSheet.getFilter();
        if (previousFilter) {
          previousFilter.remove();
        }
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

      // Creating the "Drive Audit" sheet makes it the active tab, leaving
      // the user staring at an empty sheet while the audit runs. Switch
      // them back to "Audit Status", which shows live progress.
      try {
        const statusSheet = ss.getSheetByName('Audit Status');
        if (statusSheet) {
          ss.setActiveSheet(statusSheet);
        }
      } catch (err) {
        Logger.log('Could not switch to the status sheet: ' + err.toString());
      }

      auditState.phase = 'PROCESSING';
      if (!saveAuditState(scriptProps, auditState)) {
        Logger.log('Audit was cancelled or restarted. Stopping.');
        return;
      }
    }

    // PROCESSING PHASE
    if (auditState.phase === 'PROCESSING') {
      Logger.log('Processing files...');

      if (!auditSheet) {
        throw new Error('The "Drive Audit" sheet is missing. Please run a new audit.');
      }

      // Rows below the saved position were written by an execution that
      // was killed before it could save its progress. They are processed
      // again now, so remove them rather than duplicate them.
      const lastWrittenRow = auditSheet.getLastRow();
      if (lastWrittenRow >= auditState.nextRow) {
        Logger.log('Removing ' + (lastWrittenRow - auditState.nextRow + 1) + ' unsaved rows from an interrupted run');
        auditSheet.getRange(auditState.nextRow, 1, lastWrittenRow - auditState.nextRow + 1, 16).clearContent();
      }

      updateAuditStatus('RUNNING', auditProgressMessage(auditState), auditState.filesProcessed, 0);
      let lastStatusUpdate = new Date().getTime();

      // Folder id -> { name, parent }, for resolving full folder paths.
      // The free preview only needs the folders of a handful of files.
      const folderCache = {};
      if (licensed) {
        preloadFolders(folderCache, auditState.scope);
      }

      // Work through the Drive one page of files at a time until this
      // execution runs out of time. The clock is checked before every file
      // so a slow page can't push the execution past Google's hard limit.
      while (auditState.phase === 'PROCESSING' && !outOfTime()) {
        const filesBatch = getDriveFilesBatch(auditState.pageToken, 100, auditState.scope);
        const files = filesBatch.files || [];

        const auditData = [];
        let fileIndex = auditState.pageOffset || 0;
        let freeLimitReached = false;

        for (; fileIndex < files.length; fileIndex++) {
          if (outOfTime()) {
            break;
          }

          // Free tier: the audit is a preview capped at FREE_FILE_LIMIT
          // files. Stopping with files left over means results were
          // withheld, which the final notice points out.
          if (!licensed && auditState.filesProcessed >= FREE_FILE_LIMIT) {
            freeLimitReached = true;
            break;
          }

          appendAuditRows(auditData, files[fileIndex], folderCache);
          auditState.filesProcessed++;
        }

        // Write data to sheet
        if (auditData.length > 0) {
          auditSheet.getRange(auditState.nextRow, 1, auditData.length, 16).setValues(auditData);
          auditState.nextRow += auditData.length;
          auditState.auditDataCount += auditData.length;
          Logger.log('Wrote ' + auditData.length + ' rows. Total rows: ' + auditState.auditDataCount);
        }

        const stoppedMidPage = fileIndex < files.length;

        if (stoppedMidPage) {
          // Continue this same page from this file next time
          auditState.pageOffset = fileIndex;
          if (freeLimitReached) {
            auditState.truncated = true;
            auditState.phase = 'FINALIZING';
          }
        } else {
          // An empty page is not the end of the list: only a missing
          // nextPageToken is.
          auditState.pageToken = filesBatch.nextPageToken || null;
          auditState.pageOffset = 0;

          if (!auditState.pageToken) {
            Logger.log('All files processed');
            auditState.phase = 'FINALIZING';
          } else if (!licensed && auditState.filesProcessed >= FREE_FILE_LIMIT) {
            // Free tier: preview complete — finish now instead of
            // scanning the rest of the Drive.
            Logger.log('Free tier limit of ' + FREE_FILE_LIMIT + ' files reached. Finalizing.');
            auditState.truncated = true;
            auditState.phase = 'FINALIZING';
          }
        }

        // Save progress. The rows go to the sheet first: if this
        // execution dies in between, the next one redoes this page
        // instead of skipping it.
        auditState.consecutiveErrors = 0;
        SpreadsheetApp.flush();
        if (!saveAuditState(scriptProps, auditState)) {
          Logger.log('Audit was cancelled or restarted. Stopping.');
          return;
        }

        if (new Date().getTime() - lastStatusUpdate > STATUS_UPDATE_INTERVAL_MS) {
          updateAuditStatus('RUNNING', auditProgressMessage(auditState), auditState.filesProcessed, 0);
          lastStatusUpdate = new Date().getTime();
        }
      }

      // If still processing, schedule continuation
      if (auditState.phase === 'PROCESSING') {
        Logger.log('Out of time for this execution. Total files processed: ' + auditState.filesProcessed);
        scheduleAuditContinuation(auditState);
        return;
      }
    }

    // FINALIZING PHASE
    if (auditState.phase === 'FINALIZING') {
      Logger.log('Finalizing audit...');

      // Formatting a very large sheet can be slow enough to get this
      // execution killed, after which the watchdog retries. Each retry
      // does less formatting so that finalizing always gets through.
      auditState.finalizeAttempts = (auditState.finalizeAttempts || 0) + 1;
      if (!saveAuditState(scriptProps, auditState)) {
        Logger.log('Audit was cancelled or restarted. Stopping.');
        return;
      }

      updateAuditStatus('RUNNING', 'Finalizing audit...', auditState.filesProcessed, auditState.filesProcessed);

      // Auto-resize columns and add filter
      auditSheet = ss.getSheetByName('Drive Audit');
      const lastRow = auditState.nextRow - 1;

      if (lastRow > 1) {
        if (auditState.finalizeAttempts === 1) {
          Logger.log('Auto-resizing columns...');
          auditSheet.autoResizeColumns(1, 16);
        }

        if (auditState.finalizeAttempts <= 2) {
          Logger.log('Adding filter...');
          const existingFilter = auditSheet.getFilter();
          if (existingFilter) {
            existingFilter.remove();
          }
          auditSheet.getRange(1, 1, lastRow, 16).createFilter();
        }
      }

      // Free tier: results were truncated — add a prominent notice right
      // below the results. (Placed after the filter so it stays outside
      // the filter range.)
      const showFreeLimitNotice = !licensed && auditState.truncated;

      if (showFreeLimitNotice) {
        Logger.log('Free tier: results truncated. Adding upgrade notice...');
        const noticeRow = lastRow + 2;
        auditSheet.getRange(noticeRow, 1, 1, 16).merge();
        auditSheet.getRange(noticeRow, 1)
          .setValue('🔒 FREE TIER LIMIT: only the first ' + FREE_FILE_LIMIT + ' files are shown above. ' +
            'Activate your license for only ' + LICENSE_PRICE + ' — one-time payment, lifetime access, ' +
            'no subscription — in Drive Audit > Activate License Code to see the full results.')
          .setFontSize(14)
          .setFontWeight('bold')
          .setBackground('#fce8b2')
          .setFontColor('#b45309')
          .setWrap(true)
          .setVerticalAlignment('middle');
        auditSheet.setRowHeight(noticeRow, 55);
      }

      // Create summary
      Logger.log('Creating summary sheet...');
      let summarySheet = ss.getSheetByName('Audit Summary');
      if (summarySheet) {
        summarySheet.clear();
      } else {
        summarySheet = ss.insertSheet('Audit Summary', 0);
      }

      createSummary(summarySheet, auditState.filesProcessed, auditState.auditDataCount, showFreeLimitNotice);

      const endTime = new Date();
      const totalDuration = (new Date(endTime) - new Date(auditState.startTime)) / 1000;

      Logger.log('=== AUDIT COMPLETED SUCCESSFULLY ===');
      Logger.log('End time: ' + endTime.toISOString());
      Logger.log('Total duration: ' + totalDuration + ' seconds');
      Logger.log('Files audited: ' + auditState.filesProcessed);
      Logger.log('Permission entries: ' + auditState.auditDataCount);

      // Update the final status. A truncated free-tier run is reported as
      // FREE LIMIT HIT rather than COMPLETED — the Drive was NOT fully
      // audited, and the status should lead with how to unlock the rest.
      if (showFreeLimitNotice) {
        updateAuditStatus('FREE LIMIT HIT',
          '🔒 Free preview finished: only the first ' + FREE_FILE_LIMIT + ' files were audited. ' +
          'Activate your license for only ' + LICENSE_PRICE + ' (one-time payment, lifetime access) ' +
          'in Drive Audit > Activate License Code, then run the audit again for full access. ' +
          'Permission entries: ' + auditState.auditDataCount + ', ' +
          'Duration: ' + Math.round(totalDuration) + ' seconds',
          auditState.filesProcessed, auditState.filesProcessed);
      } else {
        updateAuditStatus('COMPLETED',
          'Audit completed successfully! ' +
          'Files audited: ' + auditState.filesProcessed + ', ' +
          'Permission entries: ' + auditState.auditDataCount + ', ' +
          'Duration: ' + Math.round(totalDuration) + ' seconds',
          auditState.filesProcessed, auditState.filesProcessed);
      }

      // Clear audit state and stop the triggers
      clearAuditState(scriptProps);
      deleteAuditTriggers();

      Logger.log('Audit complete and state cleared');
    }

  } catch (error) {
    handleAuditError(scriptProps, auditState, error);
  }
}

/**
 * Handles an execution that failed. The audit keeps its progress and is
 * retried from the last saved position; it only ends with an ERROR after
 * MAX_CONSECUTIVE_ERRORS failed executions in a row.
 */
function handleAuditError(scriptProps, auditState, error) {
  Logger.log('=== ERROR DURING AUDIT ===');
  Logger.log('Error: ' + error.toString());
  Logger.log('Stack trace: ' + error.stack);

  // Retry from the progress that was actually saved: the in-memory state
  // may be ahead of it.
  let savedState = null;
  try {
    savedState = JSON.parse(scriptProps.getProperty('AUDIT_STATE'));
  } catch (parseError) {
    savedState = null;
  }

  if (savedState && auditState && savedState.runId !== auditState.runId) {
    Logger.log('Audit was restarted. Leaving the new audit alone.');
    return;
  }

  if (savedState && savedState.runId) {
    savedState.consecutiveErrors = (savedState.consecutiveErrors || 0) + 1;

    if (savedState.consecutiveErrors < MAX_CONSECUTIVE_ERRORS) {
      scriptProps.setProperty('AUDIT_STATE', JSON.stringify(savedState));
      if (!scheduleAuditContinuation(savedState)) {
        return;
      }
      updateAuditStatus('RUNNING',
        'A temporary problem interrupted the audit (attempt ' + savedState.consecutiveErrors +
        ' of ' + MAX_CONSECUTIVE_ERRORS + '): ' + error.toString() + ' — ' +
        'it will retry automatically in ' + CONTINUATION_DELAY_MINUTES + ' minute. ' +
        savedState.filesProcessed + ' files processed so far.',
        0, 0);
      return;
    }
  }

  // Update status to ERROR
  updateAuditStatus('ERROR',
    'The audit stopped because of an error: ' + error.toString() +
    (savedState ? ' — the ' + savedState.filesProcessed + ' files processed before the error are still in the "Drive Audit" sheet.' : ''),
    0, 0);

  // Clear state on error
  clearAuditState(scriptProps);
  deleteAuditTriggers();
}

/**
 * Saves the audit progress, unless the audit was cancelled or replaced by
 * a new one while this execution was running.
 *
 * @return {boolean} false if this execution's audit is no longer current
 */
function saveAuditState(scriptProps, auditState) {
  let current = null;
  try {
    current = JSON.parse(scriptProps.getProperty('AUDIT_STATE'));
  } catch (error) {
    current = null;
  }

  if (!current || current.runId !== auditState.runId) {
    return false;
  }

  scriptProps.setProperty('AUDIT_STATE', JSON.stringify(auditState));
  return true;
}

/**
 * Deletes the saved audit progress.
 */
function clearAuditState(scriptProps) {
  scriptProps.deleteProperty('AUDIT_STATE');
  scriptProps.deleteProperty('AUDIT_PAGE_TOKEN');
  scriptProps.deleteProperty('AUDIT_SCOPE');
}

/**
 * Status message shown while the audit is running.
 */
function auditProgressMessage(auditState) {
  return 'Audit in progress: ' + auditState.filesProcessed + ' files processed so far (' +
    auditState.auditDataCount + ' rows). It continues automatically in the background, ' +
    'so you can close this sheet. If "Last Updated" stops changing for more than 30 minutes, ' +
    'Google\'s daily limit on background script time has probably been reached: the audit ' +
    'picks up again by itself when the limit resets, within 24 hours. Starting a new audit ' +
    'would discard the results so far.';
}

/**
 * Schedules the audit to continue
 * Uses 1-minute intervals for fast processing
 *
 * @return {boolean} false if nothing could be scheduled (audit is paused)
 */
function scheduleAuditContinuation(auditState) {
  Logger.log('Scheduling audit continuation...');

  // Delete any existing continuation triggers
  try {
    deleteContinuationTriggers();
  } catch (error) {
    Logger.log('Error deleting continuation triggers: ' + error.toString());
  }

  const continuationCreated = createAuditTrigger('processDriveAuditBatch', function(builder) {
    return builder.after(CONTINUATION_DELAY_MINUTES * 60 * 1000);
  });
  const watchdogPresent = ensureWatchdogTrigger();

  if (!continuationCreated && !watchdogPresent) {
    // Nothing will wake the audit up again: say so instead of leaving a
    // RUNNING status that never changes.
    updateAuditStatus('PAUSED',
      'The audit could not schedule its next step, so it is paused' +
      (auditState ? ' at ' + auditState.filesProcessed + ' files' : '') + '. ' +
      'Nothing is lost: choose Drive Audit > Resume Stalled Audit to continue.',
      0, 0);
    return false;
  }

  if (auditState) {
    updateAuditStatus('RUNNING', auditProgressMessage(auditState), auditState.filesProcessed, 0);
  }
  return true;
}

/**
 * Makes sure the recurring watchdog trigger exists. It restarts the audit
 * whenever the chain of one-shot continuation triggers breaks.
 *
 * @return {boolean} whether the watchdog trigger is in place
 */
function ensureWatchdogTrigger() {
  try {
    const hasWatchdog = ScriptApp.getProjectTriggers().some(function(trigger) {
      return trigger.getHandlerFunction() === 'auditWatchdog';
    });
    if (hasWatchdog) {
      return true;
    }
  } catch (error) {
    Logger.log('Error listing triggers: ' + error.toString());
    return false;
  }

  return createAuditTrigger('auditWatchdog', function(builder) {
    return builder.everyMinutes(WATCHDOG_INTERVAL_MINUTES);
  });
}

/**
 * Creates a time-based trigger, trying a few times: creating a trigger
 * occasionally fails with a temporary server error.
 *
 * @param {string}   handler   - name of the function the trigger calls
 * @param {Function} configure - sets the schedule on the trigger builder
 * @return {boolean} whether the trigger was created
 */
function createAuditTrigger(handler, configure) {
  const MAX_ATTEMPTS = 3;

  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      configure(ScriptApp.newTrigger(handler).timeBased()).create();
      return true;
    } catch (error) {
      Logger.log('Error creating ' + handler + ' trigger (attempt ' + attempt + '): ' + error.toString());
      if (attempt < MAX_ATTEMPTS) {
        Utilities.sleep(1000 * attempt);
      }
    }
  }
  return false;
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
 * Deletes every trigger that keeps an audit going: the continuation
 * triggers and the watchdog. (The weekly schedule trigger is left alone.)
 */
function deleteAuditTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === 'processDriveAuditBatch' || handler === 'auditWatchdog') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Calls a Drive API function, retrying with a growing delay when the
 * failure is a temporary one (rate limit or server hiccup).
 */
function withRetry(apiCall) {
  const MAX_ATTEMPTS = 4;

  for (let attempt = 1; ; attempt++) {
    try {
      return apiCall();
    } catch (error) {
      const message = String((error && error.message) || error);
      const temporary = /rate limit|quota|backend error|internal error|server error|transient|timed? ?out|try again|unavailable|empty response/i.test(message);
      if (!temporary || attempt >= MAX_ATTEMPTS) {
        throw error;
      }
      Logger.log('Temporary Drive API error (attempt ' + attempt + '): ' + message);
      Utilities.sleep(Math.pow(2, attempt) * 500 + Math.floor(Math.random() * 500));
    }
  }
}

/**
 * Gets a batch of files from Google Drive. Throws if the Drive API keeps
 * failing: a failed request must not be mistaken for the end of the list.
 *
 * @param {string} pageToken - Drive API page token (null for first page)
 * @param {number} pageSize  - files per page
 * @param {string} scope     - 'all' | 'myDrive' | 'owned'
 */
function getDriveFilesBatch(pageToken, pageSize, scope) {
  scope = scope || 'all';
  Logger.log('Fetching batch of files. Scope: ' + scope + ', PageToken: ' + (pageToken || 'null'));

  const params = {
    pageSize: pageSize || 100,
    fields: 'nextPageToken, files(id, name, mimeType, owners, createdTime, modifiedTime, size, webViewLink, parents, ' +
      'permissions(id, type, role, emailAddress, domain, displayName))',
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

  const response = withRetry(function() {
    return Drive.Files.list(params);
  });

  Logger.log('Retrieved ' + (response.files ? response.files.length : 0) + ' files');
  return response;
}

/**
 * Builds the sheet rows for one file (one row per permission) and appends
 * them to auditData.
 */
function appendAuditRows(auditData, file, folderCache) {
  // The file listing already carries the permissions of files the user can
  // share. It leaves them out for shared drive items and files the user
  // can't share, which need a request of their own.
  const permissions = Array.isArray(file.permissions)
    ? file.permissions
    : getFilePermissions(file.id);
  const folderPath = getFolderPath(file.parents, folderCache);

  const fileColumns = [
    file.name,
    folderPath,
    file.id,
    file.owners && file.owners.length > 0 ? file.owners[0].emailAddress : 'Unknown',
    getFileType(file),
    file.mimeType,
    file.createdTime ? new Date(file.createdTime) : '',
    file.modifiedTime ? new Date(file.modifiedTime) : '',
    file.size || '',
    file.webViewLink || ''
  ];

  if (permissions.length === 0) {
    auditData.push(fileColumns.concat([0, '', '', '', '', '']));
    return;
  }

  permissions.forEach(function(permission) {
    auditData.push(fileColumns.concat([
      permissions.length,
      permission.type,
      permission.role,
      permission.emailAddress || '',
      permission.domain || '',
      permission.displayName || ''
    ]));
  });
}

/**
 * Loads the folders the user can see into folderCache with a few large
 * requests, so resolving folder paths doesn't take a request per folder.
 * Whatever it misses (or all of it, if it fails) is looked up one by one
 * in lookupFolder.
 *
 * @param {Object} folderCache - folder id -> { name, parent }
 * @param {string} scope       - 'all' | 'myDrive' | 'owned'
 */
function preloadFolders(folderCache, scope) {
  const deadline = new Date().getTime() + FOLDER_PRELOAD_TIME_MS;
  let pageToken = null;
  let count = 0;

  try {
    do {
      const params = {
        q: "mimeType = 'application/vnd.google-apps.folder'",
        pageSize: 1000,
        fields: 'nextPageToken, files(id, name, parents)'
      };
      if (pageToken) {
        params.pageToken = pageToken;
      }
      if (scope === 'all') {
        params.supportsAllDrives = true;
        params.includeItemsFromAllDrives = true;
      } else {
        // Files the user owns can sit in folders owned by someone else, so
        // 'owned' loads the same folders as 'myDrive'.
        params.corpora = 'user';
      }

      const response = withRetry(function() {
        return Drive.Files.list(params);
      });

      (response.files || []).forEach(function(folder) {
        folderCache[folder.id] = {
          name: folder.name,
          parent: (folder.parents && folder.parents.length > 0) ? folder.parents[0] : null
        };
        count++;
      });

      pageToken = response.nextPageToken;
    } while (pageToken && new Date().getTime() < deadline);
  } catch (error) {
    Logger.log('WARNING: could not preload folders: ' + error.toString());
  }

  Logger.log('Preloaded ' + count + ' folders');
}

/**
 * Looks up a folder's name and parent, fetching it from Drive if it isn't
 * in folderCache yet. Returns null for a folder that can't be read.
 */
function lookupFolder(folderId, folderCache) {
  if (Object.prototype.hasOwnProperty.call(folderCache, folderId)) {
    return folderCache[folderId];
  }

  try {
    const response = withRetry(function() {
      return Drive.Files.get(folderId, {
        fields: 'id, name, parents',
        supportsAllDrives: true
      });
    });
    folderCache[folderId] = {
      name: response.name,
      parent: (response.parents && response.parents.length > 0) ? response.parents[0] : null
    };
  } catch (error) {
    Logger.log('WARNING: could not resolve folder ' + folderId + ': ' + error.toString());
    folderCache[folderId] = null;
  }

  return folderCache[folderId];
}

/**
 * Resolves the full folder path for a file from its parent id(s),
 * e.g. "My Drive/Projects/2026/Reports".
 *
 * Walks up the parent chain, looking each folder up through folderCache so
 * repeated ancestors aren't re-fetched. Returns an empty string for files
 * with no parent (e.g. items shared with you that don't live in a folder
 * you can traverse).
 *
 * @param {string[]} parents     - file.parents from the Drive API
 * @param {Object}   folderCache - folder id -> { name, parent } (or null) cache
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

    const folder = lookupFolder(currentId, folderCache);
    if (!folder) {
      break;
    }

    pathParts.unshift(folder.name);
    currentId = folder.parent;
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
    const response = withRetry(function() {
      return Drive.Permissions.list(fileId, {
        fields: 'permissions(id, type, role, emailAddress, domain, displayName)',
        supportsAllDrives: true
      });
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
function createSummary(sheet, totalFiles, totalPermissions, freeLimitHit) {
  Logger.log('Generating summary data...');

  const summaryData = [
    ['Drive Audit Summary', ''],
    ['', ''],
    ['Audit Date:', new Date()],
    ['Total Files Audited:', freeLimitHit ? totalFiles + ' (free tier limit)' : totalFiles],
    ['Total Permission Entries:', totalPermissions],
    ['', '']
  ];

  // Free tier: results were truncated — show a prominent notice above
  // the next steps. Row positions below shift, so they're tracked.
  let noticeRow = 0;
  if (freeLimitHit) {
    noticeRow = summaryData.length + 1;
    summaryData.push([
      '🔒 FREE TIER LIMIT: only the first ' + FREE_FILE_LIMIT + ' files were audited. ' +
      'Activate your license for only ' + LICENSE_PRICE + ' (one-time payment, lifetime access) ' +
      'in Drive Audit > Activate License Code, then run the audit again for full results.', '']);
    summaryData.push(['', '']);
  }

  const nextStepsRow = summaryData.length + 1;
  summaryData.push(
    ['Next Steps:', ''],
    ['1. Review the "Drive Audit" sheet for detailed permissions', ''],
    ['2. Use filters to find files with specific sharing settings', ''],
    ['3. Look for files shared with "anyone" or external domains', ''],
    ['4. Set up a weekly schedule to run audits automatically', '']
  );

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
  sheet.getRange(nextStepsRow, 1, 1, 1).setFontWeight('bold').setFontSize(12);

  if (noticeRow > 0) {
    sheet.getRange(noticeRow, 1, 1, 2).merge()
      .setFontSize(14)
      .setFontWeight('bold')
      .setBackground('#fce8b2')
      .setFontColor('#b45309')
      .setWrap(true)
      .setVerticalAlignment('middle');
    sheet.setRowHeight(noticeRow, 95);
  }

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
      clearAuditState(scriptProps);

      // Delete the continuation and watchdog triggers
      deleteAuditTriggers();
      Logger.log('Deleted audit triggers');
      
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
    } else if (status === 'PAUSED') {
      icon = '⏸️';
    } else if (status === 'FREE LIMIT HIT') {
      icon = '🔒';
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
    'Free tier shows the first ' + FREE_FILE_LIMIT + ' files per audit. A license unlocks:\n' +
    LICENSE_BENEFITS + '\n\n' +
    'Get a license: use "Activate License Code" in the menu.',
    ui.ButtonSet.OK
  );
}

