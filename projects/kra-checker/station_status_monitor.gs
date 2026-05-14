/**
 * KRA Auto-Checker — Station Status Monitor
 * 
 * Runs every 1 minute via a time-driven trigger.
 * Recalculates NOW() across the Station Status sheet so the
 * formula-based Status column (Online/Stale/Offline) stays current.
 * 
 * SETUP (one time only):
 *   1. Open Rubis Stations Monitoring Allocations spreadsheet
 *   2. Extensions → Apps Script
 *   3. Paste this entire file, click Save
 *   4. Select createTrigger from dropdown, click Run
 *   5. Accept permissions prompt
 *   Done — never needs to be touched again.
 */

function forceRecalculate() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Station Status");
  if (!sheet) return;
  
  // Writing to a helper cell (Z1) forces NOW() to recalculate
  // across the entire sheet including the Status formula column.
  // Column Z is hidden so it never appears to users.
  sheet.getRange("Z1").setValue(new Date());
}

function createTrigger() {
  // Remove any existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  
  // Create a 1-minute time trigger
  ScriptApp.newTrigger("forceRecalculate")
    .timeBased()
    .everyMinutes(1)
    .create();
    
  Logger.log("Trigger created. NOW() recalculates every 1 minute.");
}
