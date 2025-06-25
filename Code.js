/**
 * Northeast Pastors Meeting Agenda Manager
 * Google Apps Script for managing weekly agendas with archive functionality
 */

// Run this when the spreadsheet opens to create the custom menu
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Agenda Manager')
    .addItem('🆕 Create New Agenda Sheet', 'createAgendaSheet')
    .addItem('📁 Archive Current Agenda', 'archiveCurrentAgenda')
    .addItem('🔄 Reset Current Agenda', 'resetCurrentAgenda')
    .addSeparator()
    .addItem('ℹ️ Help', 'showHelp')
    .addToUi();
}

// Creates the initial agenda sheet with formatting
function createAgendaSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if agenda sheet already exists
  let agendaSheet = ss.getSheetByName('Current Agenda');
  if (agendaSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert('Sheet exists', 'Current Agenda sheet already exists. Do you want to recreate it?', ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      return;
    }
    ss.deleteSheet(agendaSheet);
  }
  
  // Create new agenda sheet
  agendaSheet = ss.insertSheet('Current Agenda');
  
  // Set up the sheet structure
  setupAgendaHeaders(agendaSheet);
  setupAgendaTable(agendaSheet);
  setupActionSteps(agendaSheet);
  formatAgendaSheet(agendaSheet);
  
  // Move to front
  ss.setActiveSheet(agendaSheet);
  ss.moveActiveSheet(1);
  
  SpreadsheetApp.getUi().alert('Success!', 'Agenda sheet created successfully!', SpreadsheetApp.getUi().ButtonSet.OK);
}

// Sets up the header section with meeting info and responsibilities
function setupAgendaHeaders(sheet) {
  // Meeting title and date
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue('Northeast Pastors Meeting');
  
  sheet.getRange('A2:F2').merge();
  sheet.getRange('A2').setValue('Date: _____________ | Time: 7:30pm - 8:30pm | Location: _____________');
  
  // Meeting responsibilities header
  sheet.getRange('A4').setValue('Meeting Responsibilities');
  
  // Responsibilities grid
  const responsibilities = [
    ['MC/Facilitator:', '', 'Attendance Taker:', '', 'Note Taker:', ''],
    ['Time Keeper:', '', 'Prayer Leader:', '', 'Follow-up Coordinator:', '']
  ];
  
  sheet.getRange('A5:F6').setValues(responsibilities);
}

// Sets up the main agenda table
/******************************************************************
 * 1) REPLACE THE OLD setupAgendaTable() WITH THIS VERSION
 ******************************************************************/
function setupAgendaTable(sheet) {
  const ui = SpreadsheetApp.getUi();

  // --- Ask the user when the meeting starts --------------------
  const prompt = ui.prompt(
    'Meeting Start Time',
    'Enter meeting start time (e.g. 7:00 PM):',
    ui.ButtonSet.OK_CANCEL
  );
  if (prompt.getSelectedButton() !== ui.Button.OK) return; // user aborted
  const meetingStart = parseTimeString(prompt.getResponseText().trim());
  if (!meetingStart) {
    ui.alert('Sorry, I could not understand that time. Try again (e.g. 7:00 PM).');
    return;
  }

  // --- Table headers (now five columns) ------------------------
  sheet.getRange('A8:E8').setValues([[
    'Agenda Item', 'Start', 'Length', 'Speaker/Lead', 'Notes'
  ]]);

  // --- Raw agenda definitions: [item, length-in-minutes, speaker, notes]
  const rawItems = [
    ['Welcome, Rollcall, Prayer, Trinity Check-in', 15, 'MC',
      '5 min Trinity check-in included'],
    ['Key Developments', 10, 'Naokimi', 'LV Summit updates'],
    ['Sprint 1 Goals by Community', 30, 'Various', ''],
    ['  → Current membership active status', 5, 'TBD', 'Report next week'],
    ['  → Donor level data stats', 5, 'TBD', 'Report next week'],
    ['  → Environment enhancement project', 5, 'TBD', 'Need details – fill in'],
    ['  → Sun Check-in plan & progress', 5, 'NJ', 'NJ to report progress'],
    ['  → 3 Campaign Metrics', 10, 'Team',
      'Share plan & resources, set goals for 21 D'],
    ['True Family Tour in NY/NJ', 10, 'Event Team', ''],
    ['  → June 21 Bank Space Youth Event', 3, 'Event Team', '3 pm event details'],
    ['  → June 22 Evening at Belvedere', 3, 'Event Team', '6 pm event details'],
    ['Northeast Summit Plan & Registration', 5, 'Planning Committee',
      'Registration process update']
  ];

  // --- Build rows with start times --------------------------------------
  const rows = [];
  let runningTime = new Date(meetingStart);           // cursor for each start
  rawItems.forEach(([item, mins, speaker, notes]) => {
    const startStr = Utilities.formatDate(
      runningTime, Session.getScriptTimeZone(), 'h:mm a'
    );
    rows.push([item, startStr, `${mins} min`, speaker, notes || '']);
    runningTime = new Date(runningTime.getTime() + minutesToMillis(mins));
  });

  // --- Drop into sheet ---------------------------------------------------
  sheet.getRange('A9:E' + (8 + rows.length)).setValues(rows);
}

/******************************************************************
 * 2) REPLACE THE OLD formatAgendaSheet() WITH THIS VERSION
 ******************************************************************/
function formatAgendaSheet(sheet) {
  // Column widths for 5-column table
  sheet.setColumnWidth(1, 300); // Agenda Item
  sheet.setColumnWidth(2, 75);  // Start
  sheet.setColumnWidth(3, 65);  // Length
  sheet.setColumnWidth(4, 120); // Speaker
  sheet.setColumnWidth(5, 200); // Notes

  // ---- (everything below here is identical to what you already had,
  //       except any "A8:D…" or "A9:D…" ranges are now "A8:E…" etc.) ----
  // Main title
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#1f4e79').setFontColor('white');
  sheet.getRange('A2').setFontSize(12).setHorizontalAlignment('center')
    .setBackground('#e8f1ff').setFontStyle('italic');

  // Responsibilities header
  sheet.getRange('A4').setFontSize(14).setFontWeight('bold')
    .setBackground('#4a90e2').setFontColor('white');
  sheet.getRange('A5:F6').setBackground('#f8f9fa')
    .getFontStyles().forEach(_ => _); // keep bold names as-is

  // Agenda table header
  sheet.getRange('A8:E8').setFontWeight('bold').setBackground('#4a90e2')
    .setFontColor('white').setHorizontalAlignment('center');

  // Table content – white background
  const lastAgendaRow = sheet.getRange('A8').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  sheet.getRange(`A9:E${lastAgendaRow}`).setBackground('white');

  // Bold main agenda items
  const mains = [9, 10, 11, 17, lastAgendaRow]; // row numbers of bold items
  mains.forEach(r => sheet.getRange(`A${r}`).setFontWeight('bold'));

  // Italic / shaded sub-items
  const subs = [12, 13, 14, 15, 16, 18, 19];
  subs.forEach(r => sheet.getRange(`A${r}`).setFontStyle('italic')
                         .setBackground('#f8f9fa'));

  // Action steps header & area (unchanged)
  // … same as before …

  // Borders
  sheet.getRange('A1:F29').setBorder(true, true, true, true, true, true);
  sheet.getRange('A8:E' + lastAgendaRow).setBorder(true, true, true, true, true, true);
  sheet.getRange('A22:E29').setBorder(true, true, true, true, true, true);

  // Freeze header rows
  sheet.setFrozenRows(8);
}

/******************************************************************
 * 3) ADD THESE TWO SMALL HELPERS ANYWHERE IN YOUR FILE
 ******************************************************************/
function parseTimeString(str) {
  // Accepts “7:00 PM”, “19:00”, “7 pm”, etc.
  const match = str.match(/^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/i);
  if (!match) return null;
  let [ , h, m = '0', meridian ] = match;
  h = parseInt(h, 10);
  m = parseInt(m, 10);
  if (meridian) {
    const pm = meridian.toLowerCase() === 'pm';
    if (pm && h < 12) h += 12;
    if (!pm && h === 12) h = 0;
  }
  const d = new Date();
  d.setHours(h, m, 0, 0);
  return d;
}

const minutesToMillis = mins => mins * 60 * 1000;

// Sets up the action steps section
function setupActionSteps(sheet) {
  // Action steps header
  sheet.getRange('A22').setValue('Action Steps for All Pastors');
  
  // Action items
  const actionItems = [
    ['Immediate Actions (This Week)', '', 'Ongoing Actions', ''],
    ['☐ Start utilizing Sunday service registration form', '', '☐ Promote True Family Tour events', ''],
    ['☐ Complete PSWM Intro Course review', '', '☐ Monitor 3 Campaign Metrics weekly', ''],
    ['☐ Submit community membership data', '', '☐ Submit weekly Sun Checkin reports', ''],
    ['', '', '', ''],
    ['Next Meeting: _____________', '', 'Host: _____________', '']
  ];
  
  sheet.getRange('A24:D29').setValues(actionItems);
}

// Applies formatting to the agenda sheet
function formatAgendaSheet(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 300); // Agenda Item
  sheet.setColumnWidth(2, 80);  // Time
  sheet.setColumnWidth(3, 120); // Speaker
  sheet.setColumnWidth(4, 200); // Notes
  
  // Main title formatting
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#1f4e79').setFontColor('white');
  
  // Subtitle formatting
  sheet.getRange('A2').setFontSize(12).setHorizontalAlignment('center')
    .setBackground('#e8f1ff').setFontStyle('italic');
  
  // Responsibilities section
  sheet.getRange('A4').setFontSize(14).setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
  sheet.getRange('A5:F6').setBackground('#f8f9fa');
  
  // Make responsibility labels bold
  const responsibilityRanges = ['A5', 'C5', 'E5', 'A6', 'C6', 'E6'];
  responsibilityRanges.forEach(range => {
    sheet.getRange(range).setFontWeight('bold');
  });
  
  // Table header formatting
  sheet.getRange('A8:D8').setFontWeight('bold').setBackground('#4a90e2').setFontColor('white')
    .setHorizontalAlignment('center');
  
  // Table content formatting
  sheet.getRange('A9:D20').setBackground('white');
  
  // Make main agenda items bold
  const mainItems = ['A9', 'A10', 'A11', 'A17', 'A20'];
  mainItems.forEach(range => {
    sheet.getRange(range).setFontWeight('bold');
  });
  
  // Sub-items formatting (indented items)
  const subItems = ['A12', 'A13', 'A14', 'A15', 'A16', 'A18', 'A19'];
  subItems.forEach(range => {
    sheet.getRange(range).setFontStyle('italic').setBackground('#f8f9fa');
  });
  
  // Action steps formatting
  sheet.getRange('A22').setFontSize(14).setFontWeight('bold').setBackground('#4caf50').setFontColor('white');
  sheet.getRange('A24:D29').setBackground('#e8f5e8');
  sheet.getRange('A24').setFontWeight('bold');
  sheet.getRange('C24').setFontWeight('bold');
  
  // Add borders to main sections
  sheet.getRange('A1:F29').setBorder(true, true, true, true, true, true);
  sheet.getRange('A8:D20').setBorder(true, true, true, true, true, true);
  sheet.getRange('A22:D29').setBorder(true, true, true, true, true, true);
  
  // Freeze header rows
  sheet.setFrozenRows(8);
}

// Archives the current agenda to a separate sheet
function archiveCurrentAgenda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSheet = ss.getSheetByName('Current Agenda');
  
  if (!currentSheet) {
    SpreadsheetApp.getUi().alert('Error', 'No Current Agenda sheet found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Get current date for archive naming
  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const displayDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  
  // Get or create archive sheet
  let archiveSheet = ss.getSheetByName('Agenda Archive');
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Agenda Archive');
    setupArchiveHeaders(archiveSheet);
  }
  
  // Copy current agenda data to archive
  const agendaData = currentSheet.getRange('A1:F29').getValues();
  const nextRow = archiveSheet.getLastRow() + 2;
  
  // Add separator and date (using ' to prevent formula interpretation)
  archiveSheet.getRange(nextRow, 1, 1, 6).merge();
  archiveSheet.getRange(nextRow, 1).setValue(`MEETING: ${displayDate}`);
  archiveSheet.getRange(nextRow, 1).setFontWeight('bold').setBackground('#ffeb3b').setHorizontalAlignment('center').setFontSize(12);
  
  // Copy the agenda with formatting
  const targetRange = archiveSheet.getRange(nextRow + 1, 1, agendaData.length, agendaData[0].length);
  targetRange.setValues(agendaData);
  
  // Apply some basic formatting to the archived agenda
  const headerRange = archiveSheet.getRange(nextRow + 1, 1, 1, 6);
  headerRange.setFontWeight('bold').setBackground('#1f4e79').setFontColor('white');
  
  const tableHeaderRange = archiveSheet.getRange(nextRow + 8, 1, 1, 4);
  tableHeaderRange.setFontWeight('bold').setBackground('#4a90e2').setFontColor('white');
  
  // Add borders around the archived section
  const archiveRange = archiveSheet.getRange(nextRow, 1, agendaData.length + 1, 6);
  archiveRange.setBorder(true, true, true, true, true, true);
  
  // Confirm with user
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Archive Complete', 
    `Agenda archived successfully for ${displayDate}.\n\nDo you want to reset the current agenda for next week?`, 
    ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    resetCurrentAgenda();
  }
}

// Sets up headers for the archive sheet
function setupArchiveHeaders(sheet) {
  // Set column widths for archive sheet
  sheet.setColumnWidth(1, 300); // Agenda Item
  sheet.setColumnWidth(2, 80);  // Time
  sheet.setColumnWidth(3, 120); // Speaker
  sheet.setColumnWidth(4, 200); // Notes
  sheet.setColumnWidth(5, 100); // Extra space
  sheet.setColumnWidth(6, 100); // Extra space
  
  // Merge and format header
  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue('Northeast Pastors Meeting - Agenda Archive');
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setBackground('#1f4e79').setFontColor('white').setHorizontalAlignment('center');
  
  sheet.getRange('A2:F2').merge();
  sheet.getRange('A2').setValue('Historical record of all meeting agendas');
  sheet.getRange('A2').setFontStyle('italic').setBackground('#e8f1ff').setHorizontalAlignment('center');
}

// Resets the current agenda for a new meeting
function resetCurrentAgenda() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Current Agenda');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', 'No Current Agenda sheet found!', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Clear date and responsibilities
  sheet.getRange('A2').setValue('Date: _____________ | Time: 7:30pm - 8:30pm | Location: _____________');
  
  // Clear responsibility assignments
  const responsibilityRanges = ['B5', 'D5', 'F5', 'B6', 'D6', 'F6'];
  responsibilityRanges.forEach(range => {
    sheet.getRange(range).setValue('');
  });
  
  // Clear notes column (but keep agenda structure)
  sheet.getRange('D9:D20').clearContent();
  
  // Clear action steps checkboxes and next meeting info
  sheet.getRange('A25:A27').setValues([
    ['☐ Start utilizing Sunday service registration form'],
    ['☐ Complete PSWM Intro Course review'],
    ['☐ Submit community membership data']
  ]);
  
  sheet.getRange('C25:C27').setValues([
    ['☐ Promote True Family Tour events'],
    ['☐ Monitor 3 Campaign Metrics weekly'],
    ['☐ Submit weekly Sun Checkin reports']
  ]);
  
  sheet.getRange('A29').setValue('Next Meeting: _____________');
  sheet.getRange('C29').setValue('Host: _____________');
  
  SpreadsheetApp.getUi().alert('Success!', 'Current agenda has been reset for next meeting!', SpreadsheetApp.getUi().ButtonSet.OK);
}

// Shows help information
function showHelp() {
  const helpText = `
📋 AGENDA MANAGER HELP

🆕 Create New Agenda Sheet:
   Creates a formatted agenda sheet with all sections

📁 Archive Current Agenda:
   Saves current agenda to archive and optionally resets for next meeting

🔄 Reset Current Agenda:
   Clears notes and assignments for next meeting (keeps structure)

💡 TIPS:
• Fill in meeting responsibilities at the top before each meeting
• Update agenda items and notes during the meeting
• Archive after each meeting to keep historical records
• Use the reset function to prepare for next week

📧 Questions? Contact your meeting coordinator.
  `;
  
  SpreadsheetApp.getUi().alert('Agenda Manager Help', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

// Additional utility function to update meeting date
function updateMeetingDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Current Agenda');
  
  if (!sheet) return;
  
  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  
  sheet.getRange('A2').setValue(`Date: ${dateString} | Time: 7:30pm - 8:30pm | Location: _____________`);
}
