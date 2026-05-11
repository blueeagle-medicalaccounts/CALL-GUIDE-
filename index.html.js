// ═══════════════════════════════════════════════════════════════════
// BLUE EAGLE MEDICAL ACCOUNTS — COMPLETE SALES TRACKER
// Google Apps Script — Full Version
// Includes: Call Log, Pipeline, Stats + Daily 5PM Email Summary
// ═══════════════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet
// 2. Click Extensions → Apps Script
// 3. Delete existing code, paste this entire file
// 4. Change MANAGER_EMAIL below to your real email address
// 5. Click Save
// 6. Run "initialSetup" once (dropdown → run)
// 7. Deploy → New Deployment → Web App → Anyone → Deploy
// 8. Copy the Web App URL → paste into the Call Guide app (⚙️)
// ═══════════════════════════════════════════════════════════════════

// ── CONFIGURATION — CHANGE THESE ──
const MANAGER_EMAIL = 'info@blueeagle-medacc.com'; // Your email for daily reports
const MANAGER_NAME  = 'Owen';                       // Your name
const COMPANY_NAME  = 'Blue Eagle Medical Accounts';
const SUMMARY_HOUR  = 17; // 5 PM (24hr format) — when daily email sends
const TIMEZONE      = 'Africa/Johannesburg';

// ═══════════════════════════════════════════════════════════════════
// SECTION 1: WEB APP ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);

    ensureSheets(ss);

    logToCallLog(ss, data);

    const o = (data.outcome || '').toUpperCase();
    if (o.includes('HOT') || o.includes('WARM') || o.includes('MEETING')) {
      logToPipeline(ss, data);
    }

    updateStats(ss);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error(err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput('Blue Eagle Sales Tracker is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 2: INITIAL SETUP — Run this once manually
// ═══════════════════════════════════════════════════════════════════

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetName('Blue Eagle — Sales Tracker');

  ensureSheets(ss);
  setupDailyTrigger();

  SpreadsheetApp.getUi().alert(
    '✅ Blue Eagle Sales Tracker is ready!\n\n' +
    '• Call Log sheet created\n' +
    '• Pipeline sheet created\n' +
    '• Stats sheet created\n' +
    '• Daily 5PM email trigger set\n\n' +
    'Now deploy as Web App and paste the URL into your Call Guide.'
  );
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 3: SHEET CREATION & MANAGEMENT
// ═══════════════════════════════════════════════════════════════════

function ensureSheets(ss) {
  if (!ss.getSheetByName('Call Log'))  createCallLogSheet(ss);
  if (!ss.getSheetByName('Pipeline'))  createPipelineSheet(ss);
  if (!ss.getSheetByName('Stats'))     createStatsSheet(ss);
  if (!ss.getSheetByName('Dashboard')) createDashboardSheet(ss);
}

// ── CALL LOG ──
function createCallLogSheet(ss) {
  const s = ss.insertSheet('Call Log');

  const headers = [
    'Date','Time','Salesperson',
    'Practice Name','Doctor','Phone','Email',
    'Specialty','Area',
    'Billing Setup','System',
    'Satisfaction','Outcome','Status',
    'Pain Point','Current Bureau','Their Priority',
    'Tone','Follow-Up Date','Follow-Up Time',
    'Meeting Date','Meeting Time','Meeting Format',
    'Re-Contact Permission','Decline Reason','Notes'
  ];

  s.appendRow(headers);

  const hr = s.getRange(1, 1, 1, headers.length);
  hr.setBackground('#1565c0');
  hr.setFontColor('#ffffff');
  hr.setFontWeight('bold');
  hr.setFontSize(10);
  hr.setHorizontalAlignment('center');
  s.setFrozenRows(1);

  // Column widths
  const widths = {
    1:100, 2:80, 3:100, 4:200, 5:150, 6:120, 7:200,
    8:160, 9:140, 10:120, 11:120, 12:120, 13:180, 14:80,
    15:250, 16:160, 17:160, 18:140, 19:120, 20:100,
    21:120, 22:100, 23:120, 24:160, 25:160, 26:250
  };
  Object.entries(widths).forEach(([col, w]) => s.setColumnWidth(Number(col), w));

  return s;
}

// ── PIPELINE ──
function createPipelineSheet(ss) {
  const s = ss.insertSheet('Pipeline');

  const headers = [
    'Date Added','Salesperson',
    'Practice Name','Doctor','Phone','Email',
    'Specialty','System','Status 🔥',
    'Pain Point','Their Priority',
    'Follow-Up / Meeting Date','Time','Format',
    'Email Sent ✓','Proposal Sent ✓','Client Signed ✓',
    'Revenue (R)','Notes'
  ];

  s.appendRow(headers);

  const hr = s.getRange(1, 1, 1, headers.length);
  hr.setBackground('#0d2040');
  hr.setFontColor('#90caf9');
  hr.setFontWeight('bold');
  hr.setFontSize(10);
  s.setFrozenRows(1);

  const widths = {
    1:100, 2:110, 3:200, 4:150, 5:120, 6:200,
    7:160, 8:120, 9:130, 10:250, 11:160,
    12:140, 13:90, 14:110, 15:100, 16:120, 17:120,
    18:100, 19:250
  };
  Object.entries(widths).forEach(([col, w]) => s.setColumnWidth(Number(col), w));

  return s;
}

// ── STATS ──
function createStatsSheet(ss) {
  const s = ss.insertSheet('Stats');
  s.getRange('A1:F1').merge();
  s.getRange('A1').setValue('BLUE EAGLE MEDICAL ACCOUNTS — LIVE STATS')
    .setFontSize(14).setFontWeight('bold')
    .setFontColor('#ffffff').setBackground('#1565c0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  s.setRowHeight(1, 40);

  const statRows = [
    ['A3','B3','Total Calls (All Time)',       "=COUNTA('Call Log'!A2:A)"],
    ['A4','B4','Calls Today',                  "=COUNTIF('Call Log'!A2:A,TEXT(TODAY(),\"dd/mm/yyyy\"))"],
    ['A5','B5','Calls This Week',              "=SUMPRODUCT(('Call Log'!A2:A>=TEXT(TODAY()-WEEKDAY(TODAY(),2)+1,\"dd/mm/yyyy\"))*('Call Log'!A2:A<>\"\"))"],
    ['A7','B7','HOT Leads (Meeting Booked)',   "=COUNTIF('Call Log'!M2:M,\"HOT*\")+COUNTIF('Call Log'!M2:M,\"*Meeting*\")"],
    ['A8','B8','WARM Leads (Follow-Up)',       "=COUNTIF('Call Log'!M2:M,\"WARM*\")"],
    ['A9','B9','Cold / Not Interested',        "=COUNTIF('Call Log'!M2:M,\"COLD*\")"],
    ['A10','B10','No Answer',                  "=COUNTIF('Call Log'!M2:M,\"NO ANSWER*\")"],
    ['A12','B12','Conversion Rate (Hot/Total)',"=IF(COUNTA('Call Log'!A2:A)=0,\"No calls yet\",TEXT((COUNTIF('Call Log'!M2:M,\"HOT*\")+COUNTIF('Call Log'!M2:M,\"*Meeting*\"))/COUNTA('Call Log'!A2:A),\"0.0%\"))"],
    ['A13','B13','Proposals in Pipeline',      "=COUNTA('Pipeline'!A2:A)"],
    ['A14','B14','Clients Signed',             "=COUNTIF('Pipeline'!Q2:Q,\"YES\")"],
  ];

  statRows.forEach(([labelCell, valCell, label, formula]) => {
    s.getRange(labelCell).setValue(label).setFontWeight('bold').setFontSize(11);
    s.getRange(valCell).setFormula(formula).setFontSize(14).setFontWeight('bold').setFontColor('#1565c0');
  });

  // Section headers
  s.getRange('A6').setValue('LEAD BREAKDOWN').setFontWeight('bold').setFontColor('#888888').setFontSize(9);
  s.getRange('A11').setValue('CONVERSION').setFontWeight('bold').setFontColor('#888888').setFontSize(9);

  s.setColumnWidth(1, 260);
  s.setColumnWidth(2, 120);

  return s;
}

// ── DASHBOARD ──
function createDashboardSheet(ss) {
  const s = ss.insertSheet('Dashboard');
  s.getRange('A1:G1').merge();
  s.getRange('A1').setValue('📊 BLUE EAGLE — SALES DASHBOARD')
    .setFontSize(16).setFontWeight('bold')
    .setFontColor('#ffffff').setBackground('#1565c0')
    .setHorizontalAlignment('center');
  s.setRowHeight(1, 44);

  s.getRange('A2').setValue('Auto-calculated from Call Log. Last updated: ' + new Date().toLocaleString())
    .setFontColor('#888888').setFontStyle('italic').setFontSize(9);

  // Instructions
  const instructions = [
    ['A4', 'HOW TO USE THIS TRACKER'],
    ['A5', '1. Your salesperson uses the Blue Eagle Call Guide app on their phone'],
    ['A6', '2. Each call they log automatically appears in the Call Log tab'],
    ['A7', '3. HOT and WARM leads automatically appear in the Pipeline tab'],
    ['A8', '4. You receive a daily summary email at 5PM every weekday'],
    ['A9', '5. Manually update Pipeline: tick Email Sent, Proposal Sent, Client Signed'],
    ['A10','6. Stats tab shows live totals — refresh anytime'],
  ];

  instructions.forEach(([cell, text]) => {
    const r = s.getRange(cell);
    if (cell === 'A4') {
      r.setValue(text).setFontWeight('bold').setFontSize(12).setFontColor('#1565c0');
    } else {
      r.setValue(text).setFontSize(11).setFontColor('#333333');
    }
  });

  s.setColumnWidth(1, 600);

  return s;
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 4: DATA LOGGING
// ═══════════════════════════════════════════════════════════════════

function logToCallLog(ss, data) {
  const s = ss.getSheetByName('Call Log');
  if (!s) return;

  const o = (data.outcome || '').toUpperCase();
  let status = 'Cold';
  if (o.includes('HOT') || o.includes('MEETING')) status = 'HOT 🔥';
  else if (o.includes('WARM'))                    status = 'WARM 🌡';
  else if (o.includes('NO ANSWER'))               status = 'No Answer 📵';
  else if (o.includes('COLD'))                    status = 'Cold ❄️';

  const row = [
    data.date        || Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy'),
    data.time        || Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm'),
    data.salesperson || '',
    data.practice    || '',
    data.doctor      || '',
    data.phone       || '',
    data.email_confirmed || data.email || '',
    data.practice_type   || '',
    data.area            || '',
    data.billing         || '',
    data.system          || '',
    data.satisfaction    || '',
    data.outcome         || '',
    status,
    data.pain            || '',
    data.bureau          || '',
    data.interest        || '',
    data.tone            || '',
    data.followup_date   || '',
    data.followup_time   || '',
    data.meeting_date    || '',
    data.meeting_time    || '',
    data.meeting_format  || '',
    data.recontact       || '',
    data.decline_reason  || '',
    data.notes           || '',
  ];

  s.appendRow(row);

  // Colour the status cell
  const lastRow = s.getLastRow();
  const statusCell = s.getRange(lastRow, 14);
  if (status.includes('HOT'))      statusCell.setBackground('#fce8e6').setFontColor('#c0392b');
  else if (status.includes('WARM'))statusCell.setBackground('#fff3cd').setFontColor('#d68910');
  else if (status.includes('Cold'))statusCell.setBackground('#e8f0fe').setFontColor('#1565c0');
  else if (status.includes('No'))  statusCell.setBackground('#f5f5f5').setFontColor('#888888');
}

function logToPipeline(ss, data) {
  const s = ss.getSheetByName('Pipeline');
  if (!s) return;

  const o = (data.outcome || '').toUpperCase();
  let status = 'WARM 🌡';
  if (o.includes('HOT') || o.includes('MEETING')) status = 'HOT 🔥';

  s.appendRow([
    data.date          || Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy'),
    data.salesperson   || '',
    data.practice      || '',
    data.doctor        || '',
    data.phone         || '',
    data.email_confirmed || data.email || '',
    data.practice_type || '',
    data.system        || '',
    status,
    data.pain          || '',
    data.interest      || '',
    data.followup_date || data.meeting_date || '',
    data.followup_time || data.meeting_time || '',
    data.meeting_format|| '',
    '', '', '', // Email sent, Proposal sent, Client signed (manual)
    '',          // Revenue
    data.notes   || '',
  ]);

  // Colour status
  const lastRow = s.getLastRow();
  const sc = s.getRange(lastRow, 9);
  if (status.includes('HOT'))  sc.setBackground('#fce8e6').setFontColor('#c0392b').setFontWeight('bold');
  else                         sc.setBackground('#fff3cd').setFontColor('#d68910').setFontWeight('bold');
}

function updateStats(ss) {
  const s = ss.getSheetByName('Stats');
  if (!s) return;
  // Stats use live formulas — just update the timestamp on dashboard
  const d = ss.getSheetByName('Dashboard');
  if (d) {
    d.getRange('A2').setValue('Auto-calculated from Call Log. Last updated: ' +
      Utilities.formatDate(new Date(), TIMEZONE, 'dd MMM yyyy HH:mm'));
  }
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 5: DAILY 5PM EMAIL SUMMARY
// ═══════════════════════════════════════════════════════════════════

function setupDailyTrigger() {
  // Remove any existing daily triggers first
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailySummaryEmail') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new daily trigger at 5PM SA time
  ScriptApp.newTrigger('sendDailySummaryEmail')
    .timeBased()
    .everyDays(1)
    .atHour(SUMMARY_HOUR)
    .create();

  console.log('Daily 5PM trigger set successfully.');
}

function sendDailySummaryEmail() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const callLog = ss.getSheetByName('Call Log');
  const pipeline= ss.getSheetByName('Pipeline');

  if (!callLog) return;

  const today     = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy');
  const todayDisp = Utilities.formatDate(new Date(), TIMEZONE, 'dd MMMM yyyy');
  const dayName   = Utilities.formatDate(new Date(), TIMEZONE, 'EEEE');

  // ── GET TODAY'S CALLS ──
  const allData    = callLog.getDataRange().getValues();
  const headers    = allData[0];
  const todayCalls = allData.slice(1).filter(row => row[0] === today);

  if (todayCalls.length === 0) {
    // Still send but note no calls
    sendNoCallsEmail(todayDisp, dayName);
    return;
  }

  // ── COUNT OUTCOMES ──
  let hot = 0, warm = 0, cold = 0, noAnswer = 0;
  const hotLeads  = [];
  const warmLeads = [];

  todayCalls.forEach(row => {
    const outcome = (row[12] || '').toUpperCase();
    const status  = (row[13] || '').toUpperCase();
    const practice= row[3] || '';
    const doctor  = row[4] || '';
    const phone   = row[5] || '';
    const pain    = row[14]|| '';
    const followup= row[18]|| '';
    const meetDate= row[20]|| '';

    if (status.includes('HOT') || outcome.includes('MEETING')) {
      hot++;
      hotLeads.push({ practice, doctor, phone, pain, meetDate: meetDate || followup });
    } else if (status.includes('WARM')) {
      warm++;
      warmLeads.push({ practice, doctor, phone, pain, followup });
    } else if (status.includes('NO ANSWER') || outcome.includes('NO ANSWER')) {
      noAnswer++;
    } else {
      cold++;
    }
  });

  const total = todayCalls.length;
  const salesperson = todayCalls[0][2] || 'Your salesperson';

  // ── BUILD EMAIL HTML ──
  const html = buildEmailHTML({
    todayDisp, dayName, salesperson,
    total, hot, warm, cold, noAnswer,
    hotLeads, warmLeads
  });

  // ── SEND EMAIL ──
  GmailApp.sendEmail(
    MANAGER_EMAIL,
    `🦅 Blue Eagle Daily Sales Report — ${todayDisp}`,
    stripHtml(html),
    {
      htmlBody: html,
      name: `${COMPANY_NAME} Sales Tracker`
    }
  );

  console.log(`Daily summary sent for ${today}: ${total} calls, ${hot} HOT, ${warm} WARM`);
}

function sendNoCallsEmail(todayDisp, dayName) {
  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
      ${emailHeader()}
      <div style="padding:24px;background:#f8f9fa;">
        <p style="color:#555;font-size:15px;">Hi ${MANAGER_NAME},</p>
        <p style="color:#555;font-size:15px;">No calls were logged today (${todayDisp}).</p>
        <p style="color:#555;font-size:14px;">If calls were made, check that the Call Guide app is connected to the sheet correctly.</p>
      </div>
      ${emailFooter()}
    </div>`;

  GmailApp.sendEmail(
    MANAGER_EMAIL,
    `🦅 Blue Eagle Daily Report — ${todayDisp} (No calls logged)`,
    `No calls logged today (${todayDisp}).`,
    { htmlBody: html, name: `${COMPANY_NAME} Sales Tracker` }
  );
}

function buildEmailHTML(d) {
  const convRate = d.total > 0 ? Math.round((d.hot / d.total) * 100) : 0;

  // HOT leads rows
  let hotRows = '';
  if (d.hotLeads.length > 0) {
    hotRows = `
      <h3 style="color:#c0392b;font-size:14px;margin:24px 0 10px;">🔥 HOT LEADS — Meeting Booked (${d.hotLeads.length})</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr style="background:#fce8e6;">
          <th style="padding:8px 10px;text-align:left;color:#c0392b;border:1px solid #f5b7b1;">Practice</th>
          <th style="padding:8px 10px;text-align:left;color:#c0392b;border:1px solid #f5b7b1;">Doctor</th>
          <th style="padding:8px 10px;text-align:left;color:#c0392b;border:1px solid #f5b7b1;">Phone</th>
          <th style="padding:8px 10px;text-align:left;color:#c0392b;border:1px solid #f5b7b1;">Meeting</th>
        </tr>
        ${d.hotLeads.map(l => `
        <tr style="background:#fff;">
          <td style="padding:8px 10px;border:1px solid #f5b7b1;font-weight:bold;">${l.practice}</td>
          <td style="padding:8px 10px;border:1px solid #f5b7b1;">${l.doctor}</td>
          <td style="padding:8px 10px;border:1px solid #f5b7b1;">${l.phone}</td>
          <td style="padding:8px 10px;border:1px solid #f5b7b1;color:#c0392b;font-weight:bold;">${l.meetDate || 'TBC'}</td>
        </tr>
        <tr style="background:#fef9f9;">
          <td colspan="4" style="padding:6px 10px;border:1px solid #f5b7b1;color:#888;font-size:12px;font-style:italic;">Pain: ${l.pain || 'Not specified'}</td>
        </tr>`).join('')}
      </table>`;
  }

  // WARM leads rows
  let warmRows = '';
  if (d.warmLeads.length > 0) {
    warmRows = `
      <h3 style="color:#d68910;font-size:14px;margin:24px 0 10px;">🌡 WARM LEADS — Email Sent, Follow-Up Pending (${d.warmLeads.length})</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <tr style="background:#fff3cd;">
          <th style="padding:8px 10px;text-align:left;color:#d68910;border:1px solid #fdebd0;">Practice</th>
          <th style="padding:8px 10px;text-align:left;color:#d68910;border:1px solid #fdebd0;">Doctor</th>
          <th style="padding:8px 10px;text-align:left;color:#d68910;border:1px solid #fdebd0;">Phone</th>
          <th style="padding:8px 10px;text-align:left;color:#d68910;border:1px solid #fdebd0;">Follow-Up</th>
        </tr>
        ${d.warmLeads.map(l => `
        <tr style="background:#fff;">
          <td style="padding:8px 10px;border:1px solid #fdebd0;font-weight:bold;">${l.practice}</td>
          <td style="padding:8px 10px;border:1px solid #fdebd0;">${l.doctor}</td>
          <td style="padding:8px 10px;border:1px solid #fdebd0;">${l.phone}</td>
          <td style="padding:8px 10px;border:1px solid #fdebd0;color:#d68910;font-weight:bold;">${l.followup || 'TBC'}</td>
        </tr>
        <tr style="background:#fffef7;">
          <td colspan="4" style="padding:6px 10px;border:1px solid #fdebd0;color:#888;font-size:12px;font-style:italic;">Pain: ${l.pain || 'Not specified'}</td>
        </tr>`).join('')}
      </table>`;
  }

  return `
<!DOCTYPE html>
<html>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,sans-serif;">
<div style="max-width:620px;margin:0 auto;background:#ffffff;">

  ${emailHeader()}

  <!-- GREETING -->
  <div style="padding:24px 28px 0;">
    <p style="color:#333;font-size:15px;margin:0 0 4px;">Hi ${MANAGER_NAME},</p>
    <p style="color:#555;font-size:14px;margin:0 0 20px;">Here's your sales summary for <strong>${d.todayDisp}</strong> (${d.dayName}).</p>
  </div>

  <!-- STATS ROW -->
  <div style="padding:0 28px;">
    <table style="width:100%;border-collapse:separate;border-spacing:8px;">
      <tr>
        <td style="background:#1565c0;border-radius:10px;padding:16px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#fff;">${d.total}</div>
          <div style="font-size:11px;color:#90caf9;text-transform:uppercase;letter-spacing:1px;margin-top:4px;">Total Calls</div>
        </td>
        <td style="background:#c0392b;border-radius:10px;padding:16px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#fff;">${d.hot}</div>
          <div style="font-size:11px;color:#f1948a;text-transform:uppercase;letter-spacing:1px;margin-top:4px;">Hot 🔥</div>
        </td>
        <td style="background:#d68910;border-radius:10px;padding:16px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#fff;">${d.warm}</div>
          <div style="font-size:11px;color:#fdebd0;text-transform:uppercase;letter-spacing:1px;margin-top:4px;">Warm 🌡</div>
        </td>
        <td style="background:#566573;border-radius:10px;padding:16px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#fff;">${d.cold + d.noAnswer}</div>
          <div style="font-size:11px;color:#bfc9ca;text-transform:uppercase;letter-spacing:1px;margin-top:4px;">Cold/No Ans</div>
        </td>
        <td style="background:#1a5276;border-radius:10px;padding:16px;text-align:center;">
          <div style="font-size:32px;font-weight:bold;color:#fff;">${convRate}%</div>
          <div style="font-size:11px;color:#7fb3d3;text-transform:uppercase;letter-spacing:1px;margin-top:4px;">Hot Rate</div>
        </td>
      </tr>
    </table>
  </div>

  <!-- SALESPERSON -->
  <div style="padding:16px 28px 0;">
    <p style="color:#888;font-size:12px;margin:0;">
      Logged by: <strong style="color:#1565c0;">${d.salesperson}</strong>
    </p>
  </div>

  <!-- HOT LEADS -->
  <div style="padding:8px 28px 0;">
    ${hotRows}
  </div>

  <!-- WARM LEADS -->
  <div style="padding:8px 28px 0;">
    ${warmRows}
  </div>

  <!-- ACTION ITEMS -->
  <div style="padding:24px 28px 0;">
    <h3 style="color:#1565c0;font-size:14px;margin:0 0 12px;">✅ Action Items for Tomorrow</h3>
    <div style="background:#e8f0fe;border-left:4px solid #1565c0;padding:14px 16px;border-radius:0 8px 8px 0;font-size:13px;color:#333;line-height:1.7;">
      ${d.hot > 0 ? `• Send meeting confirmations for all <strong>${d.hot} HOT lead(s)</strong> today<br>` : ''}
      ${d.warm > 0 ? `• Follow up on all <strong>${d.warm} WARM lead(s)</strong> on their agreed dates<br>` : ''}
      ${d.cold > 0 ? `• <strong>${d.cold} cold lead(s)</strong> logged — mark for 60-day re-engagement<br>` : ''}
      ${d.noAnswer > 0 ? `• <strong>${d.noAnswer} no-answer(s)</strong> — retry at different time tomorrow<br>` : ''}
      • Review Pipeline tab and update Email Sent / Proposal Sent columns
    </div>
  </div>

  <!-- SHEET LINK -->
  <div style="padding:20px 28px;">
    <a href="https://docs.google.com/spreadsheets/d/1-23jryh9f1TgXTnvlWGk77R5AGmcxHLed9xrcwDvPcE"
       style="display:inline-block;background:#1565c0;color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:14px;">
      📊 Open Full Sales Tracker →
    </a>
  </div>

  ${emailFooter()}
</div>
</body>
</html>`;
}

function emailHeader() {
  return `
  <div style="background:linear-gradient(135deg,#0d2040,#1565c0);padding:24px 28px;text-align:center;">
    <img src="https://lirp.cdn-website.com/07efeb14/dms3rep/multi/opt/Blue+Eagle+logo-1920w.JPG"
         style="height:60px;background:#fff;padding:6px;border-radius:8px;margin-bottom:12px;" alt="Blue Eagle">
    <div style="color:#fff;font-size:18px;font-weight:bold;letter-spacing:0.5px;">Blue Eagle Medical Accounts</div>
    <div style="color:#90caf9;font-size:12px;letter-spacing:2px;text-transform:uppercase;margin-top:4px;">Daily Sales Report</div>
  </div>`;
}

function emailFooter() {
  return `
  <div style="background:#0d2040;padding:16px 28px;text-align:center;">
    <p style="color:#566573;font-size:11px;margin:0;">
      ${COMPANY_NAME} · Pretoria, South Africa<br>
      010 110 9889 · info@blueeagle-medacc.com · blueeaglemedicalaccounts.com<br><br>
      This report is generated automatically from your Blue Eagle Sales Tracker.
    </p>
  </div>`;
}

function stripHtml(html) {
  return html.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim();
}

// ═══════════════════════════════════════════════════════════════════
// SECTION 6: MANUAL TRIGGERS (run from Apps Script editor)
// ═══════════════════════════════════════════════════════════════════

// Run this to send a test email immediately
function sendTestEmail() {
  sendDailySummaryEmail();
}

// Run this to reset the daily trigger
function resetDailyTrigger() {
  setupDailyTrigger();
}

// Run this to manually add a test call (useful for testing)
function addTestCall() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheets(ss);

  const testData = {
    date: Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy'),
    time: Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm'),
    salesperson: 'Owen',
    practice: 'Test Practice — Sandton',
    doctor: 'Dr Test',
    phone: '011 000 0000',
    email: 'test@practice.co.za',
    practice_type: 'General Practitioner (GP)',
    area: 'Gauteng — Johannesburg',
    billing: 'Outsourced',
    system: 'GoodX',
    satisfaction: 'Some Issues',
    outcome: 'HOT — Meeting Booked',
    pain: 'Rejected claims not being followed up, poor communication',
    bureau: 'XYZ Billing',
    interest: 'Rejected claims recovery',
    tone: 'Very receptive and engaged',
    meeting_date: Utilities.formatDate(new Date(Date.now() + 86400000), TIMEZONE, 'dd/MM/yyyy'),
    meeting_time: 'Morning (before 12pm)',
    meeting_format: 'Online / Video Call',
    notes: 'Test entry — delete after confirming setup works'
  };

  logToCallLog(ss, testData);
  logToPipeline(ss, testData);

  SpreadsheetApp.getUi().alert('✅ Test call added to Call Log and Pipeline. Check both tabs.');
}
