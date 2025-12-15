/**
 * Creator: Galbatorix
 * Script for ONEderland Leave Request System
 * 
 * Contact the creator for support, feature requests, or issues.
 *
 * Version: 3.0
 * Date: November 24, 2025
 */

const MAINTENANCE_MODE = false; // Toggle this to false to bring site back online

// // Pop-Up Welcome Message
// function onOpen() {
//   SpreadsheetApp.getUi().alert(
//     "⚠️ Main Approval - Important Reminder",
//     "Please edit this sheet carefully.\n\n" +
//     "- Do NOT delete any tables, even if they appear empty.\n" +
//     "- If something is not needed, simply HIDE the sheet, table, or rows instead.\n" +
//     "- You may use colors, filters, or notes as needed.\n" +
//     "- Do NOT remove formulas or structures prepared for future use.\n\n" +
//     "Thank you for maintaining data consistency.\n\n" +
//     "With Love, \n- SnG IT Team",
//     SpreadsheetApp.getUi().ButtonSet.OK
//   );
// }

// ------------------------ Start of This is for Testing Only ------------------------
// const CONFIG = {
//   REPORTING_EMAILS: ["iyyan.onederland@gmail.com"],
//   GM_EMAIL: "iyyan.onederland@gmail.com",
//   HR_EMAIL: "iyyan.onederland@gmail.com",
//   SPV_MAP: {
//     "Carbon Energy": "iyyan.onederland@gmail.com",
//     "Education ONE": "iyyan.onederland@gmail.com",
//     "English Cafe": "iyyan.onederland@gmail.com",
//     "General Manager": "iyyan.onederland@gmail.com",
//     "Neurone Recruitment": "iyyan.onederland@gmail.com",
//     "ONEderland Consulting": "iyyan.onederland@gmail.com",
//     "ONEderland Enterprise Finance": "iyyan.onederland@gmail.com",
//     "ONEderland Enterprise HRGA": "iyyan.onederland@gmail.com",
//     "PeraONE Xperience": "iyyan.onederland@gmail.com",
//     "SnG OE": "iyyan.onederland@gmail.com"
//   }
// };
// ------------------------ End of This is for Testing Only ------------------------

// Global Emails Configuration
const CONFIG = {
  REPORTING_EMAILS: ["komang.onederland@gmail.com", "bcc.hrdteam@gmail.com"],
  GM_EMAIL: "ayu.septyani@educationone.net.au",
  HR_EMAIL: "dyah.onederland@gmail.com",
  SPV_MAP: {
    "Carbon Energy": "ika.widia@carbonenergy.net.au",
    "Education ONE": "cesco.neurone@gmail.com",                 
    "English Cafe": "harris.englishcafe@gmail.com",
    "General Manager": "ayu.septyani@educationone.net.au",
    "Neurone Recruitment": "cesco.neurone@gmail.com",
    "ONEderland Consulting": "ayu.karina@onederland.com.au",
    "ONEderland Enterprise Finance": "sanistya.onederland@gmail.com",
    "ONEderland Enterprise HRGA": "dyah.onederland@gmail.com",
    "PeraONE Xperience": "dyah.onederland@gmail.com",
    "SnG OE": "suma.onederland@gmail.com"
  }
};

// Column indices
const COLUMNS = {
  TIMESTAMP: 1,       // Column A
  NAME: 2,            // Column B
  DEPARTMENT: 3,      // Column C
  LEAVE_TYPE: 4,      // Column D
  START_DATE: 5,      // Column E
  END_DATE: 6,        // Column F
  REASON: 7,          // Column G
  STATUS: 8,          // Column H
  REQUESTER_EMAIL: 9, // Column I
  SPV_EMAIL: 10,      // Column J
  HR_EMAIL: 11,       // Column K
  GM_EMAIL: 12,       // Column L
  STAGE: 13,          // Column M
  DECISION: 14,       // Column N
  NOTE: 15,           // Column O
  DECISION_DATE: 16,  // Column P
  SPV_DECISION: 17,   // Column Q
  HR_DECISION: 18,    // Column R
  GM_DECISION: 19,    // Column S
  SPV_TOKEN: 20,      // Column T
  HR_TOKEN: 21,       // Column U
  GM_TOKEN: 22,       // Column V
  EMP_EMAIL: 23,      // Column W
  ANNUAL_BALANCE: 24, // Column X
  SICK_BALANCE: 25,    // Column Y
  // BEREA_BALANCE: 26     // Column Z
  // MARRIAGE_BALANCE: 27  // Column AA
  // MATERNITY_BALANCE: 28 // Column AB
  REF_ID: 43
};

/**
 * Helper function to parse dates from "d-m-Y" format.
 * @param {string} dateString The date string in "d-m-Y".
 * @return {Date} A JavaScript Date object.
 */

function getCurrentUser() {
  const cache = CacheService.getUserCache();
  const props = PropertiesService.getUserProperties();

  return {
    name: cache.get("userName") || props.getProperty("userName") || "Unknown User",
    email: cache.get("userEmail") || props.getProperty("userEmail") || "unknown@domain.com"
  };
}

function normalizeGmailAddress(email) {
  if (!email) return '';
  email = email.trim().toLowerCase();

  if (email.endsWith('@gmail.com')) {
    const [local, domain] = email.split('@');
    const normalizedLocal = local.split('+')[0].replace(/\./g, '');
    return `${normalizedLocal}@${domain}`;
  }
  return email;
}

function generateReferenceID() {
  const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
  let randomLetters = "";
  for (let i = 0; i < 7; i++) {
    randomLetters += letters.charAt(Math.floor(Math.random() * letters.length));
  }
  return `ONE-${randomLetters}/`;
}

function parseDMYDate(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split("-");
  if (parts.length === 3) {
    // new Date(year, monthIndex, day)
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  return null;
}

// Helper function to format date objects to string
function formatDate(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    // If it's already a string (potentially from sheet), return it, or handle error
    return (typeof dateObj === 'string') ? dateObj : "Invalid Date";
  }
  try {
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  } catch (e) {
    Logger.log("formatDate error: " + e.toString());
    // Fallback for environments where Session might not be available or date is weird
    const d = dateObj.getDate();
    const m = dateObj.getMonth() + 1;
    const y = dateObj.getFullYear();
    return (d < 10 ? '0' : '') + d + '-' + (m < 10 ? '0' : '') + m + '-' + y;
  }
}

function generateRandomToken(length = 16) {
  const chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890~<>!@#$%^&*';
  let token = '';
  for (let i = 0; i < length; i++) {
    token += chars[Math.floor(Math.random() * chars.length)];
  }
  return token;
}

// Get email of current Google user
const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();

function showErrorTokenPage(title, message) {
  const html = `
    <html>
      <head>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
      </head>
      <body class="bg-light d-flex align-items-center justify-content-center vh-100">
        <div class="card shadow-lg" style="max-width: 500px;">
          <div class="card-body text-center">
            <h3 class="text-danger mb-3">⚠️ ${title}</h3>
            <p class="text-muted">${message}</p>
            <hr>
            <a href="${ScriptApp.getService().getUrl()}" class="btn btn-primary mt-3">Return to Form</a>
          </div>
        </div>
      </body>
    </html>
  `;
  Logger.log(`Error Page Rendered for ${currentUserEmail} - ${title}`);
  return HtmlService.createHtmlOutput(html).setTitle(title);
}

function doGet(e) {

  if (MAINTENANCE_MODE) {
    Logger.log("Maintenance mode active — showing maintenance page.");
    return showMaintenancePage();
  }

  const page = e?.parameter?.page;
  const action = e.parameter.action;
  const row = parseInt(e.parameter.row, 10);
  const stage = e.parameter.stage;

  if (action === 'review' && row && stage) {
    Logger.log(`Action: review`);
    Logger.log(`Row: ${row}`);
    Logger.log(`Stage: ${stage}`);

    const template = HtmlService.createTemplateFromFile('rejectWithNotes');
    template.row = row;
    template.stage = stage;

    Logger.log(`Rendering rejectWithNotes template for row ${row} at stage ${stage}`);
    return template.evaluate().setTitle('Reject with Notes');
  }

  if (page === 'privacy') {
    const activeUser = Session.getActiveUser().getEmail();
    Logger.log(`[INFO]: Active user ${activeUser} read Privacy Policy Page`);
    return HtmlService.createHtmlOutputFromFile('privacy').setTitle('Privacy Policy');
  }

  if (page === 'terms') {
    const activeUser = Session.getActiveUser().getEmail();
    Logger.log(`[INFO]: Active user ${activeUser} read Terms of Usage Page`);
    return HtmlService.createHtmlOutputFromFile('terms').setTitle('Terms of Service');
  }

  const email = e.parameter.track;
  const showHistory = e.parameter.history === "true";

  if (email) {
    const activeUser = Session.getActiveUser().getEmail();
    if (activeUser && activeUser.toLowerCase() !== email.toLowerCase()) {
      const html = `
        <div class="container mt-5">
          <div class="alert alert-danger shadow p-4 rounded" role="alert">
            <div class="col-md-8">
              <h4 class="alert-heading">⛔ Access Denied</h4>
                <p>You are not authorized to view this leave request history.</p>
                <hr>
                <p class="mb-0">
                Please make sure you're signed in with the correct Google account.<br><br>
                If you need to change accounts, click the link above and sign in with the correct email.<br>
                <a href="https://accounts.google.com/AccountChooser" target="_blank" rel="noopener noreferrer" class="btn btn-link p-0 text-decoration-none">Switch Google Account</a>
                </p>
            </div>
          </div>
        </div>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      `;
      Logger.log(`Access Denied: Active user ${activeUser} tried accessing ${email}`);
      return HtmlService.createHtmlOutput(html).setTitle("Unauthorized Access");
    }

    Logger.log(`Access Granted: Showing history for ${activeUser}`);
    return HtmlService.createHtmlOutput(renderTrackingPage(email, showHistory)).setTitle("Track My Leave Request");
  }

  try {
    const activeUser = Session.getActiveUser().getEmail();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');

    if (!e || !e.parameter.action || !e.parameter.row) {

      const user = getCurrentUser(); // ← now safe

      const template = HtmlService.createTemplateFromFile('form');
      template.userEmail = user.email;
      template.userName = user.name;

      // Logger.log(`Showing form for: ${user.email}`);
      return template.evaluate()
        .setTitle("Leave Request Form")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    }

    const rowIndex = parseInt(e.parameter.row, 10);
    const action = e.parameter.action;
    // Note for rejection: If a note is crucial for rejection,
    // the email link should prompt the user or they should be instructed to add it to the URL.
    // e.g., &note=YourReasonHere. Or an intermediary HTML page could be used for a richer experience.
    const note = e.parameter.note || ''; // Approvers can manually add ?note=text to the URL if needed.
    // const stageFromParam = e.parameter.stage; // The stage link was clicked from

    if (isNaN(rowIndex) || rowIndex < 2 || !['approve', 'reject'].includes(action)) {
      Logger.log(`[INFO]: ${activeUser} - Invalid parameters. Please ensure the link is correct.`);
      return HtmlService.createHtmlOutput("Error: Invalid parameters. Please ensure the link is correct.").setTitle("Error");
    }

    // Lock to prevent concurrent modifications to the same row if possible (simple lock)
    const lock = LockService.getScriptLock();
    lock.waitLock(15000); // Wait up to 15 seconds for lock

    let nextStage = 'Final';
    let decision = action === 'approve' ? 'Approved' : 'Rejected';
    let currentStage = '';
    let name = '';
    let department = '';
    let leaveType = '';
    let requester = '';
    let startDate, endDate, reasonText;

    try {
      currentStage = sheet.getRange(rowIndex, COLUMNS.STAGE).getValue();
      const currentStatus = sheet.getRange(rowIndex, COLUMNS.STATUS).getValue();

      if (currentStatus !== 'Pending' && currentStatus !== '') {
        Logger.log(`Request at row ${rowIndex} has status '${currentStatus}' and is currently at stage '${currentStage}'`);

        // If status is Approved or Rejected, consider it final and stop further processing
        if (currentStatus === 'Approved' || currentStatus === 'Rejected') {
          const html = HtmlService.createTemplateFromFile('result');
          html.action = currentStatus.toLowerCase();
          html.stage = currentStage;

          // Try/catch in case note is blank or access fails
          try {
            html.note = sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
          } catch (error) {
            Logger.log(`Failed to fetch note for row ${rowIndex}: ${error}`);
            html.note = '';
          }

          html.nextStage = "Final";

          Logger.log(`Returning result page with status '${currentStatus}' for row ${rowIndex}`);
          return html.evaluate().setTitle("Request Processed");
        }
      }

      // Get the stageToken map
      const stageTokenColumnMap = {
        "SPV Approval": COLUMNS.SPV_TOKEN,
        "HR Review": COLUMNS.HR_TOKEN,
        "GM Review": COLUMNS.GM_TOKEN
      };

      // Get the `data` row and extract currentStage BEFORE using it
      const data = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Now it's safe to use currentStage to look up token column
      const tokenColumn = stageTokenColumnMap[currentStage];

      // Get token from link
      const tokenToUse = e.parameter.token;
      const stageParam = e.parameter.stage; // spv/hr/gm

      // Validate inputs
      // This is a critical check to ensure the token and stage are provided
      // Also this happen when the sheet 'share link' is changed to visitor. The Correct one is 'Editor'.
      if (!tokenToUse || !tokenColumn || !stageParam) {
        Logger.log(`[TOKEN CHECK] Unauthorized Access - Missing token/column/stage: tokenToUse=${tokenToUse}, tokenColumn=${tokenColumn}, stageParam=${stageParam}`);
        return showErrorTokenPage("Unauthorized Access", "Hey Stranger, what are you doing here?");
      }

      // Normalize stage and validate token
      const validStage = {
        spv: COLUMNS.SPV_TOKEN,
        hr: COLUMNS.HR_TOKEN,
        gm: COLUMNS.GM_TOKEN
      }[stageParam.toLowerCase()];

      // // Get email of current Google user
      // const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
      // const spvEmail = (data[COLUMNS.SPV_EMAIL - 1] || "").toLowerCase();
      // const hrEmail = (data[COLUMNS.HR_EMAIL - 1] || "").toLowerCase();
      // const gmEmail = (data[COLUMNS.GM_EMAIL - 1] || "").toLowerCase();

      // let expectedApprover = "";

      // if (stageParam === "spv") expectedApprover = spvEmail;
      // else if (stageParam === "hr") expectedApprover = hrEmail;
      // else if (stageParam === "gm") expectedApprover = gmEmail;
      // else return showErrorTokenPage("Invalid Stage", "Unknown stage: " + stageParam);

      // if (currentUserEmail !== expectedApprover) {
      //   Logger.log(`[TOKEN CHECK] Unauthorized Approver - Expected: ${expectedApprover}, Got: ${currentUserEmail}`);
      //   return showErrorTokenPage("Unauthorized Approver", `Whoopzz Whoopzz <code>${currentUserEmail}</code>, <br>You're not authorized to approve this request<br>for stage <b>${currentStage}</b>.`);
      // }

      // // Check saved token
      // const savedToken = sheet.getRange(rowIndex, tokenColumn).getValue();
      // Logger.log(`[TOKEN CHECK] Comparing tokens - savedToken: ${savedToken}, tokenToUse: ${tokenToUse}, stage: ${validStage} vs ${tokenColumn}`);

      // if (validStage !== tokenColumn || savedToken !== tokenToUse || savedToken.endsWith("_used")) {
      //   Logger.log(`[TOKEN CHECK] Token Mismatch or Already Used - validStage: ${validStage}, tokenColumn: ${tokenColumn}, token: ${savedToken}`);
      //   return showErrorTokenPage(
      //     "You've Already Responded",
      //     `Looks like you've already taken action on this request.<br>Current stage: <strong>${currentStage}</strong>.`
      //   );
      // }

      // --------------------- Testing Here -----------------
      // Expected approver email from sheet
      const spvEmail = (data[COLUMNS.SPV_EMAIL - 1] || "").toLowerCase();
      const hrEmail = (data[COLUMNS.HR_EMAIL - 1] || "").toLowerCase();
      const gmEmail = (data[COLUMNS.GM_EMAIL - 1] || "").toLowerCase();

      let expectedApprover = "";
      if (stageParam === "spv") expectedApprover = spvEmail;
      else if (stageParam === "hr") expectedApprover = hrEmail;
      else if (stageParam === "gm") expectedApprover = gmEmail;
      else return showErrorTokenPage("Invalid Stage", "Unknown stage: " + stageParam);

      // Get token stored in sheet for this stage
      const savedToken = sheet.getRange(rowIndex, tokenColumn).getValue();
      Logger.log(`[TOKEN CHECK] Tokens - saved: ${savedToken}, received: ${tokenToUse}`);

      // Check token existence and stage
      if (!savedToken || validStage !== tokenColumn) {
        Logger.log("[TOKEN CHECK] Token missing or wrong stage");
        return showErrorTokenPage("Invalid Link", "This approval link is not valid for the current stage.");
      }

      // Check if token already used
      if (savedToken.endsWith("_used")) {
        Logger.log("[TOKEN CHECK] Token already used");
        return showErrorTokenPage("Link Already Used", "This request has already been processed.");
      }

      // Because of limitation of USER_DEPLOYING & USER_ACCESSING in appsscript.json, we decide to disable this function.
      // As we know we have this feature to prevent un-authorized approver, since we moved  to execute the request as a code owner
      // this security check is being disable.

      // 🔒 SECURITY CHECK — Token must contain expected approver’s email
      // if (!savedToken.includes(expectedApprover)) {
      //   Logger.log(`[TOKEN CHECK] Email mismatch! Expected: ${expectedApprover}, Token: ${savedToken}`);
      //   return showErrorTokenPage(
      //     "Not Your Approval",
      //     `Oops! This link is assigned to <b>${expectedApprover}</b> — not you.`
      //   );
      // }

      // Check if token matches exactly
      if (savedToken !== tokenToUse) {
        Logger.log("[TOKEN CHECK] Token mismatch!");
        return showErrorTokenPage("Invalid Token", "The token in this link does not match our records.");
      }

      Logger.log(`[TOKEN CHECK] SUCCESS — Approver validated: ${expectedApprover}`);

      // --------------------- End Here ---------------------

      // Block re-approval if already finalized
      if (["Approved", "Rejected"].includes(currentStatus)) {
        Logger.log(`[FINALIZED CHECK] Already Finalized - Status: ${currentStatus}, Stage: ${currentStage}`);
        const html = HtmlService.createTemplateFromFile('result');
        html.action = currentStatus.toLowerCase();
        html.stage = currentStage;
        html.note = `This request was already processed as ${currentStatus}. ` + sheet.getRange(rowIndex, COLUMNS.NOTE).getValue();
        html.nextStage = "Final";
        return html.evaluate().setTitle("Request Processed");
      }

      // Invalidate the token after use
      Logger.log(`[TOKEN INVALIDATE] Marking token as used: ${tokenToUse}_used`);
      sheet.getRange(rowIndex, tokenColumn).setValue(`${tokenToUse}_used`);

      name = sheet.getRange(rowIndex, COLUMNS.NAME).getValue();
      department = sheet.getRange(rowIndex, COLUMNS.DEPARTMENT).getValue();
      leaveType = sheet.getRange(rowIndex, COLUMNS.LEAVE_TYPE).getValue();
      requester = sheet.getRange(rowIndex, COLUMNS.REQUESTER_EMAIL).getValue();
      startDate = sheet.getRange(rowIndex, COLUMNS.START_DATE).getValue();
      endDate = sheet.getRange(rowIndex, COLUMNS.END_DATE).getValue();
      reasonText = sheet.getRange(rowIndex, COLUMNS.REASON).getValue();

      // Update decision details
      sheet.getRange(rowIndex, COLUMNS.DECISION).setValue(decision + " by " + currentStage); // Be more specific
      sheet.getRange(rowIndex, COLUMNS.NOTE).setValue(note);
      sheet.getRange(rowIndex, COLUMNS.DECISION_DATE).setValue(new Date());

      // === Start Balance Validation Logic ===
      // Dynamically load leave balance data
      const balanceStartRow = 7;
      const balanceData = sheet.getRange(
        balanceStartRow,
        COLUMNS.EMP_EMAIL, // column W (23)
        sheet.getLastRow() - (balanceStartRow - 1),
        COLUMNS.SICK_BALANCE - COLUMNS.EMP_EMAIL + 1 // should be 3 columns (W to Y)
      ).getValues();

      const requesterEmail = requester.toLowerCase();
      let balance = 0;
      let sickBalance = 0;

      for (const row of balanceData) {
        if ((row[0] || "").toLowerCase() === requesterEmail) {
          balance = parseFloat(row[1]) || 0;
          sickBalance = parseFloat(row[2]) || 0;
          break;
        }
      }

      // const days = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;
      const days = calculateLeaveDays(new Date(startDate), new Date(endDate));
      // const leaveTypeLower = leaveType.toLowerCase();

      let balanceType = "Leave";
      if (leaveType.toLowerCase().includes("sick")) balanceType = "Sick Leave";
      if (leaveType.toLowerCase().includes("unpaid")) balanceType = "Unpaid";

      const applicableBalance = balanceType === "Sick Leave" ? sickBalance : balance;

      const leaveTypes = {
        annual: ["Annual Leave", "Bereavement Leave", "Career Leave", "Ceremony Leave", "Other"],
        sick: ["Sick Leave"]
      };

      const cekAnnual = leaveTypes.annual.includes(leaveType);
      const cekSick = leaveTypes.sick.includes(leaveType);
      const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ        

      if ((cekAnnual || cekSick) && days > applicableBalance) {
        const rejectionNote = performAutoReject(
          rowIndex,                          // 1
          applicableBalance,                // 2
          balanceType,                      // 3
          days,                             // 4
          name,                             // 5
          leaveType,                        // 6
          startDate,                        // 7
          endDate,                          // 8
          reasonText,                       // 9
          requester,                        // 10
          spvEmail === requester ? "Auto-Rejected" : "",  // 11
          hrEmail === requester ? "Auto-Rejected" : "",   // 12
          gmEmail === requester ? "Auto-Rejected" : "",    // 13
          refID   // 14
        );

        const html = HtmlService.createTemplateFromFile('result');
        html.action = 'reject';
        html.stage = currentStage;
        html.note = rejectionNote + "<br><br><strong>This request was automatically rejected due to insufficient balance.</strong>"; // include Auto-rejected:
        html.nextStage = "Final";

        Logger.log('Result Action:', html.action);
        Logger.log('Result Note:', html.note);

        return html.evaluate().setTitle("Auto-Rejected");
      }
      // === End Balance Validation Logic ===

      if (action === 'approve') {
        sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Pending'); // Keep pending until final approval

        switch (currentStage) {
          case "SPV Approval":

            if (leaveType === "Working From Home (WFH)") {
              Logger.log(`Auto-finalizing WFH request. Finalizing request as 'Approved by SPV'`);
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by SPV", refID);
            } else {
              nextStage = "HR Review";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              Logger.log(`Moving to next stage: ${nextStage}`);

              // Generate a new token for HR
              const hrToken = generateRandomToken();
              Logger.log(`Generated HR Token: ${hrToken}`);

              // Save token in the sheet in HR_TOKEN column
              sheet.getRange(rowIndex, COLUMNS.HR_TOKEN).setValue(hrToken);
              Logger.log(`Saved HR Token to sheet at row ${rowIndex}, column ${COLUMNS.HR_TOKEN}`);

              // Send approval email to HR
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.HR_EMAIL, nextStage, rowIndex, hrToken, refID);
              Logger.log(`Sent approval email to HR: ${CONFIG.HR_EMAIL} for row ${rowIndex}`);
            }
            break;

          case "HR Review":
            Logger.log("Stage: HR Review");
            // Best Use Cases Logic!
            // | **Submitter** | **Expected Approval Flow**            | **Expected Reject Flow**   | **Acc Flow** | **Reject Flow** |
            // | ------------- | ------------------------------------- | -------------------------- | ------------ | --------------- |
            // | Employee      | SPV → HR → Reporting (V)              | Rejected at the stage of   | Done         | Done            |
            // | SPV           | HR → GM → Reporting (V)               | ~~                         | Done         | Done            |
            // | HR            | GM → Reporting (V)                    |   ~~                       | Done         | Done            | 
            // | GM            | HR → Reporting (V)                    |     ~~                     | Done         | Done            |
            // | Unpaid Leave  | SPV(V) → HR(V) → GM(V) → Reporting (V)|       ~~                   | Done         | Done            | GM Unpaid needs to fix!
            // | WFH           | Emp -> SPV(V), SPV -> HR(V), HR -> GM(V), GM -> HR(V) |            | Done         | Done            |
            //
            //
            //
            // I don't know maybe this function still have bias on same stage, but whis is work fine.
            // Maybe for you the next person who saw this code, you can fix this better. Thanks
            // 
            // With Great Respect,
            // Galbatorix

            const allSpvEmails = Object.values(CONFIG.SPV_MAP);
            const isRequesterSPV = allSpvEmails.includes(requester);
            const isRequesterHR = requester.toLowerCase() === CONFIG.HR_EMAIL.toLowerCase();
            const needsGM = (leaveType === "Unpaid Leave") || isRequesterSPV;

            Logger.log(`Requester: ${requester}`);
            Logger.log(`Leave Type: ${leaveType}`);
            Logger.log(`Is Requester SPV? ${isRequesterSPV}`);
            Logger.log(`Is Requester HR? ${isRequesterHR}`);
            Logger.log(`Needs GM? ${needsGM}`);

            if (leaveType === "Working From Home (WFH)") {
              if (isRequesterHR) {
                Logger.log("WFH submitted by HR → escalate to GM Review");
                nextStage = "GM Review";
                sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);

                const gmToken = generateRandomToken();
                Logger.log(`Generated GM Token: ${gmToken}`);
                sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);

                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken, refID);
                Logger.log("GM approval email sent");
              } else {
                Logger.log("WFH submitted by SPV → finalize after HR Review");
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR", refID);
              }
            } else {
              nextStage = needsGM ? "GM Review" : "Final";
              Logger.log(`Next Stage: ${nextStage}`);
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);

              const gmToken = generateRandomToken();
              Logger.log(`Generated GM Token: ${gmToken}`);
              sheet.getRange(rowIndex, COLUMNS.GM_TOKEN).setValue(gmToken);

              if (needsGM) {
                Logger.log("Sending request to GM for final review");
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                sendApprovalEmail(name, leaveType, startDate, endDate, reasonText, CONFIG.GM_EMAIL, nextStage, rowIndex, gmToken, refID);
                Logger.log("GM approval email sent");
              } else {
                Logger.log("No GM needed → finalizing request");
                const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
                finalizeRequest(rowIndex, decision, note, name, requester, "Approved by HR", refID);
              }
            }
            break;

          case "GM Review":
            Logger.log("Stage: GM Review");
            Logger.log("Leave Type: " + leaveType);
            Logger.log("Requester: " + requester);

            if (leaveType === "Working From Home (WFH)") {
              Logger.log("WFH submitted → GM → Reporting");
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM", refID);
            } else {
              Logger.log("Non-WFH → GM → Finalization");
              nextStage = "Final";
              sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
              const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
              finalizeRequest(rowIndex, decision, note, name, requester, "Approved by GM", refID);
            }
            break;

          default:
            Logger.log("Unexpected Stage: " + currentStage);
            nextStage = "Error in Workflow";
            sheet.getRange(rowIndex, COLUMNS.STAGE).setValue(nextStage);
            const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ
            finalizeRequest(rowIndex, "Error", "Workflow error at stage: " + stage, name, requester, "Workflow Error", refID);
            break;
        }
      } else { // Action is 'reject'
        Logger.log(`Rejecting request at row ${rowIndex} during stage: ${currentStage}`);

        sheet.getRange(rowIndex, COLUMNS.STATUS).setValue('Rejected');
        sheet.getRange(rowIndex, COLUMNS.STAGE).setValue('Rejected at ' + currentStage);

        const rejectionNote = note || `Rejected by ${currentStage}.`;
        Logger.log(`Rejection note: ${rejectionNote}`);

        const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

        const formattedStartDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
        const formattedEndDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
        const leaveDays = Math.ceil((new Date(endDate) - new Date(startDate)) / (1000 * 60 * 60 * 24)) + 1;

        Logger.log(`Leave dates: ${formattedStartDate} to ${formattedEndDate}, total ${leaveDays} day(s)`);

        const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue(); // Column AQ

        const template = HtmlService.createTemplateFromFile("finalNotification");
        template.name = name;
        template.leaveType = leaveType;
        template.startDate = formattedStartDate;
        template.endDate = formattedEndDate;
        template.totalDays = leaveDays;
        template.reason = reasonText;
        template.finalDecision = "Rejected";
        template.finalNote = rejectionNote;
        template.updatedBalance = null;
        template.refID = refID;   // Added
        template.row = rowIndex;  // Should be fix refID not defined.

        template.spvStatus = values[COLUMNS.SPV_DECISION - 1] || "—";
        template.hrStatus = values[COLUMNS.HR_DECISION - 1] || "—";
        template.gmStatus = values[COLUMNS.GM_DECISION - 1] || "—";

        const htmlBody = template.evaluate().getContent();

        GmailApp.sendEmail(
          requester,
          `ONEderland Leave/WFH Request Rejected: ${name}`,
          '', // plain text fallback (leave blank if using only HTML)
          {
            htmlBody: htmlBody,
            name: "ONEderland Approval System"
          }
        );

        Logger.log(`Rejection email sent to: ${requester}`);

        nextStage = 'Final (Rejected)';
        Logger.log(`Stage set to: ${nextStage}`);
      }
    } finally {
      lock.releaseLock();
    }

    const html = HtmlService.createTemplateFromFile('result');
    html.action = action;
    html.stage = currentStage; // The stage that just made the decision
    html.note = note || (action === 'reject' ? "Rejected by " + currentStage : "Approved by " + currentStage);
    html.nextStage = nextStage;
    Logger.log(`[${action.toUpperCase()}] by ${currentStage} | Requester: ${requester} | Note: ${html.note} | Next: ${nextStage}`);
    return html.evaluate().setTitle("Request Processed");

  } catch (err) {
    const stackLine = (err.stack || '').split('\n')[1] || 'Line unknown';
    Logger.log("❌ Error in doGet:\n" + err.toString() + "\n📍 Location: " + stackLine);

    const errorHtml = `
        <html>
          <head>
            <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
            <title>Processing Error</title>
            <style>
              body {
                background-color: #f8f9fa;
              }
              .error-container {
                max-width: 700px;
                margin: 5% auto;
                padding: 2rem;
                background: white;
                border-radius: 15px;
                box-shadow: 0 0.5rem 1rem rgba(0,0,0,.15);
              }
              small {
                font-family: monospace;
                color: #555;
              }
            </style>
          </head>
          <body>
            <div class="container">
              <div class="error-container text-center">
                <h1 class="text-danger mb-4">⚠️ Oops! Something went wrong.</h1>
                <p class="lead">We encountered an error while processing your request.</p>
                <p>This may be due to:</p>
                <ul class="list-group list-group-flush mb-3">
                  <li class="list-group-item">🔄 An outdated or already processed request</li>
                  <li class="list-group-item">🌐 A temporary network issue</li>
                  <li class="list-group-item">⚙️ Internal processing error</li>
                </ul>
                <p class="mb-3">Please try again later or contact support if the problem persists.</p>
                <div class="alert alert-secondary text-start" role="alert">
                  <strong>Error Details:</strong><br>
                  <small>${escapeHtml(err.toString())}<br>${escapeHtml(stackLine)}</small>
                </div>
                <p class="text-muted"><strong>Support:</strong> Please screenshot this page and contact <strong>Suma Antara</strong> or <strong>Iyyan Anugrah.</strong></p>
              </div>
            </div>
          </body>
        </html>
      `;
    Logger.log(`Oops! Something went wrong. ${escapeHtml(err.toString())} - ${escapeHtml(stackLine)}`);
    return HtmlService.createHtmlOutput(errorHtml).setTitle("Processing Error");
  }
}

function submitRequest(name, email, department, leaveType, startDate, endDate, reason) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
    const spvEmail = CONFIG.SPV_MAP[department] || CONFIG.GM_EMAIL; // Default to GM if SPV not found
    const hrEmail = CONFIG.HR_EMAIL;
    const gmEmail = CONFIG.GM_EMAIL;

    const firstDate = parseDMYDate(startDate);
    const lastDate = parseDMYDate(endDate);

    // --- Normalize and rewrite email if DB already contains the canonical version (with +itn etc.)
    const data = sheet.getDataRange().getValues();
    let normalizedEmail = normalizeGmailAddress(email);
    let rewrittenEmail = email.trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const sheetEmailRaw = data[i][COLUMNS.EMP_EMAIL - 1];
      if (!sheetEmailRaw) continue;

      const sheetEmail = sheetEmailRaw.toString().trim().toLowerCase();
      if (normalizeGmailAddress(sheetEmail) === normalizedEmail) {
        // Found matching normalized email in DB → use its exact stored version
        rewrittenEmail = sheetEmail;
        break;
      }
    }

    // Update working email reference
    email = rewrittenEmail;
    normalizedEmail = normalizeGmailAddress(email);

    const isIntern = email.toLowerCase().includes('+itn@');

    // Validation: Interns can only take Unpaid Leave - ITN
    if (isIntern && leaveType !== "Unpaid Leave - ITN") {
      throw new Error("Access Denied: Your authorization is only can take 'Unpaid Leave - ITN'.<br>Contact HR/IT Team if the problem persist.");
    }

    // Optional: block non-interns from choosing ITN leave
    if (!isIntern && leaveType === "Unpaid Leave - ITN") {
      throw new Error("Access Denied: You're not allowed to take this leave.<br>Contact HR/IT Team if the problem persist.");
    }

    Logger.log(`Incoming request: ${name}, Dept: ${department}`);

    if (!firstDate || !lastDate) {
      Logger.log(`Invalid date format received: ${startDate}, ${endDate}`);
      throw new Error("Invalid date format. Please use DD-MM-YYYY.");
    }

    if (lastDate < firstDate) {
      Logger.log(`Date validation failed: endDate (${lastDate}) before startDate (${firstDate})`);
      throw new Error(
        "❌ SYSTEM ERROR: USER IQ NOT FOUND ❌<br>The last day of leave cannot be BEFORE the first day.<br>If you can time-travel, please go back and fix the economy first!<br>Until then… PLEASE enter a valid date! 😭"
      );
    }

    // Determine if the requester is the SPV of their department
    // Javascript?? this is Jawascript hahahahaha

    const ikiEmailnyaSPVngertiOra = (email.toLowerCase() === spvEmail.toLowerCase());
    const ikiEmailnyaGMngertiOra = (email.toLowerCase() === gmEmail.toLowerCase());
    const ikiEmailnyaHRngertiOra = (email.toLowerCase() === hrEmail.toLowerCase());

    let stage;
    let approvalEmail;

    const saumpaniHRprei = ["Annual Leave", "Bereavement Leave", "Career Leave", "Ceremony Leave", "Sick Leave", "Unpaid Leave - FT", "Unpaid Leave - ITN", "Other", "Working From Home (WFH)"]; // If SPV as a HR(Dyah Retno) submit all leave it must set nextstage to GM!

    if (ikiEmailnyaSPVngertiOra) {
      // Case 1: SPV submits their own leave
      if (saumpaniHRprei.includes(leaveType) && ikiEmailnyaHRngertiOra) {
        // HR acts as SPV — must go to GM
        stage = "GM Review";
        approvalEmail = gmEmail;
      } else if (leaveType === "Working From Home (WFH)") {
        // SPV requesting WFH → goes to GM
        stage = "GM Review";
        approvalEmail = gmEmail;
      } else if (leaveType === "Unpaid Leave - FT") {
        // SPV requesting Unpaid - FT → GM after HR
        stage = "HR Review";
        approvalEmail = hrEmail;
      } else if (leaveType === "Unpaid Leave - ITN") {
        // SPV requesting Unpaid - ITN → HR directly
        stage = "HR Review";
        approvalEmail = hrEmail;
      } else {
        // Default SPV case → goes to HR
        stage = "HR Review";
        approvalEmail = hrEmail;
      }

    } else if (ikiEmailnyaGMngertiOra) {
      // Case 2: GM submits any leave
      stage = "HR Review";
      approvalEmail = hrEmail;

    } else {
      // Case 3: Normal employee (not SPV or GM)
      if (leaveType === "Unpaid Leave - ITN") {
        // Intern → SPV then HR → Final
        stage = "SPV Approval";
        approvalEmail = spvEmail;
      } else if (leaveType === "Unpaid Leave - FT") {
        // Full-time unpaid leave → SPV then HR then GM
        stage = "SPV Approval";
        approvalEmail = spvEmail;
      } else if (leaveType === "Working From Home (WFH)") {
        // WFH → SPV
        stage = "SPV Approval";
        approvalEmail = spvEmail;
      } else {
        // Default → SPV Approval
        stage = "SPV Approval";
        approvalEmail = spvEmail;
      }
    }

    Logger.log(`Next stage: ${stage}, Approver: ${approvalEmail}`);

    const spvToken = generateRandomToken();
    const hrToken = generateRandomToken();
    const gmToken = generateRandomToken();

    let tokenToUse;
    if (stage === "SPV Approval") {
      tokenToUse = spvToken;
    } else if (stage === "HR Review") {
      tokenToUse = hrToken;
    } else if (stage === "GM Review") {
      tokenToUse = gmToken;
    }

    // Generate RefID for this submission
    let refID = generateReferenceID();

    // Append the request to the sheet
    const newRow = sheet.appendRow([
      new Date(), name, department, leaveType,
      firstDate, lastDate, reason,
      "Pending", email, spvEmail, hrEmail, gmEmail, stage,
      "", "", new Date(), // Decision, Note, Decision Date
      "", "", "",         // SPV_DECISION, HR_DECISION, GM_DECISION
      spvToken, hrToken, gmToken,
      "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", // fill until col AQ
      "" // placeholder for RefID
    ]);

    // Get row number
    const lastRow = sheet.getLastRow();

    // Append row number to RefID
    refID = `${refID}${lastRow}`;

    // Save into AQ
    sheet.getRange(lastRow, COLUMNS.REF_ID).setValue(refID);

    const rowIndex = sheet.getLastRow();

    // Send confirmation to requester
    Logger.log(`Sending confirmation email to ${email} with Form ID = ${refID}`);
    sendSubmissionConfirmation(email, name, leaveType, firstDate, lastDate, reason, stage, refID);

    // Send approval request to next stage
    Logger.log(`Sending approval email to SPV: ${spvEmail}, Form ID = ${refID}`);
    sendApprovalEmail(name, leaveType, firstDate, lastDate, reason, approvalEmail, stage, rowIndex, tokenToUse, refID);

    Logger.log(`Submit Request completed successfully for ${name} - ${refID}`);
    return "Success";

  } catch (e) {
    Logger.log("Submit error: " + e.toString() + " Stack: " + e.stack);
    return e.message || "Submission failed due to a server error. Please try again.";
  }
}

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function finalizeRequest(row, decision, note, name, requesterEmail, finalApprovalStageNote, refID) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
  const finalNote = note || finalApprovalStageNote || "No Notes";

  const leaveType = sheet.getRange(row, COLUMNS.LEAVE_TYPE).getValue();
  const startDate = sheet.getRange(row, COLUMNS.START_DATE).getValue();
  const endDate = sheet.getRange(row, COLUMNS.END_DATE).getValue();
  const reason = sheet.getRange(row, COLUMNS.REASON).getValue();
  const spvStatus = sheet.getRange(row, COLUMNS.SPV_DECISION).getValue();
  const hrStatus = sheet.getRange(row, COLUMNS.HR_DECISION).getValue() || '';
  const gmStatus = sheet.getRange(row, COLUMNS.GM_DECISION).getValue() || '';
  const department = sheet.getRange(row, COLUMNS.DEPARTMENT).getValue();
  const days = calculateLeaveDays(startDate, endDate);

  const leaveTypes = {
    annual: ["Annual Leave", "Bereavement Leave", "Career Leave", "Ceremony Leave", "Other"],
    sick: ["Sick Leave"]
  };

  const balanceStartRow = 7;
  const balanceRange = sheet.getRange(balanceStartRow, COLUMNS.EMP_EMAIL, sheet.getLastRow() - (balanceStartRow - 1), 3).getValues();
  const requester = requesterEmail.toLowerCase();
  let currentLeave = 0, currentSick = 0, updatedBalance = null;

  for (let i = 0; i < balanceRange.length; i++) {
    const [email, leaveBal, sickBal] = balanceRange[i];
    if (!email) continue;

    if (email.toLowerCase() === requester) {
      currentLeave = parseFloat(leaveBal) || 0;
      currentSick = parseFloat(sickBal) || 0;

      if (decision === "Approved") {
        const isAnnual = leaveTypes.annual.includes(leaveType);
        const isSick = leaveTypes.sick.includes(leaveType);

        if (isAnnual && days > currentLeave) {
          performAutoReject(row, currentLeave, "Annual", days, name, leaveType, startDate, endDate, reason, requesterEmail, spvStatus, hrStatus, gmStatus, refID);
          return;
        }

        if (isSick && days > currentSick) {
          performAutoReject(row, currentSick, "Sick", days, name, leaveType, startDate, endDate, reason, requesterEmail, spvStatus, hrStatus, gmStatus, refID);
          return;
        }

        const rowOffset = balanceStartRow + i;
        if (isAnnual) {
          const newLeave = Math.max(0, currentLeave - days);
          sheet.getRange(rowOffset, COLUMNS.ANNUAL_BALANCE).setValue(newLeave);
          updatedBalance = { leave: newLeave, sick: currentSick };
        } else if (isSick) {
          const newSick = Math.max(0, currentSick - days);
          sheet.getRange(rowOffset, COLUMNS.SICK_BALANCE).setValue(newSick);
          updatedBalance = { leave: currentLeave, sick: newSick };
        }
      }

      break;
    }
  }

  // Update final approval status
  sheet.getRange(row, COLUMNS.STATUS).setValue(decision);
  sheet.getRange(row, COLUMNS.STAGE).setValue("Completed");
  sheet.getRange(row, COLUMNS.NOTE).setValue(finalNote);

  // === SMART MERGE WEEKDAY EVENTS ===
  const wfhCalendarId = "63923981f7916d39b1e2cc1dc3f74def45df9578ee045429c2c14256114ff10a@group.calendar.google.com";
  const leaveCalendarId = "acd3tof9di4puvf3fd046naeks@group.calendar.google.com";

  const targetCalendarId = leaveType === "Working From Home (WFH)" ? wfhCalendarId : leaveCalendarId;
  const calendar = CalendarApp.getCalendarById(targetCalendarId);

  if (calendar) {
    const title = `${name}'s - ${leaveType}`;
    const description =
      `Form ID: ${refID}\n` +
      `Department: ${department}\n` +
      `Note: ${finalNote}`;

    let blockStart = null;
    let currentDate = new Date(startDate);
    const finalDate = new Date(endDate);

    while (currentDate <= finalDate) {
      const day = currentDate.getDay(); // 0=Sunday, 6=Saturday

      if (day !== 0 && day !== 6) {
        if (!blockStart) {
          blockStart = new Date(currentDate); // Start a new block
        }
      } else {
        if (blockStart) {
          const blockEnd = new Date(currentDate);
          blockEnd.setDate(blockEnd.getDate()); // End boundary (weekend)
          calendar.createAllDayEvent(title, blockStart, blockEnd, {
            description
          });
          Logger.log(`📅 Created block: ${blockStart} → ${blockEnd}`);
          blockStart = null;
        }
      }

      currentDate.setDate(currentDate.getDate() + 1);
    }

    // Create final block if exists
    if (blockStart) {
      const blockEnd = new Date(finalDate);
      blockEnd.setDate(blockEnd.getDate() + 1);
      calendar.createAllDayEvent(title, blockStart, blockEnd, {
        description
      });
    }

    Logger.log(`🎯 Completed | ${name}'s ${leaveType} created on calendar: ${targetCalendarId}`);
  } else {
    Logger.log(`❌ Calendar not found: ${targetCalendarId}`);
  }
  // === END ===

  // Prepare notification data
  const formattedStart = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  const formattedEnd = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");

  // Final Notification Email to Requester
  const template = HtmlService.createTemplateFromFile('finalNotification');
  Object.assign(template, {
    name, leaveType, reason,
    startDate: formattedStart,
    endDate: formattedEnd,
    totalDays: days,
    spvStatus, hrStatus, gmStatus,
    finalDecision: decision,
    finalNote,
    updatedBalance,
    refID,
    row
  });

  const htmlBody = template.evaluate().getContent();
  GmailApp.sendEmail(requesterEmail, `ONEderland Leave Request ${decision}: ${name}`, '', {
    htmlBody,
    name: 'ONEderland Approval System'
  });

  // Reporting Team Notification with Calendar Link
  // const calendarTitle = `${name}'s ${leaveType}`;
  // const calendarDescription = `Leave Request\nRequester: ${name}\nDepartment: ${department}\nType: ${leaveType}\nDecision: ${decision}\nNote: ${finalNote}`;
  // const calendarLink = generateCalendarLink(calendarTitle, startDate, endDate, calendarDescription);

  const reportingHtml = `
    <div style="font-family: Arial, sans-serif; background-color: #ffffff; margin: 0; padding: 0;">
      <table align="center" width="100%" style="max-width: 650px; margin: auto; border: 1px solid #e0e0e0; border-radius: 10px; box-shadow: 0 1px 6px rgba(0,0,0,0.08); background: #fff;">
        <tr>
          <td style="padding: 22px 28px;">

            <!-- Header -->
            <h2 style="margin: 0 0 15px 0; font-size: 20px; color: #2c3e50; text-align: center;">
              Leave / WFH Request:
              <span style="color: ${decision === 'Approved' ? '#28a745' : '#dc3545'};">
                ${decision}
              </span>
            </h2>

            <p style="font-size: 14.5px; color: #555; margin-top: 4px; text-align: center;">
              Final Notification — Request Complete
            </p>

            <!-- Table -->
            <table width="100%" cellpadding="0" cellspacing="0" style="margin-top: 22px; border-collapse: collapse; font-size: 15px;">
              
              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7; width:160px;">Form ID</td>
                <td style="padding:10px;">${refID}</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7;">Requester</td>
                <td style="padding:10px;">${name}</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7;">Department</td>
                <td style="padding:10px;">${department}</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7;">Leave Type</td>
                <td style="padding:10px;">${leaveType}</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7;">Dates</td>
                <td style="padding:10px;">${formattedStart} → ${formattedEnd} (${days} day${days > 1 ? 's' : ''})</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7; vertical-align: top;">Reason</td>
                <td style="padding:10px;">${reason || 'N/A'}</td>
              </tr>

              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7; vertical-align: top;">Note</td>
                <td style="padding:10px;">${finalNote || 'N/A'}</td>
              </tr>

              ${updatedBalance ? `
              <tr>
                <td style="padding:10px; font-weight:bold; background:#f7f7f7;">
                  Updated Leave Balance
                </td>
                <td style="padding:10px;">
                  Annual Leave: <strong>${updatedBalance.leave}</strong><br>
                  Sick Leave: <strong>${updatedBalance.sick}</strong>
                </td>
              </tr>
              ` : ''}

            </table>

            <!-- Calendar Note -->
            <p style="margin-top: 20px; font-size: 14px; color: #666;">
              ✔️ This request has been automatically added to the company calendar.
            </p>

            <!-- Footer -->
            <div style="margin-top: 25px; text-align: center; font-size: 13px; color: #999;">
              Thank you for using ONEderland Leave System.  
              <br>Need help? Contact HR support.
            </div>

          </td>
        </tr>
      </table>
    </div>
  `;

  CONFIG.REPORTING_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, `Reporting Notification from ${name}`, '', {
      htmlBody: reportingHtml,
      name: 'ONEderland Approval System'
    });
  });
}

function generateCalendarLink(title, startDate, endDate, description) {
  const formatDate = date => Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyyMMdd");

  const start = formatDate(startDate);
  const end = formatDate(new Date(new Date(endDate).getTime() + 86400000)); // +1 day to make it inclusive

  const details = encodeURIComponent(description);
  const calTitle = encodeURIComponent(title);

  Logger.log(`Calender Created to: ${title} for ${startDate} - ${endDate}`);
  return `https://www.google.com/calendar/render?action=TEMPLATE&text=${calTitle}&dates=${start}/${end}&details=${details}`;
}

function sendApprovalEmail(name, leaveType, startDate, endDate, reason, approverEmail, stage, row, tokenToUse, refID) {
  const baseUrl = ScriptApp.getService().getUrl();

  const template = HtmlService.createTemplateFromFile("emailtemplate");
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = formatDate(startDate);
  template.endDate = formatDate(endDate);
  template.reason = reason;
  template.stage = stage;
  template.refID = refID;
  template.baseUrl = baseUrl;

  // Extract normalized stage (lowercase) → Used for URL param
  const shortStage = stage.toLowerCase().includes("spv") ? "spv"
    : stage.toLowerCase().includes("hr") ? "hr"
      : stage.toLowerCase().includes("gm") ? "gm"
        : "unknown"; // fallback

  if (!tokenToUse || shortStage === "unknown") {
    Logger.log(`❗ Missing token or invalid stage in sendApprovalEmail for stage: ${stage}, row: ${row}`);
    return;
  }

  // Safe & encoded approval/rejection URLs
  const encodedToken = encodeURIComponent(tokenToUse);
  const noteText = `Approved at ${stage}`;
  template.approveUrl = `${baseUrl}?action=approve&stage=${shortStage}&row=${row}&token=${encodedToken}&note=${encodeURIComponent(noteText)}`;
  template.rejectUrl = `${baseUrl}?action=reject&stage=${shortStage}&row=${row}&token=${encodedToken}&note=`; // will be filled in by user

  try {
    MailApp.sendEmail({
      to: approverEmail,
      subject: `[Action Required] Leave/WFH Request: ${name} (${stage})`,
      htmlBody: template.evaluate().getContent(),
      name: "ONEderland Approval System"
    });
    Logger.log(`✅ Approval email sent to ${approverEmail} for ${name} - ${refID}`);
  } catch (e) {
    Logger.log(`❌ Failed to send approval email to ${approverEmail} for ${name} - ${refID}. Error: ${e}`);
  }
}

function sendSubmissionConfirmation(email, name, leaveType, startDate, endDate, reason, stage, refID) {
  const scriptUrl = ScriptApp.getService().getUrl();
  const trackingLink = `${scriptUrl}?track=${encodeURIComponent(email)}`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:600px; margin:auto; border:1px solid #ddd; padding:20px;">
      <h2 style="color:#2c3e50; border-bottom:1px solid #eee; padding-bottom:10px;">Leave Request Submitted</h2>
      <p>Dear ${name},</p>
      <p>Your leave request has been successfully submitted and is now awaiting <strong>${stage}</strong>.</p>
      
      <table style="width:100%;border-collapse:collapse;margin:20px 0;">
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;width:30%;"><strong>Form ID</strong></td><td style="padding:8px;border:1px solid #ddd;">${refID}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;width:30%;"><strong>Leave Type</strong></td><td style="padding:8px;border:1px solid #ddd;">${leaveType}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Dates</strong></td><td style="padding:8px;border:1px solid #ddd;">${formatDate(startDate)} to ${formatDate(endDate)}</td></tr>
        <tr><td style="padding:8px;border:1px solid #ddd;background-color:#f9f9f9;"><strong>Reason</strong></td><td style="padding:8px;border:1px solid #ddd;">${escapeHtml(reason)}</td></tr>
      </table>
      
      <p>You can track your leave request status anytime using the link below:</p>
      <p>
        <a href="${trackingLink}" target="_blank" style="display:inline-block;padding:10px 15px;background-color:#3498db;color:#fff;text-decoration:none;border-radius:5px;">
          Track My Leave Request
        </a>
      </p>

      <p>You will receive further email notifications as your request is processed.</p>
      <p style="color:#7f8c8d;font-size:0.9em;margin-top:20px;">Thank you,<br/>ONEderland Approval System</p>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: email,
      subject: "ONEderland Leave/WFH Request Submission Confirmation",
      htmlBody: htmlBody,
      name: "ONEderland Approval System"
    });
  } catch (e) {
    Logger.log(`Failed to send confirmation email to ${email}. Error: ${e.toString}`);
  }
}

// Update Balance and detect balance on colums.email

// Different beetwen parseInt and parseFloat
// | Input    | `parseInt()` | `parseFloat()` | Recommended                        |
// | -------- | ------------ | -------------- | ---------------------------------- |
// | `"1.5"`  | `1`          | `1.5`          | `parseFloat` if you need decimal   |
// | `"18,5"` | `18`         | `NaN`          | `parseFloat(...replace(',', '.'))` |
// | `"100"`  | `100`        | `100`          | Either works                       |

function getLeaveBalanceByEmail(email) {
  if (!email || typeof email !== 'string') {
    Logger.log("[ERROR] Invalid email input: %s", email);
    throw new Error("Invalid email passed to getLeaveBalanceByEmail");
  }

  const sheet = SpreadsheetApp.getActive().getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();

  // normalize input email
  const inputEmail = normalizeGmailAddress(email);

  for (let i = 1; i < data.length; i++) {
    const rowEmailRaw = data[i][COLUMNS.EMP_EMAIL - 1];
    if (!rowEmailRaw) continue;

    // normalize sheet email
    const rowEmail = normalizeGmailAddress(rowEmailRaw);

    if (rowEmail === inputEmail) {
      const leaveBalance = parseFloat(data[i][COLUMNS.ANNUAL_BALANCE - 1]) || 0;
      const sickBalance = parseFloat(data[i][COLUMNS.SICK_BALANCE - 1]) || 0;

      Logger.log("[INFO] Found balance for %s: Leave = %s, Sick = %s", inputEmail, leaveBalance, sickBalance);

      return {
        leave: leaveBalance,
        sick: sickBalance
      };
    }
  }

  Logger.log("[WARN] No balance found for email: %s", inputEmail);
  return null;
}

function calculateLeaveDays(startDate, endDate) {
  let start = new Date(startDate);
  let end = new Date(endDate);
  let count = 0;

  while (start <= end) {
    const day = start.getDay();
    if (day !== 0 && day !== 6) { // Mon-Fri only
      count++;
    }
    start.setDate(start.getDate() + 1);
  }

  return count;
}

function performAutoReject(rowIndex, balance, balanceType, days, name, leaveType, startDate, endDate, reason, requester, spvStatus, hrStatus, gmStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const rejectionNote = `Auto-rejected, Only ${balance} ${balanceType.toLowerCase()} day(s) left, but ${days} requested.`;
  const systemNote = `Your request was rejected because you only have ${balance} ${balanceType.toLowerCase()} day(s) left, but you requested ${days} day(s).`;

  // Update sheet
  sheet.getRange(rowIndex, COLUMNS.STATUS).setValue("Rejected");
  sheet.getRange(rowIndex, COLUMNS.STAGE).setValue("Completed");
  sheet.getRange(rowIndex, COLUMNS.NOTE).setValue(rejectionNote);

  // Fetch current leave balances from "Request" sheet (W: email, X: leave, Y: sick)
  const balanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const balanceStartRow = 7;
  const balanceRange = balanceSheet.getRange(balanceStartRow, COLUMNS.EMP_EMAIL, balanceSheet.getLastRow() - (balanceStartRow - 1), 3).getValues();

  let currentLeave = 0, currentSick = 0;
  const requesterLower = requester.toLowerCase();
  for (const [email, leaveBal, sickBal] of balanceRange) {
    if (email && email.toLowerCase() === requesterLower) {
      currentLeave = parseFloat(leaveBal) || 0;
      currentSick = parseFloat(sickBal) || 0;
      break;
    }
  }

  // Fetch RefID from sheet
  const refID = sheet.getRange(rowIndex, COLUMNS.REF_ID).getValue();

  // Final rejection email
  const template = HtmlService.createTemplateFromFile('finalNotification');
  template.name = name;
  template.leaveType = leaveType;
  template.startDate = Utilities.formatDate(new Date(startDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  template.endDate = Utilities.formatDate(new Date(endDate), Session.getScriptTimeZone(), "dd-MM-yyyy");
  template.totalDays = days;
  template.reason = reason;
  template.spvStatus = spvStatus;
  template.hrStatus = hrStatus;
  template.gmStatus = gmStatus;
  template.finalDecision = "Rejected";
  template.finalNote = systemNote; // Use updated extra note
  template.refID = refID;   // properly assigned
  template.updatedBalance = {
    leave: currentLeave,
    sick: currentSick
  };

  const htmlBody = template.evaluate().getContent();
  GmailApp.sendEmail(requester, `ONEderland Leave Request Rejected: ${name}`, '', {
    htmlBody: htmlBody,
    name: 'ONEderland Approval System'
  });

  Logger.log('🚨 performAutoReject triggered. Returning note: ' + rejectionNote);

  return rejectionNote;
}

function formatDateShort(date, withTime = false) {
  if (!(date instanceof Date)) date = new Date(date);
  const options = { year: 'numeric', month: 'long', day: 'numeric' };
  if (withTime) options.hour = '2-digit', options.minute = '2-digit';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), withTime ? "MMMM d, yyyy HH:mm" : "MMMM d, yyyy");
}

function renderTrackingPage(email, showHistory = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Requests");
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // Skip header

  const filteredRows = rows.filter(row => {
    const emailMatch = row[COLUMNS.REQUESTER_EMAIL - 1] === email;
    const isPending = row[COLUMNS.STATUS - 1].toLowerCase().includes("pending");
    const hasToken = !!row[COLUMNS.SPV_TOKEN - 1];
    return emailMatch && (showHistory ? true : (isPending && hasToken));
  });

  if (filteredRows.length === 0) {
    return `
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <div class="container mt-5">
        <div class="alert alert-warning">
          <h4>No ${showHistory ? "leave history" : "active pending"} requests found for <code>${email}</code></h4>
          <p>Either they have been approved/rejected or the email is incorrect.</p>
        </div>
      </div>
    `;
  }

  let html = `
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <div class="container mt-5">
      <h2 class="mb-4">${showHistory ? "Leave Request History" : "Pending Leave Requests"} for <code>${email}</code></h2>
      <table class="table table-bordered table-striped">
        <thead class="table-light">
          <tr>
            <th>Submitted</th>
            <th>Name</th>
            <th>Leave Type</th>
            <th>First Day</th>
            <th>Last Day</th>
            <th>Status</th>
            <th>Stage</th>
            <th>Form ID</th>
          </tr>
        </thead>
        <tbody>
  `;

  filteredRows.forEach(row => {
    html += `
      <tr>
        <td>${formatDateShort(row[COLUMNS.TIMESTAMP - 1], true)}</td>
        <td>${escapeHtml(row[COLUMNS.NAME - 1])}</td>
        <td>${escapeHtml(row[COLUMNS.LEAVE_TYPE - 1])}</td>
        <td>${formatDateShort(row[COLUMNS.START_DATE - 1])}</td>
        <td>${formatDateShort(row[COLUMNS.END_DATE - 1])}</td>
        <td>${escapeHtml(row[COLUMNS.STATUS - 1])}</td>
        <td>${escapeHtml(row[COLUMNS.STAGE - 1])}</td>
        <td>${escapeHtml(row[COLUMNS.REF_ID - 1])}</td>
      </tr>
    `;
  });

  return html;
}

function showMaintenancePage() {
  return HtmlService.createTemplateFromFile("maintenance")
    .evaluate()
    .setTitle("System Maintenance")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// We're using this 2 funtion to call a button in sheet!
// Adding IMPORTRANGE from another spreadsheet

// To all IT/Webdev in ONEderland Enterprise, if found this obfuscated function, hell yeah.
// You're closer to migrain ~xD wkwkwkwkwk

function _0xb0f5() {
  return [
    'getActiveSpreadsheet',
    'getActiveSheet',
    'getRange',
    'setFormula',
    '=IMPORTRANGE("1rytHgB8Td08XUIQCkvcDEptQrj4F6GD1NotDrX9ulvE","Admin-Leave-Form!I',
    ':S',
    '")'
  ];
}

const _obf = function (i) { return _0xb0f5()[i]; };

function _f1ll() {
  const ss = SpreadsheetApp[_obf(0)]();
  const sheet = ss[_obf(1)]();
  const r0 = 7, c0 = 31;

  for (let i = 0; i < 107; i++) {
    const r = 3 + i;
    const formula = _obf(4) + r + _obf(5) + r + _obf(6);
    sheet[_obf(2)](r0 + i, c0)[_obf(3)](formula);
  }
}

// Adding Button to make more easy to fill a sync leave from original data

function _0x2ab2() { const _0x5e0f6f = ['getActiveSheet', 'getActiveSpreadsheet', 'setFormula', 'getRange', 'AF', 'AI', 'AJ', 'AK', 'AL', 'AM']; _0x2ab2 = function () { return _0x5e0f6f; }; return _0x2ab2(); }
const _0xabc = function (_0x1e1e12) { return _0x2ab2()[_0x1e1e12]; };

function _0xsync() {
  const s = SpreadsheetApp[_0xabc(1)]()[_0xabc(0)]();
  const f = [_0xabc(4), _0xabc(5), _0xabc(6), _0xabc(7), _0xabc(8), _0xabc(9)]; // 'AF' to 'AM'

  for (let r = 7; r <= 107; r++) {
    for (let c = 0; c < f.length; c++) {
      s[_0xabc(3)](r, 23 + c)[_0xabc(2)]('=' + f[c] + r);
    }
  }
}
