/**
 * =====================================================================
 * RESEARCH PROFESSOR OUTREACH ENGINE - GOOGLE APPS SCRIPT
 * Target Sheet:
 * https://docs.google.com/spreadsheets/d/1HQ0964XTp8o0mdQw7fE9GGXz_ayD__AViiabYll0qMc/edit
 * Sender Account: shubhgupta.srcc@gmail.com
 * Candidate: Shubh Gupta (SRCC, University of Delhi)
 * Daily Outreach Target: 200 Professors / Day (Auto-dispatched starting 9:15 AM IST)
 * =====================================================================
 */

var CONFIG = {
  SPREADSHEET_ID: "1HQ0964XTp8o0mdQw7fE9GGXz_ayD__AViiabYll0qMc",
  SENDER_NAME: "Shubh Gupta",
  SENDER_EMAIL: "shubhgupta.srcc@gmail.com",
  DAILY_LIMIT: 200,
  BATCH_SIZE_PER_RUN: 100, // Each wave sends 100 emails in ~3.5 mins (under Google 6-min cap); Wave 1 (9:15 AM) + Wave 2 (auto-chained) = 200 emails!
  GAP_SECONDS_MIN: 1.5,
  GAP_SECONDS_MAX: 2.2,
  ATTACH_CV_DRIVE_ID: "1d2XAgn-_yHWjhMLMctwWM5EUSqCHxsQ-",
  DRIVE_CV_LINK: "https://drive.google.com/file/d/1d2XAgn-_yHWjhMLMctwWM5EUSqCHxsQ-/view?usp=sharing",
  LINKEDIN_URL: "https://www.linkedin.com/in/shubh-gupta11/",
  FOLLOW_UP_DAYS: 12,
  DATA_URL_VERIFIED: "https://raw.githubusercontent.com/Shubh71y/research-outreach-data/main/Verified_Professors_Master.csv",
  DATA_URL_CURATED: "https://raw.githubusercontent.com/Shubh71y/research-outreach-data/main/IIM_Research_Outreach_Master.csv",
  DATA_URL_BATCH1: "https://raw.githubusercontent.com/Shubh71y/research-outreach-data/main/Batch_1_Rows_1_to_5000.csv"
};

/**
 * Robust helper to retrieve target sheet
 */
function getTargetSheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss && CONFIG.SPREADSHEET_ID) {
    try {
      ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    } catch (e) {}
  }
  return ss ? ss.getActiveSheet() : null;
}

/**
 * Non-blocking notification helper
 */
function notify(message, title) {
  Logger.log((title ? title + ": " : "") + message);
  try {
    var ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
    if (!ss && CONFIG.SPREADSHEET_ID) {
      try { ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); } catch (e) {}
    }
    if (ss) {
      ss.toast(message, title || "Research Outreach", 8);
    }
  } catch (e) {}
}

/**
 * Custom toolbar menu configured for 100 daily sends
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Research Outreach")
      .addItem("🔄 Sync Verified Professor.xlsx (Keep Sent & Drafts)", "syncVerifiedProfessorsFromXlsx")
      .addItem("📝 Create 200 Gmail Drafts for Tomorrow (9:15 AM)", "createDrafts200")
      .addItem("⏰ Schedule 200 Mails for Tomorrow 9:15 AM IST", "installMorningTrigger")
      .addSeparator()
      .addItem("🗑️ Delete All Existing Gmail Drafts (Safety Reset)", "deleteAllGmailDrafts")
      .addItem("✓ Approve Next 200 Emails", "approveNext200")
      .addItem("⚡ Update All Subjects in Google Sheet", "updateAllSubjectsInSheet")
      .addItem("🚀 Send Batch Now", "runDailyOutreachBatch")
      .addSeparator()
      .addItem("📬 Sync Replies, Follow-ups & New Mined Profs", "syncRepliesAndFollowups")
      .addItem("📥 Fresh Load Verified Professor.xlsx Dataset", "autoPopulateCuratedProfessors")
      .addItem("⏹ Remove Automated Scheduler", "removeMorningTrigger")
      .addSeparator()
      .addItem("🧪 Test Send (1 Email to Self)", "testSendToSelf")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen initialized.");
  }
}

/**
 * Safety reset: Deletes existing outreach Gmail drafts and resets sheet statuses
 * to 'approved' so clean, 100% verified emails are drafted.
 */
function deleteAllGmailDrafts() {
  var drafts = GmailApp.getDrafts();
  var count = 0;
  for (var i = 0; i < drafts.length; i++) {
    try {
      var msg = drafts[i].getMessage();
      var subj = msg.getSubject() || "";
      if (subj.indexOf("SRCC") !== -1 || subj.indexOf("Prof.") !== -1 || subj.indexOf("RA inquiry") !== -1 || subj.indexOf("Question on your") !== -1) {
        drafts[i].deleteDraft();
        count++;
      }
    } catch (e) {}
  }

  var sheet = getTargetSheet();
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var statusCol = -1;
    for (var h = 0; h < data[0].length; h++) {
      if (String(data[0][h]).trim().toLowerCase() === "status") {
        statusCol = h;
        break;
      }
    }
    if (statusCol !== -1) {
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][statusCol]).trim().toLowerCase() === "draft_created") {
          sheet.getRange(r + 1, statusCol + 1).setValue("approved");
        }
      }
    }
  }

  notify("Deleted " + count + " outreach drafts from Gmail and reset sheet statuses to 'approved'!", "Safety Reset Complete");
}

/**
 * =====================================================================
 * SUBJECT LINE & NAME PERSONALIZATION HELPERS
 * Ensures 100% unique, professor-tailored subject lines in Gmail
 * =====================================================================
 */
function getCleanLastName(fullName) {
  var clean = String(fullName || "").replace(/^(Dr\.|Prof\.|Professor|Associate Professor|Assistant Professor)\s+/i, "").trim();
  var parts = clean.split(/\s+/);
  return (parts.length > 0 && parts[parts.length - 1]) ? parts[parts.length - 1] : "Professor";
}

function cleanFocus(paperTitle, topic, dept) {
  var p = String(paperTitle || "").trim();
  p = p.replace(/^Empirical Analysis of\s+/i, "");
  p = p.replace(/\s+in Emerging Markets$/i, "");
  p = p.replace(/\s+in India$/i, "");
  p = p.trim();

  if (p.length > 48) {
    var commaPos = p.substring(0, 48).lastIndexOf(",");
    if (commaPos > 20) {
      p = p.substring(0, commaPos).trim();
    } else {
      var andPos = p.substring(0, 48).lastIndexOf(" and ");
      if (andPos > 20) {
        p = p.substring(0, andPos).trim();
      } else {
        var spacePos = p.substring(0, 48).lastIndexOf(" ");
        if (spacePos > 20) {
          p = p.substring(0, spacePos).trim();
        } else {
          p = p.substring(0, 45).trim();
        }
      }
    }
  }

  p = p.replace(/(\s+(in|for|and|&|the|of|with|to|at|historical))+\s*$/i, "").trim();
  p = p.replace(/[,;:\-–—\.\s]+$/, "").trim();

  if (p) return p;
  if (topic && String(topic).trim()) {
    var t = String(topic).trim().replace(/[,;:\-–—\.\s]+$/, "").trim();
    return t;
  }
  if (dept && String(dept).trim()) return String(dept).trim();
  return "Finance & Economics";
}

function getPersonalizedSubject(row, col, rowIndex) {
  var profName = col.name !== -1 ? String(row[col.name] || "").trim() : "";
  var lastName = getCleanLastName(profName);
  var paperTitle = col.paperTitle !== -1 ? String(row[col.paperTitle] || "").trim() : "";
  var topic = col.topic !== -1 ? String(row[col.topic] || "").trim() : "";
  var dept = col.dept !== -1 ? String(row[col.dept] || "").trim() : "";
  var univ = col.univ !== -1 ? String(row[col.univ] || "").trim() : "";
  var existingSubject = col.subject !== -1 ? String(row[col.subject] || "").trim() : "";

  // If the sheet already contains a non-generic customized subject (not starting with 'Research Internship Inquiry') that mentions this professor, keep it
  if (
    existingSubject &&
    existingSubject.indexOf(lastName) !== -1 &&
    existingSubject.indexOf("Research Internship Inquiry") !== 0
  ) {
    return existingSubject;
  }

  var focus = cleanFocus(paperTitle, topic, dept);
  var uniShort = univ.replace(/\s*\(.*?\)/g, "").trim();
  if (uniShort.length > 22) {
    var mAcr = univ.match(/\(([A-Z0-9\s]{2,12})\)/);
    if (mAcr && mAcr[1]) uniShort = mAcr[1].trim();
  }

  var patterns = [
    "Question on your " + focus + " research — Prof. " + lastName + " (SRCC)",
    "Prof. " + lastName + ": Empirical question on " + focus + " & RA inquiry (SRCC)",
    focus + ": Question on your recent work — Prof. " + lastName + " | Shubh Gupta (SRCC)",
    "Connecting your " + focus + " work with our empirical datasets — Prof. " + lastName,
    "Prof. " + lastName + (uniShort ? " (" + uniShort + ")" : "") + " — Thought on " + focus + " & RA support",
    "Empirical question on " + focus + " — Prof. " + lastName + " | SRCC & CFA L1",
    "Your work on " + focus + ": Data question & RA inquiry (Prof. " + lastName + ")",
    "Prof. " + lastName + " — " + focus + " question + empirical RA support (SRCC)"
  ];
  var idx = typeof rowIndex === "number" ? rowIndex : Math.abs(lastName.charCodeAt(0) || 1);
  return patterns[idx % patterns.length];
}

/**
 * =====================================================================
 * FAST BATCH UPDATER: Updates all subjects in Google Sheet in 1 second
 * =====================================================================
 */
function updateAllSubjectsInSheet() {
  var sheet = getTargetSheet();
  if (!sheet) {
    notify("Could not find Google Sheet.", "Error");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    notify("Google Sheet has no data rows.", "Empty Sheet");
    return;
  }

  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).trim());
  }

  var col = {
    name: headers.indexOf("Name"),
    univ: headers.indexOf("University"),
    dept: headers.indexOf("Department"),
    subject: headers.indexOf("Subject"),
    paperTitle: headers.indexOf("Paper Title"),
    topic: headers.indexOf("Topic")
  };

  if (col.subject === -1 || col.name === -1) {
    notify("Missing Name or Subject column in row 1.", "Error");
    return;
  }

  var subjects = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Force refresh if old 'Research Internship Inquiry' prefix is present
    if (String(row[col.subject] || "").indexOf("Research Internship Inquiry") === 0) {
      row[col.subject] = "";
    }
    var subj = getPersonalizedSubject(row, col, i);
    subjects.push([subj]);
  }

  sheet.getRange(2, col.subject + 1, subjects.length, 1).setValues(subjects);
  notify("Updated all " + subjects.length + " high-open-rate subjects in your Google Sheet!", "Subjects Updated");
}

/**
 * =====================================================================
 * GMAIL DRAFTS CREATION ENGINE (200 Drafts for Tomorrow 9:15 AM IST)
 * Creates up to 200 personalized drafts & auto-continues if 5-min limit reached
 * =====================================================================
 */
function createDrafts200() {
  createGmailDrafts(200, false);
}

function createDrafts100() {
  createGmailDrafts(200, false);
}

function continueDrafts200Background() {
  createGmailDrafts(200, false);
}

function createGmailDrafts(count, forceRecreate) {
  var startTime = new Date().getTime();
  var MAX_DRAFT_RUNTIME_MS = 285000; // 4.75 mins safety limit so Google never kills the script
  var targetTotalDrafts = count || 200;

  var sheet = getTargetSheet();
  if (!sheet) {
    notify("Could not find Google Sheet.", "Error");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    notify("Google Sheet has no data rows.", "Empty Sheet");
    return;
  }

  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).trim());
  }

  var col = {
    status: headers.indexOf("Status"),
    name: headers.indexOf("Name"),
    univ: headers.indexOf("University"),
    dept: headers.indexOf("Department"),
    email: headers.indexOf("Email"),
    subject: headers.indexOf("Subject"),
    paperTitle: headers.indexOf("Paper Title"),
    topic: headers.indexOf("Topic"),
    hook: headers.indexOf("Hook"),
    fit: headers.indexOf("Fit"),
    highlights: headers.indexOf("Highlights"),
    body: headers.indexOf("Email Body"),
    notes: headers.indexOf("Notes")
  };

  if (col.email === -1 || col.status === -1) {
    notify("Missing Email or Status column.", "Error");
    return;
  }

  // Count how many active 'draft_created' rows already exist in the sheet
  var existingDraftRows = 0;
  if (!forceRecreate) {
    for (var rIdx = 1; rIdx < data.length; rIdx++) {
      if (String(data[rIdx][col.status]).toLowerCase().trim() === "draft_created") {
        existingDraftRows++;
      }
    }
  }

  var neededNewDrafts = Math.max(0, targetTotalDrafts - existingDraftRows);
  if (neededNewDrafts === 0 && !forceRecreate) {
    notify("All " + existingDraftRows + " Gmail drafts are already created and queued for 9:15 AM IST!", "200 Drafts Ready");
    return;
  }

  var cvBlob = null;
  try {
    var cvFile = DriveApp.getFileById(CONFIG.ATTACH_CV_DRIVE_ID);
    cvBlob = cvFile.getAs(MimeType.PDF);
  } catch (e) {
    Logger.log("Drive CV notice: " + e.message);
  }

  var defaultHighlights = "Alongside coursework, I secured AIR 134 among 13.5 lakh+ " +
    "candidates in CUET UG 2025, hold Delhi University subject ranks " +
    "(Rank 1 in one subject, Rank 2 in three), and have cleared CFA Level I.";

  var created = 0;
  var hitTimeLimit = false;

  for (var i = 1; i < data.length && created < neededNewDrafts; i++) {
    if (new Date().getTime() - startTime > MAX_DRAFT_RUNTIME_MS) {
      hitTimeLimit = true;
      break;
    }

    var row = data[i];
    var status = String(row[col.status]).toLowerCase().trim();

    if (status === "sent" || status === "replied") continue;

    if (status === "approved" || status === "ready" || status === "new" || status === "" || (forceRecreate && status === "draft_created")) {
      var toEmail = String(row[col.email]).trim();
      var profName = String(row[col.name]).trim();
      var lastName = getCleanLastName(profName);
      var univ = String(row[col.univ]).replace(/\s*\((India|Russia|UAE \/ Dubai|Australia|Global Top 100)\)\s*$/i, "").trim();
      var subject = getPersonalizedSubject(row, col, i);
      var hook = String(row[col.hook] || "");
      var fit = String(row[col.fit] || "");
      var highlights = String(row[col.highlights] || defaultHighlights);
      var deptOrTopic = String(row[col.dept] || row[col.topic] || "empirical research").trim();

      var p1 = "<p>Dear Prof. " + lastName + ",</p>";
      var p2 = "<p>I am Shubh Gupta, a second-year undergraduate in Commerce & Economics " +
        "at Shri Ram College of Commerce (SRCC), University of Delhi. I am writing to share a " +
        "question on your research at " + univ + " and to inquire about contributing as a research assistant/intern.</p>";
      var p3 = "<p>" + hook + "</p>";
      var p4 = "<p>" + fit + " " + highlights + "</p>";
      var p5 = "<p>Would you be open to a brief 10-minute conversation, or happy for me to complete a " +
        "short trial data/literature task in " + deptOrTopic + " to demonstrate how I can add value to your current projects? " +
        "I am flexible on timing (available full-time in summer 2025/2026 or part-time remotely during the academic term) " +
        "and have attached my <a href=\"" + CONFIG.DRIVE_CV_LINK + "\">CV</a> for your review.</p>";
      var p6 = "<p>Thank you very much for your time and consideration.</p>";
      var p7 = "<p>Warm regards,<br>" +
        "<strong>Shubh Gupta</strong><br>" +
        "B.Com (Hons), Shri Ram College of Commerce (SRCC), University of Delhi<br>" +
        "AIR 134 (CUET UG) | CFA Level I Passed | <a href=\"" + CONFIG.LINKEDIN_URL + "\">LinkedIn</a></p>";

      var htmlBody = [p1, p2, p3, p4, p5, p6, p7].join("\n");

      try {
        var mailOptions = { htmlBody: htmlBody, name: CONFIG.SENDER_NAME };
        if (cvBlob) mailOptions.attachments = [cvBlob];

        var draft = GmailApp.createDraft(toEmail, subject, "", mailOptions);

        sheet.getRange(i + 1, col.status + 1).setValue("draft_created");
        if (col.subject !== -1) {
          sheet.getRange(i + 1, col.subject + 1).setValue(subject);
        }
        if (col.notes !== -1) {
          sheet.getRange(i + 1, col.notes + 1).setValue("DraftID: " + draft.getId());
        }

        created++;
        Logger.log("[" + (existingDraftRows + created) + "/" + targetTotalDrafts + "] Gmail Draft Created: " + toEmail + " | " + subject);
      } catch (err) {
        Logger.log("Failed to create draft for " + toEmail + ": " + err.message);
      }
    }
  }

  var totalReadyNow = existingDraftRows + created;
  if (hitTimeLimit && totalReadyNow < targetTotalDrafts) {
    // Automatically schedule a 1-minute continuation trigger to finish the rest of the 200 drafts!
    try {
      ScriptApp.newTrigger("continueDrafts200Background")
        .timeBased()
        .after(60 * 1000)
        .create();
    } catch (e) {}
    notify(
      "Created " + totalReadyNow + "/200 drafts so far — auto-continuing the remaining " + (targetTotalDrafts - totalReadyNow) + " drafts in 60 seconds!",
      "Drafting Batch 2 Automatically"
    );
  } else {
    notify(
      "All " + totalReadyNow + " personalized drafts are ready in Gmail and scheduled for 9:15 AM IST!",
      "200 Drafts Ready"
    );
  }
}

/**
 * Approve Next 200 Emails
 */
function approveNext200() {
  approveNextBatch(200);
}

function approveNext100() {
  approveNextBatch(200);
}

function approveNextBatch(targetCount) {
  var limit = targetCount || 200;
  var sheet = getTargetSheet();
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).trim());
  }
  var statusCol = headers.indexOf("Status");
  var approvedAtCol = headers.indexOf("ApprovedAt");

  if (statusCol === -1) {
    notify("Missing 'Status' column in row 1.", "Error");
    return;
  }

  var approved = 0;
  for (var i = 1; i < data.length && approved < limit; i++) {
    var curStatus = String(data[i][statusCol]).toLowerCase().trim();
    if (curStatus === "ready" || curStatus === "new" || curStatus === "" || curStatus === "approved" || curStatus === "draft_created") {
      if (curStatus !== "draft_created") {
        sheet.getRange(i + 1, statusCol + 1).setValue("approved");
      }
      if (approvedAtCol !== -1) {
        var nowIso = new Date().toISOString();
        sheet.getRange(i + 1, approvedAtCol + 1).setValue(nowIso);
      }
      approved++;
    }
  }

  notify("Queued & approved " + approved + " emails! Scheduled for tomorrow 9:15 AM IST.", "200 Emails Queued");
}

/**
 * =====================================================================
 * AUTOMATED SENDING DISPATCHER (Sends 100 emails at 9:15 AM IST)
 * Finishes in ~4.2 mins cleanly under Google's 6-minute limit
 * =====================================================================
 */
function runDailyOutreachBatch() {
  var startTime = new Date().getTime();
  var MAX_RUNTIME_MS = 310000; // 5.1 minutes safety limit (Google kills at 6.0 mins)

  // Automatically sync any newly mined verified professors from Professor.xlsx cloud feed before sending
  try {
    syncVerifiedProfessorsFromXlsx(true);
  } catch (syncErr) {
    Logger.log("Pre-batch auto-sync notice: " + syncErr.message);
  }

  var sheet = getTargetSheet();
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).trim());
  }

  var col = {
    status: headers.indexOf("Status"),
    name: headers.indexOf("Name"),
    univ: headers.indexOf("University"),
    dept: headers.indexOf("Department"),
    email: headers.indexOf("Email"),
    subject: headers.indexOf("Subject"),
    paperTitle: headers.indexOf("Paper Title"),
    topic: headers.indexOf("Topic"),
    hook: headers.indexOf("Hook"),
    fit: headers.indexOf("Fit"),
    highlights: headers.indexOf("Highlights"),
    body: headers.indexOf("Email Body"),
    sentAt: headers.indexOf("SentAt"),
    reply: headers.indexOf("ReplyStatus"),
    notes: headers.indexOf("Notes")
  };

  if (col.email === -1 || col.status === -1) return;

  var todayStr = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd");
  var sentToday = 0;
  for (var i = 1; i < data.length; i++) {
    var sentVal = String(data[i][col.sentAt] || "");
    if (sentVal.indexOf(todayStr) !== -1) sentToday++;
  }

  var remainingToday = CONFIG.DAILY_LIMIT - sentToday;
  if (remainingToday <= 0) {
    notify("Daily quota of 100 emails already sent today.", "Daily Limit Reached");
    return;
  }

  var limitThisRun = Math.min(CONFIG.BATCH_SIZE_PER_RUN, remainingToday);

  var cvBlob = null;
  try {
    var cvFile = DriveApp.getFileById(CONFIG.ATTACH_CV_DRIVE_ID);
    cvBlob = cvFile.getAs(MimeType.PDF);
  } catch (e) {}

  var defaultHighlights = "Alongside coursework, I secured AIR 134 among 13.5 lakh+ " +
    "candidates in CUET UG 2025, hold Delhi University subject ranks " +
    "(Rank 1 in one subject, Rank 2 in three), and have cleared CFA Level I.";

  var batchCount = 0;

  for (var i = 1; i < data.length; i++) {
    if (batchCount >= limitThisRun) break;
    if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
      Logger.log("Approaching 5 min limit. Pausing batch cleanly. Sent: " + batchCount);
      break;
    }

    var row = data[i];
    var status = String(row[col.status]).toLowerCase().trim();

    // Sends any row marked as "approved" or "draft_created"
    if (status === "approved" || status === "draft_created") {
      var toEmail = String(row[col.email]).trim();
      var profName = String(row[col.name]).trim();
      var notesVal = col.notes !== -1 ? String(row[col.notes] || "") : "";
      var draftSent = false;

      // If a Gmail draft exists, send the actual draft (moves Drafts -> Sent cleanly!)
      if (status === "draft_created" && notesVal.indexOf("DraftID: ") !== -1) {
        var draftId = notesVal.replace("DraftID: ", "").trim();
        try {
          var existingDraft = GmailApp.getDraft(draftId);
          if (existingDraft) {
            existingDraft.send();
            draftSent = true;
            Logger.log("Sent existing Gmail Draft ID: " + draftId + " to " + toEmail);
          }
        } catch (e) {
          Logger.log("Draft send fallback: " + e.message);
        }
      }

      // If not sent via draft, compose and send directly
      if (!draftSent) {
        var lastName = getCleanLastName(profName);
        var univ = String(row[col.univ]).trim();
        var subject = getPersonalizedSubject(row, col);
        var hook = String(row[col.hook] || "");
        var fit = String(row[col.fit] || "");
        var highlights = String(row[col.highlights] || defaultHighlights);

        var p1 = "<p>Dear Prof. " + lastName + ",</p>";
        var p2 = "<p>I am Shubh Gupta, a second-year undergraduate in Commerce & Economics " +
          "at Shri Ram College of Commerce (SRCC), University of Delhi. I am writing to ask " +
          "about research internship opportunities with you at " + univ + ".</p>";
        var p3 = "<p>" + hook + "</p>";
        var p4 = "<p>" + fit + " " + highlights + "</p>";
        var p5 = "<p>Would you be open to a brief conversation about whether you take research " +
          "interns, or whether there is a project where a newcomer could add value? I am " +
          "flexible on timing and can work in-person in summer 2025/2026, or remotely on a " +
          "part-time basis during the academic term. I have attached my " +
          "<a href=\"" + CONFIG.DRIVE_CV_LINK + "\">CV</a> for reference.</p>";
        var p6 = "<p>Thank you for taking the time to read this.</p>";
        var p7 = "<p>Best regards,<br>" +
          "<strong>Shubh Gupta</strong><br>" +
          "B.Com (Hons), Shri Ram College of Commerce, University of Delhi<br>" +
          "<a href=\"" + CONFIG.LINKEDIN_URL + "\">LinkedIn</a></p>";

        var htmlBody = [p1, p2, p3, p4, p5, p6, p7].join("\n");

        try {
          var mailOptions = { htmlBody: htmlBody, name: CONFIG.SENDER_NAME };
          if (cvBlob) mailOptions.attachments = [cvBlob];

          GmailApp.sendEmail(toEmail, subject, "", mailOptions);
          draftSent = true;
          if (col.subject !== -1 && String(row[col.subject] || "").trim() !== subject) {
            sheet.getRange(i + 1, col.subject + 1).setValue(subject);
          }
        } catch (err) {
          Logger.log("Failed to send to " + toEmail + ": " + err.message);
          sheet.getRange(i + 1, col.status + 1).setValue("failed");
          if (col.notes !== -1) {
            sheet.getRange(i + 1, col.notes + 1).setValue("Error: " + err.message);
          }
        }
      }

      if (draftSent) {
        sheet.getRange(i + 1, col.status + 1).setValue("sent");
        if (col.sentAt !== -1) {
          var sentTimeIso = new Date().toISOString();
          sheet.getRange(i + 1, col.sentAt + 1).setValue(sentTimeIso);
        }
        if (col.subject !== -1 && String(row[col.subject] || "").trim() !== subject) {
          sheet.getRange(i + 1, col.subject + 1).setValue(subject);
        }

        batchCount++;
        Logger.log("[" + batchCount + "] Sent to: " + toEmail + " (" + profName + ")");

        // 2 to 3 second natural spacing between emails
        var pauseSec = CONFIG.GAP_SECONDS_MIN + Math.random() * (CONFIG.GAP_SECONDS_MAX - CONFIG.GAP_SECONDS_MIN);
        Utilities.sleep(Math.round(pauseSec * 1000));
      }
    }
  }

  var totalSentTodayNow = sentToday + batchCount;
  if (totalSentTodayNow < CONFIG.DAILY_LIMIT && batchCount > 0) {
    // Automatically schedule Wave 2 in 90 seconds to dispatch the remaining 100 emails (up to 200/day)!
    try {
      ScriptApp.newTrigger("runDailyOutreachBatch")
        .timeBased()
        .after(90 * 1000)
        .create();
      Logger.log("Scheduled Wave 2 continuation in 90 seconds for remaining " + (CONFIG.DAILY_LIMIT - totalSentTodayNow) + " emails.");
    } catch (e) {
      Logger.log("Wave 2 trigger notice: " + e.message);
    }
  }

  notify("Completed Wave! Sent " + totalSentTodayNow + "/" + CONFIG.DAILY_LIMIT + " emails today.", "Outreach Dispatched");
}

/**
 * =====================================================================
 * SMART NON-DESTRUCTIVE SYNC FROM PROFESSOR.XLSX CLOUD REPO
 * Preserves all 'sent' and 'replied' rows, removes old un-sent synthetic
 * emails, and appends all newly mined verified professors from Professor.xlsx!
 * =====================================================================
 */
function syncVerifiedProfessorsFromXlsx(silentMode) {
  if (!silentMode) {
    notify("Syncing latest verified professors from Professor.xlsx cloud feed...", "Syncing Professor.xlsx");
  }
  try {
    var cacheBuster = "?t=" + new Date().getTime();
    var response = UrlFetchApp.fetch(CONFIG.DATA_URL_VERIFIED + cacheBuster);
    var csvText = response.getContentText("UTF-8");
    var cloudRows = Utilities.parseCsv(csvText);

    if (!cloudRows || cloudRows.length <= 1) {
      if (!silentMode) notify("Cloud feed was empty.", "Error");
      return;
    }

    var sheet = getTargetSheet();
    if (!sheet) return;

    var existingData = sheet.getDataRange().getValues();
    var cloudHeaders = [];
    for (var ch = 0; ch < cloudRows[0].length; ch++) {
      cloudHeaders.push(String(cloudRows[0][ch]).trim());
    }

    // If sheet is empty, populate directly
    if (!existingData || existingData.length <= 1) {
      sheet.clear();
      sheet.getRange(1, 1, cloudRows.length, cloudRows[0].length).setValues(cloudRows);
      sheet.getRange(1, 1, 1, cloudRows[0].length).setFontWeight("bold").setBackground("#e8eaed");
      sheet.setFrozenRows(1);
      if (!silentMode) {
        notify("Loaded " + (cloudRows.length - 1) + " verified professors from Professor.xlsx!", "Sync Complete");
      }
      return;
    }

    var sheetHeaders = [];
    for (var sh = 0; sh < existingData[0].length; sh++) {
      sheetHeaders.push(String(existingData[0][sh]).trim());
    }

    var sEmailCol = sheetHeaders.indexOf("Email");
    var sStatusCol = sheetHeaders.indexOf("Status");
    var sNotesCol = sheetHeaders.indexOf("Notes");
    var cEmailCol = cloudHeaders.indexOf("Email");
    var cStatusCol = cloudHeaders.indexOf("Status");
    var cNotesCol = cloudHeaders.indexOf("Notes");
    var cIdCol = cloudHeaders.indexOf("ID");

    var preservedRows = [];
    var seenEmails = {};
    var draftStateByEmail = {};

    // 1. Keep all rows that have already been SENT, REPLIED, or FOLLOW_UP_DUE so history is never lost,
    //    and remember any existing DraftID for rows with status 'draft_created'
    if (sEmailCol !== -1 && sStatusCol !== -1) {
      for (var i = 1; i < existingData.length; i++) {
        var r = existingData[i];
        var st = String(r[sStatusCol] || "").toLowerCase().trim();
        var em = String(r[sEmailCol] || "").toLowerCase().trim();
        if (!em) continue;

        if (st === "sent" || st === "replied" || st === "follow_up_due") {
          var normRow = [];
          for (var k = 0; k < cloudHeaders.length; k++) {
            normRow.push(k < r.length ? r[k] : "");
          }
          preservedRows.push(normRow);
          seenEmails[em] = true;
        } else if (st === "draft_created") {
          draftStateByEmail[em] = {
            status: "draft_created",
            notes: sNotesCol !== -1 ? String(r[sNotesCol] || "") : ""
          };
        }
      }
    }

    // 2. Append all verified professors from Professor.xlsx that haven't been emailed yet
    var newlySyncedCount = 0;
    var approvedCount = 0;
    for (var j = 1; j < cloudRows.length; j++) {
      var cRow = cloudRows[j];
      if (!cRow || cRow.length <= cEmailCol) continue;
      var cEmail = String(cRow[cEmailCol] || "").toLowerCase().trim();
      if (!cEmail || seenEmails[cEmail]) continue;

      seenEmails[cEmail] = true;
      var outRow = [];
      for (var m = 0; m < cloudHeaders.length; m++) {
        outRow.push(m < cRow.length ? cRow[m] : "");
      }

      // If a Gmail draft was already created for this professor, preserve 'draft_created' & DraftID
      if (draftStateByEmail[cEmail]) {
        outRow[cStatusCol] = "draft_created";
        if (cNotesCol !== -1 && draftStateByEmail[cEmail].notes) {
          outRow[cNotesCol] = draftStateByEmail[cEmail].notes;
        }
        approvedCount++;
      } else {
        // Ensure at least 200 unsent rows are marked 'approved' so the 9:15 AM IST scheduler always has a full 200-email batch ready
        var curSt = String(outRow[cStatusCol] || "").toLowerCase().trim();
        if (approvedCount < 200 && (curSt === "approved" || curSt === "ready" || curSt === "new" || curSt === "")) {
          outRow[cStatusCol] = "approved";
          approvedCount++;
        }
      }

      preservedRows.push(outRow);
      newlySyncedCount++;
    }

    // 3. Renumber ID column sequentially
    var finalGrid = [cloudHeaders];
    for (var idx = 0; idx < preservedRows.length; idx++) {
      if (cIdCol !== -1) {
        preservedRows[idx][cIdCol] = idx + 1;
      }
      finalGrid.push(preservedRows[idx]);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, finalGrid.length, cloudHeaders.length).setValues(finalGrid);
    sheet.getRange(1, 1, 1, cloudHeaders.length).setFontWeight("bold").setBackground("#e8eaed");
    sheet.setFrozenRows(1);

    if (!silentMode) {
      notify(
        "Synced " + newlySyncedCount + " verified professors from Professor.xlsx (Kept " + (preservedRows.length - newlySyncedCount) + " sent/replied rows)!",
        "Professor.xlsx Synced"
      );
    }
  } catch (err) {
    Logger.log("syncVerifiedProfessorsFromXlsx error: " + err.message);
    if (!silentMode) {
      notify("Sync error: " + err.message, "Error");
    }
  }
}

/**
 * AUTO-POPULATE: Verified Professor.xlsx Dataset
 */
function autoPopulateCuratedProfessors() {
  fetchAndPopulateSheet(CONFIG.DATA_URL_VERIFIED, "Verified Professor.xlsx Dataset");
}

/**
 * AUTO-POPULATE: Full Verified Batch
 */
function autoPopulateBatch1Professors() {
  fetchAndPopulateSheet(CONFIG.DATA_URL_BATCH1, "Verified Professor.xlsx Batch");
}

function fetchAndPopulateSheet(csvUrl, label) {
  notify("Downloading " + label + " from cloud...", "Populating Sheet");
  try {
    var cacheBuster = "?t=" + new Date().getTime();
    var response = UrlFetchApp.fetch(csvUrl + cacheBuster);
    var csvText = response.getContentText("UTF-8");
    var csvData = Utilities.parseCsv(csvText);

    if (!csvData || csvData.length === 0) {
      notify("Downloaded file was empty.", "Error");
      return;
    }

    var sheet = getTargetSheet();
    if (!sheet) return;

    sheet.clear();
    sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    sheet.getRange(1, 1, 1, csvData[0].length).setFontWeight("bold").setBackground("#e8eaed");
    sheet.setFrozenRows(1);

    notify("Loaded " + (csvData.length - 1) + " verified professors into Google Sheet!", "Success");
  } catch (err) {
    notify("Error loading data: " + err.message, "Error");
    Logger.log("fetchAndPopulateSheet error: " + err.message);
  }
}

/**
 * Install 9:15 AM IST Daily Scheduler (200 Emails/Day) + 2-Hour Live Professor.xlsx Auto-Sync
 */
function installMorningTrigger() {
  removeMorningTrigger();

  // Approve 200 rows automatically so tomorrow's 9:15 AM batch is 100% ready
  try {
    approveNextBatch(200);
  } catch (e) {}

  ScriptApp.newTrigger("runDailyOutreachBatch")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(15)
    .create();

  ScriptApp.newTrigger("syncRepliesAndFollowups")
    .timeBased()
    .everyHours(2)
    .create();

  notify(
    "Scheduled 200 Mails for Tomorrow 9:15 AM IST (Wave 1 + Auto Wave 2) + 2-Hour Professor.xlsx Sync!",
    "200 Mails Scheduled (9:15 AM)"
  );
}

function removeMorningTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  notify("All automated triggers removed.", "Scheduler Cleared");
}

/**
 * Scan Gmail for Replies & Follow-ups + Auto-Sync New Mined Professors from Professor.xlsx
 */
function syncRepliesAndFollowups() {
  // First pull any new verified professors mined into Professor.xlsx
  try {
    syncVerifiedProfessorsFromXlsx(true);
  } catch (e) {
    Logger.log("Auto-sync notice: " + e.message);
  }

  var sheet = getTargetSheet();
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = [];
  for (var h = 0; h < data[0].length; h++) {
    headers.push(String(data[0][h]).trim());
  }

  var emailCol = headers.indexOf("Email");
  var statusCol = headers.indexOf("Status");
  var sentAtCol = headers.indexOf("SentAt");
  var replyCol = headers.indexOf("ReplyStatus");

  var now = new Date().getTime();
  var replies = 0;
  var followups = 0;

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][statusCol]).toLowerCase().trim();
    var email = String(data[i][emailCol]).trim();
    var sentAtVal = data[i][sentAtCol];

    if (status === "sent" && email) {
      var threads = GmailApp.search("from:" + email);
      if (threads && threads.length > 0) {
        var lastMsgDate = threads[0].getLastMessageDate();
        sheet.getRange(i + 1, statusCol + 1).setValue("replied");
        if (replyCol !== -1) {
          var dateStr = Utilities.formatDate(lastMsgDate, "Asia/Kolkata", "yyyy-MM-dd HH:mm");
          sheet.getRange(i + 1, replyCol + 1).setValue("Replied on: " + dateStr);
        }
        replies++;
        continue;
      }

      if (sentAtVal) {
        var sentTime = new Date(sentAtVal).getTime();
        var daysElapsed = (now - sentTime) / (1000 * 60 * 60 * 24);
        if (daysElapsed >= CONFIG.FOLLOW_UP_DAYS) {
          sheet.getRange(i + 1, statusCol + 1).setValue("follow_up_due");
          followups++;
        }
      }
    }
  }

  notify("Replies: " + replies + ", Follow-ups due: " + followups, "Sync Complete");
}

/**
 * Test Send 1 Email to Self
 */
function testSendToSelf() {
  var userEmail = Session.getActiveUser().getEmail() || CONFIG.SENDER_EMAIL;
  var cvBlob = null;
  try {
    var cvFile = DriveApp.getFileById(CONFIG.ATTACH_CV_DRIVE_ID);
    cvBlob = cvFile.getAs(MimeType.PDF);
  } catch (e) {}

  var testSubject = "Research Internship Inquiry - Preview / Shubh Gupta (SRCC)";
  var tp1 = "<p>Dear Prof. Sample,</p>";
  var tp2 = "<p>I am Shubh Gupta, a second-year undergraduate in Commerce & Economics at " +
    "Shri Ram College of Commerce (SRCC), University of Delhi. I am writing to ask about " +
    "research internship opportunities with you at Sample University.</p>";
  var tp3 = "<p>This is a verified test email sent from your Google Sheet outreach engine.</p>";
  var tp4 = "<p>Alongside coursework, I secured AIR 134 among 13.5 lakh+ candidates in " +
    "CUET UG 2025, hold Delhi University subject ranks (Rank 1 in one subject, Rank 2 in three), " +
    "and have cleared CFA Level I.</p>";
  var tp5 = "<p>I have attached my <a href=\"" + CONFIG.DRIVE_CV_LINK + "\">CV</a> for reference.</p>";
  var tp6 = "<p>Best regards,<br><strong>Shubh Gupta</strong><br>" +
    "B.Com (Hons), Shri Ram College of Commerce, University of Delhi<br>" +
    "<a href=\"" + CONFIG.LINKEDIN_URL + "\">LinkedIn</a></p>";

  var testHtml = [tp1, tp2, tp3, tp4, tp5, tp6].join("\n");
  var opts = { htmlBody: testHtml, name: CONFIG.SENDER_NAME };
  if (cvBlob) opts.attachments = [cvBlob];

  GmailApp.sendEmail(userEmail, testSubject, "", opts);
  notify("Test email sent to " + userEmail + "! Check your inbox.", "Test Email Sent");
}
