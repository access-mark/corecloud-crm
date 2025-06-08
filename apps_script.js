function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CRM Tools')
    .addItem('Send Follow-Up Emails', 'sendFollowUpEmails')
    .addItem('Update Pipeline Counts', 'updatePipeline')
    .addItem('Import LinkedIn Contacts', 'importFromLinkedInUpload')
    .addToUi();
}

function sendFollowUpEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lead Tracker");
  const data = sheet.getDataRange().getValues();
  const today = new Date();
  let sentCount = 0;

  for (let i = 1; i < data.length; i++) {
    const followUpDate = new Date(data[i][7]); // Column H: Next Follow-Up
    if (followUpDate.toDateString() === today.toDateString()) {
      const email = data[i][2]; // Column C: Email
      const name = data[i][0];  // Column A: Lead Name
      const company = data[i][1];

      MailApp.sendEmail({
        to: "mark@corecloudza.com", // Replace or personalize
        subject: `Follow-Up Reminder: ${name} @ ${company}`,
        body: `Reminder to follow up with ${name} from ${company} today.`
      });

      sentCount++;
    }
  }

  SpreadsheetApp.getUi().alert(`✅ ${sentCount} reminder(s) sent.`);
}

function updatePipeline() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lead Tracker");
  const pipelineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pipeline");
  const data = sheet.getDataRange().getValues();

  const stages = ["Prospect", "Qualified", "Proposal Sent", "Negotiation", "Closed Won", "Closed Lost"];
  const counts = Array(stages.length).fill(0);

  for (let i = 1; i < data.length; i++) {
    const status = data[i][4];
    const index = stages.indexOf(status);
    if (index !== -1) counts[index]++;
  }

  for (let i = 0; i < counts.length; i++) {
    pipelineSheet.getRange(i + 2, 2).setValue(counts[i]);
  }

  SpreadsheetApp.getUi().alert("✅ Pipeline updated.");
}

function importFromLinkedInUpload() {
  const uploadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LinkedIn Upload");
  const leadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lead Tracker");
  const uploadData = uploadSheet.getDataRange().getValues();

  for (let i = 1; i < uploadData.length; i++) {
    const row = uploadData[i];
    if (row[0] && row[1]) {
      leadSheet.appendRow([
        `${row[0]} ${row[1]}`,  // Full Name
        row[2],                 // Company
        row[3],                 // Email
        row[4],                 // Phone
        "Prospect",             // Default Status
        "LinkedIn",             // Lead Source
        "",                     // Last Contacted
        ""                      // Next Follow-Up
      ]);
    }
  }

  SpreadsheetApp.getUi().alert("✅ LinkedIn contacts imported.");
}
