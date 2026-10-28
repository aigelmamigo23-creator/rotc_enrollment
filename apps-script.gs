// Paste this into Extensions > Apps Script inside a Google Sheet.
// It saves each enrollment as a new row: Timestamp | Name | Section
// and emails you every time someone enrolls.

// Email that gets the notifications. Leave "" to use your own Google account email.
var NOTIFY_EMAIL = "";

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var name = String(data.name || "").replace(/\s+/g, " ").trim().slice(0, 80);
    var section = String(data.section || "").replace(/\s+/g, " ").trim().slice(0, 40);
    if (!name || !section) return reply({ ok: false, error: "missing" });

    var sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getLastRow() === 0) sheet.appendRow(["Timestamp", "Name", "Section"]);

    if (sheet.getLastRow() > 1) {
      var rows = sheet.getRange(2, 2, sheet.getLastRow() - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).toLowerCase() === name.toLowerCase() &&
            String(rows[i][1]).toLowerCase() === section.toLowerCase()) {
          return reply({ ok: false, error: "duplicate" });
        }
      }
    }

    sheet.appendRow([new Date(), name, section]);
    var total = sheet.getLastRow() - 1;

    // Email notification. A failed email never blocks the enrollment.
    try {
      var to = NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
      MailApp.sendEmail(to, "New ROTC enrollee: " + name,
        name + " (" + section + ") just enrolled.\nTotal enrolled: " + total);
    } catch (mailErr) {}

    return reply({ ok: true, total: total });
  } catch (err) {
    return reply({ ok: false, error: "server" });
  } finally {
    lock.releaseLock();
  }
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
