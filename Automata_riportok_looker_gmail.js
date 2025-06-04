// Script, ami egy sheet alapján automatizálja a riportkiküldést.
// Ne felejtsd el időzíteni a script lefutását dátum alapján bal oldalon a Triggers menüben.

// A helyes működéshez be kell állítani a Looker Studio-ban minden automatizálni kívánt riportnál egy "Schedule Delivery-t" a "Share" menüpontban.
// A címzett te magad legyél, a levél tárgya a riport címe lesz. Ezt fogja figyelni a script.

function sendMonthlyReports() {
  // Ide írd be a Google Sheet azonosítóját, amit figyel majd a rendszer
  // Ennek a sheet-nek 3 oszlopa kell, hogy legyen: subject, email, message
  // A subject tartalmazza, hogy mi a neve a riportnak, amit kapsz looker-ből - nem kell pontosan megadni, de figyelj, nehogy duplikált legyen
  // Az email tartalmazza, hogy kinek legyen kiküldve a riport
  // A message tartalmazza, hogy mi legyen az emailbe leírva. A PDF automatikusan csatolva lesz hozzá.
  const sheetId = "1FGZRJvgIf44Q68eqtb7dAwu5ktFJk6Z1IvFxUVcSq1Q";

  const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
  const data = sheet.getDataRange().getValues();

  //*****
  // Csinálj egy ilyen label-t Gmail-ben (Automata riportok). Csak az ide áthúzott automata Looker Studio-s emaileket scanneli be a script, ezzel biztosítva, hogy átnézd, kinek mit küldesz majd el, amikor lefut a script.
  const label = GmailApp.getUserLabelByName("Automata riportok");
  //*****

  const threads = label.getThreads();

  const configList = [];
  for (let i = 1; i < data.length; i++) {
    const [subjectPart, emailCell, message] = data[i];
    if (subjectPart && emailCell && message) {
      const emails = emailCell.toString().split(/\r?\n/).map(e => e.trim()).filter(e => e);
      configList.push({ subjectPart, emails, message });
    }
  }

  threads.forEach(thread => {
    const message = thread.getMessages()[0]; // Only process the first message in each thread
    const originalSubject = message.getSubject();
    const config = configList.find(cfg => originalSubject.includes(cfg.subjectPart));
    if (config) {
      const attachments = message.getAttachments().filter(att => att.getContentType() === "application/pdf");
      if (attachments.length > 0) {
        const modifiedSubject = originalSubject.includes(" - ") ? originalSubject.split(" - ").slice(0, -1).join(" - ") : originalSubject;
        MailApp.sendEmail({
          to: config.emails.join(','),
          subject: modifiedSubject,
          body: config.message,
          attachments: attachments
        });
        Logger.log(`Email sent -> Subject: "${modifiedSubject}", To: ${config.emails.join(', ')}, Message: "${config.message}"`);
      }
    }
  });
}
