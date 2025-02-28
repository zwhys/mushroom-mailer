var ccEmail = "REDACTED";
var overallIC = "REDACTED";
var overallICName = "REDACTED";
var password = "REDACTED";

function sendReminderEmails() {
  const sheet = SpreadsheetApp.getActive();
  const data = sheet.getDataRange().getValues();

  // Get the current date and calculate tomorrow's date
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);

  let loanDates = {};

  // Loop through each row of the data from the sheet
  data.forEach((row, index) => {
    if (index === 0) return; // Skip the header row

    // Extract relevant information from the row (e.g., loan date, recipient email, etc.)
    const recipientEmail = row[1];
    const recipientName = row[2];
    const gamesLoaned = row[5];
    const loanDate = row[6]; //letter corresponds to number - 1 (ie. A = 0, B = 1, C = 2)
    const additionalControllers = row[7];
    const emailSentStatus = row[8];

    const currentYear = new Date().getFullYear();
    loanDate.setFullYear(currentYear);

    // Check if the loan date is a valid date and the email hasn't been sent already
    if (emailSentStatus !== "Sent" && areDatesEqual(tomorrow, loanDate)) {
      if (!loanDates[loanDate]) {
        // Check for loan date conflicts (same date, different email)
        loanDates[loanDate] = recipientEmail;
        sendEmailForLoan(
          recipientEmail,
          recipientName,
          recipientEmail,
          loanDate,
          gamesLoaned,
          additionalControllers
        );
      } else if (loanDates[loanDate] !== recipientEmail) {
        sendEmailForRejection(recipientEmail, loanDate);
      }
      const cellReference = "I" + (index + 1);
      sheet.getRange(cellReference).setValue("Sent");
    }
  });
  Logger.log(`Execution completed for ${formatLoanDate(tomorrow)} `);
}

function sendEmailForLoan(
  recipientEmail,
  recipientName,
  recipientEmail,
  loanDate,
  gamesLoaned,
  additionalControllers
) {
  const subject = `REDACTED`;
  const body = `
   REDACTED
  `;

  // Get the PDF file from Google Drive using its file ID
  const pdfFileId = "REDACTED";
  const pdfFile = DriveApp.getFileById(pdfFileId);

  // Send the email with the loan reminder and the attached PDF
  MailApp.sendEmail({
    to: recipientEmail,
    cc: ccEmail,
    subject: subject,
    htmlBody: body,
    attachments: [pdfFile.getAs(MimeType.PDF)], // Attach the PDF file
  });

  // Log the action for reference
  Logger.log(
    `Email sent to ${recipientEmail} for loan on ${loanDate.toDateString()}`
  );
}

// This function sends an email for the loan conflict
function sendEmailForRejection(recipientEmail, loanDate) {
  const subject = `REDACTED`;
  const body = `
   REDACTED
  `;

  MailApp.sendEmail({
    to: recipientEmail,
    cc: ccEmail,
    subject: subject,
    htmlBody: body,
  });

  Logger.log(
    `Loan rejection email sent to ${recipientEmail} for ${loanDate.toDateString()}`
  );
}

// This function formats the loan date in a human-readable format (e.g., "15 February")
function formatLoanDate(loanDate) {
  const day = loanDate.getDate();
  const month = loanDate.toLocaleString("default", { month: "long" }); // Get full month name
  return `${day} ${month}`; // Return the formatted date
}

function areDatesEqual(tomorrow, loanDate) {
  const d1 = new Date(tomorrow);
  const d2 = new Date(loanDate);

  d1.setHours(0, 0, 0, 0);
  d2.setHours(0, 0, 0, 0);

  return d1.getTime() === d2.getTime();
}
