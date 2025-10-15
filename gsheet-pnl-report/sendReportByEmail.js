function sendReportByEmail({expenseType, year, periods, amounts}) {
  const recipientEmail = Session.getActiveUser().getEmail(); // Sends to the current user's email
  const subject = `Expense Report: ${expenseType} - ${year}`;

  const htmlBody = buildExpenseReportEmailBody(expenseType, year, periods, amounts);

  try {
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    Logger.log(`Expense report for ${expenseType} sent to ${recipientEmail}`);
    SpreadsheetApp.getUi().alert('Success', `Expense report for ${expenseType} sent to ${recipientEmail}`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log(`Error sending email: ${e.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to send email: ${e.message}. Please check your email settings or permissions.`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
