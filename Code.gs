function extractEmailsToSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var label = GmailApp.getUserLabelByName("hammad@internee.pk"); // Change label to an actual one

  // Check if label exists
  if (!label) {
    Logger.log("Label not found. Please create the label in Gmail.");
    return;
  }

  var threads = label.getThreads();
  if (threads.length === 0) {
    Logger.log("No emails found under this label.");
    return;
  }

  // Set headers if the sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "From", "Subject", "Snippet", "Attachment Links"]);
  }

  var folder = DriveApp.getFolderById("1bHXdHSZZUmgG0MXtyJh99mWedeEB_6aK"); // Change to your Drive folder ID

  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var attachments = message.getAttachments();
      var attachmentLinks = [];

      for (var k = 0; k < attachments.length; k++) {
        var file = folder.createFile(attachments[k]);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        attachmentLinks.push(file.getUrl());
      }

      // Append email data to Google Sheet
      sheet.appendRow([
        message.getDate(),
        message.getFrom(),
        message.getSubject(),
        attachmentLinks.join(", ") // Combine links if multiple attachments
      ]);
    }
  }
}
