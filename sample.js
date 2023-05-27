function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Process RFQ")
    .addItem("Track Sent Email", "getEmailData")
    .addToUi();
}

function getEmailData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // var threads = GmailApp.search('in:sent newer_than:1d subject:"RFQ"');//not working
  // var threads = GmailApp.search('in:sent:2d subject:"RFQ"');//Not working
  var threads = GmailApp.search('subject:"RFQ"');//working cool


  var data = [];
  console.log(data)
  // Adding the title row
  var titleRow = ["S.No.", "RFQ Number", "Supplier Name", "Subject", "Email Sent to", "Email Body", "Response (Yes/No)",  "Attachments (Yes/No)", "Attachment Link", "Mail Body", "Updated on"];
  data.push(titleRow);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var subject = message.getSubject();
      var body = message.getPlainBody();
      var emailSentTo = message.getTo();
      Logger.log('Processing email with subject: ' + subject);
      
      // Extracting RFQ number and Supplier Name from the email subject
      var subjectRegex = /^RFQ\s+(\d+)\s+(.*?)\s+(\w{3}\s+\d{4})/;
      var match = subject.match(subjectRegex);
      if (!match) {
        Logger.log('No match found for subject: ' + subject);
        continue;
      }
      var rfqNumber = match[1];
      var supplierName = match[2];

      // Extracting email body
      var emailBody = body.replace(/^(Supplier Name|Subject):\s+.*/gm, "");

      // Extracting attachments
      var attachments = message.getAttachments();
      var attachmentLinks = [];
      var attachmentsExist = attachments.length > 0;

      var row = [
        i + 1,
        rfqNumber, 
        supplierName, 
        subject,
        emailSentTo,
        emailBody, 
        "", 
        attachmentsExist ? "Yes" : "No", 
        "", 
        "", 
        "" 
      ];
      data.push(row);
    }
  }
  
  // Adding the data to the spreadsheet

  Logger.log('Adding ' + data.length + ' rows to the sheet');
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}
