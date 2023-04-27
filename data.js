
function getEmailData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var threads = GmailApp.search('from:shubhamtribhuvan7262@gmail.com'); // Change this to your search criteria
  var data = [];
  
  // Adding the title row
  var titleRow = ["S. No.", "RFQ Number", "Supplier Name", "Subject", "Email Sent to", "Email Body", "Response (Yes/No)", "Attachments (Yes/No)", "Attachment Link", "Mail Body", "Updated on"];
  data.push(titleRow);
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      
      // Extracting RFQ number from the email subject
      var subject = message.getSubject();
      // var rfqNumber = subject.match(/\d{6}/);
      var rfqNumber = subject.match(/\d+/);
      // Extracting Supplier Name from the email body
      var body = message.getPlainBody();
      // var supplierName = body.match(/Supplier Name: (.*)/)[1];

      var supplierNameMatch = body.match(/Supplier Name:\s*(.*)\s*Subject:/);
      var supplierName = supplierNameMatch ? supplierNameMatch[1] : "";

      
      // Adding the extracted data to a row
      var row = [
        i + 1, // S. No.
        rfqNumber, // RFQ Number
        supplierName, // Supplier Name
        subject, // Subject
        message.getTo(), // Email Sent to
        body, // Email Body
        "", // Response (Yes/No)
        "", // Attachments (Yes/No)
        "", // Attachment Link
        "", // Mail Body
        "" // Updated on
      ];
      data.push(row);
    }
  }
  
  // Adding the data to the spreadsheet
  sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}
