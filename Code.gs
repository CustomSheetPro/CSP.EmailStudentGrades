function EmailGrades() {
  sendGradesToAllStudents();
}

function sendGradesToAllStudents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var dataValues = dataRange.getValues();
  var headers = dataValues[0];

  for (var i = 1; i < dataValues.length; i++) {
    var rowData = dataValues[i];
    var dictionary = {};

    for (var j = 0; j < headers.length; j++) {
      dictionary[headers[j]] = rowData[j];
    }
    
    sendGradesToStudentViaEmail(dictionary);
  }
}

function sendGradesToStudentViaEmail(dictionary) {
  var recipient = dictionary["Email"];
  var subject = `${dictionary["Name"]} Grades`;
  var sender = Session.getActiveUser().getEmail();

  var body = `
    <p>Dear ${dictionary["Name"]},</p>
    <p>Here are your grades:</p>
    <ul>
      <li>Math: ${dictionary["Math"]}</li>
      <li>History: ${dictionary["History"]}</li>
      <li>Biology: ${dictionary["Biology"]}</li>
      <li>English: ${dictionary["English"]}</li>
    </ul>
    <p>Regards,<br>Mr. Feeny</p>
  `;

  MailApp.sendEmail({
    to: recipient,
    replyTo: sender,
    subject: subject,
    htmlBody: body
  });
}
