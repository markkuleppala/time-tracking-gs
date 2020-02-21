function onOpen() { // Add the script to spreadsheet UI
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Time tracking')
      .addItem('Send report to individuals', 'sendEmails')
      .addToUi();
}

function toColumns(values) { // Array data to JSON
  var current;
  var header = values[0];
  var ans = [];

  for (var i = 1; i < values.length; i++) {
    current = {};
    for (var j = 0; j < header.length; j++) {
      current[header[j]] = values[i][j];
    }
    ans.push(current);
  }

  return ans;
}

function getVariables() {
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt('Enter spreadsheet ID e.g. 1ho4K88TTo9EXcLjZA2_DSFGio1ThVQ5-XqrPJQX4p5Y').getResponseText();
  var month = ui.prompt('Enter spreadsheet tab name e.g. March').getResponseText();
  var PREVIOUS_BALANCE_COLUMN = ui.prompt('Enter column name for previous balance ID e.g. February 2020 balance').getResponseText();
  return [id, month, PREVIOUS_BALANCE_COLUMN];
}

function replaceCommasDots(id) { // Replace commas with dots in the given sheet
  var sheet = SpreadsheetApp.openById(id);
  var range = sheet.getRange("B2:XX999"); // Range to modify
  var data  = range.getValues();

  for (var row = 0; row < data.length; row++) {
    for (var col = 0; col < data[row].length; col++) {
      data[row][col] = (data[row][col]).toString().replace(/,/g, '.');
    }
  }
  range.setValues(data);
}

function sendEmails() {
  var userVariables = getVariables(); // Call for getVariables to ask spreadsheet variables from the user
  var id = userVariables[0]; var month = userVariables[1]; var PREVIOUS_BALANCE_COLUMN = userVariables[2];
  replaceCommasDots(id); // Replace commas with dots in the spreadsheet answers
  
  var EMAIL_SENT = "EMAIL_SENT";
  var EMAIL_COLUMN = "Email Address";
  var BALANCE_COLUMN = "BALANCE";
  
  var aliases = GmailApp.getAliases(); // Get aliases for sending the emails
  var sheetActive = SpreadsheetApp.openById(id); // Open spreadsheet
  var sheet = sheetActive.getSheetByName(month); // Open sheet
  
  if (!sheet) {
    ui.alert("Sheet not found: " + month);
    return;
  }
  
  var startRow = 1;
  var rawData = sheet.getDataRange().getValues(); // Get all data in the rows as array
  var timeData = toColumns(rawData); // Transfer the data to JSON
  var headersData = rawData[0]; // Get headers as array
  
  if (timeData[0][EMAIL_COLUMN] === undefined) {
    throw "Column '" + EMAIL_COLUMN + "' not found on sheet '" + month + "'.";
  }
  
  if (timeData[0][PREVIOUS_BALANCE_COLUMN] === undefined) {
    throw "Column '" + PREVIOUS_BALANCE_COLUMN + "' not found on sheet '" + month + "'.";
  }
  
  if (timeData[0][BALANCE_COLUMN] === undefined) {
    throw "Column '" + BALANCE_COLUMN + "' not found on sheet '" + month + "'.";
  }
  
  if (timeData[0][EMAIL_SENT] === undefined) {
    throw "Column '" + EMAIL_SENT + "' not found on sheet '" + month + "'.";
  }
  
  if (timeData.length === 0) {
    throw "No time tracking data found.";
  }
  
  var timeDataFiltered = timeData.filter(function(e) { // Filtered data as JSON, takes out rows without email or with EMAIL_SENT as EMAIL_SENT
    return (e[EMAIL_COLUMN] !== "");
  });
  
  var days = []; var offset = 4; // Non-day columns in the beginning
  for (var i = offset; i < headersData.length-offset; i++) { // Get day columns as array
    days[i-offset] = headersData[i];
  }
  
  
  var subject = "Time tracking hours - " + month + " 2020"; // Create subject
  var replyTo = "operations@smartly.io"; // Email reply to and from field
  
  var messageFirst = "Here are your time tracking markings for " + month + " 2020.<br/><br/>"; // Message in the email before table
  
  // Loop through all rows and create the email
  for (var i = 0; i < timeDataFiltered.length; i++) { // Create email for all answers
    if (timeDataFiltered[i][EMAIL_SENT] == "EMAIL_SENT") {continue;} // Skip if email has been sent already
    var email = timeDataFiltered[i][EMAIL_COLUMN]; // Email address
    var previous_balance = timeDataFiltered[i][PREVIOUS_BALANCE_COLUMN]; // Balance of the previous month
    
    
    var balance_sum = 0; // Monthly balance sum
    var day_text; var day_value; var messageSecond = '<table border="1"><tr><th>Day</th><th>Balance</th></tr>';
   
    
    for (var j = 0; j < days.length; j++) { // Create table and calculate monthly balance
      day_text = days[j]; // Day as text
      day_value = timeDataFiltered[i][day_text]; // Day's value
      messageSecond += '<tr><td>' + day_text + '</td><td style="text-align:center">' + day_value + '</td></tr>'; // Add to table
      if (!day_value) { // If no value added (null)
        day_value = 0;
      }
      balance_sum += day_value;
    }
    
    messageSecond += '<tr><th>Sum</th><th>' + balance_sum + '</th></tr></table>' // Close the table
    var balance = balance_sum + previous_balance; // Total balance
    
    var messageThird = "<br/><br/>The previous month you had " + previous_balance + " hours. In total, you have <b>" + balance + "</b> (previous + current month) hours to use.<br/>";
    
    var messageHtml = messageFirst +  messageSecond + messageThird; // Full HTML message
    var messagePlain = messageHtml.replace(/(<([^>]+)>)/ig, ""); // Full plain message
    
    GmailApp.sendEmail(email, subject, messagePlain, // Send the email
                       { 
                         from: replyTo,
                         htmlBody: messageHtml, 
                         replyTo: replyTo
                       });
    sheet.getRange(startRow + i + 1, headersData.length).setValue(EMAIL_SENT); // Add EMAIL_SENT to the last column
    SpreadsheetApp.flush(); // Update the spreadsheet data
  }
}