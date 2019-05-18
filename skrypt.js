//variables
var EMAIL_SENT = 'EMAIL_SENT';
var FIRST_REMINDER = 30;
var SECOND_REMINDER = 14;
var THIRD_REMINDER = 1;
var ROW_FIRST_EMAIL_SENT = 10;
var ROW_SECOND_EMAIL_SENT = 11;
var ROW_THIRD_EMAIL_SENT = 12;

// helpers
var numDaysBetween = function(d1, d2) {
  var diff = d1.getTime() - d2.getTime();
  return diff / (1000 * 60 * 60 * 24);
};

var addDays = function(date, days) {
  var result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
};


// functions 

 function getFormattedDate(date) {
  var year = date.getFullYear();

  var month = (1 + date.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;

  var day = date.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  
  return day  + '/' + month + '/' + year;
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Baza członków ESN Polska')
      .addItem('Force I reminder (30 days)', 'forceFirstReminder')
      .addItem('Force II reminder (14 days)', 'forceSecondReminder')
      .addItem('Force III reminder (1 day)', 'sendThirdReminder')
      .addItem('SPAM TOTALNY (wyślij wszystko)', 'forceEmails')
      .addToUi();
}

function forceEmails() {
  sendFirstReminder();
  sendSecondReminder();
  sendThirdReminder();
}

function forceFirstReminder() {
  sendFirstReminder();
}
function forceSecondReminder() {
  sendSecondReminder();
}
function forceThirdReminder() {
  sendThirdReminder();
}

function sendFirstReminder() {
  sendXXdaysReminder(FIRST_REMINDER,ROW_FIRST_EMAIL_SENT);
}
function sendSecondReminder() {
  sendXXdaysReminder(SECOND_REMINDER,ROW_SECOND_EMAIL_SENT);
}
function sendThirdReminder() {
  sendXXdaysReminder(THIRD_REMINDER,ROW_THIRD_EMAIL_SENT);
}


function FindRows() {
range = SpreadsheetApp.getActiveSheet().getLastRow();
return range;
}

function createMail(dataValid) {

  var htmlIloscDni =  
    '<body>' + 
      '<p> Cześć, <br/>' +
      '<p> Jeżeli otrzymujesz tę wiadomość oznacza to, że już  ' + dataValid + ' Twoja karta członkowska straci ważność.</p> ' +
      '<p>Twoim obowiązkiem jako członka zwyczajnego jest opłacenie składki w powyższym terminie. Prosimy o niezwłoczne dokonanie opłaty poprzez system card.esn.pl.<br/>' +
      'Brak uiszczonej w terminie składki będzie równał się odebraniu członkostwa zwyczajnego Stowarzyszenia ESN Polska.</p>' +
      '<p>Wszelkie wątpliwości dotyczące płatności składki członkowskiej prosimy kierować na adres office@esn.pl. Prosimy również o nieodpowiadanie na tę wiadomość. </p> <p> Pozdrawiamy </p>' +
    '</body>' ;
  
  return htmlIloscDni;
}


function sendXXdaysReminder(numberOfDays,rowEmailSendPosition) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dateToday = new Date();
  var startRow = 2; // First row of data to process
  var numRows = FindRows(); // Number of rows to process


  var dataRange = sheet.getRange(startRow, 1, numRows, 14);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  for (i in data) {
    var row = data[i];
    var emailAddress = row[2]; // First column
    var dataValid = row[8]; // second column
    var emailSent = row[rowEmailSendPosition];
    
    if(dataValid=="")
      return;
    
    var tempDiff = numDaysBetween(dataValid, dateToday);
    
   if (tempDiff <= numberOfDays && tempDiff > 0 && emailSent != EMAIL_SENT) {
      
     
      var dataValidFormatted = getFormattedDate(dataValid); 

      var message = createMail(dataValidFormatted);
      var subject = 'Wygasająca składka członkowska ESN Polska';
     
     MailApp.sendEmail({
       to: emailAddress,
       subject: subject,
       htmlBody: message,
     });     
      sheet.getRange(startRow + parseInt(i), rowEmailSendPosition+1).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
