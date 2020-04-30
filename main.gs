function myFunction(){
  var mailQuery = 'subject:("予約が確定")';
  var threads = GmailApp.search(mailQuery);
  var messages = GmailApp.getMessagesForThreads(threads);
  var sheet = SpreadsheetApp.getActiveSheet();
  for(var i=0; i<messages.length; i++){
    var plainBody = messages[i][0].getPlainBody();
    sheet.appendRow([plainBody.match(/次回のご予約.*/)[0].replace('次回のご予約:', '') ,])
  }
}
