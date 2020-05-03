function contact_Gmail() {
  var rowNumber = 2;  
  var mysheetname = 'Gmail解析_' + DateString(new Date());

  var GmailSS = SpreadsheetApp.create(mysheetname);
  var mySheet = GmailSS.getSheets()[0];
  mySheet.setName(mysheetname);
  mySheet.getRange(1,1).setValue("日時");
  mySheet.getRange(1,2).setValue("送信元");
  mySheet.getRange(1,3).setValue("件名");
  mySheet.getRange(1,4).setValue("本文");
  mySheet.getRange(1,5).setValue("ご予約");
  mySheet.getRange(1,6).setValue("お名前");
  mySheet.getRange(1,7).setValue("電話番号");
  mySheet.getRange(1,8).setValue("メールアドレス");
  
  var newfolder = DriveApp.createFolder(mysheetname); 
  
  var searchQuery = 'subject:(”予約が確定" OR "finalized) '; 
  var threads = GmailApp.search(searchQuery, 0, 200);

  var mymsg=[];
  
  var msgs = GmailApp.getMessagesForThreads(threads);
  
  for(var i = 0; i < msgs.length; i++) {    
    mymsg[i]=[];
    for(var j = 0; j < msgs[i].length; j++) {
      mymsg[i][0] = msgs[i][j].getDate();
      mymsg[i][1] = msgs[i][j].getFrom();
      mymsg[i][2] = msgs[i][j].getSubject();
      var nbsp = String.fromCharCode(160);
      mymsg[i][3] = msgs[i][j].getPlainBody().replace(/<("[^"]*"|'[^']*'|[^'">])*>|nbsp/g,'').replace(/&; |　/g,'').substring(0,50000);
      sheet.appendRow([
      mymsg[i][4] = fetchData(mymsg[i][3],'Appointment:','\r'),
      id
      ]);
//      mymsg[i][4] = '=MID(r[0]c[-1],FIND("Appointment",D2)+17,FIND("Name",r[0]c[-1])-(FIND("Appointment",r[0]c[-1])+17))';
//      mymsg[i][5] = '=MID(r[0]c[-2],FIND("Name",r[0]c[-2])+6,FIND("Phone",r[0]c[-2])-(FIND("Name",r[0]c[-2])+6))';
//      mymsg[i][6] = '=MID(r[0]c[-3],FIND("Phone number",r[0]c[-3])+14,FIND("email address",r[0]c[-3])-(FIND("Phone number",r[0]c[-3])+14))';
//      mymsg[i][6] = '=extractline(E2,10)';
    } 
  }
  if(mymsg.length>0){  
    GmailSS.getSheets()[0].getRange(2, 1, i, 7).setValues(mymsg); //シートに貼り付け
  }
}
                     
function DateString(date){
  return date.getFullYear().toString()
  + date.getMonth().toString()
  + date.getDate().toString()
  + date.getHours().toString()
  + date.getMinutes().toString()
  + date.getSeconds().toString();
};

function fetchData(str, pre, suf) {
  var reg = new RegExp(pre + '.*?' + suf);
  var data = str.match(reg)[0]
    .replace(pre, '')
    .replace(suf, '');
  return data;  
}