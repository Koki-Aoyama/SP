function createCalendar() 
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const myCalender = CalendarApp.getCalendarById("z19003k1@shibaurafzk.com");
  let startDate = sheet.getRange("B3").getDisplayValue();
  const startTime = sheet.getRange("C3").getDisplayValue();
  const endDate = sheet.getRange("D3").getDisplayValue();
  const endTime = sheet.getRange("E3").getDisplayValue();
  const title = sheet.getRange("F3").getValue();
  const content  = sheet.getRange("G3").getValue();
  let mailAddress = Session.getActiveUser().getUserLoginId();
  const startDataTime = new Date(startDate + " " + startTime);
  const endDataTime = new Date(endDate + " " + endTime);

  if(title != "" && content != "")
  {
    myCalender.createEvent(title,startDataTime,endDataTime,{description:content});
    addHistory(mailAddress,startDataTime,endDataTime,title,content);
    //初期化処理
    sheet.getRange("B3").setValue(new Date());
    sheet.getRange("D3").setValue(new Date());
    sheet.getRange("C3").setValue("0:00");
    sheet.getRange("E3").setValue("0:00");
    sheet.getRange("F3").setValue("");
    sheet.getRange("G3").setValue("");
  }
  else if(title == "" && content == "")
  {
    SpreadsheetApp.getUi().alert("タイトルと詳細が入力されていません")
  }
  else if(content == "")
  {
    SpreadsheetApp.getUi().alert("詳細が入力されていません")
  }
  else if (title == "")
  {
    SpreadsheetApp.getUi().alert("タイトルが入力されていません")
  }
  
}

function addHistory(address,startdate,enddate,title,content)
{
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("追加履歴");
  const histroy = [address,new Date(),startdate,enddate,title,content];
  sheet.appendRow(histroy);
}

function createDateTimeList() {
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("予定追加");
  var dateTime = [];
  const cell = sheet.getRange('C3');

  for(let date = 0;date <24;date++)
  {
    dateTime.push(`${date}:00`,`${date}:30`);
  }
  console.log(dateTime)
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(dateTime).build();
  cell.setDataValidation(rule);
  
}
