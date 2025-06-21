function setUpCalendar() { 
  var eventCalA = CalendarApp.getCalendarById("google帳號");
  var eventCalB = CalendarApp.getCalendarById("日曆的ID");
  var form = FormApp.openById("表單ID");
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var numRows = sheet.getLastRow();
  var colum = sheet.getLastColumn();
  var rang = sheet.getRange(2, 1, numRows-1, colum).getValues();
  var now = new Date();
  var early_day =now.getDate() - 2;    
  var early_min = now.getMinutes() - 1;
  var early_daytime = new Date(now);    
  var early_time = new Date(now);
  early_daytime.setDate(early_day);    
  early_time.setMinutes(early_min);

  var formResponses = form.getResponses();
  var timestamps = [], urls = [], url = [];
  for (var i = 0; i < formResponses.length; i++) {
    timestamps.push(formResponses[i].getTimestamp().setMilliseconds(0));
    urls.push(formResponses[i].getEditResponseUrl());};
  for (var j = 0; j < rang.length; j++){url.push([rang[j][0]?urls[timestamps.indexOf(rang[j][0].setMilliseconds(0))]:'']);};
  sheet.getRange(2, 40, url.length).setValues(url);
  
  for(k = 0; k < rang.length; k++){
    var request = rang[k];
    var [aa, bb] = [new Intl.NumberFormat('en-US'), "已執行"];
    var title = request[3];
    var [start_time, end_time] = [new Date(request[6]), new Date(request[7])];
    var [theday, lasttheday] = [new Date(request[6]), new Date(request[54])];
    var [stra, strb, strc, strd] = ["修改", "已取消", "已延期", sheet.getRange(k + 2, 38).getValue()];
    if(end_time <= now && (request[5] === "洽談中" || request[5] === "待執行")){sheet.getRange(k + 2, 6).setValue(bb);};
    
    var business = "業務：" + request[1] + '\n';
    var type = "類型：" + request[2] + '\n';    
    var user = "團名：" + title + '\n';
    var habitude = "分類：" + request[4] + '\n';
    var undertake = "承辦狀況：" + request[5] + '\n';    
    var company = "公司名稱：" + request[8] + '\n';
    if(request[9] !== ""){number = "統編：" + request[9] + '\n';}else{number = "";};
    var contact = "聯絡窗口：" + request[10] + '\n';
    var sellphone = "聯絡電話：" + request[11] + '\n';
    if(isNaN(request[12]) === true){person = "人數：" + request[12] + "\n";}else{person = "人數：" + aa.format(request[12]) + "人\n";};
    var tables = request[12]/10;
    var net = "專案價：NT$" + aa.format(request[14]) + "元/人\n";    
    var income = request[14] * request[12];
    if(request[15] !== ""){
      incomes = income - request[15];
      discount ="小計：NT$" + aa.format(incomes) + "元\n總價特別折扣：NT$" + aa.format(request[15]) + '元\n';
      }else{
        incomes = income;
        discount = "";
      };
    var depositM = Math.ceil(incomes * 0.3/1000) * 1000;
    var deposit = "訂單確認後需支付訂金：NT$" + aa.format(depositM) + "元\n"; 
    var calculate = incomes - depositM;
    if(calculate !== 0){payment = "尾款需支付金額：NT$" + aa.format(calculate) + "元\n";};
    if(incomes !== 0){incomess = "預估團費：NT$"+ aa.format(incomes) + "元\n";};
    if(request[18] !== ""){remark = "備註：" + '\n' + request[18] + '\n';}else{remark = "";};
    var name = "遊程名稱：" + request[20] + '\n';
    if(request[20] === "團體餐飲活動"){food = person + "保證桌數：" + Math.round(tables) + "桌/每桌10位\n";}else{food = person;};
    if(request[22] !== ""){stroke = "行程："+ '\n' + request[22] + '\n';}else{stroke = "";};
    var ordernumber = "訂單編號：" + request[37] + '\n';
    if(request[38] === ""){sheet.getRange(k + 2, 39).setFormula('=HYPERLINK(INDIRECT("AN"&ROW()),"修改")')};
    var revise = "訂單連結：" + stra.link(sheet.getRange(k + 2, 40).getValue());    
    if(request[54] === ""){sheet.getRange(k + 2, 55).setValue(request[6])};    
  
    var informationA = ordernumber + business + type + undertake;
    var informationB = name + user + habitude + company + number + contact + sellphone + food;
    var informationC = net + discount + incomess + deposit + payment;
    var informationD = stroke;
    var informationE = remark + revise;
    var allA = informationA + informationB + informationD;
    var allB = informationA + informationB + informationC + informationD + informationE;
    
    if((request[0] >= early_time && request[0] <= now) || (request[7] >= early_daytime && request[7] <= now && request[5] === bb)){
      function createEvent(calendar, title, start_time, end_time, description, color){
        var event = calendar.createEvent(title, start_time, end_time, { description: description });
        event.setColor(color);
        return event;
      }
      function deleteEvents(events, keyword){        
        for (var ch = 0; ch < events.length; ch++){
          var event = events[ch];
          var eventDescription = event.getDescription();
          var eventTitle = event.getTitle();
          if(keyword.some(keyword => eventDescription.search(keyword) > -1)){
            event.deleteEvent();
          }else if(eventTitle === title){
            event.deleteEvent();
            }
        }
      }      
      if(theday !== lasttheday){
        deleteEvents(eventCalA.getEventsForDay(lasttheday), [strd]);
        deleteEvents(eventCalB.getEventsForDay(lasttheday), [strd]);
        sheet.getRange(k + 2, 55).setValue(request[6]);
      }
      if(request[5] === strb || request[5] === strc){
        deleteEvents(eventCalA.getEventsForDay(theday), [strb]);
        deleteEvents(eventCalA.getEventsForDay(theday), [strc]);
        deleteEvents(eventCalB.getEventsForDay(theday), [strb]);
        deleteEvents(eventCalB.getEventsForDay(theday), [strc]);      
      }
      if(request[5] !== strb && request[5] !== strc){
        createEvent(eventCalA, title, start_time, end_time, allA, "11");
        createEvent(eventCalB, title, start_time, end_time, allB, "11");
      }
    }
  }
}       