function move() {
  var folders = {
    negotiating: DriveApp.getFolderById("洽談中ID"),
    toExecute: DriveApp.getFolderById("待執行ID"),
    executed: DriveApp.getFolderById("已完成ID"),
    postpone: DriveApp.getFolderById("已延期ID"),
    cancel: DriveApp.getFolderById("已取消ID"),
  }
  var templateFile = DriveApp.getFileById("範例ID");
  var sheet = SpreadsheetApp.getActiveSheet();

  var numRows = sheet.getLastRow();
  var colum = sheet.getLastColumn();
  var rang = sheet.getRange(2,1,numRows-1,colum).getValues();
  var now = new Date();
  var early_sec = now.getSeconds() - 10;
  var earlytime_move = new Date(now);
  earlytime_move.setSeconds(early_sec);
  
  for(k = 0; k < rang.length; k++){
    var request = rang[k];
    
    var title = request[3];

    if(request[0] >= earlytime_move && request[0] <= now){
      var rangeValueColumn4 = title;
      var startColumn = 51;
      var endColumn = 54;
      var setValueColumn = startColumn;
      var foundMatch = false;

      while(setValueColumn <= endColumn){
        var rangeValue = sheet.getRange(k + 2, setValueColumn).getValue();
        if(rangeValue === ""){
          sheet.getRange(k + 2, setValueColumn).setValue(title);
          break;
          }else if(rangeValue == rangeValueColumn4){
            foundMatch = true;
            break;}
        setValueColumn++;
        if(setValueColumn > endColumn){setValueColumn = startColumn;}
      };
      if(!foundMatch){sheet.getRange(k + 2, setValueColumn).setValue(title);};
    
    if(request[55] === ""){sheet.getRange(k + 2, 56).setValue(request[6])};
    if(request[56] === ""){sheet.getRange(k + 2, 57).setValue(request[5])};
    var folderName = request[2] === "銷售" ? "業務團務單-" : "公關團務單-";

    function parseDate(dateStr) {
      var formattedDate = Utilities.formatDate(dateStr, "GMT+0800", "yyyy/MM/dd").split(" ")[0].split("/");
      return {
        year: parseInt(formattedDate[0]) - 1911,
        month: formattedDate[1],
        day: formattedDate[2]
      };
    }    
    var timestampA = parseDate(sheet.getRange(k + 2, 7).getValue());
    var [yearA, monthA, dayA] = [timestampA.year, timestampA.month, timestampA.day];
    var timestampB = parseDate(sheet.getRange(k + 2, 56).getValue());
    var [yearB, monthB, dayB] = [timestampB.year, timestampB.month, timestampB.day];    

    var namedateA = folderName + yearA + monthA + dayA + "-";
    var namedateB = monthA + dayA + "-";
    var namedateC = folderName + yearB + monthB + dayB + "-";
    var namedateD = monthB + dayB + "-";
    var namedateE = folderName + title;
    var [newFolderNameA, changeNewFolderName] = [namedateA, namedateA];
    var [newFolderNameB, newFileNameB] = [namedateC, namedateD];
    var [newFileNameA, changeNewFileName] = [namedateB, namedateB];    
    var [oldFolderName, oldFileName] = [namedateE, title];
    var rangesToUpdateA = [newFolderNameA, newFolderNameB, newFileNameA, newFileNameB];
    var rangesToUpdateB =  [changeNewFolderName, changeNewFileName];
    var rangesToUpdateC = [oldFolderName, oldFileName];
    var rangeValues = [];

    for(i = 51; i <= 54; i++){rangeValues[i - 51] = sheet.getRange(k + 2, i).getValue();};
    if(rangeValues[0] !== "" && rangeValues[1] === ""){
      for(n = 0; n < rangesToUpdateA.length; n++){rangesToUpdateA[n] += rangeValues[0];};
      for(m = 0; m < rangesToUpdateB.length; m++){rangesToUpdateB[m] += rangeValues[0];};
      sheet.getRange(k + 2, 53).clearContent();
    }else if(rangeValues[1] !== "" && rangeValues[2] === ""){
      for(n = 0; n < rangesToUpdateA.length; n++){rangesToUpdateA[n] += rangeValues[0];};
      for(m = 0; m < rangesToUpdateB.length; m++){rangesToUpdateB[m] += rangeValues[1];};
      sheet.getRange(k + 2, 54).clearContent();
    }else if(rangeValues[2] !== "" && rangeValues[3] === ""){
      for(n = 0; n < rangesToUpdateA.length; n++){rangesToUpdateA[n] += rangeValues[1];};
      for(m = 0; m < rangesToUpdateB.length; m++){rangesToUpdateB[m] += rangeValues[2];};
      sheet.getRange(k + 2, 51).clearContent();
    }else if(rangeValues[3] !== "" && rangeValues[0] === ""){
      for(n = 0; n < rangesToUpdateA.length; n++){rangesToUpdateA[n] += rangeValues[2];};
      for(m = 0; m < rangesToUpdateB.length; m++){rangesToUpdateB[m] += rangeValues[3];};
      sheet.getRange(k + 2, 52).clearContent();
    }
    
    var foldersA1 = folders.negotiating.getFoldersByName(rangesToUpdateA[0]);
    var foldersA2 = folders.negotiating.getFoldersByName(rangesToUpdateB[0]);
    var foldersA3 = folders.negotiating.getFoldersByName(rangesToUpdateA[1]);
    var foldersB1 = folders.toExecute.getFoldersByName(rangesToUpdateA[0]);
    var foldersB2 = folders.toExecute.getFoldersByName(rangesToUpdateB[0]);
    var foldersB3 = folders.toExecute.getFoldersByName(rangesToUpdateA[1]);
    var foldersC1 = folders.postpone.getFoldersByName(rangesToUpdateC[0]);
    var foldersD1 = folders.cancel.getFoldersByName(rangesToUpdateC[0]);
    var foldersE1 = folders.executed.getFoldersByName(rangesToUpdateA[0]);
   
    if(request[0] >= earlytime_move && request[0] <= now){
      if(request[5] === "洽談中" || request[5] === "待執行"){
        var folderToMoveA = foldersC1.hasNext() ? foldersC1 : foldersD1;
        var folderToMoveB = foldersA1.hasNext() ? foldersA1 : foldersB1;
        var targetfolderA = request[5] === "洽談中" ? folders.negotiating : folders.toExecute;
        var tomovetargets = request[5] === "洽談中" ? foldersA1 : foldersB1;
        var olderstargetA = request[5] === "洽談中" ? foldersA3 : foldersB3;
        var changetargetA = request[5] === "洽談中" ? foldersA2 : foldersB2;
        if(foldersC1.hasNext() || foldersD1.hasNext()){
          folderToMoveA.next().moveTo(targetfolderA).setName(rangesToUpdateA[0]);
          var changeFileA = tomovetargets.next().getFilesByName(rangesToUpdateC[1]).next().setName(rangesToUpdateA[2]);
          SpreadsheetApp.open(changeFileA).getSheetByName("預訂單").getRange("O5").setValue(title);
        }else if((foldersA3.hasNext() || foldersB3.hasNext()) && rangesToUpdateA[0] !== rangesToUpdateB[0]){
          olderstargetA.next().setName(rangesToUpdateB[0]);
          var changeFileB = changetargetA.next().getFilesByName(rangesToUpdateA[3]).next().setName(rangesToUpdateB[1]);
          SpreadsheetApp.open(changeFileB).getSheetByName("預訂單").getRange("O5").setValue(title);
        }else if(foldersA1.hasNext() || foldersB1.hasNext()){folderToMoveB.next().moveTo(targetfolderA);
        }else{
          if(foldersA1.hasNext() || foldersB1.hasNext() || foldersA2.hasNext() || foldersB2.hasNext() || foldersE1.hasNext()){break;};
          var newFolder = targetfolderA.createFolder(rangesToUpdateA[0]);
          var newFile = templateFile.makeCopy(newFolder).setName(rangesToUpdateA[2]);
          SpreadsheetApp.open(newFile).getSheetByName("預訂單").getRange("O5").setValue(title);
        }
      }else if(request[5] === "已延期" || request[5] === "已取消"){
        var targetfolderB = request[5] === "已延期" ? folders.postpone : folders.cancel;
        var changetargetB = request[5] === "已延期" ? foldersC1 : foldersD1;
        var olderstargetB = foldersC1.hasNext() ? foldersC1 : foldersD1;
        var olderstargetC = foldersA1.hasNext() ? foldersA1 : foldersB1;
        if(foldersA1.hasNext() || foldersB1.hasNext()){
          olderstargetC.next().moveTo(targetfolderB).setName(rangesToUpdateC[0]);
          changetargetB.next().getFilesByName(rangesToUpdateA[2]).next().setName(rangesToUpdateC[1]);
        }else if(foldersC1.hasNext() || foldersD1.hasNext()){olderstargetB.next().moveTo(targetfolderB);}
      }
      sheet.getRange(k + 2, 56).setValue(request[6]);
    }else if(request[5] === "已執行" && (request[56] === "洽談中" || request[56] === "待執行")){
      var olderstargetD = foldersA1.hasNext() ? foldersA1 : foldersB1;
      if(foldersA1.hasNext() || foldersB1.hasNext()){olderstargetD.next().moveTo(folders.executed);}
    }
    sheet.getRange(k + 2, 57).setValue(request[5]);
  }
 }
}