function setBordersToMatchCellColors() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("E8:AG107"); // 指定的範圍
  var rangea = sheet.getRange("D7:AI7");
  var colors = range.getBackgrounds(); // 取得每個儲存格的背景色
  var colora = rangea.getBackgrounds();
  
  // 設定每個儲存格的框線
  for (var i = 0; i < colors.length; i++) {
    for (var j = 0; j < colors[0].length; j++) {
      var cell = range.getCell(i+1, j+1);
      var color = colors[i][j];
      
      cell.setBorder(false, true, false, true, false, false, color, SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  for (var k = 0; k < colora.length; k++) {
    for (var l = 0; l < colora[0].length; l++) {
      var cella = rangea.getCell(k+1, l+1);
      var coloraa = colora[k][l];

      cella.setBorder(true, false, false, false, false, false, coloraa, SpreadsheetApp.BorderStyle.SOLID);
      cella.setBorder(null, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}