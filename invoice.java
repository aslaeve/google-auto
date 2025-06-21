function arabicToChinese(number) {
  var units = ["", " 拾", " 佰", " 仟", " 萬", " 拾", " 佰", " 仟", " 萬", " 億"];
  var digits = [" 零", " 壹", " 貳", " 參", " 肆", " 伍", " 陸", " 柒", " 捌", " 玖"];

  var numberStr = String(number);
  var result = "";
  var consecutiveZeros = 0; // 記錄連續出現的零的數量

  for (var i = 0; i < numberStr.length; i++) {
    var digit = parseInt(numberStr.charAt(i));
    
    if (digit !== 0) {
      result += digits[digit] + units[numberStr.length - 1 - i];
      consecutiveZeros = 0;
    } else {
      consecutiveZeros++;
      
      // 只有在連續零的數量不超過1時才加入"零"
      if (consecutiveZeros <= 1) {
        result += digits[digit];
      }
    }
  }
  
  if (result.endsWith(digits[0])) {
    result = result.slice(0, -1); // 移除最後一個零
  }

  return result;
}