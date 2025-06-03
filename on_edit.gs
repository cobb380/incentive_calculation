function onEdit(e) {
  // 編集された範囲を取得
  var range = e.range;
  var sheet = range.getSheet();
  
  // 編集されたセルがE2またはF2の場合に処理を実行
  if ((range.getA1Notation() === 'F2' || range.getA1Notation() === 'G2') && sheet.getName() === '日給計算') {
    var eValue = sheet.getRange('F2').getValue();
    var fValue = sheet.getRange('G2').getValue();
    
    // F2の値に一致するA列の行を検索
    var aValues = sheet.getRange('B:B').getValues();
    var bValue;
    for (var i = 0; i < aValues.length; i++) {
      if (aValues[i][0] === fValue) {
        bValue = sheet.getRange('C' + (i + 1)).getValue(); // B列の値を取得
        break;
      }
    }
    
    // B列の値が見つかった場合、G2セルに計算結果を代入
    if (bValue) {
      var result = eValue * 0.7 / bValue * 10000;
      sheet.getRange('H2').setValue(result);
    } else {
      sheet.getRange('H2').setValue("一致する値が見つかりません");
    }
  }
}
