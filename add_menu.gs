function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('メニュー')
    .addItem('インセンティブを計算', 'calculateIncentives')
    .addToUi();
}