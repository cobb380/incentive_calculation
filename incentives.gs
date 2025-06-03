// 定数の定義
const HOLIDAY_SHEET_NAME = '祝日';
const INCENTIVE_SHEET_NAME = 'インセンティブ';
const BENCHMARK_SHEET_NAME = '基準';
const POSITION_SALARY_SHEET_NAME = '役職手当'; // 役職手当のシート名を定義

// メイン関数
function calculateIncentives() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var defaultSheetName = getSheetName(spreadsheet, '【対象のスケジュールシートを下記から記載してください】');
  
  var sheet = getSheetByName(spreadsheet, defaultSheetName);
  var benchmarkSheet = getSheetByName(spreadsheet, BENCHMARK_SHEET_NAME);
  var incentiveSheet = getOrCreateSheet(spreadsheet, INCENTIVE_SHEET_NAME);
  var positionSalarySheet = getSheetByName(spreadsheet, POSITION_SALARY_SHEET_NAME);

  var { year, month } = getInputYearMonth();
  if (!year || !month) return;

  // ①「基準」シートから各スタッフの基準時間とあわせて「ID」も取得する
  var benchmarkData = benchmarkSheet.getDataRange().getValues();
  var benchmarkTimes = getBenchmarkTimes(benchmarkSheet);
  var staffIds = getStaffIds(benchmarkData); // ← 新たに追加（IDとスタッフ名の紐付け）

  var holidays = getJapaneseHolidays(year, month);

  var externalSpreadsheetIds = getExternalSpreadsheetIds();
  var staffTabs1 = getStaffTabs(externalSpreadsheetIds[0]);
  var staffTabs2 = getStaffTabs(externalSpreadsheetIds[1]);

  var data = sheet.getDataRange().getValues();

  // ② IDを使えるようにパラメータで staffIds を渡す
  var incentiveData = calculateIncentiveData(
    data,
    year,
    month,
    holidays,
    benchmarkTimes,
    staffIds,     // ← 追加
    staffTabs1,
    staffTabs2
  );

  // 役職手当の計算
  var positionBonuses = calculatePositionBonuses(benchmarkSheet, positionSalarySheet, incentiveData);

  // インセンティブデータをインセンティブシートに書き込む
  incentiveSheet.getRange(1, 1, incentiveData.length, incentiveData[0].length).setValues(incentiveData);

  // ③ 役職手当や最終合計を反映する列番号を調整
  //    incentiveData[0] = [ 'ID', 'スタッフ名', '基準時間', '訪問時間', '件数確認', '合計', '時間外訪問', 'インセンティブ' ];
  //    → "役職手当" は 9 列目, "最終合計" は 10 列目になる（index 的には +1 の状態）

  for (var i = 1; i < incentiveData.length; i++) {
    var staffName = incentiveData[i][1]; // [1] がスタッフ名
    var positionBonus = positionBonuses[staffName] || 0; 
    // (i+1) 行目の 9 列目（列 A=1, B=2,... → I=9）に役職手当
    incentiveSheet.getRange(i + 1, 9).setValue(positionBonus); 

    // incentiveData[i][7] が元々のインセンティブ
    var totalIncentive = incentiveData[i][7] + positionBonus; 
    // (i+1) 行目の 10 列目に最終合計
    incentiveSheet.getRange(i + 1, 10).setValue(totalIncentive); 
  }

  // 1行目にヘッダーを表示
  incentiveSheet.getRange(1, 9).setValue('役職手当');
  incentiveSheet.getRange(1, 10).setValue('最終合計');

  setRowColors(incentiveSheet);
}

// 役職手当を計算する関数
function calculatePositionBonuses(benchmarkSheet, positionSalarySheet, incentiveData) {
  var benchmarkData = benchmarkSheet.getDataRange().getValues();
  var positionSalaryData = positionSalarySheet.getDataRange().getValues();

  var groupVisitTimes = {};
  var positionBonuses = {};

  // 各スタッフのグループ訪問時間を計算
  // incentiveData[i] = [ID, スタッフ名, 基準時間, 訪問時間, 件数確認, 合計, 時間外訪問, インセンティブ];
  for (var i = 1; i < incentiveData.length; i++) {
    var staffName = incentiveData[i][1]; // [1] がスタッフ名
    var visitTime = incentiveData[i][3]; // [3] が訪問時間
    var groupKey = getGroupKey(staffName, benchmarkData); // グループキー（事務所と職種）

    if (!groupVisitTimes[groupKey]) groupVisitTimes[groupKey] = 0;
    groupVisitTimes[groupKey] += visitTime;
  }

  // 役職手当の計算
  for (var i = 1; i < benchmarkData.length; i++) {
    var staffName = benchmarkData[i][1];
    var position = benchmarkData[i][5]; // 役職
    var groupKey = getGroupKey(staffName, benchmarkData);
    var totalGroupVisitTime = groupVisitTimes[groupKey] || 0;
    var percentage = benchmarkData[i][6] || 1; // 割合がない場合は1とする

    if (position) {
      var positionBonus = getPositionBonus(position, totalGroupVisitTime, positionSalaryData);
      positionBonuses[staffName] = positionBonus * percentage;
    }
  }

  return positionBonuses;
}

// グループキーを取得する関数（事務所と職種でグループ化）
function getGroupKey(staffName, benchmarkData) {
  for (var i = 1; i < benchmarkData.length; i++) {
    if (benchmarkData[i][1] === staffName) {
      var office = benchmarkData[i][3]; // 事務所
      var jobType = benchmarkData[i][4]; // 職種
      return `${office}_${jobType}`;
    }
  }
  return '';
}

// 役職手当の金額を取得する関数
function getPositionBonus(position, totalGroupVisitTime, positionSalaryData) {
  for (var i = 1; i < positionSalaryData.length; i++) {
    var achievement = positionSalaryData[i][1]; // 実績
    if (totalGroupVisitTime < achievement) {
      var positionIndex = positionSalaryData[0].indexOf(position); // 役職の列番号
      if (positionIndex > 0) {
        // 実績を下回った "一つ上の行" の手当を返す → i-1行目
        return positionSalaryData[i - 1][positionIndex];
      }
    }
  }
  return 0; // 該当なしの場合は 0 を返す
}

// ④ 「基準」シートから ID 列とスタッフ名列を紐づける関数を新規追加
function getStaffIds(benchmarkData) {
  // benchmarkData[i] の構成例 → [A列(ID), B列(スタッフ名), C列(別情報), D列(基準時間), ...]
  var staffIds = {};
  for (var i = 1; i < benchmarkData.length; i++) {
    var staffName = benchmarkData[i][1];  // B列
    var staffId = benchmarkData[i][0];    // A列
    staffIds[staffName] = staffId;
  }
  return staffIds;
}

// シート名を選択させる関数
function getSheetName(spreadsheet, promptText) {
  var sheets = spreadsheet.getSheets();
  var sheetNames = sheets.map(sheet => sheet.getName());

  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('シートの選択', promptText + '\n' + sheetNames.join('\n'), ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() != ui.Button.OK) {
    ui.alert('シートの選択がキャンセルされました。');
    throw new Error('シートの選択がキャンセルされました。');
  }

  var selectedSheetName = response.getResponseText();
  if (!sheetNames.includes(selectedSheetName)) {
    ui.alert('無効なシート名が入力されました。');
    throw new Error('無効なシート名が入力されました。');
  }

  return selectedSheetName;
}

// シートを取得する関数
function getSheetByName(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`対象のシート名を「${sheetName}」にしてください。（スペースも要確認）`);
  }
  return sheet;
}

// シートを取得または作成する関数
function getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return spreadsheet.insertSheet(sheetName);
  } else {
    sheet.clear();
    return sheet;
  }
}

// 年と月を入力させる関数
function getInputYearMonth() {
  var ui = SpreadsheetApp.getUi();
  var yearResponse = ui.prompt('計算対象年の入力', '年を【半角数字で】入力してください\n (例: 2024)', ui.ButtonSet.OK_CANCEL);
  if (yearResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('年の入力がキャンセルされました。');
    return {};
  }
  var monthResponse = ui.prompt('計算対象月の入力', '月を【半角数字で】入力してください\n (例: 1)', ui.ButtonSet.OK_CANCEL);
  if (monthResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('月の入力がキャンセルされました。');
    return {};
  }
  
  var year = parseInt(yearResponse.getResponseText());
  var month = parseInt(monthResponse.getResponseText());

  if (isNaN(year) || isNaN(month) || month < 1 || month > 12) {
    ui.alert('無効な年または月が入力されました。半角で入力したか確認ください。');
    return {};
  }

  return { year, month };
}

// 祝日を取得する関数（「祝日」シートから取得する）
function getJapaneseHolidays(year, month) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var holidaySheet = spreadsheet.getSheetByName(HOLIDAY_SHEET_NAME);
  if (!holidaySheet) {
    throw new Error('「' + HOLIDAY_SHEET_NAME + '」シートが見つかりません。シート名を確認してください。');
  }
  
  var holidayData = holidaySheet.getDataRange().getValues();
  var holidays = [];
  
  // ヘッダー行がある前提で、2行目（インデックス1）からループします。
  // ヘッダーがなければ i=0 からに変更してください。
  for (var i = 1; i < holidayData.length; i++) {
    var row = holidayData[i];
    // A列: 月、B列: 日
    var sheetMonth = parseInt(row[0], 10);
    var sheetDay = parseInt(row[1], 10);
    
    // 入力された月と一致する場合、指定された年と組み合わせて日付を生成
    if (sheetMonth === month) {
      var holidayDate = new Date(year, month - 1, sheetDay);
      var formattedHoliday = Utilities.formatDate(holidayDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      holidays.push(formattedHoliday);
    }
  }
  return holidays;
}

// 基準時間シートからデータを取得する関数
function getBenchmarkTimes(sheet) {
  var benchmarkData = sheet.getDataRange().getValues();
  var benchmarkTimes = {};
  for (var i = 1; i < benchmarkData.length; i++) {
    // スタッフ名は B列、基準時間は C列
    var staffName = benchmarkData[i][1];
    var benchmarkTime = benchmarkData[i][2];
    benchmarkTimes[staffName] = benchmarkTime;
  }
  return benchmarkTimes;
}


// 外部スプレッドシートIDを入力させる関数
function getExternalSpreadsheetIds() {
  var ui = SpreadsheetApp.getUi();
  var urlResponse1 = ui.prompt('件数確認表の入力（1個目）', '件数確認表のURLを入力してください \n(例: https://docs.google.com/spreadsheets/d/\n1111111111111111111111aaaaaaaaaaaaaaaaaaaaaa/edit?usp=drive_link)', ui.ButtonSet.OK_CANCEL);
  if (urlResponse1.getSelectedButton() != ui.Button.OK) {
    ui.alert('URLの入力がキャンセルされました。');
    throw new Error('外部スプレッドシート1のURL入力がキャンセルされました。');
  }
  var url1 = urlResponse1.getResponseText();
  var externalSpreadsheetId1 = extractSpreadsheetId(url1);

  var urlResponse2 = ui.prompt('件数確認表の入力（2個目）', '件数確認表のURLを入力してください\n (例: https://docs.google.com/spreadsheets/d/\n2222222222222222222222bbbbbbbbbbbbbbbbbbbbbb/edit?usp=drive_link)', ui.ButtonSet.OK_CANCEL);
  if (urlResponse2.getSelectedButton() != ui.Button.OK) {
    ui.alert('URLの入力がキャンセルされました。');
    throw new Error('外部スプレッドシート2のURL入力がキャンセルされました。');
  }
  var url2 = urlResponse2.getResponseText();
  var externalSpreadsheetId2 = extractSpreadsheetId(url2);

  return [externalSpreadsheetId1, externalSpreadsheetId2];
}

// スタッフのタブを取得する関数
function getStaffTabs(spreadsheetId) {
  try {
    var externalSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    Logger.log(`外部スプレッドシートの読み込みに成功しました: ${spreadsheetId}`);
    return externalSpreadsheet.getSheets();
  } catch (e) {
    Logger.log(`外部スプレッドシートの読み込みに失敗しました: ${e.message}`);
    throw new Error(`外部スプレッドシートの読み込みに失敗しました: ${e.message}`);
  }
}

// インセンティブデータを計算する関数
function calculateIncentiveData(
  data,
  year,
  month,
  holidays,
  benchmarkTimes,
  staffIds,      // ← 追加
  staffTabs1,
  staffTabs2
) {
  // 先頭列に「ID」を追加したヘッダーへ修正
  var incentiveData = [
    [
      'ID',           // 0
      'スタッフ名',   // 1
      '基準時間',     // 2
      '訪問時間',     // 3
      '件数確認',     // 4
      '合計',         // 5
      '時間外訪問',   // 6
      'インセンティブ' // 7
    ]
  ];

  var staffVisitTime = {};
  var staffSpecialTime = {};
  var staffWeekendOrHolidayTime = {};
  var staffOutsideBusinessHoursTime = {};

  for (var i = 1; i < data.length; i++) {
    var { 
      staff1, 
      staff2, 
      visitTimeToAdd, 
      isWeekendOrHolidayVisit, 
      isOutsideBusinessHoursVisit 
    } = processVisitData(data[i], year, month, holidays);

    if (staff1) {
      initializeStaffTimes(staff1, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime);
      updateStaffTimes(staff1, visitTimeToAdd, isWeekendOrHolidayVisit, isOutsideBusinessHoursVisit, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime);
    }

    if (staff2) {
      initializeStaffTimes(staff2, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime);
      updateStaffTimes(staff2, visitTimeToAdd, isWeekendOrHolidayVisit, isOutsideBusinessHoursVisit, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime);
    }
  }

  for (var staff in staffVisitTime) {
    var totalHours = Math.round(staffVisitTime[staff] * 100) / 100;
    var specialHours = Math.round(staffSpecialTime[staff] * 100) / 100;
    var weekendOrHolidayHours = Math.round(staffWeekendOrHolidayTime[staff] * 100) / 100;
    var outsideBusinessHours = Math.round(staffOutsideBusinessHoursTime[staff] * 100) / 100;
    var benchmarkTime = benchmarkTimes[staff] || 0;
    var countCheck = getCountCheck(staff, staffTabs1, staffTabs2);

    var totalTime = totalHours + countCheck;
    var incentive = (totalTime - benchmarkTime) > specialHours
                  ? Math.round((totalTime - benchmarkTime) * 4000)
                  : Math.round(specialHours * 4000);

    if (incentive < 0) incentive = 0; // インセンティブがマイナスにならないようにする

    // ⑤ IDを staffIds[スタッフ名] から取り出して先頭列に入れる
    var staffId = staffIds[staff] || ''; 
    incentiveData.push([
      staffId,         // [0]
      staff,           // [1]
      benchmarkTime,   // [2]
      totalHours,      // [3]
      countCheck,      // [4]
      totalTime,       // [5]
      specialHours,    // [6]
      incentive        // [7]
    ]);
  }

  return incentiveData;
}

// 訪問データを処理する関数
function processVisitData(row, year, month, holidays) {
  var staff1 = row[0].replace(/\s/g, ''); // スペースを削除
  var staff2 = row[2].replace(/\s/g, ''); // スペースを削除
  var day = parseInt(row[9]); 
  var timeStr = row[16].toString();
  var time = parseInt(timeStr);
  var visitDate = new Date(year, month - 1, day);
  var visitTimeStr = row[14];
  var visitTimeParts = visitTimeStr.split(":");
  var visitTime = new Date(year, month - 1, day, parseInt(visitTimeParts[0]), parseInt(visitTimeParts[1]));
  var isWeekendOrHolidayVisit = isWeekendOrHoliday(visitDate, holidays);
  var isOutsideBusinessHoursVisit = isOutsideBusinessHours(visitTime);

  time = convertMinutesToHours(time);
  var visitTimeToAdd = staff2 ? time / 2 : time;

  return { staff1, staff2, visitTimeToAdd, isWeekendOrHolidayVisit, isOutsideBusinessHoursVisit };
}

// スタッフの時間を初期化する関数
function initializeStaffTimes(staff, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime) {
  if (!staffVisitTime[staff]) staffVisitTime[staff] = 0;
  if (!staffSpecialTime[staff]) staffSpecialTime[staff] = 0;
  if (!staffWeekendOrHolidayTime[staff]) staffWeekendOrHolidayTime[staff] = 0;
  if (!staffOutsideBusinessHoursTime[staff]) staffOutsideBusinessHoursTime[staff] = 0;
}

// スタッフの時間を更新する関数
function updateStaffTimes(
  staff,
  visitTimeToAdd,
  isWeekendOrHolidayVisit,
  isOutsideBusinessHoursVisit,
  staffVisitTime,
  staffSpecialTime,
  staffWeekendOrHolidayTime,
  staffOutsideBusinessHoursTime
) {
  staffVisitTime[staff] += visitTimeToAdd;
  if (isWeekendOrHolidayVisit) {
    staffWeekendOrHolidayTime[staff] += visitTimeToAdd;
  }
  if (isOutsideBusinessHoursVisit) {
    staffOutsideBusinessHoursTime[staff] += visitTimeToAdd;
  }
  if (isWeekendOrHolidayVisit || isOutsideBusinessHoursVisit) {
    staffSpecialTime[staff] += visitTimeToAdd;
  }
}

// 件数確認を取得する関数
function getCountCheck(staff, staffTabs1, staffTabs2) {
  var countCheck = 0;
  for (var j = 0; j < staffTabs1.length; j++) {
    if (staffTabs1[j].getName().replace(/\s/g, '') === staff) {
      countCheck += staffTabs1[j].getRange('H40').getValue();
      break;
    }
  }
  for (var k = 0; k < staffTabs2.length; k++) {
    if (staffTabs2[k].getName().replace(/\s/g, '') === staff) {
      countCheck += staffTabs2[k].getRange('H40').getValue();
      break;
    }
  }
  return countCheck;
}

// 分を時間に変換する関数
function convertMinutesToHours(minutes) {
  switch (minutes) {
    case 19:
    case 20:
      return 0.3;
    case 29:
    case 30:
      return 0.6;
    case 40:
      return 0.7;
    case 49:
    case 50:
      return 0.8;
    case 59:
    case 60:
      return 1.0;
    case 79:
    case 80:
      return 1.2;
    case 89:
    case 90:
      return 1.3;
    case 119:
    case 120:
      return 1.6;
    default:
      return 0;
  }
}

// スプレッドシートIDを抽出する関数
function extractSpreadsheetId(url) {
  var matches = url.match(/\/d\/([a-zA-Z0-9-_]+)\//);
  if (matches && matches.length > 1) {
    return matches[1];
  } else {
    throw new Error("スプレッドシートのURLからIDを抽出できませんでした。URLを確認してください。");
  }
}

// 週末または祝日かを確認する関数
function isWeekendOrHoliday(date, holidays) {
  var day = date.getDay();
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return (day == 0 || day == 6 || holidays.includes(formattedDate));
}

// 営業時間外かを確認する関数
function isOutsideBusinessHours(time) {
  var hour = time.getHours();
  var minutes = time.getMinutes();
  // 営業時間を9:00～18:00と定義する場合、18:00以降は外とするならば
  if (hour < 9 || (hour === 18 && minutes > 0) || hour > 18) {
    return true;
  }
  return false;
}

// インセンティブシートの行に交互の背景色を設定する関数
function setRowColors(sheet) {
  var dataRange = sheet.getDataRange();
  var numRows = dataRange.getNumRows();

  // 1行目は濃い色を付けて文字を太字
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground('#d3d3d3').setFontWeight('bold');
  
  // 2行目以降は交互の背景色を設定
  for (var i = 2; i <= numRows; i++) {
    var color = (i % 2 === 0) ? '#ffffff' : '#f0f0f0';  // 偶数行は白、奇数行は薄い灰色
    sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
  }
}
