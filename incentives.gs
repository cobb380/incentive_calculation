// 定数の定義
const HOLIDAY_CALENDAR_ID = 'ja.japanese#holiday@group.v.calendar.google.com';
const INCENTIVE_SHEET_NAME = 'インセンティブ';

// メイン関数
function calculateIncentives() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var defaultSheetName = getSheetName(spreadsheet, '対象のスケジュールシートを記載してください');
  var benchmarkSheetName = getSheetName(spreadsheet, '対象の基準時間シートを選択してください');
  
  var sheet = getSheetByName(spreadsheet, defaultSheetName);
  var incentiveSheet = getOrCreateSheet(spreadsheet, INCENTIVE_SHEET_NAME);
  
  var { year, month } = getInputYearMonth();
  if (!year || !month) return;

  var benchmarkTimes = getBenchmarkTimes(spreadsheet, benchmarkSheetName);
  var holidays = getJapaneseHolidays(year, month);

  var externalSpreadsheetIds = getExternalSpreadsheetIds();
  var staffTabs1 = getStaffTabs(externalSpreadsheetIds[0]);
  var staffTabs2 = getStaffTabs(externalSpreadsheetIds[1]);

  var data = sheet.getDataRange().getValues();
  var incentiveData = calculateIncentiveData(data, year, month, holidays, benchmarkTimes, staffTabs1, staffTabs2);

  incentiveSheet.getRange(1, 1, incentiveData.length, incentiveData[0].length).setValues(incentiveData);

  setRowColors(incentiveSheet);
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
  var yearResponse = ui.prompt('対象年の入力', '年を【半角数字で】入力してください\n (例: 2024)', ui.ButtonSet.OK_CANCEL);
  if (yearResponse.getSelectedButton() != ui.Button.OK) {
    ui.alert('年の入力がキャンセルされました。');
    return {};
  }
  var monthResponse = ui.prompt('対象月の入力', '月を【半角数字で】入力してください\n (例: 1)', ui.ButtonSet.OK_CANCEL);
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

// 基準時間シートからデータを取得する関数
function getBenchmarkTimes(spreadsheet, sheetName) {
  var benchmarkSheet = getSheetByName(spreadsheet, sheetName);
  var benchmarkData = benchmarkSheet.getDataRange().getValues();
  var benchmarkTimes = {};
  for (var i = 1; i < benchmarkData.length; i++) {
    var staffName = benchmarkData[i][0];
    var benchmarkTime = benchmarkData[i][1];
    benchmarkTimes[staffName] = benchmarkTime;
  }
  return benchmarkTimes;
}

// 祝日を取得する関数
function getJapaneseHolidays(year, month) {
  var startDate = new Date(year, month - 1, 1);
  var endDate = new Date(year, month, 0);

  var holidays = CalendarApp.getCalendarById(HOLIDAY_CALENDAR_ID).getEvents(startDate, endDate);
  var holidayDates = holidays.map(function(event) {
    return Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  });

  Logger.log(holidayDates);
  return holidayDates;
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
function calculateIncentiveData(data, year, month, holidays, benchmarkTimes, staffTabs1, staffTabs2) {
  var incentiveData = [['スタッフ名', '基準時間', '訪問時間', '件数確認', '合計', '時間外訪問', 'インセンティブ']];
  var staffVisitTime = {};
  var staffSpecialTime = {};
  var staffWeekendOrHolidayTime = {};
  var staffOutsideBusinessHoursTime = {};

  for (var i = 1; i < data.length; i++) {
    var { staff1, staff2, visitTimeToAdd, isWeekendOrHolidayVisit, isOutsideBusinessHoursVisit } = processVisitData(data[i], year, month, holidays);

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
    var incentive = (totalTime - benchmarkTime) > specialHours ? Math.round((totalTime - benchmarkTime) * 4000) : Math.round(specialHours * 4000);
    if (incentive < 0) incentive = 0; // インセンティブがマイナスにならないようにする

    incentiveData.push([staff, benchmarkTime, totalHours, countCheck, totalTime, specialHours, incentive]);
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
function updateStaffTimes(staff, visitTimeToAdd, isWeekendOrHolidayVisit, isOutsideBusinessHoursVisit, staffVisitTime, staffSpecialTime, staffWeekendOrHolidayTime, staffOutsideBusinessHoursTime) {
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
      return 0.3;
    case 20:
      return 0.3;
    case 29:
      return 0.6;
    case 30:
      return 0.6;
    case 40:
      return 0.7;
    case 49:
      return 0.8;
    case 50:
      return 0.8;
    case 59:
      return 1.0;
    case 60:
      return 1.0;
    case 79:
      return 1.2;
    case 80:
      return 1.2;
    case 89:
      return 1.3;
    case 90:
      return 1.3;
    case 119:
      return 1.6;
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

// 祝日を取得する関数
function getJapaneseHolidays(year, month) {
  var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com'; // 日本の祝日カレンダーID
  var startDate = new Date(year, month - 1, 1);
  var endDate = new Date(year, month, 0); // 指定された月の最終日

  var holidays = CalendarApp.getCalendarById(calendarId).getEvents(startDate, endDate);
  var holidayDates = holidays.map(function(event) {
    return Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  });

  Logger.log(holidayDates);
  return holidayDates;
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
  return (hour < 9 || hour > 18);
}

// インセンティブシートの行に交互の背景色を設定する関数
function setRowColors(sheet) {
  var dataRange = sheet.getDataRange();
  var numRows = dataRange.getNumRows();

  // 1行目は濃い色を付けて文字を太字
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground('#d3d3d3').setFontWeight('bold');
  //2行目以降は交互の背景色を設定
  for (var i = 1; i <= numRows; i++) {
    var color = (i % 2 === 0) ? '#f0f0f0' : '#ffffff';  // 偶数行は灰色、奇数行は白色
    sheet.getRange(i, 1, 1, sheet.getLastColumn()).setBackground(color);
  }
}