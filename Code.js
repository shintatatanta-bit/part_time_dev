/**
 * Webアプリの入口 (GETリクエスト処理)
 */
function doGet(e) {
  var params = e.parameter;
  var page = params.page || "admin";
  var staffId = params.staffId || "";
  
  var template;
  if (page === "staff") {
    template = HtmlService.createTemplateFromFile("staff_view");
    template.staffId = staffId;
    template.staffName = getStaffName(staffId);
  } else {
    template = HtmlService.createTemplateFromFile("admin_view");
  }
  
  return template.evaluate()
    .setTitle("アルバイト勤怠管理システム [v1.1]")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * スタッフIDから名前を引く関数
 */
function getStaffName(staffId) {
  if (!staffId) return "未指定";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("スタッフマスタ");
  if (!sheet) return "スタッフ名不明";
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === staffId) {
      return data[i][1];
    }
  }
  return "登録なしのスタッフ";
}

/**
 * スタッフからの勤怠申請データを保存し、LINEへ通知を送る関数
 */
function saveKintaiData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) {
      throw new Error("「申請データ」シートが見つかりません。");
    }
    
    var requestId = "req_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    
    sheet.appendRow([
      requestId,
      data.staffId,
      data.staffName,
      data.workDate,
      data.startTime,
      data.endTime,
      data.breakTime,
      data.transportCost,
      timestamp
    ]);
    
    sendLineNoticeToStaff(data);
    
    return { success: true, message: "申請が完了し、LINEに控えを通知しました！" };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * スタッフマスタからLINE通知を送る関数
 */
function sendLineNoticeToStaff(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName("スタッフマスタ");
  if (!staffSheet) return;
  
  var staffData = staffSheet.getDataRange().getValues();
  var lineUserId = "";
  var assignedAccount = "";
  
  for (var i = 1; i < staffData.length; i++) {
    if (staffData[i][0] === data.staffId) {
      lineUserId = staffData[i][2];
      assignedAccount = staffData[i][3];
      break;
    }
  }
  
  if (!lineUserId) return;
  
  var accountSheet = ss.getSheetByName("アカウントマスタ");
  var accessToken = "";
  
  if (accountSheet) {
    var accountData = accountSheet.getDataRange().getValues();
    for (var k = 1; k < accountData.length; k++) {
      if (accountData[k][0] === assignedAccount) {
        accessToken = accountData[k][1];
        break;
      }
    }
  }
  
  if (!accessToken || accessToken.includes("YOUR_")) return;
  
  var messageText = 
    "【勤怠申請の控え】\n" +
    "以下の内容で受け付けました。\n\n" +
    "📅 勤務日: " + data.workDate + "\n" +
    "⏰ 時間: " + data.startTime + " ～ " + data.endTime + "\n" +
    "☕ 休憩: " + data.breakTime + "分\n" +
    "🚃 交通費: " + data.transportCost + "円\n\n" +
    "お疲れ様でした！";
  
  var url = "https://api.line.me/v2/bot/message/push";
  var payload = {
    to: lineUserId,
    messages: [{ type: "text", text: messageText }]
  };
  
  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + accessToken
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, options);
}

/**
 * 新規スタッフを登録する関数
 */
function registerNewStaff(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      throw new Error("「スタッフマスタ」シートが見つかりません。");
    }
    
    var staffId = "st_" + Utilities.getUuid().substring(0, 6);
    var timestamp = new Date();
    var webAppUrl = ScriptApp.getService().getUrl();
    var staffUrl = webAppUrl + "?page=staff&staffId=" + staffId;
    
    sheet.appendRow([
      staffId,
      formData.staffName,
      formData.lineUserId,
      formData.assignedAccount,
      "有効",
      timestamp,
      staffUrl
    ]);
    
    return {
      success: true,
      message: "スタッフ「" + formData.staffName + "」さんを登録しました！",
      staffUrl: staffUrl
    };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * 登録されている全スタッフのリストを取得する関数
 */
function getStaffList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    var staffList = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        staffList.push({
          staffId: row[0],
          staffName: row[1],
          lineUserId: row[2],
          assignedAccount: row[3],
          status: row[4],
          staffUrl: row[6] || "",
          displayChar: row[7] || row[1].substring(0, 1)
        });
      }
    }
    return staffList;
  } catch (error) {
    Logger.log("スタッフ一覧取得エラー: " + error.message);
    return [];
  }
}

/**
 * staffIdを指定してスタッフマスタからスタッフを削除する関数
 */
function deleteStaffById(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "スタッフを削除しました" };
      }
    }
    
    return { success: false, message: "削除対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "削除エラー: " + error.message };
  }
}

/**
 * 指定された期間の勤怠データを集計する関数
 */
function getKintaiSummary(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: false, message: "スタッフマスタが見つかりません" };
    var staffData = staffSheet.getDataRange().getValues();
    
    var staffMap = {};
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      if (sId) {
        staffMap[sId] = {
          staffId: sId,
          staffName: staffData[i][1],
          assignedAccount: staffData[i][3],
          daysCount: 0,
          totalCost: 0,
          totalWorkMinutes: 0,
          totalBreakMinutes: 0,
          totalWorkHoursStr: "0時間0分",
          records: []
        };
      }
    }
    
    var dataSheet = ss.getSheetByName("申請データ");
    if (!dataSheet) return { success: false, message: "申請データシートが見つかりません" };
    var kintaiData = dataSheet.getDataRange().getValues();
    
    var start = startDate ? new Date(startDate) : new Date("2000-01-01");
    var end = endDate ? new Date(endDate) : new Date("2100-12-31");
    end.setHours(23, 59, 59, 999);
    
    var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
    
    for (var j = 1; j < kintaiData.length; j++) {
      var row = kintaiData[j];
      var reqId = row[0];
      var sId = row[1];
      var workDateStr = row[3];
      var startTime = row[4];
      var endTime = row[5];
      var breakMins = Number(row[6]) || 0;
      var cost = Number(row[7]) || 0;
      
      if (!reqId || !sId || !workDateStr) continue;
      
      var workDate = new Date(workDateStr);
      if (isNaN(workDate.getTime())) continue;
      
      if (workDate >= start && workDate <= end) {
        if (staffMap[sId]) {
          staffMap[sId].daysCount += 1;
          staffMap[sId].totalCost += cost;
          staffMap[sId].totalBreakMinutes += breakMins;
          
          var formattedStart = formatTime(startTime);
          var formattedEnd = formatTime(endTime);
          var workMins = 0;
          
          if (formattedStart && formattedEnd) {
            var sParts = formattedStart.split(":");
            var eParts = formattedEnd.split(":");
            var startM = Number(sParts[0]) * 60 + Number(sParts[1]);
            var endM = Number(eParts[0]) * 60 + Number(eParts[1]);
            workMins = (endM - startM) - breakMins;
            if (workMins > 0) {
              staffMap[sId].totalWorkMinutes += workMins;
            }
          }
          
          var m = ('0' + (workDate.getMonth() + 1)).slice(-2);
          var d = ('0' + workDate.getDate()).slice(-2);
          var dayOfWeek = weekDays[workDate.getDay()];
          var formattedDate = m + '/' + d + '(' + dayOfWeek + ')';
          
          staffMap[sId].records.push({
            workDate: formattedDate,
            startTime: formattedStart,
            endTime: formattedEnd,
            breakTime: breakMins,
            transportCost: cost,
            workHoursStr: workMins > 0 ? Math.floor(workMins / 60) + "時間" + (workMins % 60) + "分" : "0時間"
          });
        }
      }
    }
    
    var summaryList = [];
    for (var key in staffMap) {
      var item = staffMap[key];
      var totH = Math.floor(item.totalWorkMinutes / 60);
      var totM = item.totalWorkMinutes % 60;
      item.totalWorkHoursStr = totH + "時間" + totM + "分";
      
      var brH = Math.floor(item.totalBreakMinutes / 60);
      var brM = item.totalBreakMinutes % 60;
      item.totalBreakHoursStr = brH > 0 ? brH + "時間" + brM + "分" : brM + "分";
      
      item.records.sort(function(a, b) {
        return new Date(a.workDate) - new Date(b.workDate);
      });
      summaryList.push(item);
    }
    
    return { success: true, summary: summaryList };
  } catch (error) {
    return { success: false, message: "集計エラー: " + error.message };
  }
}

/**
 * 時刻データを "HH:mm" 形式に変換
 */
function formatTime(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var h = ('0' + val.getHours()).slice(-2);
    var m = ('0' + val.getMinutes()).slice(-2);
    return h + ':' + m;
  }
  var str = String(val);
  if (str.indexOf("T") !== -1 || str.indexOf("1899-") !== -1) {
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      var h = ('0' + d.getHours()).slice(-2);
      var m = ('0' + d.getMinutes()).slice(-2);
      return h + ':' + m;
    }
  }
  if (str.length >= 5) {
    return str.substring(0, 5);
  }
  return str;
}

/**
 * LINEアカウント一覧を取得
 */
function getAccountList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return ["アカウントA", "アカウントB"];
    
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) accounts.push(data[i][0]);
    }
    return accounts.length > 0 ? accounts : ["アカウントA", "アカウントB"];
  } catch (error) {
    return ["アカウントA", "アカウントB"];
  }
}

/**
 * シフト予定を取得
 */
function getShiftSchedule(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) return { success: true, shifts: [] };
    
    var data = sheet.getDataRange().getValues();
    var shifts = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var shiftId = row[0];
      var workDate = row[1];
      var staffId = row[2];
      var staffName = row[3];
      
      if (!workDate || !staffId) continue;
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (dateStr.indexOf(yearMonth) === 0) {
        shifts.push({ shiftId: shiftId, workDate: dateStr, staffId: staffId, staffName: staffName });
      }
    }
    return { success: true, shifts: shifts };
  } catch (error) {
    return { success: false, message: "シフト取得エラー: " + error.message };
  }
}

/**
 * 日本の祝日マップを取得（計算ロジック内蔵）
 */
function getJapaneseHolidays(yearMonth) {
  try {
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    var holidayMap = {};
    
    var fixedHolidays = {
      "1-1": "元日", "2-11": "建国記念の日", "2-23": "天皇誕生日",
      "4-29": "昭和の日", "5-3": "憲法記念日", "5-4": "みどりの日",
      "5-5": "こどもの日", "8-11": "山の日", "11-3": "文化の日", "11-23": "勤労感謝の日"
    };

    var lastDay = new Date(year, month, 0).getDate();
    for (var day = 1; day <= lastDay; day++) {
      var d = new Date(year, month - 1, day);
      var w = d.getDay();
      var m = month;
      
      var key = year + '-' + ('0' + m).slice(-2) + '-' + ('0' + day).slice(-2);
      var hName = "";
      
      var fixKey = m + "-" + day;
      if (fixedHolidays[fixKey]) hName = fixedHolidays[fixKey];
      
      if (m === 1 && w === 1 && day >= 8 && day <= 14) hName = "成人の日";
      if (m === 4 && day === 29) hName = "昭和の日";
      if (m === 7 && w === 1 && day >= 15 && day <= 21) hName = "海の日";
      if (m === 8 && day === 11) hName = "山の日";
      if (m === 9 && w === 1 && day >= 15 && day <= 21) hName = "敬老の日";
      if (m === 10 && w === 1 && day >= 8 && day <= 14) hName = "スポーツの日";
      if (m === 11 && day === 3) hName = "文化の日";
      if (m === 11 && day === 23) hName = "勤労感謝の日";
      
      if (m === 3 && ((year === 2026 && day === 20) || (year !== 2026 && day === 21))) hName = "春分の日";
      if (m === 9 && ((year === 2026 && day === 23) || (year !== 2026 && day === 23))) hName = "秋分の日";

      if (hName) holidayMap[key] = hName;
    }
    return { success: true, holidays: holidayMap };
  } catch (error) {
    return { success: true, holidays: {} };
  }
}

/**
 * 休業日を取得
 */
function getClosedDays(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("休業日");
    if (!sheet) return { success: true, closedDays: [] };
    
    var data = sheet.getDataRange().getValues();
    var closedDays = [];
    
    for (var i = 1; i < data.length; i++) {
      var workDate = data[i][1];
      if (!workDate) continue;
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (dateStr.indexOf(yearMonth) === 0) {
        closedDays.push(dateStr);
      }
    }
    return { success: true, closedDays: closedDays };
  } catch (error) {
    return { success: false, closedDays: [], message: error.message };
  }
}

/**
 * シフト予定と休業日をまとめて一括保存する関数
 */
function saveAllSchedules(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth;
    
    // 1. シフト予定の保存
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) {
      shiftSheet = ss.insertSheet("シフト予定");
      shiftSheet.appendRow(["shiftId", "workDate", "staffId", "staffName"]);
    }
    
    var shiftData = shiftSheet.getDataRange().getValues();
    var newShiftRows = [shiftData[0]];
    
    for (var i = 1; i < shiftData.length; i++) {
      var rDate = new Date(shiftData[i][1]);
      if (isNaN(rDate.getTime())) continue;
      var rYm = rDate.getFullYear() + '-' + ('0' + (rDate.getMonth() + 1)).slice(-2);
      if (rYm !== yearMonth) {
        newShiftRows.push(shiftData[i]);
      }
    }
    
    (data.shifts || []).forEach(function(sh) {
      var sId = "shift_" + Utilities.getUuid().substring(0, 8);
      newShiftRows.push([sId, sh.workDate, sh.staffId, sh.staffName]);
    });
    
    shiftSheet.clearContents();
    if (newShiftRows.length > 0) {
      shiftSheet.getRange(1, 1, newShiftRows.length, newShiftRows[0].length).setValues(newShiftRows);
    }
    
    // 2. 休業日の保存
    var holSheet = ss.getSheetByName("休業日");
    if (!holSheet) {
      holSheet = ss.insertSheet("休業日");
      holSheet.appendRow(["holidayId", "workDate"]);
    }
    
    var holData = holSheet.getDataRange().getValues();
    var newHolRows = [holData[0]];
    
    for (var j = 1; j < holData.length; j++) {
      var hDate = new Date(holData[j][1]);
      if (isNaN(hDate.getTime())) continue;
      var hYm = hDate.getFullYear() + '-' + ('0' + (hDate.getMonth() + 1)).slice(-2);
      if (hYm !== yearMonth) {
        newHolRows.push(holData[j]);
      }
    }
    
    (data.closedDays || []).forEach(function(cDate) {
      var hId = "hol_" + Utilities.getUuid().substring(0, 8);
      newHolRows.push([hId, cDate]);
    });
    
    holSheet.clearContents();
    if (newHolRows.length > 0) {
      holSheet.getRange(1, 1, newHolRows.length, newHolRows[0].length).setValues(newHolRows);
    }
    
    return { success: true, message: "変更内容をすべて保存しました！" };
  } catch (error) {
    return { success: false, message: "保存エラー: " + error.message };
  }
}

/**
 * スタッフの「カレンダー表示文字」を更新する関数
 */
function updateStaffDisplayChar(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.staffId)) {
        // H列（インデックス7、8列目）に表示文字を書き込む
        sheet.getRange(i + 1, 8).setValue(data.displayChar);
        return { success: true, message: "表示文字を更新しました！" };
      }
    }
    
    return { success: false, message: "対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "更新エラー: " + error.message };
  }
}

/**
 * 指定された年月のフリー予定を取得する関数
 */
function getFreeSchedules(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("フリー予定");
    if (!sheet) {
      return { success: true, freeSchedules: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var freeSchedules = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var scheduleId = row[0];
      var workDate = row[1];
      var title = row[2];
      
      if (!workDate || !title) continue;
      
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      if (dateStr.indexOf(yearMonth) === 0) {
        freeSchedules.push({
          scheduleId: scheduleId,
          workDate: dateStr,
          title: title
        });
      }
    }
    
    return { success: true, freeSchedules: freeSchedules };
  } catch (error) {
    return { success: false, message: "フリー予定取得エラー: " + error.message };
  }
}

/**
 * シフト予定、休業日、フリー予定をまとめて一括保存する関数（従来のsaveAllSchedulesを拡張）
 */
function saveAllSchedulesWithFree(data) {
  try {
    // まず既存のシフトと休業日の保存関数を実行
    var res = saveAllSchedules(data);
    if (!res.success) return res;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth;
    
    // フリー予定の保存処理
    var freeSheet = ss.getSheetByName("フリー予定");
    if (!freeSheet) {
      freeSheet = ss.insertSheet("フリー予定");
      freeSheet.appendRow(["scheduleId", "workDate", "title"]);
    }
    
    var freeData = freeSheet.getDataRange().getValues();
    var newFreeRows = [freeData[0]]; // ヘッダー保持
    
    for (var i = 1; i < freeData.length; i++) {
      var fDate = new Date(freeData[i][1]);
      if (isNaN(fDate.getTime())) continue;
      var fYm = fDate.getFullYear() + '-' + ('0' + (fDate.getMonth() + 1)).slice(-2);
      if (fYm !== yearMonth) {
        newFreeRows.push(freeData[i]); // 他月のデータは保持
      }
    }
    
    (data.freeSchedules || []).forEach(function(fs) {
      var sId = "free_" + Utilities.getUuid().substring(0, 8);
      newFreeRows.push([sId, fs.workDate, fs.title]);
    });
    
    freeSheet.clearContents();
    if (newFreeRows.length > 0) {
      freeSheet.getRange(1, 1, newFreeRows.length, newFreeRows[0].length).setValues(newFreeRows);
    }
    
    return { success: true, message: "変更内容（予定・フリー項目含む）をすべて保存しました！" };
  } catch (error) {
    return { success: false, message: "一括保存エラー: " + error.message };
  }
}

/**
 * スタッフのステータス（有効/退職）を切り替える関数
 */
function toggleStaffStatus(staffId, newStatus) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      return { success: false, message: "「スタッフマスタ」シートが見つかりません。" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        // E列（インデックス4、5列目）にステータスを書き込む
        sheet.getRange(i + 1, 5).setValue(newStatus);
        return { success: true, message: "スタッフのステータスを「" + newStatus + "」に変更しました！" };
      }
    }
    
    return { success: false, message: "対象のスタッフが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "ステータス変更エラー: " + error.message };
  }
}

/**
 * 💡【変更】スタッフからのシフト希望を保存し、管理者向けの新着通知（未確認ログ）を記録する関数
 */
function saveShiftPreferences(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. シフト予定シートへの登録（従来通り）
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      sheet = ss.insertSheet("シフト予定");
      sheet.appendRow(["shiftId", "workDate", "staffId", "staffName"]);
    }
    
    var existingData = sheet.getDataRange().getValues();
    var staffId = data.staffId;
    var staffName = data.staffName;
    var targetDates = data.selectedDates || [];
    
    var newRows = [existingData[0]]; 
    for (var i = 1; i < existingData.length; i++) {
      var row = existingData[i];
      var rDateStr = String(row[1] || "");
      var rStaffId = String(row[2] || "");
      
      var isSameMonth = rDateStr.indexOf(data.yearMonth) === 0;
      if (!(isSameMonth && rStaffId === staffId)) {
        newRows.push(row);
      }
    }
    
    targetDates.forEach(function(dStr) {
      var shiftId = "shift_" + Utilities.getUuid().substring(0, 8);
      newRows.push([shiftId, dStr, staffId, staffName]);
    });
    
    sheet.clearContents();
    if (newRows.length > 0) {
      sheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    // 💡 2. 管理者向け「新着希望通知」シートへ未読ログを記録
    var notifSheet = ss.getSheetByName("希望通知ログ");
    if (!notifSheet) {
      notifSheet = ss.insertSheet("希望通知ログ");
      notifSheet.appendRow(["notifId", "staffId", "staffName", "yearMonth", "timestamp", "status"]);
    }
    
    var notifId = "notif_" + Utilities.getUuid().substring(0, 8);
    notifSheet.appendRow([
      notifId,
      staffId,
      staffName,
      data.yearMonth,
      new Date(),
      "未読"
    ]);
    
    return { success: true, message: "シフト希望を提出し、予定に反映しました！" };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
  }
}

/**
 * 指定された年月の全スタッフからのシフト希望を取得し、そのまま確定シフトとして扱う関数
 */
function getShiftPreferences(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("希望シフト");
    if (!sheet) {
      return { success: true, preferences: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var preferences = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var prefId = row[0];   
      var staffId = row[1];  
      var staffName = row[2];
      var ym = String(row[3] || "").trim();      
      var datesStr = String(row[4] || "").trim(); 
      
      if (ym === yearMonth || ym.indexOf(yearMonth) !== -1 || datesStr.indexOf(yearMonth) !== -1) {
        var datesArray = [];
        if (datesStr) {
          datesStr.split(",").forEach(function(d) {
            var trimmed = d.trim();
            if (trimmed) {
              datesArray.push(trimmed);
            }
          });
        }
        
        if (datesArray.length > 0) {
          preferences.push({
            prefId: prefId,
            staffId: staffId,
            staffName: staffName,
            yearMonth: yearMonth,
            dates: datesArray
          });
        }
      }
    }
    
    return { success: true, preferences: preferences };
  } catch (error) {
    return { success: false, message: "希望シフト取得エラー: " + error.message };
  }
}

/**
 * カレンダー上で削除されたシフト希望、または確定シフトをシートから削除する関数
 */
function deleteScheduleOrPreference(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetDate = data.workDate;
    var staffId = data.staffId;
    
    // 1. 「希望シフト」シートからの削除チェック
    var prefSheet = ss.getSheetByName("希望シフト");
    if (prefSheet) {
      var prefData = prefSheet.getDataRange().getValues();
      for (var i = 0; i < prefData.length; i++) {
        var row = prefData[i];
        if (row[1] === staffId) { // B列がstaffId
          var datesStr = String(row[4] || ""); // E列がカンマ区切りの日付
          if (datesStr.indexOf(targetDate) !== -1) {
            var datesArr = datesStr.split(",").map(function(d){ return d.trim(); });
            var newDatesArr = datesArr.filter(function(d){ return d !== targetDate; });
            
            if (newDatesArr.length > 0) {
              // まだ他の日の希望が残っている場合は、該当日の日付を抜いた文字列に更新
              prefSheet.getRange(i + 1, 5).setValue(newDatesArr.join(","));
            } else {
              // その日の希望しかなかった場合は行ごと削除
              prefSheet.deleteRow(i + 1);
            }
          }
        }
      }
    }
    
    // 2. 「シフト予定」シートからの削除チェック（もし確定シフト側を消した場合）
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (shiftSheet) {
      var shiftData = shiftSheet.getDataRange().getValues();
      for (var j = 1; j < shiftData.length; j++) {
        var sRow = shiftData[j];
        var sDate = new Date(sRow[1]);
        var sDateStr = isNaN(sDate.getTime()) ? String(sRow[1]) : Utilities.formatDate(sDate, "Asia/Tokyo", "yyyy-MM-dd");
        if (sDateStr === targetDate && sRow[2] === staffId) {
          shiftSheet.deleteRow(j + 1);
          break;
        }
      }
    }
    
    return { success: true, message: "データを削除しました" };
  } catch (error) {
    return { success: false, message: "削除エラー: " + error.message };
  }
}

/**
 * 互換性のために空の配列を返すダミー関数
 */
function getShiftPreferences(yearMonth) {
  return { success: true, preferences: [] };
}

/**
 * 互換性のために空の配列を返すダミー関数
 */
function getShiftPreferences(yearMonth) {
  return { success: true, preferences: [] };
}

/**
 * 💡【追加】LINEメッセージを一斉または個別に送信し、履歴を保存する関数
 */
function sendLineMessage(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) throw new Error("スタッフマスタが見つかりません。");
    var staffData = staffSheet.getDataRange().getValues();

    var targetStaff = [];
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      var sName = staffData[i][1];
      var lineId = staffData[i][2];
      var account = staffData[i][3];
      var status = staffData[i][4];

      if (sId && lineId && status === "有効") {
        if (data.targetType === "all" || data.targetId === sId) {
          targetStaff.push({ staffId: sId, staffName: sName, lineUserId: lineId, assignedAccount: account });
        }
      }
    }

    if (targetStaff.length === 0) {
      return { success: false, message: "送信対象の有効なスタッフが見つかりませんでした。" };
    }

    var accountSheet = ss.getSheetByName("アカウントマスタ");
    var tokens = {};
    if (accountSheet) {
      var accData = accountSheet.getDataRange().getValues();
      for (var k = 1; k < accData.length; k++) {
        if (accData[k][0]) tokens[accData[k][0]] = accData[k][1];
      }
    }

    var successCount = 0;
    var failCount = 0;

    targetStaff.forEach(function(staff) {
      var token = tokens[staff.assignedAccount];
      if (!token || token.includes("YOUR_")) {
        failCount++;
        return;
      }

      var url = "https://api.line.me/v2/bot/message/push";
      var payload = {
        to: staff.lineUserId,
        messages: [{ type: "text", text: data.messageText }]
      };

      var options = {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        var res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() === 200) {
          successCount++;
        } else {
          failCount++;
        }
      } catch (e) {
        failCount++;
      }
    });

    var histSheet = ss.getSheetByName("連絡履歴");
    if (!histSheet) {
      histSheet = ss.insertSheet("連絡履歴");
      histSheet.appendRow(["historyId", "timestamp", "targetName", "messageText", "resultSummary"]);
    }

    var historyId = "hist_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    var targetNameDisplay = (data.targetType === "all") ? "【一斉送信】全スタッフ" : (targetStaff[0] ? targetStaff[0].staffName : "個別");
    var resultStr = "成功: " + successCount + "件 / 失敗: " + failCount + "件";

    histSheet.appendRow([historyId, timestamp, targetNameDisplay, data.messageText, resultStr]);

    return { 
      success: true, 
      message: "メッセージを送信しました！（" + resultStr + "）" 
    };

  } catch (error) {
    return { success: false, message: "送信エラー: " + error.message };
  }
}

/**
 * 💡【追加】送信履歴を取得する関数
 */
function getLineHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("連絡履歴");
    if (!sheet) return { success: true, history: [] };

    var data = sheet.getDataRange().getValues();
    var history = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var timeStr = "";
        if (row[1] instanceof Date) {
          timeStr = Utilities.formatDate(row[1], "Asia/Tokyo", "yyyy/MM/dd HH:mm");
        } else {
          timeStr = String(row[1]);
        }

        history.push({
          timestamp: timeStr,
          targetName: row[2],
          messageText: row[3],
          resultSummary: row[4]
        });
      }
    }
    history.reverse();
    return { success: true, history: history };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

/**
 * 💡【変更】「過去1か月間（今日から30日前まで）」の勤怠未提出者をリストアップする関数
 */
function getKintaiPendingStaff() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 今日から30日前（約1か月間）の範囲を設定
    var today = new Date();
    var startDate = new Date();
    startDate.setDate(today.getDate() - 30);
    
    var todayStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy-MM-dd");
    var startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd");

    // 1. 過去1か月間の「シフト予定」をすべて取得
    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) return { success: true, pendingList: [] };
    var shiftData = shiftSheet.getDataRange().getValues();
    
    // キー: "staffId_yyyy-MM-dd", 値: { staffId, staffName, dateStr }
    var scheduledMap = {}; 
    for (var i = 1; i < shiftData.length; i++) {
      var row = shiftData[i];
      var wDate = row[1];
      var sId = row[2];
      var sName = row[3];
      if (!wDate || !sId) continue;
      
      var d = new Date(wDate);
      if (isNaN(d.getTime())) continue;
      var dStr = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
      
      // 30日前から今日までの範囲内のみ対象
      if (dStr >= startStr && dStr <= todayStr) {
        var key = sId + "_" + dStr;
        scheduledMap[key] = {
          staffId: sId,
          staffName: sName,
          targetDate: dStr
        };
      }
    }

    // 2. 過去1か月間の「申請データ」をすべて取得
    var dataSheet = ss.getSheetByName("申請データ");
    if (dataSheet) {
      var kintaiData = dataSheet.getDataRange().getValues();
      for (var j = 1; j < kintaiData.length; j++) {
        var kRow = kintaiData[j];
        var kStaffId = kRow[1];
        var kDate = kRow[3];
        if (!kStaffId || !kDate) continue;
        
        var kd = new Date(kDate);
        if (isNaN(kd.getTime())) continue;
        var kdStr = Utilities.formatDate(kd, "Asia/Tokyo", "yyyy-MM-dd");
        
        // すでに申請があるものは予定マップから削除（＝未提出者だけに絞る）
        var submittedKey = kStaffId + "_" + kdStr;
        if (scheduledMap[submittedKey]) {
          delete scheduledMap[submittedKey];
        }
      }
    }

    // 3. 残った未提出の予定を配列に変換（日付の新しい順に並び替え）
    var pendingList = [];
    for (var key in scheduledMap) {
      pendingList.push(scheduledMap[key]);
    }
    
    pendingList.sort(function(a, b) {
      return new Date(b.targetDate) - new Date(a.targetDate);
    });

    return { success: true, pendingList: pendingList };
  } catch (error) {
    return { success: false, message: "未提出者取得エラー: " + error.message, pendingList: [] };
  }
}

/**
 * 💡【追加】未提出者リストをLINEリマインド用に詳細情報つきで取得する関数
 */
function getKintaiPendingForReminder() {
  try {
    // 既存の未提出者取得ロジックを利用
    var res = getKintaiPendingStaff();
    if (!res.success) return res;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: true, pendingList: [] };
    var staffData = staffSheet.getDataRange().getValues();
    
    // staffIdごとのLINEの有効性をマップ化
    var staffLineMap = {};
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      var lineId = staffData[i][2];
      var status = staffData[i][4];
      if (sId && lineId && status === "有効") {
        staffLineMap[sId] = true;
      }
    }

    // LINE連携が有効な未提出者だけに絞り込み
    var validPendingList = [];
    res.pendingList.forEach(function(item) {
      if (staffLineMap[item.staffId]) {
        validPendingList.push(item);
      }
    });

    return { success: true, pendingList: validPendingList };
  } catch (error) {
    return { success: false, message: "リマインド対象取得エラー: " + error.message, pendingList: [] };
  }
}

/**
 * 💡【追加】未提出者へ個別（または一斉）にリマインドLINEを送信する関数
 */
function sendKintaiReminderLine(data) {
  try {
    var targets = data.targets || []; // 送信対象の staffId の配列
    var messageText = data.messageText;

    if (targets.length === 0) {
      return { success: false, message: "送信対象のスタッフが選択されていません。" };
    }
    if (!messageText) {
      return { success: false, message: "メッセージ本文が入力されていません。" };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) throw new Error("スタッフマスタが見つかりません。");
    var staffData = staffSheet.getDataRange().getValues();

    var staffMap = {};
    for (var i = 1; i < staffData.length; i++) {
      if (staffData[i][0]) {
        staffMap[staffData[i][0]] = {
          staffName: staffData[i][1],
          lineUserId: staffData[i][2],
          assignedAccount: staffData[i][3]
        };
      }
    }

    var accountSheet = ss.getSheetByName("アカウントマスタ");
    var tokens = {};
    if (accountSheet) {
      var accData = accountSheet.getDataRange().getValues();
      for (var k = 1; k < accData.length; k++) {
        if (accData[k][0]) tokens[accData[k][0]] = accData[k][1];
      }
    }

    var successCount = 0;
    var failCount = 0;
    var targetNames = [];

    targets.forEach(function(staffId) {
      var staff = staffMap[staffId];
      if (!staff || !staff.lineUserId) {
        failCount++;
        return;
      }

      var token = tokens[staff.assignedAccount];
      if (!token || token.includes("YOUR_")) {
        failCount++;
        return;
      }

      targetNames.push(staff.staffName);

      var url = "https://api.line.me/v2/bot/message/push";
      var payload = {
        to: staff.lineUserId,
        messages: [{ type: "text", text: messageText }]
      };

      var options = {
        method: "post",
        headers: {
          "Content-Type": "application/json; charset=UTF-8",
          "Authorization": "Bearer " + token
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        var res = UrlFetchApp.fetch(url, options);
        if (res.getResponseCode() === 200) {
          successCount++;
        } else {
          failCount++;
        }
      } catch (e) {
        failCount++;
      }
    });

    // 履歴シートへ保存
    var histSheet = ss.getSheetByName("連絡履歴");
    if (!histSheet) {
      histSheet = ss.insertSheet("連絡履歴");
      histSheet.appendRow(["historyId", "timestamp", "targetName", "messageText", "resultSummary"]);
    }

    var historyId = "hist_" + Utilities.getUuid().substring(0, 8);
    var timestamp = new Date();
    var targetNameDisplay = "【勤怠リマインド】" + targetNames.join(", ");
    var resultStr = "成功: " + successCount + "件 / 失敗: " + failCount + "件";

    histSheet.appendRow([historyId, timestamp, targetNameDisplay, messageText, resultStr]);

    return { 
      success: true, 
      message: "リマインドを送信しました！（" + resultStr + "）" 
    };

  } catch (error) {
    return { success: false, message: "リマインド送信エラー: " + error.message };
  }
}

/**
 * 💡【完全安全版】管理者向けに未読の「シフト希望申請通知」を取得する関数
 */
function getUnreadShiftNotifications() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: true, notifications: [] };
    
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: true, notifications: [] };

    var data = sheet.getDataRange().getValues();
    if (!data || data.length <= 1) return { success: true, notifications: [] };

    var notifications = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var notifId = String(row[0] || "").trim();
      var staffName = String(row[2] || "").trim();
      var yearMonth = String(row[3] || "").trim();
      var timeObj = row[4];
      var status = String(row[5] || "").trim();

      if (notifId && status === "未読") {
        var timeStr = "";
        if (timeObj instanceof Date) {
          timeStr = Utilities.formatDate(timeObj, "Asia/Tokyo", "MM/dd HH:mm");
        } else {
          timeStr = String(timeObj);
        }

        notifications.push({
          notifId: notifId,
          staffName: staffName || "スタッフ",
          yearMonth: yearMonth || "未指定",
          timestamp: timeStr
        });
      }
    }

    notifications.reverse();
    return { success: true, notifications: notifications };
  } catch (error) {
    // 万が一エラーが起きても絶対におかしな値を返さず、空の配列を返すようにする
    return { success: true, notifications: [], message: error.message };
  }
}

/**
 * 💡【追加】シフト希望通知を「既読（確認済み）」にする関数
 */
function markNotificationAsRead(notifId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: false, message: "通知シートが見つかりません" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(notifId)) {
        sheet.getRange(i + 1, 6).setValue("既読"); // F列（6列目）の状態を既読に
        return { success: true, message: "確認済みにしました" };
      }
    }
    return { success: false, message: "該当の通知が見つかりませんでした" };
  } catch (error) {
    return { success: false, message: "エラー: " + error.message };
  }
}