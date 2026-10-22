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
    .setTitle("アルバイト勤怠管理システム [v1.1 - 10/22更新]") // ←ここに更新が分かる目印を入れました！
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
 * 登録されている全スタッフのリストを取得する関数（重複を整理し、最新版に一本化）
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
      if (row[0]) { // スタッフIDが存在するもの
        staffList.push({
          staffId: row[0],
          staffName: row[1],
          lineUserId: row[2],
          assignedAccount: row[3],
          status: row[4],
          staffUrl: row[6] || "",
          displayChar: row[7] || row[1].substring(0, 1) // H列(index 7)の文字。空欄なら名前の頭文字
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
 * 指定された期間（開始日〜終了日）の勤怠データをスタッフごとに集計し、合計時間・休憩・日別明細を返す関数
 */
function getKintaiSummary(startDate, endDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. スタッフマスタから全スタッフをマップ化
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
    
    // 2. 申請データシートから該当期間のデータを集計
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
      var workDateStr = row[3]; // 勤務日
      var startTime = row[4];   // 開始時刻
      var endTime = row[5];     // 終了時刻
      var breakMins = Number(row[6]) || 0; // 休憩分
      var cost = Number(row[7]) || 0;     // 交通費
      
      if (!reqId || !sId || !workDateStr) continue;
      
      var workDate = new Date(workDateStr);
      if (isNaN(workDate.getTime())) continue;
      
      if (workDate >= start && workDate <= end) {
        if (staffMap[sId]) {
          staffMap[sId].daysCount += 1;
          staffMap[sId].totalCost += cost;
          staffMap[sId].totalBreakMinutes += breakMins;
          
          // 実働時間の計算（分単位）
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
          
          // 月/日 と 曜日を取得 (例: 10/20(火))
          var m = ('0' + (workDate.getMonth() + 1)).slice(-2);
          var d = ('0' + workDate.getDate()).slice(-2);
          var dayOfWeek = weekDays[workDate.getDay()];
          var formattedDate = m + '/' + d + '(' + dayOfWeek + ')';
          
          // 日別明細データを追加
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
    
    // 3. 合計分を変換し、日付順にソート
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
 * 時刻データを "HH:mm" 形式の文字列に変換する補助関数
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
 * LINEアカウントの一覧を取得する関数
 */
function getAccountList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) return ["アカウントA", "アカウントB"];
    
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    
    for (var i = 1; i < data.length; i++) {
      var accName = data[i][0];
      if (accName) {
        accounts.push(accName);
      }
    }
    
    return accounts.length > 0 ? accounts : ["アカウントA", "アカウントB"];
  } catch (error) {
    return ["アカウントA", "アカウントB"];
  }
}

/**
 * 指定された年月のシフト予定を一括取得する関数
 */
function getShiftSchedule(yearMonth) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      return { success: true, shifts: [] };
    }
    
    var data = sheet.getDataRange().getValues();
    var shifts = [];
    
    // 2行目からループ
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var shiftId = row[0];
      var workDate = row[1];
      var staffId = row[2];
      var staffName = row[3];
      
      if (!workDate || !staffId) continue;
      
      // 日付を "yyyy-MM-dd" 形式の文字列に変換
      var d = new Date(workDate);
      if (isNaN(d.getTime())) continue;
      
      var y = d.getFullYear();
      var m = ('0' + (d.getMonth() + 1)).slice(-2);
      var day = ('0' + d.getDate()).slice(-2);
      var dateStr = y + '-' + m + '-' + day;
      
      // 指定された年月（例: "2026-10"）に一致するものだけを抽出
      if (dateStr.indexOf(yearMonth) === 0) {
        shifts.push({
          shiftId: shiftId,
          workDate: dateStr,
          staffId: staffId,
          staffName: staffName
        });
      }
    }
    
    return { success: true, shifts: shifts };
  } catch (error) {
    return { success: false, message: "シフト取得エラー: " + error.message };
  }
}

/**
 * シフト予定を保存（配置）する関数
 */
function saveShift(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      throw new Error("「シフト予定」シートが見つかりません。作成してください。");
    }
    
    // すでに同じ日・同じスタッフのシフトが登録されているかチェック（重複登録を防ぐ）
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var existingDate = new Date(values[i][1]);
      var targetDate = new Date(data.workDate);
      
      if (
        existingDate.getFullYear() === targetDate.getFullYear() &&
        existingDate.getMonth() === targetDate.getMonth() &&
        existingDate.getDate() === targetDate.getDate() &&
        values[i][2] === data.staffId
      ) {
        // すでに存在する場合は何もしない（またはエラーを返す）
        return { success: true, message: "すでに配置されています" };
      }
    }
    
    // 新規登録
    var shiftId = "shift_" + Utilities.getUuid().substring(0, 8);
    sheet.appendRow([
      shiftId,
      data.workDate,
      data.staffId,
      data.staffName
    ]);
    
    return { success: true, message: "シフトを配置しました！" };
  } catch (error) {
    return { success: false, message: "保存エラー: " + error.message };
  }
}

/**
 * 複数のシフト予定を一括で保存する関数（毎週一括配置用）
 */
function saveShiftsBatch(shifts) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      throw new Error("「シフト予定」シートが見つかりません。");
    }
    
    var values = sheet.getDataRange().getValues();
    
    shifts.forEach(function(data) {
      var targetDate = new Date(data.workDate);
      if (isNaN(targetDate.getTime())) return;
      
      // 重複チェック
      var exists = false;
      for (var i = 1; i < values.length; i++) {
        var existingDate = new Date(values[i][1]);
        if (
          existingDate.getFullYear() === targetDate.getFullYear() &&
          existingDate.getMonth() === targetDate.getMonth() &&
          existingDate.getDate() === targetDate.getDate() &&
          values[i][2] === data.staffId
        ) {
          exists = true;
          break;
        }
      }
      
      // 存在しない場合のみ追加
      if (!exists) {
        var shiftId = "shift_" + Utilities.getUuid().substring(0, 8);
        sheet.appendRow([
          shiftId,
          data.workDate,
          data.staffId,
          data.staffName
        ]);
        // 処理中の重複を防ぐために配列にも追加
        values.push([shiftId, data.workDate, data.staffId, data.staffName]);
      }
    });
    
    return { success: true, message: "シフトを配置しました！" };
  } catch (error) {
    return { success: false, message: "一括保存エラー: " + error.message };
  }
}

/**
 * shiftIdを指定してシフト予定を削除する関数
 */
function deleteShiftById(shiftId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      return { success: false, message: "「シフト予定」シートが見つかりません。" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(shiftId)) {
        sheet.deleteRow(i + 1); // 2行目がスプレッドシートの行番号 2 (配列インデックス+1)
        return { success: true, message: "シフトを削除しました" };
      }
    }
    
    return { success: false, message: "削除対象のシフトが見つかりませんでした。" };
  } catch (error) {
    return { success: false, message: "削除エラー: " + error.message };
  }
}

/**
 * 指定された年月の日本の祝日マップを取得する関数
 */
/**
 * 指定された年月の日本の祝日マップを確実に取得する関数（計算ロジック内蔵）
 */
function getJapaneseHolidays(yearMonth) {
  try {
    var parts = yearMonth.split('-');
    var year = Number(parts[0]);
    var month = Number(parts[1]);
    
    var holidayMap = {};
    
    // 1. 固定祝日（月/日）
    var fixedHolidays = {
      "1-1": "元日",
      "2-11": "建国記念の日",
      "2-23": "天皇誕生日",
      "4-29": "昭和の日",
      "5-3": "憲法記念日",
      "5-4": "みどりの日",
      "5-5": "こどもの日",
      "7-21": "海の日", // ※2026年等の日付対応のため下記で動的計算も併用
      "8-11": "山の日",
      "9-15": "敬老の日",
      "11-3": "文化の日",
      "11-23": "勤労感謝の日"
    };

    // 2. その月の何日目が祝日か、あるいはハッピーマンデー・春分/秋分の日の簡易計算
    // 指定された年月の1日〜末日までをループして判定
    var lastDay = new Date(year, month, 0).getDate();
    
    for (var day = 1; day <= lastDay; day++) {
      var d = new Date(year, month - 1, day);
      var w = d.getDay(); // 0:日, 1:月...6:土
      var m = month;
      
      var key = year + '-' + ('0' + m).slice(-2) + '-' + ('0' + day).slice(-2);
      var hName = "";
      
      // 固定祝日のチェック
      var fixKey = m + "-" + day;
      if (fixedHolidays[fixKey]) {
        hName = fixedHolidays[fixKey];
      }
      
      // 成人の日（1月第2月曜）
      if (m === 1 && w === 1 && day >= 8 && day <= 14) hName = "成人の日";
      // 昭和の日振替など
      if (m === 4 && day === 29) hName = "昭和の日";
      // 海の日（7月第3月曜 - 2026年例）
      if (m === 7 && w === 1 && day >= 15 && day <= 21) hName = "海の日";
      // 山の日（8月11日）
      if (m === 8 && day === 11) hName = "山の日";
      // 敬老の日（9月第3月曜）
      if (m === 9 && w === 1 && day >= 15 && day <= 21) hName = "敬老の日";
      // スポーツの日（10月第2月曜）
      if (m === 10 && w === 1 && day >= 8 && day <= 14) hName = "スポーツの日";
      // 文化の日（11月3日）
      if (m === 11 && day === 3) hName = "文化の日";
      // 勤労感謝の日（11月23日）
      if (m === 11 && day === 23) hName = "勤労感謝の日";
      
      // 春分の日・秋分日の簡易計算（2026年対応など）
      if (m === 3 && ((year === 2026 && day === 20) || (year !== 2026 && day === 21))) hName = "春分の日";
      if (m === 9 && ((year === 2026 && day === 23) || (year !== 2026 && day === 23))) hName = "秋分の日";

      if (hName) {
        holidayMap[key] = hName;
      }
    }
    
    return { success: true, holidays: holidayMap };
  } catch (error) {
    return { success: true, holidays: {} };
  }
}