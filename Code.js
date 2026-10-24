/**
 * 💡 LINEからのイベント（メッセージ送信など）を受け取るWebhookの入口
 */
function doPost(e) {
  try {
    var json = JSON.parse(e.postData.contents);
    var events = json.events;
    
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      
      if (event.type === 'message' && event.message.type === 'text') {
        var userMessage = event.message.text.trim();
        var lineUserId = event.source.userId;
        var replyToken = event.replyToken;
        
        if (userMessage === "登録" || userMessage === "ID確認" || userMessage === "id") {
          saveLineIdLog(lineUserId);
          
          var webAppUrl = ScriptApp.getService().getUrl();
          var registerUrl = webAppUrl + "?page=register";

          // 💡 メッセージを3つに分割して送信します
          var messages = [
            { type: "text", text: "お疲れ様です！スタッフ登録を行います。\n\n① まず以下のご自身の【LINE ID】をコピーしてください👇" },
            { type: "text", text: lineUserId }, // IDの番号のみ（コピーしやすくするため）
            { type: "text", text: "② 次に、以下のリンクから登録フォームを開き、コピーしたLINE IDを貼り付けてください。\n\n▼ スタッフ初回登録ページ\n" + registerUrl }
          ];
          sendLineMultipleReplies(replyToken, messages);
        }
      }
    }
  } catch (error) {
    Logger.log("doPostエラー: " + error.message);
  }
  
  // 💡 確実に200 OKを返すためのレスポンス
  return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
             .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 🌐 ブラウザやURLにアクセスされたときの入口
 */
function doGet(e) {

  var params = e && e.parameter ? e.parameter : {};
  var page = params.page || "";
  
  // page=staff の場合はスタッフ用ポータルを表示
  if (page === "staff") {
    var template = HtmlService.createTemplateFromFile("staff_view");
    template.staffId = params.staffId || "";
    template.staffName = getStaffName(params.staffId);
    return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 💡 page=register の場合はスタッフ初回登録画面を表示
  if (page === "register") {
    var regTemplate = HtmlService.createTemplateFromFile("staff_register");
    return regTemplate.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 指定がない場合は管理者ダッシュボードを表示
  var adminTemplate = HtmlService.createTemplateFromFile("admin_view");
  return adminTemplate.evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * LINE IDをスプレッドシート（LINE連携ログ）に保存する関数
 */
function saveLineIdLog(lineUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("LINE連携ログ");
    if (!sheet) {
      sheet = ss.insertSheet("LINE連携ログ");
      sheet.appendRow(["timestamp", "lineUserId", "status"]);
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === lineUserId) {
        return;
      }
    }
    
    sheet.appendRow([new Date(), lineUserId, "未登録"]);
  } catch (e) {
    Logger.log("saveLineIdLogエラー: " + e.message);
  }
}

/**
 * ユーザーへ自動でリプライ（返信）を送る関数
 */
function sendLineReply(replyToken, messageText) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accountSheet = ss.getSheetByName("アカウントマスタ");
    if (!accountSheet) return;
    
    var accountData = accountSheet.getDataRange().getValues();
    var accessToken = "";
    if (accountData.length > 1 && accountData[1][1]) {
      accessToken = accountData[1][1];
    }
    
    if (!accessToken || accessToken.includes("YOUR_")) return;
    
    var url = "https://api.line.me/v2/bot/message/reply";
    var payload = {
      replyToken: replyToken,
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
  } catch (e) {
    Logger.log("sendLineReplyエラー: " + e.message);
  }
}

/**
 * 複数（配列）のメッセージをまとめて返信する関数
 */
function sendLineMultipleReplies(replyToken, messagesArray) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var accountSheet = ss.getSheetByName("アカウントマスタ");
    if (!accountSheet) return;
    
    var accountData = accountSheet.getDataRange().getValues();
    var accessToken = "";
    if (accountData.length > 1 && accountData[1][1]) {
      accessToken = accountData[1][1];
    }
    
    if (!accessToken || accessToken.includes("YOUR_")) return;
    
    var url = "https://api.line.me/v2/bot/message/reply";
    var payload = {
      replyToken: replyToken,
      messages: messagesArray
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
  } catch (e) {
    Logger.log("sendLineMultipleRepliesエラー: " + e.message);
  }
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
 * 新規スタッフを登録する関数（管理者の手動用）
 */
function registerNewStaff(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      throw new Error("「スタッフマスタ」シートが見つかりません。");
    }
    
    if (!formData.staffName || !formData.displayChar) {
      return { success: false, message: "エラー: スタッフ氏名とカレンダー表示文字は必須です。" };
    }
    
    var existingIds = [];
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) existingIds.push(String(data[i][0]));
    }
    
    var staffId;
    do {
      staffId = "st_" + Utilities.getUuid().substring(0, 8);
    } while (existingIds.indexOf(staffId) !== -1);
    
    var timestamp = new Date();
    var liffId = "2012184145-HMvvN2iP"; 
    var staffUrl = "https://liff.line.me/" + liffId + "?page=staff&staffId=" + staffId;
    
    sheet.appendRow([
      staffId,
      formData.staffName,
      "", 
      formData.assignedAccount,
      "有効",
      timestamp,
      staffUrl,
      formData.displayChar
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
 * シフト予定、休業日、フリー予定をまとめて一括保存する関数
 */
function saveAllSchedulesWithFree(data) {
  try {
    var res = saveAllSchedules(data);
    if (!res.success) return res;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var yearMonth = data.yearMonth;
    
    var freeSheet = ss.getSheetByName("フリー予定");
    if (!freeSheet) {
      freeSheet = ss.insertSheet("フリー予定");
      freeSheet.appendRow(["scheduleId", "workDate", "title"]);
    }
    
    var freeData = freeSheet.getDataRange().getValues();
    var newFreeRows = [freeData[0]];
    
    for (var i = 1; i < freeData.length; i++) {
      var fDate = new Date(freeData[i][1]);
      if (isNaN(fDate.getTime())) continue;
      var fYm = fDate.getFullYear() + '-' + ('0' + (fDate.getMonth() + 1)).slice(-2);
      if (fYm !== yearMonth) {
        newFreeRows.push(freeData[i]);
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
 * スタッフからのシフト希望を保存し、管理者向けの新着通知を記録する関数
 */
function s/**
 * スタッフからのシフト希望を保存し、管理者向けの新着通知を記録する関数（伝達事項対応版）
 */
function saveShiftPreferences(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var sheet = ss.getSheetByName("シフト予定");
    if (!sheet) {
      sheet = ss.insertSheet("シフト予定");
      sheet.appendRow(["shiftId", "workDate", "staffId", "staffName"]);
    }
    
    var existingData = sheet.getDataRange().getValues();
    var staffId = data.staffId;
    var staffName = data.staffName;
    var targetDates = data.selectedDates || [];
    var memo = data.memo || ""; // 💡 伝達事項（備考）を受け取る
    
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

    var notifSheet = ss.getSheetByName("希望通知ログ");
    if (!notifSheet) {
      notifSheet = ss.insertSheet("希望通知ログ");
      // 💡 伝達事項を記録できるように列を追加または対応（既存ログとの互換性を考慮）
      notifSheet.appendRow(["notifId", "staffId", "staffName", "yearMonth", "timestamp", "status", "memo"]);
    }
    
    var notifId = "notif_" + Utilities.getUuid().substring(0, 8);
    notifSheet.appendRow([
      notifId,
      staffId,
      staffName,
      data.yearMonth,
      new Date(),
      "未読",
      memo // 💡 ログに伝達事項を保存
    ]);
    
    return { success: true, message: "シフト希望を提出し、予定に反映しました！" };
  } catch (error) {
    return { success: false, message: "エラーが発生しました: " + error.message };
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
    
    var prefSheet = ss.getSheetByName("希望シフト");
    if (prefSheet) {
      var prefData = prefSheet.getDataRange().getValues();
      for (var i = 0; i < prefData.length; i++) {
        var row = prefData[i];
        if (row[1] === staffId) {
          var datesStr = String(row[4] || "");
          if (datesStr.indexOf(targetDate) !== -1) {
            var datesArr = datesStr.split(",").map(function(d){ return d.trim(); });
            var newDatesArr = datesArr.filter(function(d){ return d !== targetDate; });
            
            if (newDatesArr.length > 0) {
              prefSheet.getRange(i + 1, 5).setValue(newDatesArr.join(","));
            } else {
              prefSheet.deleteRow(i + 1);
            }
          }
        }
      }
    }
    
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
 * LINEアカウントのアクセストークンを考慮して一斉または個別にメッセージを送信する関数
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
 * 送信履歴を取得する関数
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
 * 過去1か月間の勤怠未提出者をリストアップする関数
 */
function getKintaiPendingStaff() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var today = new Date();
    var startDate = new Date();
    startDate.setDate(today.getDate() - 30);
    
    var todayStr = Utilities.formatDate(today, "Asia/Tokyo", "yyyy-MM-dd");
    var startStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy-MM-dd");

    var shiftSheet = ss.getSheetByName("シフト予定");
    if (!shiftSheet) return { success: true, pendingList: [] };
    var shiftData = shiftSheet.getDataRange().getValues();
    
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
      
      if (dStr >= startStr && dStr <= todayStr) {
        var key = sId + "_" + dStr;
        scheduledMap[key] = {
          staffId: sId,
          staffName: sName,
          targetDate: dStr
        };
      }
    }

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
        
        var submittedKey = kStaffId + "_" + kdStr;
        if (scheduledMap[submittedKey]) {
          delete scheduledMap[submittedKey];
        }
      }
    }

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
 * 未提出者リストをLINEリマインド用に詳細情報つきで取得する関数
 */
function getKintaiPendingForReminder() {
  try {
    var res = getKintaiPendingStaff();
    if (!res.success) return res;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    if (!staffSheet) return { success: true, pendingList: [] };
    var staffData = staffSheet.getDataRange().getValues();
    
    var staffLineMap = {};
    for (var i = 1; i < staffData.length; i++) {
      var sId = staffData[i][0];
      var lineId = staffData[i][2];
      var status = staffData[i][4];
      if (sId && lineId && status === "有効") {
        staffLineMap[sId] = true;
      }
    }

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
 * 未提出者へリマインドLINEを送信する関数
 */
function sendKintaiReminderLine(data) {
  try {
    var targets = data.targets || [];
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
 * 管理者向けに未読の「シフト希望申請通知」を取得する関数
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

        var formattedYm = yearMonth;
        if (yearMonth && yearMonth.indexOf("-") !== -1) {
          var ymParts = yearMonth.split("-");
          formattedYm = ymParts[0] + "年" + Number(ymParts[1]) + "月";
        } else if (yearMonth) {
          var ymDate = new Date(yearMonth);
          if (!isNaN(ymDate.getTime())) {
            formattedYm = Utilities.formatDate(ymDate, "Asia/Tokyo", "yyyy年M月");
          }
        }

        notifications.push({
          notifId: notifId,
          staffName: staffName || "スタッフ",
          yearMonth: formattedYm || "未指定",
          timestamp: timeStr
        });
      }
    }

    notifications.reverse();
    return { success: true, notifications: notifications };
  } catch (error) {
    return { success: true, notifications: [], message: error.message };
  }
}

/**
 * シフト希望通知を「既読（確認済み）」にする関数
 */
function markNotificationAsRead(notifId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: false, message: "通知シートが見つかりません" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(notifId)) {
        sheet.getRange(i + 1, 6).setValue("既読");
        return { success: true, message: "確認済みにしました" };
      }
    }
    return { success: false, message: "該当の通知が見つかりませんでした" };
  } catch (error) {
    return { success: false, message: "エラー: " + error.message };
  }
}

/**
 * 💡 スタッフ自身がフォームから登録を行った際の処理（アカウントごとの上限10名チェック付き）
 */
function registerStaffBySelf(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) {
      throw new Error("「スタッフマスタ」シートが見つかりません。");
    }
    
    if (!data.staffName || !data.displayChar || !data.lineUserId) {
      return { success: false, message: "エラー: 氏名、表示文字、LINE IDは必須です。" };
    }
    
    var assignedAccount = data.assignedAccount || "アカウントA"; 

    // 💡【追加】指定されたアカウントに現在「有効」なスタッフが何名いるかカウントする
    var dataRows = sheet.getDataRange().getValues();
    var activeCount = 0;
    
    for (var i = 1; i < dataRows.length; i++) {
      var rowAccount = dataRows[i][3]; // 割当アカウントの列
      var rowStatus  = dataRows[i][4]; // ステータスの列（有効 / 退職）
      
      if (rowAccount === assignedAccount && rowStatus === "有効") {
        activeCount++;
      }
    }

    // 💡【追加】すでに10名以上登録されている場合は登録をブロックする
    if (activeCount >= 10) {
      return { 
        success: false, 
        message: "⚠️ 登録エラー\n「" + assignedAccount + "」の登録上限（最大10名）に達しているため、これ以上登録できません。別のアカウントをご利用ください。" 
      };
    }
    
    var staffId = "st_" + Utilities.getUuid().substring(0, 6);
    var timestamp = new Date();
    var webAppUrl = ScriptApp.getService().getUrl();
    var staffUrl = webAppUrl + "?page=staff&staffId=" + staffId;
    
    sheet.appendRow([
      staffId,
      data.staffName,
      data.lineUserId,
      assignedAccount,
      "有効",
      timestamp,
      staffUrl,
      data.displayChar
    ]);
    
    sendWelcomeAndTestNotice(data.lineUserId, assignedAccount, data.staffName, staffUrl);
    
    return {
      success: true,
      message: "登録が完了しました！ご自身のLINEに専用ポータルURLを送信しましたのでご確認ください。"
    };
  } catch (error) {
    return { success: false, message: "登録エラー: " + error.message };
  }
}

/**
 * 登録完了時にスタッフへウェルカムメッセージと個別URLをプッシュ通知する関数
 */
function sendWelcomeAndTestNotice(lineUserId, assignedAccount, staffName, staffUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    "🎉 " + staffName + "さん、スタッフ登録が完了しました！\n\n" +
    "これより勤怠管理システムをご利用いただけます。\n" +
    "以下の【スタッフ専用ポータルURL】をタップして、無事に画面が開くか（申請テスト）をご確認ください。\n\n" +
    "🔗 スタッフ専用ポータルURL:\n" + staffUrl + "\n\n" +
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
  
  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    Logger.log("ウェルカムLINE送信失敗: " + e.message);
  }
}

/**
 * スタッフがLIFF経由で専用URLを開いたときにLINE IDを自動紐付けする関数
 */
function checkAndLinkStaffAccess(staffId, lineUserId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("スタッフマスタ");
    if (!sheet) return { success: false };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(staffId)) {
        var currentLineId = data[i][2];
        
        if (!currentLineId && lineUserId) {
          sheet.getRange(i + 1, 3).setValue(lineUserId);
          return { success: true, isNewlyLinked: true };
        }
        
        return { success: true, isNewlyLinked: false };
      }
    }
    return { success: false };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * LINEアカウントの詳細情報と現在の登録人数を取得する関数
 */
function getAccountDetailsList() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 💡 スタッフマスタから各アカウントごとの有効な登録人数をあらかじめ集計する
    var staffSheet = ss.getSheetByName("スタッフマスタ");
    var countsMap = {};
    if (staffSheet) {
      var staffData = staffSheet.getDataRange().getValues();
      for (var s = 1; s < staffData.length; s++) {
        var acc = staffData[s][3]; // 割当アカウント
        var status = staffData[s][4]; // ステータス (有効/退職)
        if (acc && status === "有効") {
          countsMap[acc] = (countsMap[acc] || 0) + 1;
        }
      }
    }

    var sheet = ss.getSheetByName("アカウントマスタ");
    if (!sheet) {
      return [
        { name: "アカウントA", addUrl: "https://lin.ee/xxxxxA", qrUrl: "", count: countsMap["アカウントA"] || 0 },
        { name: "アカウントB", addUrl: "https://lin.ee/xxxxxB", qrUrl: "", count: countsMap["アカウントB"] || 0 }
      ];
    }
    
    var data = sheet.getDataRange().getValues();
    var accounts = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        var accName = row[0];
        accounts.push({
          name: accName,
          channelId: row[1] || "",
          addUrl: row[2] || "https://lin.ee/default",
          qrUrl: row[3] || "",
          count: countsMap[accName] || 0 // 💡 登録人数をセット
        });
      }
    }
    return accounts.length > 0 ? accounts : [{ name: "アカウントA", addUrl: "", qrUrl: "", count: 0 }];
  } catch (error) {
    return [];
  }
}

/**
 * 💡 スタッフ自身が自分の直近の勤怠申請履歴を確認するための関数（最大20件）
 */
function getStaffKintaiHistory(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return { success: true, history: [] };
    
    var data = sheet.getDataRange().getValues();
    var history = [];
    
    // ヘッダーを飛ばして下から（新しい順に）走査
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var sId = row[1];
      
      if (String(sId) === String(staffId)) {
        var workDate = formatTimeOrDate(row[3]);
        var startTime = formatTime(row[4]);
        var endTime = formatTime(row[5]);
        var breakTime = row[6];
        var cost = Number(row[7]) || 0;
        var timestamp = formatTimestamp(row[8]);
        
        history.push({
          workDate: workDate,
          startTime: startTime,
          endTime: endTime,
          breakTime: breakTime,
          transportCost: cost,
          timestamp: timestamp
        });
        
        if (history.length >= 20) break;
      }
    }
    
    return { success: true, history: history };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

function formatTimeOrDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
  }
  var str = String(val);
  if (str.indexOf("T") !== -1) {
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd");
    }
  }
  return str.substring(0, 10);
}

function formatTimestamp(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  var d = new Date(val);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  }
  return String(val);
}

/**
 * 💡 スタッフ自身が自分の直近の勤怠申請履歴を確認するための関数（最大20件）
 */
function getStaffKintaiHistory(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) return { success: true, history: [] };
    
    var data = sheet.getDataRange().getValues();
    var history = [];
    
    // ヘッダーを飛ばして下から（新しい順に）走査
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var requestId = row[0]; // 一意のID (req_xxxx)
      var sId = row[1];
      
      if (String(sId) === String(staffId)) {
        var workDate = formatTimeOrDate(row[3]);
        var startTime = formatTime(row[4]);
        var endTime = formatTime(row[5]);
        var breakTime = row[6];
        var cost = Number(row[7]) || 0;
        
        history.push({
          requestId: requestId,
          workDate: workDate,
          startTime: startTime,
          endTime: endTime,
          breakTime: breakTime,
          transportCost: cost
        });
        
        if (history.length >= 20) break;
      }
    }
    
    return { success: true, history: history };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}

/**
 * 💡 スタッフが自分の申請履歴を修正してスプレッドシートを更新する関数
 */
function updateKintaiData(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("申請データ");
    if (!sheet) {
      throw new Error("「申請データ」シートが見つかりません。");
    }
    
    var sheetData = sheet.getDataRange().getValues();
    var targetRow = -1;
    
    // requestId（0列目）が一致する行を検索
    for (var i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][0]) === String(data.requestId)) {
        targetRow = i + 1; // 1行目がヘッダーなので +1
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, message: "対象の申請データが見つかりませんでした。" };
    }
    
    // 指定された行のデータを更新（勤務日、出勤時刻、退勤時刻、休憩、交通費）
    sheet.getRange(targetRow, 4).setValue(data.workDate);
    sheet.getRange(targetRow, 5).setValue(data.startTime);
    sheet.getRange(targetRow, 6).setValue(data.endTime);
    sheet.getRange(targetRow, 7).setValue(data.breakTime);
    sheet.getRange(targetRow, 8).setValue(data.transportCost);
    
    return { success: true, message: "申請内容を修正しました！" };
  } catch (error) {
    return { success: false, message: "修正エラー: " + error.message };
  }
}

function formatTimeOrDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy-MM-dd");
  }
  var str = String(val);
  if (str.indexOf("T") !== -1) {
    var d = new Date(val);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd");
    }
  }
  return str.substring(0, 10);
}

/**
 * 💡 スタッフが自分のシフト希望提出履歴（伝達事項含む）を確認するための関数
 */
function getStaffShiftHistory(staffId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("希望通知ログ");
    if (!sheet) return { success: true, history: [] };
    
    var data = sheet.getDataRange().getValues();
    var historyMap = {};
    
    // 下から（新しい順に）走査し、月ごとに最新の希望内容をまとめる
    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var sId = String(row[1] || "");
      var yearMonth = String(row[3] || "");
      var memo = String(row[6] || ""); // 伝達事項
      
      if (sId === String(staffId) && yearMonth) {
        if (!historyMap[yearMonth]) {
          historyMap[yearMonth] = {
            yearMonth: yearMonth,
            memo: memo
          };
        }
      }
    }
    
    // シフト予定シートから該当スタッフの月別希望日数を正確にカウントする
    var shiftSheet = ss.getSheetByName("シフト予定");
    var daysCountMap = {};
    if (shiftSheet) {
      var shiftData = shiftSheet.getDataRange().getValues();
      for (var j = 1; j < shiftData.length; j++) {
        var sRow = shiftData[j];
        var sDateStr = String(sRow[1] || "");
        var rowStaffId = String(sRow[2] || "");
        
        if (rowStaffId === String(staffId) && sDateStr) {
          var ym = sDateStr.substring(0, 7); // "yyyy-MM"
          daysCountMap[ym] = (daysCountMap[ym] || 0) + 1;
        }
      }
    }

    var historyList = [];
    for (var ymKey in historyMap) {
      var formattedYm = ymKey;
      if (ymKey.indexOf("-") !== -1) {
        var parts = ymKey.split("-");
        formattedYm = parts[0] + "年" + Number(parts[1]) + "月";
      }

      historyList.push({
        yearMonth: formattedYm,
        daysCount: daysCountMap[ymKey] || 0,
        memo: historyMap[ymKey].memo
      });
    }
    
    return { success: true, history: historyList };
  } catch (error) {
    return { success: false, history: [], message: error.message };
  }
}