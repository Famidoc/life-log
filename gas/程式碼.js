/**
 * 處理 GET 請求 (可用來測試 API 是否存活，或讀取資料)
 */
function doGet(e) {
  // 快取 Web App 網址供背景排程使用
  const url = ScriptApp.getService().getUrl();
  PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", url);

  // 如果有帶 action 參數，就當作 API 處理；否則回傳簡單的狀態訊息
  if (e.parameter.action === 'getRecentDiaries') {
    const offset = parseInt(e.parameter.offset) || 0;
    const limit = parseInt(e.parameter.limit) || 20;
    const data = getRecentDiaries(offset, limit);
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "Diary API is running" }))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 處理 POST 請求 (接收前端的寫入、讀取、更新、刪除等操作)
 */
function doPost(e) {
  let payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: "無效的 JSON 格式" }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  let result = {};
  
  // 路由機制：根據前端傳來的 action 決定要做什麼
  switch (payload.action) {
    // 👇 修改：讀取最新日記，支援接收 offset 與 limit 分頁參數
    case 'getRecentDiaries':
      const offset = (payload.data && payload.data.offset) ? payload.data.offset : 0;
      const limit = (payload.data && payload.data.limit) ? payload.data.limit : 20;
      result = { success: true, data: getRecentDiaries(offset, limit) };
      break;
      
    case 'saveDiaryV8':
      result = saveDiaryV8(payload.data); 
      break;
      
    case 'searchLogs':
      result = { success: true, data: searchLogs(payload.data.keyword, payload.data.searchDate) };
      break;
      
    case 'updateLog':
      result = updateLog(payload.data);
      break;
      
    case 'deleteLog':
      result = deleteLog(payload.data.id);
      break;
      
    case 'importFromSyncFolder':
      result = importFromSyncFolder();
      break;

    case 'getOnThisDay':
      result = { success: true, data: getOnThisDay() };
      break;

    case 'jumpToDate':
      result = jumpToDate(payload.data);
      break;

    case 'getOlderDiaries':
      result = getOlderDiaries(payload.data.minRow, payload.data.limit);
      break;

    case 'getNewerDiaries':
      result = getNewerDiaries(payload.data.maxRow, payload.data.limit);
      break;

    case 'sortLogs':
      result = sortLogsByDate();
      break;
      
    default:
      result = { success: false, message: "找不到對應的 API 動作：" + payload.action };
  }

  // 回傳 JSON 結果給 GitHub Pages 前端
  return ContentService.createTextOutput(JSON.stringify(result))
                       .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 存檔函式 V8.0 (Base64 解碼版)
 * 專門接收前端打包好的 Base64 檔案資料，保證多檔不漏接
 */
function saveDiaryV8(payload) {
  let debugLog = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) {
      return { success: false, message: "找不到 Logs 工作表" };
    }
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty("DRIVE_FOLDER_ID");
    const tz = Session.getScriptTimeZone();

    // --- 1. 準備共用資料 ---
    const now = new Date();
    let targetDate = now;
    if (payload.customDate) {
      const dateMatch = String(payload.customDate).trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
      if (dateMatch) {
        targetDate = new Date(
          parseInt(dateMatch[1], 10),
          parseInt(dateMatch[2], 10) - 1,
          parseInt(dateMatch[3], 10),
          12, 0, 0
        );
      } else {
        targetDate = new Date(payload.customDate);
        targetDate.setHours(12, 0, 0);
      }
    }
    const dateStr = Utilities.formatDate(targetDate, tz, "yyyy/MM/dd");
    const rawCustomTime = (payload.customTime !== undefined && payload.customTime !== null)
      ? String(payload.customTime).trim()
      : "";
    const timeStr = rawCustomTime ? rawCustomTime : Utilities.formatDate(now, tz, "HH:mm");
    const content = payload.content || "";
    const tags = payload.tags || "";

    // --- 2. ID 與冪等防重複檢查 ---
    const rawCustomId = (payload.customId !== undefined && payload.customId !== null)
      ? String(payload.customId).trim()
      : "";
    const id = rawCustomId ? rawCustomId : Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");

    if (rawCustomId) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const existingIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < existingIds.length; i++) {
          if (String(existingIds[i][0]).trim() === rawCustomId) {
            return {
              success: true,
              skipped: true,
              message: "已存在相同 ID，略過重複匯入",
              data: { id: rawCustomId }
            };
          }
        }
      }
    }

    // --- 3. 處理檔案上傳 (支援多圖，任一張失敗則回復不留殘缺資料) ---
    const files = payload.files || [];
    let uploadedCount = 0;
    let uploadedUrls = [];
    let createdFiles = [];

    if (files.length > 0) {
      if (!folderId) {
        debugLog.push("未設定 DRIVE_FOLDER_ID");
      } else {
        const folder = DriveApp.getFolderById(folderId);
        files.forEach((fileData, index) => {
          try {
            if (fileData && fileData.data) {
              const dataParts = fileData.data.split(',');
              const base64String = dataParts.length > 1 ? dataParts[1] : dataParts[0];
              const decodedBytes = Utilities.base64Decode(base64String);
              const blob = Utilities.newBlob(
                decodedBytes,
                fileData.type || 'image/jpeg',
                fileData.name || ('photo_' + index + '.jpg')
              );
              const file = folder.createFile(blob);
              file.setName(id + '_' + index + '_' + file.getName());
              createdFiles.push(file);
              uploadedUrls.push(file.getUrl());
              uploadedCount++;
            } else {
              debugLog.push('檔案 ' + (index + 1) + ' 缺少圖片資料');
            }
          } catch (e) {
            debugLog.push('檔案 ' + (index + 1) + ' 上傳失敗: ' + e.message);
          }
        });
      }

      if (uploadedCount < files.length) {
        createdFiles.forEach(file => {
          try { file.setTrashed(true); } catch (trashErr) {}
        });
        return {
          success: false,
          message: "部分照片上傳失敗: " + debugLog.join("; "),
          data: {
            uploadedCount: uploadedCount,
            expectedCount: files.length,
            debugLog: debugLog
          }
        };
      }
    }

    // --- 4. 生成 AI 溫暖生活評語 (Gemini 2.5 Flash) ---
    let aiSummary = "";
    if (payload.skipAI !== true) {
      try {
        aiSummary = generateAISummary(content);
      } catch (aiErr) {
        console.log("AI Summary 生成略過: " + aiErr.message);
      }
    }

    // 寫入試算表 (第 5 欄多圖網址以換行分隔，第 7 欄為 AI 評語)
    const mediaLinks = uploadedUrls.join("\n");
    sheet.appendRow([id, dateStr, timeStr, content, mediaLinks, tags, aiSummary, ""]);
    _formatLastRow(sheet);

    // 補寫舊日期時自動依日期排序 (批次匯入可傳入 skipSort: true 最後統一排序)
    if (payload.customDate && !payload.skipSort) {
      try {
        sortLogsByDate();
      } catch (sortErr) {
        console.log("自動排序略過: " + sortErr.message);
      }
    }

    let msg = "已儲存日記！";
    if (uploadedCount > 0) {
      msg = "✅ 完美！成功儲存日記與 " + uploadedCount + " 張照片。";
    }
    if (aiSummary) {
      msg += "\n🤖 AI 筆記：「" + aiSummary + "」";
    }

    return { 
      success: true, 
      message: msg,
      data: {
        id: id,
        ai_summary: aiSummary,
        uploadedCount: uploadedCount,
        expectedCount: files.length
      }
    };
    
  } catch (error) {
    return { success: false, message: "❌ 系統錯誤：" + error.message };
  }
}

/**
 * 依日期 (與時間、ID) 由舊到新穩定排序 Logs 工作表
 */
function sortLogsByDate() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) {
      return { success: false, message: "找不到 Logs 工作表" };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 2) {
      return { success: true, sortedCount: 0, message: "無需排序" };
    }

    const lastCol = Math.max(8, sheet.getLastColumn());
    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);

    if (range.getMergedRanges().length > 0) {
      return { success: false, message: "Logs 工作表含有合併儲存格，中止排序以避免覆寫錯誤" };
    }
    const formulas = range.getFormulas();
    for (let r = 0; r < formulas.length; r++) {
      for (let c = 0; c < formulas[r].length; c++) {
        if (formulas[r][c]) {
          return {
            success: false,
            message: "Logs 工作表第 " + (r + 2) + " 列第 " + (c + 1) + " 欄含有公式，中止排序以避免覆寫"
          };
        }
      }
    }

    const rows = range.getValues();
    const tz = Session.getScriptTimeZone();

    const keyed = rows.map((row, origIdx) => ({
      row: row,
      dateKey: _normalizeSortDateKey(row[1], tz),
      timeKey: _normalizeSortTimeKey(row[2]),
      idKey: String(row[0] || ""),
      origIdx: origIdx
    }));

    keyed.sort((a, b) => {
      if (a.dateKey !== b.dateKey) return a.dateKey < b.dateKey ? -1 : 1;
      if (a.timeKey !== b.timeKey) return a.timeKey < b.timeKey ? -1 : 1;
      if (a.idKey !== b.idKey) return a.idKey < b.idKey ? -1 : 1;
      return a.origIdx - b.origIdx;
    });

    const sortedRows = keyed.map(item => item.row);
    range.setValues(sortedRows);

    sheet.getRange(2, 2, sortedRows.length, 1).setNumberFormat("yyyy/MM/dd");
    sheet.setRowHeights(2, sortedRows.length, 30);
    sheet.getRange(2, 1, sortedRows.length, 8)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
      .setVerticalAlignment("middle");

    return {
      success: true,
      sortedCount: sortedRows.length,
      message: "已依日期由舊到新排序共 " + sortedRows.length + " 筆日記"
    };
  } catch (e) {
    return { success: false, message: "排序失敗：" + e.message };
  }
}

function _normalizeSortDateKey(val, tz) {
  if (!val) return "9999/99/99";
  if (val instanceof Date && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, tz, "yyyy/MM/dd");
  }
  const s = String(val).trim();
  const m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) {
    const d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), 12, 0, 0);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, "yyyy/MM/dd");
    }
  }
  const parsed = new Date(s.replace(/-/g, "/"));
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, "yyyy/MM/dd");
  }
  return "9999/99/99";
}

function _normalizeSortTimeKey(val) {
  if (val && typeof val.getHours === 'function' && !isNaN(val.getTime())) {
    return String(val.getHours()).padStart(2, '0') + ":" + String(val.getMinutes()).padStart(2, '0');
  }
  const s = String(val || "").trim();
  const m = s.match(/^(\d{1,2}):(\d{1,2})/);
  if (m) {
    return String(parseInt(m[1], 10)).padStart(2, '0') + ":" + String(parseInt(m[2], 10)).padStart(2, '0');
  }
  return s;
}

/**
 * 取得 Gemini API Key (優先讀取 ScriptProperties)
 */
function getGeminiApiKey() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("GEMINI_API_KEY") || "";
}

/**
 * 呼叫 Gemini 2.5 Flash 產生生活日記評語
 */
function generateAISummary(content) {
  if (!content || content.trim().length < 5) return "";
  const apiKey = getGeminiApiKey();
  if (!apiKey) return "";

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
  const prompt = "你是寫日記使用者的貼心好友與生活陪伴者。請閱讀以下使用者的日記內容，用繁體中文寫出一句 15~35 字、溫暖有同理心、富有生活哲思或鼓勵意味的簡短生活評語（請勿加上任何前綴、引號或括號，直接輸出這句話）：\n\n" + content;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      temperature: 0.7,
      maxOutputTokens: 80
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      if (json.candidates && json.candidates.length > 0 && json.candidates[0].content && json.candidates[0].content.parts.length > 0) {
        return json.candidates[0].content.parts[0].text.trim().replace(/^[\"「『]|[\"」』]$/g, "");
      }
    }
  } catch (e) {
    console.log("Gemini API 呼叫失敗: " + e.message);
  }
  return "";
}

/**
 * 取得「那年今天」歷史日記 (同月同日，且年份小於今年)
 */
function getOnThisDay() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth();
    const currentDate = today.getDate();

    const matches = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1] || row[0] === "ID") continue;
      const logDate = new Date(row[1]);
      if (isNaN(logDate.getTime())) continue;

      if (logDate.getMonth() === currentMonth && 
          logDate.getDate() === currentDate && 
          logDate.getFullYear() < currentYear) {
        matches.push(formatRowData(row));
      }
    }

    matches.sort((a, b) => new Date(b.dateStr) - new Date(a.dateStr));
    return matches;
  } catch (e) {
    console.log("getOnThisDay 錯誤: " + e.message);
    return [];
  }
}

// 輔助函式
function _formatLastRow(sheet) {
    const lastRow = sheet.getLastRow();
    sheet.setRowHeight(lastRow, 30);
    sheet.getRange(lastRow, 1, 1, 8).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.getRange(lastRow, 1, 1, 8).setVerticalAlignment("middle");
}

/**
 * 初始化 Google Sheet 資料庫結構
 * 依照 Spec v1.1 執行
 */
function setupSheet() {
  // 取得目前的試算表
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 定義我們需要的標準欄位 (Header)
  const headers = [
    "ID",           // 唯一碼 (YYYYMMDD-HHMMSS)
    "Date",         // 日期
    "Time",         // 時間
    "Content",      // 日記內文
    "Media_Links",  // 圖片/檔案的 Drive URL
    "Tags",         // 標籤 (#AI, #旅遊...)
    "AI_Summary",   // AI 自動生成的摘要
    "Related_IDs"   // AI 判斷關聯的舊日記 ID
  ];

  // 檢查是否已經有 'Logs' 這個工作表
  let sheet = ss.getSheetByName("Logs");
  
  if (!sheet) {
    // 如果沒有，就建立一個新的，並把預設的 '工作表1' 刪除 (保持整潔)
    sheet = ss.insertSheet("Logs");
    const defaultSheet = ss.getSheetByName("工作表1") || ss.getSheetByName("Sheet1");
    if (defaultSheet) ss.deleteSheet(defaultSheet);
    console.log("已建立新的 'Logs' 工作表");
  } else {
    console.log("'Logs' 工作表已存在，將更新標題列");
  }

  // 設定第一列標題
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // 美化格式：粗體、凍結第一列 (方便捲動查看)
  headerRange.setFontWeight("bold");
  sheet.setFrozenRows(1);

  // 自動調整欄寬 (讓 ID 和 時間 欄位寬一點，看起來比較舒服)
  sheet.setColumnWidth(1, 150); // ID
  sheet.setColumnWidth(4, 400); // Content (內容欄寬一點)

  console.log("資料庫初始化完成！");
}

/**
 * 建立日曆行程的函式 (維持不變)
 */
function createCalendarReminder(note) {
  const calendar = CalendarApp.getDefaultCalendar();
  const now = new Date();
  const startTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 19, 0, 0);
  const endTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 19, 15, 0);

  const event = calendar.createEvent("📝 寫日記時間 (AI 提醒)", startTime, endTime, {
    description: note,
  });

  event.addPopupReminder(0);
  console.log("✅ 已在您的 Google 日曆建立提醒！請查看 19:00 的行程。");
}

function debugProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const keys = Object.keys(props);
  
  console.log("=== 保險箱內容檢查 ===");
  if (keys.length === 0) {
    console.log("⚠️ 保險箱是空的！請去「專案設定」新增屬性。");
  } else {
    keys.forEach(key => {
      // 為了安全，我們只顯示 Key 的名稱，不顯示值
      console.log(`✅ 發現鑰匙：[${key}]`);
    });
  }
  console.log("========================");
}

/**
 * 用來觸發 Google Drive 授權視窗專用的函式
 */
function authorizeDrive() {
  // 隨便呼叫一個 Drive 功能，Google 就會跳出來問你要權限了
  DriveApp.getRootFolder();
  console.log("✅ 授權成功！");
}

/**
 * 批次匯入功能 (V6.0 修正版：真正的搬運工)
 * 邏輯：掃描 SYNC -> 複製到 DRIVE -> 建立日記 -> 刪除 SYNC 原檔
 */
function importFromSyncFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Logs");
  const props = PropertiesService.getScriptProperties();
  
  const syncFolderId = props.getProperty("SYNC_FOLDER_ID");
  const driveFolderId = props.getProperty("DRIVE_FOLDER_ID");

  if (!syncFolderId || !driveFolderId) {
    return { success: false, message: "❌ 資料夾 ID 未設定，請檢查指令碼屬性" };
  }

  const syncFolder = DriveApp.getFolderById(syncFolderId);
  const driveFolder = DriveApp.getFolderById(driveFolderId);
  const files = syncFolder.getFiles();
  
  let count = 0;
  
  // 掃描 SYNC 資料夾裡的所有檔案
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    
    // 1. 產生 ID 與 時間
    const now = new Date();
    // 為了避免同時匯入多張照片 ID 重複，我們在 ID 後面加個計數器
    const id = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + count;
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm");

    // 2. 關鍵動作：複製到倉庫 (DRIVE_FOLDER) 並改名
    // 這樣 DRIVE_FOLDER 裡就會有檔案了！
    const newFile = file.makeCopy(`${id}_${fileName}`, driveFolder);
    const fileUrl = newFile.getUrl();

    // 3. 建立日記內容
    const content = `[自動匯入] 來自電腦同步的檔案：${fileName}`;
    
    // 4. 寫入 Sheet (注意：AI 欄位給空值)
    // 格式：[ID, Date, Time, Content, ImageURL, Tags, AI_Summary, Related_IDs]
    sheet.appendRow([id, dateStr, timeStr, content, fileUrl, "#自動同步", "", ""]);
    
    // 固定行高美化
    const lastRow = sheet.getLastRow();
    sheet.setRowHeight(lastRow, 30);
    sheet.getRange(lastRow, 1, 1, 8).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // 5. 關鍵動作：刪除 SYNC 裡的舊檔 (清空還書箱)
    // 這樣您下次看 SYNC 資料夾時，已經處理完的檔案就會消失，不會重複匯入
    file.setTrashed(true);
    
    count++;
  }

  if (count === 0) {
    return { success: true, message: "📭 同步資料夾是空的，沒有新照片喔！" };
  } else {
    return { success: true, message: `✅ 成功匯入並搬運了 ${count} 張照片！` };
  }
}

/**
 * 讀取日記 (支援分頁 Load More 功能)
 * @param {number} offset 跳過幾筆資料
 * @param {number} limit 抓取幾筆資料
 */
function getRecentDiaries(offset = 0, limit = 20) {
  console.log(`🚀 開始執行 getRecentDiaries... offset=${offset}, limit=${limit}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Logs");
  
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  
  // 計算要抓取的範圍 (資料越下方越新)
  const endRow = lastRow - offset;
  if (endRow < 2) return []; // 已經沒有更舊的資料了

  // 確保不要抓到表頭 (第 1 列)
  const startRow = Math.max(2, endRow - limit + 1);
  const totalRows = endRow - startRow + 1;
  
  const data = sheet.getRange(startRow, 1, totalRows, 8).getValues();

  try {
    // 統一使用 formatRowData 處理格式，並使用 reverse 讓最新的一筆在最前面
    const result = data.reverse().map(row => formatRowData(row));
    console.log(`📤 資料準備完成，抓取筆數: ${result.length}`);
    return result;
  } catch (error) {
    console.log("❌ 錯誤: " + error.toString());
    return [];
  }
}

/**
 * 搜尋功能：支援 關鍵字 + 日期
 */
function searchLogs(keyword, searchDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Logs");
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) return [];

  // 讀取所有資料 (如果資料量破萬筆，這裡可能需要優化，但個人日記目前這樣沒問題)
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  
  // 開始過濾
  const results = data.filter(row => {
    let matchDate = true;
    let matchKeyword = true;

    // 1. 比對日期 (如果有選日期的話)
    if (searchDate) {
      const rowDateStr = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (rowDateStr !== searchDate) matchDate = false;
    }

    // 2. 比對關鍵字 (如果有輸入的話) -> 搜尋內容或標籤
    if (keyword) {
      const content = row[3].toString().toLowerCase();
      const tags = row[5].toString().toLowerCase();
      const query = keyword.toLowerCase();
      if (!content.includes(query) && !tags.includes(query)) matchKeyword = false;
    }

    return matchDate && matchKeyword;
  });

  // 格式化回傳資料 (跟 getRecentDiaries 一樣的格式)
  return results.reverse().map(row => formatRowData(row));
}

/**
 * ✨ 跳轉至指定日期並載入前後鄰接日記 (前後漫遊核心)
 * @param {Object} data - { targetDate: "YYYY-MM-DD" 或 "YYYY/MM/DD", radius: 12 }
 */
function jumpToDate(data) {
  try {
    const rawTarget = (data && data.targetDate) ? String(data.targetDate).trim() : "";
    if (!rawTarget) {
      return { success: false, message: "請提供欲跳轉的日期" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) return { success: false, message: "找不到 Logs 工作表" };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: [], message: "目前尚無日記" };
    }

    const radius = (data && data.radius) ? parseInt(data.radius, 10) : 12;

    // 正規化目標日期為 yyyy/MM/dd
    const targetNormalized = rawTarget.replace(/-/g, "/");
    const targetDateObj = new Date(targetNormalized);
    if (isNaN(targetDateObj.getTime())) {
      return { success: false, message: "日期格式不正確" };
    }
    const targetYMD = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 讀取所有日期與ID資料 (第 1, 2 欄，從第 2 列到 lastRow)
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    const rowList = [];
    const uniqueDatesSet = new Set();

    for (let i = 0; i < values.length; i++) {
      const idVal = values[i][0];
      const cellVal = values[i][1];
      if (!cellVal || idVal === "ID") continue;

      let dObj = null;
      let dStr = "";
      try {
        if (cellVal instanceof Date) {
          dObj = cellVal;
        } else {
          dObj = new Date(String(cellVal).replace(/-/g, "/"));
        }

        if (dObj && !isNaN(dObj.getTime())) {
          dStr = Utilities.formatDate(dObj, Session.getScriptTimeZone(), "yyyy/MM/dd");
          const rowNum = i + 2;
          rowList.push({ rowNumber: rowNum, dateStr: dStr, timestamp: dObj.getTime() });
          uniqueDatesSet.add(dStr);
        }
      } catch (err) {}
    }

    if (rowList.length === 0) {
      return { success: true, data: [], message: "尚無有效日期資料" };
    }

    // 所有不重複日期陣列，按時間升序排序
    const sortedUniqueDates = Array.from(uniqueDatesSet).sort((a, b) => new Date(a) - new Date(b));

    // 計算上一篇日記日期 prevDate 與下一篇日記日期 nextDate
    let prevDate = null;
    let nextDate = null;
    for (let i = 0; i < sortedUniqueDates.length; i++) {
      if (sortedUniqueDates[i] < targetYMD) {
        prevDate = sortedUniqueDates[i];
      }
      if (sortedUniqueDates[i] > targetYMD && nextDate === null) {
        nextDate = sortedUniqueDates[i];
      }
    }

    // 找出匹配或最接近 targetYMD 的 row
    let targetIndices = [];
    for (let i = 0; i < rowList.length; i++) {
      if (rowList[i].dateStr === targetYMD) {
        targetIndices.push(i);
      }
    }

    let exactMatch = targetIndices.length > 0;
    let centerIndex = -1;

    if (exactMatch) {
      centerIndex = targetIndices[Math.floor(targetIndices.length / 2)];
    } else {
      const targetTime = targetDateObj.getTime();
      let minDiff = Infinity;
      for (let i = 0; i < rowList.length; i++) {
        const diff = Math.abs(rowList[i].timestamp - targetTime);
        if (diff < minDiff) {
          minDiff = diff;
          centerIndex = i;
        }
      }
    }

    // 計算以 centerIndex 為中心，前後各取 radius 筆的範圍
    const startIndex = Math.max(0, centerIndex - radius);
    const endIndex = Math.min(rowList.length - 1, centerIndex + radius);

    const startRowNum = rowList[startIndex].rowNumber;
    const endRowNum = rowList[endIndex].rowNumber;
    const numRows = endRowNum - startRowNum + 1;

    const rawRows = sheet.getRange(startRowNum, 1, numRows, 8).getValues();

    // 轉為前端日記物件，並由新到舊倒序排列
    const diaries = [];
    for (let r = rawRows.length - 1; r >= 0; r--) {
      diaries.push(formatRowData(rawRows[r]));
    }

    const hasOlder = startRowNum > 2;
    const hasNewer = endRowNum < lastRow;

    return {
      success: true,
      data: diaries,
      targetDate: targetYMD,
      exactMatch: exactMatch,
      prevDate: prevDate,
      nextDate: nextDate,
      minRow: startRowNum,
      maxRow: endRowNum,
      hasOlder: hasOlder,
      hasNewer: hasNewer
    };
  } catch (err) {
    return { success: false, message: "跳轉日期失敗: " + err.message };
  }
}

/**
 * 載入更早的日記 (從 minRow 往前抓)
 */
function getOlderDiaries(minRow, limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) return { success: false, data: [] };

    const fetchLimit = limit ? parseInt(limit, 10) : 15;
    const targetEndRow = parseInt(minRow, 10) - 1;
    if (targetEndRow < 2) {
      return { success: true, data: [], hasOlder: false, minRow: 2 };
    }

    const startRow = Math.max(2, targetEndRow - fetchLimit + 1);
    const numRows = targetEndRow - startRow + 1;

    const rawRows = sheet.getRange(startRow, 1, numRows, 8).getValues();
    const diaries = [];
    for (let r = rawRows.length - 1; r >= 0; r--) {
      diaries.push(formatRowData(rawRows[r]));
    }

    return {
      success: true,
      data: diaries,
      minRow: startRow,
      hasOlder: startRow > 2
    };
  } catch (err) {
    return { success: false, message: "載入更早回憶失敗: " + err.message };
  }
}

/**
 * 載入更新的日記 (從 maxRow 往後抓)
 */
function getNewerDiaries(maxRow, limit) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    if (!sheet) return { success: false, data: [] };

    const fetchLimit = limit ? parseInt(limit, 10) : 15;
    const lastRow = sheet.getLastRow();
    const targetStartRow = parseInt(maxRow, 10) + 1;
    if (targetStartRow > lastRow) {
      return { success: true, data: [], hasNewer: false, maxRow: lastRow };
    }

    const endRow = Math.min(lastRow, targetStartRow + fetchLimit - 1);
    const numRows = endRow - targetStartRow + 1;

    const rawRows = sheet.getRange(targetStartRow, 1, numRows, 8).getValues();
    const diaries = [];
    for (let r = rawRows.length - 1; r >= 0; r--) {
      diaries.push(formatRowData(rawRows[r]));
    }

    return {
      success: true,
      data: diaries,
      maxRow: endRow,
      hasNewer: endRow < lastRow
    };
  } catch (err) {
    return { success: false, message: "載入更新回憶失敗: " + err.message };
  }
}

/**
 * 更新日記功能 (支援保留原圖、追加多圖、刪除舊圖連動垃圾桶)
 */
function updateLog(idOrData, newContent, newTags, filePayload) {
  try {
    let data = {};
    if (typeof idOrData === 'object' && idOrData !== null) {
      data = idOrData;
    } else {
      data = {
        id: idOrData,
        content: newContent,
        tags: newTags,
        filePayload: filePayload
      };
    }

    const id = data.id;
    const content = data.content;
    const tags = data.tags;
    const files = data.files || (data.filePayload ? [data.filePayload] : []);
    const deleteImage = data.deleteImage === true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    const sheetData = sheet.getDataRange().getValues();
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty("DRIVE_FOLDER_ID");

    // 尋找 ID 所在的那一行 (從第2列開始找)
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][0] == id) { // 第 0 欄是 ID
        // 1. 更新文字內容與標籤
        sheet.getRange(i + 1, 4).setValue(content); // Content
        sheet.getRange(i + 1, 6).setValue(tags);    // Tags

        const currentImageUrls = String(sheetData[i][4] || "").trim(); // 第 4 欄是 Media_Links

        // 2. 處理刪除舊圖 (精準單張刪除 或 整篇清空)
        if (data.deleteFileIds && Array.isArray(data.deleteFileIds) && data.deleteFileIds.length > 0) {
          data.deleteFileIds.forEach(oldId => {
            if (oldId) {
              try {
                DriveApp.getFileById(oldId).setTrashed(true);
                console.log("✅ 成功將指定舊照片移至垃圾桶：" + oldId);
              } catch (err) {
                console.log("⚠️ 移動指定照片至垃圾桶失敗: " + err.message);
              }
            }
          });
        } else if (deleteImage && currentImageUrls) {
          const matches = currentImageUrls.matchAll(/\/d\/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)/g);
          for (const m of matches) {
            const oldFileId = m[1] || m[2];
            if (oldFileId) {
              try {
                DriveApp.getFileById(oldFileId).setTrashed(true);
                console.log("✅ 成功將全部舊照片移至垃圾桶：" + oldFileId);
              } catch (err) {
                console.log("⚠️ 移動全部照片至垃圾桶失敗: " + err.message);
              }
            }
          }
        }

        // 3. 處理新上傳圖片 (支援多張)
        let newUploadedUrls = [];
        if (files && files.length > 0 && folderId) {
          const folder = DriveApp.getFolderById(folderId);
          for (let fIdx = 0; fIdx < files.length; fIdx++) {
            const fData = files[fIdx];
            if (fData && fData.data) {
              const dataParts = fData.data.split(',');
              const base64String = dataParts[1];
              const decodedBytes = Utilities.base64Decode(base64String);
              const blob = Utilities.newBlob(decodedBytes, fData.type || 'image/jpeg', fData.name || ('photo_' + fIdx + '.jpg'));
              const file = folder.createFile(blob);
              file.setName(id + '_update_' + fIdx + '_' + file.getName());
              newUploadedUrls.push(file.getUrl());
            }
          }
        }

        // 4. 組合最終圖片網址 (精準保留前端未刪除的舊圖)
        let finalUrls = [];
        if (data.keepImageUrls && Array.isArray(data.keepImageUrls)) {
          finalUrls = finalUrls.concat(data.keepImageUrls);
        } else if (!deleteImage && currentImageUrls) {
          finalUrls.push(currentImageUrls);
        }

        if (newUploadedUrls.length > 0) {
          finalUrls = finalUrls.concat(newUploadedUrls);
        }

        // 5. 若有新上傳或有刪除/調整照片，則更新第 5 欄 (Media_Links)
        if (newUploadedUrls.length > 0 || deleteImage || data.keepImageUrls) {
          sheet.getRange(i + 1, 5).setValue(finalUrls.join("\n"));
        }

        return { success: true, message: "日記已更新！" };
      }
    }
    return { success: false, message: "找不到該筆日記 ID" };
  } catch (e) {
    return { success: false, message: "更新失敗：" + e.message };
  }
}

/**
 * 輔助函式：統一資料格式化 (避免重複寫代碼)
 */
function formatRowData(row) {
  // 處理日期
  let dateStr = "";
  try { if (row[1]) dateStr = Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), "yyyy/MM/dd"); } catch (e) {}
  
  // 處理時間
  let timeStr = row[2];
  try {
    if (row[2] && typeof row[2].getHours === 'function') {
        let h = row[2].getHours().toString().padStart(2, '0');
        let m = row[2].getMinutes().toString().padStart(2, '0');
        timeStr = `${h}:${m}`;
    }
  } catch (e) {}

  return {
    id: row[0],
    dateStr: dateStr,
    time: timeStr,
    content: row[3],
    image: row[4],
    tags: row[5],
    ai_summary: row[6]
  };
}

/**
 * ⏰ 每日提醒功能
 * 設定每天晚上 19:00 寄信或透過 Google Calendar 提醒
 */

// 1. 設定觸發條件 (請手動執行一次這個函式)
function setupReminderTrigger() {
  // 先清除舊的，避免重複
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "sendDailyReminders") {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  // 設定新的：每天晚上 19:00 ~ 20:00 之間觸發
  ScriptApp.newTrigger("sendDailyReminders")
    .timeBased()
    .everyDays(1)
    .atHour(19)
    .create();
    
  console.log("✅ 每日 19:00 提醒已設定完成！");
}

/**
 * 執行提醒 (Email + Google日曆 + 那年今天回顧)
 * V10.1: 修正圖片顯示問題，改用「內嵌圖片 (CID)」技術，確保 Gmail 一定看得到圖
 */
function sendDailyReminders() {
  const email = Session.getActiveUser().getEmail();
  
  // ⭐️ 修正：改由指令碼屬性取得網址，避免觸發器環境下抓不到
  const props = PropertiesService.getScriptProperties();
  const scriptUrl = props.getProperty("WEB_APP_URL") || "請手動開啟您的 Web App 網址";
  
  // 1. 準備鼓勵語
  const quotes = [
    "忙碌了一天，留點時間與自己對話吧。",
    "今天的快樂碎片收集了嗎？快來記下來！",
    "每一天都值得被紀錄，哪怕只是平淡的日常。",
    "晚安，在休息前，把煩惱留在日記裡吧。",
    "嘿！未來的你在等著看今天的回憶呢！"
  ];
  const randomQuote = quotes[Math.floor(Math.random() * quotes.length)];

  // --- ⭐ 新增：時光機邏輯 (On This Day) ---
  let historyHtml = "";
  // 準備一個物件來裝圖片檔案 (CID Map)
  let inlineImages = {}; 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    const data = sheet.getDataRange().getValues(); 
    const today = new Date();
    
    // 篩選出「不同年份」但「同月同日」的日記
    const pastEntries = data.filter(row => {
      if (!row[1] || row[0] === "ID") return false; 
      const logDate = new Date(row[1]);
      return logDate.getMonth() === today.getMonth() && 
             logDate.getDate() === today.getDate() &&   
             logDate.getFullYear() < today.getFullYear(); 
    });

    if (pastEntries.length > 0) {
      pastEntries.sort((a, b) => new Date(b[1]) - new Date(a[1]));

      let entriesList = pastEntries.map((entry, index) => {
        const year = new Date(entry[1]).getFullYear();
        const content = entry[3]; 
        let imgTag = "";
        
        // 處理照片 (改用內嵌方式)
        if (entry[4] && entry[4].includes("drive.google.com")) {
           let fileId = "";
           let match = entry[4].match(/\/d\/([a-zA-Z0-9_-]+)/);
           if (match) fileId = match[1];
           else { match = entry[4].match(/id=([a-zA-Z0-9_-]+)/); if(match) fileId = match[1]; }
           
           if (fileId) {
             try {
               // ⭐ 關鍵：抓取檔案 Blob
               const imgBlob = DriveApp.getFileById(fileId).getBlob();
               // 給它一個獨一無二的代號 (CID)
               const cid = "pastImage_" + year + "_" + index;
               // 放入附件清單
               inlineImages[cid] = imgBlob;
               // 在 HTML 裡呼叫這個 CID
               imgTag = `<br><img src="cid:${cid}" style="max-width:100%; border-radius:8px; margin-top:15px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">`;
             } catch (e) {
               console.log("讀取圖片失敗 (可能是權限或ID錯誤): " + e.message);
             }
           }
        }

        return `
          <div style="background:white; padding:20px; margin-bottom:15px; border-left: 4px solid #5e72e4; border-radius:8px; text-align:left; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
            <div style="font-weight:bold; color:#5e72e4; margin-bottom:8px; font-size:1.1em;">📅 ${year} 年的今天</div>
            <div style="color:#444; line-height:1.6; font-size:15px;">${content}</div>
            ${imgTag}
          </div>
        `;
      }).join("");

      historyHtml = `
        <div style="margin-top:30px; padding-top:20px; border-top:1px dashed #ccc;">
          <h3 style="color:#525f7f; margin-bottom:20px; text-align:center;">🕰️ 那年今天...</h3>
          ${entriesList}
        </div>
      `;
    }
  } catch (e) {
    console.log("時光機讀取失敗: " + e.message);
  }
  // ----------------------------------------

  // 2. 寄送 Email (⭐ 記得加入 inlineImages 參數)
  MailApp.sendEmail({
    to: email,
    subject: "📖 該寫日記囉！(附上那年今天回顧)",
    htmlBody: `
      <div style="background:#f4f7fe; padding:40px 20px; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; text-align:center;">
        <div style="background:white; padding:40px; border-radius:16px; max-width:550px; margin:0 auto; box-shadow:0 4px 20px rgba(0,0,0,0.1);">
          
          <h2 style="color:#5e72e4; margin-top:0; font-size:24px;">👋 晚安，今天過得好嗎？</h2>
          <p style="font-size:16px; color:#555; line-height:1.8; font-style:italic; margin-bottom:30px;">"${randomQuote}"</p>
          
          <a href="${scriptUrl}" style="display:inline-block; background:linear-gradient(135deg, #5e72e4 0%, #825ee4 100%); color:white; text-decoration:none; padding:15px 30px; border-radius:50px; font-weight:bold; font-size:16px; box-shadow: 0 4px 15px rgba(94, 114, 228, 0.4);">
            ✍️ 點我寫日記
          </a>

          ${historyHtml}

        </div>
        <p style="color:#8898aa; font-size:12px; margin-top:30px;">Smart Life Log 自動提醒系統</p>
      </div>
    `,
    inlineImages: inlineImages // ⭐ 這行最重要！把照片貼上去
  });

  // 3. 建立 Google 日曆事項
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const now = new Date();
    const endTime = new Date(now.getTime() + 15 * 60 * 1000); 
    const event = calendar.createEvent('✍️ 寫日記時間', now, endTime, {
      description: `點擊連結開始寫日記：\n${scriptUrl}`
    });
    event.addPopupReminder(0);
  } catch (e) {}
}

/**
 * 刪除日記功能 (V12.0 連動刪除雲端硬碟照片版)
 */
function deleteLog(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Logs");
    const data = sheet.getDataRange().getValues();

    // 尋找 ID 所在的那一行 (從第2列開始找)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) { // 第 0 欄是 ID
        
        const fileUrl = data[i][4]; // 第 4 欄 (索引值) 是 Media_Links 照片網址
        
        // --- 🧹 新增：刪除 Google Drive 中的照片原檔 ---
        if (fileUrl && fileUrl.includes("drive.google.com")) {
          let fileId = "";
          // 用正規表達式從網址中抓取 File ID
          let match = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
          if (match) {
            fileId = match[1];
          } else {
             match = fileUrl.match(/id=([a-zA-Z0-9_-]+)/); 
             if (match) fileId = match[1]; 
          }
          
          if (fileId) {
            try {
              // 找到檔案並移至垃圾桶 (避免永久刪除救不回來，先丟垃圾桶是最安全的作法)
              DriveApp.getFileById(fileId).setTrashed(true);
              console.log("✅ 成功將關聯照片移至垃圾桶：" + fileId);
            } catch (e) {
              console.log("⚠️ 照片刪除失敗 (可能是權限不足或檔案已被刪除)：" + e.message);
            }
          }
        }
        // ----------------------------------------------

        sheet.deleteRow(i + 1); // 刪除該列文字資料
        return { success: true, message: "🗑️ 日記與關聯照片已成功刪除！" };
      }
    }
    return { success: false, message: "找不到該筆日記" };
  } catch (e) {
    return { success: false, message: "刪除失敗：" + e.message };
  }
}