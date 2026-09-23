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
    const props = PropertiesService.getScriptProperties();
    const folderId = props.getProperty("DRIVE_FOLDER_ID");
    
    // --- 1. 準備共用資料 ---
    const now = new Date();
    let targetDate = now;
    if (payload.customDate) {
      targetDate = new Date(payload.customDate);
      targetDate.setHours(12, 0, 0); 
    }
    const dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm"); 
    const content = payload.content;
    const tags = payload.tags;

    // --- 2. 處理檔案 (Base64 解碼) ---
    const files = payload.files || [];
    let uploadedCount = 0;

    // 情況 A：沒照片 -> 寫純文字
    if (files.length === 0) {
       const id = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
       sheet.appendRow([id, dateStr, timeStr, content, "", tags, "", ""]);
       _formatLastRow(sheet);
       return { success: true, message: "已儲存純文字日記" };
    } 
    // 情況 B：有多張照片 -> 逐一解碼上傳
    else {
      files.forEach((fileData, index) => {
        try {
          const id = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + index;
          let fileUrl = "";

          if (folderId && fileData.data) {
             const folder = DriveApp.getFolderById(folderId);
             
             // 🔥 關鍵動作：將 Base64 轉回 Blob
             // data:image/jpeg;base64,......
             const dataParts = fileData.data.split(','); // 切開檔頭
             const base64String = dataParts[1]; // 取內容
             
             // 解碼
             const decodedBytes = Utilities.base64Decode(base64String);
             const blob = Utilities.newBlob(decodedBytes, fileData.type, fileData.name);
             
             // 存檔
             const file = folder.createFile(blob); 
             file.setName(`${id}_${file.getName()}`); 
             fileUrl = file.getUrl(); 
             uploadedCount++;
          }

          sheet.appendRow([id, dateStr, timeStr, content, fileUrl, tags, "", ""]);
          _formatLastRow(sheet);

        } catch (e) {
          debugLog.push(`檔案 ${index+1} 上傳失敗: ${e.message}`);
        }
      });
    }

    // --- 3. 回報 ---
    if (uploadedCount > 0) {
       return { success: true, message: `✅ 完美！成功儲存 ${uploadedCount} 張照片。` };
    } else {
       return { success: false, message: `⚠️ 失敗：\n${debugLog.join("\n")}` };
    }
    
  } catch (error) {
    return { success: false, message: "❌ 系統錯誤：" + error.message };
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

        const currentImageUrls = String(sheetData[i][4] || "").trim(); // 第 4 欄 (索引值) 是 Media_Links 照片網址

        // 2. 處理刪除舊圖 (若 deleteImage 為 true 且原本有圖片)
        if (deleteImage && currentImageUrls) {
          const matches = currentImageUrls.matchAll(/\/d\/([a-zA-Z0-9_-]+)|id=([a-zA-Z0-9_-]+)/g);
          for (const m of matches) {
            const oldFileId = m[1] || m[2];
            if (oldFileId) {
              try {
                DriveApp.getFileById(oldFileId).setTrashed(true);
                console.log("✅ 成功將舊照片移至垃圾桶：" + oldFileId);
              } catch (err) {
                console.log("⚠️ 移動舊照片至垃圾桶失敗: " + err.message);
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
              const blob = Utilities.newBlob(decodedBytes, fData.type || 'image/jpeg', fData.name || `photo_${fIdx}.jpg`);
              const file = folder.createFile(blob);
              file.setName(`${id}_update_${fIdx}_${file.getName()}`);
              newUploadedUrls.push(file.getUrl());
            }
          }
        }

        // 4. 組合最終圖片網址
        let finalUrls = [];
        if (!deleteImage && currentImageUrls) {
          finalUrls.push(currentImageUrls); // 保留原照片網址
        }
        if (newUploadedUrls.length > 0) {
          finalUrls = finalUrls.concat(newUploadedUrls);
        }

        // 5. 若有新上傳或有執行刪除，則更新第 5 欄 (Media_Links)
        if (newUploadedUrls.length > 0 || deleteImage) {
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