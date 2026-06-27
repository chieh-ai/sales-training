/**
 *  Sales Training 互動影片學習平台
 * 後端 Google Apps Script - REST API 版本
 * 相容 GitHub Pages 前端部署
 *
 * ================================================
 *  部署說明（每次修改後需重新部署）：
 *  1. GAS 編輯器 → 「部署」→「管理部署作業」→「新增部署作業」
 *  2. 類型：「網路應用程式」
 *  3. 執行身分：「我（your@gmail.com）」
 *  4. 誰可以存取：「任何人」（不需 Google 帳號）
 *  5. 部署後複製 /exec 結尾的網址
 *  6. 貼到 index.html 最上方的 GAS_API_URL 變數
 * ================================================
 */

// ==========================================
//  doGet：回傳 API 狀態說明（前端不再從這裡取 HTML）
// ==========================================
function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'Sales Training API 運作中。' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
//  doPost：主要 REST API 進入點
//  前端用 fetch + Content-Type: text/plain 呼叫
// ==========================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var result = handleAction(body);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: '解析請求失敗: ' + err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
//  統一 action 分發器
// ==========================================
function handleAction(body) {
  var action   = body.action || '';
  var email    = (body.email || '').toLowerCase().trim();
  var password = body.password || '';

  if (!email || !password) {
    return { status: 'error', message: '請提供帳號與密碼' };
  }

  var userRow = validateUser(email, password);
  if (!userRow) {
    return { status: 'error', message: '帳號或密碼錯誤，請確認後重試。' };
  }

  try {
    if      (action === 'login')               { return getInitialDataForUser(email, userRow); }
    else if (action === 'saveWatchLog')        { return saveWatchLog(email, body.sessionData); }
    else if (action === 'toggleFavorite')      { return toggleFavorite(email, body.videoId, body.isFav); }
    else if (action === 'saveAdminVideo')      { return saveAdminVideo(email, body.data); }
    else if (action === 'deleteAdminVideo')    { return deleteAdminVideo(email, body.videoId); }
    else if (action === 'saveAdminPlaylist')   { return saveAdminPlaylist(email, body.data); }
    else if (action === 'deleteAdminPlaylist') { return deleteAdminPlaylist(email, body.id); }
    else if (action === 'saveAdminUser')       { return saveAdminUser(email, body.data); }
    else if (action === 'deleteAdminUser')     { return deleteAdminUser(email, body.targetEmail); }
    else if (action === 'changePassword')      { return changePasswordFn(email, body.newPassword); }
    else { return { status: 'error', message: '未知的 action: ' + action }; }
  } catch(err) {
    return { status: 'error', message: err.message };
  }
}

// ==========================================
//  帳號驗證（比對 Users 工作表的 Password 欄）
// ==========================================
function validateUser(email, password) {
  var users = getSheetDataAsObjects("Users");
  for (var i = 0; i < users.length; i++) {
    if (users[i].Email &&
        users[i].Email.toLowerCase() === email &&
        String(users[i].Password) === String(password)) {
      return users[i];
    }
  }
  return null;
}

// ==========================================
//  工具：讀取工作表為 Object 陣列
// ==========================================
// ==========================================
//  修改密碼
// ==========================================
function changePasswordFn(email, newPassword) {
  try {
    if (!newPassword || String(newPassword).length < 8) return { status: 'error', message: '密碼長度至少 8 碼，需含英文字母與數字' };
    if (!/[a-zA-Z]/.test(String(newPassword)) || !/[0-9]/.test(String(newPassword))) return { status: 'error', message: '密碼須包含至少一個英文字母與一個數字' };
    upsertRow("Users", "Email", email, { "Password": String(newPassword) });
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function getSheetDataAsObjects(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        var cellValue = row[j];
        if (cellValue instanceof Date) {
          var yyyy = cellValue.getFullYear();
          var mm   = String(cellValue.getMonth() + 1).padStart(2, '0');
          var dd   = String(cellValue.getDate()).padStart(2, '0');
          var hh   = String(cellValue.getHours()).padStart(2, '0');
          var min  = String(cellValue.getMinutes()).padStart(2, '0');
          var sec  = String(cellValue.getSeconds()).padStart(2, '0');
          cellValue = yyyy + '/' + mm + '/' + dd + ' ' + hh + ':' + min + ':' + sec;
        }
        obj[headers[j]] = cellValue !== undefined ? cellValue : "";
      }
    }
    result.push(obj);
  }
  return result;
}

// ==========================================
//  登入後取得初始資料
// ==========================================

// ==========================================
//  FactoryOptions 工作表：自動建立（首次執行）
// ==========================================
function ensureFactoryOptionsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("FactoryOptions");
  if (sheet) return sheet;
  sheet = ss.insertSheet("FactoryOptions");
  sheet.getRange(1, 1, 1, 3).setValues([["Type", "Code", "Label"]]);
  var rows = [
    ["Role", "AAR_MASTER",      "AAR Master"],
    ["Role", "MONO_AAR",        "Mono AAR"],
    ["Role", "RETAIL_LEAD",     "Retail Lead"],
    ["Role", "BUSINESS_EXPERT", "Business Expert"],
    ["Role", "CREATIVE_PRO",    "Creative Pro"],
    ["Cert", "IPHONE_CERT",     "iPhone 檢定"],
    ["Cert", "IPAD_CERT",       "iPad 檢定"],
    ["Cert", "MAC_CERT",        "Mac 檢定"],
    ["Cert", "WATCH_CERT",      "Apple Watch 檢定"],
    ["Cert", "AIRPODS_CERT",    "AirPods 檢定"],
    ["Cert", "TV_CERT",         "Apple TV 檢定"]
  ];
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return sheet;
}

function getInitialDataForUser(activeEmail, userRow) {
  ensureFactoryOptionsSheet();
  var users             = getSheetDataAsObjects("Users");
  var factoryOptionsData = getSheetDataAsObjects("FactoryOptions") || [];
  var playlists         = getSheetDataAsObjects("Playlists");
  var videos            = getSheetDataAsObjects("Videos");
  var channelCategories = getSheetDataAsObjects("ChannelCategories");
  var channelNames      = getSheetDataAsObjects("ChannelNames");
  var storeNames        = getSheetDataAsObjects("StoreNames");
  var categoriesData    = getSheetDataAsObjects("Categories");
  var subCategoriesData = getSheetDataAsObjects("SubCategories");
  var storeRolesData    = getSheetDataAsObjects("StoreRoles");

  var currentUser = {
    email: userRow.Email, name: userRow.Name,
    channelCategory: userRow.ChannelCategory, channelName: userRow.ChannelName,
    storeName: userRow.StoreName, storeRole: userRow.StoreRole,
    level: userRow.Level || 'L0'
  };

  var formattedVideos = [];
  for (var i = 0; i < videos.length; i++) {
    var v = videos[i];
    if (!v.ID || v.ID === "") continue;
    formattedVideos.push({
      id: v.ID, title: v.Title, category: v.Category, subCategory: v.SubCategory,
      url: v.YouTubeURL, status: v.Status, allowedRoles: v.AllowedRoles,
      isRequired: v.IsRequired, createdAt: v.CreatedAt || ""
    });
  }

  var categories = [];
  for (var i = 0; i < categoriesData.length; i++) {
    if (categoriesData[i].CategoryName && categoriesData[i].CategoryName.trim() !== "") {
      categories.push(categoriesData[i].CategoryName.trim());
    }
  }

  var subCategories = [];
  for (var i = 0; i < subCategoriesData.length; i++) {
    if (subCategoriesData[i].SubCategoryName && subCategoriesData[i].ParentCategoryName) {
      subCategories.push({
        name: subCategoriesData[i].SubCategoryName.trim(),
        parent: subCategoriesData[i].ParentCategoryName.trim(),
        order: i
      });
    }
  }

  var roles = [];
  for (var i = 0; i < storeRolesData.length; i++) {
    if (storeRolesData[i].RoleName && storeRolesData[i].RoleName.trim() !== "") {
      roles.push(storeRolesData[i].RoleName.trim());
    }
  }

  var formattedPlaylists = [];
  for (var i = 0; i < playlists.length; i++) {
    var p = playlists[i];
    if (!p.ID) continue;
    formattedPlaylists.push({
      id: p.ID, name: p.PlaylistName, targetLevel: p.TargetLevel,
      isRequired: p.IsRequired, sortOrder: p.SortOrder || 99, videoList: p.VideoList
    });
  }

  var allFavs = getSheetDataAsObjects("Favorites");
  var userFavorites = [];
  for (var i = 0; i < allFavs.length; i++) {
    if (allFavs[i].Email && allFavs[i].Email.toLowerCase() === activeEmail) {
      userFavorites.push(allFavs[i].VideoID);
    }
  }

  var allLogs = getSheetDataAsObjects("WatchLog");
  var watchLogs = [], adminWatchLogs = [];
  for (var i = 0; i < allLogs.length; i++) {
    var logData = {
      videoId:       allLogs[i].VideoID,
      viewCount:     parseInt(allLogs[i].ViewCount    || 0, 10),
      watchDuration: parseInt(allLogs[i].WatchDuration|| 0, 10),
      totalDuration: parseInt(allLogs[i].TotalDuration|| 0, 10),
      maxReachedTime:parseFloat(allLogs[i].MaxReachedTime || 0),
      isCompleted:   allLogs[i].IsCompleted || '否',
      lastWatchedAt: allLogs[i].LastWatchedAt
    };
    adminWatchLogs.push(Object.assign({ email: allLogs[i].Email }, logData));
    if (allLogs[i].Email && allLogs[i].Email.toLowerCase() === activeEmail) {
      watchLogs.push(logData);
    }
  }

  var allEvents = getSheetDataAsObjects("WatchEvents");
  var watchEvents = [];
  for (var i = 0; i < allEvents.length; i++) {
    watchEvents.push({
      eventId:  allEvents[i].EventID || "",
      email:    allEvents[i].Email   || "",
      videoId:  allEvents[i].VideoID || "",
      duration: parseInt(allEvents[i].SessionDuration || 0, 10),
      createdAt:allEvents[i].CreatedAt || ""
    });
  }

  var adminUsersFormatted = [];
  for (var i = 0; i < users.length; i++) {
    adminUsersFormatted.push({
      email: users[i].Email, name: users[i].Name,
      channelCategory: users[i].ChannelCategory, channelName: users[i].ChannelName,
      storeName: users[i].StoreName, storeRole: users[i].StoreRole,
      level: users[i].Level, hireDate: users[i].HireDate,
      factoryRoles: users[i].FactoryRoles || '',
      certifications: users[i].Certifications || ''  
    });
  }

  return {
    status: "success",
    data: {
      user: currentUser, categories: categories, subCategories: subCategories,
      videos: formattedVideos, playlists: formattedPlaylists,
      favorites: userFavorites, watchLogs: watchLogs,
      adminWatchLogs: adminWatchLogs, watchEvents: watchEvents,
      roles: roles,
      channelCategories: channelCategories.map(function(c) { return c.CategoryName; }),
      channelNames: channelNames.map(function(c) { return { name: c.ChannelName, parent: c.ParentCategory }; }),
      storeNames:   storeNames.map(function(s)   { return { name: s.StoreName,   parent: s.ParentChannel  }; }),
      adminUsers: adminUsersFormatted,
      factoryOptions: factoryOptionsData.map(function(o){ return { type: String(o.Type||''), code: String(o.Code||''), label: String(o.Label||'') }; }),
      adminFavorites: allFavs.map(function(f){ return { email: (f.Email||'').toLowerCase(), videoId: f.VideoID }; })
    }
  };
}

// ==========================================
//  儲存觀看紀錄
// ==========================================
function saveWatchLog(email, sessionData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("WatchLog");
    if (!sheet) throw new Error("找不到 WatchLog 工作表");

    var data        = sheet.getDataRange().getValues();
    var headers     = data[0];
    var emailIdx    = headers.indexOf("Email");
    var vidIdx      = headers.indexOf("VideoID");
    var viewCountIdx= headers.indexOf("ViewCount");
    var watchDurIdx = headers.indexOf("WatchDuration");
    var totalDurIdx = headers.indexOf("TotalDuration");
    var maxTimeIdx  = headers.indexOf("MaxReachedTime");
    var isCompIdx   = headers.indexOf("IsCompleted");
    var lastWatchIdx= headers.indexOf("LastWatchedAt");

    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][emailIdx] === email && data[i][vidIdx] === sessionData.videoId) {
        foundRow = i + 1; break;
      }
    }

    var timestamp = new Date().toISOString().replace('T', ' ').substring(0, 19);
    var targetCompletedStr = sessionData.isCompleted ? '是' : '否';

    if (foundRow > -1) {
      var currentViews   = parseInt(data[foundRow-1][viewCountIdx] || 0) + 1;
      var currentDur     = parseInt(data[foundRow-1][watchDurIdx]  || 0) + sessionData.duration;
      var maxTotalDur    = Math.max(parseInt(data[foundRow-1][totalDurIdx] || 0), sessionData.totalDuration || 0);
      var currentMaxTime = 0;
      if (maxTimeIdx > -1) {
        currentMaxTime = Math.max(parseFloat(data[foundRow-1][maxTimeIdx]) || 0, sessionData.maxReachedTime || 0);
      }
      var finalCompleted = data[foundRow-1][isCompIdx] === '是' ? '是' : targetCompletedStr;
      sheet.getRange(foundRow, viewCountIdx + 1).setValue(currentViews);
      sheet.getRange(foundRow, watchDurIdx  + 1).setValue(currentDur);
      sheet.getRange(foundRow, totalDurIdx  + 1).setValue(maxTotalDur);
      if (maxTimeIdx > -1) sheet.getRange(foundRow, maxTimeIdx + 1).setValue(currentMaxTime);
      sheet.getRange(foundRow, isCompIdx    + 1).setValue(finalCompleted);
      sheet.getRange(foundRow, lastWatchIdx + 1).setValue(timestamp);
    } else {
      var newRow = new Array(headers.length);
      newRow[emailIdx]     = email;
      newRow[vidIdx]       = sessionData.videoId;
      newRow[viewCountIdx] = 1;
      newRow[watchDurIdx]  = sessionData.duration;
      newRow[totalDurIdx]  = sessionData.totalDuration || 0;
      if (maxTimeIdx > -1) newRow[maxTimeIdx] = sessionData.maxReachedTime || 0;
      newRow[isCompIdx]    = targetCompletedStr;
      newRow[lastWatchIdx] = timestamp;
      sheet.appendRow(newRow);
    }

    // 同步寫入 WatchEvents 事件表
    var eventSheet = ss.getSheetByName("WatchEvents");
    if (eventSheet) {
      var evHeaders    = eventSheet.getDataRange().getValues()[0];
      var evRow        = new Array(evHeaders.length);
      var evIdIdx      = evHeaders.indexOf("EventID");
      var evEmailIdx   = evHeaders.indexOf("Email");
      var evVidIdx     = evHeaders.indexOf("VideoID");
      var evDurIdx     = evHeaders.indexOf("SessionDuration");
      var evCreatedIdx = evHeaders.indexOf("CreatedAt");
      if (evIdIdx      > -1) evRow[evIdIdx]      = "E" + new Date().getTime() + Math.random().toString(36).substring(2, 5).toUpperCase();
      if (evEmailIdx   > -1) evRow[evEmailIdx]   = email;
      if (evVidIdx     > -1) evRow[evVidIdx]     = sessionData.videoId;
      if (evDurIdx     > -1) evRow[evDurIdx]     = sessionData.duration;
      if (evCreatedIdx > -1) evRow[evCreatedIdx] = timestamp;
      eventSheet.appendRow(evRow);
    }

    return { status: "success" };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ==========================================
//  切換收藏
// ==========================================
function toggleFavorite(email, videoId, isFav) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Favorites");
    if (!sheet) throw new Error("找不到 Favorites 工作表");
    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][1] === videoId) { foundRow = i + 1; break; }
    }
    if (!isFav && foundRow > -1) sheet.deleteRow(foundRow);
    else if (isFav && foundRow === -1) sheet.appendRow([email, videoId]);
    return { status: "success" };
  } catch (e) { return { status: "error", message: e.message }; }
}

// ==========================================
//  通用工具：Upsert / Delete 列
// ==========================================
function upsertRow(sheetName, idField, idValue, rowDataObj) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error("找不到工作表: " + sheetName);
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf(idField);
  var foundRow = -1;
  if (idIdx > -1) {
    for (var i = 1; i < data.length; i++) {
      if (data[i][idIdx] === idValue) { foundRow = i + 1; break; }
    }
  }
  var newRow = new Array(headers.length);
  for (var j = 0; j < headers.length; j++) {
    var h = headers[j];
    if (rowDataObj.hasOwnProperty(h))  { newRow[j] = rowDataObj[h]; }
    else if (foundRow > -1)            { newRow[j] = data[foundRow-1][j]; }
    else                               { newRow[j] = ""; }
  }
  if (foundRow > -1) sheet.getRange(foundRow, 1, 1, headers.length).setValues([newRow]);
  else sheet.appendRow(newRow);
}

function deleteRowById(sheetName, idField, idValue) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var idIdx = data[0].indexOf(idField);
  if (idIdx === -1) return;
  for (var i = 1; i < data.length; i++) {
    if (data[i][idIdx] === idValue) { sheet.deleteRow(i + 1); return; }
  }
}

// ==========================================
//  後台 CRUD：影片
// ==========================================
function saveAdminVideo(email, data) {
  try {
    upsertRow("Videos", "ID", data.id, {
      "ID": data.id, "Title": data.title, "Category": data.category,
      "SubCategory": data.subCategory, "YouTubeURL": data.url,
      "Status": data.status, "AllowedRoles": data.allowedRoles,
      "IsRequired": data.isRequired, "CreatedAt": data.createdAt
    });
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function deleteAdminVideo(email, videoId) {
  try { deleteRowById("Videos", "ID", videoId); return { status: 'success' }; }
  catch(e) { return { status: 'error', message: e.message }; }
}

// ==========================================
//  後台 CRUD：播放列表
// ==========================================
function saveAdminPlaylist(email, data) {
  try {
    upsertRow("Playlists", "ID", data.id, {
      "ID": data.id, "PlaylistName": data.name, "TargetLevel": data.targetLevel,
      "IsRequired": data.isRequired, "SortOrder": data.sortOrder, "VideoList": data.videoList
    });
    syncVideosIsRequired();
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function deleteAdminPlaylist(email, id) {
  try {
    deleteRowById("Playlists", "ID", id);
    syncVideosIsRequired();
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

// ==========================================
//  後台 CRUD：帳號
// ==========================================
function saveAdminUser(email, data) {
  try {
    upsertRow("Users", "Email", data.email, {
      "Email": data.email, "Password": data.password, "Name": data.name,
      "HireDate": data.hireDate, "ChannelCategory": data.channelCategory,
      "ChannelName": data.channelName, "StoreName": data.storeName,
      "StoreRole": data.storeRole, "Level": data.level,
      "FactoryRoles": String(data.factoryRoles || ''),
      "Certifications": String(data.certifications || '')
    });
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function deleteAdminUser(email, targetEmail) {
  try { deleteRowById("Users", "Email", targetEmail); return { status: 'success' }; }
  catch(e) { return { status: 'error', message: e.message }; }
}

// ==========================================
//  同步影片的 IsRequired 狀態
// ==========================================
function syncVideosIsRequired() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var plSheet  = ss.getSheetByName("Playlists");
  var vidSheet = ss.getSheetByName("Videos");
  if (!plSheet || !vidSheet) return;

  var plData    = plSheet.getDataRange().getValues();
  var vidData   = vidSheet.getDataRange().getValues();
  var plHeaders = plData[0];
  var reqIdx    = plHeaders.indexOf("IsRequired");
  var listIdx   = plHeaders.indexOf("VideoList");

  var requiredVids = {};
  for (var i = 1; i < plData.length; i++) {
    if (plData[i][reqIdx] === '是' && plData[i][listIdx]) {
      var vids = plData[i][listIdx].toString().split(',');
      for (var j = 0; j < vids.length; j++) requiredVids[vids[j].trim()] = true;
    }
  }

  var vHeaders = vidData[0];
  var vIdIdx   = vHeaders.indexOf("ID");
  var vReqIdx  = vHeaders.indexOf("IsRequired");
  if (vIdIdx === -1 || vReqIdx === -1) return;

  for (var i = 1; i < vidData.length; i++) {
    var vid   = vidData[i][vIdIdx];
    var isReq = requiredVids[vid] ? '是' : '否';
    if (vidData[i][vReqIdx] !== isReq) {
      vidSheet.getRange(i + 1, vReqIdx + 1).setValue(isReq);
    }
  }
}
