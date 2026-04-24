// ============================================================
//  聚餐好選擇 V2 - Google Apps Script 後端
//  Sheet 工作表：
//    settings   - 活動設定
//    restaurants - 餐廳清單
//    dishes      - 菜色清單
//    votes       - 投票記錄
//    dish_votes  - 菜色投票
// ============================================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
const ADMIN_PASSWORD = '54222486';

// ── 主路由 ──────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  let result;
  try {
    if      (action === 'getConfig')     result = getConfig();
    else if (action === 'getStats')      result = getStats();
    else if (action === 'checkVoter')    result = checkVoter(e.parameter.voter_id);
    else if (action === 'getVote')       result = getVote(e.parameter.voter_id);
    else if (action === 'getVotes')      result = adminGetVotes(e.parameter.pwd);
    else result = { status:'error', message:'Unknown action' };
  } catch(err) {
    result = { status:'error', message: err.toString() };
  }
  const cb = e.parameter.callback;
  const json = JSON.stringify(result);
  if (cb) return ContentService.createTextOutput(`${cb}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || '';
    if      (action === 'submitVote')       result = submitVote(data);
    else if (action === 'updateVote')       result = updateVote(data);
    else if (action === 'adminSaveConfig')  result = adminSaveConfig(data);
    else if (action === 'adminAddRestaurant') result = adminAddRestaurant(data);
    else if (action === 'adminUpdateRestaurant') result = adminUpdateRestaurant(data);
    else if (action === 'adminDeleteRestaurant') result = adminDeleteRestaurant(data);
    else if (action === 'adminAddDish')     result = adminAddDish(data);
    else if (action === 'adminDeleteDish')  result = adminDeleteDish(data);
    else if (action === 'adminDeleteVote')  result = adminDeleteVote(data);
    else if (action === 'adminResetVotes')  result = adminResetVotes(data);
    else result = { status:'error', message:'Unknown action' };
  } catch(err) {
    result = { status:'error', message: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ── 初始化 ──────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // settings
  if (!ss.getSheetByName('settings')) {
    const s = ss.insertSheet('settings');
    s.appendRow(['key','value']);
    s.appendRow(['event_name','聚餐好選擇 V2']);
    s.appendRow(['event_reasons','好久不見再續前緣,慶祝分條有成下次別鬧,恭喜升官發財']);
    s.appendRow(['months','2026-05,2026-06']);
    s.appendRow(['deadline','']);
    s.appendRow(['status','open']); // open / closed
    s.setFrozenRows(1);
  }

  // restaurants
  if (!ss.getSheetByName('restaurants')) {
    const s = ss.insertSheet('restaurants');
    s.appendRow(['id','name','stars','review_count','avg_cost','map_url','emoji']);
    s.setFrozenRows(1);
  }

  // dishes
  if (!ss.getSheetByName('dishes')) {
    const s = ss.insertSheet('dishes');
    s.appendRow(['restaurant_id','dish_name']);
    s.setFrozenRows(1);
  }

  // votes
  if (!ss.getSheetByName('votes')) {
    const s = ss.insertSheet('votes');
    s.appendRow(['timestamp','voter_id','name','company','drink','restaurant_id','busy_dates','sponsor']);
    s.setFrozenRows(1);
  }

  // dish_votes
  if (!ss.getSheetByName('dish_votes')) {
    const s = ss.insertSheet('dish_votes');
    s.appendRow(['timestamp','voter_id','restaurant_id','dish_name']);
    s.setFrozenRows(1);
  }

  return { status:'ok', message:'Sheets initialized' };
}

// ── 讀取前台設定 ─────────────────────────────────────────────
function getConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // settings
  const settingsMap = {};
  const sData = ss.getSheetByName('settings').getDataRange().getValues();
  for (let i = 1; i < sData.length; i++) {
    settingsMap[String(sData[i][0])] = String(sData[i][1]);
  }

  // restaurants
  const rData = ss.getSheetByName('restaurants').getDataRange().getValues();
  const restaurants = [];
  for (let i = 1; i < rData.length; i++) {
    const r = rData[i];
    if (!r[0]) continue;
    restaurants.push({
      id: String(r[0]), name: String(r[1]), stars: String(r[2]),
      review_count: String(r[3]), avg_cost: String(r[4]),
      map_url: String(r[5]), emoji: String(r[6]) || '🍽️'
    });
  }

  // dishes
  const dData = ss.getSheetByName('dishes').getDataRange().getValues();
  const dishes = {};
  for (let i = 1; i < dData.length; i++) {
    const rid = String(dData[i][0]);
    const dish = String(dData[i][1]);
    if (!rid || !dish) continue;
    if (!dishes[rid]) dishes[rid] = [];
    dishes[rid].push(dish);
  }

  // holidays（東立材料 2026）
  const holidays = [
    '2026-01-01','2026-02-15','2026-02-16','2026-02-17',
    '2026-02-18','2026-02-19','2026-02-20','2026-02-27',
    '2026-02-28','2026-04-03','2026-04-04','2026-04-05',
    '2026-04-06','2026-05-01','2026-06-19','2026-09-25',
    '2026-09-28','2026-10-09','2026-10-10','2026-10-25',
    '2026-10-26','2026-12-25'
  ];

  return {
    status: 'ok',
    event_name:    settingsMap['event_name']    || '聚餐好選擇 V2',
    event_reasons: (settingsMap['event_reasons'] || '').split(',').filter(Boolean),
    months:        (settingsMap['months']        || '').split(',').filter(Boolean),
    deadline:      settingsMap['deadline']       || '',
    vote_status:   settingsMap['status']         || 'open',
    restaurants,
    dishes,
    holidays
  };
}

// ── 投票 ─────────────────────────────────────────────────────
function submitVote(data) {
  if (!checkPwd(data)) {} // 前台投票不需密碼
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const ds = ss.getSheetByName('dish_votes');
  const voterId = data.voter_id;
  if (findVoteRow(vs, voterId) > 0) return { status:'duplicate' };
  const now = new Date();
  const newRow = vs.getLastRow() + 1;
  vs.getRange(newRow, 1, 1, 8).setValues([[
    now, voterId, data.name, data.company, data.drink,
    data.restaurant_id,
    (data.busy_dates || []).join(','),
    data.sponsor || ''
  ]]);
  vs.getRange(newRow, 7).setNumberFormat('@STRING@');
  (data.dishes || []).forEach(dish => {
    ds.appendRow([now, voterId, data.restaurant_id, dish]);
  });
  return { status:'ok' };
}

function updateVote(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const ds = ss.getSheetByName('dish_votes');
  const voterId = data.voter_id;
  const row = findVoteRow(vs, voterId);
  if (row > 0) vs.deleteRow(row);
  deleteDishRows(ds, voterId);
  const now = new Date();
  const newRow = vs.getLastRow() + 1;
  vs.getRange(newRow, 1, 1, 8).setValues([[
    now, voterId, data.name, data.company, data.drink,
    data.restaurant_id,
    (data.busy_dates || []).join(','),
    data.sponsor || ''
  ]]);
  vs.getRange(newRow, 7).setNumberFormat('@STRING@');
  (data.dishes || []).forEach(dish => {
    ds.appendRow([now, voterId, data.restaurant_id, dish]);
  });
  return { status:'ok' };
}

function checkVoter(voterId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  return { status:'ok', voted: findVoteRow(vs, voterId) > 0 };
}

function getVote(voterId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const ds = ss.getSheetByName('dish_votes');
  const row = findVoteRow(vs, voterId);
  if (row < 0) return { status:'not_found' };
  const r = vs.getRange(row, 1, 1, 8).getValues()[0];
  const dishes = getDishesForVoter(ds, voterId);
  return {
    status:'ok', data:{
      name: String(r[2]), company: String(r[3]), drink: String(r[4]),
      restaurant_id: String(r[5]),
      busy_dates: String(r[6]).split(',').filter(s => s.match(/^\d{4}-/)),
      dishes, sponsor: String(r[7])
    }
  };
}

// ── 統計 ─────────────────────────────────────────────────────
function getStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const ds = ss.getSheetByName('dish_votes');
  const cfg = getConfig();

  const vData = vs.getDataRange().getValues();
  const dData = ds.getDataRange().getValues();
  const totalVoters = Math.max(0, vData.length - 1);

  const restaurantCount = {};
  let drinkCount = 0;
  const busyMap = [];

  for (let i = 1; i < vData.length; i++) {
    const r = vData[i];
    const rid = String(r[5]);
    const busy = String(r[6]).split(',').filter(s => s.match(/^\d{4}-/));
    const drink = String(r[4]);
    const isGoing = !rid.startsWith('other');
    restaurantCount[rid] = (restaurantCount[rid] || 0) + 1;
    if (drink === '是') drinkCount++;
    busyMap.push({ busyDates: busy, going: isGoing });
  }

  // 最高票餐廳
  let topRestId = '', topCount = 0;
  Object.entries(restaurantCount).forEach(([rid, cnt]) => {
    if (!rid.startsWith('other') && cnt > topCount) { topCount = cnt; topRestId = rid; }
  });

  // 菜色統計
  const dishCount = {};
  for (let i = 1; i < dData.length; i++) {
    if (String(dData[i][2]) === topRestId) {
      const dish = String(dData[i][3]);
      dishCount[dish] = (dishCount[dish] || 0) + 1;
    }
  }
  const topDishes = Object.entries(dishCount).sort((a,b)=>b[1]-a[1]).slice(0,10)
    .map(([name,count]) => ({ name, count }));

  // 枚舉工作日
  const holidaySet = new Set(cfg.holidays || []);
  const allWorkDays = [];
  (cfg.months || []).forEach(ym => {
    const [y, m] = ym.split('-').map(Number);
    const days = new Date(y, m, 0).getDate();
    for (let d = 1; d <= days; d++) {
      const dateKey = `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
      const dow = new Date(y, m-1, d).getDay();
      if (dow !== 0 && dow !== 6 && !holidaySet.has(dateKey)) allWorkDays.push(dateKey);
    }
  });

  const goingPeople = busyMap.filter(p => p.going);
  const goingCount = goingPeople.length;
  const availableDates = allWorkDays.map(dateKey => {
    const count = goingPeople.filter(p => !p.busyDates.includes(dateKey)).length;
    return { date: dateKey, count, total: goingCount };
  }).filter(x => x.count > 0).sort((a,b) => b.count - a.count);

  const commonDates = availableDates.filter(x => x.count === goingCount && goingCount > 0).map(x => x.date);

  return {
    status:'ok', totalVoters, drinkCount, restaurantCount,
    topRestId, topRestaurantCount: topCount,
    topDishes, availableDates, commonDates
  };
}

// ── 後台：取得投票清單 ────────────────────────────────────────
function adminGetVotes(pwd) {
  if (pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const data = vs.getDataRange().getValues();
  const votes = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    votes.push({
      row: i + 1,
      voter_id: String(r[1]), name: String(r[2]), company: String(r[3]),
      drink: String(r[4]), restaurant_id: String(r[5]),
      busy_dates: String(r[6]), sponsor: String(r[7]),
      timestamp: String(r[0])
    });
  }
  return { status:'ok', votes };
}

// ── 後台：儲存設定 ────────────────────────────────────────────
function adminSaveConfig(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName('settings');
  const rows = s.getDataRange().getValues();
  const updates = {
    event_name:    data.event_name    || '',
    event_reasons: (data.event_reasons || []).join(','),
    months:        (data.months       || []).join(','),
    deadline:      data.deadline      || '',
    status:        data.vote_status   || 'open'
  };
  for (let i = 1; i < rows.length; i++) {
    const key = String(rows[i][0]);
    if (key in updates) s.getRange(i+1, 2).setValue(updates[key]);
  }
  return { status:'ok' };
}

// ── 後台：餐廳管理 ───────────────────────────────────────────
function adminAddRestaurant(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName('restaurants');
  const id = 'r_' + Date.now();
  s.appendRow([id, data.name, data.stars, data.review_count, data.avg_cost, data.map_url, data.emoji || '🍽️']);
  return { status:'ok', id };
}

function adminUpdateRestaurant(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName('restaurants');
  const rows = s.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      s.getRange(i+1, 1, 1, 7).setValues([[
        data.id, data.name, data.stars, data.review_count,
        data.avg_cost, data.map_url, data.emoji || '🍽️'
      ]]);
      return { status:'ok' };
    }
  }
  return { status:'error', message:'找不到餐廳' };
}

function adminDeleteRestaurant(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const rs = ss.getSheetByName('restaurants');
  const ds = ss.getSheetByName('dishes');
  const rows = rs.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.id)) { rs.deleteRow(i+1); break; }
  }
  // 刪除對應菜色
  const dRows = ds.getDataRange().getValues();
  for (let i = dRows.length - 1; i >= 1; i--) {
    if (String(dRows[i][0]) === String(data.id)) ds.deleteRow(i+1);
  }
  return { status:'ok' };
}

// ── 後台：菜色管理 ───────────────────────────────────────────
function adminAddDish(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheetByName('dishes').appendRow([data.restaurant_id, data.dish_name]);
  return { status:'ok' };
}

function adminDeleteDish(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const s = ss.getSheetByName('dishes');
  const rows = s.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.restaurant_id) && String(rows[i][1]) === String(data.dish_name)) {
      s.deleteRow(i+1); break;
    }
  }
  return { status:'ok' };
}

// ── 後台：刪除投票 ───────────────────────────────────────────
function adminDeleteVote(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const vs = ss.getSheetByName('votes');
  const ds = ss.getSheetByName('dish_votes');
  const voterId = data.voter_id;
  const row = findVoteRow(vs, voterId);
  if (row > 0) vs.deleteRow(row);
  deleteDishRows(ds, voterId);
  return { status:'ok' };
}

function adminResetVotes(data) {
  if (data.pwd !== ADMIN_PASSWORD) return { status:'error', message:'密碼錯誤' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  clearSheet(ss.getSheetByName('votes'));
  clearSheet(ss.getSheetByName('dish_votes'));
  return { status:'ok' };
}

// ── 工具函式 ─────────────────────────────────────────────────
function findVoteRow(sheet, voterId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(voterId)) return i + 1;
  }
  return -1;
}

function deleteDishRows(sheet, voterId) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][1]) === String(voterId)) sheet.deleteRow(i+1);
  }
}

function getDishesForVoter(sheet, voterId) {
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(r => String(r[1]) === String(voterId)).map(r => String(r[3]));
}

function clearSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
}

function checkPwd(data) {
  return data.pwd === ADMIN_PASSWORD;
}
