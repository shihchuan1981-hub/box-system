const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'saveOrder') {
      const sheet = ss.getSheetByName('訂單紀錄');
      const rows = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.orderId) { rowIndex = i; break; }
      }
      const status = data.status || (rowIndex >= 0 ? rows[rowIndex][14] : '待出貨');
      const rowData = [
        data.orderId, data.date, data.customer, data.product,
        data.outerSize, data.outerMaterial, data.outerQty,
        data.innerSize, data.innerMaterial, data.innerQty,
        data.delivery, data.frontMark, data.sideMark,
        data.notes, status
      ];
      let rowNo;
      if (rowIndex >= 0) {
        rowNo = rowIndex + 1;
        sheet.getRange(rowNo, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
        rowNo = sheet.getLastRow();
      }
      writeSchedule(sheet, rowNo, data.schedDate, data.rush);
      return ok(rowIndex >= 0 ? '訂單已更新' : '訂單已儲存');
    }

    if (action === 'saveCustomer') {
      return ok(saveCustomer(data));
    }

    if (action === 'deleteCustomer') {
      return ok(deleteCustomer(data.customerId, data.name));
    }

    if (action === 'updateOrderStatus') {
      const sheet = ss.getSheetByName('訂單紀錄');
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] == data.orderId) {
          sheet.getRange(i+1, 15).setValue(data.status);
          break;
        }
      }
      return ok('狀態已更新');
    }

    if (action === 'setSchedule') {
      return ok(setSchedule(data.orderId, data.schedDate, data.rush));
    }

    if (action === 'updateStation') {
      return ok(updateStation(data.orderId, data.station, data.state));
    }

    if (action === 'analyzeOrder') {
      return ok(analyzeOrder(data.fileBase64, data.mediaType));
    }

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success:false, error:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══ 客戶資料：新增／修改／刪除 ═══
// 找客戶在第幾列：有客戶ID就用ID找；舊資料沒有ID的，用名稱找
function findCustomerRow(sh, id, name) {
  const rows = sh.getDataRange().getValues();
  id = String(id || '').trim();
  name = String(name || '').trim();
  for (let i = 1; i < rows.length; i++) {
    const rid = String(rows[i][0] || '').trim();
    const rname = String(rows[i][1] || '').trim();
    if (id ? rid === id : (!rid && rname === name)) return i + 1;
  }
  return -1;
}

// 編輯 → 改原本那一列；新增 → 加在最後一列
function saveCustomer(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('客戶資料');
    const rowData = [
      d.customerId || ('C' + Date.now()), d.name, d.contact, d.phone,
      d.fax, d.address, d.taxId, d.logoNote,
      d.notes, d.frontMark, d.sideMark
    ];
    if (d.isEdit) {
      const rowNo = findCustomerRow(sh, d.originalId, d.originalName);
      if (rowNo < 0) throw new Error('找不到要修改的客戶「' + d.originalName + '」，可能已被刪除，請重新整理');
      sh.getRange(rowNo, 1, 1, rowData.length).setValues([rowData]);
      return '客戶已更新';
    }
    sh.appendRow(rowData);
    return '客戶已儲存';
  } finally {
    lock.releaseLock();
  }
}

function deleteCustomer(id, name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('客戶資料');
    const rowNo = findCustomerRow(sh, id, name);
    if (rowNo < 0) throw new Error('找不到客戶「' + name + '」，可能已經刪除了');
    // 再對一次名稱，避免刪錯
    if (String(sh.getRange(rowNo, 2).getValue()).trim() !== String(name).trim()) {
      throw new Error('客戶資料對不上，請重新整理後再刪');
    }
    sh.deleteRow(rowNo);
    return '客戶已刪除';
  } finally {
    lock.releaseLock();
  }
}

// ── AI 辨識：伺服器端呼叫 Claude API，金鑰不暴露給瀏覽器 ──
function analyzeOrder(fileBase64, mediaType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    throw new Error('尚未設定 ANTHROPIC_API_KEY，請到 Apps Script 專案設定 → Script Properties 新增');
  }
  const mt = mediaType === 'application/pdf' ? 'application/pdf' : (mediaType || 'image/jpeg');
  const contentType = mt === 'application/pdf' ? 'document' : 'image';
  const year = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy');
  const prompt =
    '你是台灣紙箱工廠的訂單辨識助手。這張是客戶的紙箱訂製單，可能有一筆或多筆品項。\n' +
    '請照下面規則讀，每個品項輸出一個物件：\n' +
    '1. 尺寸：只讀表格裡「長、寬、高」三格的數字，寫成「長×寬×高」，例如 152×82×142。\n' +
    '2. 用紙：尺寸下方常有一個框寫「A X B X N」（例如 122 X 233 X 2），這不是尺寸，是用紙尺寸 A×B，N 是一個箱子要用幾張紙。放到 material（寫成 A×B）和 sheets（寫數字 N）。沒有就填空字串。\n' +
    '3. 外箱和內盒（內箱）是兩個不同的紙箱，各自有尺寸、數量、用紙，不要混在一起。「外箱數量」放 outer_qty，「內盒數量」放 inner_qty。\n' +
    '   注意：表格裡標題寫「內盒數量」的那一欄才是 inner_qty（例如 300）。「內盒印刷」底下的 QTY 是每個內盒裝幾個產品（例如 10），不是 inner_qty，要寫進 inner_mark。外箱側嘜底下的 QTY 也不是外箱數量。\n' +
    '4. 日期常常是手寫的（訂單號碼、訂貨日期、交貨日期旁邊），請仔細讀手寫字。日期寫成 ' + year + '/月/日；交貨月份比訂貨月份小代表是隔年。日期一定要含年份。\n' +
    '   交貨日期通常在訂貨日期之後幾週到兩個月，手寫的「11」很容易看成「1」，請用這點判斷。\n' +
    '5. 品名照表格最左邊那格完整抄下來（含英文、數字、引號和中文，例如 3207I-L 15" 彩盒）。\n' +
    '6. 麥頭：正嘜、側嘜照原本換行抄，每行之間用 \\n。空白沒填的變數（例如 Q\'TY:、N.W.:）照樣保留標題。\n' +
    '7. 內盒印刷：訂單上有「內盒印刷」（表格欄位或下方區塊）就代表內盒要印刷，把內容寫進 inner_mark，例如「ITEM NO : LNL65107\\nQ\'TY : 10 PCS」（表格裡的值要填進去）。沒有內盒印刷就填空字串。\n' +
    '8. 看不清楚的欄位填空字串，不要猜。\n' +
    '只回傳 JSON 陣列，不要任何說明文字或 markdown：\n' +
    '[{"customer":"客戶名稱","order_no":"訂單號碼","order_date":"訂貨日期","delivery_date":"交貨日期",' +
    '"product":"品名","item_no":"Item No.","barcode":"條碼",' +
    '"outer_size":"外箱 長×寬×高","outer_qty":"外箱數量","outer_material":"外箱用紙 A×B","outer_sheets":"外箱每箱幾張",' +
    '"inner_size":"內盒 長×寬×高","inner_qty":"內盒數量","inner_material":"內盒用紙 A×B","inner_sheets":"內盒每個幾張",' +
    '"front_mark":"正嘜","side_mark":"側嘜","inner_mark":"內盒印刷內容","notes":"印刷要求與特殊備注"}]\n' +
    '客戶名稱、訂單號碼、日期等共用資訊請在每個品項裡都填上相同的值。';

  const payload = {
    model: 'claude-opus-4-8',
    max_tokens: 3000,
    messages: [{
      role: 'user',
      content: [
        { type: contentType, source: { type: 'base64', media_type: mt, data: fileBase64 } },
        { type: 'text', text: prompt }
      ]
    }]
  };

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error(json.error.message || 'Claude API 呼叫失敗');

  const text = (json.content || []).map(c => c.text || '').join('');
  try {
    const parsed = JSON.parse(text.replace(/```json|```/g, '').trim());
    return Array.isArray(parsed) ? parsed : [parsed];
  } catch (e) {
    throw new Error('AI 回傳格式無法解析：' + text);
  }
}

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getCustomers') {
    const rows = ss.getSheetByName('客戶資料').getDataRange().getValues();
    return ok(rows);
  }
  if (action === 'getBoard') {
    try { return ok(getBoard(e.parameter.date)); }
    catch (err) { return ContentService.createTextOutput(JSON.stringify({success:false, error:err.message})).setMimeType(ContentService.MimeType.JSON); }
  }
  if (action === 'getSchedule') {
    try { return ok(getSchedule()); }
    catch (err) { return ContentService.createTextOutput(JSON.stringify({success:false, error:err.message})).setMimeType(ContentService.MimeType.JSON); }
  }
  if (action === 'getOrders') {
    const rows = ss.getSheetByName('訂單紀錄').getDataRange().getValues();
    return ok(rows);
  }
  return ok([]);
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({success:true, data:data}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 初始化工作表標題 ──
// 第一次使用時執行這個函式（選單 → 執行 → initSheets）
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = {
    '客戶資料': ['客戶ID','客戶名稱','聯絡人','電話','傳真','地址','統一編號','Logo備注','備注','正嘜範本','側嘜範本'],
    '訂單紀錄': ['訂單ID','日期','客戶名稱','品名','外箱尺寸','外箱用料','外箱數量','內箱尺寸','內箱用料','內箱數量','交貨日期','正嘜','側嘜','備注','狀態'],
    '麥頭範本': ['客戶ID','客戶名稱','正嘜範本','側嘜範本','備注'],
    '排程':     ['訂單ID','客戶名稱','交貨日期','品名','數量','狀態','備注']
  };

  Object.entries(sheets).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  });

  SpreadsheetApp.getUi().alert('初始化完成！所有工作表已建立。');
}


// ═══ 工站看板 + 每日小標 ═══
// 訂單紀錄 P 欄以後是排程用的欄位（由 setupBoard 建立）

const BOARD_COLS = ['楞別','層數','印刷','開槽','成型方式','急件','排程日期',
                    '裁切進度','印刷進度','開槽進度','釘箱進度','出貨進度','小標張數'];
const BOARD_START = 16;          // P 欄 = 第 16 欄
const STATIONS = ['裁切','印刷','開槽','釘箱','出貨'];
const TZ = 'Asia/Taipei';

// ── 第一次使用：執行一次 setupBoard ──
function setupBoard() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄');
  sh.getRange(1, BOARD_START, 1, BOARD_COLS.length).setValues([BOARD_COLS]).setFontWeight('bold');
  const last = Math.max(sh.getMaxRows(), 500);
  const dd = (col, list) => sh.getRange(2, BOARD_START + col, last - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build());
  dd(0, ['AB楞','BC楞','A楞','B楞','C楞','E楞']);
  dd(1, ['5','3']);
  dd(2, ['是','否']);
  dd(3, ['是','否']);
  dd(4, ['釘箱','糊箱','平板紙板']);
  dd(5, ['是','否']);
  sh.getRange(2, BOARD_START + 6, last - 1, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build())
    .setNumberFormat('yyyy/m/d');
  sh.getRange(1, BOARD_START, 1, 7).setBackground('#fdf6e8');      // 老闆填的欄位
  sh.getRange(1, BOARD_START + 7, 1, 5).setBackground('#eef6ee');  // 平板自動寫的欄位
  try { SpreadsheetApp.getUi().alert('工站看板欄位已建立（P 欄到 AB 欄）。'); } catch (e) {}
}

// 日期統一成 2026-10-08；「10/9」這種沒寫年份的當作今年
function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = t.match(/^(\d{1,2})[\/.-](\d{1,2})$/);
  if (m) return Utilities.formatDate(new Date(), TZ, 'yyyy') + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return t;
}
function shortDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'M/d');
  return String(v || '').replace(/^\d{4}[\/-]/, '').trim();
}
function isYes(v) { return /^(是|y|yes|v|✓|1|true)$/i.test(String(v).trim()); }

// 從「A10外箱105×64×79」「193x118x113 台分」抓出三個數字
function parseSize(s) {
  const m = String(s || '').match(/(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)(?:\s*[×xX*]\s*(\d+(?:\.\d+)?))?/);
  if (!m) return null;
  return { l: +m[1], w: +m[2], h: m[3] ? +m[3] : null };
}
function calcMaterial(size, layer, kind) {
  const s = parseSize(size);
  if (!s) return '';
  if (kind === '平板紙板' || s.h === null) return s.l + '×' + s.w;
  return layer === 3 ? (s.w + s.h) + '×' + ((s.l + s.w) * 2 + 8)
                     : (s.w + s.h + 2) + '×' + ((s.l + s.w) * 2 + 12);
}

// ── 讀出看板要用的訂單 ──
// date：只要這一天排程的單（給小標用）
// 沒給 date：回傳排程日期在今天或以前、還沒出貨的單（給看板用，前幾天沒做完的也會留著）
function getBoard(date) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    while (r.length < BOARD_START - 1 + BOARD_COLS.length) r.push('');
    if (!r[0]) continue;
    const b = r.slice(BOARD_START - 1);
    const sched = fmtDate(b[6]);
    if (!sched) continue;
    const status = String(r[14] || '');
    if (date) { if (sched !== date) continue; }
    else if (sched > today || /已出貨|已取消/.test(status) || String(b[11]).indexOf('完成') === 0) continue;

    out.push(buildOrder(r, b, sched));
  }
  return out;
}

// 一列試算表 → 看板用的訂單資料
function buildOrder(r, b, sched) {
    const isInner = !r[4] && !!r[7];
    const size = String(r[4] || r[7] || '');
    const qty = String(r[6] || r[9] || '');
    const kind = String(b[4] || '釘箱').trim();
    const layer = Number(b[1]) || (isInner || /內/.test(r[3]) ? 3 : 5);
    const hasMark = String(r[11] || '').trim() || String(r[12] || '').trim();
    const print = b[2] === '' ? !!hasMark : isYes(b[2]);
    const slot = b[3] === '' ? kind !== '平板紙板' : isYes(b[3]);
    const route = [0];
    if (print) route.push(1);
    if (slot) route.push(2);
    if (kind === '釘箱') route.push(3);
    route.push(4);
    const prog = {};
    STATIONS.forEach((s, k) => {
      const v = String(b[7 + k] || '');
      if (v.indexOf('完成') === 0) prog[k] = 'done';
      else if (v.indexOf('進行中') === 0) prog[k] = 'doing';
    });
    return {
      no: String(r[0]), cust: String(r[2] || ''), prod: String(r[3] || ''),
      size: size, qty: qty, due: shortDate(r[10]), dueKey: fmtDate(r[10]),
      flute: String(b[0] || ''), layer: layer, kind: kind,
      material: String(r[isInner ? 8 : 5] || '') || calcMaterial(size, layer, kind),
      rush: isYes(b[5]), sched: sched, note: String(r[13] || ''),
      route: route, prog: prog, copies: Number(b[12]) || 1
    };
}

// ── 排程看板：所有還沒出貨的單（含還沒排程的） ──
function getSchedule() {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄').getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    while (r.length < BOARD_START - 1 + BOARD_COLS.length) r.push('');
    if (!r[0]) continue;
    if (/已出貨|已取消/.test(String(r[14] || ''))) continue;
    const b = r.slice(BOARD_START - 1);
    if (String(b[11]).indexOf('完成') === 0) continue;
    out.push(buildOrder(r, b, fmtDate(b[6])));
  }
  return out;
}

// 寫排程日期（'' = 取消排程）和急件；undefined 代表不動
function writeSchedule(sheet, rowNo, schedDate, rush) {
  if (schedDate !== undefined && schedDate !== null) {
    const v = schedDate ? new Date(schedDate + 'T00:00:00+08:00') : '';
    sheet.getRange(rowNo, BOARD_START + 6).setValue(v);
  }
  if (rush !== undefined && rush !== null) {
    sheet.getRange(rowNo, BOARD_START + 5).setValue(rush ? '是' : '');
  }
}

function setSchedule(orderId, schedDate, rush) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄');
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) {
        writeSchedule(sh, i + 1, schedDate, rush);
        return '已排程';
      }
    }
    throw new Error('找不到訂單：' + orderId);
  } finally {
    lock.releaseLock();
  }
}

// ── 平板按「開始／完成」 ──
function updateStation(orderId, station, state) {
  const k = STATIONS.indexOf(station);
  if (k < 0) throw new Error('沒有這個工站：' + station);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄');
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderId)) {
        const time = Utilities.formatDate(new Date(), TZ, 'M/d HH:mm');
        const text = state === 'done' ? '完成 ' + time : state === 'doing' ? '進行中 ' + time : '';
        sh.getRange(i + 1, BOARD_START + 7 + k).setValue(text);
        if (station === '出貨') sh.getRange(i + 1, 15).setValue(state === 'done' ? '已出貨' : '待出貨');
        return '已更新';
      }
    }
    throw new Error('找不到訂單：' + orderId);
  } finally {
    lock.releaseLock();
  }
}
