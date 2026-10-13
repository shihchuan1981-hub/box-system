const SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ═══ 系統密碼 ═══
// 密碼存在「專案設定 → 指令碼屬性」的 APP_PASSWORD，不寫在程式碼裡
// 連續錯 20 次會鎖 10 分鐘，避免被一直亂試
function checkKey(k) {
  const pass = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if (!pass) return '系統還沒設定密碼：請到 Apps Script「專案設定 → 指令碼屬性」新增 APP_PASSWORD';
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('keyFails') || 0);
  if (fails >= 20) return '密碼錯太多次，請 10 分鐘後再試';
  if (String(k || '').trim() === String(pass).trim()) return '';
  cache.put('keyFails', String(fails + 1), 600);
  return '密碼錯誤';
}
function denied(msg) {
  return ContentService.createTextOutput(JSON.stringify({success:false, auth:true, error:msg}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const bad = checkKey(data.key);
    if (bad) return denied(bad);
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
      if (data.orderNo !== undefined) writeOrderNo(sheet, rowNo, data.orderNo);
      // 對到產品時，把層數、印刷帶進看板欄位（Q、R 欄）
      if (data.layer) sheet.getRange(rowNo, BOARD_START + 1).setValue(String(data.layer));
      if (data.print) sheet.getRange(rowNo, BOARD_START + 2).setValue(data.print);
      // 斬盒→「斬盒」、隔板→「平板紙板」（T 欄成型方式，不經釘箱站）；空字串＝改回一般箱
      if (data.form !== undefined && data.form !== null) sheet.getRange(rowNo, BOARD_START + 4).setValue(data.form);
      if (data.productId && rowIndex < 0) { try { touchProduct(data.productId); } catch (e) {} }
      return ok(rowIndex >= 0 ? '訂單已更新' : '訂單已儲存');
    }

    if (action === 'saveProduct') {
      return ok(saveProduct(data.product));
    }

    if (action === 'readShipment') {
      return ok(readShipment(data.fileBase64, data.mediaType));
    }

    if (action === 'uploadPhoto') {
      return ok(uploadPhoto(data.productId, data.image));
    }

    if (action === 'setProductPrint') {
      return ok(setProductPrint(data.productId, data.print));
    }

    if (action === 'addAlias') {
      return ok(addAlias(data.productId, data.alias));
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
      const ids = (data.orderIds || [data.orderId]).map(String);
      for (let i = 1; i < rows.length; i++) {
        if (ids.indexOf(String(rows[i][0])) >= 0) sheet.getRange(i+1, 15).setValue(data.status);
      }
      return ok('狀態已更新');
    }

    if (action === 'setSchedule') {
      return ok(setSchedule(data.orderIds || [data.orderId], data.schedDate, data.rush));
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

// ═══ 產品資料（從舊系統匯入，欄位照第一列標題） ═══
function prodSheet() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('產品資料');
  if (!sh) throw new Error('試算表裡找不到「產品資料」分頁');
  return sh;
}
function cellText(x) {
  if (x instanceof Date) return Utilities.formatDate(x, 'Asia/Taipei', 'yyyy/M/d');
  return x === null || x === undefined ? '' : String(x);
}

// 只回傳這一家客戶的產品（用客戶ID或名稱找）
function getProducts(custId, custName) {
  const v = prodSheet().getDataRange().getValues();
  const h = v[0].map(String);
  const ci = h.indexOf('客戶ID'), ni = h.indexOf('客戶名稱');
  const out = [];
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    if ((custId && String(r[ci]) === custId) || (custName && String(r[ni]) === custName)) {
      const o = {};
      h.forEach((k, j) => { if (k && k !== '才單價' && k !== '工資') o[k] = cellText(r[j]); });   // 價格不送到網頁
      out.push(o);
    }
  }
  return out;
}

function findProductRow(sh, id) {
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 1;
  return -1;
}

// 新產品：自動給下一個產品ID，存在最後一列
function saveProduct(p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = prodSheet();
    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
    let max = 0;
    ids.forEach(r => { const m = String(r[0]).match(/^P(\d+)$/); if (m) max = Math.max(max, +m[1]); });
    p['產品ID'] = 'P' + ('0000' + (max + 1)).slice(-5);
    if (!p['最後出貨']) p['最後出貨'] = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/M/d');
    if (p['出貨次數'] === undefined) p['出貨次數'] = 0;
    const row = h.map(k => p[k] === undefined ? '' : p[k]);
    sh.appendRow(row);
    const rowNo = sh.getLastRow();
    sh.getRange(rowNo, 1, 1, h.length).setNumberFormat('@');   // 全部存成文字，品名像 1/2 才不會變日期
    sh.getRange(rowNo, 1, 1, h.length).setValues([row.map(String)]);
    return p;
  } finally {
    lock.releaseLock();
  }
}

// 記住客戶的其他寫法，下次辨識直接對到
function addAlias(id, alias) {
  alias = String(alias || '').trim();
  if (!alias) return '沒有寫法';
  const sh = prodSheet();
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = h.indexOf('其他寫法') + 1;
  const rowNo = findProductRow(sh, id);
  if (rowNo < 0 || col < 1) return '找不到產品';
  const cur = String(sh.getRange(rowNo, col).getValue() || '');
  const list = cur ? cur.split(/[｜|]/) : [];
  if (list.indexOf(alias) < 0) { list.push(alias); sh.getRange(rowNo, col).setNumberFormat('@').setValue(list.join('｜')); }
  return '已記住';
}

// 接單時老闆選了有印／不印，記回產品資料，下次就不用再選
function setProductPrint(id, v) {
  if (['有印', '不印'].indexOf(v) < 0) throw new Error('印刷只能是 有印 或 不印');
  const sh = prodSheet();
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const col = h.indexOf('印刷') + 1, rowNo = findProductRow(sh, id);
  if (col < 1 || rowNo < 0) return '找不到產品';
  sh.getRange(rowNo, col).setValue(v);
  return '已更新';
}

// 下單後更新最後出貨日、出貨次數（常做的產品才會排前面）
function touchProduct(id) {
  const sh = prodSheet();
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rowNo = findProductRow(sh, id);
  if (rowNo < 0) return;
  const lc = h.indexOf('最後出貨') + 1, nc = h.indexOf('出貨次數') + 1;
  if (lc) sh.getRange(rowNo, lc).setNumberFormat('@').setValue(Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/M/d'));
  if (nc) sh.getRange(rowNo, nc).setValue((Number(sh.getRange(rowNo, nc).getValue()) || 0) + 1);
}

// ═══ 出貨麥頭照片 ═══
// 照片存在雲端硬碟「麥頭照片」資料夾，產品資料「照片」欄記「檔案ID@日期」，新的在前面，用｜隔開
// 第一次使用：在編輯器選 authorizeDrive → 執行 → 允許存取雲端硬碟
function authorizeDrive() {
  photoFolder();
  try { SpreadsheetApp.getUi().alert('雲端硬碟已授權，麥頭照片資料夾已建立。'); } catch (e) {}
}
function photoFolder() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName('麥頭照片');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('麥頭照片');
  props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  return folder;
}

function uploadPhoto(productId, image) {
  const m = String(image || '').match(/^data:(image\/\w+);base64,(.+)$/);
  if (!m) throw new Error('照片格式不對');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = prodSheet();
    const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const rowNo = findProductRow(sh, productId);
    if (rowNo < 0) throw new Error('找不到產品：' + productId);
    const row = sh.getRange(rowNo, 1, 1, h.length).getValues()[0];
    const today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/M/d');
    const name = [row[h.indexOf('客戶名稱')], row[h.indexOf('品名')], today.replace(/\//g, '-')].join('_').replace(/[\\/:*?"<>|]/g, '-') + '.jpg';
    const file = photoFolder().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));
    const pc = h.indexOf('照片') + 1, ic = h.indexOf('印刷') + 1;
    const cur = String(row[pc - 1] || '');
    const val = [file.getId() + '@' + today].concat(cur ? cur.split('｜') : []).join('｜');
    sh.getRange(rowNo, pc).setNumberFormat('@').setValue(val);
    if (ic && !String(row[ic - 1] || '').trim()) sh.getRange(rowNo, ic).setValue('有印');   // 有麥頭照片＝有印
    return {photo: val};
  } finally {
    lock.releaseLock();
  }
}

// 只給「麥頭照片」資料夾裡的檔案，其他雲端硬碟檔案一律不給
function getPhoto(id) {
  const file = DriveApp.getFileById(id);
  const folderId = photoFolder().getId();
  const parents = file.getParents();
  let inFolder = false;
  while (parents.hasNext()) if (parents.next().getId() === folderId) inFolder = true;
  if (!inFolder) throw new Error('不是麥頭照片');
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

// 讀出貨單（舊系統印的）：客戶、出貨單號、日期、每個品項
function readShipment(fileBase64, mediaType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('尚未設定 ANTHROPIC_API_KEY');
  const mt = mediaType === 'application/pdf' ? 'application/pdf' : (mediaType || 'image/jpeg');
  const prompt =
    '這是紙箱工廠的出貨單（照片可能是橫的或倒的）。請讀出：客戶名稱、出貨單號、出貨日期，以及表格每一列品項。\n' +
    '日期是民國年（例如 115/09/29），請換成西元 2026/9/29。\n' +
    '每一列：品名照抄（例如 638*12、HS330*12、A4232A*2），種類照抄（外箱、內箱、AB式…），規格照抄長*寬*高寫成「長×寬×高」，數量是數字。單價、小計不用。\n' +
    '手寫的勾、簽名不要理。看不清楚填空字串。\n' +
    '只回傳 JSON，不要說明：{"customer":"","ship_no":"","ship_date":"","items":[{"product":"","kind":"","size":"","qty":""}]}';
  const payload = {
    model: 'claude-opus-4-8',
    max_tokens: 4000,
    messages: [{role: 'user', content: [
      {type: mt === 'application/pdf' ? 'document' : 'image', source: {type: 'base64', media_type: mt, data: fileBase64}},
      {type: 'text', text: prompt}
    ]}]
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json',
    headers: {'x-api-key': apiKey, 'anthropic-version': '2023-06-01'},
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error(json.error.message || 'Claude API 呼叫失敗');
  const text = (json.content || []).map(c => c.text || '').join('');
  try { return JSON.parse(text.replace(/```json|```/g, '').trim()); }
  catch (e) { throw new Error('AI 回傳格式無法解析：' + text); }
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
    '   有些訂單的內盒和外箱是「上下兩張分開的表格」，每一列是一個品項，內盒和外箱的列數、ITEM NO 不一定對得上。\n' +
    '   這種訂單請「每一列各輸出一個物件」，不要合併、不要漏：內盒表的列只填 inner_* 和 inner_mark（outer_* 全部留空），外箱表的列只填 outer_*、front_mark、side_mark（inner_* 全部留空），product 填那一列的 ITEM NO。例如內盒 6 列、外箱 8 列就輸出 14 個物件。\n' +
    '   合併儲存格：一個尺寸格往下跨好幾列時，跨到的每一列都用同一個尺寸；表格右邊手寫的用紙（例如「=212×291×2」）也套用到它旁邊那個尺寸跨到的每一列。\n' +
    '   採購單格式（序號／品號／數量／單位／採購單價／金額）：「數量」欄才是數量（例如 15 PCS → 15），採購單價、金額、料號都不是數量。品名寫「外箱」就是外箱、「內盒」「內箱」就是內盒。\n' +
    '   外箱表的「訂單數量」那一欄是 outer_qty。同一列的 ITEM NO、QTY、N.W.、G.W.、CUFT 要填進那一列的 side_mark，例如「ITEM NO : 54925\\nQ\'TY : 10 EA\\nN.W. : 9.0 KGS\\nG.W. : 9.8 KGS\\nCUFT : 1.9」（單位照訂單側嘜區塊寫的 EA、PCS、KGS）。\n' +
    '   內盒表同一列的 ITEM NO、QTY 填進 inner_mark，例如「ITEM NO : 54332\\nQ\'TY : 10 EA」。「內盒數量」那格寫的是文字（例如「白色底有刀模」）時，inner_qty 留空，文字放 notes。\n' +
    '   表格最下面沒有 ITEM NO、但有數量和尺寸的列（例如「80」「(37+8)=45X168 三層隔板」）也要輸出一個物件：product 寫「三層隔板」這類說明，inner_qty 填數量，inner_material 填 45×168，inner_size 留空，notes 寫原文。\n' +
    '4. 日期常常是手寫的（訂單號碼、訂貨日期、交貨日期旁邊），請仔細讀手寫字。日期寫成 ' + year + '/月/日；交貨月份比訂貨月份小代表是隔年。日期一定要含年份。民國年（例如 115.10.27）要換成西元（115+1911=2026）。\n' +
    '   交貨日期通常在訂貨日期之後幾週到兩個月，手寫的「11」很容易看成「1」，請用這點判斷。\n' +
    '5. 品名照表格最左邊那格完整抄下來（含英文、數字、引號和中文，例如 3207I-L 15" 彩盒）。\n' +
    '6. 麥頭：正嘜、側嘜照原本換行抄，每行之間用 \\n。側嘜表格裡有填數字的，要把數字填進對應的標題後面（不要放到 notes）；沒填的變數（例如 Q\'TY:、N.W.:）照樣保留標題。正嘜是圖案 LOGO 時，寫出 LOGO 上的文字。\n' +
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
    max_tokens: 12000,
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
  const bad = checkKey(e.parameter.key);
  if (bad) return denied(bad);
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
  if (action === 'getPhoto') {
    try { return ok(getPhoto(e.parameter.id)); }
    catch (err) { return ContentService.createTextOutput(JSON.stringify({success:false, error:err.message})).setMimeType(ContentService.MimeType.JSON); }
  }
  if (action === 'getProducts') {
    try { return ok(getProducts(e.parameter.custId, e.parameter.custName)); }
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
                    '裁切進度','印刷進度','開槽進度','釘箱進度','出貨進度','小標張數','客戶單號'];
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
  dd(4, ['釘箱','糊箱','平板紙板','斬盒']);
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
    const isInner = !r[4] && !!(r[7] || r[8] || r[9]);
    const size = String(r[4] || r[7] || '');
    const qty = String(r[6] || r[9] || '');
    const kind = String(b[4] || '釘箱').trim();
    const layer = Number(b[1]) || (isInner || kind === '斬盒' || /內/.test(r[3]) ? 3 : 5);
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
      route: route, prog: prog, copies: Number(b[12]) || 1,
      orderNo: String(b[13] || ''), type: isInner ? '內盒' : '外箱'
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

// 一次排一整張訂單（多筆）
function setSchedule(orderIds, schedDate, rush) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('訂單紀錄');
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    const want = orderIds.map(String);
    let n = 0;
    for (let i = 1; i < ids.length; i++) {
      if (want.indexOf(String(ids[i][0])) >= 0) { writeSchedule(sh, i + 1, schedDate, rush); n++; }
    }
    if (!n) throw new Error('找不到訂單：' + want.join(', '));
    return '已排程 ' + n + ' 筆';
  } finally {
    lock.releaseLock();
  }
}

// 客戶的訂單號碼（例如 4677-2），排程看板用它把同一張訂單的品項合在一起
function writeOrderNo(sheet, rowNo, orderNo) {
  const col = BOARD_START + 13;
  if (!sheet.getRange(1, col).getValue()) sheet.getRange(1, col).setValue('客戶單號').setFontWeight('bold');
  sheet.getRange(rowNo, col).setValue(String(orderNo || ''));
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
