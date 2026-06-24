// ════════════════════════════════════════════════════════════
//  King's Corner — Google Apps Script 完整 API
//  部署方式：擴充功能 → Apps Script → 部署 → 網頁應用程式
//  執行身分：我（你的帳號）
//  存取權限：所有人（含匿名）
// ════════════════════════════════════════════════════════════

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID'; // ← 換成你的 Sheets ID

// 工作表名稱
const SH = {
  ACCOUNTS:   '帳號',
  SUPPLIERS:  '廠商',
  PRODUCTS:   '商品庫',
  CAMPAIGNS:  '團購',
  ORDERS:     '訂單',
  ADS:        '廣告',
  MARKETING:  '行銷設定',
  PROMO_CODES:'優惠碼',
};

// ── 工具函式 ──────────────────────────────────────────────
function ss()  { return SpreadsheetApp.openById(SHEET_ID); }
function ts()  { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function ok(data)  { return jsonOut({ status:'ok', ...data }); }
function err(msg)  { return jsonOut({ status:'error', message: msg }); }
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 讀取工作表 → [{col:val, ...}, ...]
function sheetToObjects(name) {
  const sheet = ss().getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      // Normalize: Sheets stores booleans as true/false (JS boolean), stringify them
      if (typeof v === 'boolean') obj[h] = String(v);
      else obj[h] = v ?? '';
    });
    return obj;
  }).filter(o => o[headers[0]] !== '' && o[headers[0]] !== null); // skip empty rows
}

// 確保工作表存在，若無則建立並設標題列
function ensureSheet(name, headers) {
  const spreadsheet = ss();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    // New sheet — write all headers
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#12100E')
      .setFontColor('#C9A84C');
    sheet.setFrozenRows(1);
  } else {
    // Existing sheet — add any missing columns
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(h => {
      if (!existing.includes(h)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(h);
        sheet.getRange(1, newCol)
          .setFontWeight('bold')
          .setBackground('#12100E')
          .setFontColor('#C9A84C');
        existing.push(h);
      }
    });
  }
  return sheet;
}

// 寫入或更新一列（依 idField 比對）
// 注意：依照 Sheets 實際欄位順序寫入，不假設與 headers 陣列順序一致
function upsertRow(sheetName, headers, record, idField) {
  const sheet = ensureSheet(sheetName, headers);
  const data = sheet.getDataRange().getValues();
  const hRow = data[0]; // Actual sheet headers (may differ from headers param)
  const idCol = hRow.indexOf(idField);

  // Build row values based on ACTUAL sheet column order
  const totalCols = Math.max(hRow.length, headers.length);
  function buildRow(actualHeaders) {
    return actualHeaders.map(h => {
      if (h === '' || h === undefined) return '';
      return record[h] !== undefined ? record[h] : '';
    });
  }

  // Update existing row
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(record[idField])) {
      const rowVals = buildRow(hRow);
      sheet.getRange(i + 1, 1, 1, hRow.length).setValues([rowVals]);
      return;
    }
  }
  // Insert new row — use actual headers after ensureSheet added any missing cols
  const freshHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowVals = buildRow(freshHeaders);
  sheet.appendRow(rowVals);
}

// 刪除一列（依 idField 比對）
function deleteRow(sheetName, idField, idValue) {
  const sheet = ss().getSheetByName(sheetName);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf(idField);
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === String(idValue)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// ── 初始化預設資料 ──────────────────────────────────────────
function initSheets() {
  // 帳號表
  const accSheet = ensureSheet(SH.ACCOUNTS, ['id','user','passHash','name','role','active','createdAt']);
  if (accSheet.getLastRow() < 2) {
    // 預設 root 帳號（密碼：root1234）
    accSheet.appendRow(['1', 'root', hashPass('root1234'), 'Timmy', 'root', 'true', ts()]);
    accSheet.appendRow(['2', 'admin', hashPass('admin1234'), '管理員', 'admin', 'true', ts()]);
  }

  // 其他工作表預建
  ensureSheet(SH.SUPPLIERS, ['id','name','cat','contact','note','createdAt']);
  ensureSheet(SH.PRODUCTS,  ['id','supplierId','name','cat','spec','price','img','active']);
  ensureSheet(SH.CAMPAIGNS, ['id','name','supplierId','deadline','minPeople','discount','status','closedStatus','pickupDate','note','products','allowedCodes','createdAt']);
  ensureSheet(SH.ORDERS,    ['id','campaignId','campaignName','supplierId','name','phone','items','itemCount','total','discounted','paid','ts']);
  ensureSheet(SH.ADS,       ['supplierId','slides']); // slides 存 JSON
  ensureSheet(SH.MARKETING,   ['key','value']);
  ensureSheet(SH.PROMO_CODES, ['id','code','name','contact','discountType','discountValue','campaignRules','status','note','createdAt']);
  // Add promoCode + allowedCodes to orders/campaigns if not exists (handled dynamically)
}

// 簡單 hash（非加密用，僅防止明文儲存）
function hashPass(pass) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, pass + 'KC_SALT_2026');
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── CORS Headers ───────────────────────────────────────────
function addCors(output) {
  // GAS doesn't support custom headers on ContentService directly
  // Use no-cors on client side
  return output;
}

// ════════════════════════════════════════════════════════════
//  GET handler
// ════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    initSheets();
    const action = e.parameter.action || '';
    const token  = e.parameter.token  || '';

    // ── 公開端點（不需登入）──
    if (action === 'getCampaigns') return getCampaignsPublic();
    if (action === 'ping')         return ok({ message: "King's Corner API OK", time: ts() });
    if (action === 'queryOrder')      return queryOrder(e.parameter.phone);
    if (action === 'validatePromoCode') return validatePromoCode(e.parameter.code, e.parameter.campaignId);
    if (action === 'validatePromoCodeMulti') {
      const ids = JSON.parse(e.parameter.campaignIds || '[]');
      return validatePromoCodeMulti(e.parameter.code, ids);
    }

    // ── 登入（公開）──
    if (action === 'login') {
      return login({ user: e.parameter.user, pass: e.parameter.pass });
    }

    // ── 訂單送出（公開）──
    if (action === 'submitOrder') {
      const body = {
        name:         e.parameter.name,
        phone:        e.parameter.phone,
        campaignId:   e.parameter.campaignId,
        campaignName: e.parameter.campaignName,
        items:        JSON.parse(e.parameter.items || '[]'),
        total:        parseFloat(e.parameter.total) || 0,
        discounted:   parseFloat(e.parameter.discounted) || 0,
        promoCode:    e.parameter.promoCode || '',
        promoDiscount:parseFloat(e.parameter.promoDiscount) || 0,
      };
      return submitOrder(body);
    }

    // ── 需要驗證的端點 ──
    const user = verifyToken(token);
    if (!user) return err('未授權，請重新登入');

    if (action === 'getAll')       return getAllData(user);
    if (action === 'getStock')     return getStock(e.parameter.campaignId);
    if (action === 'getOrders')    return getOrders(e.parameter.campaignId);
    if (action === 'getProducts')  return ok({ products: sheetToObjects(SH.PRODUCTS) });
    if (action === 'getSuppliers') return ok({ suppliers: sheetToObjects(SH.SUPPLIERS) });

    // ── 寫入端點（透過 GET 傳遞）──
    if (action === 'saveCampaign')   return saveCampaign({ campaign: JSON.parse(e.parameter.campaign || '{}') }, user);
    if (action === 'saveSupplier')   return saveSupplier({ supplier: JSON.parse(e.parameter.supplier || '{}') }, user);
    if (action === 'saveProduct')    return saveProduct({ product: JSON.parse(e.parameter.product || '{}') }, user);
    if (action === 'saveAds')        return saveAds({ supplierId: e.parameter.supplierId, slides: JSON.parse(e.parameter.slides || '[]') }, user);
    if (action === 'saveMarketing')  return saveMarketing({ settings: JSON.parse(e.parameter.settings || '{}') }, user);
    if (action === 'updatePaid')     return updatePaid({ orderId: e.parameter.orderId, paid: e.parameter.paid === 'true' }, user);
    if (action === 'saveAccount')    return saveAccount({ account: JSON.parse(e.parameter.account || '{}') }, user);
    if (action === 'deleteRecord')   return deleteRecord({ type: e.parameter.type, id: e.parameter.id }, user);
    if (action === 'importProducts')  return importProducts({ products: JSON.parse(e.parameter.products || '[]'), mode: e.parameter.mode }, user);
    if (action === 'savePromoCode')   return savePromoCode({ promoCode: JSON.parse(e.parameter.promoCode || '{}') }, user);
    if (action === 'deletePromoCode') return deletePromoCode(e.parameter.id, user);
    if (action === 'batchUpdatePaid') return batchUpdatePaid({ orderIds: JSON.parse(e.parameter.orderIds || '[]'), paid: e.parameter.paid === 'true' }, user);
    if (action === 'getPromoStats')   return getPromoStats(e.parameter.code, e.parameter.from, e.parameter.to, user);

    return err('未知的 action: ' + action);
  } catch(ex) {
    return err('GET 錯誤: ' + ex.message);
  }
}

// ════════════════════════════════════════════════════════════
//  POST handler
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    initSheets();
    const body   = JSON.parse(e.postData.contents || '{}');
    const action = body.action || '';
    const token  = body.token  || '';

    // ── 登入（不需 token）──
    if (action === 'login')       return addCorsHeaders(login(body));
    if (action === 'submitOrder') return addCorsHeaders(submitOrder(body));

    // ── 需要驗證的端點 ──
    const user = verifyToken(token);
    if (!user) return addCorsHeaders(err('未授權，請重新登入'));

    if (action === 'saveCampaign')   return addCorsHeaders(saveCampaign(body, user));
    if (action === 'saveSupplier')   return addCorsHeaders(saveSupplier(body, user));
    if (action === 'saveProduct')    return addCorsHeaders(saveProduct(body, user));
    if (action === 'saveAds')        return addCorsHeaders(saveAds(body, user));
    if (action === 'saveMarketing')  return addCorsHeaders(saveMarketing(body, user));
    if (action === 'updatePaid')     return addCorsHeaders(updatePaid(body, user));
    if (action === 'saveAccount')    return addCorsHeaders(saveAccount(body, user));
    if (action === 'deleteRecord')   return addCorsHeaders(deleteRecord(body, user));
    if (action === 'importProducts') return addCorsHeaders(importProducts(body, user));
    if (action === 'batchUpdatePaid') return addCorsHeaders(batchUpdatePaid(body, user));
    if (action === 'savePromoCode')  return addCorsHeaders(savePromoCode(body, user));
    if (action === 'deletePromoCode') return addCorsHeaders(deletePromoCode(body.id, user));

    return addCorsHeaders(err('未知的 action: ' + action));
  } catch(ex) {
    return addCorsHeaders(err('POST 錯誤: ' + ex.message));
  }
}

function addCorsHeaders(output) {
  return output;  // GAS handles CORS via deployment settings
}

function doOptions(e) {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ════════════════════════════════════════════════════════════
//  AUTH
// ════════════════════════════════════════════════════════════
// 簡易 token: base64(user:timestamp:hash)
function login(body) {
  const { user, pass } = body;
  const accounts = sheetToObjects(SH.ACCOUNTS);
  const acc = accounts.find(a => a.user === user && (a.active === 'true' || a.active === true || String(a.active).toUpperCase() === 'TRUE'));
  if (!acc) return err('帳號不存在或已停用');
  if (acc.passHash !== hashPass(pass)) return err('密碼錯誤');

  const token = generateToken(acc);
  return ok({
    token,
    user: { id: acc.id, name: acc.name, role: acc.role, user: acc.user }
  });
}

function generateToken(acc) {
  const payload = acc.id + ':' + acc.user + ':' + acc.role + ':' + new Date().getTime();
  return Utilities.base64Encode(payload + ':' + hashPass(payload));
}

function verifyToken(token) {
  if (!token) return null;
  try {
    const decoded = Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString();
    const parts = decoded.split(':');
    if (parts.length < 5) return null;
    const [id, user, role, timestamp, ...hashParts] = parts;
    const hash = hashParts.join(':');

    // Token 有效期 12 小時
    if (new Date().getTime() - parseInt(timestamp) > 12 * 3600 * 1000) return null;

    const payload = [id, user, role, timestamp].join(':');
    if (hashPass(payload) !== hash) return null;

    return { id, user, role };
  } catch(e) {
    return null;
  }
}

// ════════════════════════════════════════════════════════════
//  公開 API
// ════════════════════════════════════════════════════════════
function getCampaignsPublic() {
  const campaigns = sheetToObjects(SH.CAMPAIGNS);
  const active = campaigns
    .filter(c => c.status === 'active')
    .map(c => {
      try { c.products = JSON.parse(c.products || '[]'); } catch(e) { c.products = []; }
      // Normalize deadline: handle Date objects or strings from Sheets
      if (c.deadline) {
        if (c.deadline instanceof Date) {
          c.deadline = Utilities.formatDate(c.deadline, 'Asia/Taipei', 'yyyy-MM-dd');
        } else if (String(c.deadline).includes('T')) {
          c.deadline = String(c.deadline).split('T')[0];
        } else {
          c.deadline = String(c.deadline);
        }
      }
      return c;
    });

  const suppliers = sheetToObjects(SH.SUPPLIERS);
  const adsRaw = sheetToObjects(SH.ADS);
  const ads = {};
  adsRaw.forEach(a => {
    try { ads[a.supplierId] = JSON.parse(a.slides || '[]'); } catch(e) { ads[a.supplierId] = []; }
  });

  // Marketing settings
  const mktRows = sheetToObjects(SH.MARKETING);
  const mkt = {};
  mktRows.forEach(r => { mkt[r.key] = r.value; });

  // Add real order counts and stock info per campaign
  const allOrders = sheetToObjects(SH.ORDERS);

  // Also return closed campaigns for progress display
  const closedCamps = campaigns.filter(c => c.status === 'closed');
  closedCamps.forEach(c => {
    const campOrders = allOrders.filter(o => String(o.campaignId) === String(c.id));
    c.orderCount = campOrders.length;
  });

  active.forEach(c => {
    // Real order count
    const campOrders = allOrders.filter(o => String(o.campaignId) === String(c.id));
    c.orderCount = campOrders.length;

    // Sold qty per product
    const soldMap = {};
    campOrders.forEach(o => {
      let items = [];
      try { items = JSON.parse(o.items || '[]'); } catch(e) {}
      items.forEach(i => { soldMap[i.id] = (soldMap[i.id] || 0) + (parseInt(i.qty) || 0); });
    });

    // Add remaining stock to each product
    c.products = c.products.map(p => ({
      ...p,
      sold: soldMap[p.id] || 0,
      remaining: Math.max(0, (parseInt(p.stock) || 999) - (soldMap[p.id] || 0)),
    }));
  });

  return ok({ campaigns: active, closedCamps, suppliers, ads, marketing: mkt });
}

function submitOrder(body) {
  const { name, phone, campaignId, campaignName, items, total, discounted } = body;
  if (!name || !phone) return err('姓名與手機為必填');
  if (!/^09\d{8}$/.test(phone)) return err('手機號碼格式錯誤');
  if (!items || !items.length) return err('請選擇商品');

  // 防重複下單：同手機號 + 同團購
  const existing = sheetToObjects(SH.ORDERS);
  const dup = existing.find(o =>
    String(o.campaignId) === String(campaignId) && o.phone === phone
  );
  if (dup) return err('此手機號碼已在本次團購中下單，如需修改請聯絡管理員');

  // 取得廠商 ID
  const campaigns = sheetToObjects(SH.CAMPAIGNS);
  const camp = campaigns.find(c => String(c.id) === String(campaignId));
  const supplierId = camp ? camp.supplierId : '';

  // ── 庫存驗證 ──
  const campForStock = campaigns.find(c => String(c.id) === String(campaignId));
  if (campForStock) {
    let campProducts = [];
    try { campProducts = JSON.parse(campForStock.products || '[]'); } catch(e) {}

    // Calculate sold qty
    const existingOrders = sheetToObjects(SH.ORDERS).filter(o => String(o.campaignId) === String(campaignId));
    const soldMap = {};
    existingOrders.forEach(o => {
      let oi = [];
      try { oi = JSON.parse(o.items || '[]'); } catch(e) {}
      oi.forEach(i => { soldMap[i.id] = (soldMap[i.id] || 0) + (parseInt(i.qty) || 0); });
    });

    // Query product library once (outside loop for efficiency)
    const allProducts = sheetToObjects(SH.PRODUCTS);
    const failures = [];
    const inactiveItems = [];
    items.forEach(item => {
      const prod = campProducts.find(p => String(p.id) === String(item.id));
      if (!prod) return;
      // Check if product is active in product library (by id or by name as fallback)
      const libProd = allProducts.find(p => String(p.id) === String(item.id))
                   || allProducts.find(p => p.name === item.name);
      if (libProd) {
        const activeVal = libProd.active;
        // Normalize all forms: false, 'false', 'FALSE', FALSE
        const isInactive = activeVal === false ||
          String(activeVal).toUpperCase() === 'FALSE';
        if (isInactive) {
          inactiveItems.push(item.name);
          return;
        }
      }
      const stock = parseInt(prod.stock) || 999;
      const sold = soldMap[item.id] || 0;
      const remaining = Math.max(0, stock - sold);
      if ((parseInt(item.qty) || 0) > remaining) {
        failures.push({ name: item.name, ordered: item.qty, remaining });
      }
    });
    if (inactiveItems.length > 0) {
      return err('以下商品已停售，請移除後再下單：' + inactiveItems.join('、'));
    }

    if (failures.length > 0) {
      const msg = failures.map(f => `${f.name}：你訂 ${f.ordered} 份，剩餘 ${f.remaining} 份`).join('、');
      return err('庫存不足：' + msg);
    }
  }

  const orderId = 'KC' + new Date().getTime();
  const itemsStr = items.map(i => `${i.name}×${i.qty}`).join('、');
  const itemCount = items.reduce((s, i) => s + (parseInt(i.qty) || 0), 0);
  const promoCode    = body.promoCode    || '';
  const promoDiscount = parseFloat(body.promoDiscount) || 0;

  const shippingMethod = body.shippingMethod || '自取';
  const shippingFee    = parseFloat(body.shippingFee) || 0;
  const paymentMethod  = body.paymentMethod || '匯款';
  const address        = body.address || '';

  const headers = ['id','campaignId','campaignName','supplierId','name','phone','items','itemCount','total','discounted','promoCode','promoDiscount','shippingMethod','shippingFee','paymentMethod','address','paid','ts'];
  upsertRow(SH.ORDERS, headers, {
    id: orderId, campaignId, campaignName: campaignName || '',
    supplierId, name, phone,
    items: JSON.stringify(items),
    itemCount, total: total || 0, discounted: discounted || 0,
    promoCode: promoCode || '自然流量',
    promoDiscount: promoDiscount || 0,
    shippingMethod, shippingFee, paymentMethod, address,
    paid: 'false', ts: ts()
  }, 'id');

  // #12 LINE Notify 通知
  sendLineNotify(buildOrderNotifyMsg(name, phone, campaignName || camp?.name || '', items, discounted || 0, orderId));

  return ok({ orderId, message: '訂單已送出' });
}

// ════════════════════════════════════════════════════════════
//  Admin API
// ════════════════════════════════════════════════════════════
function getAllData(user) {
  const campaigns = sheetToObjects(SH.CAMPAIGNS).map(c => {
    try { c.products = JSON.parse(c.products || '[]'); } catch(e) { c.products = []; }
    if (c.deadline) {
      if (c.deadline instanceof Date) {
        c.deadline = Utilities.formatDate(c.deadline, 'Asia/Taipei', 'yyyy-MM-dd');
      } else if (String(c.deadline).includes('T')) {
        c.deadline = String(c.deadline).split('T')[0];
      } else {
        c.deadline = String(c.deadline);
      }
    }
    return c;
  });
  const orders = sheetToObjects(SH.ORDERS).map(o => {
    try { o.items = JSON.parse(o.items || '[]'); } catch(e) { o.items = []; }
    o.paid = o.paid === 'true';
    return o;
  });
  const adsRaw = sheetToObjects(SH.ADS);
  const ads = {};
  adsRaw.forEach(a => {
    try { ads[a.supplierId] = JSON.parse(a.slides || '[]'); } catch(e) { ads[a.supplierId] = []; }
  });
  const mktRows = sheetToObjects(SH.MARKETING);
  const marketing = {};
  mktRows.forEach(r => { marketing[r.key] = r.value; });

  // 只有 root/admin 才能看帳號列表
  let accounts = [];
  if (user.role === 'root') {
    accounts = sheetToObjects(SH.ACCOUNTS).map(a => ({ ...a, passHash: '***' }));
  }

  const promoCodes = sheetToObjects(SH.PROMO_CODES).map(p => ({
    ...p,
    status: (p.status === true || String(p.status).toUpperCase() === 'TRUE') ? 'true' : 'false'
  }));

  return ok({
    campaigns,
    orders,
    suppliers:  sheetToObjects(SH.SUPPLIERS),
    products:   sheetToObjects(SH.PRODUCTS),
    ads,
    marketing,
    accounts,
    promoCodes,
  });
}

function getOrders(campaignId) {
  let orders = sheetToObjects(SH.ORDERS).map(o => {
    try { o.items = JSON.parse(o.items || '[]'); } catch(e) { o.items = []; }
    o.paid = o.paid === 'true';
    return o;
  });
  if (campaignId) orders = orders.filter(o => String(o.campaignId) === String(campaignId));
  return ok({ orders });
}

function saveCampaign(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let c = body.campaign;
  if (typeof c === 'string') try { c = JSON.parse(c); } catch(e) {}
  // Ensure new fields have defaults
  if (!c.closedStatus) c.closedStatus = '';
  if (!c.pickupDate) c.pickupDate = '';
  if (!c || !c.name) return err('團購名稱為必填');
  if (!c.id) c.id = 'C' + new Date().getTime();
  if (!c.createdAt) c.createdAt = ts();
  c.products = JSON.stringify(c.products || []);
  const headers = ['id','name','supplierId','deadline','minPeople','discount','status','note','products','createdAt'];
  upsertRow(SH.CAMPAIGNS, headers, c, 'id');
  return ok({ id: c.id });
}

function saveSupplier(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let s = body.supplier;
  if (typeof s === 'string') try { s = JSON.parse(s); } catch(e) {}
  if (!s || !s.name) return err('廠商名稱為必填');
  if (!s.id) s.id = 'S' + new Date().getTime();
  if (!s.createdAt) s.createdAt = ts();
  const headers = ['id','name','cat','contact','note','createdAt'];
  upsertRow(SH.SUPPLIERS, headers, s, 'id');
  return ok({ id: s.id });
}

function saveProduct(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let p = body.product;
  if (typeof p === 'string') try { p = JSON.parse(p); } catch(e) {}
  if (!p || !p.name) return err('商品名稱為必填');
  if (!p.id) p.id = 'P' + new Date().getTime();
  const headers = ['id','supplierId','name','cat','spec','price','img','active'];
  upsertRow(SH.PRODUCTS, headers, p, 'id');
  return ok({ id: p.id });
}

function importProducts(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let { products, mode } = body;
  // products may be array (from POST JSON) or string (from GET)
  if (typeof products === 'string') {
    try { products = JSON.parse(products); } catch(e) { return err('products 格式錯誤'); }
  }
  if (!products || !products.length) return err('沒有可匯入的商品');

  // Auto-create missing suppliers
  const supHeaders = ['id','name','cat','contact','note','createdAt'];
  products.forEach(p => {
    if (p.supplierId && String(p.supplierId).startsWith('__new__')) {
      const supName = p.supplierName || p.supplierId.replace('__new__','');
      const existing = sheetToObjects(SH.SUPPLIERS);
      const found = existing.find(s => s.name === supName);
      if (found) {
        p.supplierId = String(found.id);
      } else {
        const newId = 'S' + new Date().getTime() + Math.random().toString(36).slice(2,5);
        upsertRow(SH.SUPPLIERS, supHeaders, {
          id: newId, name: supName, cat: '', contact: '', note: '', createdAt: ts()
        }, 'id');
        p.supplierId = newId;
      }
    }
  });

  const headers = ['id','supplierId','name','cat','spec','price','img','active'];
  const sheet = ensureSheet(SH.PRODUCTS, headers);

  if (mode === 'rebuild') {
    // Clear all data rows
    if (sheet.getLastRow() > 1) sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  const existing = sheetToObjects(SH.PRODUCTS);
  let added = 0, updated = 0, deactivated = 0;

  if (mode === 'replace_supplier') {
    // Deactivate products of affected suppliers not in this import
    const affectedSupIds = [...new Set(products.map(p => p.supplierId))];
    const csvNames = new Set(products.map(p => p.supplierId + '::' + p.name));
    existing.forEach(p => {
      if (affectedSupIds.includes(p.supplierId) && !csvNames.has(p.supplierId + '::' + p.name)) {
        p.active = 'false';
        upsertRow(SH.PRODUCTS, headers, p, 'id');
        deactivated++;
      }
    });
  }

  products.forEach(p => {
    if (!p.id) p.id = 'P' + new Date().getTime() + Math.random().toString(36).slice(2,6);
    const ex = existing.find(e => e.supplierId === p.supplierId && e.name === p.name);
    if (ex) { p.id = ex.id; updated++; } else { added++; }
    p.active = 'true';
    upsertRow(SH.PRODUCTS, headers, p, 'id');
  });

  return ok({ added, updated, deactivated });
}

function saveAds(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  const supplierId = body.supplierId;
  let slides = body.slides;
  if (typeof slides === 'string') try { slides = JSON.parse(slides); } catch(e) { slides = []; }
  const sheet = ensureSheet(SH.ADS, ['supplierId','slides']);
  const data = sheet.getDataRange().getValues();
  const sidCol = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][sidCol]) === String(supplierId)) {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(slides || []));
      return ok({});
    }
  }
  sheet.appendRow([supplierId, JSON.stringify(slides || [])]);
  return ok({});
}

function saveMarketing(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let settings = body.settings;
  if (typeof settings === 'string') try { settings = JSON.parse(settings); } catch(e) { settings = {}; }
  const sheet = ensureSheet(SH.MARKETING, ['key','value']);
  Object.entries(settings).forEach(([key, value]) => {
    const valStr = typeof value === 'object' ? JSON.stringify(value) : String(value);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(valStr);
        return;
      }
    }
    sheet.appendRow([key, valStr]);
  });
  return ok({});
}

function updatePaid(body, user) {
  if (!['root','admin','helper'].includes(user.role)) return err('權限不足');
  const { orderId, paid } = body;
  const sheet = ss().getSheetByName(SH.ORDERS);
  if (!sheet) return err('找不到訂單工作表');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const paidCol = headers.indexOf('paid');
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(orderId)) {
      sheet.getRange(i + 1, paidCol + 1).setValue(String(paid));
      return ok({});
    }
  }
  return err('找不到訂單: ' + orderId);
}

function saveAccount(body, user) {
  if (user.role !== 'root') return err('只有 ROOT 可以管理帳號');
  let acc = body.account;
  if (typeof acc === 'string') try { acc = JSON.parse(acc); } catch(e) {}
  if (!acc) return err('缺少帳號資料');
  const headers = ['id','user','passHash','name','role','active','createdAt'];

  // New account
  if (!acc.id) {
    const existing = sheetToObjects(SH.ACCOUNTS);
    if (existing.find(a => a.user === acc.user)) return err('帳號名稱已存在');
    if (!acc.pass) return err('新帳號需設定密碼');
    acc.id = 'A' + new Date().getTime();
    acc.passHash = hashPass(acc.pass);
    acc.createdAt = ts();
  } else {
    // Update — only change passHash if new pass provided
    const existing = sheetToObjects(SH.ACCOUNTS).find(a => a.id === acc.id);
    if (!existing) return err('找不到帳號');
    acc.passHash = acc.pass ? hashPass(acc.pass) : existing.passHash;
  }
  delete acc.pass;
  upsertRow(SH.ACCOUNTS, headers, acc, 'id');
  return ok({ id: acc.id });
}

function deleteRecord(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  const { type, id } = body;
  const map = { supplier: SH.SUPPLIERS, product: SH.PRODUCTS, campaign: SH.CAMPAIGNS };
  if (!map[type]) return err('無效的類型');
  deleteRow(map[type], 'id', id);
  return ok({});
}

// ════════════════════════════════════════════════════════════
//  #12 LINE Notify 訂單通知
// ════════════════════════════════════════════════════════════
// 設定方式：
//   1. 前往 https://notify-bot.line.me/ 申請 token
//   2. 把 token 填入下方 LINE_NOTIFY_TOKEN
//   3. 留空 = 關閉通知功能
const LINE_NOTIFY_TOKEN = ''; // ← 填入你的 LINE Notify token

function sendLineNotify(message) {
  if (!LINE_NOTIFY_TOKEN) return; // 未設定就跳過
  try {
    UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + LINE_NOTIFY_TOKEN },
      payload: { message: message },
      muteHttpExceptions: true,
    });
  } catch(e) {
    Logger.log('LINE Notify 失敗: ' + e.message);
  }
}

// 在 submitOrder 後呼叫，組合通知訊息
function buildOrderNotifyMsg(name, phone, campaignName, items, discounted, orderId) {
  const itemList = items.map(i => `  • ${i.name} ×${i.qty}`).join('\n');
  return `\n🛒 新訂單！\n` +
    `團購：${campaignName}\n` +
    `姓名：${name}\n` +
    `手機：${phone}\n` +
    `品項：\n${itemList}\n` +
    `金額：$${discounted}\n` +
    `訂單：${orderId}\n` +
    `時間：${ts()}`;
}

// ════════════════════════════════════════════════════════════
//  庫存查詢：計算各品項已售數量
// ════════════════════════════════════════════════════════════
function getStock(campaignId) {
  if (!campaignId) return err('缺少 campaignId');

  // Get campaign to know product stock limits
  const campaigns = sheetToObjects(SH.CAMPAIGNS);
  const camp = campaigns.find(c => String(c.id) === String(campaignId));
  if (!camp) return err('找不到團購');

  let products = [];
  try { products = JSON.parse(camp.products || '[]'); } catch(e) { products = []; }

  // Calculate sold qty per product from orders
  const orders = sheetToObjects(SH.ORDERS).filter(o => String(o.campaignId) === String(campaignId));
  const soldMap = {};
  orders.forEach(o => {
    let items = [];
    try { items = JSON.parse(o.items || '[]'); } catch(e) {}
    items.forEach(i => {
      soldMap[i.id] = (soldMap[i.id] || 0) + (parseInt(i.qty) || 0);
    });
  });

  // Build stock info per product
  const stockInfo = products.map(p => {
    const stock = parseInt(p.stock) || 999;
    const sold = soldMap[p.id] || 0;
    const remaining = Math.max(0, stock - sold);
    return {
      id: p.id,
      name: p.name,
      stock: stock,
      sold: sold,
      remaining: remaining,
      soldOut: remaining === 0,
    };
  });

  return ok({ stockInfo, campaignId });
}

// ════════════════════════════════════════════════════════════
//  #13 客人查詢自己的訂單
// ════════════════════════════════════════════════════════════
// GET ?action=queryOrder&phone=09XXXXXXXX
// 回傳該手機號碼的所有訂單（不需登入，但只能查自己的）
function queryOrder(phone) {
  if (!phone || !/^09\d{8}$/.test(phone)) return err('請輸入正確的手機號碼');

  const allOrders = sheetToObjects(SH.ORDERS).filter(o => o.phone === phone);
  if (!allOrders.length) return ok({ orders: [], message: '查無訂單' });

  const campaigns = sheetToObjects(SH.CAMPAIGNS);
  const result = allOrders.map(o => {
    const camp = campaigns.find(c => String(c.id) === String(o.campaignId)) || {};
    let items = [];
    try { items = JSON.parse(o.items || '[]'); } catch(e) {}
    return {
      orderId:      o.id,
      campaignName: camp.name || o.campaignName || '—',
      deadline:     camp.deadline || '',
      status:       camp.status || '',
      name:         o.name,
      items:        items,
      total:        o.total,
      discounted:   o.discounted,
      paid:         o.paid === 'true',
      ts:           o.ts,
      note:         camp.note || '', // 付款說明
    };
  }).sort((a, b) => String(b.orderId).localeCompare(String(a.orderId)));

  return ok({ orders: result, count: result.length });
}

// ════════════════════════════════════════════════════════════
//  #14 截止後自動關閉團購
// ════════════════════════════════════════════════════════════
// 設定方式：在 Apps Script 點「觸發條件」→ 新增觸發條件
//   函式：autoCloseCampaigns
//   事件來源：時間驅動
//   時間型觸發條件：每天（建議凌晨 1:00 - 2:00）
function autoCloseCampaigns() {
  const sheet = ss().getSheetByName(SH.CAMPAIGNS);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers = data[0];
  const statusCol = headers.indexOf('status');
  const deadlineCol = headers.indexOf('deadline');
  const nameCol = headers.indexOf('name');

  if (statusCol < 0 || deadlineCol < 0) {
    Logger.log('找不到 status 或 deadline 欄位');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let closedCount = 0;
  const closedNames = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][statusCol];
    const deadline = data[i][deadlineCol];
    const name = data[i][nameCol] || '';

    if (status !== 'active') continue; // 只處理進行中的

    const deadlineDate = new Date(deadline);
    deadlineDate.setHours(0, 0, 0, 0);

    // 截止日已過（嚴格：截止日的隔天才關）
    if (deadlineDate < today) {
      sheet.getRange(i + 1, statusCol + 1).setValue('closed');
      closedCount++;
      closedNames.push(name);
      Logger.log(`自動關閉：${name}（截止：${deadline}）`);
    }
  }

  if (closedCount > 0) {
    Logger.log(`✓ 自動關閉 ${closedCount} 個團購：${closedNames.join('、')}`);
    // 通知管理員
    if (LINE_NOTIFY_TOKEN) {
      sendLineNotify(`\n⏰ 自動關閉通知\n以下 ${closedCount} 個團購已到期自動關閉：\n${closedNames.map(n => '  • ' + n).join('\n')}`);
    }
  } else {
    Logger.log('✓ 自動關閉檢查完成，無需關閉的團購');
  }
}

// 手動觸發（可在 Apps Script 直接執行測試）
function testAutoClose() {
  autoCloseCampaigns();
}

// ════════════════════════════════════════════════════════════
//  優惠碼驗證（公開端點）
// ════════════════════════════════════════════════════════════
function validatePromoCode(code, campaignId) {
  if (!code || !code.trim()) return err('請輸入優惠碼');
  const codes = sheetToObjects(SH.PROMO_CODES);
  const promo = codes.find(c =>
    c.code.toUpperCase() === code.trim().toUpperCase() &&
    (c.status === 'true' || c.status === true || String(c.status).toUpperCase() === 'TRUE')
  );
  if (!promo) return err('優惠碼無效或已停用');

  // Parse per-campaign rules
  let campaignRules = [];
  try { campaignRules = JSON.parse(promo.campaignRules || '[]'); } catch(e) {}

  // Find rule for this specific campaign
  let rule = null;
  if (campaignId) {
    rule = campaignRules.find(r => String(r.campaignId) === String(campaignId));
  }

  // If no specific rule found, use default rule (first rule with campaignId='*') or 'none'
  if (!rule) {
    const defaultRule = campaignRules.find(r => r.campaignId === '*');
    rule = defaultRule || { campaignId: '*', discountType: 'none', discountValue: '{}' };
  }

  // Check whitelist: campaign must have this code in allowedCodes
  if (campaignId) {
    const campaigns = sheetToObjects(SH.CAMPAIGNS);
    const camp = campaigns.find(c => String(c.id) === String(campaignId));
    if (camp) {
      let allowedCodes = [];
      try { allowedCodes = JSON.parse(camp.allowedCodes || '[]'); } catch(e) {}
      if (allowedCodes.length > 0 && !allowedCodes.map(c=>c.toUpperCase()).includes(promo.code.toUpperCase())) {
        // Not in whitelist → use default behavior (not allowed)
        if (rule.discountType !== 'none') {
          return err('此優惠碼不適用本次團購');
        }
      }
    }
  }

  // Build discount config
  let discountValue = rule.discountValue || '{}';
  if (typeof discountValue === 'string' && !discountValue.startsWith('{')) {
    discountValue = JSON.stringify({ value: discountValue });
  }

  // Build human-readable label
  let cfg = {};
  try { cfg = JSON.parse(typeof discountValue === 'string' ? discountValue : JSON.stringify(discountValue)); } catch(e) {}
  let label = '純追蹤';
  if (rule.discountType === 'percent') {
    const n = parseFloat(cfg.value || 0);
    const fold = n / 10;
    label = (fold % 1 === 0 ? fold.toFixed(0) : String(n)) + '折';
  }
  else if (rule.discountType === 'amount_per_unit') label = `每份折 $${cfg.perUnit || cfg.value || '?'}`;
  else if (rule.discountType === 'amount_threshold') label = `滿$${cfg.threshold || '?'}折$${cfg.off || '?'}`;

  return ok({
    valid: true,
    code: promo.code,
    name: promo.name,
    discountType: rule.discountType,
    discountValue: discountValue,
    label,
    message: `✅ ${promo.name}・${label}`,
  });
}

// Validate promo code for multiple campaigns at once (for cart)
function validatePromoCodeMulti(code, campaignIds) {
  if (!code || !code.trim()) return err('請輸入優惠碼');
  const codes = sheetToObjects(SH.PROMO_CODES);
  const promo = codes.find(c =>
    c.code.toUpperCase() === code.trim().toUpperCase() &&
    (c.status === 'true' || c.status === true || String(c.status).toUpperCase() === 'TRUE')
  );
  if (!promo) return err('優惠碼無效或已停用');

  let campaignRules = [];
  try { campaignRules = JSON.parse(promo.campaignRules || '[]'); } catch(e) {}

  const campaigns = sheetToObjects(SH.CAMPAIGNS);
  const results = {};

  campaignIds.forEach(campaignId => {
    let rule = campaignRules.find(r => String(r.campaignId) === String(campaignId));
    if (!rule) {
      const defaultRule = campaignRules.find(r => r.campaignId === '*');
      rule = defaultRule || { campaignId: '*', discountType: 'none', discountValue: '{}' };
    }

    // Check whitelist
    const camp = campaigns.find(c => String(c.id) === String(campaignId));
    if (camp) {
      let allowedCodes = [];
      try { allowedCodes = JSON.parse(camp.allowedCodes || '[]'); } catch(e) {}
      if (allowedCodes.length > 0 && !allowedCodes.map(c=>c.toUpperCase()).includes(promo.code.toUpperCase())) {
        results[campaignId] = { discountType: 'none', discountValue: '{}', label: '此團購不適用' };
        return;
      }
    }

    let cfg = {};
    try { cfg = JSON.parse(typeof rule.discountValue === 'string' ? rule.discountValue : JSON.stringify(rule.discountValue || '{}')); } catch(e) {}
    let label = '純追蹤';
    if (rule.discountType === 'percent') {
    const n = parseFloat(cfg.value || 0);
    const fold = n / 10;
    label = (fold % 1 === 0 ? fold.toFixed(0) : String(n)) + '折';
  }
    else if (rule.discountType === 'amount_per_unit') label = `每份折 $${cfg.perUnit || cfg.value || '?'}`;
    else if (rule.discountType === 'amount_threshold') label = `滿$${cfg.threshold || '?'}折$${cfg.off || '?'}`;

    results[campaignId] = { discountType: rule.discountType, discountValue: rule.discountValue, label };
  });

  return ok({ valid: true, code: promo.code, name: promo.name, results });
}

// ════════════════════════════════════════════════════════════
//  優惠碼 CRUD
// ════════════════════════════════════════════════════════════
function savePromoCode(body, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let p = body.promoCode;
  if (typeof p === 'string') try { p = JSON.parse(p); } catch(e) {}
  if (!p || !p.code || !p.name) return err('優惠碼和名稱為必填');

  const existing = sheetToObjects(SH.PROMO_CODES);

  // Duplicate check: same code but different id = reject
  const dup = existing.find(e =>
    e.code.toUpperCase() === p.code.toUpperCase() && String(e.id) !== String(p.id)
  );
  if (dup) return err(`優惠碼「${p.code}」已存在（${dup.name}），請使用不同的代碼`);

  if (!p.id) p.id = 'PC' + new Date().getTime();
  if (!p.createdAt) p.createdAt = ts();
  const headers = ['id','code','name','contact','discountType','discountValue','campaignRules','status','note','createdAt'];
  upsertRow(SH.PROMO_CODES, headers, p, 'id');
  return ok({ id: p.id });
}

function deletePromoCode(id, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  deleteRow(SH.PROMO_CODES, 'id', id);
  return ok({});
}

// ════════════════════════════════════════════════════════════
//  優惠碼使用統計
// ════════════════════════════════════════════════════════════
function getPromoStats(code, from, to, user) {
  if (!['root','admin'].includes(user.role)) return err('權限不足');
  let orders = sheetToObjects(SH.ORDERS);

  // Filter by date range
  if (from) orders = orders.filter(o => o.ts && o.ts >= from);
  if (to)   orders = orders.filter(o => o.ts && o.ts <= to + ' 23:59:59');

  // Filter by specific code or get all
  if (code) orders = orders.filter(o => o.promoCode === code);

  // Group by promoCode + campaignName
  const statsMap = {};
  orders.forEach(o => {
    const c = o.promoCode || '自然流量';
    const camp = o.campaignName || o.campaignId || '未知';
    const key = c + '::' + camp;
    if (!statsMap[key]) statsMap[key] = { promoCode: c, campaignName: camp, count: 0, totalDiscounted: 0 };
    statsMap[key].count++;
    statsMap[key].totalDiscounted += parseFloat(o.discounted) || 0;
  });

  const stats = Object.values(statsMap).sort((a,b) => {
    if (a.promoCode !== b.promoCode) return a.promoCode.localeCompare(b.promoCode);
    return b.count - a.count;
  });

  return ok({ stats, total: orders.length });
}

// ════════════════════════════════════════════════════════════
//  批次更新付款狀態
// ════════════════════════════════════════════════════════════
function batchUpdatePaid(body, user) {
  if (!['root','admin','helper'].includes(user.role)) return err('權限不足');
  const { orderIds, paid } = body;
  if (!orderIds || !orderIds.length) return err('請選擇訂單');

  const sheet = ss().getSheetByName(SH.ORDERS);
  if (!sheet) return err('找不到訂單工作表');

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const paidCol = headers.indexOf('paid');
  if (idCol < 0 || paidCol < 0) return err('欄位結構錯誤');

  let updated = 0;
  const paidStr = String(paid);
  const idSet = new Set(orderIds.map(String));

  for (let i = 1; i < data.length; i++) {
    if (idSet.has(String(data[i][idCol]))) {
      sheet.getRange(i + 1, paidCol + 1).setValue(paidStr);
      updated++;
    }
  }

  return ok({ updated });
}

// ════════════════════════════════════════════════════════════
//  手動初始化（第一次部署後在 Apps Script 執行一次）
// ════════════════════════════════════════════════════════════
function setup() {
  initSheets();
  Logger.log('✓ King\'s Corner Sheets 初始化完成');
  Logger.log('預設帳號: root / root1234');
  Logger.log('請登入後立即修改密碼！');
  Logger.log('');
  Logger.log('⚠️  還需要手動設定：');
  Logger.log('1. LINE_NOTIFY_TOKEN：填入 LINE Notify token 以接收新訂單通知');
  Logger.log('2. 自動關閉觸發條件：觸發條件 → 新增 → autoCloseCampaigns → 每天');
}
