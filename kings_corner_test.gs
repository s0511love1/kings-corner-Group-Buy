// ════════════════════════════════════════════════════════════
//  King's Corner 自動化測試 + 診斷
//  用法：Apps Script 編輯器 → 檔案旁「＋」→ 指令碼 → 命名 kings_corner_test → 整份貼上 → 儲存
//        回到 Google Sheets 重新整理頁面 → 上方出現「🧪 KC 測試」選單
//  特色：不需要 Console、不需要 Node。結果寫在「測試結果」「診斷結果」工作表，整張複製貼回給我即可。
//  安全：測試會暫時寫入「團購 / 訂單 / 優惠碼」三張表（id 一律以 TEST_ 開頭、手機 0900000001~09），
//        結束時一定會自動刪除（包含失敗時）。若中途被中斷，執行「清除測試資料」即可。
//  ⚠ 若你已設定 LINE_NOTIFY_TOKEN，測試下單會發出通知；測試前可先暫時清空 token。
//  正式上線後可整個刪掉這個檔案，不影響主程式。
// ════════════════════════════════════════════════════════════

const KC_TEST_PHONES = ['0900000001','0900000002','0900000003','0900000004','0900000005',
                        '0900000006','0900000007','0900000008','0900000009'];
const KC_TEST_SHEET = '測試結果';
const KC_DIAG_SHEET = '診斷結果';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧪 KC 測試')
    .addItem('▶ 執行全部測試', 'kcRunAllTests')
    .addItem('🔍 診斷目前資料設定（唯讀）', 'kcDiagnose')
    .addSeparator()
    .addItem('🧹 清除測試資料', 'kcCleanup')
    .addItem('📞 修復電話號碼開頭 0', 'fixPhoneNumbers')
    .addToUi();
}

// ── 小工具 ───────────────────────────────────────────────
function _kcCall(fn, arg) {
  return JSON.parse(fn(arg).getContent());
}

function _kcTaipeiDay(offsetDays) {
  return Utilities.formatDate(new Date(Date.now() + offsetDays * 86400000), 'Asia/Taipei', 'yyyy-MM-dd');
}

function _kcDeleteRows(sheetName, pred) {
  const sh = ss().getSheetByName(sheetName);
  if (!sh) return 0;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;
  const h = data[0];
  let n = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const o = {};
    h.forEach((k, j) => { o[k] = data[i][j]; });
    if (pred(o)) { sh.deleteRow(i + 1); n++; }
  }
  return n;
}

function kcCleanup() {
  const isTest = v => String(v || '').indexOf('TEST_') === 0;
  const a = _kcDeleteRows(SH.CAMPAIGNS, o => isTest(o.id));
  const b = _kcDeleteRows(SH.ORDERS, o => isTest(o.campaignId));
  const c = _kcDeleteRows(SH.PROMO_CODES, o => isTest(o.id));
  const msg = '已清除測試資料：團購 ' + a + ' 筆、訂單 ' + b + ' 筆、優惠碼 ' + c + ' 筆';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ── 測試資料建立 ─────────────────────────────────────────
function _kcSetupFixtures() {
  const today = _kcTaipeiDay(0), tomorrow = _kcTaipeiDay(1), yesterday = _kcTaipeiDay(-1);
  const products = JSON.stringify([{ id: 'TEST_P1', name: 'ZZ測試商品', price: 100, stock: 50 }]);
  const headers = ['id','name','supplierId','deadline','minPeople','discountRule','status','closedStatus','pickupDate','note','products','allowedCodes','createdAt'];
  const rule = JSON.stringify({ type: 'amount_per_unit', perUnit: '3' });
  const base = { supplierId: 'TEST_SUP', minPeople: 1, discountRule: rule, status: 'active',
                 closedStatus: '', pickupDate: '', note: '', products: products, createdAt: ts() };
  const mk = (id, over) => upsertRow(SH.CAMPAIGNS, headers, Object.assign({}, base, { id: id, name: id }, over), 'id');

  mk('TEST_ACTIVE',       { deadline: tomorrow,  allowedCodes: '["TESTPROMO"]' });
  mk('TEST_NOWL',         { deadline: tomorrow,  allowedCodes: '[]' });
  mk('TEST_EXP_TEXT',     { deadline: yesterday, allowedCodes: '[]' });
  mk('TEST_EXP_DATE',     { deadline: new Date(yesterday + 'T00:00:00+08:00'), allowedCodes: '[]' });
  mk('TEST_TODAY_DATE',   { deadline: new Date(today + 'T00:00:00+08:00'),     allowedCodes: '[]' });
  mk('TEST_DRAFT',        { deadline: tomorrow,  status: 'draft', allowedCodes: '[]' });

  const dv = JSON.stringify({ threshold: '100', off: '10' });
  const rules = JSON.stringify([
    { campaignId: 'TEST_ACTIVE', discountType: 'amount_threshold', discountValue: dv },
    { campaignId: 'TEST_NOWL',   discountType: 'amount_threshold', discountValue: dv }
  ]);
  upsertRow(SH.PROMO_CODES,
    ['id','code','name','contact','discountType','discountValue','campaignRules','status','note','createdAt'],
    { id: 'TEST_PC1', code: 'TESTPROMO', name: '測試碼', contact: '', discountType: 'none',
      discountValue: '', campaignRules: rules, status: 'true', note: '', createdAt: ts() }, 'id');
}

// ── 主流程 ───────────────────────────────────────────────
function kcRunAllTests() {
  const rows = [];
  const rec = (id, name, expected, actual, pass, note) =>
    rows.push([id, name, pass === null ? 'SKIP' : (pass ? 'PASS' : 'FAIL'), String(expected), String(actual), note || '']);
  const has = (s, kw) => String(s || '').indexOf(kw) >= 0;
  const items = qty => [{ id: 'TEST_P1', name: 'ZZ測試商品', qty: qty, price: 1 }];
  const order = (cid, phone, extra) => Object.assign(
    { action: 'submitOrder', name: '測試', phone: phone, campaignId: cid, campaignName: cid,
      items: items(2), total: 1, discounted: 1, shippingMethod: '自取', shippingFee: 0 }, extra || {});
  const find = id => sheetToObjects(SH.ORDERS).find(o => String(o.id) === String(id));

  try {
    kcCleanup();            // 先清一次，避免上次殘留
    _kcSetupFixtures();

    // ── T1 日期正規化（不依賴試算表） ──
    const dcases = [
      ['文字 2026-06-25',                '2026-06-25',                          '2026-06-25'],
      ['ISO 字串（=台北 6/25）',          '2026-06-24T16:00:00.000Z',            '2026-06-25'],
      ['Date 物件（=台北 6/25）',         new Date('2026-06-24T16:00:00.000Z'),  '2026-06-25'],
      ['斜線 2026/6/5',                  '2026/6/5',                            '2026-06-05'],
      ['空白',                           '',                                    '']
    ];
    dcases.forEach((c, i) => { const got = toTaipeiDateStr(c[1]);
      rec('T1-' + (i + 1), '日期正規化：' + c[0], c[2], got, got === c[2]); });

    // ── T2 截止日 / 狀態驗證（後端） ──
    let r = _kcCall(submitOrder, order('TEST_EXP_TEXT', '0900000005'));
    rec('T2-1', '已截止（文字日期）被擋', 'error 含「截止」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '截止'));

    r = _kcCall(submitOrder, order('TEST_EXP_DATE', '0900000005'));
    const rawDate = (() => { const o = sheetToObjects(SH.CAMPAIGNS).find(x => x.id === 'TEST_EXP_DATE');
      return o ? (o.deadline instanceof Date ? 'Date→' + toTaipeiDateStr(o.deadline) : 'text→' + o.deadline) : '?'; })();
    rec('T2-2', '已截止（日期格式儲存格）被擋', 'error 含「截止」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '截止'),
        '儲存格讀回：' + rawDate + '；試算表時區=' + ss().getSpreadsheetTimeZone() + '，腳本時區=' + Session.getScriptTimeZone());

    r = _kcCall(submitOrder, order('TEST_TODAY_DATE', '0900000006'));
    rec('T2-3', '截止日=今天（日期格式）當天仍可下單', 'ok', r.status + ' ' + (r.message || ''), r.status === 'ok',
        '若 FAIL 且 T2-2 的時區不是 Asia/Taipei，代表時區造成日期偏移');

    r = _kcCall(submitOrder, order('TEST_DRAFT', '0900000008'));
    rec('T2-4', '草稿團購不可下單', 'error 含「非進行中」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '非進行中'));

    // ── T3 金額由後端重算（前端傳假價格 1 元） ──
    r = _kcCall(submitOrder, order('TEST_NOWL', '0900000001', { promoCode: 'TESTPROMO' }));
    let o = r.status === 'ok' ? find(r.orderId) : null;
    let it = o ? JSON.parse(o.items || '[]')[0] : {};
    rec('T3-1', '偽造單價無效：items.price 為真實單價 100', '100', it.price, Number(it.price) === 100);
    rec('T3-2', '原價小計 total = 100×2', '200', o ? o.total : '無訂單', o && Number(o.total) === 200);
    rec('T3-3', '團購每份折 $3 → 折後 194（偽造 discounted=1 無效）', '194', o ? o.discounted : '無訂單', o && Number(o.discounted) === 194);

    // ── T4 優惠碼白名單（空白名單=不允許） ──
    rec('T4-1', '白名單空白的團購：優惠碼不套用，記為自然流量', '自然流量 / 折扣0', o ? (o.promoCode + ' / ' + o.promoDiscount) : '無訂單',
        o && o.promoCode === '自然流量' && Number(o.promoDiscount) === 0);

    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000002', { promoCode: 'TESTPROMO' }));
    o = r.status === 'ok' ? find(r.orderId) : null;
    rec('T4-2', '白名單有勾選：優惠碼套用（194 → 每滿100折10 → 折10）', 'TESTPROMO / 折10 / 實付184',
        o ? (o.promoCode + ' / 折' + o.promoDiscount + ' / 實付' + o.discounted) : (r.message || '無訂單'),
        o && o.promoCode === 'TESTPROMO' && Number(o.promoDiscount) === 10 && Number(o.discounted) === 184);
    const t42Id = o ? o.id : null;

    // 優惠碼驗證 API 本身
    const v1 = _kcCall(function () { return validatePromoCode('TESTPROMO', 'TEST_NOWL'); });
    rec('T4-3', 'validatePromoCode：未勾選的團購回傳錯誤', 'error', v1.status, v1.status === 'error');
    const v2 = _kcCall(function () { return validatePromoCode('TESTPROMO', 'TEST_ACTIVE'); });
    rec('T4-4', 'validatePromoCode：已勾選的團購通過', 'ok', v2.status, v2.status === 'ok');

    // ── T5 運費計入 discounted（行為確認） ──
    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000003', { shippingFee: 65, shippingMethod: '7-11' }));
    o = r.status === 'ok' ? find(r.orderId) : null;
    rec('T5', '運費 65 不參與折扣：discounted 仍為 194，shippingFee=65', '194 / 65', o ? (o.discounted + ' / ' + o.shippingFee) : (r.message || '無訂單'), o && Number(o.discounted) === 194 && Number(o.shippingFee) === 65,
        '運費獨立欄位，不計入折扣金額');

    // ── T6 防重複 / 鎖 ──
    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000002'));
    rec('T6-1', '同手機同團購重複下單被擋', 'error 含「已在本次團購中下單」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '已在本次團購中下單'));
    const lk = LockService.getScriptLock();
    const got = lk.tryLock(3000);
    if (got) lk.releaseLock();
    rec('T6-2', '鎖在下單結束後已釋放（失敗/成功路徑皆然）', 'true', got, got === true);
    rec('T6-3', '真正的同時連點（平行請求）', '—', '無法在單一腳本內重現', null,
        'Apps Script 無法在同一次執行內平行呼叫；T6-1/T6-2 只證明「循序重複被擋、鎖不卡死」。要實測請用兩支手機同時按送出');

    // ── T7 庫存與商品驗證 ──
    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000004', { items: items(999) }));
    rec('T7-1', '超過庫存被擋（庫存50 訂999）', 'error 含「庫存不足」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '庫存不足'));
    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000004', { items: [{ id: 'NO_SUCH', name: '不存在', qty: 1, price: 1 }] }));
    rec('T7-2', '不存在的商品被擋', 'error 含「資料有誤」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '資料有誤'));

    // ── T8 電話開頭 0 ──
    const osh = ss().getSheetByName(SH.ORDERS);
    const od = osh.getDataRange().getValues();
    const pcol = od[0].indexOf('phone'), icol = od[0].indexOf('id');
    let rawPhone = '(找不到)';
    for (let i = 1; i < od.length; i++) if (String(od[i][icol]) === String(t42Id)) rawPhone = od[i][pcol];
    rec('T8-1', 'Sheets 儲存格內電話為文字且保留開頭 0', "字串 '0900000002'", typeof rawPhone + ' ' + rawPhone,
        typeof rawPhone === 'string' && rawPhone === '0900000002');
    const q = _kcCall(queryOrder, '0900000002');
    rec('T8-2', '用 09 開頭查詢訂單查得到', 'ok 且筆數≥1', q.status + ' ' + (q.count || 0), q.status === 'ok' && (q.count || 0) >= 1);

    // 舊資料（電話被吃掉 0）仍能被查詢與防重複辨識
    upsertRow(SH.ORDERS,
      ['id','campaignId','campaignName','supplierId','name','phone','items','itemCount','total','discounted','promoCode','promoDiscount','shippingMethod','shippingFee','paymentMethod','address','paid','ts'],
      { id: 'TEST_LEGACY1', campaignId: 'TEST_ACTIVE', campaignName: 'TEST_ACTIVE', supplierId: 'TEST_SUP', name: '舊資料',
        phone: 900000007, items: '[]', itemCount: 0, total: 0, discounted: 0, promoCode: '自然流量', promoDiscount: 0,
        shippingMethod: '自取', shippingFee: 0, paymentMethod: '匯款', address: '', paid: 'false', ts: ts() }, 'id');
    const q2 = _kcCall(queryOrder, '0900000007');
    rec('T8-3', '舊資料（9 碼）也能用 09 開頭查到', 'ok 且筆數≥1', q2.status + ' ' + (q2.count || 0), q2.status === 'ok' && (q2.count || 0) >= 1);
    r = _kcCall(submitOrder, order('TEST_ACTIVE', '0900000007'));
    rec('T8-4', '舊資料手機再下單仍被防重複擋住', 'error 含「已在本次團購中下單」', r.status + ' ' + (r.message || ''), r.status === 'error' && has(r.message, '已在本次團購中下單'));

  } catch (e) {
    rec('ERR', '測試執行中發生例外', '無例外', e.message + (e.stack ? ' @ ' + String(e.stack).split('\n')[0] : ''), false);
  } finally {
    const msg = kcCleanup();
    rec('CLEAN', '測試資料已清除', 'TEST_ 資料全數刪除', msg, true);
  }

  _kcWriteSheet(KC_TEST_SHEET, ['編號', '項目', '結果', '預期', '實際', '備註'], rows);
  const pass = rows.filter(x => x[2] === 'PASS').length, fail = rows.filter(x => x[2] === 'FAIL').length,
        skip = rows.filter(x => x[2] === 'SKIP').length;
  const summary = '完成：PASS ' + pass + ' / FAIL ' + fail + ' / SKIP ' + skip + '\n結果在「' + KC_TEST_SHEET + '」工作表';
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) {}
  return summary;
}

// ── 診斷（唯讀，不寫任何業務資料） ──────────────────────────
function kcDiagnose() {
  const rows = [];
  const add = (cat, item, value, verdict) => rows.push([cat, item, String(value), verdict || '']);
  const sstz = ss().getSpreadsheetTimeZone(), sctz = Session.getScriptTimeZone();
  add('時區', '試算表時區', sstz, sstz === 'Asia/Taipei' ? 'OK' : '⚠ 非台北時區，日期格式儲存格可能偏移一天');
  add('時區', '腳本時區', sctz, sctz === 'Asia/Taipei' ? 'OK' : '⚠ 建議在專案設定改為 Asia/Taipei');

  // 團購
  const csh = ss().getSheetByName(SH.CAMPAIGNS);
  if (csh && csh.getLastRow() > 1) {
    const d = csh.getDataRange().getValues(), h = d[0];
    const ci = n => h.indexOf(n);
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][ci('id')] || ''); if (!id || id.indexOf('TEST_') === 0) continue;
      const st = String(d[i][ci('status')] || '');
      if (st !== 'active') continue;
      const raw = d[i][ci('deadline')];
      const kind = raw instanceof Date ? 'Date物件' : '文字';
      const norm = toTaipeiDateStr(raw);
      const exp = isCampaignExpired(raw);
      let wl = []; try { wl = JSON.parse(d[i][ci('allowedCodes')] || '[]'); } catch (e) {}
      const dr = d[i][ci('discountRule')];
      const name = d[i][ci('name')];
      add('進行中團購', name + '｜截止日', kind + ' → ' + norm, !norm ? '⚠ 日期無法辨識（後端不會擋單）' : (exp ? '已過期（status 仍是 active，等自動關閉）' : 'OK'));
      add('進行中團購', name + '｜開放優惠碼', wl.length ? wl.join(', ') : '(空白)', wl.length ? 'OK' : '⚠ 空白＝所有優惠碼不可用（本次行為變更）');
      add('進行中團購', name + '｜團購折扣 discountRule', dr || '(無)', dr ? 'OK' : '無折扣（舊團購的 discount 已不再生效）');
    }
  } else add('進行中團購', '—', '團購工作表沒有資料', '');

  // 訂單電話
  const osh = ss().getSheetByName(SH.ORDERS);
  if (osh && osh.getLastRow() > 1) {
    const d = osh.getDataRange().getValues(), pc = d[0].indexOf('phone');
    let bad = 0, total = 0;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]).indexOf('TEST_') === 0) continue;
      total++; if (/^9\d{8}$/.test(String(d[i][pc]))) bad++;
    }
    const fmt = osh.getRange(2, pc + 1).getNumberFormat();
    add('訂單電話', '總筆數 / 缺 0 的筆數', total + ' / ' + bad, bad ? '⚠ 請執行「修復電話號碼開頭 0」' : 'OK');
    add('訂單電話', '電話欄第一筆儲存格格式', fmt, fmt === '@' ? 'OK（純文字）' : '⚠ 非純文字；修復函式或下一筆新單會設定');
  }

  // 優惠碼
  const psh = ss().getSheetByName(SH.PROMO_CODES);
  if (psh && psh.getLastRow() > 1) {
    const d = psh.getDataRange().getValues(), h = d[0];
    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][h.indexOf('id')] || ''); if (!id || id.indexOf('TEST_') === 0) continue;
      let rules = []; try { rules = JSON.parse(d[i][h.indexOf('campaignRules')] || '[]'); } catch (e) {}
      const seen = {}; let dup = false;
      rules.forEach(r => { if (seen[r.campaignId]) dup = true; seen[r.campaignId] = true; });
      add('優惠碼', d[i][h.indexOf('code')] + '｜規則數', rules.length, dup ? '⚠ 同一團購有重複規則（只有第一條會生效）' : 'OK');
    }
  }

  _kcWriteSheet(KC_DIAG_SHEET, ['類別', '項目', '值', '判定'], rows);
  const warn = rows.filter(x => String(x[3]).indexOf('⚠') >= 0).length;
  const msg = '診斷完成，共 ' + rows.length + ' 項，警告 ' + warn + ' 項\n結果在「' + KC_DIAG_SHEET + '」工作表';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  return msg;
}

// ── 寫結果工作表 ─────────────────────────────────────────
function _kcWriteSheet(name, header, rows) {
  const book = ss();
  let sh = book.getSheetByName(name);
  if (!sh) sh = book.insertSheet(name);
  sh.clear();
  const all = [header].concat(rows.length ? rows : [header.map(() => '')]);
  const range = sh.getRange(1, 1, all.length, header.length);
  range.setNumberFormat('@');           // 全部當文字，避免 '09...' 或 '-' 被轉型
  range.setValues(all);
  range.setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#12100E').setFontColor('#C9A84C');
  sh.setFrozenRows(1);
  const widths = header.length === 6 ? [70, 300, 70, 220, 300, 320] : [100, 330, 360, 330];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const rc = header.indexOf('結果');
  if (rc >= 0) {
    rows.forEach((r, i) => {
      const bg = r[rc] === 'PASS' ? '#d9ead3' : r[rc] === 'FAIL' ? '#f4cccc' : '#e0e0e0';
      sh.getRange(i + 2, rc + 1).setBackground(bg);
    });
  } else {
    rows.forEach((r, i) => { if (String(r[3]).indexOf('⚠') >= 0) sh.getRange(i + 2, 4).setBackground('#fff2cc'); });
  }
  sh.activate();
}
