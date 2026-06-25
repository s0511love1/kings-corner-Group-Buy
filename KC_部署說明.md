# King's Corner 團購平台 — 部署說明 v2.2

> 最後更新：2026-06-25

---

## 系統架構

```
客人 / 團購主
    ↓ 瀏覽器（GitHub Pages）
Google Apps Script（後端 API）
    ↓ 讀寫
Google Sheets（資料庫，8 個工作表）
    ↑
管理員 / 員工（admin.html）
```

費用：全部 $0

---

## 檔案清單

| 檔案 | 用途 |
|---|---|
| `kings_corner_index.html` | 客戶端訂購頁（含購物車） |
| `kings_corner_admin.html` | 後台管理系統 |
| `kings_corner_gas.js` | Google Apps Script 程式碼 |

---

## 部署步驟

### Phase 1：建立 Google Sheets

1. 新增空白試算表，命名「**King's Corner 後台**」
2. 複製網址中的 Sheets ID：
   ```
   https://docs.google.com/spreadsheets/d/【ID在這裡】/edit
   ```

### Phase 2：部署 Google Apps Script

1. 試算表 → 擴充功能 → Apps Script
2. 全選刪除，貼上 `kings_corner_gas.js` 全部內容
3. 第 8 行填入你的 Sheets ID：
   ```javascript
   const SHEET_ID = '你的ID';
   ```
4. Ctrl+S 儲存
5. 函式選單選 **`setup`** → ▶ 執行（第一次需要授權）
6. 執行完成後 Sheets 應出現 **8 個工作表**
7. 部署 → 新增部署作業 → 類型：**網頁應用程式**
   - 執行身分：**我（你的帳號）**
   - 存取權限：**所有人**
8. 複製 `exec` 結尾的部署網址

### Phase 3：上傳 GitHub Pages

1. 新增 Repository（Public，勾選 README）
2. 上傳三個 HTML 檔案
3. Settings → Pages → Branch: main / root → Save
4. 約 1 分鐘後網站上線

### Phase 4：設定 Admin

1. 開啟 admin 網址，首次會要求填入 GAS 部署網址
2. 登入：`root` / `root1234`
3. **立即修改 root 密碼**

---

## 更新程式碼

**GAS 有改動：**
全選貼上 → 填回 SHEET_ID → 儲存 → 部署 → 鉛筆 → **建立新版本** → 部署
> 若有新增欄位，部署後需重新執行一次 `setup()`，會自動補上缺少的欄位，不影響現有資料。

**HTML 有改動：**
GitHub → Add file → Upload files → 覆蓋 → Commit（約 1 分鐘生效）

---

## Sheets 欄位說明（重要）

### 團購工作表欄位順序
```
id | name | supplierId | deadline | minPeople | discount | status | closedStatus | pickupDate | note | products | allowedCodes | createdAt
```

### 優惠碼工作表欄位順序
```
id | code | name | contact | discountType | discountValue | campaignRules | status | note | createdAt
```

### 訂單工作表欄位順序
```
id | campaignId | campaignName | supplierId | name | phone | items | itemCount | total | discounted | promoCode | promoDiscount | shippingMethod | shippingFee | paymentMethod | address | paid | ts
```

> **重要：** 欄位順序必須與上方一致。若用 `ensureSheet` 補欄後欄位跑到最右側，需手動在 Sheets 調整欄位順序，再重新新增資料。

---

## Trouble Shooting

### T1：Admin「無法連線」
GAS URL 末尾必須是 `/exec`，部署設定「存取權限」必須是「所有人」。

### T2：帳號不存在或已停用
確認 Sheets「帳號」工作表有資料，若為空重新執行 `setup()`。

### T3：Index 看到假資料
必須上傳到 GitHub Pages，`file://` 本機無法連線 GAS。

### T4：新增/編輯儲存失敗（400 Bad Request）
確認使用最新版 admin.html（已改為 POST JSON body）。

### T5：CSV 匯入失敗
確認使用最新版 admin.html（分批 20 筆 POST）。

### T6：優惠碼 `campaignRules` 欄顯示時間
欄位順序錯位，手動在 Sheets 調整後重新新增資料。正確順序見上方。

### T7：優惠碼狀態全部顯示「停用」
確認使用最新版 GAS（TRUE 正規化已修復）。

### T8：GAS 重新部署後改動沒有生效
部署 → 管理部署作業 → 鉛筆 → **建立新版本**（不是選舊版本）。

### T9：首頁預設顯示某個廠商而非「全部」
確認使用最新版 index.html。

### T10：購物車返回首頁後畫面只剩 Banner
確認使用最新版 index.html。

### T11：載入後不顯示任何團購
確認使用最新版 index.html，supplierId 型別比對已修復。

### T12：停用商品仍可送出訂單
確認 GAS 已重新部署（submitOrder 已加入 active 欄位驗證，支援 name fallback）。

### T13：已截止團購沒有顯示在首頁進度區
- 確認 GAS 已重新部署（getCampaignsPublic 已回傳 closedCamps）
- 確認執行過 `setup()`（新增 closedStatus / pickupDate 欄位）
- 確認團購狀態為 `closed`（不是 `done`，`done` 不顯示）
- 截止日過了但狀態還是 `active` → 執行 `createAutoCloseTrigger()` 設定自動關閉

### T14：autoFillProductImages 執行後商品沒有圖片
- 確認「圖片清單」工作表 B欄名稱與商品庫 name 欄位**完全一致**（含空格、括號）
- 執行記錄會列出未配對的商品名稱，依提示修正後重新執行
- 圖片網址若為 Google Drive 分享連結，GAS 會自動轉換格式，無需手動處理

---

## 商品圖片管理

### 廠商提供圖片流程

1. 給廠商 `KC_廠商圖片範本.md`，請廠商填寫商品名稱 + Google Drive 分享連結
2. 收到後，把資料貼到 Sheets「**圖片清單**」工作表：
   - B欄：商品名稱（必須與商品庫完全一致）
   - C欄：圖片網址（任何格式皆可，GAS 自動轉換）
3. GAS 執行 `autoFillProductImages()` → 自動寫入商品庫

### 支援的圖片網址格式（自動轉換）

| 廠商給的格式 | 處理方式 |
|---|---|
| Google Drive `/file/d/ID/view` | 自動轉成可顯示格式 |
| Google Drive `open?id=ID` | 自動轉成可顯示格式 |
| Imgur 頁面 `imgur.com/xxx` | 自動補 `.jpg` |
| 直接圖片網址 `.jpg/.png` | 直接使用，不轉換 |

### 執行記錄說明

```
→ 已自動轉換 3 個網址格式
✓ 自動填入完成（嚴格比對）
  成功寫入：12 個商品
  名稱未完全符合，跳過：2 個
  ── 以下商品找不到完全一致的名稱 ──
    ✗ 泰式酸辣雞胸（180g）
```

跳過的商品 → 到「圖片清單」B欄把名稱改成與商品庫完全一致 → 再執行一次

### 自動關閉過期團購

GAS 新增 `createAutoCloseTrigger()` 函式，**執行一次**即可設定每日凌晨 1 點自動關閉到期團購：

1. Apps Script → 函式選單選 `createAutoCloseTrigger` → ▶ 執行
2. 允許授權（ScriptApp 需要額外授權）
3. 執行完成，往後每天自動跑，不需手動管理

---

## 日常維護

- GAS 每次修改都必須重新部署（建立新版本）
- 有新增欄位時，部署後執行 `setup()` 自動補欄，不影響現有資料
- Token 有效期 12 小時，閒置後需重新登入
- 備註範本上限 5 個，達上限需至 Sheets「行銷設定」刪除 `note_template_X`
- 優惠碼不允許重複（不分大小寫）
- GitHub Pages 更新需約 1 分鐘，看到舊版請 Ctrl+Shift+R
- 購物車 localStorage 有效期 24 小時，超過自動清空
- 圖片清單工作表可手動新增或修改，隨時可重新執行 `autoFillProductImages()` 覆寫
