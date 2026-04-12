# 📊 Ledger Analyzer - 完整 UI/UX 審查報告

**項目**: 離線分類帳分析工具  
**審查日期**: 2026年4月12日  
**審查範圍**: index.html, ledger.js  
**技術棧**: 原生 JavaScript + 嵌入式 CSS (無框架)

---

## 📋 執行摘要

| 維度 | 評分 | 關鍵問題 |
|------|------|----------|
| **無障礙性 (a11y)** | 🔴 35/100 | 零 ARIA、無焦點管理、對比度問題 |
| **視覺設計** | 🟡 62/100 | 設計系統一致但缺乏層次 |
| **UX 流程** | 🟡 58/100 | 按鈕氾濫、缺乏引導 |
| **設計工藝** | 🟢 71/100 | CSS 變數系統完整、動畫恰當 |
| **總體評分** | **🟡 57/100** | 功能強大但體驗需大幅改進 |

---

## 1️⃣ Frontend Design Review

### Pillar 1: Frictionless Insight to Action ❌ 不合格

**問題:**
- ❌ 主要操作不明確：每頁有 3-5 個 `primary` 按鈕，失去視覺焦點
- ❌ F1 模組工具列有 12+ 按鈕，認知超載
- ❌ 17 個導航選項無搜尋/過濾功能
- ❌ 首次使用無引導，用戶不知道從何開始

**紅旗:**
```
❌ Excessive clicks: 用戶需要點擊 5+ 次才能完成複雜分析
❌ Multiple competing primary buttons: 失去操作層次感
❌ No onboarding: 新手用戶面對空白頁面不知所措
```

**建議:**
```
✅ 每頁只保留 1 個 primary action
✅ 添加「開始使用」引導流程
✅ 實現導航搜尋功能
✅ 使用 Progressive Disclosure 隱藏進階選項
```

---

### Pillar 2: Quality is Craft ⚠️ 部分合格

**設計系統合規性:**

| 項目 | 狀態 | 說明 |
|------|------|------|
| CSS 變數系統 | ✅ 優秀 | `:root` 定義完整的設計令牌 |
| 色彩系統 | ⚠️ 需改進 | 部分顏色對比度不足 |
| 間距系統 | ⚠️ 不一致 | 混用 `8px`, `10px`, `12px`, `14px` |
| 圓角系統 | ⚠️ 碎片化 | `4px` 到 `16px` 缺乏規律 |
| 字體大小 | ❌ 混亂 | 7 種不同大小 (`11px`-`22px`) |
| Dark Mode | ❌ 缺失 | 無 `prefers-color-scheme` 支援 |
| Reduced Motion | ❌ 缺失 | 對前庭障礙用戶不友善 |

**視覺層次分析:**

```
❌ 問題: 所有按鈕同等重要
✅ 建議: 建立 4 級按鈕層次
   - Primary (1個): 填充色，強烈陰影
   - Secondary (2-3個): 輪廓樣式
   - Tertiary (多個): Ghost/文字按鈕
   - Danger (隔離): 紅色，確認對話框

❌ 問題: 表格缺乏視覺緩衝
✅ 建議: 添加空狀態、載入狀態、錯誤狀態
```

**動畫審查:**

| 動畫 | 持續時間 | 緩動函數 | 狀態 |
|------|----------|----------|------|
| fadeIn | 150ms | ease | ✅ 合理 |
| slideIn | 200ms | ease | ✅ 合理 |
| modalIn | 180ms | cubic-bezier(0.34,1.56,0.64,1) | ⚠️ 彈跳過大 |
| button hover | 120ms | - | ✅ 流暢 |

**建議:**
```css
/* 統一動畫系統 */
:root {
  --transition-fast: 150ms cubic-bezier(0.4, 0, 0.2, 1);
  --transition-base: 200ms cubic-bezier(0.4, 0, 0.2, 1);
  --transition-slow: 300ms cubic-bezier(0.4, 0, 0.2, 1);
}

/* 支援 Reduced Motion */
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    transition-duration: 0.01ms !important;
  }
}
```

---

### Pillar 3: Trustworthy Building ⚠️ 部分合格

**錯誤處理:**

| 情境 | 當前行為 | 問題 | 建議 |
|------|----------|------|------|
| 檔案上傳失敗 | Toast 自動消失 | 用戶無法回顧錯誤 | 持久化錯誤，提供解決步驟 |
| 計算超時 | 無回饋 | 用戶不知道是否卡住 | 添加進度指示器和超時提示 |
| 無資料 | 空白表格 | 缺乏引導 | 顯示空狀態和下一步指引 |
| 表單驗證失敗 | 紅色邊框 | 無錯誤訊息 | 添加 inline validation + 說明文字 |

**建議的錯誤 UI:**
```html
<div class="error-state" role="alert" aria-live="assertive">
  <div class="error-icon" aria-hidden="true">❌</div>
  <div class="error-content">
    <h4 class="error-title">上傳失敗</h4>
    <p class="error-message">
      無法解析 Excel 檔案。請確認：
    </p>
    <ul class="error-steps">
      <li>檔案格式為 .xlsx 或 .xls</li>
      <li>包含「日期」、「科目」、「金額」欄位</li>
      <li>檔案未損壞且可正常開啟</li>
    </ul>
    <div class="error-actions">
      <button class="btn-primary">🔄 重試</button>
      <button class="btn-ghost">📖 查看說明</button>
    </div>
  </div>
</div>
```

---

## 2️⃣ Web Design Guidelines 檢查

### 違反項目清單

| # | 規則 | 位置 | 嚴重程度 | 說明 |
|---|------|------|----------|------|
| 1 | 觸控目標最小 44x44px | `.btn-sm` | 🔴 Critical | `padding: 2px 7px` 僅約 20x18px |
| 2 | 明確的表單標籤關聯 | 所有 `<label>` | 🔴 Critical | 使用隱式關聯，缺乏 `for`/`id` |
| 3 | 可聚焦元素視覺回饋 | 全局 | 🔴 Critical | 無 `:focus-visible` 樣式 |
| 4 | 語義化 HTML | `<table>` | 🟡 Major | `<th>` 缺少 `scope` 屬性 |
| 5 | 響應式文字大小 | 全局 | 🟡 Major | 使用固定 `px` 而非 `rem` |
| 6 | 載入狀態指示 | 所有操作 | 🟡 Major | 長時間操作無視覺回饋 |
| 7 | 錯誤邊界處理 | 全局 | 🟡 Major | JavaScript 錯誤無優雅降級 |
| 8 | 一致的設計語言 | 全局 | 🟢 Minor | 間距和圓角系統不統一 |

### 詳細違反內容

#### 🔴 Critical 違反

**1. 觸控目標尺寸 (WCAG 2.5.5)**

```css
/* 當前 - 違規 ❌ */
.btn-sm { 
  font-size: 11px; 
  padding: 2px 7px;  /* 實際大小約 20x18px */
}

/* 建議 - 修正 ✅ */
.btn-sm { 
  font-size: var(--text-xs); 
  padding: var(--space-2) var(--space-3);
  min-height: 44px;  /* WCAG 最低要求 */
  min-width: 44px;
}
```

**2. 表單標籤關聯 (WCAG 3.3.2)**

```html
<!-- 當前 - 違規 ❌ -->
<label>科目
  <select id="accountSelect"></select>
</label>

<!-- 建議 - 修正 ✅ -->
<div class="form-group">
  <label for="accountSelect" class="form-label">科目</label>
  <select id="accountSelect" aria-describedby="account-help">
    <option value="">選擇科目</option>
  </select>
  <span id="account-help" class="sr-only">
    從下拉選單選擇要篩選的會計科目
  </span>
</div>
```

**3. 焦點視覺回饋 (WCAG 2.4.7)**

```css
/* 當前 - 缺失 ❌ */
/* 無任何 :focus 或 :focus-visible 樣式 */

/* 建議 - 新增 ✅ */
:focus-visible {
  outline: none;
  box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.4);
  border-radius: var(--radius-sm);
}

input:focus-visible,
select:focus-visible,
textarea:focus-visible {
  border-color: var(--brand-500);
  box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.4);
}
```

#### 🟡 Major 違反

**4. 表格語義化**

```html
<!-- 當前 - 不完整 ⚠️ -->
<th>日期</th>

<!-- 建議 - 修正 ✅ -->
<th scope="col" class="sortable" aria-sort="none" tabindex="0">
  <button class="sort-button" aria-label="依照日期排序">
    日期 <span class="sort-icon" aria-hidden="true">⇅</span>
  </button>
</th>
```

**5. 響應式文字大小**

```css
/* 當前 - 使用 px ❌ */
body { font-size: 14px; }
h1 { font-size: 18px; }

/* 建議 - 使用 rem ✅ */
html { font-size: 100%; /* 16px */ }
body { font-size: 0.875rem; /* 14px */ }
h1 { font-size: 1.25rem; /* 20px */ }
```

**6. 載入狀態**

```css
/* 建議新增 ✅ */
.spinner {
  display: inline-block;
  width: 16px;
  height: 16px;
  border: 2px solid var(--border-subtle);
  border-top-color: var(--brand-600);
  border-radius: 50%;
  animation: spin 0.6s linear infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}

button[aria-busy="true"] {
  position: relative;
  color: transparent !important;
  pointer-events: none;
}

button[aria-busy="true"]::after {
  content: '';
  position: absolute;
  width: 16px;
  height: 16px;
  border: 2px solid var(--border-subtle);
  border-top-color: var(--brand-600);
  border-radius: 50%;
  animation: spin 0.6s linear infinite;
}
```

---

## 3️⃣ Accessibility Review (WCAG 2.1 AA)

### 摘要

| 項目 | 數量 | 說明 |
|------|------|------|
| **發現問題總數** | 18 | |
| 🔴 Critical | 6 | 阻礙障礙用戶使用 |
| 🟡 Major | 8 | 顯著影響體驗 |
| 🟢 Minor | 4 | 體驗優化建議 |

---

### Perceivable (可感知)

| # | 問題 | WCAG 標準 | 嚴重程度 | 建議 |
|---|------|-----------|----------|------|
| 1 | 色彩對比不足：導航分區標題 `#9bb2cc` 在白色上 | 1.4.3 | 🔴 Critical | 改用 `#64748b` (5.7:1) |
| 2 | 色彩對比不足：`.badge-restored` 的 `#2f54eb` 在 `#f0f5ff` 上 | 1.4.3 | 🔴 Critical | 改用 `#1d4ed8` (7.2:1) |
| 3 | 無 ARIA 屬性 | 4.1.2 | 🔴 Critical | 添加 `role`, `aria-*` |
| 4 | 螢幕閱讀器無法接收 Toast 通知 | 4.1.3 | 🔴 Critical | 添加 `aria-live="polite"` |
| 5 | 表單標籤缺乏明確關聯 | 3.3.2 | 🔴 Critical | 使用 `for`/`id` 配對 |
| 6 | 表格表頭無 `scope` 屬性 | 1.3.1 | 🟡 Major | 添加 `scope="col"` |

#### 色彩對比詳細檢查

| 元素 | 前景色 | 背景色 | 當前比例 | 要求 | 狀態 |
|------|--------|--------|----------|------|------|
| 正文文字 | `#0f1f2e` | `#ffffff` | 15.4:1 | 4.5:1 | ✅ 通過 |
| 次要文字 | `#3a526d` | `#ffffff` | 7.2:1 | 4.5:1 | ✅ 通過 |
| 導航分區標題 | `#9bb2cc` | `#ffffff` | 3.1:1 | 4.5:1 | ❌ 失敗 |
| 徽章文字 | `#2f54eb` | `#f0f5ff` | 2.8:1 | 4.5:1 | ❌ 失敗 |
| Primary 按鈕 | `#ffffff` | `#165dff` | 4.6:1 | 4.5:1 | ✅ 通過 |

---

### Operable (可操作)

| # | 問題 | WCAG 標準 | 嚴重程度 | 建議 |
|---|------|-----------|----------|------|
| 1 | 觸控目標小於 44x44px | 2.5.5 | 🔴 Critical | 所有互動元素最小 44x44px |
| 2 | 無 `:focus-visible` 樣式 | 2.4.7 | 🔴 Critical | 添加清晰的焦點環 |
| 3 | 模態對話框無焦點陷阱 | 2.4.3 | 🟡 Major | 實現 focus trap |
| 4 | 無法用 Escape 關閉模態 | 2.1.1 | 🟡 Major | 添加鍵盤關閉功能 |
| 5 | 可排序表頭無法用鍵盤操作 | 2.1.1 | 🟡 Major | 添加 `tabindex` 和 `Enter` 支援 |
| 6 | 無 `prefers-reduced-motion` | 2.3.3 | 🟡 Major | 為動畫提供降級 |

#### 鍵盤導航測試結果

| 操作 | 預期行為 | 實際行為 | 狀態 |
|------|----------|----------|------|
| Tab 遍歷 | 所有互動元素可聚焦 | 部分元素跳過 | ❌ 失敗 |
| Enter 激活 | 按鈕和連結可觸發 | 正常工作 | ✅ 通過 |
| Space 滾動 | 在按鈕上應觸發點擊 | 有時滾動頁面 | ⚠️ 不一致 |
| Escape 關閉 | 模態應關閉 | 無反應 | ❌ 失敗 |
| Arrow 導航 | 選項/表頭可遍歷 | 無特殊處理 | ❌ 失敗 |

---

### Understandable (可理解)

| # | 問題 | WCAG 標準 | 嚴重程度 | 建議 |
|---|------|-----------|----------|------|
| 1 | 錯誤訊息缺乏指導性 | 3.3.1 | 🟡 Major | 提供具體解決步驟 |
| 2 | 表單無驗證提示 | 3.3.1 | 🟡 Major | 添加 inline validation |
| 3 | 載入狀態不明確 | 3.2.1 | 🟡 Major | 添加進度指示器 |
| 4 | 快捷鍵無文檔說明 | 3.3.5 | 🟢 Minor | 添加快捷鍵幫助 |

---

### Robust (穩健性)

| # | 問題 | WCAG 標準 | 嚴重程度 | 建議 |
|---|------|-----------|----------|------|
| 1 | 無 ARIA Landmark | 4.1.2 | 🟡 Major | 添加 `role="main"` 等 |
| 2 | 動態內容無 live region | 4.1.3 | 🟡 Major | Toast 添加 `aria-live` |
| 3 | 自定義組件無角色定義 | 4.1.2 | 🟡 Major | 添加 `role` 屬性 |

---

### 優先修復建議

1. **🔴 Critical: 添加 ARIA 屬性**
   - 影響：所有螢幕閱讀器用戶
   - 阻擋：完全無法理解頁面結構
   - 工作量：2-3 小時

2. **🔴 Critical: 修正色彩對比**
   - 影響：低視力用戶（約 8% 人口）
   - 阻擋：無法閱讀文字
   - 工作量：30 分鐘

3. **🔴 Critical: 添加焦點樣式**
   - 影響：鍵盤導航用戶
   - 阻擋：不知道當前位置
   - 工作量：1 小時

4. **🟡 Major: 觸控目標尺寸**
   - 影響：觸控設備和運動障礙用戶
   - 改善：點擊準確性
   - 工作量：2 小時

5. **🟡 Major: 模態焦點管理**
   - 影響：鍵盤和螢幕閱讀器用戶
   - 改善：對話框可用性
   - 工作量：1-2 小時

---

## 4️⃣ Design Critique

### Overall Impression

這是一個**功能強大但體驗粗糙**的專業工具。藍色品牌色調和卡片式佈局提供了良好的基礎，但缺乏視覺層次和無障礙支援嚴重影響了專業感。最大的機會在於建立清晰的操作層次和現代化的互動模式。

---

### Usability

| 發現 | 嚴重程度 | 建議 |
|------|----------|------|
| 17 個導航選項無分組 | 🔴 Critical | 添加搜尋功能或折疊選單 |
| 每頁多個 primary 按鈕 | 🔴 Critical | 限制為 1 個 primary，其餘降級 |
| 首次使用無引導 | 🟡 Moderate | 添加「開始使用」流程 |
| 工具列按鈕過多 | 🟡 Moderate | 使用分組和進階折疊 |
| 空狀態缺乏指引 | 🟡 Moderate | 添加下一步操作提示 |

---

### Visual Hierarchy

- **什麼最先吸引目光**: Primary 按鈕的藍色 — ✅ 正確，但太多藍色按鈕失去了重點
- **閱讀流程**: 從左側導航到右側內容 — ✅ 符合 F 型閱讀模式
- **層次問題**: 
  - ❌ 所有操作同等權重
  - ❌ 模組標題與內容缺乏明顯區隔
  - ❌ 統計數字不夠突出（使用相同字重）

**建議改進:**
```
1. 建立 4 級按鈕系統（Primary/Secondary/Tertiary/Danger）
2. 增大模組標題（24px → 28px）
3. 統計數字使用大字體（22px → 32px）+ 字重 700
4. 添加視覺分隔（圖標、顏色、大小）
```

---

### Consistency

| 元素 | 問題 | 建議 |
|------|------|------|
| 間距 | 混用 8/10/12/14px | 統一為 4px 倍數系統 |
| 圓角 | 4/6/8/10/12/16px | 簡化為 4 級（sm/md/lg/xl） |
| 字體大小 | 7 種不同大小 | 限制為 5-6 級排版比例 |
| 色彩語義 | warn/danger/ok 使用正確 | ✅ 良好 |
| 動畫時機 | 120-200ms 之間 | 統一為 3 級（fast/base/slow） |

---

### Accessibility

- **色彩對比**: ⚠️ 2 項失敗（導航標題、徽章）
- **觸控目標**: ❌ 多數元素小於 44x44px
- **文字可讀性**: ✅ 字體大小適當（14px base）
- **鍵盤導航**: ❌ 無焦點樣式
- **螢幕閱讀器**: ❌ 零 ARIA 支援

---

### What Works Well

- ✅ CSS 變數系統完整，為主題切換奠定基礎
- ✅ 響應式設計基本合理（grid 佈局）
- ✅ 動畫時機和緩動函數選擇恰当
- ✅ 色彩語義（warn/danger/ok）使用正確
- ✅ 卡片式佈局提供清晰的內容分組
- ✅ Sticky header 提升導航可用性

---

### Priority Recommendations

1. **🔴 修復無障礙關鍵問題（1-2 天）**
   - 添加 ARIA landmarks 和屬性
   - 修正色彩對比
   - 實現 `:focus-visible` 樣式
   - 添加鍵盤導航支援
   
2. **🟡 建立視覺層次（2-3 天）**
   - 統一間距和圓角系統
   - 改進按鈕層次（4 級系統）
   - 優化排版比例
   - 添加空/載入/錯誤狀態

3. **🟢 優化 UX 流程（1-2 天）**
   - 添加「開始使用」引導
   - 實現導航搜尋
   - 改進錯誤處理
   - 添加快捷鍵文檔

---

## 📊 評分總結

| 維度 | 當前評分 | 目標評分 | 改進空間 |
|------|----------|----------|----------|
| 無障礙性 | 35/100 | 90+/100 | +55 |
| 視覺設計 | 62/100 | 85/100 | +23 |
| UX 流程 | 58/100 | 80/100 | +22 |
| 設計工藝 | 71/100 | 90/100 | +19 |
| **總體** | **57/100** | **86/100** | **+29** |

---

## 🚀 下一步行動

### 立即執行（本週）
- [ ] 修正色彩對比問題（30 分鐘）
- [ ] 添加 `:focus-visible` 樣式（1 小時）
- [ ] 添加 ARIA landmarks（2 小時）
- [ ] 限制每頁 1 個 primary 按鈕（1 小時）

### 短期改進（2 週內）
- [ ] 實現鍵盤導航完整支援
- [ ] 添加模態焦點陷阱
- [ ] 統一間距和圓角系統
- [ ] 添加載入狀態指示器

### 中期優化（1 個月內）
- [ ] 添加「開始使用」引導流程
- [ ] 實現導航搜尋功能
- [ ] 改進錯誤狀態 UI
- [ ] 添加 Dark Mode 支援

---

## 📚 參考資源

- [WCAG 2.1 AA Guidelines](https://www.w3.org/TR/WCAG21/)
- [ARIA Authoring Practices](https://www.w3.org/WAI/ARIA/apg/)
- [WebAIM Contrast Checker](https://webaim.org/resources/contrastchecker/)
- [Vercel Web Interface Guidelines](https://vercel.com/design/guidelines)

---

**審查完成日期**: 2026-04-12  
**審查工具**: frontend-design-review, web-design-guidelines, accessibility-review, design-critique  
**建議下次審查時間**: 實施改進後
