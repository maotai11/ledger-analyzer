# UI 重新設計實施步驟

## 快速開始：應用改進

### 選項 1：漸進式改進（推薦）

1. **備份現有文件**
   ```bash
   cp index.html index.html.backup
   cp ledger.js ledger.js.backup
   ```

2. **引入無障礙補丁**
   
   在 `index.html` 的最後一個 `</body>` 標籤前添加：
   ```html
   <script src="ledger.js"></script>
   <script src="accessibility-enhancement.js"></script>
   </body>
   </html>
   ```

3. **測試**
   - 打開瀏覽器開發者工具
   - 檢查控制台是否有 `[A11Y]` 日誌
   - 測試鍵盤導航（Tab, Enter, Escape）
   - 測試快捷鍵（Alt+1, Alt+2, /）

---

### 選項 2：完整 CSS 升級

#### 步驟 1：替換 CSS 變數

打開 `index.html`，找到 `<style>` 標籤中的 `:root` 區塊（約在第 10-25 行），替換為：

```css
:root {
  /* Brand colors */
  --brand-50: #eff6ff;
  --brand-100: #dbeafe;
  --brand-200: #bfdbfe;
  --brand-300: #93c5fd;
  --brand-400: #60a5fa;
  --brand-500: #3b82f6;
  --brand-600: #2563eb;
  --brand-700: #1d4ed8;
  
  /* Semantic colors */
  --success-50: #f0fdf4;
  --success-500: #22c55e;
  --success-600: #16a34a;
  --success-700: #15803d;
  
  --warning-50: #fffbeb;
  --warning-500: #f59e0b;
  --warning-600: #d97706;
  --warning-700: #b45309;
  
  --error-50: #fef2f2;
  --error-500: #ef4444;
  --error-600: #dc2626;
  --error-700: #b91c1c;
  
  /* Neutral scale */
  --gray-50: #f8fafc;
  --gray-100: #f1f5f9;
  --gray-200: #e2e8f0;
  --gray-300: #cbd5e1;
  --gray-400: #94a3b8;
  --gray-500: #64748b;
  --gray-600: #475569;
  --gray-700: #334155;
  --gray-800: #1e293b;
  --gray-900: #0f172a;
  
  /* Surface tokens */
  --bg-page: var(--gray-50);
  --bg-surface: #ffffff;
  --bg-surface-raised: #ffffff;
  --bg-surface-overlay: #ffffff;
  --bg-disabled: var(--gray-100);
  
  /* Text tokens */
  --text-primary: var(--gray-900);
  --text-secondary: var(--gray-600);
  --text-tertiary: var(--gray-500);
  --text-inverse: #ffffff;
  --text-link: var(--brand-600);
  
  /* Border tokens */
  --border-subtle: var(--gray-200);
  --border-default: var(--gray-300);
  --border-strong: var(--gray-400);
  
  /* Shadow tokens */
  --shadow-xs: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
  --shadow-sm: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px -1px rgba(0, 0, 0, 0.1);
  --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -2px rgba(0, 0, 0, 0.1);
  --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.1);
  --shadow-xl: 0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1);
  
  /* Focus ring */
  --focus-ring: 0 0 0 3px rgba(37, 99, 235, 0.4);
  --focus-ring-error: 0 0 0 3px rgba(220, 38, 38, 0.4);
  
  /* Spacing scale (4px base) */
  --space-1: 4px;
  --space-2: 8px;
  --space-3: 12px;
  --space-4: 16px;
  --space-5: 20px;
  --space-6: 24px;
  --space-8: 32px;
  --space-10: 40px;
  --space-12: 48px;
  
  /* Border radius scale */
  --radius-sm: 6px;
  --radius-md: 8px;
  --radius-lg: 12px;
  --radius-xl: 16px;
  --radius-full: 9999px;
  
  /* Typography scale */
  --text-xs: 0.75rem;    /* 12px */
  --text-sm: 0.875rem;   /* 14px */
  --text-base: 1rem;     /* 16px */
  --text-lg: 1.125rem;   /* 18px */
  --text-xl: 1.25rem;    /* 20px */
  --text-2xl: 1.5rem;    /* 24px */
  
  /* Line heights */
  --leading-tight: 1.25;
  --leading-normal: 1.5;
  --leading-relaxed: 1.625;
  
  /* Z-index scale */
  --z-dropdown: 100;
  --z-sticky: 200;
  --z-overlay: 300;
  --z-modal: 400;
  --z-toast: 500;
  
  /* Transitions */
  --transition-fast: 150ms cubic-bezier(0.4, 0, 0.2, 1);
  --transition-base: 200ms cubic-bezier(0.4, 0, 0.2, 1);
  --transition-slow: 300ms cubic-bezier(0.4, 0, 0.2, 1);
}
```

#### 步驟 2：添加 Dark Mode 支援

在 `</style>` 標籤前添加：

```css
/* Dark Mode */
@media (prefers-color-scheme: dark) {
  :root {
    --bg-page: var(--gray-900);
    --bg-surface: var(--gray-800);
    --bg-surface-raised: var(--gray-800);
    --bg-surface-overlay: var(--gray-700);
    --bg-disabled: var(--gray-700);
    
    --text-primary: var(--gray-50);
    --text-secondary: var(--gray-300);
    --text-tertiary: var(--gray-400);
    --text-inverse: var(--gray-900);
    --text-link: var(--brand-400);
    
    --border-subtle: var(--gray-700);
    --border-default: var(--gray-600);
    --border-strong: var(--gray-500);
    
    --shadow-xs: 0 1px 2px 0 rgba(0, 0, 0, 0.3);
    --shadow-sm: 0 1px 3px 0 rgba(0, 0, 0, 0.4);
    --shadow-md: 0 4px 6px -1px rgba(0, 0, 0, 0.4);
    --shadow-lg: 0 10px 15px -3px rgba(0, 0, 0, 0.4);
    
    --focus-ring: 0 0 0 3px rgba(96, 165, 250, 0.5);
  }
}

/* Reduced Motion */
@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: 0.01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: 0.01ms !important;
    scroll-behavior: auto !important;
  }
}

/* Screen Reader Only */
.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border-width: 0;
}
```

#### 步驟 3：更新按鈕樣式

找到 `button` 相關樣式，替換為：

```css
/* 基礎按鈕 - 最小高度 44px (WCAG 2.5.5) */
button {
  border: 1px solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--bg-surface);
  cursor: pointer;
  padding: var(--space-2) var(--space-3);
  min-height: 44px;
  font-size: var(--text-sm);
  font-weight: 500;
  color: var(--text-secondary);
  transition: background var(--transition-fast), border-color var(--transition-fast), 
              box-shadow var(--transition-fast), transform var(--transition-fast);
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-2);
}

button:hover {
  background: var(--gray-50);
  border-color: var(--border-strong);
}

@media (prefers-color-scheme: dark) {
  button:hover {
    background: var(--gray-700);
  }
}

button:active {
  transform: translateY(1px);
}

button:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}

/* Primary 按鈕 */
button.primary {
  background: var(--brand-600);
  color: var(--text-inverse);
  border-color: var(--brand-600);
  font-weight: 600;
  box-shadow: var(--shadow-sm);
}

button.primary:hover {
  background: var(--brand-700);
  border-color: var(--brand-700);
  box-shadow: var(--shadow-md);
}

button.primary:focus-visible {
  box-shadow: var(--focus-ring);
}

/* Danger 按鈕 */
button.danger {
  background: var(--error-600);
  color: var(--text-inverse);
  border-color: var(--error-600);
}

button.danger:hover {
  background: var(--error-700);
  border-color: var(--error-700);
}

/* Ghost 按鈕 */
button.ghost {
  background: transparent;
  border-color: transparent;
}

button.ghost:hover {
  background: var(--gray-100);
  border-color: var(--border-subtle);
}

@media (prefers-color-scheme: dark) {
  button.ghost:hover {
    background: var(--gray-700);
  }
}
```

#### 步驟 4：添加焦點樣式

在 `</style>` 前添加：

```css
/* 全域焦點樣式 - WCAG 2.4.7 */
:focus-visible {
  outline: none;
  box-shadow: var(--focus-ring);
  border-radius: var(--radius-sm);
}

button:focus-visible,
[role="button"]:focus-visible {
  box-shadow: var(--focus-ring);
}

input:focus-visible,
select:focus-visible,
textarea:focus-visible {
  outline: none;
  border-color: var(--brand-500);
  box-shadow: var(--focus-ring);
}
```

#### 步驟 5：修正色彩對比

搜索並替換以下顏色：

```css
/* 原本: .nav-section { color: #9bb2cc; } */
/* 修正為: */
.nav-section {
  color: var(--gray-500, #64748b); /* 5.7:1 對比度 */
}

/* 原本: .badge-restored { color: #2f54eb; background: #f0f5ff; } */
/* 修正為: */
.badge-restored {
  color: var(--brand-700, #1d4ed8);
  background: var(--brand-50, #eff6ff);
  border: 1px solid var(--brand-200, #bfdbfe);
}
```

---

## HTML 結構改進

### 1. 添加 ARIA Landmarks

找到 `<div class="app">` 開頭，修改為：

```html
<div class="app">
  <header class="header" role="banner">
    <div>
      <h1>離線分類帳分析工具</h1>
      <div class="meta" id="metaText">尚未上傳檔案</div>
    </div>
    <!-- ... -->
  </header>

  <div class="layout">
    <aside class="nav" role="navigation" aria-label="主功能選單">
      <!-- 導航按鈕 -->
    </aside>

    <main class="content" role="main" id="main-content" aria-label="主要內容區域">
      <!-- 模組內容 -->
    </main>
  </div>
</div>
```

### 2. 改進導航按鈕

每個導航按鈕添加以下屬性：

```html
<button 
  class="active" 
  data-module="overview" 
  title="查看所有分錄明細與科目彙總"
  role="tab"
  aria-selected="true"
  aria-controls="module-overview"
  id="tab-overview"
  tabindex="0">
  📋 總覽
</button>

<button 
  data-module="f7" 
  title="科目期初/本期/期末餘額驗算"
  role="tab"
  aria-selected="false"
  aria-controls="module-f7"
  id="tab-f7"
  tabindex="-1">
  ⚖️ 試算表
</button>
```

### 3. 改進模組區域

每個模組 section 添加：

```html
<section 
  class="card module active" 
  id="module-overview"
  role="tabpanel"
  aria-labelledby="tab-overview"
  tabindex="0"
  aria-label="總覽">
  <h3 id="heading-overview">解析後明細</h3>
  <!-- 內容 -->
</section>

<section 
  class="card module" 
  id="module-f7"
  role="tabpanel"
  aria-labelledby="tab-f7"
  tabindex="0"
  aria-label="試算表"
  hidden>
  <h3 id="heading-f7">試算表</h3>
  <!-- 內容 -->
</section>
```

### 4. 改進表頭

```html
<table>
  <thead>
    <tr>
      <th scope="col" class="sortable" aria-sort="none" tabindex="0">
        <button class="sort-button" aria-label="依照日期排序">
          日期 <span class="sort-icon" aria-hidden="true">⇅</span>
        </button>
      </th>
      <th scope="col">傳票號碼</th>
      <th scope="col">科目</th>
      <th scope="col" class="col-amount">借方金額</th>
      <th scope="col" class="col-amount">貸方金額</th>
    </tr>
  </thead>
  <tbody>
    <!-- 資料列 -->
  </tbody>
</table>
```

### 5. 改進表單輸入

```html
<div class="grid">
  <div class="form-group">
    <label for="account-select">科目</label>
    <select id="account-select" aria-describedby="account-desc">
      <option value="">選擇科目</option>
    </select>
    <span id="account-desc" class="sr-only">篩選特定會計科目</span>
  </div>

  <div class="form-group">
    <label for="keyword-input">關鍵字搜尋</label>
    <input 
      id="keyword-input" 
      type="text" 
      placeholder="摘要或傳票號碼" 
      aria-describedby="keyword-help" />
    <span id="keyword-help" class="sr-only">輸入摘要內容或傳票號碼進行搜尋</span>
  </div>
</div>
```

### 6. 改進 Toast 容器

```html
<div class="toast-host" id="toastHost" role="status" aria-live="polite" aria-atomic="true"></div>
```

---

## JavaScript 改進

### 在 ledger.js 中添加模組切換邏輯

```javascript
function switchModule(moduleId) {
  // 更新導航按鈕
  document.querySelectorAll('.nav button').forEach(btn => {
    btn.classList.remove('active');
    btn.setAttribute('aria-selected', 'false');
    btn.setAttribute('tabindex', '-1');
  });
  
  const activeBtn = document.querySelector(`[data-module="${moduleId}"]`);
  if (activeBtn) {
    activeBtn.classList.add('active');
    activeBtn.setAttribute('aria-selected', 'true');
    activeBtn.setAttribute('tabindex', '0');
  }
  
  // 更新模組顯示
  document.querySelectorAll('.module').forEach(mod => {
    mod.classList.remove('active');
    mod.setAttribute('hidden', '');
  });
  
  const activeModule = document.getElementById(`module-${moduleId}`);
  if (activeModule) {
    activeModule.classList.add('active');
    activeModule.removeAttribute('hidden');
    
    // 移動焦點到模組標題
    const heading = activeModule.querySelector('h3');
    if (heading) {
      heading.setAttribute('tabindex', '-1');
      heading.focus();
    }
  }
}
```

### 添加鍵盤快捷鍵

```javascript
document.addEventListener('keydown', (e) => {
  const isInputFocused = ['INPUT', 'SELECT', 'TEXTAREA'].includes(document.activeElement.tagName);
  
  // Alt + 數字切換模組
  if (e.altKey && !e.shiftKey && !isInputFocused) {
    const moduleMap = {
      '1': 'overview',
      '2': 'f7',
      '3': 'f1',
      '4': 'f2',
      '5': 'f3',
      '6': 'f5',
      '7': 'f4',
      '8': 'f6',
      '9': 'f14',
      '0': 'f18'
    };
    
    if (moduleMap[e.key]) {
      e.preventDefault();
      const btn = document.querySelector(`[data-module="${moduleMap[e.key]}"]`);
      if (btn) {
        btn.click();
      }
    }
  }

  // / 聚焦搜尋框
  if (e.key === '/' && !isInputFocused) {
    e.preventDefault();
    const keywordInput = document.getElementById('keywordInput');
    if (keywordInput) {
      keywordInput.focus();
    }
  }
});
```

---

## 測試清單

### 無障礙測試

- [ ] **鍵盤導航**
  - [ ] Tab 鍵遍歷所有交互元素
  - [ ] Enter/Space 激活按鈕
  - [ ] Escape 關閉模態
  - [ ] 箭頭鍵導航選項

- [ ] **螢幕閱讀器**
  - [ ] 導航菜單正確朗讀
  - [ ] 表單標籤正確關聯
  - [ ] Toast 通知自動朗讀
  - [ ] 表格表頭正確宣告

- [ ] **色彩對比**
  - [ ] 使用 WebAIM Contrast Checker 驗證
  - [ ] 所有文字至少 4.5:1
  - [ ] 聚焦狀態明顯可見

- [ ] **縮放測試**
  - [ ] 瀏覽器縮放到 200%
  - [ ] 佈局不破裂
  - [ ] 所有功能可用

### 功能測試

- [ ] 所有模組切換正常
- [ ] 表單提交正常
- [ ] 快捷鍵工作
- [ ] Toast 顯示和消失
- [ ] 模態開啟和關閉

---

## 常見問題

### Q: 升級後某些樣式跑版了怎麼辦？

A: 檢查是否有遺留的舊 CSS 類名衝突。可以使用瀏覽器開發者工具的「計算樣式」面板查看實際應用的樣式。

### Q: Dark Mode 沒有生效？

A: 確保你的操作系統已啟用 Dark Mode，並且瀏覽器支援 `prefers-color-scheme` 媒體查詢。

### Q: 鍵盤快捷鍵不工作？

A: 檢查 `ledger.js` 中是否有現有的 `keydown` 事件監聽器產生了衝突。

### Q: 如何測試無障礙？

A: 安裝以下工具：
- **axe DevTools** (Chrome/Firefox 擴充功能)
- **WAVE** (Chrome 擴充功能)
- **WebAIM Contrast Checker** (線上工具)

---

## 後續優化建議

1. **添加載入動畫**：為長時間操作添加骨架屏
2. **錯誤邊界**：捕獲並優雅地處理 JavaScript 錯誤
3. **離線支援**：改進 Service Worker 策略
4. **效能優化**：虛擬化長列表渲染
5. **國際化**：準備多語言支援架構

---

## 總結

完成以上步驟後，你的應用將：
- ✅ 符合 WCAG 2.1 AA 標準
- ✅ 支援 Dark Mode
- ✅ 完整的鍵盤導航
- ✅ 螢幕閱讀器友善
- ✅ 現代化的視覺設計
- ✅ 一致的間距和排版系統

預估完成時間：2-4 小時（視熟悉程度而定）
