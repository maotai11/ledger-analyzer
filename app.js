const state = {
  transactions: [],
  items: new Map(),
  groups: [],
  groupRules: [],
  offsetMatches: [],
  offsetSuggested: [],
  offsetPoolOutputs: [],
  table: null,
  fuse: null,
};

const dom = {
  fileInput: document.getElementById('fileInput'),
  fileStatus: document.getElementById('fileStatus'),
  yearSelect: document.getElementById('yearSelect'),
  itemSelect: document.getElementById('itemSelect'),
  searchInput: document.getElementById('searchInput'),
  similarityThreshold: document.getElementById('similarityThreshold'),
  batchMoveBtn: document.getElementById('batchMoveBtn'),
  batchOffsetBtn: document.getElementById('batchOffsetBtn'),
  batchMarkBtn: document.getElementById('batchMarkBtn'),
  refreshGroupsBtn: document.getElementById('refreshGroupsBtn'),
  dataStats: document.getElementById('dataStats'),
  exportSummaryBtn: document.getElementById('exportSummaryBtn'),
  groupSummary: document.getElementById('groupSummary'),
  addGroupRuleBtn: document.getElementById('addGroupRuleBtn'),
  autoGroupBtn: document.getElementById('autoGroupBtn'),
  groupRuleList: document.getElementById('groupRuleList'),
  groupRulePanel: document.getElementById('groupRulePanel'),
  offsetTolerance: document.getElementById('offsetTolerance'),
  offsetMaxSize: document.getElementById('offsetMaxSize'),
  offsetPoolTargets: document.getElementById('offsetPoolTargets'),
  runOffsetBtn: document.getElementById('runOffsetBtn'),
  runOffsetPoolBtn: document.getElementById('runOffsetPoolBtn'),
  resetOffsetBtn: document.getElementById('resetOffsetBtn'),
  offsetResults: document.getElementById('offsetResults'),
  anomalyMethod: document.getElementById('anomalyMethod'),
  anomalyThreshold: document.getElementById('anomalyThreshold'),
  anomalyScope: document.getElementById('anomalyScope'),
  runAnomalyBtn: document.getElementById('runAnomalyBtn'),
  anomalyResults: document.getElementById('anomalyResults'),
  tabButtons: document.querySelectorAll('.tab-btn'),
  pages: document.querySelectorAll('.page'),
  ledgerHosts: document.querySelectorAll('.ledger-host'),
  groupRuleHosts: document.querySelectorAll('.group-rule-host'),
};

const DEFAULT_GROUPS = [
  { id: 'ungrouped', name: '未分類' },
];

const DEFAULT_CATEGORIES = ['未分類', '保險費', '租金', '水電費', '其他'];
const DEFAULT_PROJECTS = ['A', 'B', 'C', '其他'];
const DEFAULT_GROUP_NAMES = ['未分類', '群組1', '群組2'];

function toDecimal(value) {
  try {
    return new Decimal(String(value).replace(/,/g, '').trim() || 0);
  } catch {
    return new Decimal(0);
  }
}

function normalizeText(value) {
  if (!value) return '';
  return String(value)
    .replace(/_x000D_/g, ' ')
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseRocDate(text) {
  if (!text) return null;
  const m = String(text).match(/(\d{2,3})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const rocYear = Number(m[1]);
  const adYear = rocYear + 1911;
  const month = Number(m[2]);
  const day = Number(m[3]);
  return { text: m[0], rocYear, adYear, month, day, date: new Date(adYear, month - 1, day) };
}

function formatAmount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return '';
  return num.toLocaleString('zh-TW', { maximumFractionDigits: 2 });
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

function insuranceLabel(txn) {
  const summary = txn.summary || '';
  let keyword = '';
  if (summary.includes('勞健保')) keyword = '勞健保';
  else if (summary.includes('健保')) keyword = '健保';
  else if (summary.includes('勞保')) keyword = '勞保';
  if (!keyword || !txn.date) return null;
  const month = pad2(txn.date.getMonth() + 1);
  const day = pad2(txn.date.getDate());
  return `${txn.rocYear}年${month}-${day}${keyword}`;
}

function longestCommonPrefix(values) {
  if (!values.length) return '';
  let prefix = values[0];
  for (let i = 1; i < values.length; i += 1) {
    while (!values[i].startsWith(prefix) && prefix) {
      prefix = prefix.slice(0, -1);
    }
    if (!prefix) break;
  }
  return prefix.trim();
}

function longestCommonSuffix(values) {
  if (!values.length) return '';
  let suffix = values[0];
  for (let i = 1; i < values.length; i += 1) {
    while (!values[i].endsWith(suffix) && suffix) {
      suffix = suffix.slice(1);
    }
    if (!suffix) break;
  }
  return suffix.trim();
}

function smartGroupLabel(rule, list) {
  const summaries = list.map((t) => normalizeText(t.summary)).filter(Boolean);
  let base = rule.displayName || rule.name;
  if (summaries.length >= 2 && !rule.displayName) {
    const prefix = longestCommonPrefix(summaries);
    const suffix = longestCommonSuffix(summaries);
    const useSuffix = suffix.length >= 2 && suffix.length >= prefix.length;
    if (useSuffix) {
      const parts = summaries.map((s) => s.replace(suffix, '').trim()).filter(Boolean);
      if (parts.length) base = `${suffix}(${parts.join(',')})`;
    } else if (prefix.length >= 2) {
      const parts = summaries.map((s) => s.replace(prefix, '').trim()).filter(Boolean);
      if (parts.length) base = `${prefix}(${parts.join(',')})`;
    }
  }
  if (!rule.useDateRange) return base;
  const dates = list.map((t) => t.date).filter(Boolean).sort((a, b) => a - b);
  if (!dates.length) return base;
  const min = dates[0];
  const max = dates[dates.length - 1];
  const minRoc = min.getFullYear() - 1911;
  const maxRoc = max.getFullYear() - 1911;
  const minLabel = `${pad2(min.getMonth() + 1)}-${pad2(min.getDate())}`;
  const maxLabel = `${pad2(max.getMonth() + 1)}-${pad2(max.getDate())}`;
  const range = minRoc === maxRoc ? `${minRoc}年${minLabel}${minLabel === maxLabel ? '' : `~${maxLabel}`}` : `${minRoc}年${minLabel}~${maxRoc}年${maxLabel}`;
  return `${range}${base}`;
}

function similarity(a, b) {
  const left = normalizeText(a);
  const right = normalizeText(b);
  if (!left || !right) return 0;
  const dist = Levenshtein.get(left, right);
  const maxLen = Math.max(left.length, right.length) || 1;
  return 1 - dist / maxLen;
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: null });

  let headerRowIndex = -1;
  let headerMap = {};

  rows.forEach((row, idx) => {
    if (headerRowIndex !== -1) return;
    const dateIdx = row.findIndex((v) => String(v || '').includes('日期'));
    const voucherIdx = row.findIndex((v) => String(v || '').includes('傳票'));
    const summaryIdx = row.findIndex((v) => String(v || '').includes('摘要'));
    if (dateIdx !== -1 && voucherIdx !== -1 && summaryIdx !== -1) {
      headerRowIndex = idx;
      headerMap = {
        date: dateIdx,
        voucher: voucherIdx,
        summary: summaryIdx,
        debit: row.findIndex((v) => String(v || '').includes('借方')),
        credit: row.findIndex((v) => String(v || '').includes('貸方')),
        drcr: row.findIndex((v) => String(v || '').includes('借/貸')),
        balance: row.findIndex((v) => String(v || '').includes('餘額')),
      };
    }
  });

  if (headerRowIndex === -1) {
    throw new Error('找不到標題列，請確認檔案格式。');
  }

  const items = new Map();
  const transactions = [];
  let currentItem = null;

  for (let i = headerRowIndex + 1; i < rows.length; i += 1) {
    const row = rows[i];
    const first = normalizeText(row[0]);

    if (first.startsWith('項')) {
      const raw = normalizeText(row[0]);
      const match = raw.match(/項\s*目[:：]?\s*([0-9A-Za-z]+)\s*([^()]+)\((借|貸)\)/);
      if (match) {
        const code = match[1];
        const name = normalizeText(match[2]);
        const side = match[3];
        currentItem = { id: code, code, name, side };
      } else {
        currentItem = { id: raw, code: raw, name: raw, side: '' };
      }
      if (!items.has(currentItem.id)) {
        items.set(currentItem.id, { ...currentItem });
      }
      continue;
    }

    if (row[headerMap.date] === '日期') continue;

    if (['上期結轉', '月計:', '累計:'].includes(first)) continue;

    const dateInfo = parseRocDate(row[headerMap.date]);
    if (!dateInfo) continue;

    const summary = normalizeText(row[headerMap.summary]);
    const debit = toDecimal(row[headerMap.debit]).toNumber();
    const credit = toDecimal(row[headerMap.credit]).toNumber();
    const drcr = normalizeText(row[headerMap.drcr]);
    const amount = drcr.includes('借') ? (debit || credit) : drcr.includes('貸') ? (credit || debit) : (debit || credit);

    transactions.push({
      id: `${transactions.length + 1}`,
      itemId: currentItem ? currentItem.id : 'UNKNOWN',
      itemName: currentItem ? currentItem.name : 'UNKNOWN',
      itemSide: currentItem ? currentItem.side : '',
      dateText: dateInfo.text,
      date: dateInfo.date,
      rocYear: dateInfo.rocYear,
      adYear: dateInfo.adYear,
      voucher: normalizeText(row[headerMap.voucher]),
      summary,
      debit,
      credit,
      drcr,
      amount: amount || 0,
      balance: toDecimal(row[headerMap.balance]).toNumber(),
      group: '未分類',
      project: '',
      category: '',
      marked: false,
    });
  }

  return { transactions, items };
}

function initGroups() {
  state.groups = DEFAULT_GROUPS.map((g) => ({ ...g }));
  state.groupRules = [{ id: `gr-${Date.now()}`, name: '未分類', keywords: '', displayName: '', useDateRange: false }];
}

function populateSelectors() {
  const years = Array.from(new Set(state.transactions.map((t) => t.rocYear))).sort((a, b) => a - b);
  dom.yearSelect.innerHTML = `<option value="all">全部</option>` + years.map((y) => `<option value="${y}">${y}</option>`).join('');

  const items = Array.from(state.items.values());
  dom.itemSelect.innerHTML = `<option value="all">全部</option>` + items.map((item) => `<option value="${item.id}">${item.code} ${item.name}</option>`).join('');
}

function buildFuse() {
  state.fuse = new Fuse(state.transactions, {
    includeScore: true,
    keys: ['summary', 'voucher', 'itemName'],
    threshold: 0.35,
  });
}

function getFilteredTransactions() {
  let rows = state.transactions;
  const year = dom.yearSelect.value;
  const itemId = dom.itemSelect.value;

  if (year !== 'all') rows = rows.filter((t) => t.rocYear === Number(year));
  if (itemId !== 'all') rows = rows.filter((t) => t.itemId === itemId);

  const keyword = normalizeText(dom.searchInput.value);
  if (keyword) {
    const result = state.fuse.search(keyword).map((r) => r.item.id);
    const set = new Set(result);
    rows = rows.filter((t) => set.has(t.id));
  }

  return rows;
}

function renderTable() {
  if (state.table) {
    state.table.setData(getFilteredTransactions());
    return;
  }

  const selectValues = (field, defaults) => {
    const values = new Set(defaults);
    state.transactions.forEach((t) => {
      if (t[field]) values.add(t[field]);
    });
    return Array.from(values);
  };

  state.table = new Tabulator('#ledgerTable', {
    data: getFilteredTransactions(),
    layout: 'fitColumns',
    height: 520,
    selectable: true,
    reactiveData: true,
    groupBy: ['itemName'],
    groupHeader: function (value, count, data) {
      const total = data.reduce((acc, cur) => acc.plus(cur.amount || 0), new Decimal(0));
      const debitSum = data.reduce((acc, cur) => acc.plus(cur.debit || 0), new Decimal(0));
      const creditSum = data.reduce((acc, cur) => acc.plus(cur.credit || 0), new Decimal(0));
      const diff = debitSum.minus(creditSum);
      return `${value}（${count}筆） 合計:${formatAmount(total)} 借:${formatAmount(debitSum)} 貸:${formatAmount(creditSum)} 差額:${formatAmount(diff)}`;
    },
    columns: [
      { title: '日期', field: 'dateText', width: 90 },
      { title: '傳票號碼', field: 'voucher', width: 110, editor: 'input', headerFilter: 'input' },
      { title: '摘要', field: 'summary', editor: 'input', widthGrow: 2, headerFilter: 'input' },
      { title: '借/貸', field: 'drcr', width: 70 },
      { title: '借方', field: 'debit', editor: 'input', formatter: (cell) => formatAmount(cell.getValue()) },
      { title: '貸方', field: 'credit', editor: 'input', formatter: (cell) => formatAmount(cell.getValue()) },
      { title: '金額', field: 'amount', editor: 'input', formatter: (cell) => formatAmount(cell.getValue()) },
      { title: '餘額', field: 'balance', formatter: (cell) => formatAmount(cell.getValue()) },
      {
        title: '分類', field: 'category', editor: 'select',
        editorParams: () => ({ values: selectValues('category', DEFAULT_CATEGORIES) }),
      },
      {
        title: '專案', field: 'project', editor: 'select',
        editorParams: () => ({ values: selectValues('project', DEFAULT_PROJECTS) }),
      },
      {
        title: '群組', field: 'group', editor: 'select',
        editorParams: () => ({ values: selectValues('group', DEFAULT_GROUP_NAMES) }),
      },
      {
        title: '標記', field: 'marked', formatter: 'tickCross', editor: true, width: 70,
      },
    ],
    cellEdited: function (cell) {
      const row = cell.getRow().getData();
      const field = cell.getField();
      if (['debit', 'credit', 'amount'].includes(field)) {
        const amount = toDecimal(row.amount || 0).abs();
        if (field === 'amount') {
          if (row.drcr.includes('借')) {
            row.debit = amount.toNumber();
            row.credit = 0;
          } else if (row.drcr.includes('貸')) {
            row.credit = amount.toNumber();
            row.debit = 0;
          }
        } else if (field === 'debit' && row.drcr.includes('借')) {
          row.amount = toDecimal(row.debit || 0).toNumber();
        } else if (field === 'credit' && row.drcr.includes('貸')) {
          row.amount = toDecimal(row.credit || 0).toNumber();
        }
        cell.getRow().update(row);
        buildSummaryOutput();
      }
    },
  });
}

function mountLedgerTable(pageId) {
  const host = document.querySelector(`#${pageId} .ledger-host`);
  const tableEl = document.getElementById('ledgerTable');
  if (host && tableEl && !host.contains(tableEl)) {
    host.appendChild(tableEl);
  }
  if (state.table) state.table.redraw(true);
}

function mountGroupRules(pageId) {
  const host = document.querySelector(`#${pageId} .group-rule-host`);
  if (host && dom.groupRulePanel && !host.contains(dom.groupRulePanel)) {
    host.appendChild(dom.groupRulePanel);
  }
}

function updateStats() {
  const rows = getFilteredTransactions();
  dom.dataStats.textContent = `筆數：${rows.length}`;
}

function renderGroupRules() {
  if (!dom.groupRuleList) return;
  dom.groupRuleList.innerHTML = '';
  state.groupRules.forEach((rule) => {
    const card = document.createElement('div');
    card.className = 'group-rule-card';
    card.innerHTML = `
      <label>
        群組名稱
        <input data-field="name" data-id="${rule.id}" value="${rule.name}" />
      </label>
      <label>
        關鍵字（逗號分隔）
        <input data-field="keywords" data-id="${rule.id}" value="${rule.keywords}" />
      </label>
      <label>
        輸出名稱覆寫（可留空）
        <input data-field="displayName" data-id="${rule.id}" value="${rule.displayName}" />
      </label>
      <label class="inline">
        <input type="checkbox" data-field="useDateRange" data-id="${rule.id}" ${rule.useDateRange ? 'checked' : ''} />
        輸出含日期範圍
      </label>
      ${rule.name !== '未分類' ? `<button data-remove="${rule.id}">刪除</button>` : ''}
    `;
    dom.groupRuleList.appendChild(card);
  });

  dom.groupRuleList.querySelectorAll('input').forEach((input) => {
    const handler = (e) => {
      const rule = state.groupRules.find((r) => r.id === e.target.dataset.id);
      if (!rule) return;
      if (e.target.type === 'checkbox') {
        rule[e.target.dataset.field] = e.target.checked;
      } else {
        rule[e.target.dataset.field] = e.target.value;
      }
      buildSummaryOutput();
    };
    input.addEventListener('input', handler);
    input.addEventListener('change', handler);
  });

  dom.groupRuleList.querySelectorAll('button[data-remove]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.remove;
      state.groupRules = state.groupRules.filter((r) => r.id !== id);
      renderGroupRules();
      buildSummaryOutput();
    });
  });
}

function autoGroupByKeywords() {
  const rules = state.groupRules.filter((r) => r.keywords && r.keywords.trim());
  state.transactions.forEach((t) => {
    let assigned = '未分類';
    for (const rule of rules) {
      const keywords = rule.keywords.split(/[，,\\s]+/).map((k) => k.trim()).filter(Boolean);
      if (keywords.some((kw) => t.summary.includes(kw))) {
        assigned = rule.name;
        break;
      }
    }
    t.group = assigned;
  });
  renderTable();
  buildSummaryOutput();
}

function copySummary() {
  navigator.clipboard.writeText(dom.groupSummary.textContent || '').then(() => {
    dom.groupSummary.textContent += '\n\n(已複製)';
    setTimeout(buildSummaryOutput, 600);
  });
}

function resetOffset() {
  state.offsetMatches = [];
  state.offsetSuggested = [];
  state.offsetPoolOutputs = [];
  renderOffsetResults();
}

function runOffsetMatching() {
  state.offsetMatches = [];
  state.offsetSuggested = [];

  const rows = getFilteredTransactions();
  const tolerance = toDecimal(dom.offsetTolerance.value || 0);
  const maxSize = Number(dom.offsetMaxSize.value || 3);

  const debits = rows.filter((t) => t.drcr.includes('借'));
  const credits = rows.filter((t) => t.drcr.includes('貸'));

  const usedCredits = new Set();

  debits.forEach((d) => {
    const exact = credits.find((c) => !usedCredits.has(c.id)
      && toDecimal(Math.abs(d.amount)).minus(Math.abs(c.amount)).abs().lte(tolerance)
      && normalizeText(d.summary) === normalizeText(c.summary)
      && d.itemId === c.itemId);
    if (exact) {
      usedCredits.add(exact.id);
      state.offsetMatches.push({ id: `m-${Date.now()}-${state.offsetMatches.length}`, type: 'exact', debitId: d.id, creditIds: [exact.id] });
    }
  });

  const remainingCredits = credits.filter((c) => !usedCredits.has(c.id));
  const simThreshold = Number(dom.similarityThreshold.value || 0.82);

  debits.forEach((d) => {
    if (state.offsetMatches.some((m) => m.debitId === d.id)) return;
    const candidates = remainingCredits.filter((c) => c.itemId === d.itemId && similarity(d.summary, c.summary) >= simThreshold);
    if (!candidates.length) return;
    const target = toDecimal(Math.abs(d.amount));

    const combo = findBestCombo(candidates, target, maxSize, tolerance);
    if (combo.length) {
      state.offsetSuggested.push({ id: `s-${Date.now()}-${state.offsetSuggested.length}`, debitId: d.id, creditIds: combo.map((c) => c.id) });
    }
  });

  renderOffsetResults();
}

function findBestCombo(candidates, target, maxSize, tolerance) {
  let best = [];
  let bestDiff = target;

  function dfs(start, combo, sum) {
    const diff = sum.minus(target).abs();
    if (diff.lt(bestDiff)) {
      bestDiff = diff;
      best = combo.slice();
    }
    if (combo.length === maxSize || sum.gt(target.plus(tolerance))) return;
    for (let i = start; i < candidates.length; i += 1) {
      const next = candidates[i];
      dfs(i + 1, combo.concat(next), sum.plus(Math.abs(next.amount)));
      if (bestDiff.lte(tolerance)) return;
    }
  }

  dfs(0, [], new Decimal(0));
  return bestDiff.lte(tolerance) ? best : [];
}

function renderOffsetResults() {
  const txnById = new Map(state.transactions.map((t) => [t.id, t]));

  const matchedHtml = state.offsetMatches.length
    ? `<h4>已匹配</h4><ul>${state.offsetMatches.map((m) => {
      const debit = txnById.get(m.debitId);
      const credits = m.creditIds.map((id) => txnById.get(id)).filter(Boolean);
      const debitAmount = toDecimal(Math.abs(debit?.amount || 0));
      const creditSum = credits.reduce((acc, cur) => acc.plus(Math.abs(cur.amount || 0)), new Decimal(0));
      const diff = debitAmount.minus(creditSum);
      const diffText = diff.abs().gt(0) ? `（差額 ${formatAmount(diff)}）` : '';
      return `<li><span class="tag">${m.type}</span>借：${debit.dateText} ${debit.summary} ${formatAmount(debit.amount)} => 貸：${credits.map((c) => `${c.dateText} ${c.summary} ${formatAmount(c.amount)}`).join(' / ')} ${diffText}<button data-remove="${m.id}">移除</button></li>`;
    }).join('')}</ul>`
    : '<h4>已匹配</h4><div>尚無。</div>';

  const suggestedHtml = state.offsetSuggested.length
    ? `<h4>待確認（建議）</h4><ul>${state.offsetSuggested.map((m) => {
      const debit = txnById.get(m.debitId);
      const credits = m.creditIds.map((id) => txnById.get(id)).filter(Boolean);
      const debitAmount = toDecimal(Math.abs(debit?.amount || 0));
      const creditSum = credits.reduce((acc, cur) => acc.plus(Math.abs(cur.amount || 0)), new Decimal(0));
      const diff = debitAmount.minus(creditSum);
      const diffText = diff.abs().gt(0) ? `（差額 ${formatAmount(diff)}）` : '';
      return `<li><span class="tag">建議</span>借：${debit.dateText} ${debit.summary} ${formatAmount(debit.amount)} => 貸：${credits.map((c) => `${c.dateText} ${c.summary} ${formatAmount(c.amount)}`).join(' / ')} ${diffText}<button data-accept="${m.id}">採用</button></li>`;
    }).join('')}</ul>`
    : '<h4>待確認（建議）</h4><div>尚無。</div>';

  const unmatched = getFilteredTransactions().filter((t) => !state.offsetMatches.some((m) => m.debitId === t.id || m.creditIds.includes(t.id)));
  const bySummary = _.groupBy(unmatched, 'summary');
  const unmatchedHtml = Object.keys(bySummary).length
    ? `<h4>未沖銷摘要（可複製）</h4><div>${Object.entries(bySummary).map(([summary, list]) => {
      const total = list.reduce((acc, cur) => acc.plus(cur.amount || 0), new Decimal(0));
      return `${summary}(${formatAmount(total)})`;
    }).join('<br/>')}</div>`
    : '<h4>未沖銷摘要（可複製）</h4><div>全部已沖銷。</div>';

  const poolHtml = state.offsetPoolOutputs.length
    ? `<h4>沖銷數字池結果</h4><div>${state.offsetPoolOutputs.map((l) => `${l}`).join('<br/>')}</div>`
    : '<h4>沖銷數字池結果</h4><div>尚未執行。</div>';

  dom.offsetResults.innerHTML = `${matchedHtml}${suggestedHtml}${unmatchedHtml}${poolHtml}`;

  dom.offsetResults.querySelectorAll('button[data-accept]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.accept;
      const idx = state.offsetSuggested.findIndex((s) => s.id === id);
      if (idx === -1) return;
      const suggestion = state.offsetSuggested.splice(idx, 1)[0];
      state.offsetMatches.push({ id: `m-${Date.now()}-${state.offsetMatches.length}`, type: 'suggested', debitId: suggestion.debitId, creditIds: suggestion.creditIds });
      renderOffsetResults();
    });
  });

  dom.offsetResults.querySelectorAll('button[data-remove]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.remove;
      state.offsetMatches = state.offsetMatches.filter((m) => m.id !== id);
      renderOffsetResults();
    });
  });
}

function runOffsetPool() {
  const targets = dom.offsetPoolTargets.value.split(/[\n,]/).map((v) => toDecimal(v)).filter((v) => v.gt(0));
  if (!targets.length) {
    state.offsetPoolOutputs = ['請輸入沖銷數字池金額。'];
    renderOffsetResults();
    return;
  }
  const rows = getFilteredTransactions();
  const matchedIds = new Set(state.offsetMatches.flatMap((m) => [m.debitId, ...m.creditIds]));
  const pool = rows.filter((t) => !matchedIds.has(t.id));
  const maxSize = Number(dom.offsetMaxSize.value || 3);
  const tolerance = toDecimal(dom.offsetTolerance.value || 0);
  const outputs = [];

  targets.forEach((target) => {
    const byItem = _.groupBy(pool, 'itemId');
    let best = { combo: [], diff: null, itemName: '' };
    Object.values(byItem).forEach((list) => {
      const combo = findBestCombo(list, target, maxSize, tolerance);
      if (!combo.length) return;
      const sum = combo.reduce((acc, cur) => acc.plus(Math.abs(cur.amount || 0)), new Decimal(0));
      const diff = sum.minus(target).abs();
      if (best.diff === null || diff.lt(best.diff)) {
        best = { combo, diff, itemName: combo[0].itemName || combo[0].itemId };
      }
    });
    if (!best.combo.length) {
      outputs.push(`目標 ${formatAmount(target)}: 未找到可沖銷組合`);
      return;
    }
    outputs.push(`目標 ${formatAmount(target)}: ${best.combo.map((c) => `${c.dateText} ${c.summary} ${formatAmount(c.amount)}`).join(' / ')} (差額 ${formatAmount(best.diff)} | 項目 ${best.itemName})`);
  });

  state.offsetPoolOutputs = outputs;
  renderOffsetResults();
}

function runAnomaly() {
  const method = dom.anomalyMethod.value;
  const threshold = Number(dom.anomalyThreshold.value || 2.5);
  const scope = dom.anomalyScope.value;
  const rows = scope === 'all' ? state.transactions : getFilteredTransactions();
  const values = rows.map((t) => Math.abs(t.amount)).filter((v) => v > 0);
  if (!values.length) {
    dom.anomalyResults.textContent = '沒有可用的金額資料。';
    return;
  }

  let outputs = [];
  if (method === 'z') {
    const mean = values.reduce((a, b) => a + b, 0) / values.length;
    const variance = values.reduce((acc, v) => acc + (v - mean) ** 2, 0) / values.length;
    const std = Math.sqrt(variance) || 1;
    outputs = rows.filter((t) => Math.abs((Math.abs(t.amount) - mean) / std) >= threshold)
      .map((t) => `${t.dateText} ${t.voucher} ${t.summary} ${formatAmount(t.amount)}`);
  } else {
    const sorted = values.slice().sort((a, b) => a - b);
    const q1 = sorted[Math.floor(sorted.length * 0.25)];
    const q3 = sorted[Math.floor(sorted.length * 0.75)];
    const iqr = q3 - q1;
    const lower = q1 - threshold * iqr;
    const upper = q3 + threshold * iqr;
    outputs = rows.filter((t) => Math.abs(t.amount) < lower || Math.abs(t.amount) > upper)
      .map((t) => `${t.dateText} ${t.voucher} ${t.summary} ${formatAmount(t.amount)}`);
  }

  dom.anomalyResults.textContent = outputs.length ? outputs.join('\n') : '未偵測到異常金額。';
}

function handleBatchMove() {
  const rows = state.table.getSelectedData();
  if (!rows.length) return;
  const target = prompt('輸入要移動到的分類名稱');
  if (!target) return;
  rows.forEach((row) => {
    row.category = target;
  });
  state.table.replaceData(getFilteredTransactions());
  buildSummaryOutput();
}

function handleBatchOffset() {
  const rows = state.table.getSelectedData();
  if (!rows.length) return;
  const baseItem = rows[0].itemId;
  const sameItemRows = rows.filter((r) => r.itemId === baseItem);
  const debit = sameItemRows.filter((r) => r.drcr.includes('借'));
  const credit = sameItemRows.filter((r) => r.drcr.includes('貸'));
  if (!debit.length || !credit.length) return;
  debit.forEach((d) => {
    state.offsetMatches.push({ id: `m-${Date.now()}-${state.offsetMatches.length}`, type: 'manual', debitId: d.id, creditIds: credit.map((c) => c.id) });
  });
  renderOffsetResults();
}

function handleBatchMark() {
  const rows = state.table.getSelectedData();
  rows.forEach((row) => {
    row.marked = true;
  });
  state.table.replaceData(getFilteredTransactions());
}

function refreshGroupSummary() {
  buildSummaryOutput();
}

function buildSimilarSummarySuggestions() {
  const rows = getFilteredTransactions();
  const threshold = Number(dom.similarityThreshold.value || 0.82);
  const pairs = [];

  for (let i = 0; i < rows.length; i += 1) {
    for (let j = i + 1; j < rows.length; j += 1) {
      const sim = similarity(rows[i].summary, rows[j].summary);
      if (sim >= threshold && rows[i].summary && rows[j].summary) {
        pairs.push(`${rows[i].summary} <-> ${rows[j].summary} (${sim.toFixed(2)})`);
      }
    }
  }

  return pairs.slice(0, 50);
}

function buildSummaryOutput() {
  const rows = getFilteredTransactions();
  const grouped = _.groupBy(rows, 'category');
  const lines = [];

  Object.entries(grouped).forEach(([category, list]) => {
    const total = list.reduce((acc, cur) => acc.plus(cur.amount || 0), new Decimal(0));
    const debitSum = list.reduce((acc, cur) => acc.plus(cur.debit || 0), new Decimal(0));
    const creditSum = list.reduce((acc, cur) => acc.plus(cur.credit || 0), new Decimal(0));
    lines.push(`${category || '未分類'} ${formatAmount(total)} (借${formatAmount(debitSum)} / 貸${formatAmount(creditSum)})`);
  });

  const groupedByRule = _.groupBy(rows, 'group');
  const groupLines = [];
  Object.entries(groupedByRule).forEach(([groupName, list]) => {
    const rule = state.groupRules.find((r) => r.name === groupName);
    const label = rule ? smartGroupLabel(rule, list) : (groupName || '未分類');
    const total = list.reduce((acc, cur) => acc.plus(cur.amount || 0), new Decimal(0));
    const debitSum = list.reduce((acc, cur) => acc.plus(cur.debit || 0), new Decimal(0));
    const creditSum = list.reduce((acc, cur) => acc.plus(cur.credit || 0), new Decimal(0));
    groupLines.push(`${label} ${formatAmount(total)} (借${formatAmount(debitSum)} / 貸${formatAmount(creditSum)})`);
  });
  if (groupLines.length) {
    lines.push('\n摘要分組：');
    lines.push(groupLines.join('\n'));
  }

  const summaryItems = _.groupBy(rows, 'summary');
  const crossItem = [];
  Object.entries(summaryItems).forEach(([summary, list]) => {
    const itemSet = new Set(list.map((r) => r.itemName));
    if (summary && itemSet.size > 1) {
      crossItem.push(`${summary} -> ${Array.from(itemSet).join(', ')}`);
    }
  });
  if (crossItem.length) {
    lines.push('\n跨項目同摘要：');
    lines.push(crossItem.slice(0, 30).join('\n'));
  }

  const insuranceMap = new Map();
  rows.forEach((t) => {
    const label = insuranceLabel(t);
    if (!label) return;
    const total = insuranceMap.get(label) || new Decimal(0);
    insuranceMap.set(label, total.plus(t.amount || 0));
  });
  if (insuranceMap.size) {
    lines.push('\n保險費輸出：');
    insuranceMap.forEach((total, label) => {
      lines.push(`${label}(${formatAmount(total)})`);
    });
  }

  const similar = buildSimilarSummarySuggestions();
  if (similar.length) {
    lines.push('\n相似摘要建議：');
    lines.push(similar.join('\n'));
  }

  dom.groupSummary.textContent = lines.join('\n');
}

function copySummary() {
  navigator.clipboard.writeText(dom.groupSummary.textContent || '').then(() => {
    dom.groupSummary.textContent += '\n\n(已複製)';
    setTimeout(buildSummaryOutput, 600);
  });
}

function setFileStatus(text) {
  dom.fileStatus.textContent = text;
}

function wireEvents() {
  dom.tabButtons.forEach((btn) => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.page;
      dom.tabButtons.forEach((b) => b.classList.toggle('active', b === btn));
      dom.pages.forEach((page) => page.classList.toggle('active', page.id === target));
      mountLedgerTable(target);
      mountGroupRules(target);
      localStorage.setItem('ledger-active-page', target);
    });
  });

  dom.fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const result = parseWorkbook(evt.target.result);
        state.transactions = result.transactions;
        state.items = result.items;
        initGroups();
        populateSelectors();
        buildFuse();
        renderTable();
        updateStats();
        renderGroupRules();
        buildSummaryOutput();
        const activePage = document.querySelector('.page.active')?.id || 'page-ledger';
        mountLedgerTable(activePage);
        mountGroupRules(activePage);
        setFileStatus(`已載入 ${file.name}，共 ${state.transactions.length} 筆明細`);
      } catch (err) {
        setFileStatus(`解析失敗：${err.message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  [dom.yearSelect, dom.itemSelect, dom.searchInput, dom.similarityThreshold].forEach((el) => {
    el.addEventListener('input', () => {
      renderTable();
      updateStats();
      buildSummaryOutput();
    });
  });

  dom.batchMoveBtn.addEventListener('click', handleBatchMove);
  dom.batchOffsetBtn.addEventListener('click', handleBatchOffset);
  dom.batchMarkBtn.addEventListener('click', handleBatchMark);
  dom.refreshGroupsBtn.addEventListener('click', refreshGroupSummary);
  if (dom.addGroupRuleBtn) {
    dom.addGroupRuleBtn.addEventListener('click', () => {
      state.groupRules.push({ id: `gr-${Date.now()}`, name: '新群組', keywords: '', displayName: '', useDateRange: false });
      renderGroupRules();
    });
  }
  if (dom.autoGroupBtn) {
    dom.autoGroupBtn.addEventListener('click', autoGroupByKeywords);
  }

  dom.exportSummaryBtn.addEventListener('click', copySummary);
  dom.runOffsetBtn.addEventListener('click', runOffsetMatching);
  dom.runOffsetPoolBtn.addEventListener('click', runOffsetPool);
  dom.resetOffsetBtn.addEventListener('click', resetOffset);
  dom.runAnomalyBtn.addEventListener('click', runAnomaly);
}

initGroups();
wireEvents();
renderGroupRules();

const savedPage = localStorage.getItem('ledger-active-page');
if (savedPage) {
  const btn = Array.from(dom.tabButtons).find((b) => b.dataset.page === savedPage);
  if (btn) btn.click();
}
