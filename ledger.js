const DEFAULT_COLUMN_MAP = { date: 0, voucherNo: 4, summary: 10, debit: 14, credit: 17, drCr: 19, balance: 21 };

const AppState = {
  meta: { company: '', parsedAt: null, skippedRows: 0, igCount: { header: 0, monthly: 0, opening: 0, accountHeader: 0, crossPageDup: 0, invalid: 0, other: 0 }, columnMap: { ...DEFAULT_COLUMN_MAP } },
  accounts: {}, transactions: [],
  grouping: { groups: [], ungrouped: [], draftItems: [], mode: 'applied', draftRenameMap: {}, draftExtraGroups: [], copyText: '' },
  groupingStore: {}, activeGroupingKey: 'all',
  offset: { pairs: [], forcedUnmatchedIds: [], manualPairIds: [], manualMatches: [], lastUnmatchedIds: [], unmatchedGroups: [], copyText: '', suggestThreshold: 80, timeWindowDays: 14, subsetMaxK: 4, subsetTimeLimitMs: 1200, unmatchedView: 'review' },
  pool: { candidateIds: [], results: [], groups: [], ungrouped: [], copyText: '' }, anomaly: { results: [] }, crossLink: { mode: 'keyword', query: '', results: [] },
  gap: { results: [], periodicSuggestions: [] }, dupVoucher: { results: [] }, trendAlert: { results: [], groups: [], ungrouped: [], copyText: '' }, todos: [],
};

const hasFuse = typeof Fuse !== 'undefined';
const hasDecimal = typeof Decimal !== 'undefined';
let searchEngine = null;
let poolWorker = null;

const dom = {
  fileInput: document.getElementById('fileInput'), resetBtn: document.getElementById('resetBtn'), metaText: document.getElementById('metaText'),
  stats: document.getElementById('stats'), accountSelect: document.getElementById('accountSelect'), keywordInput: document.getElementById('keywordInput'),
  navButtons: Array.from(document.querySelectorAll('.nav button[data-module]')), modules: Array.from(document.querySelectorAll('.module')),
  overviewList: document.getElementById('overviewList'), runF1Btn: document.getElementById('runF1Btn'), applyF1Btn: document.getElementById('applyF1Btn'),
  f1NewGroupName: document.getElementById('f1NewGroupName'), f1AddGroupBtn: document.getElementById('f1AddGroupBtn'), exportF1Btn: document.getElementById('exportF1Btn'),
  copyF1TextBtn: document.getElementById('copyF1TextBtn'), f1CopyOutput: document.getElementById('f1CopyOutput'),
  f1Result: document.getElementById('f1Result'), f1List: document.getElementById('f1List'), runF2Btn: document.getElementById('runF2Btn'),
  exportF2Btn: document.getElementById('exportF2Btn'), copyF2TextBtn: document.getElementById('copyF2TextBtn'), f2ManualMatchBtn: document.getElementById('f2ManualMatchBtn'), f2ResetManualBtn: document.getElementById('f2ResetManualBtn'),
  f2Tolerance: document.getElementById('f2Tolerance'), f2Result: document.getElementById('f2Result'), f2UnmatchedSummary: document.getElementById('f2UnmatchedSummary'),
  f2List: document.getElementById('f2List'), f3Direction: document.getElementById('f3Direction'), f3Target: document.getElementById('f3Target'),
  f3Tolerance: document.getElementById('f3Tolerance'), runF3Btn: document.getElementById('runF3Btn'), f3Result: document.getElementById('f3Result'),
  f3List: document.getElementById('f3List'), runF4Btn: document.getElementById('runF4Btn'), exportF4Btn: document.getElementById('exportF4Btn'),
  f4Result: document.getElementById('f4Result'), f4List: document.getElementById('f4List'), f5Mode: document.getElementById('f5Mode'),
  f5Query: document.getElementById('f5Query'), f5Amount: document.getElementById('f5Amount'), f5Tolerance: document.getElementById('f5Tolerance'),
  runF5Btn: document.getElementById('runF5Btn'), f5Result: document.getElementById('f5Result'), f5List: document.getElementById('f5List'),
  f6Keyword: document.getElementById('f6Keyword'), f6Frequency: document.getElementById('f6Frequency'), f6CustomCount: document.getElementById('f6CustomCount'),
  f6From: document.getElementById('f6From'), f6To: document.getElementById('f6To'), runF6Btn: document.getElementById('runF6Btn'),
  f6Result: document.getElementById('f6Result'), f6List: document.getElementById('f6List'), runF14Btn: document.getElementById('runF14Btn'),
  exportF14Btn: document.getElementById('exportF14Btn'), f14Result: document.getElementById('f14Result'), f14List: document.getElementById('f14List'),
  f18Keyword: document.getElementById('f18Keyword'), f18Threshold: document.getElementById('f18Threshold'), runF18Btn: document.getElementById('runF18Btn'),
  copyF18TextBtn: document.getElementById('copyF18TextBtn'), exportF18Btn: document.getElementById('exportF18Btn'), f18Result: document.getElementById('f18Result'), f18List: document.getElementById('f18List'),
  addTodoBtn: document.getElementById('addTodoBtn'), todoToggleBtn: document.getElementById('todoToggleBtn'), todoPanel: document.getElementById('todoPanel'),
  todoVoucher: document.getElementById('todoVoucher'), todoContent: document.getElementById('todoContent'),
  copyTodoAllBtn: document.getElementById('copyTodoAllBtn'), todoList: document.getElementById('todoList'), todoBadge: document.getElementById('todoBadge'), toastHost: document.getElementById('toastHost'),
};

function escapeHtml(v) { return String(v ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;'); }
function cleanText(v) { return v == null ? '' : String(v).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }
function cleanSummary(v) { return cleanText(String(v ?? '').replace(/_x000D_\n/g, ' ').replace(/\r?\n/g, ' ')); }
function fmtAmount(v) { const n = Number(v); return Number.isFinite(n) ? n.toLocaleString('zh-TW', { maximumFractionDigits: 2 }) : ''; }
function dec(v) { if (!hasDecimal) return Number(v || 0) || 0; try { return new Decimal(String(v ?? '').replace(/,/g, '').trim() || 0); } catch { return new Decimal(0); } }
function decToNum(v) { if (hasDecimal && v && typeof v.toNumber === 'function') return v.toNumber(); return Number(v || 0); }
function decKey(v) { return hasDecimal ? dec(v).toFixed(2) : Number(v || 0).toFixed(2); }
function absDeltaWithin(a, b, t) { if (!hasDecimal) return Math.abs(Number(a) - Number(b)) <= Number(t); return dec(a).minus(dec(b)).abs().lte(dec(t)); }
function getSignedAmount(txn) {
  const debit = Number(txn.debit || 0);
  const credit = Number(txn.credit || 0);
  if (txn.accountNormalSide === '貸') return credit - debit;
  if (txn.accountNormalSide === '借') return debit - credit;
  return debit - credit;
}
function fmtSigned(v) {
  const n = Number(v || 0);
  const sign = n >= 0 ? '+' : '-';
  return `${sign}${fmtAmount(Math.abs(n))}`;
}
function asGroupAnchor(groupId) { return `f1-group-${String(groupId || '').replace(/[^A-Za-z0-9_-]/g, '')}`; }
function asScopedAnchor(scope, groupId) { return `${scope}-${String(groupId || '').replace(/[^A-Za-z0-9_-]/g, '')}`; }
function parseRocPeriod(text) { const m = cleanText(text).match(/^(\d{2,3})-(\d{1,2})$/); return m ? `${m[1]}-${String(Number(m[2])).padStart(2, '0')}` : null; }
function toast(msg, level = 'INFO') { const el = document.createElement('div'); el.className = 'toast'; el.innerHTML = `<strong>${level}</strong> ${escapeHtml(msg)}`; dom.toastHost.appendChild(el); setTimeout(() => el.remove(), level === 'ERROR' ? 5000 : 3000); }
function toRocYear(yearNum) { return yearNum > 1911 ? yearNum - 1911 : yearNum; }
function splitGroupNames(input) {
  return Array.from(new Set(
    String(input || '')
      .replace(/[，、；;]/g, '\n')
      .split(/\r?\n+/)
      .map((x) => cleanText(x))
      .filter(Boolean),
  ));
}
function formatSummaryAmountList(rows) {
  return rows.map((t) => `${t.summary || '(空白摘要)'}(${fmtSigned(getSignedAmount(t))})`).join('、');
}
function getIsoWeek(dateObj) {
  const d = new Date(Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return { year: d.getUTCFullYear(), week: weekNo };
}
function extractSummaryDates(summary, fallbackTxn) {
  const text = cleanText(summary || '');
  const out = [];
  const reYmd = /(\d{2,4})\s*(?:年|[\/\-.])\s*(\d{1,2})\s*(?:月|[\/\-.])\s*(\d{1,2})\s*日?/g;
  const reYm = /(\d{2,4})\s*(?:年|[\/\-.])\s*(\d{1,2})\s*月?/g;
  let m;
  while ((m = reYmd.exec(text)) !== null) {
    const y = Number(m[1]); const mm = Number(m[2]); const dd = Number(m[3]);
    const ad = y > 1911 ? y : y + 1911;
    const dt = new Date(ad, mm - 1, dd);
    if (!Number.isNaN(dt.getTime())) out.push(dt);
  }
  while ((m = reYm.exec(text)) !== null) {
    const y = Number(m[1]); const mm = Number(m[2]);
    const ad = y > 1911 ? y : y + 1911;
    const dt = new Date(ad, mm - 1, 1);
    if (!Number.isNaN(dt.getTime())) out.push(dt);
  }
  if (out.length) return out;
  return fallbackTxn?.date ? [fallbackTxn.date] : [];
}
function f6PeriodKeysFromTxn(txn, frequency) {
  const dates = extractSummaryDates(txn.summary, txn);
  const keys = dates.map((dt) => {
    const roc = dt.getFullYear() - 1911;
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    if (frequency === 'yearly') return `${roc}`;
    if (frequency === 'monthly' || frequency === 'custom') return `${roc}-${mm}`;
    if (frequency === 'daily') return `${roc}-${mm}-${dd}`;
    const wk = getIsoWeek(dt);
    return `${wk.year - 1911}-W${String(wk.week).padStart(2, '0')}`;
  });
  return Array.from(new Set(keys));
}
function emptyGroupingState() {
  return { groups: [], ungrouped: [], draftItems: [], mode: 'applied', draftRenameMap: {}, draftExtraGroups: [], copyText: '' };
}
function cloneGroupingState(s) {
  return {
    groups: (s?.groups || []).map((g) => ({
      id: g.id,
      name: g.name,
      transactionIds: [...(g.transactionIds || [])],
      rule: g.rule ? { mode: g.rule.mode || 'A', keyword: g.rule.keyword || '', threshold: Number(g.rule.threshold ?? 70) } : { mode: 'A', keyword: '', threshold: 70 },
    })),
    ungrouped: [...(s?.ungrouped || [])],
    draftItems: (s?.draftItems || []).map((d) => ({ txnId: d.txnId, proposed: d.proposed, source: d.source || d.proposed })),
    mode: s?.mode || 'applied',
    draftRenameMap: { ...(s?.draftRenameMap || {}) },
    draftExtraGroups: [...(s?.draftExtraGroups || [])],
    copyText: s?.copyText || '',
  };
}
function currentGroupingKey() { return dom.accountSelect.value || 'all'; }
function saveCurrentGroupingState() { AppState.groupingStore[AppState.activeGroupingKey || 'all'] = cloneGroupingState(AppState.grouping); }
function loadGroupingState(key) {
  const useKey = key || 'all';
  AppState.grouping = AppState.groupingStore[useKey] ? cloneGroupingState(AppState.groupingStore[useKey]) : emptyGroupingState();
  AppState.activeGroupingKey = useKey;
}

function extractPeriodsFromSummary(summary) {
  const out = [];
  const text = cleanText(summary);
  if (!text) return out;

  // 114/03, 2025-03, 114年03月, 2025年3月
  const re = /(\d{2,4})\s*(?:年|[\/\-.])\s*(\d{1,2})\s*月?/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const yy = toRocYear(Number(m[1]));
    const mm = Number(m[2]);
    if (Number.isFinite(yy) && Number.isFinite(mm) && mm >= 1 && mm <= 12) {
      out.push(`${yy}-${String(mm).padStart(2, '0')}`);
    }
  }

  // 114-01~03 / 2025-01至2025-03 / 114年1月到3月
  const range = text.match(/(\d{2,4})\s*(?:年|[\/\-.])\s*(\d{1,2})\s*月?\s*[~\-至到]+\s*(?:(\d{2,4})\s*(?:年|[\/\-.]))?\s*(\d{1,2})\s*月?/);
  if (range) {
    const y1 = toRocYear(Number(range[1]));
    const m1 = Number(range[2]);
    const y2 = range[3] ? toRocYear(Number(range[3])) : y1;
    const m2 = Number(range[4]);
    if (Number.isFinite(y1) && Number.isFinite(y2) && y1 === y2 && m1 >= 1 && m1 <= 12 && m2 >= 1 && m2 <= 12 && m1 <= m2) {
      for (let mth = m1; mth <= m2; mth += 1) out.push(`${y1}-${String(mth).padStart(2, '0')}`);
    }
  }

  return Array.from(new Set(out));
}

function stripDateTokens(summary) {
  return cleanText(summary)
    .replace(/\d{2,4}\s*(?:年|[\/\-.])\s*\d{1,2}\s*月?/g, ' ')
    .replace(/\d{1,2}\s*月/g, ' ')
    .replace(/[~\-至到]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function guessSummaryKeyword(rows) {
  const freq = new Map();
  rows.forEach((t) => {
    const stem = stripDateTokens(t.summary);
    if (!stem || stem.length < 2) return;
    freq.set(stem, (freq.get(stem) || 0) + 1);
  });
  let best = '';
  let bestN = 0;
  for (const [k, v] of freq.entries()) {
    if (v > bestN) {
      best = k;
      bestN = v;
    }
  }
  return best;
}

function parseROCDate(text) {
  const m = cleanText(text).match(/(\d{2,3})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!m) return null;
  const rocYear = Number(m[1]); const month = Number(m[2]); const day = Number(m[3]);
  const date = new Date(rocYear + 1911, month - 1, day); if (Number.isNaN(date.getTime())) return null;
  const mm = String(month).padStart(2, '0'); const dd = String(day).padStart(2, '0');
  return { date, dateISO: `${rocYear + 1911}-${mm}-${dd}`, dateROC: `${rocYear}-${mm}-${dd}`, periodROC: `${rocYear}-${mm}` };
}

function detectColumnMap(rows) {
  for (const row of rows) {
    const h = row.map((x) => cleanText(x));
    const date = h.findIndex((x) => x === '日期');
    const voucherNo = h.findIndex((x) => x.includes('傳票'));
    const summary = h.findIndex((x) => x === '摘要');
    if (date >= 0 && voucherNo >= 0 && summary >= 0) {
      return { date, voucherNo, summary, debit: h.findIndex((x) => x.includes('借方')), credit: h.findIndex((x) => x.includes('貸方')), drCr: h.findIndex((x) => x.includes('借/貸')), balance: h.findIndex((x) => x.includes('餘額')) };
    }
  }
  return { ...DEFAULT_COLUMN_MAP };
}

function parseAccountHeader(text) {
  const compact = cleanText(text).replace(/\s+/g, '');
  const m = compact.match(/^項目[:：]?([A-Za-z0-9]+)([^()]+)\((借|貸)\)/);
  if (m) return { code: m[1], name: m[2], normalSide: m[3] };
  const m2 = compact.match(/^項目[:：]?([A-Za-z0-9]+)(.+)$/);
  return m2 ? { code: m2[1], name: m2[2], normalSide: '' } : null;
}

function looksLikeHeaderRow(colA) {
  return ((colA.includes('分') && colA.includes('類') && colA.includes('帳')) || colA.includes('製表日期') || colA.includes('頁次') || colA === '日期' || colA.includes('公司'));
}

function pickSummary(row, map) {
  const direct = cleanSummary(row[map.summary]);
  if (direct) return direct;
  const deny = new Set([map.date, map.voucherNo, map.debit, map.credit, map.drCr, map.balance].filter((x) => x >= 0));
  for (let i = 0; i < row.length; i += 1) {
    if (deny.has(i)) continue;
    const cell = cleanSummary(row[i]);
    if (!cell || /^[\d,.-]+$/.test(cell)) continue;
    if (cell.length >= 2) return cell;
  }
  return '';
}
function parseWorkbook(arrayBuffer, fileName) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const map = detectColumnMap(rows);

  const accounts = {}; const txns = []; const seen = new Set();
  const ig = { header: 0, monthly: 0, opening: 0, accountHeader: 0, crossPageDup: 0, invalid: 0, other: 0 };
  let currentAccount = null;

  for (const row of rows) {
    const colA = cleanText(row[map.date]);
    if (!colA) { ig.other += 1; continue; }
    if (looksLikeHeaderRow(colA)) { ig.header += 1; continue; }
    if (colA.startsWith('月計:') || colA.startsWith('累計:')) { ig.monthly += 1; continue; }
    if (colA.startsWith('項')) {
      const acc = parseAccountHeader(colA);
      if (acc) {
        currentAccount = acc;
        if (!accounts[acc.code]) accounts[acc.code] = { ...acc, openingBalance: null, transactionIds: [] };
      }
      ig.accountHeader += 1; continue;
    }
    if (colA === '上期結轉') {
      if (currentAccount && map.balance >= 0 && accounts[currentAccount.code]) accounts[currentAccount.code].openingBalance = decToNum(dec(row[map.balance]));
      ig.opening += 1; continue;
    }

    const voucherNo = cleanText(row[map.voucherNo]); const d = parseROCDate(colA);
    if (!currentAccount || !voucherNo || !d) { ig.invalid += 1; continue; }

    const debit = map.debit >= 0 ? dec(row[map.debit]) : dec(0);
    const credit = map.credit >= 0 ? dec(row[map.credit]) : dec(0);
    const dupKey = `${currentAccount.code}||${voucherNo}||${d.dateISO}||${decKey(debit)}||${decKey(credit)}`;
    if (seen.has(dupKey)) { ig.crossPageDup += 1; continue; }
    seen.add(dupKey);

    const txn = {
      id: Math.random().toString(36).slice(2, 10), accountCode: currentAccount.code, accountName: currentAccount.name,
      accountNormalSide: currentAccount.normalSide, voucherNo, date: d.date, dateISO: d.dateISO, dateROC: d.dateROC,
      periodROC: d.periodROC, summary: pickSummary(row, map), debit: decToNum(debit), credit: decToNum(credit),
      drCr: map.drCr >= 0 ? cleanText(row[map.drCr]) : '', balance: map.balance >= 0 ? decToNum(dec(row[map.balance])) : 0,
    };
    txns.push(txn); accounts[currentAccount.code].transactionIds.push(txn.id);
  }

  AppState.accounts = accounts; AppState.transactions = txns; AppState.meta.company = fileName; AppState.meta.parsedAt = new Date().toISOString();
  AppState.meta.skippedRows = Object.values(ig).reduce((a, b) => a + b, 0); AppState.meta.igCount = ig; AppState.meta.columnMap = map;
  AppState.grouping = emptyGroupingState();
  AppState.groupingStore = {};
  AppState.groupingStore.all = cloneGroupingState(AppState.grouping);
  AppState.activeGroupingKey = 'all';
  AppState.offset.unmatchedGroups = [];

  if (hasFuse) searchEngine = new Fuse(txns, { keys: ['summary', 'voucherNo', 'accountName'], threshold: 0.35 });
  toast(`解析完成：${txns.length} 筆有效分錄｜${Object.keys(accounts).length} 個科目｜${ig.crossPageDup} 筆跨頁重複已合併｜${AppState.meta.skippedRows} 列已忽略`);
  renderBase();
}

function getFilteredTransactions() {
  let rows = AppState.transactions;
  const account = dom.accountSelect.value; const keyword = cleanText(dom.keywordInput.value);
  if (account && account !== 'all') rows = rows.filter((t) => t.accountCode === account);
  if (keyword) {
    if (searchEngine) {
      const ids = new Set(searchEngine.search(keyword).map((x) => x.item.id)); rows = rows.filter((t) => ids.has(t.id));
    } else rows = rows.filter((t) => t.summary.includes(keyword) || t.voucherNo.includes(keyword));
  }
  return rows;
}

function renderTxnList(el, rows, note = '', opts = {}) {
  const collapsible = opts.collapsible !== false;
  const show = rows.slice(0, 200);
  if (!collapsible) {
    el.innerHTML = `<p class="muted">${escapeHtml(note || `顯示 ${show.length}/${rows.length} 筆`)}</p><div class="table-wrap"><table><thead><tr><th>#</th><th class="col-date">日期</th><th class="col-voucher">傳票</th><th>科目</th><th class="col-summary">摘要</th><th class="col-amount">摘要金額(符號)</th><th class="col-amount">餘額</th></tr></thead><tbody>${show.map((t, idx) => `<tr><td>${idx + 1}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.voucherNo)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}${t.accountNormalSide ? `（${escapeHtml(t.accountNormalSide)}）` : ''}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td></tr>`).join('')}</tbody></table></div>`;
    return;
  }
  el.innerHTML = `<p class="muted">${escapeHtml(note || `顯示 ${show.length}/${rows.length} 筆`)}</p>${show.map((t, idx) => `<details class="card" style="margin:8px 0;padding:8px 10px;"><summary style="cursor:pointer;list-style:none;"><strong>${idx + 1}.</strong> ${escapeHtml(t.dateROC)}｜${escapeHtml(t.voucherNo)}｜[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}｜${escapeHtml((t.summary || '(空白摘要)').slice(0, 50))}｜${fmtSigned(getSignedAmount(t))}</summary><div style="margin-top:8px;"><div><strong>摘要：</strong>${escapeHtml(t.summary || '(空白摘要)')}</div><div><strong>日期：</strong>${escapeHtml(t.dateROC)} (${escapeHtml(t.dateISO)})</div><div><strong>傳票：</strong>${escapeHtml(t.voucherNo)}</div><div><strong>科目：</strong>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}${t.accountNormalSide ? `（${escapeHtml(t.accountNormalSide)}）` : ''}</div><div><strong>借方 / 貸方：</strong>${fmtAmount(t.debit)} / ${fmtAmount(t.credit)}</div><div><strong>摘要金額(符號)：</strong>${fmtSigned(getSignedAmount(t))}</div><div><strong>餘額：</strong>${fmtAmount(t.balance)}</div></div></details>`).join('')}`;
}

function renderBase() {
  renderAccountSelect(); renderStats();
  const rows = getFilteredTransactions();
  renderTxnList(dom.overviewList, rows, `目前篩選 ${rows.length} 筆`, { collapsible: false });
  renderTxnList(dom.f1List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f2List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f3List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f4List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f5List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f6List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f14List, rows, `目前篩選 ${rows.length} 筆`);
  renderTxnList(dom.f18List, rows, `目前篩選 ${rows.length} 筆`);
  if (AppState.grouping.mode === 'draft') renderF1Draft();
  else if (AppState.grouping.groups.length) renderF1Output();
  else {
    dom.f1Result.innerHTML = '<p class="muted">尚未分組。請先按「預覽分組名稱」。</p>';
    dom.f1CopyOutput.textContent = '(尚無分組摘要)';
    dom.applyF1Btn.textContent = '套用分組';
  }
}

function renderAccountSelect() {
  const current = dom.accountSelect.value || 'all';
  const options = ['<option value="all">全科目</option>'];
  Object.values(AppState.accounts).sort((a, b) => a.code.localeCompare(b.code)).forEach((a) => options.push(`<option value="${escapeHtml(a.code)}">[${escapeHtml(a.code)}] ${escapeHtml(a.name)}${a.normalSide ? `（${escapeHtml(a.normalSide)}）` : ''}</option>`));
  dom.accountSelect.innerHTML = options.join('');
  const exists = current === 'all' || Object.prototype.hasOwnProperty.call(AppState.accounts, current);
  dom.accountSelect.value = exists ? current : 'all';
}

function renderStats() {
  if (!AppState.meta.parsedAt) return;
  const ig = AppState.meta.igCount;
  dom.metaText.textContent = `檔案：${AppState.meta.company}｜解析時間：${new Date(AppState.meta.parsedAt).toLocaleString('zh-TW')}`;
  dom.stats.textContent = `有效分錄 ${AppState.transactions.length}｜科目 ${Object.keys(AppState.accounts).length}｜IGN-1:${ig.header} IGN-2:${ig.monthly} IGN-3:${ig.opening} IGN-4:${ig.accountHeader} IGN-5:${ig.crossPageDup} 其他:${ig.invalid + ig.other}｜摘要欄索引:${AppState.meta.columnMap.summary}`;
}

function normalizeForToken(s) { return cleanText(s).replace(/^[*@#＊＠＃]+/, ''); }
function bigramSet(s) { const t = cleanText(s); const out = new Set(); for (let i = 0; i < t.length - 1; i += 1) out.add(t.slice(i, i + 2)); return out; }
function jaccard(a, b) { const sa = bigramSet(a), sb = bigramSet(b); if (!sa.size || !sb.size) return 0; let inter = 0; sa.forEach((x) => { if (sb.has(x)) inter += 1; }); return inter / (sa.size + sb.size - inter); }

function buildF2UnmatchedSummary(rows) {
  const byAccount = new Map();
  rows.forEach((t) => {
    if (!byAccount.has(t.accountCode)) byAccount.set(t.accountCode, { name: t.accountName, items: [] });
    byAccount.get(t.accountCode).items.push(`${t.summary || '(空白摘要)'}(${fmtSigned(getSignedAmount(t))})`);
  });
  return Array.from(byAccount.entries()).map(([code, x]) => `[${code}] ${x.name}: ${x.items.join('、')}`).join(' | ');
}

function buildF2UnmatchedGroups(rows) {
  const map = new Map();
  rows.forEach((t) => {
    const name = deriveF1GroupName(t.summary);
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(t.id);
  });
  AppState.offset.unmatchedGroups = Array.from(map.entries()).map(([name, ids], idx) => ({ id: `u${idx + 1}`, name, transactionIds: ids }));
}

function renderF2UnmatchedOrganizer(rows) {
  const txMap = new Map(rows.map((r) => [r.id, r]));
  const options = AppState.offset.unmatchedGroups.map((g) => `<option value="${g.id}">${escapeHtml(g.name)}</option>`).join('');
  const copyItems = AppState.offset.unmatchedGroups.map((g) => {
    const list = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = list.reduce((a, b) => a + getSignedAmount(b), 0);
    return { text: `${g.name}(${fmtSigned(total)})`, anchorId: asScopedAnchor('f2-group', g.id) };
  });
  const copyRows = [];
  AppState.offset.unmatchedGroups.forEach((g) => {
    g.transactionIds.forEach((id) => { const t = txMap.get(id); if (t) copyRows.push(t); });
  });
  AppState.offset.copyText = formatSummaryAmountList(copyRows);
  const groupsHtml = AppState.offset.unmatchedGroups.map((g) => {
    const list = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = list.reduce((a, b) => a + getSignedAmount(b), 0);
    return `<details class="card" id="${asScopedAnchor('f2-group', g.id)}" style="margin:8px 0;" open><summary><input data-f2-rename="${g.id}" value="${escapeHtml(g.name)}" />（${list.length}筆，${fmtSigned(total)}） <button data-f2-del-group="${g.id}" style="margin-left:8px;">刪除群組</button></summary>
      <div style="margin-top:8px;"><button data-f2-back="1">回到分組摘要</button></div>
      <div style="margin-top:8px;"><label><input type="checkbox" data-f2-check-all="${g.id}" /> 全選</label> <select data-f2-batch-target="${g.id}"><option value="">批次移動到...</option>${options}</select> <button data-f2-batch-move="${g.id}">批次移動</button> <button data-f2-batch-del="${g.id}">批次刪除</button></div>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th><th>操作</th></tr></thead><tbody>
      ${list.map((t) => `<tr><td><input type="checkbox" data-f2-pick-item="${t.id}" data-f2-from="${g.id}" /> ${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td><td><select data-f2-target="${t.id}"><option value="">移動到...</option>${options}</select> <button data-f2-move="${t.id}" data-f2-from="${g.id}">移動</button> <button data-f2-del="${t.id}" data-f2-from="${g.id}">刪除</button></td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');
  dom.f2UnmatchedSummary.innerHTML = `${copyItems.length ? `<div class="muted">點選分組可跳到明細：</div><div style="margin-top:6px;">${copyItems.map((x) => `<button data-f2-jump="${x.anchorId}" style="margin:0 6px 6px 0;">${escapeHtml(x.text)}</button>`).join('')}</div>` : '<p class="muted">(目前無未沖帳)</p>'}${groupsHtml}`;
}

function daysBetween(a, b) {
  if (!(a instanceof Date) || !(b instanceof Date)) return null;
  const ms = Math.abs(a.getTime() - b.getTime());
  return Math.round(ms / 86400000);
}

function withinDays(a, b, winDays) {
  if (winDays == null) return true;
  const d = daysBetween(a, b);
  if (d == null) return true;
  return d <= winDays;
}

function sumByIds(txMap, ids, side) {
  return (ids || []).map((id) => txMap.get(id)).filter(Boolean).reduce((acc, t) => acc + Number(side === 'debit' ? t.debit : t.credit), 0);
}

function bestSummarySimilarityPct(txMap, leftIds, rightIds) {
  let best = 0;
  const left = (leftIds || []).map((id) => txMap.get(id)).filter(Boolean);
  const right = (rightIds || []).map((id) => txMap.get(id)).filter(Boolean);
  left.forEach((a) => {
    right.forEach((b) => {
      const s = jaccard(a.summary, b.summary) * 100;
      if (s > best) best = s;
    });
  });
  return best;
}

function subsetSumK(cands, target, tol, kMax, timeLimitMs) {
  // cands: [{id, amount}]
  const start = Date.now();
  const out = [];
  const list = cands.slice().sort((a, b) => b.amount - a.amount);

  function dfs(idx, picks, sum) {
    if (Date.now() - start > timeLimitMs) return;
    const delta = sum - target;
    if (Math.abs(delta) <= tol) {
      out.push({ ids: [...picks], total: sum, delta });
      // 不要爆量
      if (out.length >= 30) return;
      // 仍可繼續找更小 delta（但限制時間）
    }
    if (picks.length >= kMax) return;
    for (let i = idx; i < list.length; i += 1) {
      if (Date.now() - start > timeLimitMs) return;
      const next = list[i];
      const nextSum = sum + next.amount;
      // pruning: 若已經超過 target+tol 且接下來都是正數，仍可以繼續(因為可能找到別的組合)，但簡單剪枝
      if (nextSum > target + tol && picks.length + 1 >= kMax) continue;
      picks.push(next.id);
      dfs(i + 1, picks, nextSum);
      picks.pop();
      if (out.length >= 30) return;
    }
  }

  dfs(0, [], 0);
  // 依 abs(delta) 排序
  return out.sort((a, b) => Math.abs(a.delta) - Math.abs(b.delta));
}

function f2SuggestMatches(rows, tol, thresholdPct = 80, winDays = 14, kMax = 4, timeLimitMs = 1200) {
  const th = Math.max(0, Math.min(100, Number(thresholdPct) || 80));
  const txMap = new Map(rows.map((r) => [r.id, r]));
  const debits = rows.filter((x) => x.debit > 0);
  const credits = rows.filter((x) => x.credit > 0);
  const suggestions = [];

  // 借方多筆 ≈ 貸方單筆
  for (const c of credits) {
    const pool = debits
      .filter((d) => withinDays(d.date, c.date, winDays))
      .map((d) => ({ id: d.id, amount: Number(d.debit || 0) }))
      .filter((x) => x.amount > 0 && x.amount <= Number(c.credit || 0) + tol);

    if (pool.length < 2) continue;
    const subsets = subsetSumK(pool, Number(c.credit || 0), tol, kMax, timeLimitMs);
    const best = subsets.find((s) => s.ids.length >= 2);
    if (!best) continue;

    const debitIds = best.ids;
    const creditIds = [c.id];
    const sim = bestSummarySimilarityPct(txMap, debitIds, creditIds);
    if (sim < th) continue;

    const total = sumByIds(txMap, debitIds, 'debit');
    const delta = total - Number(c.credit || 0);
    const voucherMatch = debitIds.some((id) => cleanText(txMap.get(id)?.voucherNo) === cleanText(c.voucherNo));
    const dayDiff = Math.min(...debitIds.map((id) => daysBetween(txMap.get(id)?.date, c.date)).filter((x) => x != null), null);
    suggestions.push({
      id: `s_${c.id}_${Math.random().toString(36).slice(2, 7)}`,
      debitIds,
      creditIds,
      reason: { kind: `${debitIds.length}→1`, delta, simPct: sim, voucherMatch, dayDiff, target: Number(c.credit || 0) },
    });
  }

  // 貸方多筆 ≈ 借方單筆
  for (const d of debits) {
    const pool = credits
      .filter((c) => withinDays(d.date, c.date, winDays))
      .map((c) => ({ id: c.id, amount: Number(c.credit || 0) }))
      .filter((x) => x.amount > 0 && x.amount <= Number(d.debit || 0) + tol);

    if (pool.length < 2) continue;
    const subsets = subsetSumK(pool, Number(d.debit || 0), tol, kMax, timeLimitMs);
    const best = subsets.find((s) => s.ids.length >= 2);
    if (!best) continue;

    const debitIds = [d.id];
    const creditIds = best.ids;
    const sim = bestSummarySimilarityPct(txMap, debitIds, creditIds);
    if (sim < th) continue;

    const total = sumByIds(txMap, creditIds, 'credit');
    const delta = Number(d.debit || 0) - total;
    const voucherMatch = creditIds.some((id) => cleanText(txMap.get(id)?.voucherNo) === cleanText(d.voucherNo));
    const dayDiff = Math.min(...creditIds.map((id) => daysBetween(txMap.get(id)?.date, d.date)).filter((x) => x != null), null);
    suggestions.push({
      id: `s_${d.id}_${Math.random().toString(36).slice(2, 7)}`,
      debitIds,
      creditIds,
      reason: { kind: `1→${creditIds.length}`, delta, simPct: sim, voucherMatch, dayDiff, target: Number(d.debit || 0) },
    });
  }

  // 排序：先看 abs(delta) 再看相似度
  return suggestions
    .sort((a, b) => (Math.abs(a.reason.delta) - Math.abs(b.reason.delta)) || (b.reason.simPct - a.reason.simPct))
    .slice(0, 20);
}

function renderF2UnmatchedEditor(rows) {
  AppState.offset.lastUnmatchedIds = rows.map((r) => r.id);
  AppState.offset.copyText = formatSummaryAmountList(rows);

  const view = AppState.offset.unmatchedView || 'review';
  const tol = Number(dom.f2Tolerance.value || 0.01);
  const suggestTh = Number(AppState.offset.suggestThreshold ?? 80);

  // View: group（像 F1 一樣分組）
  if (view === 'group') {
    buildF2UnmatchedGroups(rows);
    // 在 organizer 上方加一個返回按鈕
    const header = `
      <div class="toolbar">
        <button data-f2-view="review">回到未沖帳清單</button>
        <button data-f2-suggest-refresh="1">重新掃描沖帳</button>
      </div>
      <div class="muted" style="margin-top:6px;">未沖帳分組模式：操作與 F1 類似（改名、批次移動、刪除群組…）。</div>
    `;
    renderF2UnmatchedOrganizer(rows);
    dom.f2UnmatchedSummary.innerHTML = header + dom.f2UnmatchedSummary.innerHTML;
    dom.f2List.innerHTML = '';
    return;
  }

  // View: review（先確認未沖帳是否有遺漏，再決定是否進入分組）
  const winDays = Number(AppState.offset.timeWindowDays ?? 14);
  const kMax = Number(AppState.offset.subsetMaxK ?? 4);
  const timeLimitMs = Number(AppState.offset.subsetTimeLimitMs ?? 1200);

  const suggestions = f2SuggestMatches(rows, tol, suggestTh, winDays, kMax, timeLimitMs);
  const txMap = new Map(rows.map((r) => [r.id, r]));

  const suggestHtml = suggestions.length
    ? `<div style="margin-top:6px;display:flex;flex-wrap:wrap;gap:6px;">
        ${suggestions.map((s) => {
          const d0 = txMap.get(s.debitIds?.[0]);
          const c0 = txMap.get(s.creditIds?.[0]);
          const label = `${s.reason.kind}｜Δ${fmtAmount(s.reason.delta)}｜${Math.round(s.reason.simPct)}%｜${s.reason.voucherMatch ? '同傳票' : '不同傳票'}｜±${s.reason.dayDiff ?? '-'}天`;
          const payload = `${(s.debitIds || []).join(',')}|${(s.creditIds || []).join(',')}`;
          const title = `${(d0?.summary || '').slice(0, 14)} ↔ ${(c0?.summary || '').slice(0, 14)}`;
          return `<button data-f2-suggest="${escapeHtml(payload)}" title="${escapeHtml(title)}">${escapeHtml(label)}</button>`;
        }).join('')}
      </div>`
    : `<div class="muted" style="margin-top:6px;">（目前無建議沖帳）</div>`;

  dom.f2UnmatchedSummary.innerHTML = `
    <div class="muted">未沖帳剩餘 ${rows.length} 筆。你可以先檢查是否有遺漏，再按「進入分組」用 F1 同款操作整理。</div>
    <div class="toolbar">
      <button data-f2-view="group">進入分組（未沖帳）</button>
      <button data-f2-suggest-refresh="1">重新掃描沖帳</button>
    </div>
    <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;align-items:end;">
      <label>建議沖帳相似度門檻(%)
        <input data-f2-suggest-th type="number" min="0" max="100" step="1" value="${escapeHtml(String(Number.isFinite(suggestTh) ? suggestTh : 80))}" style="width:110px;" />
      </label>
      <label>時間窗(天)
        <input data-f2-win type="number" min="0" max="365" step="1" value="${escapeHtml(String(Number.isFinite(winDays) ? winDays : 14))}" style="width:110px;" />
      </label>
      <label>最多筆數(k)
        <input data-f2-kmax type="number" min="2" max="8" step="1" value="${escapeHtml(String(Number.isFinite(kMax) ? kMax : 4))}" style="width:110px;" />
      </label>
      <button data-f2-suggest-refresh="1">更新建議</button>
    </div>
    <div class="muted" style="margin-top:6px;">建議沖帳（金額平衡 + 摘要高度雷同 + subset sum；點一下會自動勾選建議組合）：</div>
    ${suggestHtml}
  `;

  const maxShow = 800;
  const shown = rows.slice(0, maxShow);
  dom.f2List.innerHTML = `<p class="muted">未沖帳清單 ${rows.length} 筆（顯示前 ${shown.length} 筆）｜勾選 2 筆（一借一貸）→「勾選加入沖帳」</p>
    <div class="table-wrap"><table><thead><tr><th>勾選</th><th>日期</th><th>傳票</th><th>科目</th><th>摘要</th><th class="col-amount">簽帳金額</th></tr></thead><tbody>
    ${shown.map((t) => `<tr id="f2-row-${t.id}"><td><input type="checkbox" data-f2-pick="${t.id}" /></td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.voucherNo)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td></tr>`).join('')}
    </tbody></table></div>`;
}

function getF2UnmatchedRowsFromState() {
  const base = getFilteredTransactions();
  const map = new Map(base.map((t) => [t.id, t]));
  return AppState.offset.lastUnmatchedIds.map((id) => map.get(id)).filter(Boolean);
}

function deriveF1GroupName(summary) {
  // F1: 預設「保留原始摘要」當作分組名稱（只做最小清理）。
  // 這樣同傳票號、摘要雷同但只有數字不同時，不會被 regex 吃掉。
  let s = normalizeForToken(summary || '');
  if (!s) return '其他';

  // 去掉 Excel 解析常見雜訊與分隔符，保留主要文字與數字。
  if (s.includes('|')) s = cleanText(s.split('|').slice(-1)[0]);
  s = s.replace(/[()（）【】\[\]{}]/g, ' ').replace(/\s+/g, ' ').trim();

  // 避免組名太長（仍保留差異用的尾碼/數字）：
  // - 優先保留前面內容
  // - 太長就截斷
  const maxLen = 40;
  if (s.length > maxLen) s = `${s.slice(0, maxLen).trim()}…`;

  return s || '其他';
}

function ensureOtherGroup() {
  let g = AppState.grouping.groups.find((x) => x.name === '其他');
  if (!g) {
    g = { id: `g${Date.now()}`, name: '其他', transactionIds: [], rule: { mode: 'A', keyword: '', threshold: 70 } };
    AppState.grouping.groups.push(g);
  }
  if (!g.rule) g.rule = { mode: 'A', keyword: '', threshold: 70 };
  return g;
}

function f1NormalizeForMatch(s) {
  // 給「關鍵字選取」用：保留原始摘要但壓縮空白。
  return cleanText(s || '');
}

function f1RuleMatch(txn, rule) {
  const keyword = cleanText(rule?.keyword || '');
  if (!keyword) return false;
  const text = f1NormalizeForMatch(txn?.summary || '');
  if (!text) return false;

  const mode = rule?.mode || 'A';
  if (mode === 'B') {
    const th = Number(rule?.threshold ?? 70);
    const pct = Number.isFinite(th) ? th : 70;
    const score = jaccard(text, keyword) * 100;
    return score >= pct;
  }
  // A: 直接包含
  return text.includes(keyword);
}

function syncF1DraftEditsFromUI() {
  dom.f1Result.querySelectorAll('input[data-f1-draft-name]').forEach((inp) => {
    const from = inp.getAttribute('data-f1-draft-name') || '';
    const to = cleanText(inp.value) || from || '其他';
    if (from) AppState.grouping.draftRenameMap[from] = to;
  });
}
function syncF1AppliedEditsFromUI() {
  dom.f1Result.querySelectorAll('input[data-f1-rename]').forEach((inp) => {
    const gid = inp.getAttribute('data-f1-rename') || '';
    const g = AppState.grouping.groups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(inp.value) || g.name;
  });
}

function renderF1Draft() {
  dom.applyF1Btn.textContent = '套用分組';
  dom.f1CopyOutput.textContent = '(尚未套用分組)';
  const rows = getFilteredTransactions();
  const draft = (AppState.grouping.mode === 'draft' && AppState.grouping.draftItems.length)
    ? AppState.grouping.draftItems
    : rows.map((t) => {
      const base = deriveF1GroupName(t.summary);
      return { txnId: t.id, proposed: base, source: base };
    });
  AppState.grouping.draftItems = draft;
  AppState.grouping.mode = 'draft';

  const baseNames = Array.from(new Set(draft.map((d) => d.proposed)));
  const names = Array.from(new Set(baseNames.concat(AppState.grouping.draftExtraGroups))).sort((a, b) => a.localeCompare(b));
  const blocks = names.map((name, idx) => {
    const finalName = AppState.grouping.draftRenameMap[name] || name;
    return `<div class="card" style="margin:8px 0;padding:8px;"><strong>${idx + 1}.</strong>
      <div style="margin-top:6px;"><label>分組名稱 <input data-f1-draft-name="${escapeHtml(name)}" value="${escapeHtml(finalName)}" /></label> <button data-f1-draft-del="${escapeHtml(name)}">刪除</button></div>
    </div>`;
  }).join('');
  dom.f1Result.innerHTML = `<p class="muted">已產生 ${names.length} 個候選分組名稱（此步驟僅預覽名稱，不會先分組）。請先調整名稱，再按「套用分組」。</p>${blocks || '<p class="muted">無可分組資料。</p>'}`;
}

function applyF1GroupingFromDraft() {
  const draft = AppState.grouping.draftItems || [];
  if (!draft.length) {
    toast('請先按「預覽分組名稱」', 'WARN');
    return;
  }

  const renameMap = new Map();
  const originalNames = [];
  dom.f1Result.querySelectorAll('input[data-f1-draft-name]').forEach((inp) => {
    const from = inp.getAttribute('data-f1-draft-name') || '';
    const to = cleanText(inp.value) || '其他';
    renameMap.set(from, to);
    AppState.grouping.draftRenameMap[from] = to;
    if (from) originalNames.push(from);
  });

  const groups = new Map();
  draft.forEach((d) => {
    const target = renameMap.get(d.proposed) || d.proposed || '其他';
    if (!groups.has(target)) groups.set(target, []);
    groups.get(target).push(d.txnId);
  });

  // Keep user-added preview names as normal candidate groups after apply.
  originalNames.forEach((from) => {
    const target = renameMap.get(from) || from || '其他';
    if (!groups.has(target)) groups.set(target, []);
  });

  AppState.grouping.groups = Array.from(groups.entries()).map(([name, ids], idx) => ({ id: `g${idx + 1}`, name, transactionIds: ids, rule: { mode: 'A', keyword: '', threshold: 70 } }));
  AppState.grouping.ungrouped = [];
  AppState.grouping.mode = 'applied';
  ensureOtherGroup();
  renderF1Output();
}

function applyF1GroupingEdits() {
  if (!AppState.grouping.groups.length) {
    toast('目前沒有可套用的分組', 'WARN');
    return;
  }
  const merged = new Map();
  const order = [];
  const seenTxn = new Set();
  AppState.grouping.groups.forEach((g) => {
    const name = cleanText(g.name) || '其他';
    if (!merged.has(name)) {
      merged.set(name, []);
      order.push(name);
    }
    g.transactionIds.forEach((id) => {
      if (seenTxn.has(id)) return;
      seenTxn.add(id);
      merged.get(name).push(id);
    });
  });
  AppState.grouping.groups = order.map((name, idx) => {
    const prev = (AppState.grouping.groups || []).find((g) => cleanText(g.name) === name);
    return { id: `g${idx + 1}`, name, transactionIds: merged.get(name) || [], rule: prev?.rule ? { ...prev.rule } : { mode: 'A', keyword: '', threshold: 70 } };
  });
  AppState.grouping.ungrouped = AppState.grouping.ungrouped.filter((id) => !seenTxn.has(id));
  AppState.grouping.mode = 'applied';
  ensureOtherGroup();
  renderF1Output();
  toast('已套用編輯後分組');
}

function renderF1Output() {
  dom.applyF1Btn.textContent = '套用編輯後分組';
  const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
  const accountLastBalance = new Map();
  AppState.transactions.forEach((t) => accountLastBalance.set(t.accountCode, Number(t.balance || 0)));

  const groupOptions = AppState.grouping.groups.map((g) => `<option value="${g.id}">${escapeHtml(g.name)}</option>`).join('');
  const cards = AppState.grouping.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const sumSigned = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    const byAccount = new Map();
    txns.forEach((t) => {
      byAccount.set(t.accountCode, (byAccount.get(t.accountCode) || 0) + getSignedAmount(t));
    });
    const checks = Array.from(byAccount.entries()).map(([acc, total]) => {
      const lastBal = accountLastBalance.get(acc) || 0;
      const ok = Math.abs(total - lastBal) <= 0.01;
      return `<div class="${ok ? 'ok' : 'danger'}">[${escapeHtml(acc)}] 分組合計 ${fmtSigned(total)} / 最後餘額 ${fmtAmount(lastBal)} ${ok ? 'OK' : '不一致'}</div>`;
    }).join('');
    const anchorId = asGroupAnchor(g.id);
    const rule = g.rule || (g.rule = { mode: 'A', keyword: '', threshold: 70 });
    const mode = rule.mode || 'A';
    const th = Number.isFinite(Number(rule.threshold)) ? Number(rule.threshold) : 70;

    return `<details class="card" id="${anchorId}" style="margin:8px 0;" open>
      <summary style="cursor:pointer;"><strong>群組：</strong><input data-f1-rename="${g.id}" value="${escapeHtml(g.name)}" style="margin-left:8px;min-width:180px;" />｜筆數 ${txns.length}｜群組簽帳合計 ${fmtSigned(sumSigned)} <button data-f1-del-group="${g.id}" style="margin-left:8px;">刪除群組</button></summary>
      <div style="margin-top:8px;"><button data-f1-back="1">回到分組摘要</button></div>

      <div style="margin-top:8px;display:flex;flex-wrap:wrap;gap:8px;align-items:end;">
        <label style="min-width:220px;">本組關鍵字
          <input data-f1-rule-keyword="${g.id}" value="${escapeHtml(rule.keyword || '')}" placeholder="例如：勞保費 / 發票AA123" />
        </label>
        <label>模式
          <select data-f1-rule-mode="${g.id}">
            <option value="A" ${mode === 'A' ? 'selected' : ''}>A：包含</option>
            <option value="B" ${mode === 'B' ? 'selected' : ''}>B：相似度</option>
          </select>
        </label>
        <label>相似度門檻(%)
          <input data-f1-rule-threshold="${g.id}" type="number" min="0" max="100" step="1" value="${escapeHtml(String(th))}" style="width:110px;" />
        </label>
        <button data-f1-rule-select="${g.id}">只勾選命中</button>
        <button data-f1-rule-clear="${g.id}">清除勾選</button>
      </div>

      <div style="margin-top:8px;">${checks || '<div class="muted">此群組無資料</div>'}</div>
      <div style="margin-top:8px;">
        <label><input type="checkbox" data-f1-check-all="${g.id}" /> 全選</label>
        <select data-f1-batch-target="${g.id}" style="margin-left:8px;"><option value="">批次移動到...</option>${groupOptions}</select>
        <button data-f1-batch-move="${g.id}">批次移動</button>
        <button data-f1-batch-del="${g.id}">批次刪除</button>
      </div>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th><th>操作</th></tr></thead><tbody>
      ${txns.map((t) => `<tr><td><input type="checkbox" data-f1-pick="${t.id}" data-f1-from="${g.id}" /> ${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td><td><select data-f1-target="${t.id}"><option value="">移動到...</option>${groupOptions}</select> <button data-f1-move="${t.id}" data-f1-from="${g.id}">移動</button> <button data-f1-del="${t.id}" data-f1-from="${g.id}">刪除</button></td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');
  dom.f1Result.innerHTML = `<p class="muted">群組數 ${AppState.grouping.groups.length}｜未分組 ${AppState.grouping.ungrouped.length}</p>${cards || '<p class="muted">目前無群組。</p>'}`;

  const copyItems = AppState.grouping.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    return { text: `${g.name}(${fmtSigned(total)})`, anchorId: asGroupAnchor(g.id) };
  });
  const copyRows = [];
  AppState.grouping.groups.forEach((g) => {
    g.transactionIds.forEach((id) => { const t = txMap.get(id); if (t) copyRows.push(t); });
  });
  AppState.grouping.copyText = formatSummaryAmountList(copyRows);
  dom.f1CopyOutput.innerHTML = copyItems.length
    ? `<div class="muted">點選分組可跳到明細：</div><div style="margin-top:6px;">${copyItems.map((x) => `<button data-f1-jump="${x.anchorId}" style="margin:0 6px 6px 0;">${escapeHtml(x.text)}</button>`).join('')}</div>`
    : '(尚無分組摘要)';
}

function runF1Grouping() {
  AppState.grouping.groups = [];
  AppState.grouping.ungrouped = [];
  AppState.grouping.copyText = '';
  AppState.grouping.draftItems = [];
  AppState.grouping.draftRenameMap = {};
  AppState.grouping.draftExtraGroups = [];
  renderF1Draft();
}

function runF2() {
  const tol = Number(dom.f2Tolerance.value || 0.01);
  const rows = getFilteredTransactions();
  const txMap = new Map(rows.map((r) => [r.id, r]));

  // Drop stale ids not in current filtered rows
  AppState.offset.forcedUnmatchedIds = AppState.offset.forcedUnmatchedIds.filter((id) => txMap.has(id));
  AppState.offset.manualPairIds = AppState.offset.manualPairIds.filter((p) => txMap.has(p.debitId) && txMap.has(p.creditId));
  AppState.offset.manualMatches = (AppState.offset.manualMatches || []).filter((m) => (m.debitIds || []).every((id) => txMap.has(id)) && (m.creditIds || []).every((id) => txMap.has(id)));

  const forced = new Set(AppState.offset.forcedUnmatchedIds);
  const manualPairs = [];
  const used = new Set();

  AppState.offset.manualPairIds.forEach((p) => {
    const d = txMap.get(p.debitId);
    const c = txMap.get(p.creditId);
    if (!d || !c) return;
    manualPairs.push({
      id: `p_${d.id}_${c.id}`,
      kind: '1↔1',
      debitIds: [d.id],
      creditIds: [c.id],
      debitTotal: d.debit,
      creditTotal: c.credit,
      confidence: 'manual',
      reason: { delta: Number(d.debit || 0) - Number(c.credit || 0), simPct: jaccard(d.summary, c.summary) * 100, voucherMatch: cleanText(d.voucherNo) === cleanText(c.voucherNo), dayDiff: daysBetween(d.date, c.date) },
    });
    used.add(d.id);
    used.add(c.id);
  });

  // Multi-line manual matches
  (AppState.offset.manualMatches || []).forEach((m) => {
    const debitIds = (m.debitIds || []).filter((id) => txMap.has(id));
    const creditIds = (m.creditIds || []).filter((id) => txMap.has(id));
    if (!debitIds.length || !creditIds.length) return;

    debitIds.forEach((id) => used.add(id));
    creditIds.forEach((id) => used.add(id));

    const debitTotal = sumByIds(txMap, debitIds, 'debit');
    const creditTotal = sumByIds(txMap, creditIds, 'credit');
    manualPairs.push({
      id: m.id,
      kind: `${debitIds.length}↔${creditIds.length}`,
      debitIds,
      creditIds,
      debitTotal,
      creditTotal,
      confidence: 'manual',
      reason: m.reason || { delta: debitTotal - creditTotal, simPct: bestSummarySimilarityPct(txMap, debitIds, creditIds), voucherMatch: false, dayDiff: null },
    });
  });

  const debits = rows.filter((x) => x.debit > 0 && !forced.has(x.id) && !used.has(x.id));
  const credits = rows.filter((x) => x.credit > 0 && !forced.has(x.id) && !used.has(x.id));
  const autoPairs = [];

  debits.forEach((d) => {
    let best = null;
    let score = -1;
    credits.forEach((c) => {
      if (used.has(c.id)) return;
      if (!absDeltaWithin(d.debit, c.credit, tol)) return;
      const s = jaccard(d.summary, c.summary);
      if (s > score) { score = s; best = c; }
    });
    if (!best) return;
    used.add(d.id);
    used.add(best.id);
    autoPairs.push({ debit: d, credit: best, debitTotal: d.debit, creditTotal: best.credit, confidence: score >= 0.8 ? 'high' : score >= 0.5 ? 'medium' : 'low' });
  });

  const pairs = manualPairs.concat(autoPairs);
  AppState.offset.pairs = pairs;

  const matched = new Set();
  pairs.forEach((p) => {
    // autoPairs keep old shape (debit/credit objects)
    if (p?.debit?.id && p?.credit?.id) {
      matched.add(p.debit.id);
      matched.add(p.credit.id);
      return;
    }
    (p.debitIds || []).forEach((id) => matched.add(id));
    (p.creditIds || []).forEach((id) => matched.add(id));
  });
  const remain = rows.filter((r) => !matched.has(r.id) || forced.has(r.id));

  dom.f2Result.innerHTML = pairs.length
    ? `<div class="table-wrap"><table><thead><tr><th>類型</th><th>借方</th><th class="col-amount">借方合計</th><th>貸方</th><th class="col-amount">貸方合計</th><th>Δ</th><th>相似%</th><th>信心度</th><th>操作</th></tr></thead><tbody>${pairs.map((p, idx) => {
      // old autoPairs shape
      if (p?.debit?.id && p?.credit?.id) {
        const sim = jaccard(p.debit.summary, p.credit.summary) * 100;
        return `<tr><td>1↔1</td><td>${escapeHtml(p.debit.summary || '(空白摘要)')}</td><td class="col-amount">${fmtAmount(p.debitTotal)}</td><td>${escapeHtml(p.credit.summary || '(空白摘要)')}</td><td class="col-amount">${fmtAmount(p.creditTotal)}</td><td>${fmtAmount(Number(p.debitTotal || 0) - Number(p.creditTotal || 0))}</td><td>${Math.round(sim)}</td><td>${p.confidence}</td><td><button data-f2-unpair="${idx}">移動至未沖帳</button></td></tr>`;
      }
      const txMap2 = new Map(rows.map((r) => [r.id, r]));
      const dText = (p.debitIds || []).map((id) => txMap2.get(id)?.summary).filter(Boolean).slice(0, 2).join(' / ') || '(借方)';
      const cText = (p.creditIds || []).map((id) => txMap2.get(id)?.summary).filter(Boolean).slice(0, 2).join(' / ') || '(貸方)';
      const delta = Number(p.debitTotal || 0) - Number(p.creditTotal || 0);
      const sim = Number(p.reason?.simPct ?? bestSummarySimilarityPct(txMap2, p.debitIds, p.creditIds));
      return `<tr><td>${escapeHtml(p.kind || '多筆')}</td><td>${escapeHtml(dText)}</td><td class="col-amount">${fmtAmount(p.debitTotal)}</td><td>${escapeHtml(cText)}</td><td class="col-amount">${fmtAmount(p.creditTotal)}</td><td>${fmtAmount(delta)}</td><td>${Math.round(sim)}</td><td>${p.confidence}</td><td><button data-f2-unpair="${idx}">移動至未沖帳</button></td></tr>`;
    }).join('')}</tbody></table></div>`
    : `<p class="muted">命中 0 筆配對，未沖帳清單如下。</p>`;

  renderF2UnmatchedEditor(remain);
}

function bruteSubset(candidates, target, tolerance, timeLimit) {
  const start = Date.now(); const out = []; const n = candidates.length;
  const cents = candidates.map((c) => ({ ...c, cent: Math.round(c.amount * 100) }));
  const tgt = Math.round(target * 100); const tol = Math.round(tolerance * 100);
  let interrupted = false;
  for (let mask = 1; mask < (1 << n); mask += 1) {
    if (Date.now() - start > timeLimit) { interrupted = true; break; }
    let sum = 0; const picks = [];
    for (let i = 0; i < n; i += 1) if (mask & (1 << i)) { sum += cents[i].cent; picks.push(cents[i].id); }
    if (Math.abs(sum - tgt) <= tol) out.push({ id: Math.random().toString(36).slice(2, 10), transactionIds: picks, total: sum / 100, delta: (sum - tgt) / 100 });
    if (out.length >= 50) break;
  }
  return { results: out, interrupted, elapsed: Date.now() - start };
}

function ensureF3OtherGroup() {
  let g = AppState.pool.groups.find((x) => x.name === '其他');
  if (!g) {
    g = { id: `p${Date.now()}`, name: '其他', transactionIds: [] };
    AppState.pool.groups.push(g);
  }
  return g;
}

function buildF3GroupsFromResults(rows) {
  const rowIds = new Set(rows.map((r) => r.id));
  AppState.pool.groups = (AppState.pool.results || []).map((r, idx) => ({
    id: `p${idx + 1}`,
    name: `組合${idx + 1}`,
    transactionIds: Array.from(new Set((r.transactionIds || []).filter((id) => rowIds.has(id)))),
  }));
  const used = new Set(AppState.pool.groups.flatMap((g) => g.transactionIds));
  AppState.pool.ungrouped = AppState.pool.candidateIds.filter((id) => rowIds.has(id) && !used.has(id));
}

function renderF3Groups(rows, noteText = '') {
  const txMap = new Map(rows.map((t) => [t.id, t]));
  const groupOptions = AppState.pool.groups.map((g) => `<option value="${g.id}">${escapeHtml(g.name)}</option>`).join('');
  const cards = AppState.pool.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    const anchorId = asScopedAnchor('f3-group', g.id);
    return `<details class="card" id="${anchorId}" style="margin:8px 0;" open>
      <summary><strong>群組：</strong><input data-f3-rename="${g.id}" value="${escapeHtml(g.name)}" style="margin-left:8px;min-width:180px;" />｜筆數 ${txns.length}｜合計 ${fmtSigned(total)} <button data-f3-del-group="${g.id}" style="margin-left:8px;">刪除群組</button></summary>
      <div style="margin-top:8px;"><button data-f3-back="1">回到分組摘要</button></div>
      <div style="margin-top:8px;"><label><input type="checkbox" data-f3-check-all="${g.id}" /> 全選</label> <select data-f3-batch-target="${g.id}"><option value="">批次移動到...</option>${groupOptions}</select> <button data-f3-batch-move="${g.id}">批次移動</button> <button data-f3-batch-del="${g.id}">批次刪除</button></div>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th><th>操作</th></tr></thead><tbody>
      ${txns.map((t) => `<tr><td><input type="checkbox" data-f3-pick="${t.id}" data-f3-from="${g.id}" /> ${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td><td><select data-f3-target="${t.id}"><option value="">移動到...</option>${groupOptions}</select> <button data-f3-move="${t.id}" data-f3-from="${g.id}">移動</button> <button data-f3-del="${t.id}" data-f3-from="${g.id}">刪除</button></td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');

  const copyItems = AppState.pool.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    const detail = txns.map((t) => `${t.summary || '(空白摘要)'}(${fmtSigned(getSignedAmount(t))})`).join('、');
    return { buttonText: `${g.name}(${fmtSigned(total)})`, detailText: `${g.name}:${detail}`, anchorId: asScopedAnchor('f3-group', g.id) };
  });
  const copyRows = [];
  AppState.pool.groups.forEach((g) => {
    g.transactionIds.forEach((id) => { const t = txMap.get(id); if (t) copyRows.push(t); });
  });
  AppState.pool.copyText = formatSummaryAmountList(copyRows);

  dom.f3Result.innerHTML = `<p class="muted">${escapeHtml(noteText || `結果 ${AppState.pool.groups.length} 組`)}</p>
    <div class="toolbar"><button data-f3-copy="1">複製分組摘要</button></div>
    ${copyItems.length ? `<div class="muted" style="margin-top:6px;">點選分組可跳到明細：</div><div style="margin-top:6px;">${copyItems.map((x) => `<button data-f3-jump="${x.anchorId}" style="margin:0 6px 6px 0;">${escapeHtml(x.buttonText)}</button>`).join('')}</div>` : '<p class="muted">尚無命中分組</p>'}
    ${cards}`;

  const ungroupedRows = AppState.pool.ungrouped.map((id) => txMap.get(id)).filter(Boolean);
  renderTxnList(dom.f3List, ungroupedRows, `未分組 ${ungroupedRows.length} 筆`);
}

function runF3() {
  const rows = getFilteredTransactions(); const direction = dom.f3Direction.value;
  const target = Number(dom.f3Target.value || 0); const tolerance = Number(dom.f3Tolerance.value || 0.01);
  const candidates = rows.filter((x) => (direction === 'debit' ? x.debit > 0 : x.credit > 0)).map((x) => ({ id: x.id, amount: direction === 'debit' ? x.debit : x.credit }));
  AppState.pool.candidateIds = candidates.map((x) => x.id);

  if (candidates.length > 200) {
    AppState.pool.results = [];
    AppState.pool.groups = [];
    AppState.pool.ungrouped = AppState.pool.candidateIds.slice();
    AppState.pool.copyText = '';
    dom.f3Result.innerHTML = `<p class="danger">目前 ${candidates.length} 筆，請縮小至 200 筆以下。</p>`;
    renderTxnList(dom.f3List, rows.filter((r) => AppState.pool.candidateIds.includes(r.id)), `候選 ${candidates.length} 筆`);
    return;
  }

  if (candidates.length <= 30) {
    const r = bruteSubset(candidates, target, tolerance, 3000); AppState.pool.results = r.results;
    buildF3GroupsFromResults(rows);
    renderF3Groups(rows, r.results.length ? `結果 ${r.results.length} 組｜耗時 ${r.elapsed}ms${r.interrupted ? '｜已中斷' : ''}` : `命中 0 組，候選 ${candidates.length} 筆仍可見。`);
    return;
  }

  if (!poolWorker) poolWorker = new Worker('pool.worker.js');
  dom.runF3Btn.disabled = true; dom.f3Result.innerHTML = `<p class="muted">計算中...</p>`;
  poolWorker.onmessage = (ev) => {
    AppState.pool.results = ev.data.results || []; dom.runF3Btn.disabled = false;
    buildF3GroupsFromResults(rows);
    renderF3Groups(rows, AppState.pool.results.length ? `結果 ${AppState.pool.results.length} 組｜耗時 ${ev.data.elapsed}ms${ev.data.interrupted ? '｜已中斷' : ''}` : `命中 0 組，候選 ${candidates.length} 筆仍可見。`);
  };
  poolWorker.postMessage({ candidates, target, tolerance, timeLimit: 3000 });
}
function percentile(sorted, p) { if (!sorted.length) return 0; const idx = (sorted.length - 1) * p; const lo = Math.floor(idx); const hi = Math.ceil(idx); if (lo === hi) return sorted[lo]; return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo); }

function runF4() {
  const rows = getFilteredTransactions(); const results = [];
  const byA = new Map();
  rows.forEach((t) => {
    const amount = t.debit > 0 ? t.debit : t.credit;
    const k = `${t.accountCode}||${t.summary}||${decKey(amount)}`;
    if (!byA.has(k)) byA.set(k, []);
    byA.get(k).push(t);
  });
  byA.forEach((items) => { if (items.length >= 2 && new Set(items.map((x) => x.voucherNo)).size >= 2) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'A', severity: 'WARN', accountCode: items[0].accountCode, transactionIds: items.map((x) => x.id), description: '摘要+金額重複，傳票不同' }); });

  const amounts = rows.map((t) => Math.abs((t.debit || 0) - (t.credit || 0))).filter((x) => x > 0).sort((a, b) => a - b);
  if (amounts.length) {
    const q1 = percentile(amounts, 0.25), q3 = percentile(amounts, 0.75), iqr = q3 - q1, low = q1 - 1.5 * iqr, high = q3 + 1.5 * iqr;
    rows.forEach((t) => { const a = Math.abs((t.debit || 0) - (t.credit || 0)); if (a < low || a > high) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'B', severity: 'WARN', accountCode: t.accountCode, transactionIds: [t.id], description: 'IQR 離群' }); });
  }

  const byC = new Map();
  rows.forEach((t) => { const k = `${t.voucherNo}||${t.accountCode}`; if (!byC.has(k)) byC.set(k, []); byC.get(k).push(t); });
  byC.forEach((items) => { if (items.length >= 2) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'C', severity: 'INFO', accountCode: items[0].accountCode, transactionIds: items.map((x) => x.id), description: '同傳票同科目多筆' }); });

  rows.forEach((t) => {
    if (t.accountNormalSide === '借' && t.credit > 0) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'D', severity: 'WARN', accountCode: t.accountCode, transactionIds: [t.id], description: '借方科目出現貸方發生額' });
    if (t.accountNormalSide === '貸' && t.debit > 0) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'D', severity: 'WARN', accountCode: t.accountCode, transactionIds: [t.id], description: '貸方科目出現借方發生額' });
    const amount = Math.abs((t.debit || 0) - (t.credit || 0));
    if (amount >= 100000 && amount % 1000 === 0) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'E', severity: 'INFO', accountCode: t.accountCode, transactionIds: [t.id], description: '整數金額大額' });
    if (t.accountNormalSide === '借' && t.balance < 0) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'F', severity: 'ERROR', accountCode: t.accountCode, transactionIds: [t.id], description: '借方科目餘額<0' });
  });

  const digits = rows.map((t) => { const amount = Math.abs((t.debit || 0) - (t.credit || 0)); const m = String(Math.floor(amount)).match(/[1-9]/); return m ? Number(m[0]) : null; }).filter((x) => x != null);
  if (digits.length >= 30) {
    const exp = [0, 0.301, 0.176, 0.125, 0.097, 0.079, 0.067, 0.058, 0.051, 0.046]; const c = Array(10).fill(0); digits.forEach((d) => c[d] += 1);
    let chi2 = 0; for (let d = 1; d <= 9; d += 1) { const e = exp[d] * digits.length; chi2 += ((c[d] - e) ** 2) / (e || 1); }
    if (chi2 > 15.507) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'G', severity: 'WARN', accountCode: 'all', transactionIds: rows.map((r) => r.id), description: 'Benford chi2=' + chi2.toFixed(2) });
  }

  const byH = new Map();
  rows.forEach((t) => {
    const n = Number(cleanText(t.voucherNo).replace(/\D/g, ''));
    if (!Number.isFinite(n) || !n) return;
    const k = `${t.accountCode}||${t.periodROC}`;
    if (!byH.has(k)) byH.set(k, []);
    byH.get(k).push({ n, t });
  });

  // H: 以「科目 + 月份」分組，顯示缺的區間，並帶前後摘要方便展開查看。
  byH.forEach((items, key) => {
    items.sort((a, b) => a.n - b.n);
    for (let i = 1; i < items.length; i += 1) {
      const prev = items[i - 1];
      const curr = items[i];
      const gap = curr.n - prev.n;
      if (gap <= 1) continue;
      const missingFrom = prev.n + 1;
      const missingTo = curr.n - 1;
      const missText = missingFrom === missingTo ? `${missingFrom}` : `${missingFrom}~${missingTo}`;
      const [accountCode, periodROC] = String(key).split('||');
      const before = prev.t;
      const after = curr.t;
      const desc = `傳票缺號 ${prev.n}→${curr.n} 缺 ${missText}｜[${accountCode}] ${periodROC}`;

      results.push({
        id: Math.random().toString(36).slice(2, 10),
        rule: 'H',
        severity: 'WARN',
        accountCode: before.accountCode,
        transactionIds: [before.id, after.id],
        description: desc,
        meta: {
          accountCode,
          periodROC,
          beforeVoucher: before.voucherNo,
          afterVoucher: after.voucherNo,
          beforeSummary: before.summary,
          afterSummary: after.summary,
          missingFrom,
          missingTo,
        },
      });
    }
  });

  AppState.anomaly.results = results;
  const hit = new Set(results.flatMap((r) => r.transactionIds)); const remain = rows.filter((r) => !hit.has(r.id));
  if (!results.length) {
    dom.f4Result.innerHTML = `<p class="muted">命中 0 筆異常，剩餘 ${remain.length} 筆仍顯示。</p>`;
    renderTxnList(dom.f4List, remain, `未命中異常 ${remain.length} 筆`);
    return;
  }
  const txMap = new Map(rows.map((t) => [t.id, t]));
  const groupMap = new Map();
  results.forEach((r) => {
    const k = `${r.rule}||${r.severity}||${r.description}`;
    if (!groupMap.has(k)) groupMap.set(k, { rule: r.rule, severity: r.severity, description: r.description, ids: [] });
    groupMap.get(k).ids.push(...r.transactionIds);
  });
  const cards = Array.from(groupMap.values()).map((g, idx) => {
    const ids = Array.from(new Set(g.ids));
    const list = ids.map((id) => txMap.get(id)).filter(Boolean);
    return `<details class="card" style="margin:8px 0;" open><summary><strong>${idx + 1}. [${g.rule}] ${escapeHtml(g.description)}</strong>｜${g.severity}｜${list.length} 筆</summary>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>科目</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th></tr></thead><tbody>
      ${list.map((t) => `<tr><td>${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');
  dom.f4Result.innerHTML = `<p class="muted">異常分組 ${groupMap.size} 組</p>${cards}`;
  renderTxnList(dom.f4List, remain, `未命中異常 ${remain.length} 筆`);
}

function runF5() {
  const rows = getFilteredTransactions(); const mode = dom.f5Mode.value;
  const query = cleanText(dom.f5Query.value); const amount = Number(dom.f5Amount.value || 0); const tol = Number(dom.f5Tolerance.value || 0.01);
  let matched = [];
  if (mode === 'keyword') { if (query.length < 2) return toast('關鍵字至少2字'); matched = rows.filter((t) => t.summary.includes(query)); }
  else if (mode === 'voucher') matched = rows.filter((t) => t.voucherNo === query);
  else if (mode === 'amount') matched = rows.filter((t) => absDeltaWithin(Math.abs((t.debit || 0) - (t.credit || 0)), Math.abs(amount), tol));
  else { if (query.length < 2) return toast('複合模式關鍵字至少2字'); matched = rows.filter((t) => t.summary.includes(query) && absDeltaWithin(Math.abs((t.debit || 0) - (t.credit || 0)), Math.abs(amount), tol)); }
  if (matched.length > 500) matched = matched.slice(0, 500);
  AppState.crossLink = { mode, query, results: matched };
  const hit = new Set(matched.map((m) => m.id)); const remain = rows.filter((r) => !hit.has(r.id));
  dom.f5Result.innerHTML = matched.length ? `<p class="muted">命中 ${matched.length} 筆</p>` : `<p class="muted">命中 0 筆，仍顯示剩餘 ${remain.length} 筆。</p>`;
  renderTxnList(dom.f5Result, matched, `命中 ${matched.length} 筆`);
  renderTxnList(dom.f5List, remain, `剩餘 ${remain.length} 筆`);
}

function runF6() {
  const rows = getFilteredTransactions();
  const inputKeyword = cleanText(dom.f6Keyword.value);
  const frequency = dom.f6Frequency.value;
  const customCount = Number(dom.f6CustomCount.value || 1);
  const from = parseRocPeriod(dom.f6From.value);
  const to = parseRocPeriod(dom.f6To.value);

  const autoKeyword = inputKeyword || guessSummaryKeyword(rows);
  let base = rows;
  if (autoKeyword) base = base.filter((t) => t.summary.includes(autoKeyword));

  // Use date tokens in summary first; fallback to voucher date.
  const byPeriod = new Map();
  base.forEach((t) => {
    const keys = f6PeriodKeysFromTxn(t, frequency);
    keys.forEach((period) => {
      if ((from || to) && (t.periodROC < (from || '000-00') || t.periodROC > (to || '999-99'))) return;
      if (!byPeriod.has(period)) byPeriod.set(period, { count: 0, txnIds: [] });
      const x = byPeriod.get(period);
      x.count += 1;
      x.txnIds.push(t.id);
    });
  });

  const expected = frequency === 'custom' ? customCount : 1;
  const periods = Array.from(byPeriod.keys()).sort();
  const results = periods.map((p) => {
    const actual = byPeriod.get(p)?.count || 0;
    return { period: p, expected, actual, status: actual < expected ? 'missing' : actual > expected ? 'excess' : 'ok' };
  });

  const counts = results.map((r) => r.actual);
  const mean = counts.length ? counts.reduce((a, b) => a + b, 0) / counts.length : 0;
  const std = counts.length ? Math.sqrt(counts.reduce((a, b) => a + (b - mean) ** 2, 0) / counts.length) : 0;
  const suggestions = counts.length >= 3 && std < 0.5 ? [{ keyword: autoKeyword || '(auto)', expected }] : [];
  AppState.gap.results = results;
  AppState.gap.periodicSuggestions = suggestions;

  const okPeriods = new Set(results.filter((r) => r.status === 'ok').map((r) => r.period));
  const txMap = new Map(base.map((t) => [t.id, t]));
  const okIds = new Set();
  okPeriods.forEach((p) => (byPeriod.get(p)?.txnIds || []).forEach((id) => okIds.add(id)));
  const okRows = Array.from(okIds).map((id) => txMap.get(id)).filter(Boolean);

  dom.f6Result.innerHTML = `<div class="table-wrap"><table><thead><tr><th>期間</th><th>預期</th><th>實際</th><th>狀態</th></tr></thead><tbody>${results.map((r) => `<tr><td>${r.period}</td><td>${r.expected}</td><td>${r.actual}</td><td>${r.status === 'ok' ? 'OK' : r.status === 'missing' ? 'MISSING' : 'EXCESS'}</td></tr>`).join('')}</tbody></table></div>${
    suggestions.length ? `<p class="muted">偵測到週期性模式 ${escapeHtml(JSON.stringify(suggestions))}</p>` : ''
  }`;
  dom.f6List.innerHTML = `<div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>摘要</th><th class="col-amount">金額</th></tr></thead><tbody>${okRows.map((t) => `<tr><td>${escapeHtml(t.voucherNo)}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td></tr>`).join('')}</tbody></table></div><p class="muted">時間斷層符合頻率 ${okRows.length} 筆｜關鍵字:${escapeHtml(autoKeyword || '(未指定)')}</p>`;
}

function runF14() {
  const rows = getFilteredTransactions(); const map = new Map();
  rows.forEach((t) => { const amount = t.debit > 0 ? t.debit : t.credit; const k = `${t.accountCode}||${t.voucherNo}||${decKey(amount)}`; if (!map.has(k)) map.set(k, []); map.get(k).push(t); });
  const results = [];
  map.forEach((items) => { if (items.length >= 2 && new Set(items.map((x) => x.dateISO)).size >= 2) results.push({ id: Math.random().toString(36).slice(2, 10), voucherNo: items[0].voucherNo, accountCode: items[0].accountCode, accountName: items[0].accountName, entries: items, severity: 'WARN' }); });
  AppState.dupVoucher.results = results;
  const hit = new Set(results.flatMap((r) => r.entries.map((e) => e.id))); const remain = rows.filter((r) => !hit.has(r.id));
  dom.f14Result.innerHTML = results.length ? `<div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>科目</th><th>日期集合</th><th>筆數</th><th>嚴重度</th></tr></thead><tbody>${results.map((r) => `<tr><td>${escapeHtml(r.voucherNo)}</td><td>[${escapeHtml(r.accountCode)}] ${escapeHtml(r.accountName)}</td><td>${escapeHtml(Array.from(new Set(r.entries.map((e) => e.dateROC))).join(', '))}</td><td>${r.entries.length}</td><td class="warn">${r.severity}</td></tr>`).join('')}</tbody></table></div>` : `<p class="muted">命中 0 筆 F14，剩餘 ${remain.length} 筆仍顯示。</p>`;
  renderTxnList(dom.f14List, remain, `未命中 F14 ${remain.length} 筆`);
}

function ensureF18OtherGroup() {
  let g = AppState.trendAlert.groups.find((x) => x.name === '其他');
  if (!g) {
    g = { id: `t${Date.now()}`, name: '其他', sourceKeyword: '', transactionIds: [] };
    AppState.trendAlert.groups.push(g);
  }
  return g;
}

function renderF18Groups(rows) {
  const txMap = new Map(rows.map((t) => [t.id, t]));
  const groupOptions = AppState.trendAlert.groups.map((g) => `<option value="${g.id}">${escapeHtml(g.name)}</option>`).join('');
  const cards = AppState.trendAlert.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    const anchorId = asScopedAnchor('f18-group', g.id);
    const trend = AppState.trendAlert.results.find((r) => r.keyword === (g.sourceKeyword || g.name));
    const trendHtml = trend?.monthlyData?.length
      ? `<div class="table-wrap" style="margin:8px 0;"><table><thead><tr><th>月份</th><th class="col-amount">金額</th><th class="col-amount">前月</th><th class="col-amount">變動率(%)</th><th>狀態</th></tr></thead><tbody>${trend.monthlyData.map((m) => `<tr><td>${m.period}</td><td class="col-amount">${fmtAmount(m.amount)}</td><td class="col-amount">${m.prevAmount == null ? '—' : fmtAmount(m.prevAmount)}</td><td class="col-amount">${m.changeRate == null ? '—' : `${m.changeRate.toFixed(1)}%`}</td><td>${m.changeRate == null ? '—' : m.flagged ? '<span class="danger">超標</span>' : '<span class="ok">正常</span>'}</td></tr>`).join('')}</tbody></table></div>`
      : '<p class="muted">此關鍵字無月資料</p>';
    return `<details class="card" id="${anchorId}" style="margin:8px 0;" open>
      <summary><strong>群組：</strong><input data-f18-rename="${g.id}" value="${escapeHtml(g.name)}" style="margin-left:8px;min-width:180px;" />｜筆數 ${txns.length}｜合計 ${fmtSigned(total)} <button data-f18-del-group="${g.id}" style="margin-left:8px;">刪除群組</button></summary>
      <div style="margin-top:8px;"><button data-f18-back="1">回到分組摘要</button></div>
      ${trendHtml}
      <div style="margin-top:8px;"><label><input type="checkbox" data-f18-check-all="${g.id}" /> 全選</label> <select data-f18-batch-target="${g.id}"><option value="">批次移動到...</option>${groupOptions}</select> <button data-f18-batch-move="${g.id}">批次移動</button> <button data-f18-batch-del="${g.id}">批次刪除</button></div>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th><th>操作</th></tr></thead><tbody>
      ${txns.map((t) => `<tr><td><input type="checkbox" data-f18-pick="${t.id}" data-f18-from="${g.id}" /> ${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td><td><select data-f18-target="${t.id}"><option value="">移動到...</option>${groupOptions}</select> <button data-f18-move="${t.id}" data-f18-from="${g.id}">移動</button> <button data-f18-del="${t.id}" data-f18-from="${g.id}">刪除</button></td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');
  const copyItems = AppState.trendAlert.groups.map((g) => {
    const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
    const total = txns.reduce((a, b) => a + getSignedAmount(b), 0);
    return { text: `${g.name}(${fmtSigned(total)})`, anchorId: asScopedAnchor('f18-group', g.id) };
  });
  const copyRows = [];
  AppState.trendAlert.groups.forEach((g) => g.transactionIds.forEach((id) => { const t = txMap.get(id); if (t) copyRows.push(t); }));
  AppState.trendAlert.copyText = formatSummaryAmountList(copyRows);
  dom.f18Result.innerHTML = `<div class="muted">點選分組可跳到明細：</div><div style="margin-top:6px;">${copyItems.map((x) => `<button data-f18-jump="${x.anchorId}" style="margin:0 6px 6px 0;">${escapeHtml(x.text)}</button>`).join('')}</div>${cards}`;
  const remainRows = AppState.trendAlert.ungrouped.map((id) => txMap.get(id)).filter(Boolean);
  renderTxnList(dom.f18List, remainRows, `未分組 ${remainRows.length} 筆`);
}

function runF18() {
  const keywords = splitGroupNames(dom.f18Keyword.value);
  const threshold = Number(dom.f18Threshold.value || 20);
  if (!keywords.length) return toast('請輸入摘要關鍵字');
  const allRows = getFilteredTransactions();
  const results = [];
  const groups = [];
  const hit = new Set();

  keywords.forEach((keyword, idx) => {
    const rows = allRows.filter((t) => t.summary.includes(keyword));
    if (!rows.length) {
      groups.push({ id: `t${idx + 1}`, name: keyword, sourceKeyword: keyword, transactionIds: [] });
      results.push({ keyword, monthlyData: [] });
      return;
    }
    rows.forEach((t) => hit.add(t.id));
    const by = new Map();
    rows.forEach((t) => {
      const amount = getSignedAmount(t);
      if (!by.has(t.periodROC)) by.set(t.periodROC, { period: t.periodROC, amount: 0, txnIds: [] });
      const x = by.get(t.periodROC);
      x.amount += amount;
      x.txnIds.push(t.id);
    });
    const monthly = Array.from(by.values()).sort((a, b) => a.period.localeCompare(b.period));
    monthly.forEach((m, i) => {
      const prev = monthly[i - 1];
      m.prevAmount = prev ? prev.amount : null;
      if (!prev || prev.amount === 0) { m.changeRate = null; m.flagged = false; } else { m.changeRate = ((m.amount - prev.amount) / prev.amount) * 100; m.flagged = Math.abs(m.changeRate) > threshold; }
    });
    results.push({ keyword, monthlyData: monthly });
    groups.push({ id: `t${idx + 1}`, name: keyword, sourceKeyword: keyword, transactionIds: rows.map((r) => r.id) });
  });

  AppState.trendAlert.results = results;
  AppState.trendAlert.groups = groups;
  AppState.trendAlert.ungrouped = allRows.filter((t) => !hit.has(t.id)).map((t) => t.id);
  renderF18Groups(allRows);
}
function csvEscape(v) { const s = String(v ?? ''); return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replaceAll('"', '""')}"` : s; }
function downloadCsv(name, headers, rows) {
  const data = [headers.join(',')].concat(rows.map((r) => r.map(csvEscape).join(','))).join('\n');
  const blob = new Blob(['\ufeff' + data], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); URL.revokeObjectURL(a.href);
}

function renderTodos() {
  dom.todoBadge.textContent = String(AppState.todos.length);
  dom.todoList.innerHTML = AppState.todos.map((t) => `<li data-todo-id="${t.id}"><input data-todo-voucher="${t.id}" value="${escapeHtml(t.voucherNo)}" style="width:120px;" /> <input data-todo-content="${t.id}" value="${escapeHtml(t.content)}" style="min-width:260px;" /> <button data-todo-save="${t.id}">儲存</button> <button data-todo-del="${t.id}">刪除</button></li>`).join('');
}

function debounce(fn, wait = 180) {
  let timer = null;
  return (...args) => {
    if (timer) clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function bindEvents() {
  dom.fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    try { parseWorkbook(await file.arrayBuffer(), file.name); } catch (err) { toast(`解析失敗：${err.message}`, 'ERROR'); }
  });

  dom.resetBtn.addEventListener('click', () => location.reload());
  dom.accountSelect.addEventListener('change', () => {
    saveCurrentGroupingState();
    loadGroupingState(currentGroupingKey());
    renderBase();
  });
  dom.keywordInput.addEventListener('input', renderBase);

  dom.navButtons.forEach((btn) => btn.addEventListener('click', () => {
    dom.navButtons.forEach((b) => b.classList.toggle('active', b === btn));
    const target = btn.dataset.module;
    dom.modules.forEach((m) => m.classList.toggle('active', m.id === `module-${target}`));
  }));

  dom.runF1Btn.addEventListener('click', runF1Grouping);
  dom.applyF1Btn.addEventListener('click', () => {
    if (AppState.grouping.mode === 'draft' && AppState.grouping.draftItems.length) {
      applyF1GroupingFromDraft();
      return;
    }
    if (!AppState.grouping.groups.length && AppState.grouping.draftItems.length) {
      applyF1GroupingFromDraft();
      return;
    }
    syncF1AppliedEditsFromUI();
    applyF1GroupingEdits();
  });
  dom.runF2Btn.addEventListener('click', runF2);
  dom.runF3Btn.addEventListener('click', runF3);
  dom.runF4Btn.addEventListener('click', runF4);
  dom.runF5Btn.addEventListener('click', runF5);
  dom.runF6Btn.addEventListener('click', runF6);
  dom.runF14Btn.addEventListener('click', runF14);
  dom.runF18Btn.addEventListener('click', runF18);
  dom.copyF18TextBtn.addEventListener('click', async () => {
    const text = AppState.trendAlert.copyText || '';
    if (!text.trim()) return;
    try {
      await navigator.clipboard.writeText(text);
      toast('F18 分組摘要已複製');
    } catch {
      toast('無法存取剪貼簿，請手動複製', 'WARN');
    }
  });

  dom.f18Result.addEventListener('input', (e) => {
    const gid = e.target?.dataset?.f18Rename;
    if (!gid) return;
    const g = AppState.trendAlert.groups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(e.target.value) || g.name;
  });

  dom.f18Result.addEventListener('change', (e) => {
    const gid = e.target?.dataset?.f18Rename;
    if (gid) {
      const g = AppState.trendAlert.groups.find((x) => x.id === gid);
      if (!g) return;
      g.name = cleanText(e.target.value) || g.name;
      renderF18Groups(getFilteredTransactions());
      return;
    }
    const allGroupId = e.target?.dataset?.f18CheckAll;
    if (!allGroupId) return;
    const checked = !!e.target.checked;
    dom.f18Result.querySelectorAll(`input[data-f18-pick][data-f18-from="${allGroupId}"]`).forEach((cb) => {
      cb.checked = checked;
    });
  });

  dom.f18Result.addEventListener('click', (e) => {
    const jump = e.target?.dataset?.f18Jump;
    if (jump) {
      const target = document.getElementById(jump);
      if (!target) return;
      target.open = true;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const back = e.target?.dataset?.f18Back;
    if (back) {
      dom.f18Result.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const delGroupId = e.target?.dataset?.f18DelGroup;
    if (delGroupId) {
      const group = AppState.trendAlert.groups.find((g) => g.id === delGroupId);
      if (!group) return;
      const other = ensureF18OtherGroup();
      if (other.id !== group.id) group.transactionIds.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      AppState.trendAlert.groups = AppState.trendAlert.groups.filter((g) => g.id !== delGroupId);
      renderF18Groups(getFilteredTransactions());
      return;
    }
    const batchMoveGroupId = e.target?.dataset?.f18BatchMove;
    if (batchMoveGroupId) {
      const sel = dom.f18Result.querySelector(`select[data-f18-batch-target="${batchMoveGroupId}"]`);
      const toId = sel?.value;
      if (!toId || toId === batchMoveGroupId) return;
      const from = AppState.trendAlert.groups.find((g) => g.id === batchMoveGroupId);
      const to = AppState.trendAlert.groups.find((g) => g.id === toId);
      if (!from || !to) return;
      const picks = Array.from(dom.f18Result.querySelectorAll(`input[data-f18-pick][data-f18-from="${batchMoveGroupId}"]:checked`)).map((x) => x.getAttribute('data-f18-pick')).filter(Boolean);
      if (!picks.length) return toast('請先勾選要移動的明細', 'WARN');
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!to.transactionIds.includes(id)) to.transactionIds.push(id); });
      renderF18Groups(getFilteredTransactions());
      return;
    }
    const batchDelGroupId = e.target?.dataset?.f18BatchDel;
    if (batchDelGroupId) {
      const from = AppState.trendAlert.groups.find((g) => g.id === batchDelGroupId);
      if (!from) return;
      const picks = Array.from(dom.f18Result.querySelectorAll(`input[data-f18-pick][data-f18-from="${batchDelGroupId}"]:checked`)).map((x) => x.getAttribute('data-f18-pick')).filter(Boolean);
      if (!picks.length) return toast('請先勾選要刪除的明細', 'WARN');
      const other = ensureF18OtherGroup();
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      renderF18Groups(getFilteredTransactions());
      return;
    }
    const moveId = e.target?.dataset?.f18Move;
    if (moveId) {
      const fromId = e.target.dataset.f18From;
      const sel = dom.f18Result.querySelector(`select[data-f18-target="${moveId}"]`);
      const toId = sel?.value;
      if (!toId || toId === fromId) return;
      const from = AppState.trendAlert.groups.find((g) => g.id === fromId);
      const to = AppState.trendAlert.groups.find((g) => g.id === toId);
      if (!from || !to) return;
      from.transactionIds = from.transactionIds.filter((id) => id !== moveId);
      if (!to.transactionIds.includes(moveId)) to.transactionIds.push(moveId);
      renderF18Groups(getFilteredTransactions());
      return;
    }
    const delId = e.target?.dataset?.f18Del;
    if (delId) {
      const fromId = e.target.dataset.f18From;
      const from = AppState.trendAlert.groups.find((g) => g.id === fromId);
      if (!from) return;
      const other = ensureF18OtherGroup();
      from.transactionIds = from.transactionIds.filter((id) => id !== delId);
      if (!other.transactionIds.includes(delId)) other.transactionIds.push(delId);
      renderF18Groups(getFilteredTransactions());
    }
  });

  dom.f3Result.addEventListener('input', (e) => {
    const gid = e.target?.dataset?.f3Rename;
    if (!gid) return;
    const g = AppState.pool.groups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(e.target.value) || g.name;
  });

  dom.f3Result.addEventListener('change', (e) => {
    const gid = e.target?.dataset?.f3Rename;
    if (gid) {
      const g = AppState.pool.groups.find((x) => x.id === gid);
      if (!g) return;
      g.name = cleanText(e.target.value) || g.name;
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
      return;
    }
    const allGroupId = e.target?.dataset?.f3CheckAll;
    if (!allGroupId) return;
    const checked = !!e.target.checked;
    dom.f3Result.querySelectorAll(`input[data-f3-pick][data-f3-from="${allGroupId}"]`).forEach((cb) => {
      cb.checked = checked;
    });
  });

  dom.f3Result.addEventListener('click', async (e) => {
    const copy = e.target?.dataset?.f3Copy;
    if (copy) {
      const text = AppState.pool.copyText || '';
      if (!text.trim()) return;
      try {
        await navigator.clipboard.writeText(text);
        toast('F3 分組摘要已複製');
      } catch {
        toast('無法存取剪貼簿，請手動複製', 'WARN');
      }
      return;
    }
    const jump = e.target?.dataset?.f3Jump;
    if (jump) {
      const target = document.getElementById(jump);
      if (!target) return;
      target.open = true;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const back = e.target?.dataset?.f3Back;
    if (back) {
      dom.f3Result.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const delGroupId = e.target?.dataset?.f3DelGroup;
    if (delGroupId) {
      const group = AppState.pool.groups.find((g) => g.id === delGroupId);
      if (!group) return;
      const other = ensureF3OtherGroup();
      if (other.id !== group.id) {
        group.transactionIds.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      }
      AppState.pool.groups = AppState.pool.groups.filter((g) => g.id !== delGroupId);
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
      return;
    }
    const batchMoveGroupId = e.target?.dataset?.f3BatchMove;
    if (batchMoveGroupId) {
      const sel = dom.f3Result.querySelector(`select[data-f3-batch-target="${batchMoveGroupId}"]`);
      const toId = sel?.value;
      if (!toId || toId === batchMoveGroupId) return;
      const from = AppState.pool.groups.find((g) => g.id === batchMoveGroupId);
      const to = AppState.pool.groups.find((g) => g.id === toId);
      if (!from || !to) return;
      const picks = Array.from(dom.f3Result.querySelectorAll(`input[data-f3-pick][data-f3-from="${batchMoveGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f3-pick'))
        .filter(Boolean);
      if (!picks.length) return toast('請先勾選要移動的明細', 'WARN');
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!to.transactionIds.includes(id)) to.transactionIds.push(id); });
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
      return;
    }
    const batchDelGroupId = e.target?.dataset?.f3BatchDel;
    if (batchDelGroupId) {
      const from = AppState.pool.groups.find((g) => g.id === batchDelGroupId);
      if (!from) return;
      const picks = Array.from(dom.f3Result.querySelectorAll(`input[data-f3-pick][data-f3-from="${batchDelGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f3-pick'))
        .filter(Boolean);
      if (!picks.length) return toast('請先勾選要刪除的明細', 'WARN');
      const other = ensureF3OtherGroup();
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
      return;
    }
    const moveId = e.target?.dataset?.f3Move;
    if (moveId) {
      const fromId = e.target.dataset.f3From;
      const sel = dom.f3Result.querySelector(`select[data-f3-target="${moveId}"]`);
      const toId = sel?.value;
      if (!toId || toId === fromId) return;
      const from = AppState.pool.groups.find((g) => g.id === fromId);
      const to = AppState.pool.groups.find((g) => g.id === toId);
      if (!from || !to) return;
      from.transactionIds = from.transactionIds.filter((id) => id !== moveId);
      if (!to.transactionIds.includes(moveId)) to.transactionIds.push(moveId);
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
      return;
    }
    const delId = e.target?.dataset?.f3Del;
    if (delId) {
      const fromId = e.target.dataset.f3From;
      const from = AppState.pool.groups.find((g) => g.id === fromId);
      if (!from) return;
      const other = ensureF3OtherGroup();
      from.transactionIds = from.transactionIds.filter((id) => id !== delId);
      if (!other.transactionIds.includes(delId)) other.transactionIds.push(delId);
      renderF3Groups(getFilteredTransactions(), `結果 ${AppState.pool.groups.length} 組`);
    }
  });

  dom.f2Result.addEventListener('click', (e) => {
    const idxText = e.target?.dataset?.f2Unpair;
    if (idxText == null) return;
    const idx = Number(idxText);
    const pair = AppState.offset.pairs[idx];
    if (!pair) return;

    // autoPairs (legacy shape)
    if (pair?.debit?.id && pair?.credit?.id) {
      if (!AppState.offset.forcedUnmatchedIds.includes(pair.debit.id)) AppState.offset.forcedUnmatchedIds.push(pair.debit.id);
      if (!AppState.offset.forcedUnmatchedIds.includes(pair.credit.id)) AppState.offset.forcedUnmatchedIds.push(pair.credit.id);
      AppState.offset.manualPairIds = AppState.offset.manualPairIds.filter((p) => !(p.debitId === pair.debit.id && p.creditId === pair.credit.id));
      runF2();
      return;
    }

    // manualMatches (multi-line)
    const ids = (pair.debitIds || []).concat(pair.creditIds || []).filter(Boolean);
    ids.forEach((id2) => {
      if (!AppState.offset.forcedUnmatchedIds.includes(id2)) AppState.offset.forcedUnmatchedIds.push(id2);
    });

    // remove from manualMatches by id
    AppState.offset.manualMatches = (AppState.offset.manualMatches || []).filter((m) => m.id !== pair.id);

    runF2();
  });

  dom.f2UnmatchedSummary.addEventListener('input', (e) => {
    const gid = e.target?.dataset?.f2Rename;
    if (!gid) return;
    const g = AppState.offset.unmatchedGroups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(e.target.value) || g.name;
  });

  dom.f2UnmatchedSummary.addEventListener('click', (e) => {
    const view = e.target?.dataset?.f2View;
    if (view) {
      AppState.offset.unmatchedView = view === 'group' ? 'group' : 'review';
      renderF2UnmatchedEditor(getF2UnmatchedRowsFromState());
      return;
    }

    const suggest = e.target?.dataset?.f2Suggest;
    if (suggest) {
      const [debitCsv, creditCsv] = String(suggest).split('|');
      const debitIds = (debitCsv || '').split(',').map((x) => cleanText(x)).filter(Boolean);
      const creditIds = (creditCsv || '').split(',').map((x) => cleanText(x)).filter(Boolean);
      if (!debitIds.length || !creditIds.length) return;
      // 先清除現有勾選
      dom.f2List.querySelectorAll('input[data-f2-pick]').forEach((cb) => { cb.checked = false; });
      // 自動勾選建議的多筆
      debitIds.concat(creditIds).forEach((id) => {
        const cb = dom.f2List.querySelector(`input[data-f2-pick="${id}"]`);
        if (cb) cb.checked = true;
      });
      const row = document.getElementById(`f2-row-${debitIds[0]}`) || document.getElementById(`f2-row-${creditIds[0]}`);
      if (row) row.scrollIntoView({ behavior: 'smooth', block: 'center' });
      toast(`已自動勾選建議沖帳：借${debitIds.length}／貸${creditIds.length}`);
      return;
    }
    const refresh = e.target?.dataset?.f2SuggestRefresh;
    if (refresh) {
      runF2();
      return;
    }

    const jump = e.target?.dataset?.f2Jump;
    if (jump) {
      const target = document.getElementById(jump);
      if (!target) return;
      target.open = true;
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const back = e.target?.dataset?.f2Back;
    if (back) {
      dom.f2UnmatchedSummary.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    const delGroupId = e.target?.dataset?.f2DelGroup;
    if (delGroupId) {
      const groups = AppState.offset.unmatchedGroups;
      const target = groups.find((g) => g.id === delGroupId);
      if (!target) return;
      let other = groups.find((g) => g.name === '其他');
      if (!other) {
        other = { id: `u${Date.now()}`, name: '其他', transactionIds: [] };
        groups.push(other);
      }
      if (other.id !== target.id) {
        target.transactionIds.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      }
      AppState.offset.unmatchedGroups = groups.filter((g) => g.id !== delGroupId);
      renderF2UnmatchedOrganizer(getF2UnmatchedRowsFromState());
      return;
    }

    const batchMoveGroupId = e.target?.dataset?.f2BatchMove;
    const batchDelGroupId = e.target?.dataset?.f2BatchDel;
    if (batchMoveGroupId) {
      const sel = dom.f2UnmatchedSummary.querySelector(`select[data-f2-batch-target="${batchMoveGroupId}"]`);
      const toId = sel?.value;
      if (!toId || toId === batchMoveGroupId) return;
      const from = AppState.offset.unmatchedGroups.find((g) => g.id === batchMoveGroupId);
      const to = AppState.offset.unmatchedGroups.find((g) => g.id === toId);
      if (!from || !to) return;
      const picks = Array.from(dom.f2UnmatchedSummary.querySelectorAll(`input[data-f2-pick-item][data-f2-from="${batchMoveGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f2-pick-item'))
        .filter(Boolean);
      if (!picks.length) {
        toast('請先勾選要移動的明細', 'WARN');
        return;
      }
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!to.transactionIds.includes(id)) to.transactionIds.push(id); });
      renderF2UnmatchedOrganizer(getF2UnmatchedRowsFromState());
      toast(`已批次移動 ${picks.length} 筆`);
      return;
    }
    if (batchDelGroupId) {
      const from = AppState.offset.unmatchedGroups.find((g) => g.id === batchDelGroupId);
      if (!from) return;
      const picks = Array.from(dom.f2UnmatchedSummary.querySelectorAll(`input[data-f2-pick-item][data-f2-from="${batchDelGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f2-pick-item'))
        .filter(Boolean);
      if (!picks.length) {
        toast('請先勾選要刪除的明細', 'WARN');
        return;
      }
      let other = AppState.offset.unmatchedGroups.find((g) => g.name === '其他');
      if (!other) {
        other = { id: `u${Date.now()}`, name: '其他', transactionIds: [] };
        AppState.offset.unmatchedGroups.push(other);
      }
      from.transactionIds = from.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      renderF2UnmatchedOrganizer(getF2UnmatchedRowsFromState());
      toast(`已批次刪除 ${picks.length} 筆（移至其他）`);
      return;
    }

    const moveId = e.target?.dataset?.f2Move;
    const delId = e.target?.dataset?.f2Del;
    if (moveId) {
      const fromId = e.target.dataset.f2From;
      const sel = dom.f2UnmatchedSummary.querySelector(`select[data-f2-target="${moveId}"]`);
      const toId = sel?.value;
      if (!toId || toId === fromId) return;
      const from = AppState.offset.unmatchedGroups.find((g) => g.id === fromId);
      const to = AppState.offset.unmatchedGroups.find((g) => g.id === toId);
      if (!from || !to) return;
      from.transactionIds = from.transactionIds.filter((id) => id !== moveId);
      if (!to.transactionIds.includes(moveId)) to.transactionIds.push(moveId);
      renderF2UnmatchedOrganizer(getF2UnmatchedRowsFromState());
      return;
    }
    if (delId) {
      const fromId = e.target.dataset.f2From;
      const from = AppState.offset.unmatchedGroups.find((g) => g.id === fromId);
      if (!from) return;
      let other = AppState.offset.unmatchedGroups.find((g) => g.name === '其他');
      if (!other) {
        other = { id: `u${Date.now()}`, name: '其他', transactionIds: [] };
        AppState.offset.unmatchedGroups.push(other);
      }
      from.transactionIds = from.transactionIds.filter((id) => id !== delId);
      if (!other.transactionIds.includes(delId)) other.transactionIds.push(delId);
      renderF2UnmatchedOrganizer(getF2UnmatchedRowsFromState());
    }
  });

  dom.f2UnmatchedSummary.addEventListener('change', (e) => {
    const th = e.target?.dataset?.f2SuggestTh;
    if (th) {
      const n = Number(e.target.value);
      AppState.offset.suggestThreshold = Number.isFinite(n) ? Math.max(0, Math.min(100, n)) : 80;
      return;
    }
    const win = e.target?.dataset?.f2Win;
    if (win) {
      const n = Number(e.target.value);
      AppState.offset.timeWindowDays = Number.isFinite(n) ? Math.max(0, Math.min(365, n)) : 14;
      return;
    }
    const km = e.target?.dataset?.f2Kmax;
    if (km) {
      const n = Number(e.target.value);
      AppState.offset.subsetMaxK = Number.isFinite(n) ? Math.max(2, Math.min(8, n)) : 4;
    }

    const allGroupId = e.target?.dataset?.f2CheckAll;
    if (!allGroupId) return;
    const checked = !!e.target.checked;
    dom.f2UnmatchedSummary.querySelectorAll(`input[data-f2-pick-item][data-f2-from="${allGroupId}"]`).forEach((cb) => {
      cb.checked = checked;
    });
  });

  dom.f2ManualMatchBtn.addEventListener('click', () => {
    const picks = Array.from(dom.f2List.querySelectorAll('input[data-f2-pick]:checked'))
      .map((x) => x.getAttribute('data-f2-pick'))
      .filter(Boolean);

    if (picks.length < 2) {
      toast('請至少勾選 2 筆（可多對一/一對多）', 'WARN');
      return;
    }

    const all = getFilteredTransactions();
    const rows = picks.map((id) => all.find((t) => t.id === id)).filter(Boolean);
    if (rows.length !== picks.length) {
      toast('部分勾選項目已不在清單中', 'WARN');
      return;
    }

    const debitIds = rows.filter((t) => t.debit > 0).map((t) => t.id);
    const creditIds = rows.filter((t) => t.credit > 0).map((t) => t.id);
    if (!debitIds.length || !creditIds.length) {
      toast('需同時包含借方與貸方項目', 'WARN');
      return;
    }

    const txMap = new Map(all.map((t) => [t.id, t]));
    const debitTotal = sumByIds(txMap, debitIds, 'debit');
    const creditTotal = sumByIds(txMap, creditIds, 'credit');
    const tol = Number(dom.f2Tolerance.value || 0.01);
    if (!absDeltaWithin(debitTotal, creditTotal, tol)) {
      toast(`借貸不平（借${fmtAmount(debitTotal)} / 貸${fmtAmount(creditTotal)}，容差${tol}）`, 'WARN');
      return;
    }

    const id = `m${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
    const simPct = bestSummarySimilarityPct(txMap, debitIds, creditIds);
    const dayDiff = (() => {
      const d0 = txMap.get(debitIds[0])?.date;
      const ds = debitIds.concat(creditIds).map((x) => txMap.get(x)?.date).filter((x) => x instanceof Date);
      if (!d0 || !ds.length) return null;
      return Math.max(...ds.map((dt) => daysBetween(dt, d0)).filter((x) => x != null));
    })();
    const voucherMatch = (() => {
      const v = cleanText(txMap.get(debitIds[0])?.voucherNo);
      if (!v) return false;
      return creditIds.some((cid) => cleanText(txMap.get(cid)?.voucherNo) === v);
    })();

    const exists = (AppState.offset.manualMatches || []).some((m) => {
      const a = (m.debitIds || []).slice().sort().join(',');
      const b = (m.creditIds || []).slice().sort().join(',');
      return a === debitIds.slice().sort().join(',') && b === creditIds.slice().sort().join(',');
    });
    if (exists) {
      toast('此沖帳組合已存在', 'WARN');
      return;
    }

    if (!AppState.offset.manualMatches) AppState.offset.manualMatches = [];
    AppState.offset.manualMatches.push({
      id,
      debitIds,
      creditIds,
      createdAt: Date.now(),
      reason: { kind: `${debitIds.length}↔${creditIds.length}`, delta: debitTotal - creditTotal, simPct, voucherMatch, dayDiff },
    });

    // 被套用的 items 不應再卡在 forcedUnmatched
    const usedIds = new Set(debitIds.concat(creditIds));
    AppState.offset.forcedUnmatchedIds = AppState.offset.forcedUnmatchedIds.filter((x) => !usedIds.has(x));

    runF2();
  });

  dom.f2ResetManualBtn.addEventListener('click', () => {
    AppState.offset.forcedUnmatchedIds = [];
    AppState.offset.manualPairIds = [];
    AppState.offset.manualMatches = [];
    AppState.offset.unmatchedGroups = [];
    runF2();
  });

  dom.copyF2TextBtn.addEventListener('click', async () => {
    const text = AppState.offset.copyText || '';
    if (!text.trim()) return;
    try {
      await navigator.clipboard.writeText(text);
      toast('F2 未沖帳整理已複製');
    } catch {
      toast('無法存取剪貼簿，請手動複製', 'WARN');
    }
  });

  dom.f1Result.addEventListener('input', (e) => {
    const draftFrom = e.target?.dataset?.f1DraftName;
    if (draftFrom) {
      AppState.grouping.draftRenameMap[draftFrom] = cleanText(e.target.value) || draftFrom;
      return;
    }
    const gid = e.target?.dataset?.f1Rename;
    if (!gid) return;
    const g = AppState.grouping.groups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(e.target.value) || g.name;
  });

  dom.f1Result.addEventListener('change', (e) => {
    const draftFrom = e.target?.dataset?.f1DraftName;
    if (draftFrom) {
      const newName = cleanText(e.target.value) || draftFrom;
      if (newName !== draftFrom) {
        AppState.grouping.draftItems.forEach((d) => {
          if (d.proposed === draftFrom) d.proposed = newName;
        });
        AppState.grouping.draftExtraGroups = AppState.grouping.draftExtraGroups.map((x) => (x === draftFrom ? newName : x));
        delete AppState.grouping.draftRenameMap[draftFrom];
        AppState.grouping.draftRenameMap[newName] = newName;
        renderF1Draft();
      }
      return;
    }
    const gid = e.target?.dataset?.f1Rename;
    if (!gid) return;
    const g = AppState.grouping.groups.find((x) => x.id === gid);
    if (!g) return;
    g.name = cleanText(e.target.value) || g.name;
    renderF1Output();
  });

  dom.f1Result.addEventListener('click', (e) => {
    const draftDel = e.target?.dataset?.f1DraftDel;
    if (draftDel) {
      // Remove preview name from current candidate set, but keep source for future restore.
      AppState.grouping.draftItems.forEach((d) => {
        if (d.proposed === draftDel) d.proposed = '其他';
      });
      AppState.grouping.draftExtraGroups = AppState.grouping.draftExtraGroups.filter((x) => x !== draftDel && x !== '其他');
      delete AppState.grouping.draftRenameMap[draftDel];
      AppState.grouping.mode = 'draft';
      renderF1Draft();
      return;
    }
    const back = e.target?.dataset?.f1Back;
    if (back) {
      dom.f1CopyOutput.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }

    // F1: 本組關鍵字/相似度 勾選輔助
    const ruleSelectId = e.target?.dataset?.f1RuleSelect;
    if (ruleSelectId) {
      const group = AppState.grouping.groups.find((g) => g.id === ruleSelectId);
      if (!group) return;
      if (!group.rule) group.rule = { mode: 'A', keyword: '', threshold: 70 };
      const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
      const hits = new Set();
      group.transactionIds.forEach((id) => {
        const t = txMap.get(id);
        if (t && f1RuleMatch(t, group.rule)) hits.add(id);
      });
      dom.f1Result.querySelectorAll(`input[data-f1-pick][data-f1-from="${ruleSelectId}"]`).forEach((cb) => {
        cb.checked = hits.has(cb.getAttribute('data-f1-pick'));
      });
      toast(`已勾選命中 ${hits.size} 筆`);
      return;
    }
    const ruleClearId = e.target?.dataset?.f1RuleClear;
    if (ruleClearId) {
      dom.f1Result.querySelectorAll(`input[data-f1-pick][data-f1-from="${ruleClearId}"]`).forEach((cb) => { cb.checked = false; });
      toast('已清除勾選');
      return;
    }

    const delGroupId = e.target?.dataset?.f1DelGroup;
    if (delGroupId) {
      const group = AppState.grouping.groups.find((g) => g.id === delGroupId);
      if (!group) return;
      const other = ensureOtherGroup();
      if (other.id !== group.id) {
        group.transactionIds.forEach((id) => { if (!other.transactionIds.includes(id)) other.transactionIds.push(id); });
      }
      AppState.grouping.groups = AppState.grouping.groups.filter((g) => g.id !== delGroupId);
      renderF1Output();
      return;
    }

    const moveTxnId = e.target?.dataset?.f1Move;
    const delTxnId = e.target?.dataset?.f1Del;
    const batchMoveGroupId = e.target?.dataset?.f1BatchMove;
    const batchDelGroupId = e.target?.dataset?.f1BatchDel;
    if (batchMoveGroupId) {
      const sel = dom.f1Result.querySelector(`select[data-f1-batch-target="${batchMoveGroupId}"]`);
      const toId = sel?.value;
      if (!toId || toId === batchMoveGroupId) return;
      const fromGroup = AppState.grouping.groups.find((g) => g.id === batchMoveGroupId);
      const toGroup = AppState.grouping.groups.find((g) => g.id === toId);
      if (!fromGroup || !toGroup) return;
      const picks = Array.from(dom.f1Result.querySelectorAll(`input[data-f1-pick][data-f1-from="${batchMoveGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f1-pick'))
        .filter(Boolean);
      if (!picks.length) {
        toast('請先勾選要移動的明細', 'WARN');
        return;
      }
      fromGroup.transactionIds = fromGroup.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!toGroup.transactionIds.includes(id)) toGroup.transactionIds.push(id); });
      renderF1Output();
      toast(`已批次移動 ${picks.length} 筆`);
      return;
    }
    if (batchDelGroupId) {
      const fromGroup = AppState.grouping.groups.find((g) => g.id === batchDelGroupId);
      if (!fromGroup) return;
      const picks = Array.from(dom.f1Result.querySelectorAll(`input[data-f1-pick][data-f1-from="${batchDelGroupId}"]:checked`))
        .map((x) => x.getAttribute('data-f1-pick'))
        .filter(Boolean);
      if (!picks.length) {
        toast('請先勾選要刪除的明細', 'WARN');
        return;
      }
      fromGroup.transactionIds = fromGroup.transactionIds.filter((id) => !picks.includes(id));
      picks.forEach((id) => { if (!AppState.grouping.ungrouped.includes(id)) AppState.grouping.ungrouped.push(id); });
      renderF1Output();
      toast(`已批次刪除 ${picks.length} 筆`);
      return;
    }
    if (moveTxnId) {
      const fromId = e.target.dataset.f1From;
      const sel = dom.f1Result.querySelector(`select[data-f1-target="${moveTxnId}"]`);
      const toId = sel?.value;
      if (!toId || toId === fromId) return;
      const fromGroup = AppState.grouping.groups.find((g) => g.id === fromId);
      const toGroup = AppState.grouping.groups.find((g) => g.id === toId);
      if (!fromGroup || !toGroup) return;
      fromGroup.transactionIds = fromGroup.transactionIds.filter((id) => id !== moveTxnId);
      if (!toGroup.transactionIds.includes(moveTxnId)) toGroup.transactionIds.push(moveTxnId);
      renderF1Output();
      return;
    }
    if (delTxnId) {
      const fromId = e.target.dataset.f1From;
      const fromGroup = AppState.grouping.groups.find((g) => g.id === fromId);
      if (!fromGroup) return;
      fromGroup.transactionIds = fromGroup.transactionIds.filter((id) => id !== delTxnId);
      if (!AppState.grouping.ungrouped.includes(delTxnId)) AppState.grouping.ungrouped.push(delTxnId);
      renderF1Output();
    }
  });

  // Draft rename persistence: 當你在「預覽分組名稱」直接改名字，任何重繪都要保留。
  dom.f1Result.addEventListener('input', (e) => {
    const from = e.target?.dataset?.f1DraftName;
    if (!from) return;
    const to = cleanText(e.target.value) || from || '其他';
    AppState.grouping.draftRenameMap[from] = to;
  });

  dom.f1Result.addEventListener('change', (e) => {
    const allGroupId = e.target?.dataset?.f1CheckAll;
    if (allGroupId) {
      const checked = !!e.target.checked;
      dom.f1Result.querySelectorAll(`input[data-f1-pick][data-f1-from="${allGroupId}"]`).forEach((cb) => {
        cb.checked = checked;
      });
      return;
    }

    const modeGid = e.target?.dataset?.f1RuleMode;
    if (modeGid) {
      const g = AppState.grouping.groups.find((x) => x.id === modeGid);
      if (!g) return;
      if (!g.rule) g.rule = { mode: 'A', keyword: '', threshold: 70 };
      g.rule.mode = e.target.value === 'B' ? 'B' : 'A';
      return;
    }

    const thGid = e.target?.dataset?.f1RuleThreshold;
    if (thGid) {
      const g = AppState.grouping.groups.find((x) => x.id === thGid);
      if (!g) return;
      if (!g.rule) g.rule = { mode: 'A', keyword: '', threshold: 70 };
      const n = Number(e.target.value);
      g.rule.threshold = Number.isFinite(n) ? Math.max(0, Math.min(100, n)) : 70;
    }
  });

  dom.f1Result.addEventListener('input', (e) => {
    const gid = e.target?.dataset?.f1RuleKeyword;
    if (!gid) return;
    const g = AppState.grouping.groups.find((x) => x.id === gid);
    if (!g) return;
    if (!g.rule) g.rule = { mode: 'A', keyword: '', threshold: 70 };
    g.rule.keyword = String(e.target.value || '');
  });

  dom.f1CopyOutput.addEventListener('click', (e) => {
    const anchorId = e.target?.dataset?.f1Jump;
    if (!anchorId) return;
    const target = document.getElementById(anchorId);
    if (!target) return;
    target.open = true;
    target.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  dom.f1AddGroupBtn.addEventListener('click', () => {
    // 先同步「預覽分組名稱」裡手動改過的名稱，避免 renderF1Draft() 重新渲染時被洗掉。
    if (AppState.grouping.mode === 'draft') syncF1DraftEditsFromUI();

    const names = splitGroupNames(dom.f1NewGroupName.value);
    if (!names.length) return;
    if (AppState.grouping.mode === 'draft') {
      const existing = new Set(AppState.grouping.draftItems.map((d) => d.proposed).concat(AppState.grouping.draftExtraGroups));
      let added = 0;
      names.forEach((name) => {
        // If deleted before, restore rows by source keyword first.
        let restored = 0;
        AppState.grouping.draftItems.forEach((d) => {
          if ((d.source || d.proposed) === name && d.proposed !== name) {
            d.proposed = name;
            restored += 1;
          }
        });
        if (restored > 0) {
          existing.add(name);
          added += 1;
          return;
        }
        if (existing.has(name)) return;
        existing.add(name);
        AppState.grouping.draftExtraGroups.push(name);
        added += 1;
      });
      if (!added) {
        toast('分組名稱已存在', 'WARN');
        return;
      }
      dom.f1NewGroupName.value = '';
      AppState.grouping.mode = 'draft';
      renderF1Draft();
      toast(`已新增 ${added} 個候選名稱`);
      return;
    }
    const existing = new Set(AppState.grouping.groups.map((g) => g.name));
    let added = 0;
    names.forEach((name, idx) => {
      if (existing.has(name)) return;
      existing.add(name);
      AppState.grouping.groups.push({ id: `g${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 6)}`, name, transactionIds: [], rule: { mode: 'A', keyword: '', threshold: 70 } });
      added += 1;
    });
    if (!added) {
      toast('分組名稱已存在', 'WARN');
      return;
    }
    dom.f1NewGroupName.value = '';
    renderF1Output();
    toast(`已新增 ${added} 個分組`);
  });

  dom.copyF1TextBtn.addEventListener('click', async () => {
    const text = AppState.grouping.copyText || '';
    if (!text.trim()) return;
    try {
      await navigator.clipboard.writeText(text);
      toast('F1 分組摘要已複製');
    } catch {
      toast('無法存取剪貼簿，請手動複製', 'WARN');
    }
  });

  dom.todoToggleBtn.addEventListener('click', () => {
    dom.todoPanel.classList.toggle('collapsed');
    dom.todoToggleBtn.textContent = dom.todoPanel.classList.contains('collapsed') ? '展開' : '隱藏';
  });

  const autoF2 = debounce(() => { if (AppState.transactions.length) runF2(); });
  const autoF3 = debounce(() => { if (AppState.transactions.length) runF3(); }, 260);
  const autoF5 = debounce(() => { if (AppState.transactions.length) runF5(); });
  const autoF6 = debounce(() => { if (AppState.transactions.length) runF6(); });
  const autoF18 = debounce(() => { if (AppState.transactions.length) runF18(); });

  dom.f2Tolerance.addEventListener('input', autoF2);
  dom.f3Direction.addEventListener('change', autoF3);
  dom.f3Target.addEventListener('input', autoF3);
  dom.f3Tolerance.addEventListener('input', autoF3);
  dom.f5Mode.addEventListener('change', autoF5);
  dom.f5Query.addEventListener('input', autoF5);
  dom.f5Amount.addEventListener('input', autoF5);
  dom.f5Tolerance.addEventListener('input', autoF5);
  dom.f6Keyword.addEventListener('input', autoF6);
  dom.f6Frequency.addEventListener('change', autoF6);
  dom.f6CustomCount.addEventListener('input', autoF6);
  dom.f6From.addEventListener('input', autoF6);
  dom.f6To.addEventListener('input', autoF6);
  dom.f18Keyword.addEventListener('input', autoF18);
  dom.f18Threshold.addEventListener('input', autoF18);

  dom.exportF1Btn.addEventListener('click', () => {
    const map = new Map(AppState.transactions.map((t) => [t.id, t])); const rows = [];
    AppState.grouping.groups.forEach((g) => {
      const txns = g.transactionIds.map((id) => map.get(id)).filter(Boolean);
      txns.forEach((t) => rows.push([g.name, t.voucherNo, t.dateROC, t.summary, t.debit, t.credit]));
      rows.push([`${g.name}_合計`, '', '', '', txns.reduce((a, b) => a + b.debit, 0), txns.reduce((a, b) => a + b.credit, 0)]);
    });
    downloadCsv(`F1_${Date.now()}.csv`, ['組名', '傳票號碼', '日期', '摘要', '借方金額', '貸方金額'], rows);
  });

  dom.exportF2Btn.addEventListener('click', () => {
    const rows = AppState.offset.pairs.map((p) => [p.debit.summary, p.debitTotal, p.credit.summary, p.creditTotal, p.confidence]);
    downloadCsv(`F2_${Date.now()}.csv`, ['借方摘要', '借方金額', '貸方摘要', '貸方金額', '信心度'], rows);
  });

  dom.exportF4Btn.addEventListener('click', () => {
    downloadCsv(`F4_${Date.now()}.csv`, ['規則', '嚴重度', '科目代碼', '命中分錄', '說明'], AppState.anomaly.results.map((r) => [r.rule, r.severity, r.accountCode, r.transactionIds.join('|'), r.description]));
  });

  dom.exportF14Btn.addEventListener('click', () => {
    const rows = [];
    AppState.dupVoucher.results.forEach((r) => {
      const s = [...r.entries].sort((a, b) => a.dateISO.localeCompare(b.dateISO));
      if (s.length >= 2) rows.push([r.voucherNo, r.accountCode, r.accountName, s[0].dateROC, s[0].summary, s[0].debit || s[0].credit, s[1].dateROC, s[1].summary, s[1].debit || s[1].credit, r.severity]);
    });
    downloadCsv(`F14_${Date.now()}.csv`, ['傳票號碼', '科目代碼', '科目名稱', '日期1', '摘要1', '金額1', '日期2', '摘要2', '金額2', '狀態'], rows);
  });

  dom.exportF18Btn.addEventListener('click', () => {
    const rows = [];
    AppState.trendAlert.results.forEach((hit) => {
      (hit.monthlyData || []).forEach((m) => rows.push([hit.keyword, m.period, m.amount, m.prevAmount == null ? '' : m.prevAmount, m.prevAmount == null ? '' : (m.amount - m.prevAmount), m.changeRate == null ? '—' : m.changeRate.toFixed(2), m.flagged ? 'Y' : 'N', m.txnIds.length]));
    });
    downloadCsv(`F18_${Date.now()}.csv`, ['關鍵字', '月份', '金額', '前月金額', '變動金額', '變動率(%)', '是否超出門檻', '涉及傳票數'], rows);
  });

  dom.addTodoBtn.addEventListener('click', () => {
    const voucherNo = cleanText(dom.todoVoucher.value); const content = cleanText(dom.todoContent.value);
    if (!voucherNo) return toast('傳票號碼為必填', 'ERROR');
    if (!content) return toast('請填寫待辦內容', 'ERROR');
    AppState.todos.push({ id: Math.random().toString(36).slice(2, 10), voucherNo, content });
    dom.todoVoucher.value = ''; dom.todoContent.value = ''; renderTodos();
  });
  dom.todoList.addEventListener('click', (e) => {
    const saveId = e.target?.dataset?.todoSave;
    if (saveId) {
      const item = AppState.todos.find((t) => t.id === saveId);
      if (!item) return;
      const voucherInput = dom.todoList.querySelector(`input[data-todo-voucher="${saveId}"]`);
      const contentInput = dom.todoList.querySelector(`input[data-todo-content="${saveId}"]`);
      const voucherNo = cleanText(voucherInput?.value);
      const content = cleanText(contentInput?.value);
      if (!voucherNo || !content) return toast('傳票與內容不可空白', 'WARN');
      item.voucherNo = voucherNo;
      item.content = content;
      renderTodos();
      toast('已更新待辦');
      return;
    }
    const delId = e.target?.dataset?.todoDel;
    if (delId) {
      AppState.todos = AppState.todos.filter((t) => t.id !== delId);
      renderTodos();
    }
  });
  dom.copyTodoAllBtn.addEventListener('click', async () => {
    const text = AppState.todos.map((t) => `${t.voucherNo} ${t.content}`).join('\n');
    if (!text.trim()) return;
    try {
      await navigator.clipboard.writeText(text);
      toast('待辦已全部複製');
    } catch {
      toast('無法存取剪貼簿，請手動複製', 'WARN');
    }
  });
}

function init() {
  bindEvents();
  renderTodos();
  toast(`LibDetector: Fuse ${hasFuse ? 'OK' : 'fallback'} / Decimal ${hasDecimal ? 'OK' : 'missing'}`);
}

init();
