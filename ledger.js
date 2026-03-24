const DEFAULT_COLUMN_MAP = { date: 0, voucherNo: 4, summary: 10, debit: 14, credit: 17, drCr: 19, balance: 21 };

const AppState = {
  meta: { company: '', parsedAt: null, skippedRows: 0, igCount: { header: 0, monthly: 0, opening: 0, accountHeader: 0, crossPageDup: 0, invalid: 0, other: 0 }, columnMap: { ...DEFAULT_COLUMN_MAP }, excludedRows: [] },
  accounts: {}, transactions: [],
  grouping: { groups: [], ungrouped: [], draftItems: [], mode: 'applied', draftRenameMap: {}, draftExtraGroups: [], copyText: '', keywordRules: [] },
  groupingStore: {}, activeGroupingKey: 'all',
  offset: { pairs: [], forcedUnmatchedIds: [], manualPairIds: [], manualMatches: [], lastUnmatchedIds: [], unmatchedGroups: [], copyText: '', suggestThreshold: 80, timeWindowDays: 14, subsetMaxK: 4, subsetTimeLimitMs: 1200, unmatchedView: 'review' },
  pool: { candidateIds: [], results: [], groups: [], ungrouped: [], copyText: '' }, anomaly: { results: [], reviewed: new Set() }, crossLink: { mode: 'keyword', query: '', results: [] },
  gap: { results: [], periodicSuggestions: [] }, dupVoucher: { results: [] }, trendAlert: { results: [], groups: [], ungrouped: [], copyText: '' }, todos: [],
  memo: { notes: {}, rules: [], missingResults: [] },
};

const hasFuse = typeof Fuse !== 'undefined';
const hasDecimal = typeof Decimal !== 'undefined';
let searchEngine = null;
let poolWorker = null;
// Overview sort & pagination state
let _ovSort = { col: 'dateROC', dir: 'asc' };
let _ovPage = 1;
const OV_PAGE_SIZE = 150;
// ---- Company prefix cross-file comparison ----
function getCompanyKey(fileName) {
  return String(fileName || '').replace(/\.[^.]+$/, '').slice(0, 8);
}
function loadPrevCompanyData() {
  try {
    return { key: localStorage.getItem('ledger_company_key') || '', summaries: JSON.parse(localStorage.getItem('ledger_prev_summaries') || '{}') };
  } catch { return { key: '', summaries: {} }; }
}
function saveCompanyData(fileName, transactions) {
  try {
    const key = getCompanyKey(fileName);
    const summaries = {};
    for (const t of transactions) {
      if (!summaries[t.accountCode]) summaries[t.accountCode] = [];
      const s = cleanText(t.summary || '');
      if (s && !summaries[t.accountCode].includes(s)) summaries[t.accountCode].push(s);
    }
    localStorage.setItem('ledger_company_key', key);
    localStorage.setItem('ledger_prev_summaries', JSON.stringify(summaries));
  } catch { /* ignore */ }
}
function renderCrossFileResult(fileName, transactions) {
  const el = dom.crossFileResult;
  if (!el) return;
  const newKey = getCompanyKey(fileName);
  const prev = loadPrevCompanyData();
  if (!prev.key || !Object.keys(prev.summaries).length) { el.innerHTML = ''; return; }
  if (prev.key !== newKey) {
    el.innerHTML = `<div class="card" style="border-color:var(--warn);background:#fff8f0;margin-bottom:10px;"><strong style="color:var(--warn);">⚠️ 偵測到不同公司</strong><div class="muted" style="margin-top:4px;">前次公司代碼：<code>${escapeHtml(prev.key)}</code>，本次：<code>${escapeHtml(newKey)}</code>。備忘與規則已保留，跨期比對不適用。</div></div>`;
    return;
  }
  const newSummaries = {};
  for (const t of transactions) {
    if (!newSummaries[t.accountCode]) newSummaries[t.accountCode] = new Set();
    const s = cleanText(t.summary || ''); if (s) newSummaries[t.accountCode].add(s);
  }
  const missing = [], added = [];
  const allAccts = new Set([...Object.keys(prev.summaries), ...Object.keys(newSummaries)]);
  for (const code of allAccts) {
    const prevSet = new Set(prev.summaries[code] || []);
    const newSet = newSummaries[code] || new Set();
    const gone = [...prevSet].filter((s) => !newSet.has(s));
    const appeared = [...newSet].filter((s) => !prevSet.has(s));
    if (gone.length) missing.push({ code, summaries: gone });
    if (appeared.length) added.push({ code, summaries: appeared });
  }
  if (!missing.length && !added.length) {
    el.innerHTML = `<div class="card" style="border-color:var(--ok);background:#f6ffed;margin-bottom:10px;"><strong style="color:var(--ok);">✓ 跨期比對：摘要完整</strong><div class="muted" style="margin-top:4px;">與前次同公司檔案相比，所有摘要均已存在，未發現缺少分錄。</div></div>`;
    return;
  }
  const missingHtml = missing.length
    ? `<div style="margin-top:8px;"><strong class="danger">前次有、本次無（可能缺少分錄，共 ${missing.reduce((s, x) => s + x.summaries.length, 0)} 筆）：</strong><ul style="margin:4px 0 0;padding-left:16px;font-size:13px;">${missing.map((m) => `<li>[${escapeHtml(m.code)}] ${m.summaries.map(escapeHtml).join('、')}</li>`).join('')}</ul></div>` : '';
  const addedHtml = added.length
    ? `<div style="margin-top:8px;"><strong style="color:var(--ok);">本次新增摘要（${added.reduce((s, x) => s + x.summaries.length, 0)} 筆）：</strong><ul style="margin:4px 0 0;padding-left:16px;font-size:13px;">${added.map((a) => `<li>[${escapeHtml(a.code)}] ${a.summaries.map(escapeHtml).join('、')}</li>`).join('')}</ul></div>` : '';
  el.innerHTML = `<div class="card" style="border-color:var(--warn);background:#fffbe6;margin-bottom:10px;"><strong style="color:var(--warn);">🔄 跨期比對（同公司 ${escapeHtml(newKey)}…）</strong>${missingHtml}${addedHtml}</div>`;
}
// User settings persistence (localStorage) for F2 parameters
function loadUserSettings() {
  try {
    const sTh = Number(localStorage.getItem("f2_suggestThreshold"));
    const win = Number(localStorage.getItem("f2_timeWindowDays"));
    const kmax = Number(localStorage.getItem("f2_subsetMaxK"));
    const tol = localStorage.getItem("f2_tolerance");
    if (Number.isFinite(sTh)) AppState.offset.suggestThreshold = sTh;
    if (Number.isFinite(win)) AppState.offset.timeWindowDays = win;
    if (Number.isFinite(kmax)) AppState.offset.subsetMaxK = kmax;
    if (tol != null && dom?.f2Tolerance) { dom.f2Tolerance.value = String(tol); AppState.offset.tolerance = Number(tol); }
  } catch { /* ignore */ }
}
function persistOffsetSetting(key, value) {
  try { localStorage.setItem(key, String(value)); } catch { /* ignore */ }
}

const dom = {
  fileInput: document.getElementById('fileInput'), resetBtn: document.getElementById('resetBtn'), metaText: document.getElementById('metaText'),
  stats: document.getElementById('stats'), accountSelect: document.getElementById('accountSelect'), keywordInput: document.getElementById('keywordInput'),
  navButtons: Array.from(document.querySelectorAll('.nav button[data-module]')), modules: Array.from(document.querySelectorAll('.module')),
  overviewList: document.getElementById('overviewList'), crossFileResult: document.getElementById('crossFileResult'), clearLocalBtn: document.getElementById('clearLocalBtn'), runF1Btn: document.getElementById('runF1Btn'), applyF1Btn: document.getElementById('applyF1Btn'),
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
  runF7Btn: document.getElementById('runF7Btn'), exportF7Btn: document.getElementById('exportF7Btn'), f7Result: document.getElementById('f7Result'),
  exportF3Btn: document.getElementById('exportF3Btn'), exportF5Btn: document.getElementById('exportF5Btn'), exportF6Btn: document.getElementById('exportF6Btn'),
  f3MinAmt: document.getElementById('f3MinAmt'), f3MaxAmt: document.getElementById('f3MaxAmt'),
  mergeF1Btn: document.getElementById('mergeF1Btn'), sortF1Btn: document.getElementById('sortF1Btn'),
  clearF4ReviewedBtn: document.getElementById('clearF4ReviewedBtn'),
  periodFrom: document.getElementById('periodFrom'), periodTo: document.getElementById('periodTo'),
  f9AccountSelect: document.getElementById('f9AccountSelect'), f9RuleAccount: document.getElementById('f9RuleAccount'),
  f9Memo: document.getElementById('f9Memo'), f9SaveMemoBtn: document.getElementById('f9SaveMemoBtn'), f9MemoStatus: document.getElementById('f9MemoStatus'),
  f9RuleKeyword: document.getElementById('f9RuleKeyword'), f9RuleFreq: document.getElementById('f9RuleFreq'), f9RuleAmount: document.getElementById('f9RuleAmount'),
  f9AddRuleBtn: document.getElementById('f9AddRuleBtn'), f9ClearRulesBtn: document.getElementById('f9ClearRulesBtn'), f9RuleList: document.getElementById('f9RuleList'),
  runF9Btn: document.getElementById('runF9Btn'), copyF9RequestBtn: document.getElementById('copyF9RequestBtn'), exportF9Btn: document.getElementById('exportF9Btn'), f9Result: document.getElementById('f9Result'),
  addTodoBtn: document.getElementById('addTodoBtn'), todoToggleBtn: document.getElementById('todoToggleBtn'), todoPanel: document.getElementById('todoPanel'),
  todoVoucher: document.getElementById('todoVoucher'), todoContent: document.getElementById('todoContent'),
  copyTodoAllBtn: document.getElementById('copyTodoAllBtn'), todoList: document.getElementById('todoList'), todoBadge: document.getElementById('todoBadge'), toastHost: document.getElementById('toastHost'),
  runF10Btn: document.getElementById('runF10Btn'), f10Result: document.getElementById('f10Result'), f10Tolerance: document.getElementById('f10Tolerance'),
  runF11Btn: document.getElementById('runF11Btn'), f11Result: document.getElementById('f11Result'), f11Field: document.getElementById('f11Field'),
  runF13Btn: document.getElementById('runF13Btn'), f13Result: document.getElementById('f13Result'), f13AccountSelect: document.getElementById('f13AccountSelect'), f13AnomalyMode: document.getElementById('f13AnomalyMode'),
  runF15Btn: document.getElementById('runF15Btn'), f15Result: document.getElementById('f15Result'), f15Sort: document.getElementById('f15Sort'), f15MinCount: document.getElementById('f15MinCount'), f15MaxCount: document.getElementById('f15MaxCount'),
  workbenchCard: document.getElementById('workbenchCard'), workbenchBody: document.getElementById('workbenchBody'), workbenchToggle: document.getElementById('workbenchToggle'),
  f1TabGroup: document.getElementById('f1TabGroup'), f1TabExclusion: document.getElementById('f1TabExclusion'), f1ExclusionView: document.getElementById('f1ExclusionView'), f1ExclusionResult: document.getElementById('f1ExclusionResult'),
  f1KeywordRulesBtn: document.getElementById('f1KeywordRulesBtn'), f1KeywordRulesPanel: document.getElementById('f1KeywordRulesPanel'),
  kwRuleName: document.getElementById('kwRuleName'), kwRuleKeywords: document.getElementById('kwRuleKeywords'), kwRuleExclude: document.getElementById('kwRuleExclude'), kwRuleTarget: document.getElementById('kwRuleTarget'), kwRulePriority: document.getElementById('kwRulePriority'), kwAddRuleBtn: document.getElementById('kwAddRuleBtn'), kwRuleList: document.getElementById('kwRuleList'),
  f1MainContent: document.getElementById('f1MainContent'),
};

function escapeHtml(v) { return String(v ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;'); }
function cleanText(v) { return v == null ? '' : String(v).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim(); }

function normalizeSummary(raw) {
  if (!raw) return '';
  let s = String(raw);
  // Remove carriage returns and _x000D_ artifacts
  s = s.replace(/_x000D_/gi, '').replace(/\r/g, '');
  // Normalize full-width spaces and common punctuation
  s = s.replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
  // Strip leading serial/sequence numbers like "123.", "A001-", "甲01."
  s = s.replace(/^[\d０-９]{1,6}[.\-、\s。]+/, '');
  s = s.replace(/^[A-Za-z]{0,3}[\d０-９]{2,6}[.\-\s]+/, '');
  // Strip trailing ROC/AD date patterns like 114/01/15 or 2025-01-15
  s = s.replace(/[\s　]+\d{2,4}[-/]\d{1,2}[-/]\d{1,2}$/, '');
  // Strip parenthetical internal code patterns like (AB-123) if code-like
  s = s.replace(/\s*\([A-Z0-9\-_]{2,12}\)\s*$/, '');
  // Final trim
  s = s.trim();
  return s || String(raw).trim();
}

function classifyRow(row, map) {
  const texts = row.map((c) => String(c == null ? '' : c).trim());
  const joined = texts.join(' ');
  const hasAnyNum = row.some((c) => typeof c === 'number' && c !== 0);
  const hasVoucher = map.voucherNo >= 0 && texts[map.voucherNo] && texts[map.voucherNo].length >= 1;
  // Month/period totals
  if (/月計|本月合計|月份合計|小計|本期合計/.test(joined)) return 'month_total';
  // Cumulative totals
  if (/累計|年度累計|至今|本年累計/.test(joined)) return 'cumulative_total';
  // Opening balance
  if (/上期結轉|前期餘額|期初餘額|期初|上期餘額/.test(joined)) return 'opening_balance';
  // Table header row (column labels)
  if (/日期|摘要|借方|貸方|傳票|科目/.test(joined) && !hasAnyNum) return 'table_header';
  // Account header
  if (/科目：|科目:/.test(joined) && !hasVoucher) return 'account_header';
  // Transaction: has voucher or has numeric amounts
  if (hasVoucher && hasAnyNum) return 'transaction';
  if (hasAnyNum && !hasVoucher) return 'transaction'; // might be a valid entry
  // No numbers at all → likely header/page
  if (!hasAnyNum) return 'page_header';
  return 'unknown';
}

function applyKeywordRules(summaryNorm, rules) {
  if (!rules || !rules.length) return null;
  const sorted = [...rules].filter((r) => r.enabled).sort((a, b) => (b.priority || 0) - (a.priority || 0));
  for (const rule of sorted) {
    const kws = (rule.keywords || []).map((k) => k.toLowerCase());
    const excl = (rule.excludeWords || []).map((k) => k.toLowerCase());
    const low = summaryNorm.toLowerCase();
    if (kws.length && !kws.some((k) => low.includes(k))) continue;
    if (excl.some((k) => low.includes(k))) continue;
    return rule.targetGroup;
  }
  return null;
}

const STORAGE_VERSION = 1;

function loadKeywordRules() {
  try {
    const raw = localStorage.getItem('ledger_keyword_rules');
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (parsed.version !== STORAGE_VERSION) return [];
    return Array.isArray(parsed.rules) ? parsed.rules : [];
  } catch { return []; }
}

function saveKeywordRules(rules) {
  try {
    localStorage.setItem('ledger_keyword_rules', JSON.stringify({ version: STORAGE_VERSION, rules }));
  } catch { /* ignore */ }
}

function loadWhitelist() {
  try { return JSON.parse(localStorage.getItem('ledger_exclusion_whitelist') || '[]'); }
  catch { return []; }
}

function saveWhitelist(list) {
  try { localStorage.setItem('ledger_exclusion_whitelist', JSON.stringify(list)); } catch { /* ignore */ }
}
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
  const excludedRows = [];
  let currentAccount = null;

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];
    const colA = cleanText(row[map.date]);
    if (!colA) {
      ig.other += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: classifyRow(row, map), reason: '空白列', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }
    if (looksLikeHeaderRow(colA)) {
      ig.header += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: 'page_header', reason: '頁首/製表標頭', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }
    if (colA.startsWith('月計:') || colA.startsWith('累計:')) {
      ig.monthly += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: classifyRow(row, map) || 'month_total', reason: '月計/累計列', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }
    if (colA.startsWith('項')) {
      const acc = parseAccountHeader(colA);
      if (acc) {
        currentAccount = acc;
        if (!accounts[acc.code]) accounts[acc.code] = { ...acc, openingBalance: null, transactionIds: [] };
      }
      ig.accountHeader += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: 'account_header', reason: '科目標頭', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }
    if (colA === '上期結轉') {
      if (currentAccount && map.balance >= 0 && accounts[currentAccount.code]) accounts[currentAccount.code].openingBalance = decToNum(dec(row[map.balance]));
      ig.opening += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: 'opening_balance', reason: '上期結轉列', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }

    const voucherNo = cleanText(row[map.voucherNo]); const d = parseROCDate(colA);
    if (!currentAccount || !voucherNo || !d) {
      ig.invalid += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: classifyRow(row, map), reason: '無效列（缺科目/傳票/日期）', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: true });
      continue;
    }

    const debit = map.debit >= 0 ? dec(row[map.debit]) : dec(0);
    const credit = map.credit >= 0 ? dec(row[map.credit]) : dec(0);
    const txnSummary = pickSummary(row, map);
    const dupKey = `${currentAccount.code}||${voucherNo}||${d.dateISO}||${decKey(debit)}||${decKey(credit)}||${txnSummary}`;
    if (seen.has(dupKey)) {
      ig.crossPageDup += 1;
      excludedRows.push({ rowIndex: rowIdx, rowType: 'transaction', reason: '跨頁重複列', rawContent: row.slice(0, 8).map((c) => String(c == null ? '' : c)), whitelisted: false, recoverable: false });
      continue;
    }
    seen.add(dupKey);

    const rawSummary = txnSummary;
    const summaryNormalized = normalizeSummary(rawSummary);
    const txn = {
      id: Math.random().toString(36).slice(2, 10), accountCode: currentAccount.code, accountName: currentAccount.name,
      accountNormalSide: currentAccount.normalSide, voucherNo, date: d.date, dateISO: d.dateISO, dateROC: d.dateROC,
      periodROC: d.periodROC, summary: txnSummary, rawSummary, summaryNormalized,
      debit: decToNum(debit), credit: decToNum(credit),
      drCr: map.drCr >= 0 ? cleanText(row[map.drCr]) : '', balance: map.balance >= 0 ? decToNum(dec(row[map.balance])) : 0,
      hasMultiSummaryInVoucher: false, keywordGroupName: '', defaultGroupName: '', manualGroupName: '', effectiveGroupName: '',
    };
    txns.push(txn); accounts[currentAccount.code].transactionIds.push(txn.id);
  }

  // Compute hasMultiSummaryInVoucher
  const voucherSummaryMap = {};
  for (const t of txns) {
    if (!voucherSummaryMap[t.voucherNo]) voucherSummaryMap[t.voucherNo] = new Set();
    voucherSummaryMap[t.voucherNo].add(t.rawSummary);
  }
  for (const t of txns) {
    t.hasMultiSummaryInVoucher = (voucherSummaryMap[t.voucherNo]?.size || 0) > 1;
  }

  // Apply keyword rules and compute effectiveGroupName
  const kwRules = loadKeywordRules();
  for (const t of txns) {
    t.keywordGroupName = applyKeywordRules(t.summaryNormalized, kwRules) || '';
    t.effectiveGroupName = t.manualGroupName || t.keywordGroupName || t.defaultGroupName || t.summaryNormalized;
  }

  AppState.accounts = accounts; AppState.transactions = txns; AppState.meta.company = fileName; AppState.meta.parsedAt = new Date().toISOString();
  AppState.meta.skippedRows = Object.values(ig).reduce((a, b) => a + b, 0); AppState.meta.igCount = ig; AppState.meta.columnMap = map;
  AppState.meta.excludedRows = excludedRows;
  AppState.grouping = emptyGroupingState();
  AppState.grouping.keywordRules = kwRules;
  AppState.groupingStore = {};
  AppState.groupingStore.all = cloneGroupingState(AppState.grouping);
  AppState.activeGroupingKey = 'all';
  AppState.offset.unmatchedGroups = [];

  if (hasFuse) searchEngine = new Fuse(txns, { keys: ['summary', 'voucherNo', 'accountName'], threshold: 0.35 });
  toast(`解析完成：${txns.length} 筆有效分錄｜${Object.keys(accounts).length} 個科目｜${ig.crossPageDup} 筆跨頁重複已合併｜${AppState.meta.skippedRows} 列已忽略`);
  // Cross-file comparison (must run before saveCompanyData overwrites previous)
  renderCrossFileResult(fileName, txns);
  saveCompanyData(fileName, txns);
  // Clear stale F9 scan result so user reruns against new file
  AppState.memo.missingResults = [];
  if (dom.f9Result) dom.f9Result.innerHTML = '<p class="muted">已載入新檔案，請重新執行「掃描缺少分錄」。</p>';
  renderBase();
}

function getFilteredTransactions() {
  let rows = AppState.transactions;
  const account = dom.accountSelect.value; const keyword = cleanText(dom.keywordInput.value);
  const pFrom = parseRocPeriod(dom.periodFrom?.value || '');
  const pTo = parseRocPeriod(dom.periodTo?.value || '');
  if (account && account !== 'all') rows = rows.filter((t) => t.accountCode === account);
  if (keyword) {
    if (searchEngine) {
      const ids = new Set(searchEngine.search(keyword).map((x) => x.item.id)); rows = rows.filter((t) => ids.has(t.id));
    } else rows = rows.filter((t) => t.summary.includes(keyword) || t.voucherNo.includes(keyword));
  }
  if (pFrom) rows = rows.filter((t) => t.periodROC >= pFrom);
  if (pTo) rows = rows.filter((t) => t.periodROC <= pTo);
  return rows;
}

function renderTxnList(el, rows, note = '', opts = {}) {
  const collapsible = opts.collapsible !== false;
  const limit = opts.limit || 250;
  const show = rows.slice(0, limit);
  const hasMore = rows.length > limit;
  if (!collapsible) {
    el.innerHTML = `<p class="muted">${escapeHtml(note || `顯示 ${show.length}/${rows.length} 筆`)}${hasMore ? ` <span style="color:var(--warn);">（僅顯示前 ${limit} 筆，請用篩選縮小範圍）</span>` : ''}</p><div class="table-wrap"><table><thead><tr><th>#</th><th class="col-date">日期</th><th class="col-voucher">傳票</th><th>科目</th><th class="col-summary">摘要</th><th class="col-amount">金額(符號)</th><th class="col-amount">餘額</th></tr></thead><tbody>${show.map((t, idx) => `<tr><td>${idx + 1}</td><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.voucherNo)}</td><td>${escapeHtml(t.accountName)}${t.accountNormalSide ? `（${escapeHtml(t.accountNormalSide)}）` : ''} <span class="muted" style="font-size:11px;">(${escapeHtml(t.accountCode)})</span></td><td class="col-summary">${escapeHtml(t.rawSummary || t.summary || '(空白摘要)')}${t.hasMultiSummaryInVoucher ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要</span>' : ''}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td></tr>`).join('')}</tbody></table></div>`;
    return;
  }
  el.innerHTML = `<p class="muted">${escapeHtml(note || `顯示 ${show.length}/${rows.length} 筆`)}${hasMore ? ` <span style="color:var(--warn);">（顯示前 ${limit} 筆）</span>` : ''}</p>${show.map((t, idx) => {
    const multiPill = t.hasMultiSummaryInVoucher ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要</span>' : '';
    const normSameAsRaw = (t.summaryNormalized || '') === (t.rawSummary || '');
    const normDisplay = normSameAsRaw ? '（與原始相同）' : escapeHtml(t.summaryNormalized || '');
    const groupSrc = t.manualGroupName ? '手動' : t.keywordGroupName ? '關鍵字規則' : t.defaultGroupName ? '預設規則' : 'none';
    return `<details class="card" style="margin:6px 0;padding:8px 10px;"><summary style="cursor:pointer;list-style:none;"><strong>${idx + 1}.</strong> ${escapeHtml(t.dateROC)}｜${escapeHtml(t.voucherNo)}｜${escapeHtml(t.accountName)}｜${escapeHtml((t.rawSummary || t.summary || '(空白摘要)').slice(0, 60))}｜${fmtSigned(getSignedAmount(t))}${multiPill}</summary><div style="margin-top:8px;display:grid;gap:4px;font-size:13px;"><div><strong>原始摘要：</strong>${escapeHtml(t.rawSummary || t.summary || '(空白摘要)')}</div><div><strong>正規化摘要：</strong>${normDisplay}</div><div><strong>預設分組：</strong>${escapeHtml(t.defaultGroupName || '（無）')}</div><div><strong>關鍵字分組：</strong>${escapeHtml(t.keywordGroupName || '（無）')}</div><div><strong>手動分組：</strong>${escapeHtml(t.manualGroupName || '（無）')}</div><div><strong>有效分組：</strong>${escapeHtml(t.effectiveGroupName || t.summaryNormalized || t.rawSummary || '（無）')}</div><div><strong>分組來源：</strong>${escapeHtml(groupSrc)}</div><div><strong>傳票號碼：</strong>${escapeHtml(t.voucherNo)}</div><div><strong>日期：</strong>${escapeHtml(t.dateROC)} (${escapeHtml(t.dateISO)})</div><div><strong>科目：</strong>${escapeHtml(t.accountName)}${t.accountNormalSide ? `（${escapeHtml(t.accountNormalSide)}）` : ''} <span class="muted" style="font-size:11px;">[${escapeHtml(t.accountCode)}]</span></div><div><strong>借方 / 貸方：</strong>${fmtAmount(t.debit)} / ${fmtAmount(t.credit)}</div><div><strong>金額(符號)：</strong>${fmtSigned(getSignedAmount(t))}</div><div><strong>餘額：</strong>${fmtAmount(t.balance)}</div><div><strong>多摘要傳票：</strong>${t.hasMultiSummaryInVoucher ? '是' : '否'}</div></div></details>`;
  }).join('')}`;
}

function _ovSortRows(rows) {
  const { col, dir } = _ovSort;
  return rows.slice().sort((a, b) => {
    let va, vb;
    if (col === 'amount') { va = getSignedAmount(a); vb = getSignedAmount(b); }
    else if (col === 'balance') { va = Number(a.balance || 0); vb = Number(b.balance || 0); }
    else if (col === 'debit') { va = Number(a.debit || 0); vb = Number(b.debit || 0); }
    else if (col === 'credit') { va = Number(a.credit || 0); vb = Number(b.credit || 0); }
    else { va = String(a[col] || ''); vb = String(b[col] || ''); }
    let cmp = typeof va === 'number' ? va - vb : String(va).localeCompare(String(vb), 'zh-TW');
    return dir === 'desc' ? -cmp : cmp;
  });
}

function renderOverviewTable(rows) {
  if (!dom.overviewList) return;
  const sorted = _ovSortRows(rows);
  const totalPages = Math.max(1, Math.ceil(sorted.length / OV_PAGE_SIZE));
  if (_ovPage > totalPages) _ovPage = totalPages;
  const start = (_ovPage - 1) * OV_PAGE_SIZE;
  const show = sorted.slice(start, start + OV_PAGE_SIZE);

  const cols = [
    { key: 'dateROC', label: '日期' },
    { key: 'voucherNo', label: '傳票' },
    { key: 'accountCode', label: '科目' },
    { key: 'summary', label: '摘要' },
    { key: 'amount', label: '金額(符號)' },
    { key: 'balance', label: '餘額' },
  ];
  const thRow = cols.map((c) => {
    const isActive = _ovSort.col === c.key;
    const icon = isActive ? (_ovSort.dir === 'asc' ? '↑' : '↓') : '⇕';
    const cls = `sortable${isActive ? ` sort-${_ovSort.dir}` : ''}`;
    return `<th class="${cls}" data-ov-sort="${c.key}">${escapeHtml(c.label)} <span class="sort-icon">${icon}</span></th>`;
  }).join('');

  const pageNav = totalPages > 1
    ? `<div class="pagination">
        <button data-ov-page="prev" ${_ovPage <= 1 ? 'disabled' : ''}>‹ 上頁</button>
        <span class="page-info">第 ${_ovPage} / ${totalPages} 頁（共 ${rows.length} 筆）</span>
        <button data-ov-page="next" ${_ovPage >= totalPages ? 'disabled' : ''}>下頁 ›</button>
        <select data-ov-jump style="padding:4px 6px;font-size:12px;">${Array.from({ length: totalPages }, (_, i) => `<option value="${i + 1}"${i + 1 === _ovPage ? ' selected' : ''}>第 ${i + 1} 頁</option>`).join('')}</select>
      </div>`
    : `<p class="muted" style="margin:6px 0;">共 ${rows.length} 筆</p>`;

  dom.overviewList.innerHTML = `${pageNav}<div class="table-wrap"><table><thead><tr><th>#</th>${thRow}</tr></thead><tbody>
    ${show.map((t, i) => `<tr>
      <td style="color:#9bb2cc;font-size:12px;">${start + i + 1}</td>
      <td class="col-date">${escapeHtml(t.dateROC)}</td>
      <td class="col-voucher">${escapeHtml(t.voucherNo)}</td>
      <td>${escapeHtml(t.accountName)}${t.accountNormalSide ? `（${escapeHtml(t.accountNormalSide)}）` : ''} <span class="muted" style="font-size:11px;">(${escapeHtml(t.accountCode)})</span></td>
      <td class="col-summary">${escapeHtml(t.rawSummary || t.summary || '(空白摘要)')}${t.hasMultiSummaryInVoucher ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要</span>' : ''}</td>
      <td class="col-amount">${fmtSigned(getSignedAmount(t))}</td>
      <td class="col-amount">${fmtAmount(t.balance)}</td>
    </tr>`).join('')}
  </tbody></table></div>`;
}

function renderOverviewSummary(rows) {
  if (!rows.length || !dom.overviewList) return;
  const byAcc = new Map();
  rows.forEach((t) => {
    if (!byAcc.has(t.accountCode)) byAcc.set(t.accountCode, { code: t.accountCode, name: t.accountName, normalSide: t.accountNormalSide, debit: 0, credit: 0, count: 0, lastBal: 0 });
    const x = byAcc.get(t.accountCode);
    x.debit += t.debit || 0; x.credit += t.credit || 0; x.count += 1; x.lastBal = t.balance || 0;
  });
  const accs = Array.from(byAcc.values()).sort((a, b) => a.code.localeCompare(b.code));
  if (!accs.length) return;
  const totalDebit = accs.reduce((s, a) => s + a.debit, 0);
  const totalCredit = accs.reduce((s, a) => s + a.credit, 0);
  const summaryEl = document.getElementById('overviewSummary');
  if (!summaryEl) return;
  summaryEl.innerHTML = `<div style="margin-bottom:8px;font-weight:600;color:#2a4668;">篩選結果科目彙總（${accs.length} 個科目，共 ${rows.length} 筆）</div>
    <div class="table-wrap"><table style="min-width:700px;"><thead><tr>
      <th>科目代碼</th><th>科目名稱</th><th>借/貸</th><th>筆數</th>
      <th class="col-amount">借方合計</th><th class="col-amount">貸方合計</th>
      <th class="col-amount">淨額</th><th class="col-amount">末筆餘額</th>
    </tr></thead><tbody>
    ${accs.map((a) => {
      const net = a.normalSide === '貸' ? a.credit - a.debit : a.debit - a.credit;
      return `<tr>
        <td>${escapeHtml(a.code)}</td><td>${escapeHtml(a.name)}</td><td>${escapeHtml(a.normalSide || '—')}</td><td>${a.count}</td>
        <td class="col-amount">${fmtAmount(a.debit)}</td><td class="col-amount">${fmtAmount(a.credit)}</td>
        <td class="col-amount" style="font-weight:600;">${fmtSigned(net)}</td>
        <td class="col-amount">${fmtAmount(a.lastBal)}</td>
      </tr>`;
    }).join('')}
    </tbody><tfoot><tr style="font-weight:600;background:#f7fafe;">
      <td colspan="4">合計</td>
      <td class="col-amount">${fmtAmount(totalDebit)}</td><td class="col-amount">${fmtAmount(totalCredit)}</td>
      <td class="col-amount">${fmtSigned(totalDebit - totalCredit)}</td><td></td>
    </tr></tfoot></table></div>`;
}

function renderWorkbench() {
  const card = dom.workbenchCard;
  const body = dom.workbenchBody;
  if (!card || !body) return;
  card.style.display = '';

  const txns = AppState.transactions;
  const excluded = AppState.meta.excludedRows || [];

  // Row type breakdown
  const typeCount = {};
  for (const r of excluded) typeCount[r.rowType] = (typeCount[r.rowType] || 0) + 1;
  const typeSummary = Object.entries(typeCount).map(([t, n]) => `${t} ${n}`).join('、') || '無';

  // Top 20 ungrouped (no defaultGroupName, keywordGroupName, or manualGroupName)
  const ungrouped = {};
  for (const t of txns) {
    if (!t.defaultGroupName && !t.keywordGroupName && !t.manualGroupName) {
      const key = t.summaryNormalized || t.rawSummary || t.summary;
      ungrouped[key] = (ungrouped[key] || 0) + 1;
    }
  }
  const top20ungrouped = Object.entries(ungrouped).sort((a, b) => b[1] - a[1]).slice(0, 20);

  // Top 20 singletons (appear only once)
  const summaryCount = {};
  for (const t of txns) {
    const key = t.summaryNormalized || t.rawSummary || t.summary;
    summaryCount[key] = (summaryCount[key] || 0) + 1;
  }
  const singletons = Object.entries(summaryCount).filter(([, n]) => n === 1).slice(0, 20);

  body.innerHTML = `
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:10px;margin-bottom:12px;">
      <div class="card" style="padding:10px;text-align:center;">
        <div style="font-size:24px;font-weight:700;color:var(--brand);">${txns.length}</div>
        <div class="muted">有效分錄</div>
      </div>
      <div class="card" style="padding:10px;text-align:center;">
        <div style="font-size:24px;font-weight:700;color:var(--warn);">${excluded.length}</div>
        <div class="muted">被排除列 <button id="goExclusionBtn" style="font-size:11px;padding:2px 6px;margin-left:4px;">查看</button></div>
      </div>
      <div class="card" style="padding:10px;text-align:center;">
        <div style="font-size:24px;font-weight:700;color:var(--ink2);">${txns.filter((t) => t.hasMultiSummaryInVoucher).length}</div>
        <div class="muted">多摘要傳票分錄 <button id="goF4FromWorkbench" style="font-size:11px;padding:2px 6px;margin-left:4px;">查看異常</button></div>
      </div>
      <div class="card" style="padding:10px;text-align:center;">
        <div style="font-size:24px;font-weight:700;color:var(--ink);">${Object.keys(ungrouped).length}</div>
        <div class="muted">未分組摘要</div>
      </div>
      <div class="card" style="padding:10px;text-align:center;">
        <div style="font-size:24px;font-weight:700;color:var(--ink2);">${singletons.length}</div>
        <div class="muted">孤筆摘要</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
      <div>
        <div class="list-title" style="margin-top:0;">列型別統計</div>
        <div class="muted" style="font-size:12px;margin-top:4px;">${escapeHtml(typeSummary)}</div>
        <div class="list-title">未分組摘要 Top 20</div>
        ${top20ungrouped.length ? `<ul style="margin:4px 0 0;padding-left:16px;font-size:12px;">${top20ungrouped.map(([k, n]) => `<li>${escapeHtml(k)} <span class="pill">${n}筆</span></li>`).join('')}</ul>` : '<div class="muted" style="font-size:12px;margin-top:4px;">（全部已分組）</div>'}
      </div>
      <div>
        <div class="list-title" style="margin-top:0;">孤筆摘要 Top 20</div>
        ${singletons.length ? `<ul style="margin:4px 0 0;padding-left:16px;font-size:12px;">${singletons.map(([k]) => `<li>${escapeHtml(k)}</li>`).join('')}</ul>` : '<div class="muted" style="font-size:12px;margin-top:4px;">（無孤筆）</div>'}
        <div style="margin-top:12px;display:flex;flex-wrap:wrap;gap:6px;">
          <button id="goF1FromWorkbench" class="primary">前往摘要分組 →</button>
          <button id="goF15FromWorkbench">查看孤筆</button>
        </div>
      </div>
    </div>
  `;

  document.getElementById('goF1FromWorkbench')?.addEventListener('click', () => {
    document.querySelector('[data-module="f1"]')?.click();
  });
  document.getElementById('goExclusionBtn')?.addEventListener('click', () => {
    document.querySelector('[data-module="f1"]')?.click();
    setTimeout(() => { dom.f1TabExclusion?.click(); }, 100);
  });
  document.getElementById('goF15FromWorkbench')?.addEventListener('click', () => {
    document.querySelector('[data-module="f15"]')?.click();
  });
  document.getElementById('goF4FromWorkbench')?.addEventListener('click', () => {
    document.querySelector('[data-module="f4"]')?.click();
  });
}

function renderExclusionViewer() {
  const el = dom.f1ExclusionResult;
  if (!el) return;
  const rows = AppState.meta.excludedRows || [];
  if (!rows.length) {
    el.innerHTML = '<p class="muted">無排除記錄。</p>';
    return;
  }
  const map = AppState.meta.columnMap || {};
  el.innerHTML = `
    <div class="muted" style="margin-bottom:8px;">共 ${rows.length} 列被排除。點「加入白名單」可避免下次誤排；「恢復此列」暫時將此列加回明細並重新分組。</div>
    <div class="table-wrap">
      <table>
        <thead><tr>
          <th>列索引</th><th>列型別</th><th>排除原因</th><th>內容預覽</th><th>操作</th>
        </tr></thead>
        <tbody>
          ${rows.map((r, i) => `
            <tr style="${r.whitelisted || r.restored ? 'opacity:0.5;' : ''}">
              <td>${r.rowIndex + 1}</td>
              <td><span class="pill">${escapeHtml(r.rowType)}</span></td>
              <td>${escapeHtml(r.reason)}</td>
              <td style="font-size:12px;max-width:300px;word-break:break-word;">${r.rawContent.filter(Boolean).slice(0, 5).map(escapeHtml).join(' | ')}</td>
              <td style="white-space:nowrap;">
                ${r.recoverable && !r.whitelisted && !r.restored ? `<button data-excl-whitelist="${i}" style="font-size:11px;padding:2px 6px;">加入白名單</button>` : ''}
                ${!r.restored ? `<button data-excl-restore="${i}" style="font-size:11px;padding:2px 6px;margin-left:4px;" ${r.restored ? 'disabled' : ''}>恢復此列</button>` : '<span class="ok" style="font-size:11px;">已恢復</span>'}
                ${r.whitelisted ? ' <span class="ok" style="font-size:11px;">已白名單</span>' : ''}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
  el.addEventListener('click', (e) => {
    const wlIdx = e.target?.dataset?.exclWhitelist;
    if (wlIdx != null) {
      const row = rows[parseInt(wlIdx)];
      if (!row) return;
      row.whitelisted = true;
      const wl = loadWhitelist();
      wl.push(row.rawContent.filter(Boolean).join('|'));
      saveWhitelist(wl);
      renderExclusionViewer();
      toast('已加入白名單');
      return;
    }
    const restIdx = e.target?.dataset?.exclRestore;
    if (restIdx != null) {
      const row = rows[parseInt(restIdx)];
      if (!row || row.restored) return;
      row.restored = true;
      const rc = row.rawContent || [];
      const vNo = (map.voucherNo >= 0 ? rc[map.voucherNo] : '') || rc[1] || '';
      const joined = rc.filter(Boolean).join(' | ');
      const restoredTxn = {
        id: 'restored-' + Math.random().toString(36).slice(2),
        restored: true,
        accountCode: '', accountName: '（恢復列）',
        accountNormalSide: '',
        voucherNo: vNo,
        date: new Date(), dateISO: '', dateROC: rc[0] || '',
        periodROC: '',
        rawSummary: joined,
        summary: joined,
        summaryNormalized: joined,
        debit: 0, credit: 0, balance: 0,
        defaultGroupName: '', keywordGroupName: '', manualGroupName: '',
        effectiveGroupName: '（恢復列）',
        hasMultiSummaryInVoucher: false,
      };
      AppState.transactions.push(restoredTxn);
      renderBase();
      renderExclusionViewer();
      toast('已恢復 1 列，請重新分組');
    }
  });
}

function renderKeywordRuleList() {
  const el = dom.kwRuleList;
  if (!el) return;
  const rules = AppState.grouping.keywordRules || [];
  if (!rules.length) { el.innerHTML = '<p class="muted">尚無規則。</p>'; return; }
  el.innerHTML = rules.map((r, i) => `
    <div style="display:flex;align-items:center;gap:8px;padding:6px 8px;border:1px solid var(--line);border-radius:8px;margin-bottom:4px;background:#fff;">
      <input type="checkbox" ${r.enabled ? 'checked' : ''} data-kw-toggle="${i}" />
      <span style="flex:1;font-size:13px;"><strong>${escapeHtml(r.name)}</strong> → <em>${escapeHtml(r.targetGroup)}</em> &nbsp;<span class="muted">[${(r.keywords||[]).map(escapeHtml).join(', ')}]</span></span>
      <span class="muted" style="font-size:11px;">優先序 ${r.priority}</span>
      <button data-kw-delete="${i}" style="font-size:11px;padding:2px 6px;color:var(--danger);">刪除</button>
    </div>
  `).join('');
  el.addEventListener('click', (e) => {
    const del = e.target?.dataset?.kwDelete;
    if (del != null) {
      AppState.grouping.keywordRules.splice(parseInt(del), 1);
      saveKeywordRules(AppState.grouping.keywordRules);
      renderKeywordRuleList();
      return;
    }
  });
  el.addEventListener('change', (e) => {
    const toggle = e.target?.dataset?.kwToggle;
    if (toggle != null) {
      AppState.grouping.keywordRules[parseInt(toggle)].enabled = e.target.checked;
      saveKeywordRules(AppState.grouping.keywordRules);
    }
  });
}

function renderBase() {
  renderAccountSelect(); renderStats(); renderF9AccountSelects();
  const rows = getFilteredTransactions();
  renderOverviewSummary(rows);
  renderOverviewTable(rows);
  renderWorkbench();
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
  dom.stats.textContent = `有效分錄 ${AppState.transactions.length}｜科目 ${Object.keys(AppState.accounts).length}｜略過：標題行 ${ig.header} 月計 ${ig.monthly} 上期結轉 ${ig.opening} 科目標頭 ${ig.accountHeader} 跨頁重複 ${ig.crossPageDup} 無效 ${ig.invalid + ig.other}｜摘要欄第 ${AppState.meta.columnMap.summary + 1} 欄`;
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
      const sa = a.summaryNormalized || a.summary;
      const sb = b.summaryNormalized || b.summary;
      const s = jaccard(sa, sb) * 100;
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
  const totalBudgetMs = 5000; // 全部 suggest 最多跑 5 秒，避免大量未沖帳時凍結 UI
  const totalStart = Date.now();
  const th = Math.max(0, Math.min(100, Number(thresholdPct) || 80));
  const txMap = new Map(rows.map((r) => [r.id, r]));
  const debits = rows.filter((x) => x.debit > 0);
  const credits = rows.filter((x) => x.credit > 0);
  const suggestions = [];

  // 借方多筆 ≈ 貸方單筆
  for (const c of credits) {
    if (Date.now() - totalStart > totalBudgetMs) break;
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
    if (Date.now() - totalStart > totalBudgetMs) break;
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
      <div class="muted" style="margin-top:6px;">未沖帳分組模式：可改名、批次移動、刪除群組，操作方式同摘要分組。</div>
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
          return `<button data-f2-suggest="${escapeHtml(payload)}" title="${escapeHtml(title)}">${escapeHtml(label)}</button> <button data-f2-apply="${escapeHtml(payload)}" title="一鍵套用此建議">⚡套用</button>`;
        }).join('')}
      </div>`
    : `<div class="muted" style="margin-top:6px;">（目前無建議沖帳）</div>`;

  dom.f2UnmatchedSummary.innerHTML = `
    <div class="muted">未沖帳剩餘 ${rows.length} 筆。可先確認是否有遺漏，再按「進入分組」整理歸類。</div>
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
  // 帳齡：以未沖帳清單中最晚一筆日期為基準計算
  const refDateF2 = rows.reduce((mx, t) => (t.date instanceof Date && t.date > mx ? t.date : mx), new Date(0));
  const hasRef = refDateF2.getTime() > 0;
  dom.f2List.innerHTML = `<p class="muted">未沖帳清單 ${rows.length} 筆（顯示前 ${shown.length} 筆）｜勾選 2 筆（一借一貸）→「勾選加入沖帳」</p>
    <div class="table-wrap"><table><thead><tr><th>勾選</th><th>日期</th><th>帳齡(天)</th><th>傳票</th><th>科目</th><th>摘要</th><th class="col-amount">簽帳金額</th></tr></thead><tbody>
    ${shown.map((t) => {
      const ageDays = hasRef && t.date instanceof Date ? daysBetween(t.date, refDateF2) : null;
      const ageText = ageDays !== null ? String(ageDays) : '—';
      const ageStyle = ageDays !== null && ageDays > 90 ? 'color:#cf1322;font-weight:600;' : ageDays > 30 ? 'color:#d46b08;' : 'color:#5f7692;';
      const multiPillF2 = t.hasMultiSummaryInVoucher ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要</span>' : '';
      return `<tr id="f2-row-${t.id}"><td><input type="checkbox" data-f2-pick="${t.id}" /></td><td>${escapeHtml(t.dateROC)}</td><td style="${ageStyle}">${ageText}</td><td>${escapeHtml(t.voucherNo)}</td><td>${escapeHtml(t.accountName)} <span class="muted" style="font-size:11px;">[${escapeHtml(t.accountCode)}]</span></td><td>${escapeHtml(t.rawSummary || t.summary || '(空白摘要)')}${multiPillF2}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td></tr>`;
    }).join('')}
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
  dom.f1Result.innerHTML = `<p class="muted">已產生 ${names.length} 個候選分組名稱（此步驟僅預覽名稱，不會先分組）。請先調整名稱，再按「套用分組」。</p>
    <div class="toolbar" style="margin-bottom:8px;">
      <button data-f1-bulk-delete="1" style="border-color:#ffb8b8;color:var(--danger);">批量刪除分組…</button>
    </div>
    ${blocks || '<p class="muted">無可分組資料。</p>'}`;
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
  // Update defaultGroupName and effectiveGroupName on transactions
  const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
  AppState.grouping.groups.forEach((g) => {
    g.transactionIds.forEach((id) => {
      const t = txMap.get(id);
      if (t) {
        t.defaultGroupName = g.name;
        t.effectiveGroupName = t.manualGroupName || t.keywordGroupName || t.defaultGroupName || t.summaryNormalized || t.rawSummary;
      }
    });
  });
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
  // Update defaultGroupName and effectiveGroupName on transactions
  const txMapEdit = new Map(AppState.transactions.map((t) => [t.id, t]));
  AppState.grouping.groups.forEach((g) => {
    g.transactionIds.forEach((id) => {
      const t = txMapEdit.get(id);
      if (t) {
        t.defaultGroupName = g.name;
        t.effectiveGroupName = t.manualGroupName || t.keywordGroupName || t.defaultGroupName || t.summaryNormalized || t.rawSummary;
      }
    });
  });
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
      if (!byAccount.has(t.accountCode)) byAccount.set(t.accountCode, { debit: 0, credit: 0, signed: 0 });
      const x = byAccount.get(t.accountCode);
      x.debit += t.debit || 0; x.credit += t.credit || 0; x.signed += getSignedAmount(t);
    });
    const checks = Array.from(byAccount.entries()).map(([acc, totals]) => {
      const lastBal = accountLastBalance.get(acc) ?? 0;
      return `<div class="muted">[${escapeHtml(acc)}] 本組：借 ${fmtAmount(totals.debit)} ／貸 ${fmtAmount(totals.credit)} ／淨額 ${fmtSigned(totals.signed)}｜科目末筆餘額 ${fmtAmount(lastBal)}</div>`;
    }).join('');
    const anchorId = asGroupAnchor(g.id);
    const rule = g.rule || (g.rule = { mode: 'A', keyword: '', threshold: 70 });
    const mode = rule.mode || 'A';
    const th = Number.isFinite(Number(rule.threshold)) ? Number(rule.threshold) : 70;

    // Determine group source badge from transactions
    const manualCount = txns.filter((t) => t.manualGroupName).length;
    const kwCount = txns.filter((t) => t.keywordGroupName).length;
    let srcBadge;
    if (manualCount > 0 && manualCount === txns.length) {
      srcBadge = '<span class="pill" style="font-size:11px;background:#f6ffed;color:#237804;border-color:#b7eb8f;">手動</span>';
    } else if (kwCount > txns.length / 2) {
      srcBadge = '<span class="pill" style="font-size:11px;background:#e6f7ff;color:#0050b3;border-color:#91d5ff;">關鍵字</span>';
    } else {
      srcBadge = '<span class="pill" style="font-size:11px;background:#f0f0f0;color:#595959;border-color:#d9d9d9;">預設規則</span>';
    }
    const hasMulti = txns.some((t) => t.hasMultiSummaryInVoucher);
    const multiWarning = hasMulti ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要傳票</span>' : '';
    // Use effectiveGroupName as display header (fallback to g.name)
    const effectiveName = txns.length > 0 ? (txns[0].effectiveGroupName || g.name) : g.name;

    return `<details class="card" id="${anchorId}" style="margin:8px 0;" open>
      <summary style="cursor:pointer;"><strong>群組：</strong><input data-f1-rename="${g.id}" value="${escapeHtml(g.name)}" style="margin-left:8px;min-width:180px;" /> ${srcBadge}${multiWarning}｜有效名稱：<em style="color:#0050b3;">${escapeHtml(effectiveName)}</em>｜筆數 ${txns.length}｜群組簽帳合計 ${fmtSigned(sumSigned)} <button data-f1-del-group="${g.id}" style="margin-left:8px;">刪除群組</button></summary>
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
      <div style="margin-top:8px;">
      ${txns.map((t) => {
        const multiPill = t.hasMultiSummaryInVoucher ? ' <span class="pill" style="background:#fff1f0;color:#cf1322;border-color:#ffa39e;font-size:11px;">多摘要</span>' : '';
        const normSame = (t.summaryNormalized || '') === (t.rawSummary || '');
        const normDisp = normSame ? '（與原始相同）' : escapeHtml(t.summaryNormalized || '');
        const tGroupSrc = t.manualGroupName ? '手動' : t.keywordGroupName ? '關鍵字規則' : t.defaultGroupName ? '預設規則' : 'none';
        return `<details style="border:1px solid var(--line);border-radius:8px;margin:4px 0;padding:4px 8px;background:#fafbfd;">
          <summary style="cursor:pointer;list-style:none;font-size:13px;"><input type="checkbox" data-f1-pick="${t.id}" data-f1-from="${g.id}" style="margin-right:4px;" />${escapeHtml(t.dateROC)}｜${escapeHtml(t.voucherNo || '')}｜${escapeHtml(t.accountName)}｜${escapeHtml((t.rawSummary || t.summary || '(空白摘要)').slice(0, 60))}｜${fmtSigned(getSignedAmount(t))}${multiPill}
            <select data-f1-target="${t.id}" style="margin-left:6px;font-size:12px;"><option value="">移動到...</option>${groupOptions}</select>
            <button data-f1-move="${t.id}" data-f1-from="${g.id}" style="font-size:11px;padding:1px 5px;">移動</button>
            <button data-f1-del="${t.id}" data-f1-from="${g.id}" style="font-size:11px;padding:1px 5px;">刪除</button>
          </summary>
          <div style="margin-top:6px;display:grid;gap:3px;font-size:12px;padding-left:8px;">
            <div><strong>原始摘要：</strong>${escapeHtml(t.rawSummary || t.summary || '(空白摘要)')}</div>
            <div><strong>正規化摘要：</strong>${normDisp}</div>
            <div><strong>預設分組：</strong>${escapeHtml(t.defaultGroupName || '（無）')}</div>
            <div><strong>關鍵字分組：</strong>${escapeHtml(t.keywordGroupName || '（無）')}</div>
            <div><strong>手動分組：</strong>${escapeHtml(t.manualGroupName || '（無）')}</div>
            <div><strong>有效分組：</strong>${escapeHtml(t.effectiveGroupName || '（無）')}</div>
            <div><strong>分組來源：</strong>${escapeHtml(tGroupSrc)}</div>
            <div><strong>傳票號碼：</strong>${escapeHtml(t.voucherNo)}</div>
            <div><strong>日期：</strong>${escapeHtml(t.dateROC)}</div>
            <div><strong>借方 / 貸方：</strong>${fmtAmount(t.debit)} / ${fmtAmount(t.credit)}</div>
            <div><strong>多摘要傳票：</strong>${t.hasMultiSummaryInVoucher ? '是' : '否'}</div>
          </div>
        </details>`;
      }).join('')}
      </div>
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

function showF1BulkDeleteModal() {
  const allNames = Array.from(new Set(
    (AppState.grouping.draftItems || []).map((d) => d.proposed)
      .concat(AppState.grouping.draftExtraGroups || [])
  )).filter((n) => n && n !== '其他').sort((a, b) => a.localeCompare(b, 'zh-TW'));

  if (!allNames.length) { toast('沒有可刪除的分組', 'WARN'); return; }

  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `<div class="modal-box">
    <div class="modal-header">
      <h4>批量刪除分組</h4>
      <button data-modal-close style="background:transparent;border:none;font-size:18px;color:#9bb2cc;padding:0 4px;">✕</button>
    </div>
    <div style="display:flex;align-items:center;gap:10px;padding:6px 12px;background:#f7fafe;border-radius:8px;">
      <label style="display:flex;align-items:center;gap:8px;font-size:13px;font-weight:600;cursor:pointer;">
        <input type="checkbox" id="modalSelectAll" style="width:15px;height:15px;" /> 全選（${allNames.length} 個）
      </label>
      <span class="muted" id="modalSelectedCount" style="margin-left:auto;">已選 0 個</span>
    </div>
    <div class="modal-body" style="max-height:320px;">
      ${allNames.map((n) => `<label class="modal-item">
        <input type="checkbox" class="modal-group-cb" value="${escapeHtml(n)}" />
        <span style="flex:1;">${escapeHtml(n)}</span>
      </label>`).join('')}
    </div>
    <div class="modal-footer">
      <span class="muted" style="font-size:12px;">選取後將把該分組的分錄移入「其他」</span>
      <div style="display:flex;gap:8px;">
        <button data-modal-close>取消</button>
        <button id="modalConfirmDel" class="danger">刪除選取</button>
      </div>
    </div>
  </div>`;
  document.body.appendChild(overlay);

  const updateCount = () => {
    const n = overlay.querySelectorAll('.modal-group-cb:checked').length;
    document.getElementById('modalSelectedCount').textContent = `已選 ${n} 個`;
    const btn = document.getElementById('modalConfirmDel');
    if (btn) btn.disabled = n === 0;
  };

  const selectAll = document.getElementById('modalSelectAll');
  selectAll.addEventListener('change', (e) => {
    overlay.querySelectorAll('.modal-group-cb').forEach((cb) => { cb.checked = e.target.checked; });
    updateCount();
  });
  overlay.querySelectorAll('.modal-group-cb').forEach((cb) => cb.addEventListener('change', () => {
    selectAll.indeterminate = true;
    updateCount();
  }));
  overlay.querySelectorAll('[data-modal-close]').forEach((btn) => btn.addEventListener('click', () => overlay.remove()));
  overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });

  document.getElementById('modalConfirmDel').addEventListener('click', () => {
    const toDelete = new Set(Array.from(overlay.querySelectorAll('.modal-group-cb:checked')).map((cb) => cb.value));
    if (!toDelete.size) return;
    AppState.grouping.draftItems.forEach((d) => { if (toDelete.has(d.proposed)) d.proposed = '其他'; });
    AppState.grouping.draftExtraGroups = (AppState.grouping.draftExtraGroups || []).filter((x) => !toDelete.has(x));
    toDelete.forEach((n) => delete AppState.grouping.draftRenameMap[n]);
    AppState.grouping.mode = 'draft';
    overlay.remove();
    renderF1Draft();
    toast(`已批量刪除 ${toDelete.size} 個分組`);
  });
  updateCount();
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
      const s = jaccard(d.summaryNormalized || d.summary, c.summaryNormalized || c.summary);
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

function renderF3CandidateList(rows) {
  const direction = dom.f3Direction.value;
  const minAmt = Number(dom.f3MinAmt?.value || 0) || 0;
  const maxAmt = Number(dom.f3MaxAmt?.value || 0) || 0;
  let cands = rows.filter((t) => direction === 'debit' ? t.debit > 0 : t.credit > 0);
  if (minAmt > 0) cands = cands.filter((t) => (direction === 'debit' ? t.debit : t.credit) >= minAmt);
  if (maxAmt > 0) cands = cands.filter((t) => (direction === 'debit' ? t.debit : t.credit) <= maxAmt);
  const show = cands.slice(0, 300);
  dom.f3List.innerHTML = `<p class="muted">候選 ${cands.length} 筆（點擊金額欄可自動填入目標）${cands.length > 300 ? '｜僅顯示前 300 筆' : ''}</p>
    <div class="table-wrap"><table><thead><tr>
      <th class="col-date">日期</th><th class="col-voucher">傳票</th><th>科目</th><th class="col-summary">摘要</th>
      <th class="col-amount" title="點擊金額可填入目標">金額 ▶ 點填目標</th>
    </tr></thead><tbody>
    ${show.map((t) => {
      const amt = direction === 'debit' ? t.debit : t.credit;
      return `<tr><td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.voucherNo)}</td><td>[${escapeHtml(t.accountCode)}]</td><td>${escapeHtml((t.summary || '(空白)').slice(0, 35))}</td><td class="col-amount" style="cursor:pointer;color:#165dff;text-decoration:underline dotted;" data-f3-set-target="${amt}">${fmtAmount(amt)}</td></tr>`;
    }).join('')}
    </tbody></table></div>`;
}

function runF3() {
  const rows = getFilteredTransactions(); const direction = dom.f3Direction.value;
  const target = Number(dom.f3Target.value || 0); const tolerance = Number(dom.f3Tolerance.value || 0.01);
  const minAmt = Number(dom.f3MinAmt?.value || 0) || 0;
  const maxAmt = Number(dom.f3MaxAmt?.value || 0) || 0;
  const candidates = rows.filter((x) => {
    const amt = direction === 'debit' ? x.debit : x.credit;
    if (amt <= 0) return false;
    if (minAmt > 0 && amt < minAmt) return false;
    if (maxAmt > 0 && amt > maxAmt) return false;
    return true;
  }).map((x) => ({ id: x.id, amount: direction === 'debit' ? x.debit : x.credit }));
  AppState.pool.candidateIds = candidates.map((x) => x.id);

  if (candidates.length > 200) {
    AppState.pool.results = [];
    AppState.pool.groups = [];
    AppState.pool.ungrouped = AppState.pool.candidateIds.slice();
    AppState.pool.copyText = '';
    dom.f3Result.innerHTML = `<p class="danger">目前 ${candidates.length} 筆，超過 200 筆上限。請縮小科目範圍或設定金額上下限篩選。</p>`;
    renderF3CandidateList(rows);
    return;
  }

  if (candidates.length <= 30) {
    const r = bruteSubset(candidates, target, tolerance, 3000); AppState.pool.results = r.results;
    buildF3GroupsFromResults(rows);
    renderF3CandidateList(rows);
    renderF3Groups(rows, r.results.length ? `結果 ${r.results.length} 組｜耗時 ${r.elapsed}ms${r.interrupted ? '｜已中斷' : ''}` : `命中 0 組，候選 ${candidates.length} 筆仍可見。`);
    return;
  }

  const runWorkerFallback = () => {
    dom.runF3Btn.disabled = false;
    const r = bruteSubset(candidates, target, tolerance, 3000);
    AppState.pool.results = r.results;
    buildF3GroupsFromResults(rows);
    renderF3CandidateList(rows);
    renderF3Groups(rows, r.results.length ? `結果 ${r.results.length} 組｜耗時 ${r.elapsed}ms${r.interrupted ? '｜已中斷（同步模式）' : ''}` : `命中 0 組，候選 ${candidates.length} 筆仍可見。`);
  };
  try {
    if (!poolWorker) poolWorker = new Worker('pool.worker.js');
    dom.runF3Btn.disabled = true; dom.f3Result.innerHTML = `<p class="muted">計算中...</p>`;
    const workerTimeout = setTimeout(() => { poolWorker = null; runWorkerFallback(); toast('Worker 逾時，改用同步模式', 'WARN'); }, 8000);
    poolWorker.onerror = () => { clearTimeout(workerTimeout); poolWorker = null; runWorkerFallback(); };
    poolWorker.onmessage = (ev) => {
      clearTimeout(workerTimeout);
      AppState.pool.results = ev.data.results || []; dom.runF3Btn.disabled = false;
      buildF3GroupsFromResults(rows);
      renderF3CandidateList(rows);
      renderF3Groups(rows, AppState.pool.results.length ? `結果 ${AppState.pool.results.length} 組｜耗時 ${ev.data.elapsed}ms${ev.data.interrupted ? '｜已中斷' : ''}` : `命中 0 組，候選 ${candidates.length} 筆仍可見。`);
    };
    poolWorker.postMessage({ candidates, target, tolerance, timeLimit: 3000 });
  } catch {
    poolWorker = null;
    runWorkerFallback();
    toast('Worker 不可用，已改用同步模式', 'WARN');
  }
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
    if (amount >= 100000 && Math.round(amount) % 1000 === 0) results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'E', severity: 'INFO', accountCode: t.accountCode, transactionIds: [t.id], description: '整數金額大額' });
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

  // Rule I: 週末交易（週六/週日），依科目彙整
  const weekendByAcc = new Map();
  rows.forEach((t) => {
    const dow = t.date instanceof Date ? t.date.getDay() : -1;
    if (dow !== 0 && dow !== 6) return;
    if (!weekendByAcc.has(t.accountCode)) weekendByAcc.set(t.accountCode, []);
    weekendByAcc.get(t.accountCode).push(t);
  });
  weekendByAcc.forEach((txns, acc) => {
    results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'I', severity: 'INFO', accountCode: acc, transactionIds: txns.map((t) => t.id), description: `週末交易 ${txns.length} 筆 [${acc}]（${txns.map((t) => `${t.dateROC}(${['日','一','二','三','四','五','六'][t.date.getDay()]})`).slice(0, 5).join('、')}${txns.length > 5 ? '…' : ''}）` });
  });

  // Rule J: 月末集中（某月 ≥5 筆且月末3日佔比 >50%）
  const byMonthLast = new Map();
  rows.forEach((t) => {
    if (!(t.date instanceof Date)) return;
    const lastDay = new Date(t.date.getFullYear(), t.date.getMonth() + 1, 0).getDate();
    const isLastThree = t.date.getDate() >= lastDay - 2;
    const key = `${t.accountCode}||${t.periodROC}`;
    if (!byMonthLast.has(key)) byMonthLast.set(key, { total: 0, lastThree: 0, txnIds: [], acc: t.accountCode, period: t.periodROC });
    const x = byMonthLast.get(key);
    x.total++;
    if (isLastThree) { x.lastThree++; x.txnIds.push(t.id); }
  });
  byMonthLast.forEach((v) => {
    if (v.total >= 5 && v.lastThree / v.total > 0.5) {
      results.push({ id: Math.random().toString(36).slice(2, 10), rule: 'J', severity: 'WARN', accountCode: v.acc, transactionIds: v.txnIds, description: `月末集中：[${v.acc}] ${v.period} 月末3日佔 ${Math.round(v.lastThree / v.total * 100)}%（${v.lastThree}/${v.total} 筆）` });
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
    const reviewKey = `${g.rule}||${g.description}`;
    const isReviewed = AppState.anomaly.reviewed.has(reviewKey);
    const cardStyle = isReviewed ? 'margin:8px 0;opacity:0.5;' : 'margin:8px 0;';
    return `<details class="card" style="${cardStyle}" ${isReviewed ? '' : 'open'}><summary><strong>${idx + 1}. [${g.rule}] ${escapeHtml(g.description)}</strong>｜${g.severity}｜${list.length} 筆 <button data-f4-review="${escapeHtml(reviewKey)}" style="margin-left:12px;font-size:12px;">${isReviewed ? '取消已審閱' : '✓ 標記已審閱'}</button></summary>
      <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>科目</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th></tr></thead><tbody>
      ${list.map((t) => `<tr><td>${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td></tr>`).join('')}
      </tbody></table></div>
    </details>`;
  }).join('');
  const reviewedCount = Array.from(groupMap.keys()).filter((k) => AppState.anomaly.reviewed.has(k)).length;
  dom.f4Result.innerHTML = `<p class="muted">異常分組 ${groupMap.size} 組${reviewedCount ? `｜已審閱 ${reviewedCount} 組` : ''}</p>${cards}`;
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
  renderTxnList(dom.f5Result, matched, matched.length ? `命中 ${matched.length} 筆` : `命中 0 筆，仍顯示剩餘 ${remain.length} 筆`);
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
  renderTxnList(dom.f14List, remain, `未命中重複傳票 ${remain.length} 筆`);
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
      ? `<div style="margin:8px 0;"><button data-f18-copy-table="${escapeHtml(g.sourceKeyword || g.name)}" style="margin-bottom:6px;font-size:12px;">複製月份表格（貼入 Excel）</button><div class="table-wrap"><table><thead><tr><th>月份</th><th class="col-amount">金額</th><th class="col-amount">前月</th><th class="col-amount">變動率(%)</th><th>狀態</th></tr></thead><tbody>${trend.monthlyData.map((m) => `<tr><td>${m.period}</td><td class="col-amount">${fmtAmount(m.amount)}</td><td class="col-amount">${m.prevAmount == null ? '—' : fmtAmount(m.prevAmount)}</td><td class="col-amount">${m.changeRate == null ? '—' : `${m.changeRate.toFixed(1)}%`}</td><td>${m.changeRate == null ? '—' : m.flagged ? '<span class="danger">超標</span>' : '<span class="ok">正常</span>'}</td></tr>`).join('')}</tbody></table></div></div>`
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
// ---- F9 Account Memo & Request List ----
function loadMemoStorage() {
  try {
    const n = JSON.parse(localStorage.getItem('memo_notes') || '{}');
    const r = JSON.parse(localStorage.getItem('memo_rules') || '[]');
    AppState.memo.notes = (n && typeof n === 'object') ? n : {};
    AppState.memo.rules = Array.isArray(r) ? r : [];
  } catch { AppState.memo.notes = {}; AppState.memo.rules = []; }
}
function saveMemoStorage() {
  try {
    localStorage.setItem('memo_notes', JSON.stringify(AppState.memo.notes));
    localStorage.setItem('memo_rules', JSON.stringify(AppState.memo.rules));
  } catch { /* ignore */ }
}
function renderF9AccountSelects() {
  const allOpt = '<option value="">（全科目）</option>';
  const opts = Object.values(AppState.accounts).sort((a, b) => a.code.localeCompare(b.code))
    .map((a) => `<option value="${escapeHtml(a.code)}">[${escapeHtml(a.code)}] ${escapeHtml(a.name)}${a.normalSide ? `（${escapeHtml(a.normalSide)}）` : ''}</option>`).join('');
  if (dom.f9AccountSelect) dom.f9AccountSelect.innerHTML = allOpt + opts;
  if (dom.f9RuleAccount) dom.f9RuleAccount.innerHTML = allOpt + opts;
  // F13 只列實際科目（不含「全科目」選項）
  if (dom.f13AccountSelect) dom.f13AccountSelect.innerHTML = '<option value="">（請選擇科目）</option>' + opts;
}
function renderF9Memo() {
  const code = dom.f9AccountSelect?.value || '';
  if (dom.f9Memo) dom.f9Memo.value = AppState.memo.notes[code] || '';
  if (dom.f9MemoStatus) dom.f9MemoStatus.textContent = '';
}
function renderF9Rules() {
  if (!dom.f9RuleList) return;
  if (!AppState.memo.rules.length) { dom.f9RuleList.innerHTML = '<p class="muted">尚無規則。</p>'; return; }
  const freqLabel = (f) => f === 'monthly' ? '每月' : f === 'yearly' ? '每年' : '只查有無';
  dom.f9RuleList.innerHTML = `<div class="table-wrap"><table style="min-width:500px;"><thead><tr><th>科目</th><th>關鍵字</th><th>頻率</th><th>預期金額</th><th>操作</th></tr></thead><tbody>
    ${AppState.memo.rules.map((r, idx) => {
      const accName = r.accountCode ? `[${escapeHtml(r.accountCode)}] ${escapeHtml(AppState.accounts[r.accountCode]?.name || r.accountCode)}` : '（全科目）';
      return `<tr><td>${accName}</td><td>${escapeHtml(r.keyword)}</td><td>${freqLabel(r.frequency)}</td><td>${r.expectedAmount ? fmtAmount(r.expectedAmount) : '—'}</td><td><button data-f9-del-rule="${idx}">刪除</button></td></tr>`;
    }).join('')}
  </tbody></table></div>`;
}
function runF9Missing() {
  const results = [];
  for (const rule of AppState.memo.rules) {
    let rows = AppState.transactions;
    if (rule.accountCode) rows = rows.filter((t) => t.accountCode === rule.accountCode);
    if (rule.keyword) rows = rows.filter((t) => (t.summary || '').includes(rule.keyword));
    const accName = rule.accountCode ? (AppState.accounts[rule.accountCode]?.name || rule.accountCode) : '全科目';
    if (rule.frequency === 'any') {
      results.push({ rule, accName, missing: [], found: rows.length, note: rows.length ? `找到 ${rows.length} 筆` : '找不到任何相關分錄' });
      continue;
    }
    const byPeriod = new Map();
    rows.forEach((t) => {
      const period = rule.frequency === 'yearly' ? (t.periodROC || '').slice(0, t.periodROC?.indexOf('-') > 0 ? t.periodROC.indexOf('-') : 3) : t.periodROC;
      if (!period) return;
      if (!byPeriod.has(period)) byPeriod.set(period, []);
      byPeriod.get(period).push(t);
    });
    // Build full expected period range from ALL transactions of the account
    let allAccRows = AppState.transactions;
    if (rule.accountCode) allAccRows = allAccRows.filter((t) => t.accountCode === rule.accountCode);
    const allPeriods = new Set();
    allAccRows.forEach((t) => {
      const p = rule.frequency === 'yearly' ? (t.periodROC || '').slice(0, t.periodROC?.indexOf('-') > 0 ? t.periodROC.indexOf('-') : 3) : t.periodROC;
      if (p) allPeriods.add(p);
    });
    const sorted = Array.from(allPeriods).sort();
    const missing = [];
    if (sorted.length >= 2) {
      if (rule.frequency === 'monthly') {
        const parseP = (p) => { const [y, m] = p.split('-').map(Number); return { y, m }; };
        const { y: fy, m: fm } = parseP(sorted[0]);
        const { y: ly, m: lm } = parseP(sorted[sorted.length - 1]);
        let cy = fy, cm = fm;
        while (cy < ly || (cy === ly && cm <= lm)) {
          const key = `${cy}-${String(cm).padStart(2, '0')}`;
          if (!byPeriod.has(key)) missing.push(key);
          cm++; if (cm > 12) { cm = 1; cy++; }
        }
      } else {
        const fy = Number(sorted[0]); const ly = Number(sorted[sorted.length - 1]);
        for (let y = fy; y <= ly; y++) { if (!byPeriod.has(String(y))) missing.push(String(y)); }
      }
    }
    results.push({ rule, accName, missing, found: byPeriod.size, note: '' });
  }
  AppState.memo.missingResults = results;
  return results;
}
function renderF9Result(results) {
  if (!results.length) { dom.f9Result.innerHTML = '<p class="muted">尚無規則，請先新增規則。</p>'; return; }
  const cards = results.map((r, idx) => {
    const accLabel = r.rule.accountCode ? `[${escapeHtml(r.rule.accountCode)}] ${escapeHtml(r.accName)}` : '全科目';
    const freqLabel = r.rule.frequency === 'monthly' ? '每月' : r.rule.frequency === 'yearly' ? '每年' : '只查有無';
    const amtLabel = r.rule.expectedAmount ? `｜預期金額 ${fmtAmount(r.rule.expectedAmount)}` : '';
    const statusHtml = r.missing.length
      ? `<div class="danger" style="margin-top:6px;"><strong>缺少 ${r.missing.length} 個期間：</strong>${escapeHtml(r.missing.join('、'))}</div>`
      : r.note
        ? `<div class="warn" style="margin-top:6px;">${escapeHtml(r.note)}</div>`
        : `<div class="ok" style="margin-top:6px;">全部期間均有分錄 ✓（共 ${r.found} 個期間）</div>`;
    // Show memo for this account
    const memo = r.rule.accountCode ? (AppState.memo.notes[r.rule.accountCode] || '') : '';
    const memoHtml = memo ? `<div class="muted" style="margin-top:4px;font-style:italic;">備忘：${escapeHtml(memo)}</div>` : '';
    return `<details class="card" style="margin:8px 0;" open>
      <summary><strong>${idx + 1}. ${accLabel}</strong>｜${escapeHtml(r.rule.keyword || '(無關鍵字)')}｜${freqLabel}${amtLabel}</summary>
      ${memoHtml}${statusHtml}
    </details>`;
  }).join('');
  const totalMissing = results.reduce((s, r) => s + r.missing.length, 0);
  dom.f9Result.innerHTML = `<p class="muted">掃描 ${results.length} 條規則｜共缺少 <strong class="${totalMissing ? 'danger' : 'ok'}">${totalMissing}</strong> 個期間</p>${cards}`;
}
function buildF9RequestText() {
  const results = AppState.memo.missingResults || [];
  if (!results.length) return '（尚無掃描結果，請先按「掃描缺少分錄」）';
  const missGroups = results.filter((r) => r.missing.length > 0 || (r.rule.frequency === 'any' && !r.found));
  let text = `索取清單\n製表時間：${new Date().toLocaleString('zh-TW')}\n${'='.repeat(40)}\n\n`;
  if (!missGroups.length) { text += '✓ 所有期望分錄均已找到，無缺少項目。\n'; return text; }
  missGroups.forEach((r) => {
    const accLabel = r.rule.accountCode ? `[${r.rule.accountCode}] ${r.accName}` : '全科目';
    const freqLabel = r.rule.frequency === 'monthly' ? '（每月）' : r.rule.frequency === 'yearly' ? '（每年）' : '';
    text += `科目：${accLabel}\n`;
    text += `項目：${r.rule.keyword || '(無關鍵字)'}${freqLabel}`;
    if (r.rule.expectedAmount) text += `（預期金額約 ${fmtAmount(r.rule.expectedAmount)} 元）`;
    text += '\n';
    if (r.missing.length) text += `缺少期間：${r.missing.join('、')}\n`;
    else text += `缺少：找不到任何相關分錄\n`;
    const memo = r.rule.accountCode ? (AppState.memo.notes[r.rule.accountCode] || '') : '';
    if (memo) text += `備忘：${memo}\n`;
    text += '\n';
  });
  text += `${'='.repeat(40)}\n請儘速提供上述缺少之憑證及相關資料，謝謝。`;
  return text;
}
// ---- end F9 ----

// ---- F7 Trial Balance ----
function computeTrialBalance() {
  const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
  const rows = Object.values(AppState.accounts).sort((a, b) => a.code.localeCompare(b.code)).map((acc) => {
    const txns = (acc.transactionIds || []).map((id) => txMap.get(id)).filter(Boolean);
    const totalDebit = txns.reduce((s, t) => s + (t.debit || 0), 0);
    const totalCredit = txns.reduce((s, t) => s + (t.credit || 0), 0);
    const opening = acc.openingBalance ?? 0;
    let closing;
    if (acc.normalSide === '貸') closing = opening + totalCredit - totalDebit;
    else closing = opening + totalDebit - totalCredit;
    const lastTxn = txns[txns.length - 1];
    const lastBal = lastTxn ? lastTxn.balance : null;
    const balOk = lastBal === null || Math.abs(closing - lastBal) < 0.02;
    return { code: acc.code, name: acc.name, normalSide: acc.normalSide || '—', opening, totalDebit, totalCredit, closing, lastBal, balOk, txnCount: txns.length };
  });
  const grandDebit = rows.reduce((s, r) => s + r.totalDebit, 0);
  const grandCredit = rows.reduce((s, r) => s + r.totalCredit, 0);
  const balanced = Math.abs(grandDebit - grandCredit) < 0.02;
  return { rows, grandDebit, grandCredit, balanced };
}

function renderTrialBalance() {
  if (!AppState.transactions.length) return toast('請先上傳分類帳', 'WARN');
  const { rows, grandDebit, grandCredit, balanced } = computeTrialBalance();
  const balClass = balanced ? 'ok' : 'danger';
  const balLabel = balanced ? '借貸平衡 ✓' : `借貸不平衡！差額 ${fmtAmount(Math.abs(grandDebit - grandCredit))}`;
  const badRows = rows.filter((r) => !r.balOk);
  const warnHtml = badRows.length
    ? `<div class="warn" style="margin-bottom:8px;">⚠ 下列科目期末餘額與帳上最後一筆餘額不符：${badRows.map((r) => `[${escapeHtml(r.code)}] ${escapeHtml(r.name)}`).join('、')}</div>`
    : '';
  dom.f7Result.innerHTML = `${warnHtml}<div class="${balClass}" style="font-weight:600;margin-bottom:8px;">${escapeHtml(balLabel)}</div>
<div class="table-wrap"><table style="min-width:800px;"><thead><tr>
  <th>科目代碼</th><th>科目名稱</th><th>借/貸</th><th>筆數</th>
  <th class="col-amount">期初餘額</th>
  <th class="col-amount">本期借方</th>
  <th class="col-amount">本期貸方</th>
  <th class="col-amount">期末餘額（計算）</th>
  <th class="col-amount">帳上末筆餘額</th>
  <th>核對</th>
</tr></thead><tbody>
${rows.map((r) => `<tr>
  <td>${escapeHtml(r.code)}</td><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.normalSide)}</td><td>${r.txnCount}</td>
  <td class="col-amount">${fmtAmount(r.opening)}</td>
  <td class="col-amount">${fmtAmount(r.totalDebit)}</td>
  <td class="col-amount">${fmtAmount(r.totalCredit)}</td>
  <td class="col-amount">${fmtAmount(r.closing)}</td>
  <td class="col-amount">${r.lastBal === null ? '—' : fmtAmount(r.lastBal)}</td>
  <td>${r.balOk ? '<span class="ok">OK</span>' : '<span class="danger">不一致</span>'}</td>
</tr>`).join('')}
</tbody><tfoot><tr style="font-weight:600;background:#f7fafe;">
  <td colspan="5">合計</td>
  <td class="col-amount">${fmtAmount(grandDebit)}</td>
  <td class="col-amount">${fmtAmount(grandCredit)}</td>
  <td class="col-amount" colspan="3"><span class="${balClass}">${escapeHtml(balLabel)}</span></td>
</tr></tfoot></table></div>`;
}

// ---- end F7 ----

// ---- F10 傳票借貸平衡驗證 ----
function runF10() {
  const rows = getFilteredTransactions();
  if (!rows.length) return toast('請先上傳分類帳', 'WARN');
  const tol = Math.max(0, Number(dom.f10Tolerance?.value ?? 0.02) || 0.02);
  const byVoucher = new Map();
  rows.forEach((t) => {
    if (!byVoucher.has(t.voucherNo)) byVoucher.set(t.voucherNo, { debit: 0, credit: 0, accounts: new Set(), txns: [] });
    const v = byVoucher.get(t.voucherNo);
    v.debit += t.debit || 0;
    v.credit += t.credit || 0;
    v.accounts.add(`[${t.accountCode}] ${t.accountName}`);
    v.txns.push(t);
  });
  const unbalanced = []; const balanced = [];
  byVoucher.forEach((v, vno) => {
    const delta = v.debit - v.credit;
    const entry = { voucherNo: vno, debit: v.debit, credit: v.credit, delta, accounts: Array.from(v.accounts), txns: v.txns };
    if (Math.abs(delta) > tol) unbalanced.push(entry); else balanced.push(entry);
  });
  unbalanced.sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta));
  if (!unbalanced.length) {
    dom.f10Result.innerHTML = `<p class="ok" style="font-weight:600;">✓ 全部 ${byVoucher.size} 張傳票均借貸平衡（容差 ${tol}）。</p>`;
    return;
  }
  const cards = unbalanced.map((u, idx) => `<details class="card" style="margin:8px 0;" open>
    <summary><strong>${idx + 1}. 傳票 ${escapeHtml(u.voucherNo)}</strong>｜借 ${fmtAmount(u.debit)}｜貸 ${fmtAmount(u.credit)}｜<span class="danger">差額 ${fmtSigned(u.delta)}</span>｜${u.accounts.map(escapeHtml).join('、')}</summary>
    <div class="table-wrap" style="margin-top:8px;"><table><thead><tr><th>日期</th><th>科目</th><th>摘要</th><th class="col-amount">借方</th><th class="col-amount">貸方</th></tr></thead><tbody>
    ${u.txns.map((t) => `<tr><td>${escapeHtml(t.dateROC)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${t.debit ? fmtAmount(t.debit) : ''}</td><td class="col-amount">${t.credit ? fmtAmount(t.credit) : ''}</td></tr>`).join('')}
    </tbody></table></div>
  </details>`).join('');
  dom.f10Result.innerHTML = `<p class="muted">傳票數 ${byVoucher.size}｜<span class="ok">平衡 ${balanced.length}</span>｜<span class="danger">不平衡 ${unbalanced.length}</span>（容差 ${tol}）</p>${cards}`;
}
// ---- end F10 ----

// ---- F11 科目月份矩陣 ----
function runF11() {
  const rows = getFilteredTransactions();
  if (!rows.length) return toast('請先上傳分類帳', 'WARN');
  const field = dom.f11Field?.value || 'net';
  const accounts = Array.from(new Set(rows.map((t) => t.accountCode))).sort((a, b) => a.localeCompare(b));
  const months = Array.from(new Set(rows.map((t) => t.periodROC))).sort();
  const matrix = new Map();
  rows.forEach((t) => {
    const key = `${t.accountCode}||${t.periodROC}`;
    if (!matrix.has(key)) matrix.set(key, { net: 0, debit: 0, credit: 0, count: 0 });
    const x = matrix.get(key);
    x.net += getSignedAmount(t); x.debit += t.debit || 0; x.credit += t.credit || 0; x.count += 1;
  });
  const accInfo = new Map(Object.values(AppState.accounts).map((a) => [a.code, a]));
  const getValue = (acc, mon) => {
    const v = matrix.get(`${acc}||${mon}`);
    if (!v) return null;
    return field === 'count' ? v.count : field === 'debit' ? v.debit : field === 'credit' ? v.credit : v.net;
  };
  const allVals = [];
  accounts.forEach((acc) => months.forEach((mon) => { const v = getValue(acc, mon); if (v !== null) allVals.push(Math.abs(v)); }));
  const maxVal = Math.max(...allVals, 1);
  const cellStyle = (val) => {
    if (val === null || val === 0) return '';
    const alpha = (0.07 + Math.min(1, Math.abs(val) / maxVal) * 0.28).toFixed(2);
    if (field === 'count') return `background:rgba(22,93,255,${alpha});`;
    return val > 0 ? `background:rgba(35,120,4,${alpha});` : `background:rgba(207,19,34,${alpha});`;
  };
  const fmt = (val) => val === null ? '' : field === 'count' ? String(val) : fmtSigned(val);
  const fmtTotal = (val) => field === 'count' ? String(val) : fmtSigned(val);
  const header = `<tr><th>科目</th>${months.map((m) => `<th class="col-amount" style="white-space:nowrap;">${escapeHtml(m)}</th>`).join('')}<th class="col-amount">合計</th></tr>`;
  const bodyRows = accounts.map((acc) => {
    const info = accInfo.get(acc);
    const rowVals = months.map((m) => getValue(acc, m));
    const rowTotal = rowVals.reduce((s, v) => s + (v !== null ? v : 0), 0);
    const cells = rowVals.map((v, i) => `<td class="col-amount" style="${cellStyle(v)}">${fmt(v)}</td>`).join('');
    return `<tr><td style="white-space:nowrap;" title="${escapeHtml(info?.name || acc)}">[${escapeHtml(acc)}]<span class="muted"> ${escapeHtml((info?.name || '').slice(0, 6))}</span></td>${cells}<td class="col-amount" style="font-weight:600;">${fmtTotal(rowTotal)}</td></tr>`;
  }).join('');
  const colTotals = months.map((m) => accounts.reduce((s, acc) => { const v = getValue(acc, m); return s + (v !== null ? v : 0); }, 0));
  const grandTotal = colTotals.reduce((a, b) => a + b, 0);
  const footRow = `<tr style="font-weight:600;background:#f7fafe;"><td>合計</td>${colTotals.map((v) => `<td class="col-amount">${fmtTotal(v)}</td>`).join('')}<td class="col-amount">${fmtTotal(grandTotal)}</td></tr>`;
  const fieldLabel = { net: '淨額', debit: '借方合計', credit: '貸方合計', count: '筆數' }[field];
  dom.f11Result.innerHTML = `<p class="muted">${accounts.length} 個科目 × ${months.length} 個月份｜顯示：${fieldLabel}｜顏色深淺代表相對金額大小</p>
    <div class="table-wrap"><table style="min-width:${Math.max(600, months.length * 110 + 220)}px;">
      <thead>${header}</thead><tbody>${bodyRows}</tbody><tfoot>${footRow}</tfoot>
    </table></div>`;
}
// ---- end F11 ----

// ---- F13 科目餘額走勢 ----
function runF13() {
  const accCode = dom.f13AccountSelect?.value;
  if (!accCode) return toast('請選擇科目', 'WARN');
  const accInfo = AppState.accounts[accCode];
  if (!accInfo) return toast('找不到科目資料', 'WARN');
  const anomalyOnly = dom.f13AnomalyMode?.value === 'anomaly';
  const keyword = cleanText(dom.keywordInput.value);
  let rows = AppState.transactions.filter((t) => t.accountCode === accCode);
  if (keyword) rows = rows.filter((t) => t.summary.includes(keyword) || t.voucherNo.includes(keyword));
  rows = rows.slice().sort((a, b) => {
    const dt = a.date - b.date;
    if (dt !== 0) return dt;
    return (a.voucherNo || '').localeCompare(b.voucherNo || '');
  });
  if (!rows.length) { dom.f13Result.innerHTML = '<p class="muted">此科目無分錄資料（或關鍵字無命中）。</p>'; return; }
  const opening = accInfo.openingBalance ?? 0;
  const normalSide = accInfo.normalSide || '';
  let running = opening;
  const items = rows.map((t) => {
    if (normalSide === '貸') running = running + (t.credit || 0) - (t.debit || 0);
    else running = running + (t.debit || 0) - (t.credit || 0);
    const bookBal = t.balance || 0;
    const balOk = Math.abs(running - bookBal) < 0.02;
    const isAnomaly = !balOk || (normalSide === '借' && bookBal < 0) || (normalSide === '貸' && bookBal > 0);
    return { t, runningBal: running, bookBal, balOk, isAnomaly };
  });
  const display = anomalyOnly ? items.filter((x) => x.isAnomaly) : items;
  const anomalyCount = items.filter((x) => x.isAnomaly).length;
  if (!display.length) {
    dom.f13Result.innerHTML = `<p class="ok" style="font-weight:600;">✓ ${rows.length} 筆分錄，無餘額異常。期末餘額 ${fmtAmount(running)}。</p>`;
    return;
  }
  dom.f13Result.innerHTML = `
    <p class="muted">[${escapeHtml(accCode)}] ${escapeHtml(accInfo.name || '')}（${escapeHtml(normalSide || '—')}）｜期初 ${fmtAmount(opening)}｜${rows.length} 筆｜期末（計算）${fmtAmount(running)}｜${anomalyCount ? `<span class="danger">${anomalyCount} 筆餘額異常</span>` : '<span class="ok">餘額無異常</span>'}</p>
    <div class="table-wrap"><table style="min-width:860px;">
      <thead><tr><th>#</th><th class="col-date">日期</th><th class="col-voucher">傳票</th><th class="col-summary">摘要</th>
        <th class="col-amount">借方</th><th class="col-amount">貸方</th>
        <th class="col-amount">計算餘額</th><th class="col-amount">帳上餘額</th><th>核對</th>
      </tr></thead>
      <tbody>
      ${display.map(({ t, runningBal, bookBal, balOk, isAnomaly }, i) => `<tr style="${isAnomaly ? 'background:#fff2f0;' : ''}">
        <td>${rows.indexOf(t) + 1}</td>
        <td>${escapeHtml(t.dateROC)}</td><td>${escapeHtml(t.voucherNo)}</td>
        <td>${escapeHtml((t.summary || '(空白摘要)').slice(0, 40))}</td>
        <td class="col-amount">${t.debit ? fmtAmount(t.debit) : ''}</td>
        <td class="col-amount">${t.credit ? fmtAmount(t.credit) : ''}</td>
        <td class="col-amount" style="font-weight:600;">${fmtAmount(runningBal)}</td>
        <td class="col-amount">${fmtAmount(bookBal)}</td>
        <td>${balOk ? '<span class="ok">OK</span>' : `<span class="danger">差 ${fmtSigned(runningBal - bookBal)}</span>`}</td>
      </tr>`).join('')}
      </tbody>
    </table></div>`;
}
// ---- end F13 ----

// ---- F15 摘要頻率分析 ----
function runF15() {
  const rows = getFilteredTransactions();
  if (!rows.length) return toast('請先上傳分類帳', 'WARN');
  const sortMode = dom.f15Sort?.value || 'count_asc';
  const minCount = Math.max(1, Number(dom.f15MinCount?.value || 1) || 1);
  const maxCount = Number(dom.f15MaxCount?.value || 0) || 0;
  const freq = new Map();
  rows.forEach((t) => {
    const key = cleanText(t.summary) || '(空白摘要)';
    if (!freq.has(key)) freq.set(key, { summary: key, count: 0, totalDebit: 0, totalCredit: 0, accounts: new Set(), minDate: t.dateROC, maxDate: t.dateROC });
    const x = freq.get(key);
    x.count++; x.totalDebit += t.debit || 0; x.totalCredit += t.credit || 0;
    x.accounts.add(t.accountCode);
    if (t.dateROC < x.minDate) x.minDate = t.dateROC;
    if (t.dateROC > x.maxDate) x.maxDate = t.dateROC;
  });
  let entries = Array.from(freq.values()).filter((x) => x.count >= minCount && (maxCount <= 0 || x.count <= maxCount));
  if (sortMode === 'count_asc') entries.sort((a, b) => a.count - b.count || (b.totalDebit + b.totalCredit) - (a.totalDebit + a.totalCredit));
  else if (sortMode === 'amount_desc') entries.sort((a, b) => (b.totalDebit + b.totalCredit) - (a.totalDebit + a.totalCredit));
  else entries.sort((a, b) => b.count - a.count || (b.totalDebit + b.totalCredit) - (a.totalDebit + a.totalCredit));
  const show = entries.slice(0, 500);
  const singletons = Array.from(freq.values()).filter((x) => x.count === 1).length;
  dom.f15Result.innerHTML = `
    <p class="muted">共 ${freq.size} 種不同摘要｜孤筆（出現1次）<strong class="warn">${singletons}</strong> 種｜篩選後 ${entries.length} 種（顯示前 ${show.length}）</p>
    <div class="table-wrap"><table style="min-width:720px;">
      <thead><tr><th>#</th><th>摘要</th><th>次數</th><th class="col-amount">借方合計</th><th class="col-amount">貸方合計</th><th>涉及科目</th><th>日期範圍</th></tr></thead>
      <tbody>
      ${show.map((x, idx) => {
        const cStyle = x.count === 1 ? 'color:#d46b08;font-weight:600;' : x.count >= 12 ? 'color:#237804;font-weight:600;' : '';
        const dateRange = x.minDate === x.maxDate ? x.minDate : `${x.minDate}～${x.maxDate}`;
        return `<tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(x.summary)}</td>
          <td style="${cStyle}">${x.count}${x.count === 1 ? ' ⚠' : ''}</td>
          <td class="col-amount">${x.totalDebit ? fmtAmount(x.totalDebit) : '—'}</td>
          <td class="col-amount">${x.totalCredit ? fmtAmount(x.totalCredit) : '—'}</td>
          <td style="font-size:12px;">${Array.from(x.accounts).map(escapeHtml).join('、')}</td>
          <td style="font-size:12px;white-space:nowrap;">${escapeHtml(dateRange)}</td>
        </tr>`;
      }).join('')}
      </tbody>
    </table></div>`;
}
// ---- end F15 ----

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

function copyText(text, label = '已複製') {
  if (!text || !text.trim()) { toast('沒有可複製的內容', 'WARN'); return; }
  navigator.clipboard.writeText(text).then(
    () => toast(label),
    () => {
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
      const box = document.createElement('div');
      box.style.cssText = 'background:#fff;border-radius:12px;padding:16px;width:min(520px,92vw);display:flex;flex-direction:column;gap:10px;box-shadow:0 8px 32px rgba(0,0,0,.25);';
      box.innerHTML = '<strong style="font-size:15px;">複製文字（file:// 限制，請手動複製）</strong><p style="font-size:13px;color:#5f7692;margin:0;">請按 Ctrl+A 全選後再按 Ctrl+C 複製：</p>';
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.cssText = 'width:100%;height:200px;font:13px/1.5 monospace;resize:vertical;border:1px solid #d6dee8;border-radius:8px;padding:8px;';
      ta.readOnly = true;
      const closeBtn = document.createElement('button');
      closeBtn.textContent = '關閉';
      closeBtn.style.cssText = 'align-self:flex-end;padding:7px 20px;border:1px solid #d6dee8;border-radius:8px;background:#fff;cursor:pointer;';
      closeBtn.onclick = () => overlay.remove();
      overlay.addEventListener('click', (ev) => { if (ev.target === overlay) overlay.remove(); });
      box.append(ta, closeBtn);
      overlay.append(box);
      document.body.append(overlay);
      setTimeout(() => { ta.focus(); ta.select(); }, 50);
      toast('瀏覽器限制，請在彈出框手動複製', 'WARN');
    }
  );
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
  dom.f4Result.addEventListener('click', (e) => {
    const reviewKey = e.target?.dataset?.f4Review;
    if (!reviewKey) return;
    if (AppState.anomaly.reviewed.has(reviewKey)) AppState.anomaly.reviewed.delete(reviewKey);
    else AppState.anomaly.reviewed.add(reviewKey);
    // Re-render without re-running scan
    const rows = getFilteredTransactions();
    const txMap = new Map(rows.map((t) => [t.id, t]));
    const groupMap = new Map();
    AppState.anomaly.results.forEach((r) => {
      const k = `${r.rule}||${r.severity}||${r.description}`;
      if (!groupMap.has(k)) groupMap.set(k, { rule: r.rule, severity: r.severity, description: r.description, ids: [] });
      groupMap.get(k).ids.push(...r.transactionIds);
    });
    const cards = Array.from(groupMap.values()).map((g, idx) => {
      const ids = Array.from(new Set(g.ids));
      const list = ids.map((id) => txMap.get(id)).filter(Boolean);
      const rKey = `${g.rule}||${g.description}`;
      const isReviewed = AppState.anomaly.reviewed.has(rKey);
      const cardStyle = isReviewed ? 'margin:8px 0;opacity:0.5;' : 'margin:8px 0;';
      return `<details class="card" style="${cardStyle}" ${isReviewed ? '' : 'open'}><summary><strong>${idx + 1}. [${g.rule}] ${escapeHtml(g.description)}</strong>｜${g.severity}｜${list.length} 筆 <button data-f4-review="${escapeHtml(rKey)}" style="margin-left:12px;font-size:12px;">${isReviewed ? '取消已審閱' : '✓ 標記已審閱'}</button></summary>
        <div class="table-wrap"><table><thead><tr><th>傳票號碼</th><th>日期</th><th>科目</th><th>摘要</th><th class="col-amount">簽帳金額</th><th class="col-amount">餘額</th></tr></thead><tbody>
        ${list.map((t) => `<tr><td>${escapeHtml(t.voucherNo || '')}</td><td>${escapeHtml(t.dateROC)}</td><td>[${escapeHtml(t.accountCode)}] ${escapeHtml(t.accountName)}</td><td>${escapeHtml(t.summary || '(空白摘要)')}</td><td class="col-amount">${fmtSigned(getSignedAmount(t))}</td><td class="col-amount">${fmtAmount(t.balance)}</td></tr>`).join('')}
        </tbody></table></div>
      </details>`;
    }).join('');
    const reviewedCount = Array.from(groupMap.keys()).filter((k2) => AppState.anomaly.reviewed.has(k2)).length;
    dom.f4Result.innerHTML = `<p class="muted">異常分組 ${groupMap.size} 組${reviewedCount ? `｜已審閱 ${reviewedCount} 組` : ''}</p>${cards}`;
  });
  dom.clearF4ReviewedBtn?.addEventListener('click', () => {
    AppState.anomaly.reviewed.clear();
    if (AppState.anomaly.results.length) runF4();
    toast('已清除全部已審閱標記');
  });
  dom.runF5Btn.addEventListener('click', runF5);
  dom.runF6Btn.addEventListener('click', runF6);
  dom.runF14Btn.addEventListener('click', runF14);
  dom.runF18Btn.addEventListener('click', runF18);
  dom.copyF18TextBtn.addEventListener('click', () => copyText(AppState.trendAlert.copyText || '', '金額變動分組摘要已複製'));

  // F18 copy monthly table as tab-separated
  dom.f18Result?.addEventListener('click', (e) => {
    if (!e.target?.dataset?.f18CopyTable) return;
    const keyword = e.target.dataset.f18CopyTable;
    const hit = AppState.trendAlert.results.find((r) => r.keyword === keyword);
    if (!hit?.monthlyData?.length) return toast('無月份資料', 'WARN');
    const header = '月份\t金額\t前月金額\t變動率(%)';
    const rows = hit.monthlyData.map((m) => `${m.period}\t${m.amount}\t${m.prevAmount ?? ''}\t${m.changeRate != null ? m.changeRate.toFixed(2) : ''}`);
    copyText([header, ...rows].join('\n'), '月份趨勢表已複製（可貼入 Excel）');
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

  dom.f3Result.addEventListener('click', (e) => {
    const copy = e.target?.dataset?.f3Copy;
    if (copy) {
      copyText(AppState.pool.copyText || '', '數字池分組摘要已複製');
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

    const apply = e.target?.dataset?.f2Apply;
    if (apply) {
      const [debitCsv, creditCsv] = String(apply).split('|');
      const debitIds = (debitCsv || '').split(',').map((x) => cleanText(x)).filter(Boolean);
      const creditIds = (creditCsv || '').split(',').map((x) => cleanText(x)).filter(Boolean);
      if (!debitIds.length || !creditIds.length) return;
      const all = getFilteredTransactions();
      const txMap2 = new Map(all.map((t) => [t.id, t]));
      const debitTotal = sumByIds(txMap2, debitIds, 'debit');
      const creditTotal = sumByIds(txMap2, creditIds, 'credit');
      const tol2 = Number(dom.f2Tolerance.value || 0.01);
      if (!AppState.offset.manualMatches) AppState.offset.manualMatches = [];
      const alreadyExists = AppState.offset.manualMatches.some((m) => {
        return (m.debitIds || []).slice().sort().join(',') === debitIds.slice().sort().join(',') &&
               (m.creditIds || []).slice().sort().join(',') === creditIds.slice().sort().join(',');
      });
      if (alreadyExists) { toast('此沖帳組合已存在', 'WARN'); return; }
      AppState.offset.manualMatches.push({
        id: `m${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
        debitIds, creditIds, createdAt: Date.now(),
        reason: { kind: `${debitIds.length}↔${creditIds.length}`, delta: debitTotal - creditTotal, simPct: bestSummarySimilarityPct(txMap2, debitIds, creditIds), voucherMatch: false, dayDiff: null },
      });
      const usedIds = new Set(debitIds.concat(creditIds));
      AppState.offset.forcedUnmatchedIds = AppState.offset.forcedUnmatchedIds.filter((x) => !usedIds.has(x));
      runF2();
      toast(`已套用建議沖帳（借${debitIds.length}筆／貸${creditIds.length}筆）`);
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
    if (th !== undefined) {
      const n = Number(e.target.value);
      if (Number.isFinite(n) && n >= 0 && n <= 100) { AppState.offset.suggestThreshold = n; persistOffsetSetting('f2_suggestThreshold', n); }
      return;
    }
    const win = e.target?.dataset?.f2Win;
    if (win !== undefined) {
      const n = Number(e.target.value);
      if (Number.isFinite(n) && n >= 0) { AppState.offset.timeWindowDays = n; persistOffsetSetting('f2_timeWindowDays', n); }
      return;
    }
    const km = e.target?.dataset?.f2Kmax;
    if (km !== undefined) {
      const n = Number(e.target.value);
      if (Number.isFinite(n) && n >= 2) { AppState.offset.subsetMaxK = n; persistOffsetSetting('f2_subsetMaxK', n); }
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

  dom.copyF2TextBtn.addEventListener('click', () => copyText(AppState.offset.copyText || '', '未沖帳整理已複製'));

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
    if (e.target?.dataset?.f1BulkDelete) { showF1BulkDeleteModal(); return; }
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

  dom.mergeF1Btn?.addEventListener('click', () => {
    if (!AppState.grouping.groups.length) return toast('請先套用分組', 'WARN');
    syncF1AppliedEditsFromUI();
    const merged = new Map();
    const order = [];
    AppState.grouping.groups.forEach((g) => {
      const name = cleanText(g.name) || '其他';
      if (!merged.has(name)) { merged.set(name, { ...g, transactionIds: [] }); order.push(name); }
      g.transactionIds.forEach((id) => { if (!merged.get(name).transactionIds.includes(id)) merged.get(name).transactionIds.push(id); });
    });
    AppState.grouping.groups = order.map((name) => merged.get(name));
    renderF1Output();
    toast(`合併完成，現有 ${AppState.grouping.groups.length} 個群組`);
  });
  dom.sortF1Btn?.addEventListener('click', () => {
    if (!AppState.grouping.groups.length) return toast('請先套用分組', 'WARN');
    syncF1AppliedEditsFromUI();
    AppState.grouping.groups.sort((a, b) => cleanText(a.name).localeCompare(cleanText(b.name), 'zh-TW'));
    renderF1Output();
    toast('已依名稱排序');
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

  dom.copyF1TextBtn.addEventListener('click', () => copyText(AppState.grouping.copyText || '', '分組摘要已複製'));

  dom.todoToggleBtn.addEventListener('click', () => {
    dom.todoPanel.classList.toggle('collapsed');
    dom.todoToggleBtn.textContent = dom.todoPanel.classList.contains('collapsed') ? '展開' : '隱藏';
  });

  const autoF2 = debounce(() => { if (AppState.transactions.length) runF2(); });
  const autoF3 = debounce(() => { if (AppState.transactions.length) runF3(); }, 260);
  const autoF5 = debounce(() => {
    if (!AppState.transactions.length) return;
    const mode = dom.f5Mode.value;
    const query = cleanText(dom.f5Query.value);
    if ((mode === 'keyword' || mode === 'combo') && query.length < 2) return; // 靜默略過，避免打字中途彈 toast
    runF5();
  });
  const autoF6 = debounce(() => { if (AppState.transactions.length) runF6(); });
  const autoF18 = debounce(() => {
    if (!AppState.transactions.length) return;
    if (!splitGroupNames(dom.f18Keyword.value).length) return; // 靜默略過空關鍵字
    runF18();
  });

  dom.f2Tolerance.addEventListener('input', (e) => {
    persistOffsetSetting('f2_tolerance', e.target.value);
    autoF2();
  });
  dom.f3Direction.addEventListener('change', autoF3);
  dom.f3Target.addEventListener('input', autoF3);
  dom.f3Tolerance.addEventListener('input', autoF3);
  dom.f3MinAmt?.addEventListener('input', () => { if (AppState.transactions.length) renderF3CandidateList(getFilteredTransactions()); });
  dom.f3MaxAmt?.addEventListener('input', () => { if (AppState.transactions.length) renderF3CandidateList(getFilteredTransactions()); });
  dom.f3List.addEventListener('click', (e) => {
    const amt = e.target?.dataset?.f3SetTarget;
    if (amt == null) return;
    dom.f3Target.value = amt;
    toast(`目標金額已設為 ${fmtAmount(Number(amt))}`);
    if (AppState.transactions.length) runF3();
  });
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
    const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
    const rows = AppState.offset.pairs.map((p) => {
      // autoPairs shape: { debit: txn, credit: txn, ... }
      if (p?.debit?.id) return [p.debit.summary || '', p.debitTotal, p.credit?.summary || '', p.creditTotal, p.confidence];
      // manualPairs / manualMatches shape: { debitIds: [], creditIds: [], ... }
      const dText = (p.debitIds || []).map((id) => txMap.get(id)?.summary).filter(Boolean).join(' / ') || '(借方)';
      const cText = (p.creditIds || []).map((id) => txMap.get(id)?.summary).filter(Boolean).join(' / ') || '(貸方)';
      return [dText, p.debitTotal, cText, p.creditTotal, p.confidence];
    });
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

  dom.runF7Btn.addEventListener('click', renderTrialBalance);

  dom.exportF7Btn.addEventListener('click', () => {
    if (!AppState.transactions.length) return toast('請先上傳分類帳', 'WARN');
    const { rows, grandDebit, grandCredit, balanced } = computeTrialBalance();
    const csvRows = rows.map((r) => [r.code, r.name, r.normalSide, r.txnCount, r.opening, r.totalDebit, r.totalCredit, r.closing, r.lastBal ?? '', r.balOk ? 'OK' : '不一致']);
    csvRows.push(['合計', '', '', '', '', grandDebit, grandCredit, '', '', balanced ? '借貸平衡' : '借貸不平衡']);
    downloadCsv(`F7_試算表_${Date.now()}.csv`, ['科目代碼', '科目名稱', '借貸', '筆數', '期初餘額', '本期借方', '本期貸方', '期末餘額(計算)', '帳上末筆餘額', '核對'], csvRows);
  });

  dom.exportF3Btn.addEventListener('click', () => {
    const txMap = new Map(AppState.transactions.map((t) => [t.id, t]));
    const rows = [];
    AppState.pool.groups.forEach((g) => {
      const txns = g.transactionIds.map((id) => txMap.get(id)).filter(Boolean);
      txns.forEach((t) => rows.push([g.name, t.voucherNo, t.dateROC, t.accountCode, t.accountName, t.summary, t.debit, t.credit]));
    });
    AppState.pool.ungrouped.forEach((id) => { const t = txMap.get(id); if (t) rows.push(['(未分組)', t.voucherNo, t.dateROC, t.accountCode, t.accountName, t.summary, t.debit, t.credit]); });
    if (!rows.length) return toast('尚無數字池結果', 'WARN');
    downloadCsv(`F3_${Date.now()}.csv`, ['組名', '傳票號碼', '日期', '科目代碼', '科目名稱', '摘要', '借方金額', '貸方金額'], rows);
  });

  dom.exportF5Btn.addEventListener('click', () => {
    const results = AppState.crossLink.results;
    if (!results.length) return toast('尚無跨科目連結結果', 'WARN');
    const csvRows = results.map((t) => [t.voucherNo, t.dateROC, t.accountCode, t.accountName, t.summary, t.debit, t.credit, t.balance]);
    downloadCsv(`F5_${Date.now()}.csv`, ['傳票號碼', '日期', '科目代碼', '科目名稱', '摘要', '借方金額', '貸方金額', '餘額'], csvRows);
  });

  dom.exportF6Btn.addEventListener('click', () => {
    const results = AppState.gap.results;
    if (!results.length) return toast('尚無時間斷層結果', 'WARN');
    const csvRows = results.map((r) => [r.period, r.expected, r.actual, r.status === 'ok' ? 'OK' : r.status === 'missing' ? 'MISSING' : 'EXCESS']);
    downloadCsv(`F6_${Date.now()}.csv`, ['期間', '預期次數', '實際次數', '狀態'], csvRows);
  });

  // ---- Period filter ----
  const autoRebase = debounce(() => { if (AppState.transactions.length) renderBase(); }, 300);
  dom.periodFrom.addEventListener('input', autoRebase);
  dom.periodTo.addEventListener('input', autoRebase);

  // ---- F9 event handlers ----
  dom.f9AccountSelect.addEventListener('change', renderF9Memo);

  dom.f9SaveMemoBtn.addEventListener('click', () => {
    const code = dom.f9AccountSelect.value || '';
    const text = cleanText(dom.f9Memo.value);
    if (text) AppState.memo.notes[code] = text;
    else delete AppState.memo.notes[code];
    saveMemoStorage();
    if (dom.f9MemoStatus) dom.f9MemoStatus.textContent = `已儲存（${new Date().toLocaleTimeString('zh-TW')}）`;
    toast('備忘已儲存');
  });

  dom.f9AddRuleBtn.addEventListener('click', () => {
    const keyword = cleanText(dom.f9RuleKeyword.value);
    if (!keyword) return toast('請輸入摘要關鍵字', 'WARN');
    const rule = {
      id: `r${Date.now()}`,
      accountCode: dom.f9RuleAccount.value || '',
      keyword,
      frequency: dom.f9RuleFreq.value,
      expectedAmount: Number(dom.f9RuleAmount.value) || 0,
    };
    AppState.memo.rules.push(rule);
    saveMemoStorage();
    dom.f9RuleKeyword.value = ''; dom.f9RuleAmount.value = '';
    renderF9Rules();
    toast('規則已新增');
  });

  dom.f9ClearRulesBtn.addEventListener('click', () => {
    if (!AppState.memo.rules.length) return;
    AppState.memo.rules = [];
    saveMemoStorage();
    renderF9Rules();
    dom.f9Result.innerHTML = '';
    toast('已清除全部規則');
  });

  dom.f9RuleList.addEventListener('click', (e) => {
    const idxStr = e.target?.dataset?.f9DelRule;
    if (idxStr == null) return;
    AppState.memo.rules.splice(Number(idxStr), 1);
    saveMemoStorage();
    renderF9Rules();
  });

  dom.runF9Btn.addEventListener('click', () => {
    if (!AppState.transactions.length) return toast('請先上傳分類帳', 'WARN');
    if (!AppState.memo.rules.length) return toast('請先新增期望分錄規則', 'WARN');
    const results = runF9Missing();
    renderF9Result(results);
    renderF9Rules();
  });

  dom.copyF9RequestBtn.addEventListener('click', () => copyText(buildF9RequestText(), '索取清單已複製'));

  dom.exportF9Btn.addEventListener('click', () => {
    const results = AppState.memo.missingResults || [];
    if (!results.length) return toast('請先掃描缺少分錄', 'WARN');
    const csvRows = [];
    results.forEach((r) => {
      const accLabel = r.rule.accountCode || '全科目';
      if (r.missing.length) {
        r.missing.forEach((p) => csvRows.push([accLabel, r.accName, r.rule.keyword, r.rule.frequency, p, '缺少', r.rule.expectedAmount || '']));
      } else {
        csvRows.push([accLabel, r.accName, r.rule.keyword, r.rule.frequency, '', r.rule.frequency === 'any' && !r.found ? '找不到分錄' : 'OK', r.rule.expectedAmount || '']);
      }
    });
    downloadCsv(`F9_索取清單_${Date.now()}.csv`, ['科目代碼', '科目名稱', '關鍵字', '頻率', '缺少期間', '狀態', '預期金額'], csvRows);
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
  dom.copyTodoAllBtn.addEventListener('click', () => {
    copyText(AppState.todos.map((t) => `${t.voucherNo} ${t.content}`).join('\n'), '待辦已全部複製');
  });

  // ---- Overview sort & pagination ----
  dom.overviewList.addEventListener('click', (e) => {
    const sortCol = e.target?.closest?.('[data-ov-sort]')?.dataset?.ovSort;
    if (sortCol) {
      if (_ovSort.col === sortCol) _ovSort.dir = _ovSort.dir === 'asc' ? 'desc' : 'asc';
      else { _ovSort.col = sortCol; _ovSort.dir = 'asc'; }
      _ovPage = 1;
      renderOverviewTable(getFilteredTransactions());
      return;
    }
    const page = e.target?.dataset?.ovPage;
    if (page === 'prev') { _ovPage = Math.max(1, _ovPage - 1); renderOverviewTable(getFilteredTransactions()); return; }
    if (page === 'next') { _ovPage += 1; renderOverviewTable(getFilteredTransactions()); return; }
  });
  dom.overviewList.addEventListener('change', (e) => {
    const jump = e.target?.dataset?.ovJump;
    if (jump) { _ovPage = Number(jump) || 1; renderOverviewTable(getFilteredTransactions()); }
  });

  // ---- Clear local storage ----
  dom.clearLocalBtn?.addEventListener('click', () => {
    if (!confirm('確定清除所有本地儲存的備忘、規則、設定及跨期比對記錄？此操作無法復原。')) return;
    localStorage.clear();
    AppState.memo.notes = {}; AppState.memo.rules = []; AppState.memo.missingResults = [];
    if (dom.f9Result) dom.f9Result.innerHTML = '';
    if (dom.crossFileResult) dom.crossFileResult.innerHTML = '';
    renderF9Rules();
    toast('本地資料已全部清除');
  });
  // ---- 進階分析 ----
  dom.runF10Btn?.addEventListener('click', runF10);
  dom.runF11Btn?.addEventListener('click', runF11);
  dom.f11Field?.addEventListener('change', () => { if (AppState.transactions.length && dom.f11Result?.innerHTML.trim()) runF11(); });
  dom.runF13Btn?.addEventListener('click', runF13);
  dom.f13AccountSelect?.addEventListener('change', () => { if (dom.f13Result) dom.f13Result.innerHTML = ''; });
  dom.f13AnomalyMode?.addEventListener('change', () => { if (AppState.transactions.length && dom.f13Result?.innerHTML.trim()) runF13(); });
  dom.runF15Btn?.addEventListener('click', runF15);
  dom.f15Sort?.addEventListener('change', () => { if (AppState.transactions.length && dom.f15Result?.innerHTML.trim()) runF15(); });

  // ---- Workbench toggle ----
  dom.workbenchToggle?.addEventListener('click', () => {
    const body = dom.workbenchBody;
    if (!body) return;
    const hidden = body.style.display === 'none';
    body.style.display = hidden ? '' : 'none';
    dom.workbenchToggle.textContent = hidden ? '收起' : '展開';
  });

  // ---- F1 tabs: 分組 / 排除記錄 ----
  dom.f1TabGroup?.addEventListener('click', () => {
    dom.f1TabGroup.classList.add('primary');
    if (dom.f1TabExclusion) dom.f1TabExclusion.classList.remove('primary');
    if (dom.f1ExclusionView) dom.f1ExclusionView.style.display = 'none';
    if (dom.f1MainContent) dom.f1MainContent.style.display = '';
    if (dom.f1KeywordRulesPanel) dom.f1KeywordRulesPanel.style.display = 'none';
  });
  dom.f1TabExclusion?.addEventListener('click', () => {
    if (dom.f1TabExclusion) dom.f1TabExclusion.classList.add('primary');
    if (dom.f1TabGroup) dom.f1TabGroup.classList.remove('primary');
    if (dom.f1ExclusionView) dom.f1ExclusionView.style.display = '';
    if (dom.f1MainContent) dom.f1MainContent.style.display = 'none';
    if (dom.f1KeywordRulesPanel) dom.f1KeywordRulesPanel.style.display = 'none';
    renderExclusionViewer();
  });

  // ---- F1 keyword rules panel ----
  dom.f1KeywordRulesBtn?.addEventListener('click', () => {
    const panel = dom.f1KeywordRulesPanel;
    if (!panel) return;
    panel.style.display = panel.style.display === 'none' ? '' : 'none';
    if (panel.style.display !== 'none') renderKeywordRuleList();
  });

  dom.kwAddRuleBtn?.addEventListener('click', () => {
    const name = dom.kwRuleName?.value.trim();
    const keywords = (dom.kwRuleKeywords?.value || '').split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    const excludeWords = (dom.kwRuleExclude?.value || '').split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    const targetGroup = dom.kwRuleTarget?.value.trim();
    const priority = parseInt(dom.kwRulePriority?.value) || 10;
    if (!name || !keywords.length || !targetGroup) { toast('請填寫規則名稱、關鍵字、目標分組'); return; }
    if (!AppState.grouping.keywordRules) AppState.grouping.keywordRules = [];
    const rule = { id: Math.random().toString(36).slice(2), name, enabled: true, keywords, excludeWords, targetGroup, priority, matchMode: 'contains-any' };
    AppState.grouping.keywordRules.push(rule);
    saveKeywordRules(AppState.grouping.keywordRules);
    // Recompute keywordGroupName and effectiveGroupName for all existing transactions
    for (const t of AppState.transactions) {
      t.keywordGroupName = applyKeywordRules(t.summaryNormalized, AppState.grouping.keywordRules) || '';
      t.effectiveGroupName = t.manualGroupName || t.keywordGroupName || t.defaultGroupName || t.summaryNormalized || t.rawSummary;
    }
    renderKeywordRuleList();
    renderWorkbench();
    if (dom.kwRuleName) dom.kwRuleName.value = '';
    if (dom.kwRuleKeywords) dom.kwRuleKeywords.value = '';
    if (dom.kwRuleExclude) dom.kwRuleExclude.value = '';
    if (dom.kwRuleTarget) dom.kwRuleTarget.value = '';
    toast('規則已新增，分組已即時更新');
  });
}

function init() {
  loadUserSettings();
  loadMemoStorage();
  AppState.grouping.keywordRules = loadKeywordRules();
  bindEvents();
  renderTodos();
  renderF9Rules();
  toast(`LibDetector: Fuse ${hasFuse ? 'OK' : 'fallback'} / Decimal ${hasDecimal ? 'OK' : 'missing'}`);
}

init();
