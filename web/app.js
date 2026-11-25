/* O2O 红包核销看板 · 前端 */
const state = {
  tab: 'overview',
  meta: null,
  summary: null,
  dims: null,
  check: 'C1',
  checkQuery: { page: 1, size: 30, status: '', period: '', channel: '', q: '' },
  moduleNo: 1,
  moduleQuery: { page: 1, size: 30, period: '', channel: '', q: '' },
};

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];

async function api(path, opts) {
  const res = await fetch(path, opts);
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`);
  return data;
}

function toast(msg, ms = 2600) {
  const el = $('#toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.add('hidden'), ms);
}

const fmt = (v, digits = 2) => {
  if (v === null || v === undefined || v === '') return '—';
  const n = Number(v);
  if (Number.isNaN(n)) return v;
  return Number.isInteger(n) ? n.toLocaleString('zh-CN') : n.toLocaleString('zh-CN', { maximumFractionDigits: digits });
};
const typeText = t => ({ retail: '零售红包', order: '下单红包' }[t] || (t || '—'));
const statusText = s => ({ PASS: '合规', FAIL: '异常', WARN: '预警' }[s] || s);
/* Excel 来的内容（门店名/型号/错误信息等）渲染前一律转义，防底表注入 */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* ---------- 导航 ---------- */
$$('.nav-item').forEach(btn => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));

function switchTab(tab) {
  state.tab = tab;
  $$('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  $$('.pane').forEach(p => p.classList.toggle('hidden', p.id !== `pane-${tab}`));
  const titles = {
    overview: ['概览', '四项核对的整体结论与异常清单'],
    import: ['数据导入', '拖入底表即可入库，重复导入按主键覆盖'],
    data: ['数据明细', '下单、发放、政策、期初四类原始数据'],
    checks: ['核销结果', '逐条核对记录，可按状态与月份筛选'],
    rules: ['核对口径', '判定规则与参数，改口径改配置文件'],
  };
  const [t, s] = titles[tab] || ['—', ''];
  $('#pageTitle').textContent = t;
  $('#pageSub').textContent = s;

  if (tab === 'overview') loadOverview();
  if (tab === 'data') { loadDims().then(loadModule); }
  if (tab === 'checks') { loadDims().then(loadChecks); }
  if (tab === 'rules') loadRules();
}

/* ---------- 概览 ---------- */
async function loadOverview() {
  const [meta, summary] = await Promise.all([api('/api/meta'), api('/api/summary')]);
  state.meta = meta;
  state.summary = summary;

  const dbParts = (meta.db_file || '—').split('/');
  $('#sideDb').textContent = `我的库：${dbParts.length > 1 ? dbParts.slice(-2).join('/') : dbParts[0]}`;
  $('#sideDb').title = meta.db_file || '';
  $('#sideRun').textContent = meta.run_id ? `批次 #${meta.run_id}` : '尚未核对';
  $('#sideSpan').textContent = meta.periods_span || '—';
  $('#overviewSpan').textContent = meta.periods_span ? `核对区间 ${meta.periods_span}` : '';

  const c = meta.counts;
  const totals = summary.checks.reduce((a, s) => ({
    total: a.total + s.total, fail: a.fail + s.fail, warn: a.warn + s.warn,
  }), { total: 0, fail: 0, warn: 0 });

  $('#kpis').innerHTML = [
    { label: '核对项数', value: fmt(totals.total), foot: `批次 #${meta.run_id || '—'}`, cls: '' },
    { label: '异常', value: fmt(totals.fail), foot: '超出政策标准或超发', cls: 'fail' },
    { label: '预警', value: fmt(totals.warn), foot: '数据不完备，需补数复核', cls: '' },
    { label: '下单明细', value: fmt(c.order), foot: '行', cls: '' },
    { label: '发放明细', value: fmt(c.issue), foot: '行', cls: '' },
    { label: '政策 / 期初', value: `${fmt(c.policy)} / ${fmt(c.opening)}`, foot: '条', cls: '' },
  ].map(k => `<div class="kpi ${k.cls}">
      <div class="k-label">${k.label}</div>
      <div class="k-value">${k.value}</div>
      <div class="k-foot">${k.foot}</div>
    </div>`).join('');

  $('#summaryTable tbody').innerHTML = summary.checks.map(s => `
    <tr>
      <td><b>${s.code}</b> ${s.name}</td>
      <td class="num">${fmt(s.total)}</td>
      <td class="num">${fmt(s.pass)}</td>
      <td class="num ${s.fail ? 'diff-fail' : ''}">${fmt(s.fail)}</td>
      <td class="num">${fmt(s.warn)}</td>
      <td class="num">${fmt(s.diff)}</td>
      <td><button class="subtab" data-jump="${s.code}">查看</button></td>
    </tr>`).join('') || '<tr><td colspan="7" class="empty">还没有核对结果</td></tr>';

  $$('#summaryTable [data-jump]').forEach(b => b.addEventListener('click', () => {
    state.check = b.dataset.jump;
    $$('#checkTabs .subtab').forEach(x => x.classList.toggle('active', x.dataset.check === state.check));
    switchTab('checks');
  }));

  let fails;
  try {
    fails = await api('/api/results?status=FAIL&size=30&page=1');
  } catch (err) {
    $('#failTable tbody').innerHTML = `<tr><td colspan="8" class="empty">加载异常清单失败：${esc(err.message)}</td></tr>`;
    return;
  }
  $('#failTable tbody').innerHTML = fails.rows.length ? fails.rows.map(r => `
    <tr>
      <td>${esc(r.period) || '—'}</td>
      <td>${esc(r.store_name) || '—'}</td>
      <td>${esc(r.model_code) || '—'}</td>
      <td>${esc(r.left_label ? r.left_label.replace(/（.*）/, '') : typeText(r.red_type)) || '—'}</td>
      <td class="num">${fmt(r.left_value)}</td>
      <td class="num">${fmt(r.right_value)}</td>
      <td class="num diff-fail">${r.diff === null ? '—' : fmt(r.diff)}</td>
      <td>${esc(r.note) || '—'}</td>
    </tr>`).join('') : '<tr><td colspan="8" class="empty">没有异常，全部合规</td></tr>';
}

/* ---------- 维度 ---------- */
async function loadDims() {
  if (state.dims) return state.dims;
  state.dims = await api('/api/dims');
  const opts = (arr, val = 'value', text = 'label') =>
    arr.map(x => {
      const v = typeof x === 'string' ? x : x[val];
      const t = typeof x === 'string' ? x : x[text];
      return `<option value="${esc(v)}">${esc(t)}</option>`;
    }).join('');

  $('#dPeriod').innerHTML = '<option value="">全部月份</option>' + opts(state.dims.periods);
  $('#cPeriod').innerHTML = '<option value="">全部月份</option>' + opts(state.dims.periods);
  $('#dChannel').innerHTML = '<option value="">全部渠道</option>' + opts(state.dims.channels);
  $('#cChannel').innerHTML = '<option value="">全部渠道</option>' + opts(state.dims.channels);
  return state.dims;
}

/* ---------- 数据明细 ---------- */
async function loadModules() {
  const { modules } = await api('/api/modules');
  $('#moduleTabs').innerHTML = modules.map(m =>
    `<button class="subtab ${m.no === state.moduleNo ? 'active' : ''}" data-no="${m.no}">${m.name}</button>`).join('');
  $$('#moduleTabs .subtab').forEach(b => b.addEventListener('click', () => {
    state.moduleNo = Number(b.dataset.no);
    state.moduleQuery.page = 1;
    $$('#moduleTabs .subtab').forEach(x => x.classList.toggle('active', Number(x.dataset.no) === state.moduleNo));
    loadModule();
  }));
}

async function loadModule() {
  if (!$('#moduleTabs').children.length) await loadModules();
  const qs = new URLSearchParams({ no: state.moduleNo, ...state.moduleQuery });
  const data = await api(`/api/module?${qs}`);
  const table = $('#moduleTable');
  table.innerHTML = `
    <thead><tr>${data.columns.map(c => `<th class="${/qty|amount|value/.test(c.key) ? 'num' : ''}">${c.label}</th>`).join('')}</tr></thead>
    <tbody>${data.rows.length ? data.rows.map(r => `<tr>${data.columns.map(c => {
      const v = r[c.key];
      const cls = /qty|amount|value/.test(c.key) ? 'num' : '';
      const text = c.key === 'red_type' ? typeText(v) : (typeof v === 'number' ? fmt(v) : (v ?? '—'));
      return `<td class="${cls}">${esc(text)}</td>`;
    }).join('')}</tr>`).join('') : `<tr><td colspan="${data.columns.length}" class="empty">没有数据</td></tr>`}</tbody>`;

  renderPager($('#dPager'), data, q => { Object.assign(state.moduleQuery, q); loadModule(); });
}

function renderPager(el, data, onChange) {
  const pages = Math.max(1, Math.ceil(data.total / data.size));
  el.innerHTML = `
    <div>共 ${fmt(data.total)} 条 · 第 ${data.page}/${pages} 页</div>
    <div class="pages">
      <button data-p="1" ${data.page === 1 ? 'disabled' : ''}>首页</button>
      <button data-p="${data.page - 1}" ${data.page === 1 ? 'disabled' : ''}>上一页</button>
      <button data-p="${data.page + 1}" ${data.page >= pages ? 'disabled' : ''}>下一页</button>
      <button data-p="${pages}" ${data.page >= pages ? 'disabled' : ''}>末页</button>
    </div>`;
  $$('button[data-p]', el).forEach(b => b.addEventListener('click', () => {
    const p = Number(b.dataset.p);
    if (p >= 1 && p <= pages) onChange({ page: p });
  }));
}

/* ---------- 核销结果 ---------- */
async function loadChecks() {
  const qs = new URLSearchParams({ check: state.check, ...state.checkQuery });
  const data = await api(`/api/results?${qs}`);
  const table = $('#checkTable');
  table.innerHTML = `
    <thead><tr><th>月份</th><th>渠道</th><th>门店</th><th>型号</th><th>类型</th>
      <th class="num">${esc(data.left_label)}</th><th class="num">${esc(data.right_label)}</th>
      <th class="num">差异</th><th>状态</th><th>说明</th></tr></thead>
    <tbody>${data.rows.length ? data.rows.map(r => `
      <tr>
        <td>${esc(r.period) || '—'}</td>
        <td>${esc(r.channel_code) || '—'}</td>
        <td>${esc(r.store_name) || '—'}</td>
        <td>${esc(r.model_code) || '—'}</td>
        <td>${esc(typeText(r.red_type))}</td>
        <td class="num">${fmt(r.left_value)}</td>
        <td class="num">${fmt(r.right_value)}</td>
        <td class="num ${r.status === 'FAIL' ? 'diff-fail' : ''}">${r.diff === null ? '—' : fmt(r.diff)}</td>
        <td><span class="badge ${r.status}">${statusText(r.status)}</span></td>
        <td>${esc(r.note)}</td>
      </tr>`).join('') : '<tr><td colspan="10" class="empty">没有记录</td></tr>'}</tbody>`;
  renderPager($('#cPager'), data, q => { Object.assign(state.checkQuery, q); loadChecks(); });
}

$$('#checkTabs .subtab').forEach(b => b.addEventListener('click', () => {
  state.check = b.dataset.check;
  state.checkQuery.page = 1;
  $$('#checkTabs .subtab').forEach(x => x.classList.toggle('active', x.dataset.check === state.check));
  loadChecks();
}));

$('#cStatus').addEventListener('change', e => { state.checkQuery.status = e.target.value; state.checkQuery.page = 1; loadChecks(); });
$('#cPeriod').addEventListener('change', e => { state.checkQuery.period = e.target.value; state.checkQuery.page = 1; loadChecks(); });
$('#cChannel').addEventListener('change', e => { state.checkQuery.channel = e.target.value; state.checkQuery.page = 1; loadChecks(); });
let cTimer;
$('#cSearch').addEventListener('input', e => {
  clearTimeout(cTimer);
  cTimer = setTimeout(() => { state.checkQuery.q = e.target.value.trim(); state.checkQuery.page = 1; loadChecks(); }, 350);
});

$('#dPeriod').addEventListener('change', e => { state.moduleQuery.period = e.target.value; state.moduleQuery.page = 1; loadModule(); });
$('#dChannel').addEventListener('change', e => { state.moduleQuery.channel = e.target.value; state.moduleQuery.page = 1; loadModule(); });
let dTimer;
$('#dSearch').addEventListener('input', e => {
  clearTimeout(dTimer);
  dTimer = setTimeout(() => { state.moduleQuery.q = e.target.value.trim(); state.moduleQuery.page = 1; loadModule(); }, 350);
});

/* ---------- 导出 ---------- */
function download(url) {
  const a = document.createElement('a');
  a.href = url;
  a.click();
}

$('#btnCExport').addEventListener('click', () => {
  const qs = new URLSearchParams({ type: 'results', check: state.check, ...state.checkQuery });
  download(`/api/export?${qs}`);
});
$('#btnDExport').addEventListener('click', () => {
  const qs = new URLSearchParams({ type: 'module', no: state.moduleNo, ...state.moduleQuery });
  download(`/api/export?${qs}`);
});
$('#btnExportTop').addEventListener('click', () => {
  const qs = new URLSearchParams({ type: 'results', check: state.check, ...state.checkQuery });
  download(`/api/export?${qs}`);
});

/* ---------- 重跑核对 ---------- */
$('#btnRecheck').addEventListener('click', async e => {
  const btn = e.target;
  btn.disabled = true;
  btn.textContent = '核对中…';
  try {
    const r = await api('/api/recheck', { method: 'POST' });
    toast(`核对完成：异常 ${r.summary.checks.reduce((a, s) => a + s.fail, 0)} 项`);
    state.dims = null;
    switchTab('overview');
  } catch (err) {
    toast(`核对失败：${err.message}`, 4000);
  } finally {
    btn.disabled = false;
    btn.textContent = '重跑核对';
  }
});

/* ---------- 上传 ---------- */
const drop = $('#drop');
const fileInput = $('#fileInput');
drop.addEventListener('click', () => fileInput.click());
['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.add('over');
}));
['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => {
  e.preventDefault(); drop.classList.remove('over');
}));
drop.addEventListener('drop', e => uploadFiles(e.dataTransfer.files));
fileInput.addEventListener('change', e => uploadFiles(e.target.files));

async function uploadFiles(files) {
  if (!files || !files.length) return;
  for (const file of files) {
    toast(`正在导入 ${file.name} …`, 6000);
    try {
      const res = await fetch(`/api/import?name=${encodeURIComponent(file.name)}`, {
        method: 'POST', body: file,
      });
      const data = await res.json();
      if (!res.ok || data.error) throw new Error(data.error || `HTTP ${res.status}`);
      showImportResult(file.name, data);
      toast(`已导入 ${data.written} 行（失败 ${data.failed} 行），建议重跑核对`);
    } catch (err) {
      toast(`导入失败：${err.message}`, 4500);
      showImportResult(file.name, { error: err.message });
    }
  }
  state.dims = null;
  loadOverview();
}

function showImportResult(name, data) {
  const box = $('#importResult');
  box.classList.remove('hidden');
  if (data.error) {
    box.innerHTML = `<div class="ir-head">${esc(name)}</div><div class="ir-body err">导入失败：${esc(data.error)}</div>`;
    return;
  }
  const rows = (data.sheets || []).map(s => `
    <div>${esc(s.sheet)}（${esc(s.type)}）：写入 ${s.written} 行，失败 ${s.failed} 行
    ${s.errors && s.errors.length ? `<div class="err">${s.errors.map(esc).join('<br>')}</div>` : ''}</div>`).join('');
  box.innerHTML = `<div class="ir-head">${esc(name)} · 批次 #${data.batch_id}</div>
    <div class="ir-body">共写入 ${data.written} 行，失败 ${data.failed} 行<div style="margin-top:6px">${rows}</div></div>`;
}

/* ---------- 口径 ---------- */
async function loadRules() {
  const meta = state.meta || await api('/api/meta');
  $('#ruleJudge').textContent = meta.judge_from;
  $('#ruleBase').textContent = meta.baseline;
}

/* ---------- 启动 ---------- */
(async function init() {
  try {
    await loadOverview();
  } catch (e) {
    toast(`接口异常：${e.message}`, 5000);
  }
})();
