/* =============================================================================
 * 工资个税核算 · 前端
 * 无框架原生实现，兼容 Chrome / Safari / Edge 桌面端与移动端
 * ============================================================================= */
(function () {
  'use strict';
  var T = window.TaxCore;
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };
  var Store = window.StoreLocal;
  var Excel = window.ExcelBrowser;
  var MAX = Store.MAX_RECORDS;

  var state = {
    records: [],
    settings: { carry: T.defaultCarry(), yearBase: {} },
    draft: null,
    computed: [],
    result: null,
    preview: null,
    exportScope: 'month',
    importMode: 'merge',
    selected: new Set(),
    compareOpen: false,
    /* 派生字段的计算模式：'manual' = 手动，缺省 = 自动。
       模式只能由点击胶囊切换，不根据数值差异推断，也不随记录存档 */
    mode: {}
  };

  /* ------------------------------------------------------------- 工具 */
  function el(tag, cls) { var e = document.createElement(tag); if (cls) e.className = cls; return e; }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fmt(v) {
    var n = Math.round((Number(v) || 0) * 100) / 100;
    try { return n.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
    catch (e) { return n.toFixed(2); }
  }
  function yen(v) { return '¥' + fmt(v); }
  function nowYM(d) { d = d || new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0'); }
  function parseNumInput(v) {
    if (v === '' || v === null || v === undefined) return 0;
    return T.num(String(v).replace(/[,，\s¥￥%]/g, ''));
  }
  /** 百分比字段：界面按 % 显示，内部存小数 */
  function fmtPct(v) {
    var n = (Number(v) || 0) * 100;
    var s = Math.abs(n - Math.round(n)) < 1e-9 ? String(Math.round(n)) : n.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
    return s;
  }
  /** 派生字段需要回填输入框时，按字段类型格式化 */
  function fmtFieldValue(f, d) {
    var v = d[f.key];
    if (f.kind === 'ratio') return fmtPct(v);
    return fmt(v);
  }
  function isRatio(f) { return f.kind === 'ratio'; }
  /** 派生字段是否处于手动模式（缺省即自动，默认自动） */
  function isManual(key) { return state.mode[key] === 'manual'; }
  /** 当前草稿里处于手动模式的派生字段集合，供 compute / previewOne 显式告知后端口径 */
  function manualFlags() {
    var m = {};
    T.DERIVED_KEYS.forEach(function (k) { if (isManual(k)) m[k] = true; });
    return m;
  }
  /** 派生字段的公式值（只读探针，不改动草稿） */
  function formulaValue(key) {
    if (!state.draft) return 0;
    return T.derivedValues(state.draft)[key];
  }

  var toastTimer;
  function toast(msg, kind) {
    var t = $('#toast');
    t.textContent = msg;
    t.className = 'toast show' + (kind === 'err' ? ' err' : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.className = 'toast'; }, kind === 'err' ? 4200 : 2600);
  }

  var modalCb = null;
  function confirmDialog(title, bodyHtml, okText, cb) {
    $('#modal-title').textContent = title;
    $('#modal-body').innerHTML = bodyHtml;
    $('#modal-ok').textContent = okText;
    modalCb = cb;
    $('#modal').hidden = false;
  }
  function closeModal() { $('#modal').hidden = true; modalCb = null; }

  /* ------------------------------------------------------------- 本地存储 */
  function loadData() {
    var d = Store.load();
    state.records = d.records || [];
    state.settings = d.settings || { carry: T.defaultCarry(), yearBase: {} };
    if (!state.settings.carry) state.settings.carry = T.defaultCarry();
    if (!state.settings.yearBase) state.settings.yearBase = {};
    updateChip();
    return true;
  }

  function updateChip() {
    $('#archive-chip').textContent = '已存档 ' + state.records.length + ' / ' + MAX + ' 月';
  }

  /* ------------------------------------------------------------- 月份 */
  function latestPayMonth() {
    var best = null;
    state.records.forEach(function (r) {
      if (!best || T.monthKey(r.payMonth) > T.monthKey(best)) best = r.payMonth;
    });
    return best;
  }
  function prevRecord(payMonth) {
    var k = T.monthKey(payMonth), best = null;
    state.records.forEach(function (r) {
      if (T.monthKey(r.payMonth) < k && (!best || T.monthKey(r.payMonth) > T.monthKey(best.payMonth))) best = r;
    });
    return best;
  }
  function nextRecord(payMonth) {
    var k = T.monthKey(payMonth), best = null;
    state.records.forEach(function (r) {
      if (T.monthKey(r.payMonth) > k && (!best || T.monthKey(r.payMonth) < T.monthKey(best.payMonth))) best = r;
    });
    return best;
  }

  function newDraftFor(payMonth) {
    var d = T.emptyRecord();
    d.payMonth = T.fmtMonth(payMonth);
    d.periodMonth = T.nextMonth(d.payMonth, -1) || d.payMonth;
    var prev = prevRecord(d.payMonth);
    if (prev) {
      T.FIELDS.forEach(function (f) {
        if (state.settings.carry[f.key]) d[f.key] = prev[f.key];
      });
    }
    return d;
  }

  function editMonth(payMonth) {
    payMonth = T.fmtMonth(payMonth);
    var rec = state.records.filter(function (r) { return T.fmtMonth(r.payMonth) === payMonth; })[0];
    /* 模式不存档：每次载入月份一律回到自动，只有点击才会变成手动 */
    state.mode = {};
    if (rec) {
      state.draft = Object.assign(T.emptyRecord(), rec);
    } else {
      state.draft = newDraftFor(payMonth);
    }
    state.draft.payMonth = payMonth;
    syncMonthSelects();
    recompute();
    syncFields();
  }

  function syncMonthSelects() {
    var d = state.draft; if (!d) return;
    var p = T.parseMonth(d.periodMonth) || T.parseMonth(d.payMonth);
    var q = T.parseMonth(d.payMonth) || p;
    setSel('#period-year', p.year); setSel('#period-month', p.month);
    setSel('#pay-year', q.year); setSel('#pay-month', q.month);
  }

  function syncFields() {
    var d = state.draft; if (!d) return;
    T.FIELDS.forEach(function (f) {
      var inp = $('#f-' + f.key);
      if (!inp) return;
      if (kindOf(f) === 'derived') { writeDerived(f); return; }
      inp.value = fmtFieldValue(f, d);
    });
    refreshDerivedModes();
    var saved = state.records.some(function (r) { return T.fmtMonth(r.payMonth) === d.payMonth; });
    var tag = $('#edit-state');
    tag.textContent = saved ? '已存档 · 第 ' + (T.parseMonth(d.payMonth) || {}).month + ' 期' : '未存档（保存后生效）';
    tag.className = 'state-tag' + (saved ? ' saved' : '');
    $('#btn-delete-one').disabled = !saved;
  }

  /** 字段种类：显式 kind 优先，未标注则视为纯录入 */
  function kindOf(f) { return f.kind || 'input'; }

  /** 把派生字段的当前值写进输入框（公式值在 recompute 后已回写到草稿） */
  function writeDerived(f) {
    var inp = $('#f-' + f.key);
    if (!inp || !state.draft) return;
    if (inp === document.activeElement) return;   // 正在编辑时不打断用户
    inp.value = fmt(T.num(state.draft[f.key]));
  }

  /** 同步派生字段的模式胶囊外观与输入框可写状态 */
  function refreshDerivedModes() {
    T.DERIVED_KEYS.forEach(function (k) {
      var inp = $('#f-' + k);
      var pill = $('#mode-' + k);
      if (!inp || !pill) return;
      var manual = isManual(k);
      inp.readOnly = !manual;
      inp.classList.toggle('is-derived', !manual);
      inp.classList.toggle('is-manual', manual);
      pill.textContent = manual ? '手动' : '自动';
      pill.classList.toggle('on', manual);
      pill.setAttribute('aria-pressed', manual ? 'true' : 'false');
      pill.title = manual ? '点击切换为自动模式（数值将回到公式结果）' : '点击切换为手动模式（可自行填写）';
      pill.setAttribute('aria-label', '计算模式：' + (manual ? '手动' : '自动') + '，点击切换');
      var wrap = inp.closest('.field');
      if (wrap) wrap.classList.toggle('is-manual', manual);
    });
  }

  /** 点击胶囊切换自动 / 手动；这是改变模式的唯一入口 */
  function toggleDerivedMode(f) {
    if (!state.draft) return;
    var toManual = !isManual(f.key);
    if (toManual) {
      state.mode[f.key] = 'manual';
    } else {
      delete state.mode[f.key];
      /* 切回自动：立刻用公式值覆盖，手填值不再保留 */
      state.draft[f.key] = formulaValue(f.key);
    }
    /* 先让输入框失焦：否则 syncFields 会认为「正在编辑」而跳过回填，
       切回自动后框里仍是手填的旧值 */
    var box = $('#f-' + f.key);
    if (box && box === document.activeElement) box.blur();
    recompute();
    syncFields();
    toast(f.label + ' 已调整为' + (toManual ? '手动' : '自动') + '模式');
    if (toManual) {
      var inp = $('#f-' + f.key);
      if (inp) { inp.focus(); inp.select(); }
    }
  }

  /* ------------------------------------------------------------- 计算 */
  function recompute() {
    var list = state.records.map(function (r) { return Object.assign({}, r); });
    var d = state.draft;
    /* 派生字段先按模式定稿，再交给 compute：
       自动 = 永远等于公式值（不比对、不推断，改任何输入项都立刻跟着变）
       手动 = 完全以用户填的值为准，公式不参与 */
    if (d) {
      var fv = T.derivedValues(d);
      T.DERIVED_KEYS.forEach(function (k) {
        if (!isManual(k)) d[k] = fv[k];
      });
    }
    if (d) {
      var clone = Object.assign({}, d, { id: d.payMonth });
      clone.__manual = manualFlags();     // 显式告知后端哪些字段是手动，避免被公式冲掉
      var i = -1;
      for (var k = 0; k < list.length; k++) if (T.fmtMonth(list[k].payMonth) === d.payMonth) { i = k; break; }
      if (i >= 0) list[i] = clone; else list.push(clone);
    }
    list.forEach(function (r) { r.id = r.id || T.fmtMonth(r.payMonth); });
    state.computed = T.compute(list, { yearBase: state.settings.yearBase || {} });
    state.result = d ? state.computed.filter(function (c) { return T.fmtMonth(c.payMonth) === d.payMonth; })[0] : null;
    renderResult();
    renderHistory();
    renderExportOptions();
  }

  function renderResult() {
    var r = state.result, d = state.draft;
    if (!r || !d) return;
    var pv = T.previewOne(d, null, manualFlags());
    state.preview = pv;
    $('#sum-gross').textContent = yen(pv.gross);
    $('#sum-social').textContent = yen(pv.socialTotal);
    /* 派生字段同步进输入框，并刷新锁定外观。
       正在编辑的框不要格式化，否则光标会跳到末尾、盖掉用户输入 */
    T.DERIVED_KEYS.forEach(function (k) {
      var inp = $('#f-' + k);
      if (!inp || inp === document.activeElement) return;
      if (!isManual(k)) inp.value = fmt(T.num(d[k]));
    });
    refreshDerivedModes();
    $('#r-tax').textContent = yen(r.tax);
    var br = T.lookup(T.NEW_BRACKETS, r.cumTaxable);
    /* 适用预扣率与速算扣除数并入这一行副标题，不再在结果表里单占一格 */
    var pctTxt = (br.rate * 100).toFixed(0) + '%' + (br.level ? '（第' + br.level + '档）' : '');
    $('#r-tax-sub').textContent = '适用预扣率 ' + pctTxt + ' · 速算扣除数 ' + fmt(br.deduct) +
      ' · 第 ' + r.seq + ' 期' + (r.tax === 0 && r.cumTaxable > 0 ? ' · 本月无需预扣' : '');
    $('#r-net').textContent = yen(r.netPay);
    $('#r-current').textContent = yen(r.currentPay);
    $('#r-gross').textContent = yen(r.gross);
    $('#r-social').textContent = yen(r.socialTotal);
    $('#r-deduct').textContent = yen(r.totalDeduction);
    $('#r-old').textContent = yen(r.oldTax);
    renderAnomalies(r);
    renderCum(r);
  }

  function renderAnomalies(r) {
    var box = $('#anomaly-box');
    box.innerHTML = '';
    var errs = (r.anomalies || []).filter(function (a) { return a.level === 'error'; });
    var warns = (r.anomalies || []).filter(function (a) { return a.level === 'warn'; });
    var infos = (r.anomalies || []).filter(function (a) { return a.level === 'info'; });
    if (!errs.length && !warns.length && !infos.length) {
      var ok = el('div', 'anomaly-ok');
      ok.innerHTML = '<span class="ico" style="background:var(--ok);width:16px;height:16px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;color:#fff;font-size:11px">✓</span><span>本月数据与计算逻辑一致，未发现异常</span>';
      box.appendChild(ok);
      return;
    }
    var list = el('div', 'anomaly-list');
    errs.concat(warns).concat(infos).forEach(function (a) {
      var div = el('div', 'anomaly ' + a.level);
      var ico = el('span', 'ico');
      ico.textContent = a.level === 'error' ? '!' : (a.level === 'warn' ? '!' : 'i');
      var span = el('span'); span.textContent = a.message;
      div.appendChild(ico); div.appendChild(span);
      list.appendChild(div);
    });
    box.appendChild(list);
  }

  function renderCum(r) {
    var prevPaid = T.round2(r.cumTaxPaid - r.tax);
    var br = T.lookup(T.NEW_BRACKETS, r.cumTaxable);
    var rows = [
      ['累计收入（税前工资累计）', yen(r.cumIncome)],
      ['累计减除费用（5000 × ' + r.seq + ' 期）', yen(r.cumBasic)],
      ['累计社保公积金', yen(r.cumSocial)],
      ['累计专项附加扣除', yen(r.cumSpecial)],
      ['累计应纳税所得额', yen(r.cumTaxable)],
      ['适用预扣率 / 速算扣除数', (br.rate * 100).toFixed(0) + '% / ¥' + fmt(br.deduct)],
      ['累计应缴税额', yen(r.cumTaxDue)],
      ['上月累计已缴税额', yen(prevPaid)],
      ['本月应预扣预缴个税', yen(r.tax)],
      ['累计已缴税额', yen(r.cumTaxPaid)]
    ];
    var dl = el('dl', 'kv');
    rows.forEach(function (row, i) {
      var dt = el('dt'); dt.textContent = row[0];
      var dd = el('dd'); dd.textContent = row[1];
      if (i >= 8) { dt.className = 'tot'; dd.className = 'tot'; }
      dl.appendChild(dt); dl.appendChild(dd);
    });
    var box = $('#cum-detail'); box.innerHTML = '';
    box.appendChild(dl);
    /* 本月社保公积金拆项，便于核对与公司工资条是否一致 */
    var dl2 = el('dl', 'kv');
    [
      ['本月个人社保', yen(r.personalSocial)],
      ['本月个人公积金', yen(r.personalFund)],
      ['本月社保公积金合计', yen(r.socialTotal)]
    ].forEach(function (row, i) {
      var dt = el('dt'); dt.textContent = row[0];
      var dd = el('dd'); dd.textContent = row[1];
      if (i === 2) { dt.className = 'tot'; dd.className = 'tot'; }
      dl2.appendChild(dt); dl2.appendChild(dd);
    });
    var sub = el('div', 'cum-sub');
    sub.appendChild(dl2);
    box.appendChild(sub);
  }

  /* ------------------------------------------------------------- 明细 */
  function renderHistory() {
    var list = $('#hist-list'), empty = $('#hist-empty');
    list.innerHTML = '';
    var ordered = state.computed.slice().sort(function (a, b) { return T.monthKey(b.payMonth) - T.monthKey(a.payMonth); });
    $('#hist-count').textContent = '共 ' + ordered.length + ' 个月 · 上限 ' + MAX + ' 个月';
    if (!ordered.length) { empty.hidden = false; return; }
    empty.hidden = true;

    ordered.forEach(function (c) {
      var bad = (c.anomalies || []).filter(function (a) { return a.level === 'error'; });
      var warn = (c.anomalies || []).filter(function (a) { return a.level === 'warn'; });
      var row = el('div', 'hrow');
      row.setAttribute('data-pm', c.payMonth);

      var lab = el('label', 'check');
      lab.addEventListener('click', function (e) { e.stopPropagation(); });
      var cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = state.selected.has(c.payMonth);
      cb.setAttribute('data-pm', c.payMonth);
      cb.addEventListener('click', function (e) { e.stopPropagation(); });
      cb.addEventListener('change', function () {
        if (cb.checked) state.selected.add(c.payMonth); else state.selected.delete(c.payMonth);
        updateSelCount();
      });
      lab.appendChild(cb);

      var main = el('div', 'hrow-main');
      var title = el('div', 'hrow-title');
      title.textContent = c.periodMonth + ' 所属 → ' + c.payMonth + ' 发放';
      var seq = el('span', 'seq'); seq.textContent = '第 ' + c.seq + ' 期';
      title.appendChild(seq);
      var meta = el('div', 'hrow-meta');
      meta.textContent = '应发 ' + yen(c.gross) + ' · 社保公积金 ' + yen(c.socialTotal) + ' · 到手 ' + yen(c.netPay);
      var badge = el('span', 'badge ' + (bad.length ? 'err' : (warn.length ? 'warn' : 'ok')));
      badge.textContent = bad.length ? bad.length + ' 项异常' : (warn.length ? warn.length + ' 项提醒' : '正常');
      meta.appendChild(badge);
      main.appendChild(title); main.appendChild(meta);

      var num = el('div', 'hrow-num');
      var b = el('b'); b.textContent = yen(c.tax);
      var s = el('span'); s.textContent = '本月个税';
      num.appendChild(b); num.appendChild(s);

      row.appendChild(lab); row.appendChild(main); row.appendChild(num);
      row.addEventListener('click', function () {
        editMonth(c.payMonth);
        switchView('calc');
      });
      list.appendChild(row);
    });
    updateSelCount();
  }
  function findRec(payMonth) {
    return state.records.filter(function (r) { return T.fmtMonth(r.payMonth) === payMonth; })[0] || T.emptyRecord();
  }
  function updateSelCount() {
    $('#sel-count').textContent = '已选 ' + state.selected.size + ' 条';
    $('#btn-bulk-delete').disabled = state.selected.size === 0;
    var all = $('#chk-all');
    all.checked = state.selected.size > 0 && state.selected.size === state.records.length;
    all.indeterminate = state.selected.size > 0 && state.selected.size < state.records.length;
  }

  /* ------------------------------------------------------------- 导出选项 */
  function renderExportOptions() {
    var sel = $('#export-value');
    var cur = sel.value;
    sel.innerHTML = '';
    var hint = $('#export-hint');
    var row = $('#scope-value-row');
    var ordered = state.computed.slice().sort(function (a, b) { return T.monthKey(a.payMonth) - T.monthKey(b.payMonth); });
    if (state.exportScope === 'month') {
      row.hidden = false;
      $('label[for=export-value]').textContent = '选择月份';
      ordered.forEach(function (c) {
        var o = el('option'); o.value = c.payMonth; o.textContent = c.payMonth + ' 发放（' + c.periodMonth + ' 所属）';
        sel.appendChild(o);
      });
      sel.value = cur && ordered.some(function (c) { return c.payMonth === cur; }) ? cur : (ordered.length ? ordered[ordered.length - 1].payMonth : '');
      hint.textContent = '按月：只导出所选发放月这一条记录。';
    } else if (state.exportScope === 'year') {
      row.hidden = false;
      $('label[for=export-value]').textContent = '选择年份';
      var years = {};
      ordered.forEach(function (c) { years[String(c.payMonth).slice(0, 4)] = (years[String(c.payMonth).slice(0, 4)] || 0) + 1; });
      Object.keys(years).sort().reverse().forEach(function (y) {
        var o = el('option'); o.value = y; o.textContent = y + ' 年（' + years[y] + ' 个月）';
        sel.appendChild(o);
      });
      sel.value = cur && years[cur] ? cur : (Object.keys(years)[0] || '');
      hint.textContent = '按年：导出该年度已存档的全部月份；不满一年则导出该年现有全部。';
    } else {
      row.hidden = true;
      hint.textContent = '全部：导出当前存档的 ' + ordered.length + ' 个月（上限 ' + MAX + ' 个月，更早的会被自动覆盖）。';
    }
    $('#btn-export').disabled = ordered.length === 0;
    /* 选项是动态填充的，自定义下拉的按钮文字要跟着刷新 */
    if (sel._ddRefresh) sel._ddRefresh();
  }

  /* ------------------------------------------------------------- 视图 */
  function switchView(v) {
    state.view = v;
    $$('.tab').forEach(function (t) { t.classList.toggle('is-active', t.getAttribute('data-view') === v); });
    ['calc', 'history', 'io', 'settings'].forEach(function (k) {
      var s = $('#view-' + k);
      if (k === v) { s.hidden = false; } else { s.hidden = true; }
    });
    window.scrollTo({ top: 0, behavior: 'auto' });
  }

  /* ------------------------------------------------------------- 构建 UI */
  function buildYearOptions() {
    var y0 = new Date().getFullYear();
    [['#period-year', '#period-month'], ['#pay-year', '#pay-month']].forEach(function (pair) {
      var ys = $(pair[0]), ms = $(pair[1]);
      /* 框内只放纯数字，年 / 月 由框外的静态文字承担（.sel-unit） */
      for (var y = y0 - 10; y <= y0 + 2; y++) {
        var o = el('option'); o.value = String(y); o.textContent = String(y); ys.appendChild(o);
      }
      for (var m = 1; m <= 12; m++) {
        var o2 = el('option'); o2.value = String(m); o2.textContent = String(m); ms.appendChild(o2);
      }
    });
  }

  /* ------------------------------------------------- 字段分组编排（界面层）
     编排原则：
     1) 每组自成一格网，成员数 = 列数的整数倍（4→4列、3→3列、2→2列、单项→整行），
        因此不会出现「一行只剩一个、旁边大片留白」的残行；
     2) 每格只有「单行标题 + 单行输入框」，高度恒定，同行输入框严格平齐——
        「手动改」移到标题行右端，字段说明收进 ⓘ 浮层，两者都不再占第二行。
        原先压板式的灰字说明长短不一，正是同行其他格子被拉高、底部堆空白的主因。 */
  var SUBGROUPS = [
    { id: 'income-base', group: 'income', title: '基本构成', cols: 4,
      keys: ['base', 'housingSubsidy', 'performance', 'phoneAllowance'] },
    { id: 'income-other', group: 'income', title: '其他', cols: 2, note: '含奖金、开门红等临时收入',
      keys: ['otherPreTax', 'missedPunch'] },
    { id: 'income-leave', group: 'income', title: '请假', cols: 3,
      keys: ['leaveHours', 'standardHours', 'leaveDeduction'] },
    { id: 'pre-social', group: 'preTax', title: '个人社保', cols: 3,
      keys: ['socialRatio', 'socialRate', 'personalSocial'] },
    { id: 'pre-fund', group: 'preTax', title: '个人公积金', cols: 2,
      keys: ['fundRate', 'personalFund'] },
    { id: 'pre-special', group: 'preTax', title: '其他扣除', cols: 1,
      keys: ['specialDeduction'] },
    { id: 'post-all', group: 'postTax', cols: 2,
      keys: ['utility', 'rent', 'otherDeduction', 'advancePaid'] },
    { id: 'cmp-all', group: 'compare', cols: 3,
      keys: ['psGross', 'psTax', 'psNet'] }
  ];

  function fieldByKey(k) {
    return T.FIELDS.filter(function (f) { return f.key === k; })[0];
  }

  function closeAllHints() {
    $$('.hint-pop.show').forEach(function (p) { p.classList.remove('show'); });
  }

  /** 标题行右端的 ⓘ 说明：点开显示，避免说明文字把格子撑高 */
  function buildHint(f) {
    var b = el('button', 'hint-btn'); b.type = 'button';
    b.textContent = 'i';
    b.setAttribute('aria-label', f.label + '：' + f.hint);
    b.title = f.hint;
    var pop = el('span', 'hint-pop'); pop.textContent = f.hint;
    b.appendChild(pop);
    b.addEventListener('click', function (e) {
      e.stopPropagation();
      var was = pop.classList.contains('show');
      closeAllHints();
      if (!was) pop.classList.add('show');
    });
    return b;
  }

  /* ---------------------------------------------- 自定义下拉（日式风格）
   原生 <select> 的弹出列表是操作系统默认样式，无法用 CSS 定制，
   与整页的和纸风格脱节。这里保留原生 select 作为数据源（隐藏），
   视觉层换成自定义按钮 + 可滚动列表面板：
   - 列表懒渲染（每次打开时读 select.options，动态填充的导出月份天然支持）
   - 打开时自动把当前选中项滚到视野中央
   - 底部空间不足时向上弹
   - 键盘可用：Enter/Space/↓ 打开，↑↓ 移动，Enter 确认，Esc 关闭
   - 选中后向原生 select 派发真实 change 事件，所有既有监听照常工作 */
  function closeAllDd() {
    $$('.dd-panel.show').forEach(function (p) {
      p.classList.remove('show');
      var b = p.parentElement.querySelector('.dd-btn');
      if (b) b.setAttribute('aria-expanded', 'false');
    });
  }

  function enhanceSelect(sel) {
    if (sel.closest('.dd')) return;
    var dd = el('div', 'dd');
    sel.parentNode.insertBefore(dd, sel);
    dd.appendChild(sel);
    sel.classList.add('dd-src');

    var btn = el('button', 'dd-btn'); btn.type = 'button';
    btn.setAttribute('aria-haspopup', 'listbox');
    btn.setAttribute('aria-expanded', 'false');
    var panel = el('div', 'dd-panel');
    var list = el('div', 'dd-list');
    list.setAttribute('role', 'listbox');
    panel.appendChild(list);
    dd.appendChild(btn); dd.appendChild(panel);

    function label() {
      var o = sel.options[sel.selectedIndex];
      btn.textContent = o ? o.text : '';
    }
    function buildList() {
      list.innerHTML = '';
      [].slice.call(sel.options).forEach(function (o) {
        var it = el('div', 'dd-item' + (String(o.value) === String(sel.value) ? ' sel hl' : ''));
        it.textContent = o.text;
        it.setAttribute('role', 'option');
        it.setAttribute('data-value', o.value);
        it.addEventListener('click', function () { pick(o.value); });
        list.appendChild(it);
      });
    }
    function open() {
      closeAllDd();
      buildList();
      panel.classList.add('show');
      btn.setAttribute('aria-expanded', 'true');
      var s = list.querySelector('.dd-item.sel');
      /* 手动把选中项滚到列表垂直居中：scrollIntoView 依赖布局时机，
         实测会偏差约 30px，直接算 scrollTop 最稳 */
      if (s) list.scrollTop = s.offsetTop + s.offsetHeight / 2 - list.clientHeight / 2;
      var need = Math.min(list.scrollHeight, 238);
      panel.classList.toggle('up', btn.getBoundingClientRect().bottom + need + 12 > window.innerHeight);
    }
    function close() {
      panel.classList.remove('show');
      btn.setAttribute('aria-expanded', 'false');
    }
    function pick(v) {
      if (String(sel.value) !== String(v)) {
        sel.value = v;
        sel.dispatchEvent(new Event('change', { bubbles: true }));
      }
      label(); close(); btn.focus();
    }
    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      if (panel.classList.contains('show')) close(); else open();
    });
    btn.addEventListener('keydown', function (e) {
      var opened = panel.classList.contains('show');
      if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (!opened) { open(); return; }
      }
      if (!opened) return;
      var items = [].slice.call(list.querySelectorAll('.dd-item'));
      var idx = items.findIndex(function (x) { return x.classList.contains('hl'); });
      if (e.key === 'ArrowDown') { e.preventDefault(); idx = Math.min(idx + 1, items.length - 1); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); idx = Math.max(idx - 1, 0); }
      else if (e.key === 'Enter') {
        e.preventDefault();
        if (items[idx]) pick(items[idx].getAttribute('data-value'));
        return;
      } else if (e.key === 'Escape') { close(); return; }
      else return;
      items.forEach(function (x) { x.classList.remove('hl'); });
      if (items[idx]) { items[idx].classList.add('hl'); if (items[idx].scrollIntoView) items[idx].scrollIntoView({ block: 'nearest' }); }
    });
    sel.addEventListener('change', label);
    /* 外部代码直接改 select.value 后调用它刷新按钮文字 */
    sel._ddRefresh = label;
    label();
  }

  function setSel(id, v) {
    var s = $(id);
    if (!s) return;
    s.value = v;
    if (s._ddRefresh) s._ddRefresh();
  }

  function buildFieldGrids() {
    SUBGROUPS.forEach(function (sg) {
      var box = $('#grid-' + sg.group); if (!box) return;
      var wrap = el('div', 'subgroup');
      if (sg.title || sg.note) {
        var head = el('div', 'sub-head');
        if (sg.title) { var t = el('span', 'sub-title'); t.textContent = sg.title; head.appendChild(t); }
        if (sg.note) { var n = el('em', 'sub-note'); n.textContent = sg.note; head.appendChild(n); }
        head.appendChild(el('i', 'sub-rule'));
        wrap.appendChild(head);
      }
      var g = el('div', 'g g' + sg.cols);
      sg.keys.forEach(function (k) {
        var f = fieldByKey(k);
        if (f) g.appendChild(buildOneField(f));
      });
      wrap.appendChild(g);
      box.appendChild(wrap);
    });
  }

  function buildOneField(f) {
    var kind = kindOf(f);
    var wrap = el('div', 'field');
    if (kind === 'derived') wrap.classList.add('field-derived');

    /* ── 第一行：标题 + 右侧小控件（高度恒定） ── */
    var head = el('div', 'field-head');
    var lab = el('label'); lab.setAttribute('for', 'f-' + f.key); lab.textContent = f.label;
    head.appendChild(lab);
    if (kind === 'derived') {
      /* 一个可点胶囊，直接显示当前模式；点它才切换，不会自己跳 */
      var pill = el('button', 'mode-pill');
      pill.type = 'button';
      pill.id = 'mode-' + f.key;
      pill.textContent = '自动';
      pill.title = '点击切换为手动模式（可自行填写）';
      pill.setAttribute('aria-pressed', 'false');
      pill.addEventListener('click', function () { toggleDerivedMode(f); });
      head.appendChild(pill);
    }
    if (f.hint) head.appendChild(buildHint(f));
    wrap.appendChild(head);

    /* ── 第二行：输入框（¥ / % / 时 后缀用绝对定位，不影响行高） ── */
    var unit = el('div', 'unit');
    if (kind === 'ratio') unit.classList.add('unit-pct');
    else if (kind === 'hours') unit.classList.add('unit-hr');

    var input = document.createElement('input');
    input.type = 'text'; input.autocomplete = 'off';
    input.inputMode = 'decimal';
    input.id = 'f-' + f.key; input.setAttribute('data-key', f.key);
    input.setAttribute('data-kind', kind);

    input.addEventListener('input', function () {
      if (!state.draft) return;
      if (kind === 'derived' && !isManual(f.key)) return;   // 自动模式下不可写
      var v = parseNumInput(input.value);
      state.draft[f.key] = kind === 'ratio' ? v / 100 : v;
      recompute();
    });
    /* 派生字段即使通过键盘、粘贴、右键菜单改动也要拦住 */
    ['keydown', 'paste', 'cut', 'drop'].forEach(function (ev) {
      input.addEventListener(ev, function (e) {
        if (kind === 'derived' && !isManual(f.key)) e.preventDefault();
      });
    });
    input.addEventListener('focus', function () {
      if (kind === 'derived' && !isManual(f.key)) { input.blur(); return; }
      input.select();
    });
    input.addEventListener('blur', function () {
      if (kind === 'ratio') input.value = fmtPct(state.draft[f.key]);
      else if (kind !== 'derived' || isManual(f.key)) input.value = fmt(parseNumInput(input.value));
      else input.value = fmt(T.num(state.draft[f.key]));
    });
    unit.appendChild(input);
    wrap.appendChild(unit);
    return wrap;
  }

  function buildCarryList() {
    var box = $('#carry-list'); box.innerHTML = '';
    T.FIELDS.filter(function (f) { return ['income', 'preTax', 'postTax'].indexOf(f.group) >= 0; }).forEach(function (f) {
      var lab = el('label', 'carry-item');
      var cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = !!state.settings.carry[f.key];
      cb.setAttribute('data-key', f.key);
      var s = el('span'); s.textContent = f.label;
      lab.appendChild(cb); lab.appendChild(s);
      box.appendChild(lab);
    });
  }

  function buildYearBase() {
    var box = $('#yearbase-list'); box.innerHTML = '';
    var yb = state.settings.yearBase || {};
    var years = Object.keys(yb).sort();
    if (!years.length) {
      var p = el('p', 'formula'); p.textContent = '暂无。从 1 月开始记录的话不需要填写。';
      box.appendChild(p);
    }
    years.forEach(function (y) {
      var row = el('div', 'yb-row');
      var yl = el('div', 'yb-y'); yl.textContent = y + '年';
      row.appendChild(yl);
      [['income', '累计收入'], ['social', '累计社保'], ['special', '累计专项'], ['taxPaid', '累计已缴税']].forEach(function (f) {
        var inp = document.createElement('input');
        inp.type = 'text'; inp.inputMode = 'decimal';
        inp.value = fmt(yb[y][f[0]]);
        inp.setAttribute('data-y', y); inp.setAttribute('data-f', f[0]);
        inp.title = f[1];
        row.appendChild(inp);
      });
      var del = el('button', 'btn danger-ghost sm');
      del.type = 'button'; del.textContent = '删除';
      del.addEventListener('click', function () {
        delete state.settings.yearBase[y];
        buildYearBase(); recompute();
      });
      row.appendChild(del);
      box.appendChild(row);
    });
    if (years.length) {
      var head = el('div', 'yb-row');
      var h0 = el('div', 'yb-h'); h0.textContent = '';
      head.appendChild(h0);
      ['累计收入', '累计社保', '累计专项', '累计已缴税'].forEach(function (t) {
        var h = el('div', 'yb-h'); h.textContent = t; head.appendChild(h);
      });
      box.insertBefore(head, box.firstChild);
    }
  }

  function buildRateTables() {
    [['#tbl-new', T.NEW_BRACKETS, '累计预扣预缴应纳税所得额'], ['#tbl-old', T.OLD_BRACKETS, '月应纳税所得额']].forEach(function (cfg) {
      var box = $(cfg[0]); box.innerHTML = '';
      var tb = el('table', 'tbl');
      var thead = el('thead');
      var trh = el('tr');
      ['级数', cfg[2], '税率', '速算扣除数'].forEach(function (h) {
        var th = el('th'); th.textContent = h; trh.appendChild(th);
      });
      thead.appendChild(trh); tb.appendChild(thead);
      var tbody = el('tbody');
      cfg[1].forEach(function (b) {
        var tr = el('tr');
        tr.setAttribute('data-level', b.level);
        [b.level, b.desc, (b.rate * 100).toFixed(0) + '%', fmt(b.deduct)].forEach(function (v) {
          var td = el('td'); td.textContent = v; tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      tb.appendChild(tbody); box.appendChild(tb);
    });
  }
  function highlightRateTable() {
    var lv = state.result ? state.result.rateLevel : 0;
    $$('#tbl-new tr[data-level]').forEach(function (tr) {
      tr.classList.toggle('is-hit', String(lv) === tr.getAttribute('data-level'));
    });
    var ov = state.result ? T.lookup(T.OLD_BRACKETS, state.result.oldTaxable).level : 0;
    $$('#tbl-old tr[data-level]').forEach(function (tr) {
      tr.classList.toggle('is-hit', String(ov) === tr.getAttribute('data-level'));
    });
  }

  /* ------------------------------------------------------------- 保存 / 删除 */
  async function saveDraft() {
    if (!state.draft) return;
    var d = state.draft;
    if (!T.parseMonth(d.payMonth)) { toast('工资发放月无效', 'err'); return; }
    var r = Store.upsert(d);
    if (r.error) { toast(r.error, 'err'); return; }
    if (r.saved === false) { toast('保存失败：浏览器存储空间不足', 'err'); return; }
    state.records = r.data.records;
    updateChip();
    recompute(); syncFields();
    if (r.dropped && r.dropped.length) {
      toast('已保存。达到 ' + MAX + ' 个月上限，自动覆盖了最早的月份：' + r.dropped.join('、'), 'err');
    } else {
      toast('已保存 ' + d.payMonth);
    }
  }

  async function deleteMonths(ids, label) {
    var r = Store.remove(ids);
    state.records = r.data.records;
    ids.forEach(function (i) { state.selected.delete(i); });
    updateChip();
    if (state.draft && ids.indexOf(state.draft.payMonth) >= 0) {
      var nx = nextRecord(state.draft.payMonth) || prevRecord(state.draft.payMonth);
      editMonth(nx ? nx.payMonth : nowYM());
    } else { recompute(); syncFields(); }
    toast('已删除 ' + label);
  }

  /* ------------------------------------------------------------- 导出 / 导入 */
  function doExport() {
    var list = state.records.slice().sort(function (a, b) { return T.monthKey(a.payMonth) - T.monthKey(b.payMonth); });
    var tag;
    if (state.exportScope === 'month') {
      var mv = $('#export-value').value || '';
      list = list.filter(function (r) { return T.fmtMonth(r.payMonth) === T.fmtMonth(mv); });
      tag = T.fmtMonth(mv);
    } else if (state.exportScope === 'year') {
      var yv = $('#export-value').value || '';
      list = list.filter(function (r) { return String(r.payMonth).slice(0, 4) === String(yv); });
      tag = yv + '年';
    } else {
      tag = '全部';
    }
    if (!list.length) { toast('所选范围内没有数据', 'err'); return; }

    Excel.exportArrayBuffer(list, state.settings).then(function (buf) {
      var name = '工资个税核算_' + tag + '_' + nowYM() + '.xlsx';
      download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name);
      toast('已导出 ' + name);
    }).catch(function (e) {
      toast('导出失败：' + (e && e.message ? e.message : e), 'err');
    });
  }
  function download(blob, name) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 3000);
  }
  function fileToArrayBuffer(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(new Error('文件读取失败')); };
      fr.readAsArrayBuffer(file);
    });
  }
  function doImport(file) {
    if (!file) return;
    if (!/\.xlsx$/i.test(file.name) && !/\.xlsm$/i.test(file.name)) {
      toast('请上传 .xlsx 格式的模板文件', 'err'); return;
    }
    fileToArrayBuffer(file).then(function (buf) {
      var parsed;
      try { parsed = Excel.parseWorkbook(buf); }
      catch (e) { toast(e.message || '无法识别该文件，请使用本程序导出的模板', 'err'); return; }
      if (!parsed.records.length) { toast('文件中没有读取到有效数据行', 'err'); return; }

      var mode = state.importMode === 'replace' ? 'replace' : 'merge';
      var records = state.records.slice();
      var added = 0, updated = 0;

      parsed.records.forEach(function (nr) {
        var pm = T.fmtMonth(nr.payMonth);
        var i = records.findIndex(function (r) { return T.fmtMonth(r.payMonth) === pm; });
        var rec = Object.assign({}, nr, { payMonth: pm, id: pm });
        if (i >= 0) { records[i] = Object.assign(records[i], rec); updated++; }
        else { records.push(rec); added++; }
      });

      var dropped = [];
      records.sort(function (a, b) { return T.monthKey(b.payMonth) - T.monthKey(a.payMonth); });
      if (records.length > MAX) {
        dropped = records.splice(MAX).map(function (r) { return r.payMonth; });
      }

      var settings = state.settings;
      if (parsed.carry) settings.carry = Object.assign(T.defaultCarry(), parsed.carry);
      if (Object.keys(parsed.yearBase || {}).length) {
        settings.yearBase = mode === 'replace' ? parsed.yearBase : Object.assign({}, settings.yearBase, parsed.yearBase);
      }
      state.records = records;
      state.settings = settings;
      Store.replaceAll(records, settings);
      updateChip();
      buildCarryList(); buildYearBase();

      var box = $('#import-result');
      box.hidden = false;
      box.className = 'imp-res';
      box.textContent = '导入完成：新增 ' + added + ' 个月，覆盖 ' + updated + ' 个月' +
        (dropped.length ? '；超出 ' + MAX + ' 个月上限，已自动移除最早的 ' + dropped.join('、') : '') + '。';
      var last = latestPayMonth();
      editMonth(last || nowYM());
      recompute();
      toast('导入成功');
    }).catch(function (e) {
      toast(e.message || '导入失败', 'err');
    });
  }

  /* ------------------------------------------------------------- 事件绑定 */
  function bind() {
    $$('.tab').forEach(function (t) {
      t.addEventListener('click', function () { switchView(t.getAttribute('data-view')); });
    });

    ['#period-year', '#period-month'].forEach(function (s) {
      $(s).addEventListener('change', function () {
        if (!state.draft) return;
        state.draft.periodMonth = $('#period-year').value + '-' + String($('#period-month').value).padStart(2, '0');
        state.draft.periodMonth = T.fmtMonth(state.draft.periodMonth);
        recompute();
      });
    });
    ['#pay-year', '#pay-month'].forEach(function (s) {
      $(s).addEventListener('change', function () {
        var pm = $('#pay-year').value + '-' + String($('#pay-month').value).padStart(2, '0');
        editMonth(T.fmtMonth(pm));
      });
    });

    $('#btn-prev-month').addEventListener('click', function () {
      var p = prevRecord(state.draft.payMonth);
      editMonth(p ? p.payMonth : T.nextMonth(state.draft.payMonth, -1));
    });
    $('#btn-next-month').addEventListener('click', function () {
      var n = nextRecord(state.draft.payMonth) ||
        state.records.filter(function (r) { return T.monthKey(r.payMonth) > T.monthKey(state.draft.payMonth); })[0];
      editMonth(n ? n.payMonth : T.nextMonth(state.draft.payMonth, 1));
    });

    $('#toggle-compare').addEventListener('click', function () {
      state.compareOpen = !state.compareOpen;
      this.setAttribute('aria-expanded', String(state.compareOpen));
      $('#body-compare').hidden = !state.compareOpen;
    });

    $('#btn-save').addEventListener('click', saveDraft);
    $('#btn-delete-one').addEventListener('click', function () {
      if (!state.draft) return;
      var pm = state.draft.payMonth;
      confirmDialog('删除该月存档',
        '确定要删除 <strong>' + esc(pm) + '</strong> 发放的这条记录吗？<br>删除后该月不再参与累计计算，且无法撤销。',
        '确认删除',
        function () { closeModal(); deleteMonths([pm], pm); });
    });

    $('#chk-all').addEventListener('change', function () {
      if (this.checked) state.records.forEach(function (r) { state.selected.add(T.fmtMonth(r.payMonth)); });
      else state.selected.clear();
      renderHistory();
    });
    $('#btn-bulk-delete').addEventListener('click', function () {
      var ids = Array.from(state.selected).sort(function (a, b) { return T.monthKey(b) - T.monthKey(a); });
      if (!ids.length) return;
      confirmDialog('批量删除 ' + ids.length + ' 条存档',
        '即将删除以下月份，删除后无法撤销：<ul><li>' + ids.map(esc).join('</li><li>') + '</li></ul>',
        '确认删除 ' + ids.length + ' 条',
        function () { closeModal(); deleteMonths(ids, ids.length + ' 条'); });
    });
    $('#btn-goto-calc').addEventListener('click', function () { switchView('calc'); });

    $$('#export-scope .seg-btn').forEach(function (b) {
      b.addEventListener('click', function () {
        state.exportScope = b.getAttribute('data-scope');
        $$('#export-scope .seg-btn').forEach(function (x) { x.classList.toggle('is-active', x === b); });
        renderExportOptions();
      });
    });
    $('#btn-export').addEventListener('click', doExport);
    $('#btn-template').addEventListener('click', function () {
      Excel.exportArrayBuffer([], state.settings).then(function (buf) {
        download(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), '工资个税核算_导入模板.xlsx');
        toast('模板已下载');
      }).catch(function () { toast('模板下载失败', 'err'); });
    });

    var dz = $('#dropzone'), fi = $('#file-input');
    dz.addEventListener('click', function () { fi.click(); });
    fi.addEventListener('change', function () { if (fi.files[0]) doImport(fi.files[0]); fi.value = ''; });
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('is-over'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('is-over'); });
    });
    dz.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files[0]) doImport(e.dataTransfer.files[0]);
    });
    $$('#import-mode .seg-btn').forEach(function (b) {
      b.addEventListener('click', function () {
        state.importMode = b.getAttribute('data-mode');
        $$('#import-mode .seg-btn').forEach(function (x) { x.classList.toggle('is-active', x === b); });
      });
    });

    $('#btn-save-carry').addEventListener('click', function () {
      var carry = T.defaultCarry();
      $$('#carry-list input[data-key]').forEach(function (cb) { carry[cb.getAttribute('data-key')] = cb.checked; });
      var r = Store.saveSettings({ carry: carry });
      if (r.saved === false) { toast('保存失败：浏览器存储空间不足', 'err'); return; }
      state.settings.carry = r.settings.carry;
      toast('沿用设置已保存');
    });
    $('#btn-yb-add').addEventListener('click', function () {
      var y = ($('#yb-year').value || '').trim();
      if (!/^\d{4}$/.test(y)) { toast('请输入四位年份，如 2026', 'err'); return; }
      state.settings.yearBase[y] = { income: 0, social: 0, special: 0, taxPaid: 0 };
      $('#yb-year').value = '';
      buildYearBase();
    });
    $('#btn-save-yb').addEventListener('click', function () {
      var yb = {};
      $$('#yearbase-list input[data-y]').forEach(function (inp) {
        var y = inp.getAttribute('data-y'), f = inp.getAttribute('data-f');
        yb[y] = yb[y] || { income: 0, social: 0, special: 0, taxPaid: 0 };
        yb[y][f] = parseNumInput(inp.value);
      });
      var r = Store.saveSettings({ yearBase: yb });
      if (r.saved === false) { toast('保存失败：浏览器存储空间不足', 'err'); return; }
      state.settings.yearBase = r.settings.yearBase;
      buildYearBase(); recompute();
      toast('起算基数已保存');
    });

    $('#btn-reset').addEventListener('click', function () {
      confirmDialog('清空全部数据',
        '将删除全部 <strong>' + state.records.length + '</strong> 条存档，此操作<strong>不可撤销</strong>。<br>建议先导出 Excel 备份。',
        '确认清空',
        function () {
          closeModal();
          Store.reset();
          state.records = []; state.selected.clear();
          updateChip(); editMonth(nowYM()); recompute();
          toast('已清空全部数据');
        });
    });

    $('#modal-cancel').addEventListener('click', closeModal);
    $('#modal-ok').addEventListener('click', function () { if (modalCb) modalCb(); });
    $('#modal').addEventListener('click', function (e) { if (e.target === this) closeModal(); });

    /* 点空白处收起字段说明与自定义下拉 */
    document.addEventListener('click', function (e) {
      if (!e.target.closest || !e.target.closest('.hint-btn')) closeAllHints();
      if (!e.target.closest || !e.target.closest('.dd')) closeAllDd();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') { closeAllDd(); closeAllHints(); }
    });

    // 税率表高亮随结果更新
    var mo = new MutationObserver(function () { highlightRateTable(); });
    mo.observe($('#anomaly-box'), { childList: true });
  }

  /* ------------------------------------------------------------- 启动 */
  function init() {
    buildYearOptions();
    $$('select').forEach(enhanceSelect);
    buildFieldGrids();
    buildRateTables();
    bind();
    loadData();
    buildCarryList();
    buildYearBase();
    var last = latestPayMonth();
    editMonth(last || nowYM());
    highlightRateTable();
    switchView('calc');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
