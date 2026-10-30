/* =============================================================================
 * 工资个税核算 · 计算内核（前后端共用）
 *
 * 计算口径严格依据《个税计算逻辑.xls》与《工资条.xls》，并按用户 2026-10-16 的修正调整：
 *
 *   请假扣发     = 请假时数 ÷ 标准计薪时数 × (基本工资 + 绩效奖 + 电话津贴)
 *   个人社保     = (基本工资 + 绩效奖) × 社保折算比例 × 社保综合缴纳比例
 *   个人公积金   = (基本工资 + 绩效奖) × 公积金个人缴纳比例
 *
 *   应发合计     = 基本工资 + 住宿补贴 + 绩效奖 + 电话津贴 + 其他(税前) - 漏打卡 - 请假扣发
 *   当月收入     = 应发合计                    ← 请假扣发已含在应发合计内，不可再扣一次
 *   累计应纳税所得额 = 累计收入 - 5000×月序 - 累计社保公积金 - 累计专项附加扣除
 *   本月个税     = (累计应纳税所得额 × 预扣率 - 速算扣除数) - 累计已预缴税额
 *                  （为负值时暂不退税，按 0 计）
 *   扣发合计     = 个人社保 + 个人公积金 + 个税 + 水电费 + 房租 + 其他扣发
 *   实发合计     = 应发合计 - 扣发合计
 *   本次发放     = 实发合计 - 前期已发
 *
 *   个人社保与个人公积金之和进入【累计社保公积金】，参与个税计算。
 *   请假扣发 / 个人社保 / 个人公积金 默认由公式派生，允许手动覆盖（覆盖后标黄提示）。
 *
 *   月序：按【工资发放月】在该自然年内的序号计（1~12），跨年重新累计。
 *   旧算法（月度税率表）仅作参考对比，不参与实发计算。
 * ============================================================================= */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TaxCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------------------------------------------------------------- 基础工具 */

  /** 四舍五入到 2 位（先消除浮点噪声，避免 53690.315 → 53690.31） */
  function round2(v) {
    if (typeof v !== 'number' || !isFinite(v)) return 0;
    var t = Math.round(v * 1e6) / 1e6;
    return Math.round(t * 100 + (t >= 0 ? 1e-6 : -1e-6)) / 100;
  }

  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    var n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[,，\s¥￥]/g, ''));
    return isFinite(n) ? n : 0;
  }

  /** "2026-01" → { year:2026, month:1 } */
  function parseMonth(s) {
    if (!s) return null;
    var m = String(s).trim().replace(/[/.]/g, '-').match(/^(\d{4})-(\d{1,2})$/);
    if (!m) return null;
    return { year: parseInt(m[1], 10), month: parseInt(m[2], 10) };
  }

  function fmtMonth(s) {
    var p = parseMonth(s);
    return p ? p.year + '-' + String(p.month).padStart(2, '0') : String(s || '');
  }

  /** 所属月 + 1 个月 = 发放月 */
  function nextMonth(s, delta) {
    var p = parseMonth(s);
    if (!p) return null;
    var t = p.year * 12 + (p.month - 1) + (delta === undefined ? 1 : delta);
    return Math.floor(t / 12) + '-' + String((t % 12) + 1).padStart(2, '0');
  }

  function monthKey(s) {
    var p = parseMonth(s);
    return p ? p.year * 12 + (p.month - 1) : -Infinity;
  }

  /* ------------------------------------------------------------ 税率表（新） */
  /* 累计预扣预缴应纳税所得额 —— 年度七级 */
  var NEW_BRACKETS = [
    { level: 1, max: 36000, rate: 0.03, deduct: 0, desc: 'A ≤ 36000' },
    { level: 2, max: 144000, rate: 0.10, deduct: 2520, desc: '36000 ＜ A ≤ 144000' },
    { level: 3, max: 300000, rate: 0.20, deduct: 16920, desc: '144000 ＜ A ≤ 300000' },
    { level: 4, max: 420000, rate: 0.25, deduct: 31920, desc: '300000 ＜ A ≤ 420000' },
    { level: 5, max: 660000, rate: 0.30, deduct: 52920, desc: '420000 ＜ A ≤ 660000' },
    { level: 6, max: 960000, rate: 0.35, deduct: 85920, desc: '660000 ＜ A ≤ 960000' },
    { level: 7, max: Infinity, rate: 0.45, deduct: 181920, desc: 'A ＞ 960000' }
  ];

  /* 税率表（旧）—— 月度七级，仅作对比参考 */
  var OLD_BRACKETS = [
    { level: 1, max: 3000, rate: 0.03, deduct: 0, desc: 'A ≤ 3000' },
    { level: 2, max: 12000, rate: 0.10, deduct: 210, desc: '3000 ＜ A ≤ 12000' },
    { level: 3, max: 25000, rate: 0.20, deduct: 1410, desc: '12000 ＜ A ≤ 25000' },
    { level: 4, max: 35000, rate: 0.25, deduct: 2660, desc: '25000 ＜ A ≤ 35000' },
    { level: 5, max: 55000, rate: 0.30, deduct: 4410, desc: '35000 ＜ A ≤ 55000' },
    { level: 6, max: 80000, rate: 0.35, deduct: 7160, desc: '55000 ＜ A ≤ 80000' },
    { level: 7, max: Infinity, rate: 0.45, deduct: 15160, desc: 'A ＞ 80000' }
  ];

  function lookup(brackets, amount) {
    if (amount <= 0) return { level: 0, rate: 0, deduct: 0, desc: '—' };
    for (var i = 0; i < brackets.length; i++) {
      if (amount <= brackets[i].max) return brackets[i];
    }
    return brackets[brackets.length - 1];
  }

  function taxByBracket(brackets, amount) {
    if (amount <= 0) return 0;
    var b = lookup(brackets, amount);
    return round2(amount * b.rate - b.deduct);
  }

  /* ------------------------------------------------------------------ 字段表
   *
   * kind 取值：
   *   'input'   纯手动录入
   *   'ratio'   百分比录入（界面显示 %，内部存小数），可参与沿用
   *   'hours'   小时数录入
   *   'derived' 由公式派生，公式成立时界面灰显；用户手改则标记 covered
   *
   * group：income / preTax / postTax / compare —— 对应四个录入分区
   */

  var FIELDS = [
    /* ---- 收入项 ---- */
    { key: 'base', label: '基本工资', group: 'income', carry: true },
    { key: 'housingSubsidy', label: '住宿补贴', group: 'income', carry: true },
    { key: 'performance', label: '绩效奖', group: 'income', carry: false },
    { key: 'phoneAllowance', label: '电话津贴', group: 'income', carry: true },
    { key: 'otherPreTax', label: '其他（税前）', group: 'income', carry: false, hint: '如奖金 / 春节开门红等' },
    { key: 'missedPunch', label: '漏打卡', group: 'income', carry: false, hint: '税前扣除。第二次 10 元，第三次 20 元，以此类推' },
    { key: 'leaveHours', label: '请假时数', group: 'income', kind: 'hours', carry: false, hint: '本月请假小时数，用于计算请假扣发' },
    { key: 'standardHours', label: '标准计薪时数', group: 'income', kind: 'hours', carry: true, hint: '月度标准计薪小时数，每月一般固定' },
    { key: 'leaveDeduction', label: '请假扣发', group: 'income', kind: 'derived', carry: false, hint: '由请假时数自动算出，也可手动修改' },

    /* ---- 税前扣除 ---- */
    { key: 'socialRatio', label: '社保折算比例', group: 'preTax', kind: 'ratio', carry: true, hint: '参与个人社保计算，按 % 填写' },
    { key: 'socialRate', label: '社保综合缴纳比例', group: 'preTax', kind: 'ratio', carry: true, hint: '参与个人社保计算，按 % 填写' },
    { key: 'fundRate', label: '公积金个人缴纳比例', group: 'preTax', kind: 'ratio', carry: true, hint: '参与个人公积金计算，按 % 填写' },
    { key: 'personalSocial', label: '个人社保', group: 'preTax', kind: 'derived', carry: true, hint: '由基本工资、绩效奖与社保比例算出，也可手动修改' },
    { key: 'personalFund', label: '个人公积金', group: 'preTax', kind: 'derived', carry: true, hint: '由基本工资、绩效奖与公积金比例算出，也可手动修改' },
    { key: 'specialDeduction', label: '专项附加扣除', group: 'preTax', carry: true, hint: '当月总额' },

    /* ---- 税后扣除 ---- */
    { key: 'utility', label: '水电费', group: 'postTax', carry: true, hint: '税后扣除' },
    { key: 'rent', label: '房租', group: 'postTax', carry: true, hint: '税后扣除' },
    { key: 'otherDeduction', label: '其他扣发', group: 'postTax', carry: false, hint: '如借车费、饭卡充值费' },
    { key: 'advancePaid', label: '前期已发', group: 'postTax', carry: false },

    /* ---- 公司工资条比对 ---- */
    { key: 'psGross', label: '工资条应发合计', group: 'compare', carry: false },
    { key: 'psTax', label: '工资条扣税', group: 'compare', carry: false },
    { key: 'psNet', label: '工资条实发合计', group: 'compare', carry: false }
  ];

  /** 公式派生字段：值等于公式结果时视为「未覆盖」 */
  var DERIVED_KEYS = FIELDS.filter(function (f) { return f.kind === 'derived'; })
    .map(function (f) { return f.key; });

  var RATIO_KEYS = FIELDS.filter(function (f) { return f.kind === 'ratio'; })
    .map(function (f) { return f.key; });

  /** 参与「累计社保公积金」的字段 */
  var SOCIAL_KEYS = ['personalSocial', 'personalFund'];

  /** 旧模板兼容：曾用「五险一金」一个字段承载社保 + 公积金 */
  var LEGACY_MAP = { socialInsurance: 'personalSocial' };

  var FIELD_MAP = {};
  FIELDS.forEach(function (f) { FIELD_MAP[f.key] = f; });

  var CARRY_FIELDS = FIELDS.filter(function (f) { return f.carry; }).map(function (f) { return f.key; });

  function emptyRecord() {
    var r = { id: null, periodMonth: '', payMonth: '', note: '' };
    FIELDS.forEach(function (f) { r[f.key] = 0; });
    return r;
  }

  /* ------------------------------------------------------ 派生字段公式（唯一实现处）

     界面上不展示这些公式，只在文档与本注释中保留，便于日后核对。
     三个公式都对外暴露，前端预览与后端存档共用同一份，避免两处口径漂移。 */

  /** 请假扣发 = 请假时数 ÷ 标准计薪时数 × (基本工资 + 绩效奖 + 电话津贴) */
  function calcLeaveDeduction(r) {
    var h = num(r.leaveHours), std = num(r.standardHours);
    if (h <= 0 || std <= 0) return 0;
    var wage = num(r.base) + num(r.performance) + num(r.phoneAllowance);
    return round2(h / std * wage);
  }

  /** 个人社保 = (基本工资 + 绩效奖) × 社保折算比例 × 社保综合缴纳比例 */
  function calcPersonalSocial(r) {
    var wage = num(r.base) + num(r.performance);
    return round2(wage * num(r.socialRatio) * num(r.socialRate));
  }

  /** 个人公积金 = (基本工资 + 绩效奖) × 公积金个人缴纳比例 */
  function calcPersonalFund(r) {
    var wage = num(r.base) + num(r.performance);
    return round2(wage * num(r.fundRate));
  }

  /** 三个派生字段的公式值 */
  function derivedValues(r) {
    return {
      leaveDeduction: calcLeaveDeduction(r),
      personalSocial: calcPersonalSocial(r),
      personalFund: calcPersonalFund(r)
    };
  }

  /**
   * 解析派生字段。
   *
   * 判定规则：只有「公式算了、且用户填的数与公式不同」才算手动覆盖。
   * 反过来，值为 0 时无法区分「用户就是想填 0」还是「字段为空」，
   * 因此把「公式为 0」和「用户值为 0」都视为未覆盖，取公式值。
   * 用户若确实需要 0，可在界面把该字段切到「手动」模式——前端会通过 manualMap 显式告知，
   * 此时 resolveDerived 收到的值虽为 0，但 derivedValues 因缺少时数/比例也为 0，
   * 公式与用户值相等，仍不会被改写。
   *
   * @param {Object} [manualMap] 界面显式切到「手动」的字段，{ key: true }。
   *                 命中时无条件保留录入值，不参与数值比对 —— 模式由用户点击决定，不由程序推断。
   * @returns {{ resolved: Object, covered: Object }}
   *          resolved —— 三个字段的公式值
   *          covered  —— { key: true } 表示该字段被手动覆盖
   */
  function resolveDerived(r, manualMap) {
    var f = derivedValues(r);
    var covered = {};
    manualMap = manualMap || {};
    DERIVED_KEYS.forEach(function (k) {
      var formula = f[k];
      var manual = num(r[k]);
      if (manualMap[k]) { covered[k] = true; return; }
      if (formula <= 0) {
        /* 公式算不出结果：保留录入值，不视为覆盖 */
        covered[k] = false;
        return;
      }
      /* 公式有结果：用户填了不同的非零值才算覆盖，否则一律用公式值 */
      if (manual !== 0 && Math.abs(round2(manual - formula)) > 0.005) {
        covered[k] = true;
      } else {
        r[k] = formula;
        covered[k] = false;
      }
    });
    return { resolved: f, covered: covered };
  }

  /* -------------------------------------------------------------- 核心计算 */

  /**
   * @param {Array} records 记录数组（顺序任意，内部按发放月排序）
   * @param {Object} opts   { yearBase: { "2026": {income, social, special, taxPaid} } }
   * @returns {Array} 每条记录附加计算结果的数组（保持与输入同序，通过 id 关联）
   */
  function compute(records, opts) {
    opts = opts || {};
    var yearBase = opts.yearBase || {};
    var list = records.map(function (r) { return Object.assign({}, r); });
    list.sort(function (a, b) { return monthKey(a.payMonth) - monthKey(b.payMonth); });
    // 派生字段（请假扣发 / 个人社保 / 个人公积金）先统一解析，
    // 后续所有累计与个税计算都基于解析后的值，避免中途改口径
    var coveredMap = {};
    list.forEach(function (r) {
      var rr = resolveDerived(r, r.__manual);
      coveredMap[r.id] = rr.covered;
    });

    // 按发放月所在自然年分组累计
    var state = {};   // year -> { income, social, special, taxPaid, months:Set }
    var prevRate = {}; // year -> 上月适用预扣率
    var out = [];

    for (var i = 0; i < list.length; i++) {
      var r = list[i];
      var pm = parseMonth(r.payMonth) || parseMonth(r.periodMonth) || { year: 0, month: 0 };
      var year = pm.year || 0;
      var seq = pm.month || 0; // 月序 1~12

      if (!state[year]) {
        var b = yearBase[String(year)] || {};
        state[year] = {
          income: num(b.income),
          social: num(b.social),
          special: num(b.special),
          taxPaid: num(b.taxPaid),
          months: []
        };
        prevRate[year] = 0;
      }
      var st = state[year];

      /* ---- 派生字段已在循环外解析，这里直接从记录取值 ---- */
      var dv = { covered: coveredMap[r.id] || { leaveDeduction: false, personalSocial: false, personalFund: false } };
      var personalSocial = num(r.personalSocial);
      var personalFund = num(r.personalFund);
      var socialTotal = round2(personalSocial + personalFund);

      /* ---- 收入 ---- */
      /* 请假扣发已包含在应发合计内，当月收入即应发合计，不可再扣一次 */
      var gross = round2(num(r.base) + num(r.housingSubsidy) + num(r.performance) +
        num(r.phoneAllowance) + num(r.otherPreTax) - num(r.missedPunch) - num(r.leaveDeduction));
      var monthlyIncome = gross;

      /* ---- 累计 ---- */
      st.income = round2(st.income + monthlyIncome);
      st.social = round2(st.social + socialTotal);
      st.special = round2(st.special + num(r.specialDeduction));
      st.months.push(seq);

      var cumBasic = 5000 * seq;
      var cumTaxable = round2(st.income - cumBasic - st.social - st.special);
      var br = lookup(NEW_BRACKETS, cumTaxable);
      var cumTaxDue = taxByBracket(NEW_BRACKETS, cumTaxable);

      var rawTax = round2(cumTaxDue - st.taxPaid);
      var tax = rawTax < 0 ? 0 : rawTax;         // 负值时暂不退税
      st.taxPaid = round2(st.taxPaid + tax);

      /* ---- 旧算法（月度税率表，仅参考） ---- */
      var oldTaxable = round2(monthlyIncome - 5000 - socialTotal - num(r.specialDeduction));
      var oldBr = lookup(OLD_BRACKETS, oldTaxable);
      var oldTax = taxByBracket(OLD_BRACKETS, oldTaxable);

      /* ---- 税后 ---- */
      var totalDeduction = round2(socialTotal + tax +
        num(r.utility) + num(r.rent) + num(r.otherDeduction));
      var netPay = round2(gross - totalDeduction);
      var currentPay = round2(netPay - num(r.advancePaid));

      var res = {
        id: r.id,
        periodMonth: r.periodMonth,
        payMonth: r.payMonth,
        seq: seq,
        year: year,
        gross: gross,
        monthlyIncome: monthlyIncome,
        leaveDeduction: num(r.leaveDeduction),
        personalSocial: personalSocial,
        personalFund: personalFund,
        socialTotal: socialTotal,
        covered: coveredMap[r.id] || dv.covered,
        cumIncome: st.income,
        cumBasic: cumBasic,
        cumSocial: st.social,
        cumSpecial: st.special,
        cumTaxable: cumTaxable,
        rate: br.rate,
        rateLevel: br.level,
        cumTaxDue: cumTaxDue,
        cumTaxPaid: st.taxPaid,
        tax: tax,
        rawTax: rawTax,
        oldTaxable: oldTaxable,
        oldTax: oldTax,
        oldRate: oldBr.rate,
        totalDeduction: totalDeduction,
        netPay: netPay,
        currentPay: currentPay,
        anomalies: []
      };

      res.anomalies = detect(r, res, {
        prevRate: prevRate[year],
        months: st.months,
        hasYearBase: !!yearBase[String(year)]
      });
      prevRate[year] = br.rate;

      out.push(res);
    }

    /* ---- 跨记录校验：年度缺口、月份重复 ---- */
    var byYear = {}, dupCount = {};
    list.forEach(function (r) {
      var p = parseMonth(r.payMonth) || parseMonth(r.periodMonth);
      if (!p) return;
      (byYear[p.year] = byYear[p.year] || []).push(p.month);
      var k = p.year + '-' + p.month;
      dupCount[k] = (dupCount[k] || 0) + 1;
    });

    var index = {};
    out.forEach(function (o) { index[o.id] = o; });

    list.forEach(function (r) {
      var p = parseMonth(r.payMonth) || parseMonth(r.periodMonth);
      if (!p) return;
      var o = index[r.id];
      if (!o) return;

      if (dupCount[p.year + '-' + p.month] > 1) {
        o.anomalies.push(anomaly('error', 'duplicate', '发放月 ' + fmtMonth(r.payMonth) + ' 存在 ' + dupCount[p.year + '-' + p.month] + ' 条重复记录，累计税额会重复累加'));
      }
      var months = byYear[p.year];
      var min = Math.min.apply(null, months);
      var max = Math.max.apply(null, months);
      var gaps = [];
      for (var m = min; m <= max; m++) if (months.indexOf(m) === -1) gaps.push(m + '月');
      if (gaps.length) {
        o.anomalies.push(anomaly('warn', 'gap', '本年度缺少 ' + gaps.join('、') + '，累计数据不完整，个税计算会偏低'));
      }
      var yb = opts.yearBase || {};
      var hasBase = !!yb[String(p.year)];
      if (min !== 1 && !hasBase && p.month === min) {
        o.anomalies.push(anomaly('warn', 'noJan', '本年度最早记录为 ' + min + ' 月，未从 1 月起算。若此前月份已发过工资，请在「设置 → 年度累计起算基数」中补录，否则累计税额不准'));
      }
    });

    // 保持与输入相同的顺序返回
    var order = {};
    records.forEach(function (r, i) { order[r.id] = i; });
    out.sort(function (a, b) { return (order[a.id] === undefined ? 0 : order[a.id]) - (order[b.id] === undefined ? 0 : order[b.id]); });
    return out;
  }

  function anomaly(level, code, message) {
    return { level: level, code: code, message: message };
  }

  /** 单条记录的异常/提示判定 */
  function detect(r, res, ctx) {
    var a = [];
    var EPS = 0.01;

    /* —— 与公司工资条比对 —— */
    if (num(r.psGross) > 0 && Math.abs(round2(num(r.psGross) - res.gross)) > EPS) {
      a.push(anomaly('error', 'grossDiff', '工资条应发合计 ¥' + round2(num(r.psGross)) +
        ' 与计算值 ¥' + res.gross + ' 相差 ¥' + round2(num(r.psGross) - res.gross)));
    }
    if (r.psTax !== null && r.psTax !== undefined && String(r.psTax) !== '') {
      var d = round2(num(r.psTax) - res.tax);
      if (Math.abs(d) > EPS) {
        a.push(anomaly('error', 'taxDiff', '扣税与计算值相差 ¥' + d + '（工资条 ¥' +
          round2(num(r.psTax)) + ' / 计算 ¥' + res.tax + '）' + (d > 0 ? '，公司多扣' : '，公司少扣')));
      }
    }
    if (num(r.psNet) > 0) {
      var dn = round2(num(r.psNet) - res.netPay);
      if (Math.abs(dn) > EPS) {
        a.push(anomaly('error', 'netDiff', '实发合计与计算值相差 ¥' + dn + '（工资条 ¥' +
          round2(num(r.psNet)) + ' / 计算 ¥' + res.netPay + '）' + (dn < 0 ? '，到手少了' : '，到手多了')));
      }
    }

    /* —— 数据合理性 —— */
    FIELDS.forEach(function (f) {
      if (num(r[f.key]) < 0) a.push(anomaly('error', 'negative', f.label + ' 不能为负数'));
    });

    if (res.gross <= 0) a.push(anomaly('error', 'noIncome', '应发合计为 ¥' + res.gross + '，请检查收入项'));

    /* —— 派生字段：能算出公式值时做交叉校验 —— */
    var dv = derivedValues(r);
    if (dv.leaveDeduction > 0 && res.covered && res.covered.leaveDeduction) {
      a.push(anomaly('warn', 'leaveManual', '请假扣发为手动填写 ¥' + num(r.leaveDeduction) +
        '，公式结果为 ¥' + dv.leaveDeduction + '。若非病假折算、跨月补扣等特殊情况，请核对'));
    }
    if (dv.leaveDeduction > 0 && num(r.leaveDeduction) > num(r.base) + num(r.performance) + num(r.phoneAllowance)) {
      a.push(anomaly('error', 'leaveOver', '请假扣发 ¥' + num(r.leaveDeduction) +
        ' 超过了基本工资 + 绩效奖 + 电话津贴之和 ¥' +
        round2(num(r.base) + num(r.performance) + num(r.phoneAllowance)) + '，请核对请假时数'));
    }
    if (num(r.leaveHours) > 0 && num(r.standardHours) <= 0) {
      a.push(anomaly('error', 'stdHours', '已填请假时数 ' + num(r.leaveHours) +
        ' 小时，但标准计薪时数为 0，无法计算请假扣发。请在「收入项 → 标准计薪时数」中填写'));
    }
    if (num(r.leaveHours) > num(r.standardHours) && num(r.standardHours) > 0) {
      a.push(anomaly('warn', 'leaveHoursHigh', '请假时数 ' + num(r.leaveHours) +
        ' 小时已超过标准计薪时数 ' + num(r.standardHours) + ' 小时，请核对'));
    }
    if (dv.personalSocial > 0 && res.covered && res.covered.personalSocial) {
      a.push(anomaly('warn', 'socialManual', '个人社保为手动填写 ¥' + num(r.personalSocial) +
        '，公式结果为 ¥' + dv.personalSocial + '。若因社保基数上下限封顶产生差异属正常'));
    }
    if (dv.personalFund > 0 && res.covered && res.covered.personalFund) {
      a.push(anomaly('warn', 'fundManual', '个人公积金为手动填写 ¥' + num(r.personalFund) +
        '，公式结果为 ¥' + dv.personalFund + '。若因缴存基数调整产生差异属正常'));
    }

    /* —— 比例合理性 —— */
    var pctSum = num(r.socialRatio) * num(r.socialRate);
    if (num(r.socialRatio) > 0 && pctSum > 0.5) {
      a.push(anomaly('warn', 'socialRateHigh', '社保折算比例 × 社保综合缴纳比例 = ' +
        (pctSum * 100).toFixed(2) + '%，个人社保负担明显偏高，请核对比例设置'));
    }
    if (num(r.fundRate) > 0.3) {
      a.push(anomaly('warn', 'fundRateHigh', '公积金个人缴纳比例 ' + (num(r.fundRate) * 100).toFixed(2) +
        '% 偏高，通常为 5%~12%，请核对'));
    }
    if (num(r.socialRatio) > 0 && (num(r.socialRate) <= 0 || num(r.fundRate) <= 0)) {
      a.push(anomaly('info', 'rateMissing', '社保折算比例已填，但社保综合缴纳比例或公积金个人缴纳比例为 0，对应扣款会算成 0'));
    }

    if (num(r.advancePaid) > res.netPay) {
      a.push(anomaly('error', 'advanceOver', '前期已发 ¥' + num(r.advancePaid) + ' 超过实发合计 ¥' + res.netPay + '，本次发放为负'));
    }
    if (num(r.specialDeduction) > 8000) {
      a.push(anomaly('warn', 'specialHigh', '专项附加扣除 ¥' + num(r.specialDeduction) + ' 偏高，请核对是否符合扣除标准'));
    }
    if (res.socialTotal > res.gross && res.gross > 0) {
      a.push(anomaly('warn', 'socialOver', '个人社保与个人公积金合计 ¥' + res.socialTotal + ' 超过应发合计，请核对'));
    }

    var pp = parseMonth(r.periodMonth), pm2 = parseMonth(r.payMonth);
    if (pp && pm2 && monthKey(r.payMonth) < monthKey(r.periodMonth)) {
      a.push(anomaly('error', 'monthOrder', '工资发放月早于所属月，请核对（通常发放月 = 所属月 + 1 个月）'));
    }
    if (pp && pm2 && monthKey(r.payMonth) > monthKey(r.periodMonth) + 1) {
      a.push(anomaly('warn', 'monthSpan', '发放月比所属月晚 ' + (monthKey(r.payMonth) - monthKey(r.periodMonth)) + ' 个月，请确认税款所属期是否正确'));
    }

    /* —— 提示级 —— */
    if (res.rawTax < 0) {
      a.push(anomaly('info', 'taxFloor', '按累计法计算本月应退 ¥' + Math.abs(res.rawTax) + '，依规暂不退税，本月个税按 0 计（年度汇算时多退少补）'));
    }
    if (ctx && ctx.prevRate && res.rate > ctx.prevRate) {
      a.push(anomaly('info', 'bracketJump', '本月预扣率由 ' + pct(ctx.prevRate) + ' 跳档至 ' + pct(res.rate) +
        '（累计应纳税所得额 ¥' + res.cumTaxable + '），到手工资下降属正常现象'));
    }
    if (res.cumTaxable < 0 && res.seq > 0) {
      a.push(anomaly('info', 'negativeTaxable', '累计应纳税所得额为负，本月无需预扣个税'));
    }
    return a;
  }

  function pct(v) { return (v * 100).toFixed(0) + '%'; }

  /** 单条即时计算（不依赖累计，用于表单实时预览） */
  function previewOne(record, ctx, manualMap) {
    ctx = ctx || {};
    var r = Object.assign({}, record);
    var dv = resolveDerived(r, manualMap);
    var socialTotal = round2(num(r.personalSocial) + num(r.personalFund));
    var gross = round2(num(r.base) + num(r.housingSubsidy) + num(r.performance) +
      num(r.phoneAllowance) + num(r.otherPreTax) - num(r.missedPunch) - num(r.leaveDeduction));
    var monthlyIncome = gross;
    var oldTaxable = round2(monthlyIncome - 5000 - socialTotal - num(r.specialDeduction));
    var oldTax = taxByBracket(OLD_BRACKETS, oldTaxable);
    return {
      gross: gross,
      monthlyIncome: monthlyIncome,
      leaveDeduction: num(r.leaveDeduction),
      personalSocial: num(r.personalSocial),
      personalFund: num(r.personalFund),
      socialTotal: socialTotal,
      covered: dv.covered,
      oldTaxable: oldTaxable,
      oldTax: oldTax
    };
  }

  /** 默认沿用策略 */
  function defaultCarry() {
    var o = {};
    FIELDS.forEach(function (f) { if (f.key !== 'psGross' && f.key !== 'psTax' && f.key !== 'psNet') o[f.key] = !!f.carry; });
    return o;
  }

  return {
    round2: round2,
    num: num,
    parseMonth: parseMonth,
    fmtMonth: fmtMonth,
    nextMonth: nextMonth,
    monthKey: monthKey,
    NEW_BRACKETS: NEW_BRACKETS,
    OLD_BRACKETS: OLD_BRACKETS,
    lookup: lookup,
    taxByBracket: taxByBracket,
    FIELDS: FIELDS,
    FIELD_MAP: FIELD_MAP,
    CARRY_FIELDS: CARRY_FIELDS,
    DERIVED_KEYS: DERIVED_KEYS,
    RATIO_KEYS: RATIO_KEYS,
    SOCIAL_KEYS: SOCIAL_KEYS,
    LEGACY_MAP: LEGACY_MAP,
    emptyRecord: emptyRecord,
    compute: compute,
    detect: detect,
    previewOne: previewOne,
    defaultCarry: defaultCarry,
    /* 派生字段公式（界面不展示，供程序与测试调用） */
    calcLeaveDeduction: calcLeaveDeduction,
    calcPersonalSocial: calcPersonalSocial,
    calcPersonalFund: calcPersonalFund,
    derivedValues: derivedValues,
    resolveDerived: resolveDerived
  };
});
