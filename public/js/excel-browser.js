/* =============================================================================
 * Excel 模板（浏览器版）—— 导出与导入使用完全相同的列结构，用户无需二次处理
 * Sheet1 工资数据（主表） / Sheet2 沿用设置 / Sheet3 年度起算基数 / Sheet4 填写说明
 *
 * 与 lib/excel.js 逻辑一致，仅把 Node 依赖换成浏览器全局：
 *   - XLSX 来自 vendor/xlsx.full.min.js（SheetJS 0.20.3 官方社区版）
 *   - JSZip 来自 vendor/jszip.min.js（用于导出后为表头补深蓝底白字样式）
 *   - T    来自 shared/tax-core.js（UMD，window.TaxCore）
 * ============================================================================= */
(function () {
  'use strict';
  var XLSX = window.XLSX;
  var JSZip = window.JSZip;
  var T = window.TaxCore;

  /* 主表列定义：input=true 的列在导入时被读取，其余为计算结果列（导入时忽略） */
  var COLS = [
    { key: 'periodMonth', label: '工资所属月', input: true, w: 12, t: 's' },
    { key: 'payMonth', label: '工资发放月', input: true, w: 12, t: 's' },
    /* ---- 收入项 ---- */
    { key: 'base', label: '基本工资', input: true, w: 12, t: 'n' },
    { key: 'housingSubsidy', label: '住宿补贴', input: true, w: 11, t: 'n' },
    { key: 'performance', label: '绩效奖', input: true, w: 11, t: 'n' },
    { key: 'phoneAllowance', label: '电话津贴', input: true, w: 11, t: 'n' },
    { key: 'otherPreTax', label: '其他（税前）', input: true, w: 13, t: 'n' },
    { key: 'missedPunch', label: '漏打卡', input: true, w: 10, t: 'n' },
    { key: 'leaveHours', label: '请假时数', input: true, w: 10, t: 'n' },
    { key: 'standardHours', label: '标准计薪时数', input: true, w: 13, t: 'n' },
    { key: 'leaveDeduction', label: '请假扣发', input: true, w: 11, t: 'n' },
    /* ---- 税前扣除 ---- */
    { key: 'socialRatio', label: '社保折算比例(%)', input: true, w: 15, t: 'n' },
    { key: 'socialRate', label: '社保综合缴纳比例(%)', input: true, w: 18, t: 'n' },
    { key: 'fundRate', label: '公积金个人缴纳比例(%)', input: true, w: 19, t: 'n' },
    { key: 'personalSocial', label: '个人社保', input: true, w: 11, t: 'n' },
    { key: 'personalFund', label: '个人公积金', input: true, w: 11, t: 'n' },
    { key: 'specialDeduction', label: '专项附加扣除', input: true, w: 13, t: 'n' },
    /* ---- 税后扣除 ---- */
    { key: 'utility', label: '水电费', input: true, w: 10, t: 'n' },
    { key: 'rent', label: '房租', input: true, w: 10, t: 'n' },
    { key: 'otherDeduction', label: '其他扣发', input: true, w: 11, t: 'n' },
    { key: 'advancePaid', label: '前期已发', input: true, w: 11, t: 'n' },
    /* ---- 公司工资条比对 ---- */
    { key: 'psGross', label: '工资条应发合计', input: true, w: 15, t: 'n' },
    { key: 'psTax', label: '工资条扣税', input: true, w: 13, t: 'n' },
    { key: 'psNet', label: '工资条实发合计', input: true, w: 15, t: 'n' },
    { key: 'note', label: '备注', input: true, w: 20, t: 's' },
    /* ↓ 以下为计算列，导入时忽略 */
    { key: 'gross', label: '应发合计(计算)', input: false, w: 14, t: 'n' },
    { key: 'monthlyIncome', label: '当月收入(计算)', input: false, w: 14, t: 'n' },
    { key: 'socialTotal', label: '本月社保公积金(计算)', input: false, w: 18, t: 'n' },
    { key: 'cumIncome', label: '税前工资累计', input: false, w: 14, t: 'n' },
    { key: 'cumBasic', label: '累计减除费用', input: false, w: 14, t: 'n' },
    { key: 'cumSocial', label: '累计社保公积金', input: false, w: 15, t: 'n' },
    { key: 'cumSpecial', label: '累计专项附加扣除', input: false, w: 16, t: 'n' },
    { key: 'cumTaxable', label: '累计应纳税所得额', input: false, w: 16, t: 'n' },
    { key: 'rate', label: '适用预扣率', input: false, w: 11, t: 's' },
    { key: 'tax', label: '本月个税(计算)', input: false, w: 14, t: 'n' },
    { key: 'cumTaxPaid', label: '已缴税费累计', input: false, w: 14, t: 'n' },
    { key: 'oldTax', label: '旧算法参考个税', input: false, w: 15, t: 'n' },
    { key: 'totalDeduction', label: '扣发合计', input: false, w: 12, t: 'n' },
    { key: 'netPay', label: '实发合计(计算)', input: false, w: 14, t: 'n' },
    { key: 'currentPay', label: '本次发放', input: false, w: 12, t: 'n' },
    { key: 'anomalyText', label: '异常提示', input: false, w: 46, t: 's' }
  ];

  var INPUT_COUNT = COLS.filter(function (c) { return c.input; }).length;
  var PCT_COLS = { socialRatio: 1, socialRate: 1, fundRate: 1 };
  var NUMFMT = '#,##0.00';

  /* ------------------------------------------------------------------ 导出 */
  function buildWorkbook(records, settings) {
    settings = settings || {};
    var computed = T.compute(records, { yearBase: settings.yearBase || {} });
    var byId = {};
    computed.forEach(function (c, i) { byId[c.id != null ? c.id : i] = c; });

    var INPUT_KEYS = {};
    T.FIELDS.forEach(function (f) { INPUT_KEYS[f.key] = 1; });
    var DERIVED = {};
    T.DERIVED_KEYS.forEach(function (k) { DERIVED[k] = 1; });

    var aoa = [COLS.map(function (c) { return c.label; })];
    records.forEach(function (r, i) {
      var c = byId[r.id != null ? r.id : i] || {};
      var row = COLS.map(function (col) {
        if (col.key === 'rate') return (c.rate || 0) > 0 ? (c.rate * 100).toFixed(0) + '%' : '—';
        if (col.key === 'anomalyText') {
          return (c.anomalies || []).filter(function (a) { return a.level !== 'info'; }).map(function (a) { return a.message; }).join('；') || '正常';
        }
        if (col.key === 'note') return r.note || '';
        if (DERIVED[col.key]) return c[col.key] != null ? c[col.key] : 0;
        if (INPUT_KEYS[col.key]) {
          if (PCT_COLS[col.key]) return Math.round(T.num(r[col.key]) * 1e6) / 1e4;
          return T.num(r[col.key]);
        }
        return c[col.key] != null ? c[col.key] : 0;
      });
      aoa.push(row);
    });

    var ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = COLS.map(function (c) { return { wch: c.w }; });
    ws['!freeze'] = { xSplit: '2', ySplit: '1' };
    ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: COLS.length - 1, r: Math.max(aoa.length - 1, 0) } }) };
    styleHeader(ws, COLS.length);

    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '工资数据');

    /* Sheet2 沿用设置 */
    var carry = settings.carry || T.defaultCarry();
    var cAoa = [['字段', '字段名', '是否沿用上月']];
    T.FIELDS.filter(function (f) { return ['income', 'preTax', 'postTax'].indexOf(f.group) >= 0; }).forEach(function (f) {
      cAoa.push([f.key, f.label, carry[f.key] ? '是' : '否']);
    });
    var ws2 = XLSX.utils.aoa_to_sheet(cAoa);
    ws2['!cols'] = [{ wch: 20 }, { wch: 18 }, { wch: 14 }];
    styleHeader(ws2, 3);
    XLSX.utils.book_append_sheet(wb, ws2, '沿用设置');

    /* Sheet3 年度起算基数 */
    var yb = settings.yearBase || {};
    var yAoa = [['年度', '累计收入', '累计社保公积金', '累计专项附加扣除', '累计已预缴税额']];
    Object.keys(yb).sort().forEach(function (y) {
      var b = yb[y];
      yAoa.push([y, T.num(b.income), T.num(b.social), T.num(b.special), T.num(b.taxPaid)]);
    });
    var ws3 = XLSX.utils.aoa_to_sheet(yAoa);
    ws3['!cols'] = [{ wch: 10 }, { wch: 14 }, { wch: 16 }, { wch: 18 }, { wch: 16 }];
    styleHeader(ws3, 5);
    XLSX.utils.book_append_sheet(wb, ws3, '年度起算基数');

    /* Sheet4 填写说明 */
    var notes = [
      ['填写说明'],
      [''],
      ['1. 请填写「工资数据」表中前 ' + INPUT_COUNT + ' 列（浅色区），其后的灰色区为系统计算列，导入时会被忽略，请勿手动修改。'],
      ['2. 「工资所属月」「工资发放月」请填 年-月 格式，例如 2026-01。通常发放月 = 所属月 + 1 个月。'],
      ['3. 三个比例列（社保折算比例、社保综合缴纳比例、公积金个人缴纳比例）按百分数填写，例如社保综合缴纳比例填 10.5 表示 10.5%。'],
      ['4. 「请假扣发」「个人社保」「个人公积金」是公式计算值，模板里已算好；若贵司口径特殊需要手改，直接覆盖即可。'],
      ['5. 「工资条应发合计」「工资条扣税」「工资条实发合计」三列用于与公司工资条比对，没有可留空。'],
      ['6. 「沿用设置」表中「是否沿用上月」填 是/否，可控制新建月份时哪些字段自动带出上月数值。'],
      ['7. 「年度起算基数」用于年中才开始记录的情况：填写该年度最早记录之前已累计的金额，留空表示从 1 月起算。'],
      [''],
      ['计算口径（界面上不展示，仅本表存档备查）'],
      ['请假扣发 = 请假时数 ÷ 标准计薪时数 × (基本工资 + 绩效奖 + 电话津贴)'],
      ['个人社保 = (基本工资 + 绩效奖) × 社保折算比例 × 社保综合缴纳比例'],
      ['个人公积金 = (基本工资 + 绩效奖) × 公积金个人缴纳比例'],
      ['应发合计 = 基本工资 + 住宿补贴 + 绩效奖 + 电话津贴 + 其他(税前) − 漏打卡 − 请假扣发'],
      ['当月收入 = 应发合计（请假扣发已含在应发合计内，不再重复扣减）'],
      ['累计应纳税所得额 = 累计收入 − 5000×月序 − 累计社保公积金 − 累计专项附加扣除'],
      ['本月个税 = 累计应纳税所得额 × 预扣率 − 速算扣除数 − 累计已预缴税额（为负时暂不退税，按 0 计）'],
      ['扣发合计 = 个人社保 + 个人公积金 + 本月个税 + 水电费 + 房租 + 其他扣发'],
      ['实发合计 = 应发合计 − 扣发合计；本次发放 = 实发合计 − 前期已发'],
      ['月序按【工资发放月】在该年度内的序号计（1~12），跨年重新累计。'],
      [''],
      ['新算法（累计预扣预缴）税率表'],
      ['级数', '累计应纳税所得额', '预扣率', '速算扣除数']
    ].concat(T.NEW_BRACKETS.map(function (b) { return [b.level, b.desc, (b.rate * 100).toFixed(0) + '%', b.deduct]; }))
      .concat([
        [''],
        ['旧算法（月度）税率表 —— 仅作参考对比'],
        ['级数', '月应纳税所得额', '税率', '速算扣除数']
      ])
      .concat(T.OLD_BRACKETS.map(function (b) { return [b.level, b.desc, (b.rate * 100).toFixed(0) + '%', b.deduct]; }));

    var ws4 = XLSX.utils.aoa_to_sheet(notes);
    ws4['!cols'] = [{ wch: 14 }, { wch: 26 }, { wch: 12 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws4, '填写说明');

    return wb;
  }

  function styleHeader(ws, nCols) {
    for (var c = 0; c < nCols; c++) {
      for (var r = 1; r < 2000; r++) {
        var addr = XLSX.utils.encode_cell({ c: c, r: r });
        var cell = ws[addr];
        if (!cell) continue;
        if (cell.t === 'n') cell.z = NUMFMT;
      }
    }
  }

  /* 后处理：为表头加上深蓝底白字样式，并冻结前两列与表头行。
     浏览器端用 JSZip 直接改 xlsx 内部 styles.xml，逻辑与 lib/excel.js 一致 */
  function polish(arrayBuffer) {
    return JSZip.loadAsync(arrayBuffer).then(function (zip) {
      var stylesFile = zip.file('xl/styles.xml');
      if (!stylesFile) return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
      return stylesFile.async('string').then(function (styles) {
        var bump = function (xml, tag, snippet) {
          var re = new RegExp('<' + tag + ' count="(\\d+)">');
          var m = xml.match(re);
          if (!m) return { xml: xml, idx: -1 };
          var n = parseInt(m[1], 10);
          var out = xml.replace(re, '<' + tag + ' count="' + (n + 1) + '">');
          out = out.replace('</' + tag + '>', snippet + '</' + tag + '>');
          return { xml: out, idx: n };
        };

        var f = bump(styles, 'fonts',
          '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/><family val="2"/></font>');
        styles = f.xml;
        var p = bump(styles, 'fills',
          '<fill><patternFill patternType="solid"><fgColor rgb="FF4A6C8C"/><bgColor indexed="64"/></patternFill></fill>');
        styles = p.xml;
        var x = bump(styles, 'cellXfs',
          '<xf numFmtId="0" fontId="' + f.idx + '" fillId="' + p.idx + '" borderId="0" xfId="0"' +
          ' applyFont="1" applyFill="1" applyAlignment="1">' +
          '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>');
        styles = x.xml;
        if (x.idx < 0) return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
        zip.file('xl/styles.xml', styles);

        var sheets = ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml'];
        var tasks = sheets.map(function (name) {
          var entry = zip.file(name);
          if (!entry) return Promise.resolve();
          return entry.async('string').then(function (s) {
            s = s.replace(/<c r="([A-Z]+)1"([^>]*)>/g,
              function (m, col, rest) { return '<c r="' + col + '1" s="' + x.idx + '"' + rest + '>'; });
            if (name.indexOf('sheet1.xml') >= 0) {
              s = s.replace(/<sheetView workbookViewId="0"\/>/,
                '<sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" ' +
                'activePane="bottomRight" state="frozen"/><selection pane="bottomRight"/></sheetView>');
            }
            zip.file(name, s);
          });
        });
        return Promise.all(tasks).then(function () {
          return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
        });
      });
    }).catch(function () {
      return arrayBuffer; // 样式增强失败不影响文件内容
    });
  }

  function exportArrayBuffer(records, settings) {
    var wb = buildWorkbook(records, settings);
    var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, compression: true });
    return polish(buf);
  }

  /* ------------------------------------------------------------------ 导入 */
  function toMonth(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0');
    if (typeof v === 'number') {
      if (v > 20000 && v < 60000) {
        var d = XLSX.SSF.parse_date_code(v);
        return d.y + '-' + String(d.m).padStart(2, '0');
      }
      var s = String(Math.round(v));
      if (/^\d{6}$/.test(s)) return s.slice(0, 4) + '-' + s.slice(4);
      return '';
    }
    var s2 = String(v).trim().replace(/[/.年月]/g, '-').replace(/-+/g, '-');
    var m = s2.match(/^(\d{4})-(\d{1,2})/);
    if (m) return m[1] + '-' + String(parseInt(m[2], 10)).padStart(2, '0');
    return '';
  }

  function toNum(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    var n = parseFloat(String(v).replace(/[,，\s¥￥元%]/g, ''));
    return isFinite(n) ? n : 0;
  }

  function roundPct(v) {
    if (!isFinite(v)) return 0;
    return Math.round(v * 1e6) / 1e8;
  }

  function parseWorkbook(arrayBuffer) {
    var wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    var warnings = [];

    var name1 = wb.SheetNames.find(function (n) { return n.replace(/\s/g, '') === '工资数据'; }) || wb.SheetNames[0];
    var rows = XLSX.utils.sheet_to_json(wb.Sheets[name1], { header: 1, defval: '', blankrows: false });
    if (!rows.length) throw new Error('工作表「工资数据」为空');

    var headerIdx = -1;
    for (var i = 0; i < Math.min(rows.length, 10); i++) {
      var line = rows[i].map(function (x) { return String(x || '').trim(); });
      if (line.indexOf('工资所属月') >= 0 && line.indexOf('工资发放月') >= 0) { headerIdx = i; break; }
    }
    if (headerIdx < 0) throw new Error('未找到表头，请使用本程序导出的模板文件');

    var header = rows[headerIdx].map(function (x) { return String(x || '').trim(); });
    var idx = {};
    COLS.forEach(function (c) {
      var p = header.indexOf(c.label);
      if (p >= 0) idx[c.key] = p;
    });
    Object.keys(T.LEGACY_MAP).forEach(function (legacy) {
      if (idx[legacy] !== undefined) return;
      var legacyLabel = { socialInsurance: '五险一金' }[legacy];
      var p = header.indexOf(legacyLabel);
      if (p >= 0) idx[legacy] = p;
    });
    if (idx.periodMonth === undefined && idx.payMonth === undefined) {
      throw new Error('模板缺少「工资所属月」「工资发放月」列');
    }

    var records = [];
    for (var r2 = headerIdx + 1; r2 < rows.length; r2++) {
      var row = rows[r2];
      if (!row) continue;
      var periodMonth = toMonth(row[idx.periodMonth]);
      var payMonth = toMonth(idx.payMonth !== undefined ? row[idx.payMonth] : '');
      if (!periodMonth && !payMonth) continue;
      if (!payMonth) payMonth = T.nextMonth(periodMonth, 1) || periodMonth;
      if (!periodMonth) periodMonth = T.nextMonth(payMonth, -1) || payMonth;

      var rec = T.emptyRecord();
      T.FIELDS.forEach(function (f) {
        if (idx[f.key] === undefined) return;
        if (PCT_COLS[f.key]) {
          rec[f.key] = roundPct(toNum(row[idx[f.key]]));
        } else {
          rec[f.key] = toNum(row[idx[f.key]]);
        }
      });
      if (idx.socialInsurance !== undefined) {
        var legacyTotal = toNum(row[idx.socialInsurance]);
        if (legacyTotal > 0 && T.num(rec.personalSocial) === 0 && T.num(rec.personalFund) === 0) {
          rec.personalSocial = legacyTotal;
          warnings.push('第 ' + (r2 - headerIdx) + ' 行：旧模板「五险一金」¥' + legacyTotal +
            ' 已全部记入「个人社保」。建议核对后拆分到个人社保与个人公积金两列');
        }
      }
      rec.note = idx.note !== undefined ? String(row[idx.note] || '').trim() : '';
      rec.periodMonth = periodMonth;
      rec.payMonth = payMonth;
      records.push(rec);
    }

    /* Sheet2 沿用设置 */
    var carry = null;
    var n2 = wb.SheetNames.find(function (n) { return n.replace(/\s/g, '') === '沿用设置'; });
    if (n2) {
      var r2data = XLSX.utils.sheet_to_json(wb.Sheets[n2], { header: 1, defval: '' });
      var h2 = (r2data[0] || []).map(function (x) { return String(x || '').trim(); });
      var hi = r2data.findIndex(function (row) { return row.some(function (x) { return String(x || '').trim() === '是否沿用上月'; }); });
      if (hi < 0) hi = h2.indexOf('是否沿用上月') >= 0 ? 0 : -1;
      if (hi >= 0) {
        var head = (r2data[hi] || []).map(function (x) { return String(x || '').trim(); });
        var iKey = head.indexOf('字段'), iVal = head.indexOf('是否沿用上月');
        carry = T.defaultCarry();
        for (var i2 = hi + 1; i2 < r2data.length; i2++) {
          var row2 = r2data[i2];
          if (!row2 || !row2[iKey]) continue;
          var k = String(row2[iKey]).trim();
          if (Object.prototype.hasOwnProperty.call(carry, k)) carry[k] = String(row2[iVal] || '').trim() === '是';
        }
      }
    }

    /* Sheet3 年度起算基数 */
    var yearBase = {};
    var n3 = wb.SheetNames.find(function (n) { return n.replace(/\s/g, '') === '年度起算基数'; });
    if (n3) {
      var r3data = XLSX.utils.sheet_to_json(wb.Sheets[n3], { header: 1, defval: '' });
      var hi3 = r3data.findIndex(function (row) { return row.some(function (x) { return String(x || '').trim() === '年度'; }); });
      if (hi3 < 0) hi3 = 0;
      for (var i3 = hi3 + 1; i3 < r3data.length; i3++) {
        var row3 = r3data[i3];
        if (!row3 || row3[0] === '' || row3[0] === null || row3[0] === undefined) continue;
        var y = String(row3[0]).trim();
        if (!/^\d{4}$/.test(y)) continue;
        yearBase[y] = {
          income: toNum(row3[1]), social: toNum(row3[2]),
          special: toNum(row3[3]), taxPaid: toNum(row3[4])
        };
      }
    }

    return { records: records, carry: carry, yearBase: yearBase, warnings: warnings };
  }

  window.ExcelBrowser = {
    COLS: COLS,
    INPUT_COUNT: INPUT_COUNT,
    PCT_COLS: PCT_COLS,
    buildWorkbook: buildWorkbook,
    exportArrayBuffer: exportArrayBuffer,
    parseWorkbook: parseWorkbook,
    toMonth: toMonth,
    toNum: toNum,
    roundPct: roundPct
  };
})();
