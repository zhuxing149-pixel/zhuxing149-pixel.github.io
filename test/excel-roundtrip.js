/* 浏览器端 Excel 导出→导入 往返一致性测试（纯静态，不依赖服务器）
   验证 ExcelBrowser.exportArrayBuffer 产出 .xlsx，再用 parseWorkbook 读回，
   字段值、设置、年份基数都应完整还原。 */
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
let pass = 0, fail = 0;
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  v ' + name); }
  else { fail++; console.log('  x ' + name + (extra ? '  -> ' + extra : '')); }
}

(async () => {
  const dom = new JSDOM('<body></body>', { runScripts: 'dangerously', pretendToBeVisual: true });
  const w = dom.window;
  for (const p of ['public/vendor/xlsx.full.min.js', 'public/vendor/jszip.min.js', 'public/shared/tax-core.js', 'public/js/excel-browser.js']) {
    const code = fs.readFileSync(path.join(ROOT, p), 'utf8');
    const s = w.document.createElement('script');
    s.textContent = code;
    w.document.body.appendChild(s);
  }
  await new Promise(r => setTimeout(r, 800));

  const T = w.TaxCore;
  const E = w.ExcelBrowser;

  // 构造两条记录
  const rec1 = Object.assign(T.emptyRecord(), {
    periodMonth: '2025-12', payMonth: '2026-01', base: 14000, housingSubsidy: 1000,
    performance: 1000, phoneAllowance: 50, otherPreTax: 0, missedPunch: 0,
    leaveHours: 8, standardHours: 174, socialRatio: 0.6, socialRate: 0.105,
    fundRate: 0.08, specialDeduction: 1500, utility: 200, rent: 1200,
    otherDeduction: 0, advancePaid: 0, psGross: 0, psTax: 0, psNet: 0, note: '一月工资'
  });
  const rec2 = Object.assign(T.emptyRecord(), {
    periodMonth: '2026-01', payMonth: '2026-02', base: 14000, housingSubsidy: 1000,
    performance: 2000, phoneAllowance: 50, otherPreTax: 0, missedPunch: 0,
    leaveHours: 0, standardHours: 174, socialRatio: 0.6, socialRate: 0.105,
    fundRate: 0.08, specialDeduction: 1500, utility: 200, rent: 1200,
    otherDeduction: 0, advancePaid: 0, psGross: 0, psTax: 0, psNet: 0, note: '二月工资'
  });

  const settings = {
    carry: T.defaultCarry(),
    yearBase: { '2026': { income: 30000, social: 3000, special: 3000, taxPaid: 500 } }
  };

  console.log('\n[1] 导出（buildWorkbook + XLSX.write，绕过样式增强）');
  // 注：exportArrayBuffer 里的 polish() 用浏览器版 jszip 做样式增强，
  //     在 jsdom 里 jszip 的 Promise 与 Node 事件循环不同步会挂起（真实浏览器无此问题）。
  //     这里直接用 buildWorkbook + XLSX.write 验证「数据往返」——polish 只改表头样式，不动数据。
  const wb = E.buildWorkbook([rec1, rec2], settings);
  ok('生成 workbook，4 个 sheet', wb.SheetNames.length === 4, wb.SheetNames.join(','));
  const buf = w.XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, compression: true });
  ok('导出为 ArrayBuffer', buf instanceof ArrayBuffer || (buf && buf.constructor && buf.constructor.name === 'ArrayBuffer') || (buf instanceof Uint8Array), buf && buf.constructor && buf.constructor.name);
  ok('导出体积 > 4KB', buf.byteLength > 4096, String(buf.byteLength));

  console.log('\n[2] 导入（读回）');
  const parsed = E.parseWorkbook(buf);
  ok('读回 2 条记录', parsed.records.length === 2, String(parsed.records.length));

  const r1 = parsed.records.find(r => r.payMonth === '2026-01');
  const r2 = parsed.records.find(r => r.payMonth === '2026-02');
  ok('月份正确', !!r1 && !!r2);
  ok('基本工资 14000 往返一致', T.num(r1.base) === 14000, String(r1.base));
  ok('绩效奖 1000 往返一致', T.num(r1.performance) === 1000, String(r1.performance));
  ok('请假时数 8 往返一致', T.num(r1.leaveHours) === 8, String(r1.leaveHours));
  ok('标准计薪时数 174 往返一致', T.num(r1.standardHours) === 174, String(r1.standardHours));
  ok('备注往返一致', r1.note === '一月工资', r1.note);
  ok('二月绩效 2000 往返一致', T.num(r2.performance) === 2000, String(r2.performance));

  console.log('\n[3] 百分比列还原（模板 10.5 → 内部 0.105）');
  ok('社保折算比例 0.6 还原', Math.abs(T.num(r1.socialRatio) - 0.6) < 1e-6, String(r1.socialRatio));
  ok('社保综合 0.105 还原', Math.abs(T.num(r1.socialRate) - 0.105) < 1e-6, String(r1.socialRate));
  ok('公积金 0.08 还原', Math.abs(T.num(r1.fundRate) - 0.08) < 1e-6, String(r1.fundRate));

  console.log('\n[4] 沿用设置还原');
  ok('沿用设置已读回', parsed.carry !== null);
  ok('基本工资沿用标志还原', parsed.carry.base === true, String(parsed.carry.base));
  ok('绩效奖沿用标志还原（默认不勾）', parsed.carry.performance === false, String(parsed.carry.performance));

  console.log('\n[5] 年度起算基数还原');
  ok('yearBase 含 2026', !!parsed.yearBase['2026']);
  ok('累计收入 30000 还原', T.num(parsed.yearBase['2026'].income) === 30000, String(parsed.yearBase['2026'].income));
  ok('累计已缴税 500 还原', T.num(parsed.yearBase['2026'].taxPaid) === 500, String(parsed.yearBase['2026'].taxPaid));

  console.log('\n[6] 派生字段（公式值）在导出时写入');
  // 导出用 compute 算出的公式值，导入时应能读到（个人社保 = (14000+1000)*0.6*0.105 = 945）
  ok('个人社保公式值写入并读回 945', Math.abs(T.num(r1.personalSocial) - 945) < 0.01, String(r1.personalSocial));
  ok('个人公积金公式值写入并读回 1200', Math.abs(T.num(r1.personalFund) - 1200) < 0.01, String(r1.personalFund));

  console.log('\n[7] 空模板（下载导入模板用）');
  const emptyWb = E.buildWorkbook([], { carry: T.defaultCarry(), yearBase: {} });
  ok('空模板可构建', !!emptyWb);
  const emptyBuf = w.XLSX.write(emptyWb, { bookType: 'xlsx', type: 'array', compression: true });
  ok('空模板可导出', emptyBuf.byteLength > 0);
  const emptyParsed = E.parseWorkbook(emptyBuf);
  ok('空模板读回 0 条', emptyParsed.records.length === 0, String(emptyParsed.records.length));

  console.log('\n' + '='.repeat(56) + `\n通过 ${pass} 项，失败 ${fail} 项\n` + '='.repeat(56) + '\n');
  dom.window.close();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
