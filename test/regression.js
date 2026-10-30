/* 回归验证：用《工资条.xls》《个税计算逻辑.xls》的真实样本逐月比对 */
const T = require('../shared/tax-core.js');

let pass = 0, fail = 0;
function eq(name, actual, expected, tol) {
  tol = tol === undefined ? 0.005 : tol;
  const ok = Math.abs(actual - expected) <= tol;
  if (ok) { pass++; console.log(`  v ${name}: ${actual}`); }
  else { fail++; console.log(`  x ${name}: 实际 ${actual} / 期望 ${expected}`); }
}

/* ------------------------------------------------------------------ */
console.log('\n【样本 A】工资条.xls 7 个月（2018-12 所属 ~ 2019-06 所属，2019-01~07 发放）');
console.log('  基本工资 14000 / 住宿补贴 1000 / 绩效奖 1000 / 电话津贴 50');
console.log('  社保公积金合计 999.3（按新口径拆分：个人社保 800 + 个人公积金 199.3）');
console.log('  专项附加 1500；请假沿用原表的请假扣发金额（手改覆盖）\n');

/* 原表只有一个「五险一金 999.3」，新口径拆成个人社保 + 个人公积金。
   这里按 800 / 199.3 拆分，两者之和仍是 999.3，
   因此累计社保公积金、个税、实发全部与拆分前一致 —— 这正是本次改造要保住的等价性。 */
const SOC = 800, FUND = 199.3;

const A = [
  // 所属月,   发放月,  基本, 住宿, 绩效, 电话, 其他税前, 漏打卡, 请假, 水电, 房租, 其他扣发, 前期已发, 期望个税
  ['2018-12', '2019-01', 14000, 1000, 1000, 50, 0, 0, 0, 200, 1200, 0, 0, 256.52],
  ['2019-01', '2019-02', 14000, 1000, 1000, 50, 0, 0, 0, 250, 1200, 0, 7000, 256.52],
  ['2019-02', '2019-03', 14000, 1000, 1000, 50, 0, 0, 184.48, 200, 1200, 0, 0, 250.99],
  ['2019-03', '2019-04', 14000, 1000, 1000, 50, 20000, 10, 0, 150, 1300, 0, 0, 2116.80],
  ['2019-04', '2019-05', 14000, 1000, 1000, 50, 0, 0, 46.12, 300, 1300, 0, 0, 850.46],
  ['2019-05', '2019-06', 15000, 1000, 2000, 50, 2000, 0, 0, 400, 1300, 0, 0, 1255.07],
  ['2019-06', '2019-07', 15000, 1000, 2000, 50, 0, 0, 103.74, 400, 1300, 100, 0, 1044.70]
];

const recA = A.map((r, i) => ({
  id: 'a' + i, periodMonth: r[0], payMonth: r[1],
  base: r[2], housingSubsidy: r[3], performance: r[4], phoneAllowance: r[5],
  otherPreTax: r[6], missedPunch: r[7], leaveDeduction: r[8],
  /* 请假时数与标准计薪时数留空（会触发 stdHours 提示），请假扣发按原表金额手动覆盖 */
  leaveHours: 0, standardHours: 0,
  /* 三个人为 0，公式算不出值 → 视为手动录入，不标记为覆盖 */
  socialRatio: 0, socialRate: 0, fundRate: 0,
  personalSocial: SOC, personalFund: FUND,
  specialDeduction: 1500,
  utility: r[9], rent: r[10], otherDeduction: r[11], advancePaid: r[12],
  psGross: 0, psTax: 0, psNet: 0
}));

const outA = T.compute(recA, {});
const expNet = [13394.18, 13344.18, 13215.23, 31473.90, 12554.12, 16095.63, 14102.26];
// 注：原表第 7 月自身不一致——工资条.xls 请假扣发为 103.74，但其个税 1050.17 实为
// 按个税计算逻辑.xls 的 48.99 算得。此处按工资条的 103.74 口径期望 834.70，
// 另在【样本 A2】中用 48.99 口径交叉验证原表的 1050.17 / 840.17。
const expOld = [645.07, 645.07, 626.62, 4475.18, 640.46, 1100.14, 834.70];

outA.forEach((o, i) => {
  console.log(`  ── ${o.periodMonth} 所属 / ${o.payMonth} 发放（第 ${o.seq} 期）──`);
  /* 新口径的应发合计 = 旧口径 − 当月请假扣发 */
  eq('应发合计', o.gross, [16050, 16050, 15865.52, 36040, 16003.88, 20050, 17946.26][i]);
  eq('本月个税', o.tax, A[i][13]);
  eq('实发合计', o.netPay, expNet[i]);
  eq('旧算法参考', o.oldTax, expOld[i]);
  eq('累计已缴税额', o.cumTaxPaid, [256.52, 513.04, 764.03, 2880.83, 3731.29, 4986.36, 6031.06][i]);
});

eq('全年个税合计', outA.reduce((s, o) => s + o.tax, 0), 6031.06);

/* ------------------------------------------------------------------ */
console.log('\n【样本 A2】第 7 月改用个税计算逻辑.xls 的请假扣发 48.99（交叉验证原表 1050.17 / 840.17）');
const recA2 = recA.map(r => Object.assign({}, r));
recA2[6].leaveDeduction = 48.99;
const outA2 = T.compute(recA2, {});
eq('本月个税', outA2[6].tax, 1050.17);
eq('旧算法参考', outA2[6].oldTax, 840.17);
eq('累计收入', outA2[6].cumIncome, 138060.41);

/* ------------------------------------------------------------------ */
console.log('\n【样本 B】个税计算逻辑.xls 首行（应发 160500，验证高档税率与旧算法）');
const recB = [{
  id: 'b0', periodMonth: '2018-12', payMonth: '2019-01',
  base: 160500, housingSubsidy: 0, performance: 0, phoneAllowance: 0,
  otherPreTax: 0, missedPunch: 0, leaveDeduction: 0,
  leaveHours: 0, standardHours: 0, socialRatio: 0, socialRate: 0, fundRate: 0,
  personalSocial: 999.3, personalFund: 0,
  specialDeduction: 1500,
  utility: 0, rent: 0, otherDeduction: 0, advancePaid: 0, psGross: 0, psTax: 0, psNet: 0
}];
const outB = T.compute(recB, {});
eq('累计应纳税所得额', outB[0].cumTaxable, 153000.7);
eq('本月应预扣预缴税额', outB[0].tax, 13680.14);
eq('旧算法个税', outB[0].oldTax, 53690.32);
eq('适用预扣率', outB[0].rate, 0.20);

/* ------------------------------------------------------------------ */
console.log('\n【样本 C】跨年重置 + 负税额归零 + 异常检测');
const blank = () => ({
  housingSubsidy: 0, performance: 0, phoneAllowance: 0, otherPreTax: 0, missedPunch: 0,
  leaveDeduction: 0, leaveHours: 0, standardHours: 0,
  socialRatio: 0, socialRate: 0, fundRate: 0, personalSocial: 0, personalFund: 0,
  specialDeduction: 0, utility: 0, rent: 0, otherDeduction: 0, advancePaid: 0,
  psGross: 0, psTax: 0, psNet: 0
});
const recC = [
  Object.assign(blank(), { id: 'c1', periodMonth: '2019-11', payMonth: '2019-12', base: 3000 }),
  Object.assign(blank(), { id: 'c2', periodMonth: '2019-12', payMonth: '2020-01', base: 3000 })
];
const outC = T.compute(recC, {});
eq('2019-12（第12期）累计减除费用', outC[0].cumBasic, 60000);
eq('2020-01 月序重置为', outC[1].seq, 1);
eq('2020-01 累计收入重置为', outC[1].cumIncome, 3000);
eq('低收入个税为 0', outC[1].tax, 0);

const recD = [Object.assign(blank(), {
  id: 'd1', periodMonth: '2020-02', payMonth: '2020-03',
  base: 14000, personalSocial: 999.3, specialDeduction: 1500,
  psGross: 16000, psTax: 999, psNet: 12000
})];
const outD = T.compute(recD, {});
const codes = outD[0].anomalies.map(a => a.code);
console.log('  检出异常码:', codes.join(', '));
eq('检出 应发合计不符', codes.includes('grossDiff') ? 1 : 0, 1);
eq('检出 扣税不符', codes.includes('taxDiff') ? 1 : 0, 1);
eq('检出 实发不符', codes.includes('netDiff') ? 1 : 0, 1);

/* ------------------------------------------------------------------ */
console.log('\n【样本 E】三个派生公式（本次新增，含覆盖与解锁语义）');

/* E1 请假扣发 = 请假时数 / 标准计薪时数 × (基本 + 绩效 + 电话津贴) */
const e1 = Object.assign(blank(), { base: 14000, performance: 1000, phoneAllowance: 50, leaveHours: 8, standardHours: 174 });
eq('请假扣发 8/174×15050', T.calcLeaveDeduction(e1), 691.95);
const e1b = Object.assign(blank(), { base: 14000, performance: 1000, phoneAllowance: 50, leaveHours: 0, standardHours: 174 });
eq('请假 0 小时 → 扣发 0', T.calcLeaveDeduction(e1b), 0);
const e1c = Object.assign(blank(), { base: 14000, performance: 1000, phoneAllowance: 50, leaveHours: 8, standardHours: 0 });
eq('标准计薪时数 0 → 不产生扣发（界面另有报错提示）', T.calcLeaveDeduction(e1c), 0);
const e1d = Object.assign(blank(), { base: 14000, performance: 1000, phoneAllowance: 50, leaveHours: 240, standardHours: 174 });
eq('请假 240 小时（整月）', T.calcLeaveDeduction(e1d), 20758.62);

/* E2 个人社保 = (基本 + 绩效) × 折算比例 × 综合比例 */
const e2 = Object.assign(blank(), { base: 14000, performance: 1000, socialRatio: 0.6, socialRate: 0.105 });
eq('个人社保 15000×0.6×0.105', T.calcPersonalSocial(e2), 945);
const e2b = Object.assign(blank(), { base: 14000, performance: 1000, socialRatio: 1, socialRate: 0.105 });
eq('折算比例=1 时等同于直接乘以综合比例', T.calcPersonalSocial(e2b), 1575);
const e2c = Object.assign(blank(), { base: 14000, performance: 1000, socialRatio: 0, socialRate: 0.105 });
eq('折算比例为 0 → 个人社保 0（界面另有提示）', T.calcPersonalSocial(e2c), 0);

/* E3 个人公积金 = (基本 + 绩效) × 公积金比例 */
const e3 = Object.assign(blank(), { base: 14000, performance: 1000, fundRate: 0.08 });
eq('个人公积金 15000×8%', T.calcPersonalFund(e3), 1200);

/* E4 派生字段的「公式值 / 手改覆盖」判定 */
const e4 = Object.assign(blank(), {
  base: 14000, performance: 1000, phoneAllowance: 50, leaveHours: 8, standardHours: 174,
  socialRatio: 0.6, socialRate: 0.105, fundRate: 0.08
});
const rv0 = T.resolveDerived(e4);
eq('公式可算时留言空值 → 取公式值', rv0.resolved.leaveDeduction, 691.95);
eq('未被标记为手动覆盖', rv0.covered.leaveDeduction ? 1 : 0, 0);

const e5 = Object.assign({}, e4, { leaveDeduction: 500 });
const rv1 = T.resolveDerived(e5);
eq('手填 500 与公式 691.95 不符 → 保留 500', e5.leaveDeduction, 500);
eq('resolved 里公式值仍是 691.95', rv1.resolved.leaveDeduction, 691.95);
eq('被标记为手动覆盖', rv1.covered.leaveDeduction ? 1 : 0, 1);
const e6 = Object.assign({}, e4, { personalSocial: 999.3 });
const rv2 = T.resolveDerived(e6);
eq('手改个人社保 999.3 → 保留', e6.personalSocial, 999.3);
eq('个人社保标记为覆盖', rv2.covered.personalSocial ? 1 : 0, 1);
eq('未手改的个人公积金仍取公式值', e6.personalFund, 1200);
eq('个人公积金未被标记覆盖', rv2.covered.personalFund ? 1 : 0, 0);

/* 比例未填时公式算不出值，此时手填的三个金额应被保留且不标覆盖 */
const e6b = Object.assign(blank(), {
  base: 14000, performance: 1000, personalSocial: 999.3, personalFund: 199,
  leaveDeduction: 184.48
});
const rv3 = T.resolveDerived(e6b);
eq('公式算不出时保留手填的请假扣发', e6b.leaveDeduction, 184.48);
eq('公式算不出时不标记为覆盖', rv3.covered.leaveDeduction ? 1 : 0, 0);
eq('公式算不出时保留手填的个人社保', e6b.personalSocial, 999.3);

/* E5 全链路：请假 + 比例驱动，逐项手算核对
   注意月序按【发放月】计：下面用 2026-01 发放，月序 = 1，累计减除费用 5000 */
console.log('\n【样本 F】全链路：请假/比例公式驱动，逐项手算核对（发放月 2026-01，第 1 期）');
const e7 = Object.assign(blank(), {
  id: 'f1', periodMonth: '2025-12', payMonth: '2026-01',
  base: 14000, housingSubsidy: 1000, performance: 2000, phoneAllowance: 50,
  leaveHours: 16, standardHours: 174,
  socialRatio: 0.6, socialRate: 0.105, fundRate: 0.08,
  specialDeduction: 1500
});
const outE7 = T.compute([e7], {});
const o = outE7[0];
eq('月序 = 1', o.seq, 1);
// 请假扣发 = 16/174 × 16050 = 1475.862... → 1475.86
eq('请假扣发', o.leaveDeduction, 1475.86);
// 个人社保 = 16000 × 0.6 × 0.105 = 1008
eq('个人社保', o.personalSocial, 1008);
// 个人公积金 = 16000 × 0.08 = 1280
eq('个人公积金', o.personalFund, 1280);
eq('本月社保公积金合计', o.socialTotal, 2288);
// 应发合计 = 14000+1000+2000+50 − 0 − 1475.86 = 15574.14
eq('应发合计（已含请假扣发）', o.gross, 15574.14);
eq('当月收入 = 应发合计（不再重复扣请假）', o.monthlyIncome, 15574.14);
eq('累计社保公积金', o.cumSocial, 2288);
// 累计应纳税所得额 = 15574.14 − 5000 − 2288 − 1500 = 6786.14
eq('累计应纳税所得额', o.cumTaxable, 6786.14);
// 个税 = 6786.14 × 3% = 203.58
eq('本月个税', o.tax, 203.58);
// 扣发合计 = 2288 + 203.58 = 2491.58
eq('扣发合计', o.totalDeduction, 2491.58);
// 实发合计 = 15574.14 − 2491.58 = 13082.56
eq('实发合计', o.netPay, 13082.56);

/* E6 覆盖一致性：手改成与公式不同的值，实发与个税应随之变化。
   注意：改成 0 不构成覆盖（0 与空值无法区分，会退回公式值），
   所以这里改成别的非零值来验证覆盖确实生效。 */
const e8 = Object.assign({}, e7, {
  id: 'f2',
  leaveDeduction: 2000,   // 公式 1475.86 → 覆盖为 2000
  personalSocial: 800,    // 公式 1008   → 覆盖为 800
  personalFund: 1000      // 公式 1280   → 覆盖为 1000
});
const o8 = T.compute([e8], {})[0];
// 应发合计 = 17050 − 2000 = 15050
eq('覆盖后应发合计', o8.gross, 15050);
eq('覆盖后本月社保公积金合计', o8.socialTotal, 1800);
eq('个人社保标记为覆盖', o8.covered.personalSocial ? 1 : 0, 1);
eq('请假扣发标记为覆盖', o8.covered.leaveDeduction ? 1 : 0, 1);
// 累计应纳税所得额 = 15050 − 5000 − 1800 − 1500 = 6750
eq('覆盖后累计应纳税所得额', o8.cumTaxable, 6750);
// 个税 = 6750 × 3% = 202.50
eq('覆盖后本月个税', o8.tax, 202.50);

/* 公式算不出值时（缺时数 / 缺比例），填多少就保留多少 */
const e9 = Object.assign(blank(), {
  id: 'f3', periodMonth: '2025-12', payMonth: '2026-01',
  base: 14000, housingSubsidy: 1000, performance: 1000, phoneAllowance: 50,
  leaveDeduction: 184.48, personalSocial: 999.3
});
const o9 = T.compute([e9], {})[0];
eq('无公式时保留手填请假扣发', o9.leaveDeduction, 184.48);
eq('无公式时保留手填个人社保', o9.personalSocial, 999.3);
eq('无公式时个人公积金为 0', o9.personalFund, 0);
eq('无公式时社保公积金合计', o9.socialTotal, 999.3);
// 应发合计 = 14000+1000+1000+50 − 184.48 = 15865.52
eq('无公式时应发合计', o9.gross, 15865.52);

/* ------------------------------------------------------------------ */
console.log('\n【样本 A3】新口径等价性：请假扣发移入应发合计后，实发与个税必须与原来一致');
/* 旧口径：应发 = 各项 − 漏打卡；当月收入 = 应发 − 请假；扣发合计含请假
   新口径：应发 = 各项 − 漏打卡 − 请假；当月收入 = 应发；扣发合计不含请假
   两者数学上等价，此处用样本 A 的 7 个月逐条比对 netPay 与 tax 的合计值 */
const grossSum = outA.reduce((s, o) => s + o.gross, 0);
const netSum = outA.reduce((s, o) => s + o.netPay, 0);
const taxSum = outA.reduce((s, o) => s + o.tax, 0);
// 这 7 个月的应发合计比旧口径各少了当月的请假扣发 0/0/184.48/0/46.12/0/103.74，共 334.34
eq('应发合计合计（新口径，已扣请假）', grossSum, 138005.66);
eq('实发合计合计（与旧口径完全一致）', netSum, 114179.50);
eq('个税合计（与旧口径完全一致）', taxSum, 6031.06);
// 实发合计 = 应发合计 − 扣发合计，逐条自洽
outA.forEach((o, i) => {
  eq('第' + (i + 1) + '期 实发 = 应发 − 扣发', T.round2(o.gross - o.totalDeduction), o.netPay);
});

/* ------------------------------------------------------------------ */
console.log('\n【样本 D】四舍五入边界');
eq('53690.315 → ', T.round2(153000.7 * 0.45 - 15160), 53690.32);
eq('2655.8199999999997 → ', T.round2(2655.8199999999997), 2655.82);
eq('1710.1400000000012 → ', T.round2(1710.1400000000012), 1710.14);

console.log(`\n${'='.repeat(56)}\n通过 ${pass} 项，失败 ${fail} 项\n${'='.repeat(56)}\n`);
process.exit(fail ? 1 : 0);
