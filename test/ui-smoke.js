/* 运行时冒烟：用 jsdom 真实加载页面并执行前端脚本，捕获运行时错误与交互结果。
   纯静态版 —— 不依赖 server.js，vendor 库 + 本地存储层全部在内存里注入。 */
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..');

let pass = 0, fail = 0;
const errors = [];
function ok(name, cond, extra) {
  if (cond) { pass++; console.log('  v ' + name); }
  else { fail++; console.log('  x ' + name + (extra ? '  -> ' + extra : '')); }
}

(async () => {
  const html = fs.readFileSync(path.join(ROOT, 'public/index.html'), 'utf8');
  const dom = new JSDOM(html, {
    url: 'http://localhost/',
    runScripts: 'dangerously',
    resources: undefined,
    pretendToBeVisual: true
  });
  const w = dom.window;

  w.URL.createObjectURL = () => 'blob:x';
  w.URL.revokeObjectURL = () => { };
  w.onerror = (m) => errors.push('window.onerror: ' + m);
  const origErr = console.error;
  w.console = Object.assign({}, console, {
    error: (...a) => { errors.push('console.error: ' + a.join(' ')); origErr(...a); }
  });

  console.log('\n[1] 页面加载与脚本执行');
  // 手动注入脚本（jsdom 不加载外部资源）：vendor 库 → 核心 → 业务层
  const scripts = [
    'public/vendor/xlsx.full.min.js',
    'public/vendor/jszip.min.js',
    'public/shared/tax-core.js',
    'public/js/excel-browser.js',
    'public/js/store-local.js',
    'public/js/app.js'
  ];
  for (const p of scripts) {
    const code = fs.readFileSync(path.join(ROOT, p), 'utf8');
    const s = w.document.createElement('script');
    s.textContent = code;
    w.document.body.appendChild(s);
  }
  await new Promise(r => setTimeout(r, 900));

  ok('TaxCore 已加载', typeof w.TaxCore === 'object');
  ok('SheetJS 已加载', typeof w.XLSX === 'object' && typeof w.XLSX.utils === 'object');
  ok('JSZip 已加载', typeof w.JSZip === 'function');
  ok('ExcelBrowser 已加载', typeof w.ExcelBrowser === 'object');
  ok('StoreLocal 已加载', typeof w.StoreLocal === 'object');
  ok('无运行时错误', errors.length === 0, errors.join(' | '));

  const $ = s => w.document.querySelector(s);
  const $$ = s => Array.from(w.document.querySelectorAll(s));

  console.log('\n[2] 初始渲染');
  ok('收入项字段已生成 9 个', $$('#grid-income input').length === 9, String($$('#grid-income input').length));
  ok('税前扣除字段 6 个', $$('#grid-preTax input').length === 6, String($$('#grid-preTax input').length));
  ok('税后扣除字段 4 个', $$('#grid-postTax input').length === 4, String($$('#grid-postTax input').length));
  ok('比对字段 3 个', $$('#grid-compare input').length === 3, String($$('#grid-compare input').length));
  ok('沿用设置项 19 个', $$('#carry-list input[data-key]').length === 19, String($$('#carry-list input[data-key]').length));
  ok('税率表已生成', $$('#tbl-new tbody tr').length === 7 && $$('#tbl-old tbody tr').length === 7);
  ok('存档计数显示', /已存档 \d+ \/ 24 月/.test($('#archive-chip').textContent), $('#archive-chip').textContent);
  ok('已无存档凭证卡片', !$('#token-value') && !$('#btn-copy-token') && !$('#btn-restore-token'));

  console.log('\n[2b] 新增字段与公式字段就位');
  ok('请假时数在收入项', !!$('#f-leaveHours') && $('#f-leaveHours').closest('#grid-income') !== null);
  ok('标准计薪时数在收入项', !!$('#f-standardHours') && $('#f-standardHours').closest('#grid-income') !== null);
  ok('请假扣发在收入项', !!$('#f-leaveDeduction') && $('#f-leaveDeduction').closest('#grid-income') !== null);
  ok('社保折算比例在税前扣除', !!$('#f-socialRatio') && $('#f-socialRatio').closest('#grid-preTax') !== null);
  ok('社保综合缴纳比例在税前扣除', !!$('#f-socialRate') && $('#f-socialRate').closest('#grid-preTax') !== null);
  ok('公积金个人缴纳比例在税前扣除', !!$('#f-fundRate') && $('#f-fundRate').closest('#grid-preTax') !== null);
  ok('个人社保在税前扣除', !!$('#f-personalSocial') && $('#f-personalSocial').closest('#grid-preTax') !== null);
  ok('个人公积金在税前扣除', !!$('#f-personalFund') && $('#f-personalFund').closest('#grid-preTax') !== null);
  ok('已无「五险一金」字段', !$('#f-socialInsurance'));
  ok('三个派生字段默认只读', ['leaveDeduction', 'personalSocial', 'personalFund']
    .every(k => $('#f-' + k).readOnly), 'readOnly=' + ['leaveDeduction', 'personalSocial', 'personalFund'].map(k => $('#f-' + k).readOnly).join(','));
  ok('派生字段带模式胶囊且默认显示「自动」', ['leaveDeduction', 'personalSocial', 'personalFund']
    .every(k => $('#mode-' + k) && $('#mode-' + k).textContent === '自动'),
    ['leaveDeduction', 'personalSocial', 'personalFund'].map(k => ($('#mode-' + k) || {}).textContent).join(','));
  ok('页面不展示应发合计公式', !/基本工资 \+/.test($('#view-calc').textContent));
  ok('页面不展示社保/公积金公式', !/社保折算比例 ×/.test($('#view-calc').textContent));
  ok('页面不展示请假扣发公式', !/标准计薪时数 ×/.test($('#view-calc').textContent));
  /* 核算视图不得出现公式表达式（累计明细只列金额，不写算式） */
  ok('核算视图不展示个税算式', !/× 预扣率/.test($('#view-calc').textContent),
    $('#view-calc').textContent.match(/.{0,20}× 预扣率.{0,20}/) || '');
  ok('页面不展示「÷」分母式', !/÷/.test($('#view-calc').textContent));

  console.log('\n[3] 只显示当前视图（[hidden] 生效）');
  const visible = ['calc', 'history', 'io', 'settings'].filter(v => !$('#view-' + v).hasAttribute('hidden'));
  ok('仅「核算」视图可见', visible.length === 1 && visible[0] === 'calc', visible.join(','));

  console.log('\n[4] 录入数据 → 实时计算（把发放月设为当年 1 月，即第 1 期）');
  const y0 = new Date().getFullYear();
  $('#pay-year').value = String(y0);
  $('#pay-month').value = '1';
  $('#pay-month').dispatchEvent(new w.Event('change', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('发放月切换生效', $('#pay-month').value === '1' && $('#period-month').value === '12', $('#period-month').value);

  const setVal = (k, v) => {
    const el = $('#f-' + k);
    el.value = String(v);
    el.dispatchEvent(new w.Event('input', { bubbles: true }));
  };
  /* 第一个测试月：比例留空、手填金额不生效（派生字段只读），
     因此统一用「比例驱动」录入，与真实使用方式一致 */
  setVal('base', 14000); setVal('housingSubsidy', 1000); setVal('performance', 1000);
  setVal('phoneAllowance', 50); setVal('specialDeduction', 1500);
  setVal('utility', 200); setVal('rent', 1200);
  setVal('standardHours', 174);
  /* 社保折算比例 60%、社保综合 10.5%、公积金 8%
     → 个人社保 = (14000+1000)×0.6×0.105 = 945
     → 个人公积金 = 15000×8% = 1200 */
  setVal('socialRatio', 60); setVal('socialRate', 10.5); setVal('fundRate', 8);
  await new Promise(r => setTimeout(r, 80));
  ok('应发合计 = 16050', $('#sum-gross').textContent === '¥16,050.00', $('#sum-gross').textContent);
  ok('个人社保自动算出 945', $('#f-personalSocial').value === '945.00', $('#f-personalSocial').value);
  ok('个人公积金自动算出 1200', $('#f-personalFund').value === '1,200.00', $('#f-personalFund').value);
  ok('社保公积金合计 = 2145', $('#sum-social').textContent === '¥2,145.00', $('#sum-social').textContent);
  /* 累计应纳税所得额 = 16050 − 5000 − 2145 − 1500 = 7405 → 税 = 7405×3% = 222.15 */
  ok('本月个税 = 222.15', $('#r-tax').textContent === '¥222.15', $('#r-tax').textContent);
  /* 实发 = 16050 − 2145 − 222.15 − 200 − 1200 = 12282.85 */
  ok('实发合计 = 12282.85', $('#r-net').textContent === '¥12,282.85', $('#r-net').textContent);
  /* 适用预扣率已并入个税副标题（原先会在结果表里单独占一行） */
  ok('适用预扣率并入个税副标题', /适用预扣率 3%/.test($('#r-tax-sub').textContent), $('#r-tax-sub').textContent);
  ok('结果区不再有独立的税率格', $('.result-grid').children.length === 6, String($('.result-grid').children.length));
  ok('比例字段按百分比显示', $('#f-socialRate').value === '10.5', $('#f-socialRate').value);

  console.log('\n[4b] 请假扣发由公式算出并进入应发合计');
  setVal('leaveHours', 8);
  /* 请假扣发 = 8/174 × (14000+1000+50) = 691.95；应发 = 16050 − 691.95 = 15358.05 */
  await new Promise(r => setTimeout(r, 80));
  ok('请假扣发自动算出 691.95', $('#f-leaveDeduction').value === '691.95', $('#f-leaveDeduction').value);
  ok('应发合计随请假下调为 15358.05', $('#sum-gross').textContent === '¥15,358.05', $('#sum-gross').textContent);
  /* 累计应纳税所得额 = 15358.05 − 5000 − 2145 − 1500 = 6713.05 → 税 = 201.39 */
  ok('个税随之下调为 201.39', $('#r-tax').textContent === '¥201.39', $('#r-tax').textContent);

  console.log('\n[4c] 派生字段模式胶囊：默认自动，只有点击才切换');
  const modePill = k => $('#mode-' + k);
  ok('三个派生字段都有模式胶囊',
    ['leaveDeduction', 'personalSocial', 'personalFund'].every(k => !!modePill(k)));
  ok('默认全部显示「自动」',
    ['leaveDeduction', 'personalSocial', 'personalFund'].every(k => modePill(k).textContent === '自动'),
    ['leaveDeduction', 'personalSocial', 'personalFund'].map(k => modePill(k).textContent).join('/'));
  ok('自动模式下输入框只读', $('#f-leaveDeduction').readOnly === true
    && $('#f-personalSocial').readOnly === true && $('#f-personalFund').readOnly === true);

  /* 只读语义：浏览器会阻止用户在只读框输入，这里模拟该行为。
     判定标准是「草稿值未被改写」——直接看页面显示是否仍是公式值。 */
  const beforeLeave = $('#f-leaveDeduction').value;
  $('#f-leaveDeduction').dispatchEvent(new w.KeyboardEvent('keydown', { key: '1', bubbles: true }));
  await new Promise(r => setTimeout(r, 60));
  ok('自动模式下键盘输入被拦截（值未变）', $('#f-leaveDeduction').value === beforeLeave, $('#f-leaveDeduction').value);

  /* 点胶囊 → 手动 */
  modePill('leaveDeduction').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('点击后胶囊显示「手动」', modePill('leaveDeduction').textContent === '手动', modePill('leaveDeduction').textContent);
  ok('手动模式输入框可编辑', $('#f-leaveDeduction').readOnly === false);
  ok('提示「已调整为手动模式」', /已调整为手动模式/.test($('#toast').textContent), $('#toast').textContent);
  $('#f-leaveDeduction').value = '1000';
  $('#f-leaveDeduction').dispatchEvent(new w.Event('input', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  /* 编辑中不格式化（避免光标跳位），失焦后才加千分位 */
  ok('手改后请假扣发取 1000', parseFloat($('#f-leaveDeduction').value.replace(/,/g, '')) === 1000, $('#f-leaveDeduction').value);
  ok('手改后标黄提示手动', $('#f-leaveDeduction').classList.contains('is-manual'),
    $('#f-leaveDeduction').className);
  ok('手改后应发合计随之变化', $('#sum-gross').textContent === '¥15,050.00', $('#sum-gross').textContent);
  ok('手改后出现提示', /手动填写/.test($('#anomaly-box').textContent), $('#anomaly-box').textContent.slice(0, 100));
  $('#f-leaveDeduction').dispatchEvent(new w.Event('blur', { bubbles: true }));
  ok('失焦后格式化为 1,000.00', $('#f-leaveDeduction').value === '1,000.00', $('#f-leaveDeduction').value);

  /* 再点一次 → 自动，数值立刻回到公式 */
  modePill('leaveDeduction').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('再点后胶囊回到「自动」', modePill('leaveDeduction').textContent === '自动', modePill('leaveDeduction').textContent);
  ok('提示「已调整为自动模式」', /已调整为自动模式/.test($('#toast').textContent), $('#toast').textContent);
  ok('切回自动立刻回到公式值 691.95', $('#f-leaveDeduction').value === '691.95', $('#f-leaveDeduction').value);
  ok('切回自动后取消标黄', !$('#f-leaveDeduction').classList.contains('is-manual'));

  console.log('\n[4c-2] 关键回归：改输入项不会让派生字段自己跳成手动');
  /* 过去的缺陷：沿用上月的值与新公式不一致 → 被误判为「手动」并显示旧值。
     现在模式只由点击决定，自动值必须实时等于公式。 */
  setVal('standardHours', 200);
  await new Promise(r => setTimeout(r, 80));
  /* 请假扣发 = 8/200 × 15050 = 602.00 */
  ok('改标准计薪时数后胶囊仍为「自动」', modePill('leaveDeduction').textContent === '自动', modePill('leaveDeduction').textContent);
  ok('改标准计薪时数后请假扣发实时变为 602.00', $('#f-leaveDeduction').value === '602.00', $('#f-leaveDeduction').value);
  setVal('socialRate', 11);
  await new Promise(r => setTimeout(r, 80));
  /* 个人社保 = 15000 × 0.6 × 0.11 = 990.00 */
  ok('改社保比例后个人社保实时变为 990.00', $('#f-personalSocial').value === '990.00', $('#f-personalSocial').value);
  ok('改社保比例后胶囊仍为「自动」', modePill('personalSocial').textContent === '自动', modePill('personalSocial').textContent);
  ok('社保公积金合计随之变为 2190.00', $('#sum-social').textContent === '¥2,190.00', $('#sum-social').textContent);
  setVal('standardHours', 174); setVal('socialRate', 10.5);
  await new Promise(r => setTimeout(r, 80));
  ok('改回原值后请假扣发回到 691.95', $('#f-leaveDeduction').value === '691.95', $('#f-leaveDeduction').value);
  ok('改回原值后个人社保回到 945.00', $('#f-personalSocial').value === '945.00', $('#f-personalSocial').value);

  console.log('\n[4c-3] 手动模式下填 0 也要被尊重（不被公式冲掉）');
  modePill('personalSocial').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 60));
  setVal('personalSocial', 0);
  await new Promise(r => setTimeout(r, 80));
  ok('手动填 0 后社保公积金合计只剩公积金 1200', $('#sum-social').textContent === '¥1,200.00', $('#sum-social').textContent);
  modePill('personalSocial').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('切回自动后个人社保恢复 945.00', $('#f-personalSocial').value === '945.00', $('#f-personalSocial').value);
  ok('切回自动后社保公积金合计恢复 2145.00', $('#sum-social').textContent === '¥2,145.00', $('#sum-social').textContent);

  console.log('\n[4d] 异常校验：标准计薪时数为 0');
  setVal('standardHours', 0);
  await new Promise(r => setTimeout(r, 80));
  ok('提示标准计薪时数缺失', /标准计薪时数/.test($('#anomaly-box').textContent), $('#anomaly-box').textContent.slice(0, 140));
  setVal('standardHours', 174);
  await new Promise(r => setTimeout(r, 80));

  console.log('\n[5] 与公司工资条比对 → 异常告警');
  setVal('psTax', 999);
  await new Promise(r => setTimeout(r, 60));
  const anom = $$('#anomaly-box .anomaly').map(e => e.className);
  ok('检出异常条目', anom.length >= 1, JSON.stringify(anom));
  ok('含 error 级异常', anom.some(c => c.includes('error')));
  ok('异常文案指向多扣', /多扣/.test($('#anomaly-box').textContent), $('#anomaly-box').textContent.slice(0, 120));

  console.log('\n[6] 保存到本地存储（localStorage）');
  $('#btn-save').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 500));
  ok('状态变为已存档', $('#edit-state').textContent.indexOf('已存档') === 0, $('#edit-state').textContent);
  ok('存档计数更新为 1', /已存档 1 \/ 24 月/.test($('#archive-chip').textContent), $('#archive-chip').textContent);

  console.log('\n[6b] 模式不存档：重新载入该月一律回到自动');
  modePill('personalFund').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 60));
  ok('先切到手动作为前提', modePill('personalFund').textContent === '手动', modePill('personalFund').textContent);
  /* 切走再切回，等价于下次打开这个月份 */
  $('#pay-month').value = '2';
  $('#pay-month').dispatchEvent(new w.Event('change', { bubbles: true }));
  await new Promise(r => setTimeout(r, 120));
  $('#pay-month').value = '1';
  $('#pay-month').dispatchEvent(new w.Event('change', { bubbles: true }));
  await new Promise(r => setTimeout(r, 160));
  ok('重新载入后胶囊回到「自动」', modePill('personalFund').textContent === '自动', modePill('personalFund').textContent);
  ok('重新载入后输入框恢复只读', $('#f-personalFund').readOnly === true);
  ok('重新载入后数值为公式值 1,200.00', $('#f-personalFund').value === '1,200.00', $('#f-personalFund').value);

  console.log('\n[7] 切换到明细视图');
  $$('.tab').find(t => t.getAttribute('data-view') === 'history').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 120));
  ok('明细视图可见', !$('#view-history').hasAttribute('hidden'));
  ok('列表渲染 1 行', $$('#hist-list .hrow').length === 1, String($$('#hist-list .hrow').length));
  ok('行内显示社保公积金合计', /社保公积金/.test($('#hist-list').textContent), $('#hist-list').textContent.slice(0, 120));

  console.log('\n[8] 批量删除确认弹窗');
  $('#chk-all').checked = true;
  $('#chk-all').dispatchEvent(new w.Event('change', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('已选 1 条', /已选 1 条/.test($('#sel-count').textContent), $('#sel-count').textContent);
  $('#btn-bulk-delete').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 80));
  ok('弹出二次确认', !$('#modal').hasAttribute('hidden'));
  ok('确认框列出月份', /\d{4}-\d{2}/.test($('#modal-body').textContent), $('#modal-body').textContent);
  ok('按钮为确认删除', $('#modal-ok').textContent.indexOf('确认删除') === 0, $('#modal-ok').textContent);
  $('#modal-cancel').dispatchEvent(new w.Event('click', { bubbles: true }));
  ok('可取消关闭', $('#modal').hasAttribute('hidden'));

  console.log('\n[9] 导入导出视图');
  $$('.tab').find(t => t.getAttribute('data-view') === 'io').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 100));
  ok('导出范围默认按月', $('#export-scope .seg-btn.is-active').getAttribute('data-scope') === 'month');
  ok('月份下拉已填充', $$('#export-value option').length >= 1, String($$('#export-value option').length));
  $$('#export-scope .seg-btn').find(b => b.getAttribute('data-scope') === 'all')
    .dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 60));
  ok('切到全部时隐藏选择行', $('#scope-value-row').hasAttribute('hidden'));

  console.log('\n[10] 设置页');
  $$('.tab').find(t => t.getAttribute('data-view') === 'settings').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 100));
  const carryKeys = $$('#carry-list input[data-key]').map(i => i.getAttribute('data-key'));
  ok('沿用策略默认勾选基本工资', $$('#carry-list input[data-key]').find(i => i.getAttribute('data-key') === 'base').checked);
  ok('绩效奖默认不勾选', !$$('#carry-list input[data-key]').find(i => i.getAttribute('data-key') === 'performance').checked);
  ok('沿用列表含「请假时数」', carryKeys.includes('leaveHours'));
  ok('沿用列表含「标准计薪时数」', carryKeys.includes('standardHours'));
  ok('沿用列表含「社保折算比例」', carryKeys.includes('socialRatio'));
  ok('沿用列表含「社保综合缴纳比例」', carryKeys.includes('socialRate'));
  ok('沿用列表含「公积金个人缴纳比例」', carryKeys.includes('fundRate'));
  ok('标准计薪时数默认勾选沿用', $$('#carry-list input[data-key]').find(i => i.getAttribute('data-key') === 'standardHours').checked);
  ok('请假时数默认不勾选沿用', !$$('#carry-list input[data-key]').find(i => i.getAttribute('data-key') === 'leaveHours').checked);
  ok('社保折算比例默认勾选沿用', $$('#carry-list input[data-key]').find(i => i.getAttribute('data-key') === 'socialRatio').checked);
  ok('个税税率表页无计算明细公式', !/本月个税 = 累计应纳税所得额/.test($('#view-settings').textContent));

  console.log('\n[11] 布局编排：语义分组 / 无残行 / 标题行控件');
  $$('.tab').find(t => t.getAttribute('data-view') === 'calc').dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 100));

  /* 分组结构：子组标题 + 组内网格 */
  const subs = $$('#grid-income .subgroup, #grid-preTax .subgroup, #grid-postTax .subgroup, #grid-compare .subgroup');
  ok('四个字段区都被拆成子组', subs.length === 8, String(subs.length));

  const keysIn = g => Array.from(g.querySelectorAll('.field input')).map(i => i.getAttribute('data-key'));
  const colsOf = g => Number(Array.from(g.classList).filter(c => /^g\d$/.test(c))[0].slice(1));
  const joinKeys = k => k.join(',');

  const leaveGroup = $$('#grid-income .g').filter(g => keysIn(g).indexOf('leaveDeduction') >= 0)[0];
  const socialGroup = $$('#grid-preTax .g').filter(g => keysIn(g).indexOf('personalSocial') >= 0)[0];
  const fundGroup = $$('#grid-preTax .g').filter(g => keysIn(g).indexOf('personalFund') >= 0)[0];
  ok('请假三件套同一组', joinKeys(keysIn(leaveGroup)) === 'leaveHours,standardHours,leaveDeduction',
    joinKeys(keysIn(leaveGroup)));
  ok('社保三件套同一组', joinKeys(keysIn(socialGroup)) === 'socialRatio,socialRate,personalSocial',
    joinKeys(keysIn(socialGroup)));
  ok('公积金两项同一组', joinKeys(keysIn(fundGroup)) === 'fundRate,personalFund', joinKeys(keysIn(fundGroup)));

  /* 零残行：每组「成员数 = 列数整数倍」，否则末行会留大片空白 */
  const remainders = $$('.fields .g').map(g => {
    const cols = colsOf(g);
    return { cols: cols, n: g.children.length, rest: g.children.length % cols };
  }).filter(x => x.rest !== 0);
  ok('每个组的成员数都能整除列数（无残行）', remainders.length === 0, JSON.stringify(remainders));

  /* 每格只有「标题行 + 输入框行」两行，说明文字不再占第三行 */
  ok('每格结构统一为标题行 + 输入框', $$('.fields .field').every(f => f.children.length === 2), '');
  ok('页面已不再压板式灰字说明', $$('.fields .field .tip').length === 0, String($$('.fields .field .tip').length));
  ok('说明改为 ⓘ 浮层', $$('.fields .field .hint-btn').length > 0, String($$('.fields .field .hint-btn').length));
  ok('ⓘ 浮层默认收起', $$('.fields .hint-pop.show').length === 0, '');
  $('[id="f-socialRatio"]').closest('.field').querySelector('.hint-btn')
    .dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 50));
  ok('点击 ⓘ 显示说明', $$('.fields .hint-pop.show').length === 1, '');
  ok('ⓘ 说明不含计算公式', !/[×*÷]|基本工资\s*\+|时数\s*\//.test($('.fields .hint-pop.show').textContent),
    ($('.fields .hint-pop.show') || {}).textContent);

  /* 模式胶囊落在标题行里，不占额外高度 */
  ok('模式胶囊在标题行内', $$('.mode-pill').length === 3 && $$('.mode-pill').every(b => b.closest('.field-head')),
    String($$('.mode-pill').length));

  console.log('\n[12] 月份选择：单位外置、框内纯数字');
  ok('年份选项为纯数字', $$('#period-year option').every(o => /^\d+$/.test(o.textContent)), '');
  ok('月份选项为纯数字', $$('#period-month option').every(o => /^\d+$/.test(o.textContent)), '');
  ok('年单位在框外', $('#period-year').closest('.sel-cell').querySelector('.sel-unit').textContent === '年');
  ok('月单位在框外', $('#period-month').closest('.sel-cell').querySelector('.sel-unit').textContent === '月');
  ok('单位不在框内', $('#period-year').textContent.indexOf('年') < 0, $('#period-year').textContent.slice(0, 20));

  console.log('\n[13] 自定义下拉组件');
  ok('全部原生下拉都被组件化', $$('select').every(s => s.closest('.dd') && s.classList.contains('dd-src')),
    String($$('select').filter(s => !s.closest('.dd')).length));
  ok('原生下拉已视觉隐藏', $$('select.dd-src').length === 5, String($$('select.dd-src').length));
  ok('每个下拉都有按钮与面板', $$('.dd').every(d => d.querySelector('.dd-btn') && d.querySelector('.dd-panel')), '');
  ok('面板默认收起', $$('.dd-panel.show').length === 0, '');

  const ddb = $('#pay-year').closest('.dd').querySelector('.dd-btn');
  ddb.dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 50));
  const panel = $('#pay-year').closest('.dd').querySelector('.dd-panel');
  ok('点击按钮弹出面板', panel.classList.contains('show'), '');
  ok('列表项数与选项数一致', panel.querySelectorAll('.dd-item').length === $('#pay-year').options.length,
    panel.querySelectorAll('.dd-item').length + '/' + $('#pay-year').options.length);
  ok('当前选中项带标记', !!panel.querySelector('.dd-item.sel'), '');
  ok('按钮文字与选中项一致', ddb.textContent === '2026', ddb.textContent);

  const target = panel.querySelectorAll('.dd-item')[2];
  target.dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 60));
  ok('点选后原生 select 值同步', $('#pay-year').value === target.getAttribute('data-value'),
    $('#pay-year').value + '/' + target.getAttribute('data-value'));
  ok('点选后按钮文字同步', ddb.textContent === target.textContent, ddb.textContent);
  ok('点选后面板收起', !panel.classList.contains('show'), '');

  ddb.dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 40));
  w.document.body.dispatchEvent(new w.Event('click', { bubbles: true }));
  await new Promise(r => setTimeout(r, 40));
  ok('点击空白处收起面板', !panel.classList.contains('show'), '');

  ok('全程无运行时错误', errors.length === 0, errors.join(' | '));

  console.log('\n' + '='.repeat(56) + `\n通过 ${pass} 项，失败 ${fail} 项\n` + '='.repeat(56) + '\n');
  dom.window.close();
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
