/* =============================================================================
 * 浏览器本地存储层（纯静态版）
 * 用 localStorage 替代原来的服务端文件存储，数据只留在这台设备的浏览器里。
 * 接口与原来的 lib/store.js 保持一致（load / save / upsert / remove / reset），
 * 只是把「按 token 分文件」换成「单个 localStorage key」。
 * ============================================================================= */
(function () {
  'use strict';
  var T = window.TaxCore;

  var LS_DATA = 'stx.data.v1';
  var MAX_RECORDS = 24; // 两年 / 24 个月滚动存档

  function blank() {
    return {
      version: 1,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      records: [],
      settings: { carry: T.defaultCarry(), yearBase: {} }
    };
  }

  function load() {
    var raw = null;
    try { raw = localStorage.getItem(LS_DATA); } catch (e) { /* 隐私模式/配额满时忽略 */ }
    var d = null;
    if (raw) { try { d = JSON.parse(raw); } catch (e) { d = null; } }
    if (!d || typeof d !== 'object') return blank();
    d.records = Array.isArray(d.records) ? d.records : [];
    d.settings = d.settings || {};
    if (!d.settings.carry) d.settings.carry = T.defaultCarry();
    if (!d.settings.yearBase) d.settings.yearBase = {};
    return d;
  }

  function save(data) {
    data.updatedAt = new Date().toISOString();
    try {
      localStorage.setItem(LS_DATA, JSON.stringify(data));
      return true;
    } catch (e) {
      // 配额满（隐私模式下 Safari 约 5MB 限制）：抛给上层提示
      return false;
    }
  }

  /** 按发放月去重后写入；超出 24 个月自动覆盖最早的月份 */
  function upsert(record) {
    var d = load();
    var payMonth = T.fmtMonth(record.payMonth);
    if (!payMonth) return { error: '工资发放月无效' };
    var rec = Object.assign({}, record, { payMonth: payMonth, id: payMonth });
    rec.periodMonth = T.fmtMonth(rec.periodMonth) || T.nextMonth(payMonth, -1) || payMonth;
    T.FIELDS.forEach(function (f) {
      rec[f.key] = f.key === 'note' ? String(rec[f.key] || '') : T.num(rec[f.key]);
    });

    var i = d.records.findIndex(function (r) { return T.fmtMonth(r.payMonth) === payMonth; });
    var dropped = [];
    if (i >= 0) d.records[i] = rec;
    else d.records.push(rec);

    d.records.sort(function (a, b) { return T.monthKey(b.payMonth) - T.monthKey(a.payMonth); }); // 新→旧
    if (d.records.length > MAX_RECORDS) {
      dropped = d.records.splice(MAX_RECORDS).map(function (r) { return r.payMonth; });
    }
    var ok = save(d);
    return { data: d, dropped: dropped, saved: ok };
  }

  function remove(ids) {
    var d = load();
    var set = new Set((ids || []).map(function (x) { return String(x); }));
    var before = d.records.length;
    d.records = d.records.filter(function (r) {
      return !set.has(String(r.id)) && !set.has(T.fmtMonth(r.payMonth));
    });
    var ok = save(d);
    return { data: d, removed: before - d.records.length, saved: ok };
  }

  function reset() {
    var d = blank();
    save(d);
    return d;
  }

  function saveSettings(patch) {
    var d = load();
    if (patch.carry) d.settings.carry = Object.assign(T.defaultCarry(), patch.carry);
    if (patch.yearBase) d.settings.yearBase = patch.yearBase;
    var ok = save(d);
    return { settings: d.settings, saved: ok };
  }

  /** 导入场景：整体替换 records + settings，保留 createdAt */
  function replaceAll(records, settings) {
    var d = load();
    d.records = records || [];
    d.settings = settings || { carry: T.defaultCarry(), yearBase: {} };
    var ok = save(d);
    return { data: d, saved: ok };
  }

  window.StoreLocal = {
    MAX_RECORDS: MAX_RECORDS,
    load: load,
    save: save,
    upsert: upsert,
    remove: remove,
    reset: reset,
    saveSettings: saveSettings,
    replaceAll: replaceAll,
    blank: blank
  };
})();
