// 画面テストの実行役。URL の #名前 のシナリオを1つ動かし、結果を <pre id="__result"> に書く（run.mjs が読む）。
(function () {
  'use strict';
  const name = decodeURIComponent(location.hash.slice(1));
  const out = (obj) => {
    const pre = document.createElement('pre');
    pre.id = '__result';
    pre.textContent = JSON.stringify(obj);
    document.body.appendChild(pre);
  };
  if (name === '__list') return out({ list: Object.keys(window.SCENARIOS) });
  const sc = window.SCENARIOS[name];
  if (!sc) return out({ name, pass: false, errors: ['シナリオがありません'] });

  const errors = [];
  const fmt = (v) => (typeof v === 'string' ? v : JSON.stringify(v));
  const t = {
    $: (s) => document.querySelector(s),
    $$: (s) => [...document.querySelectorAll(s)],
    calls: window.__calls,
    sleep: (ms) => new Promise((r) => setTimeout(r, ms)),
    async until(fn, msg, ms = 2000) {
      const end = Date.now() + ms;
      while (Date.now() < end) { try { if (fn()) return true; } catch (e) { /* 待つ */ } await t.sleep(20); }
      errors.push('待っても条件が成り立たない：' + msg);
      return false;
    },
    ok(cond, msg) { if (!cond) errors.push(msg); },
    eq(a, b, msg) { if (fmt(a) !== fmt(b)) errors.push(`${msg}\n    期待: ${fmt(b)}\n    実際: ${fmt(a)}`); },
    has(s, sub, msg) { if (!String(s).includes(sub)) errors.push(`${msg}\n    「${sub}」が含まれていない: ${fmt(s)}`); },
    // 実際に表示されているか（hidden 属性ではなく、CSS の結果で判定する）
    visible: (el) => !!el && el.offsetParent !== null && getComputedStyle(el).visibility !== 'hidden',
    click(el) { (typeof el === 'string' ? document.querySelector(el) : el).click(); },
    toast: () => document.getElementById('toast').textContent,
  };

  (async () => {
    await t.sleep(60); // Office.onReady とアドインの初期化を待つ
    try { await sc.run(t); } catch (e) { errors.push('例外：' + (e && e.stack ? e.stack : e)); }
    out({ name, title: sc.title, pass: errors.length === 0, errors });
  })();
})();
