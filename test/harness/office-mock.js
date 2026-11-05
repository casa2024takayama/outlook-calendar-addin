// Office.js の偽物。Outlook の外でアドインの画面を動かし、Outlook に渡した値と Slack への送信を記録する。
// 読み込み順：office-mock.js → scenarios.js（window.__MOCK と開いている予定を設定）→ アドイン本体 → runner.js
(function () {
  'use strict';
  const M = () => window.__MOCK || {};
  const calls = (window.__calls = { fetch: [], appointmentForms: [], browser: [], handlers: {}, roamingSaved: 0 });
  const ok = (value) => ({ status: 'succeeded', value });

  window.fetch = async (url, opts) => {
    calls.fetch.push({ url, opts });
    if (M().fetchFail) throw new Error('network down');
    return { type: 'opaque', ok: false, status: 0 };
  };

  const roaming = {};
  window.Office = {
    HostType: { Outlook: 'Outlook' },
    CoercionType: { Text: 'text', Html: 'html' },
    AsyncResultStatus: { Succeeded: 'succeeded', Failed: 'failed' },
    EventType: { ItemChanged: 'olkItemSelectedChanged' },
    onReady(cb) { setTimeout(() => cb({ host: M().host === undefined ? 'Outlook' : M().host }), 0); return Promise.resolve(); },
    context: {
      mailbox: {
        item: undefined,
        displayNewAppointmentForm(p) { calls.appointmentForms.push(p); },
        addHandlerAsync(type, h) { calls.handlers[type] = h; },
      },
      roamingSettings: {
        get: (k) => roaming[k],
        set: (k, v) => { roaming[k] = v; },
        saveAsync: (cb) => { calls.roamingSaved++; cb(ok()); },
      },
      ui: { openBrowserWindow(u) { calls.browser.push(u); } },
    },
  };

  // 予定・メールの組み立て（テストから使う）
  const prop = (v) => ({ getAsync: (a, b) => (typeof a === 'function' ? a : b)(ok(v)) });
  window.MockItems = {
    // 表示画面の予定（招待された予定など）
    readAppt: (o) => ({ itemType: 'appointment', itemId: o.id, subject: o.subject, start: o.start, end: o.end, location: o.location || '' }),
    // 編集画面の予定（自分が主催）。id なし＝未保存
    composeAppt: (o) => ({
      itemType: 'appointment', subject: prop(o.subject), start: prop(o.start), end: prop(o.end), location: prop(o.location || ''),
      isAllDayEvent: o.allDay === undefined ? undefined : prop(o.allDay),
      getItemIdAsync: (cb) => cb(o.id ? ok(o.id) : { status: 'failed', error: { message: 'not saved' } }),
    }),
    // 受信メール
    mail: (o) => ({
      itemType: 'message', subject: o.subject, itemId: o.itemId || 'mail-' + o.subject, internetMessageId: o.internetMessageId || '<' + o.subject + '@test>',
      dateTimeCreated: o.received || new Date(), from: { displayName: o.from || '総務部', emailAddress: 'soumu@example.com' },
      body: { getAsync: (type, cb) => cb(o.bodyFail ? { status: 'failed', error: { message: 'x' } } : ok(o.body)) },
    }),
  };
})();
