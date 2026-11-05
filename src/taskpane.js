// Outlook アドイン：開いているメールから日付を読み取り、確認してから「新しい予定」画面を入力済みで開く。
// 予定は直接書き込まない（保存は Outlook でユーザーが押す）。本文はこの画面の中だけで処理し、外部へは送らない。
(function () {
  'use strict';
  const VERSION = '0.1.0';
  const OUTLOOK_WEB = 'https://outlook.cloud.microsoft/calendar/deeplink/compose';
  const STORE_KEY = 'olCalAddin.v1';
  const P = window.MailDateParser;
  const $ = (id) => document.getElementById(id);
  const pad = (n) => String(n).padStart(2, '0');

  let state = { mail: null, candidates: [], shortTitle: '', mailKey: '', prefix: true };
  let inOutlook = false;

  // ---------- 小物 ----------
  let toastTimer;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => (t.hidden = true), 4000);
  }
  const toDate = (d, t) => { const [y, m, dd] = d.split('-').map(Number); const [h, mi] = (t || '00:00').split(':').map(Number); return new Date(y, m - 1, dd, h, mi); };
  const fmtAt = (iso) => { const d = new Date(iso); return `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
  const fmtMailDate = (d) => d ? `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}` : '';

  // ---------- 登録済みの記録（このPCの Outlook 内ストレージ・メールごと） ----------
  const loadStore = () => { try { return JSON.parse(localStorage.getItem(STORE_KEY)) || {}; } catch (e) { return {}; } };
  function saveDone(c, on) {
    if (!c.key || !state.mailKey) return;
    try {
      const s = loadStore();
      const a = (s[state.mailKey] = s[state.mailKey] || {});
      if (on) a[c.key] = { at: new Date().toISOString(), subject: c.subject, start: c.startDate };
      else delete a[c.key];
      if (!Object.keys(a).length) delete s[state.mailKey];
      localStorage.setItem(STORE_KEY, JSON.stringify(s));
    } catch (e) { /* 保存できない環境でも動作は続ける */ }
    c.done = on ? new Date().toISOString() : '';
  }

  // ---------- 予定の中身 ----------
  const fullSubject = (c) => (state.prefix && c.subject !== state.shortTitle ? `${c.subject}｜${state.shortTitle}` : c.subject);
  const bodyText = (c) => {
    const m = state.mail || {};
    return [
      c.context && `■ メールの記載\n${c.context}`,
      `■ メール\n件名：${m.subject || ''}` + (m.from ? `\n差出人：${m.from}` : '') + (m.received ? `\n受信：${fmtMailDate(m.received)}` : ''),
    ].filter(Boolean).join('\n\n');
  };
  // 終日は Outlook アプリの予定画面に直接指定できないため、0:00〜翌日0:00 で開く（画面で「終日」をオンにすると日付がそのまま合う）
  function times(c) {
    if (c.allDay) return { start: toDate(c.startDate), end: toDate(P.nextDay(c.endDate || c.startDate)) };
    return { start: toDate(c.startDate, c.startTime), end: toDate(c.endDate || c.startDate, c.endTime || P.addHour(c.startTime)) };
  }
  function webUrl(c) {
    const q = new URLSearchParams();
    q.set('subject', fullSubject(c));
    q.set('body', bodyText(c).replace(/\n/g, '<br>'));
    if (c.location) q.set('location', c.location);
    if (c.allDay) { q.set('allday', 'true'); q.set('startdt', c.startDate); q.set('enddt', P.nextDay(c.endDate || c.startDate)); }
    else { q.set('startdt', `${c.startDate}T${c.startTime}:00`); q.set('enddt', `${c.endDate || c.startDate}T${c.endTime || P.addHour(c.startTime)}:00`); }
    return `${OUTLOOK_WEB}?${q.toString()}`;
  }

  function openForm(c) {
    const { start, end } = times(c);
    const params = { subject: fullSubject(c), start, end, location: c.location || '', body: bodyText(c), requiredAttendees: [], optionalAttendees: [], resources: [] };
    if (!inOutlook) {
      toast(`（試し）予定画面を開きます\n${params.subject}\n${start.toLocaleString()} 〜 ${end.toLocaleString()}${c.allDay ? '\n→ 画面で「終日」をオン' : ''}`);
      return true;
    }
    try {
      Office.context.mailbox.displayNewAppointmentForm(params);
      if (c.allDay) toast('予定画面で「終日」をオンにしてから保存してください');
      return true;
    } catch (e) {
      toast('予定画面を開けませんでした：' + e.message);
      return false;
    }
  }
  function openWeb(c) {
    const url = webUrl(c);
    try {
      if (inOutlook && Office.context.ui.openBrowserWindow) { Office.context.ui.openBrowserWindow(url); return; }
    } catch (e) { /* 下の window.open へ */ }
    window.open(url, '_blank', 'noopener');
  }

  // ---------- 画面 ----------
  function render() {
    const list = $('list');
    const m = state.mail;
    if (!m) { $('mail').textContent = ''; list.innerHTML = '<p class="empty">メールを選ぶと日付を読み取ります</p>'; return; }
    const found = state.candidates.filter((c) => c.raw);
    $('mail').textContent = `${m.subject || '（件名なし）'}（日付候補 ${found.length} 件・登録済み ${found.filter((c) => c.done).length} 件）`;
    list.textContent = '';
    if (!state.candidates.length) {
      list.innerHTML = '<p class="empty">本文に日付が見つかりませんでした。<br>添付ファイルや画像に書かれているかもしれません。<br>「＋ 手動で追加」から入れられます。</p>';
      return;
    }
    // 未来のもの → 過去のもの の順
    const ordered = [...state.candidates].sort((a, b) => (a.note === '過去の日付です') - (b.note === '過去の日付です'));
    ordered.forEach((c) => list.appendChild(card(c)));
  }

  function card(c) {
    const el = document.createElement('div');
    el.className = 'c' + (c.done ? ' done' : c.checked ? '' : ' off');
    el.innerHTML = `
      <div class="st2"></div>
      <div class="r"><input type="text" class="sj" placeholder="件名"></div>
      <div class="r"><input type="date" class="sd"><input type="time" class="st" step="300"></div>
      <div class="r"><span>〜</span><input type="date" class="ed"><input type="time" class="et" step="300"><label><input type="checkbox" class="ad">終日</label></div>
      <div class="r"><input type="text" class="lc" placeholder="場所（任意）"></div>
      <div class="r"><button class="b pri go">予定を作成</button><button class="lnk web" hidden>Web版で終日として開く</button></div>
      <div class="hint adhint" hidden>終日の予定は、開いた画面で「終日」をオンにしてください</div>
      <div class="raw"></div>`;
    const q = (s) => el.querySelector(s);
    q('.sj').value = c.subject; q('.sd').value = c.startDate; q('.ed').value = c.endDate;
    q('.st').value = c.startTime; q('.et').value = c.endTime; q('.ad').checked = c.allDay; q('.lc').value = c.location || '';
    q('.raw').textContent = c.raw ? `読み取り元：${c.context.replace(/\n/g, ' / ')}` : '手動入力';
    if (c.warn) { const w = document.createElement('div'); w.className = 'warn'; w.textContent = `⚠ ${c.warn}`; q('.raw').before(w); }
    if (c.note && !c.done) { const n = document.createElement('div'); n.className = 'hint'; n.textContent = c.note; q('.raw').before(n); }

    const sync = () => { q('.st').disabled = q('.et').disabled = c.allDay; q('.web').hidden = q('.adhint').hidden = !c.allDay; };
    sync();
    q('.sj').oninput = (e) => (c.subject = e.target.value);
    q('.lc').oninput = (e) => (c.location = e.target.value);
    q('.sd').onchange = (e) => { c.startDate = e.target.value; if (!c.endDate || c.endDate < c.startDate) { c.endDate = c.startDate; q('.ed').value = c.endDate; } };
    q('.ed').onchange = (e) => (c.endDate = e.target.value);
    q('.st').onchange = (e) => { c.startTime = e.target.value; if (!c.endTime || c.endTime <= c.startTime) { c.endTime = P.addHour(c.startTime); q('.et').value = c.endTime; } };
    q('.et').onchange = (e) => (c.endTime = e.target.value);
    q('.ad').onchange = (e) => {
      c.allDay = e.target.checked;
      if (!c.allDay && !c.startTime) { c.startTime = '10:00'; c.endTime = '11:00'; q('.st').value = c.startTime; q('.et').value = c.endTime; }
      sync();
    };

    const st2 = q('.st2');
    if (c.key && c.done) {
      st2.innerHTML = '<span class="badge">✓ 登録済み</span><span class="at"></span><button class="lnk undo">未登録に戻す</button>';
      st2.querySelector('.at').textContent = `${fmtAt(c.done)} に作成`;
      st2.querySelector('.undo').onclick = () => { saveDone(c, false); render(); };
      q('.go').textContent = 'もう一度開く';
      q('.go').classList.remove('pri');
    } else if (c.key) {
      st2.innerHTML = '<span>未登録</span><button class="lnk mark">登録済みにする</button>';
      st2.querySelector('.mark').onclick = () => { saveDone(c, true); render(); };
    } else st2.remove();

    const after = () => { if (c.key && !c.done) { saveDone(c, true); render(); } };
    q('.go').onclick = () => {
      if (!c.startDate) return toast('日付を入れてください');
      if (!c.allDay && !c.startTime) return toast('時刻を入れるか「終日」にしてください');
      if (openForm(c)) after();
    };
    q('.web').onclick = () => { if (!c.startDate) return toast('日付を入れてください'); openWeb(c); after(); };
    return el;
  }

  // ---------- 読み取り ----------
  function analyze(mail, mailKey) {
    state.mail = mail;
    state.mailKey = mailKey || '';
    const r = P.extract(mail);
    state.shortTitle = r.shortTitle;
    const stored = (loadStore()[state.mailKey] || {});
    r.candidates.forEach((c) => {
      c.key = `${c.startDate}|${c.startTime}|${c.subject}`;
      c.done = (stored[c.key] || {}).at || '';
    });
    state.candidates = r.candidates;
    state.defaultLocation = r.location;
    render();
  }

  function loadCurrentItem() {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) { state.mail = null; render(); return; }
    $('list').innerHTML = '<p class="empty">読み取り中…</p>';
    item.body.getAsync(Office.CoercionType.Text, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) { $('list').innerHTML = '<p class="empty">本文を読み取れませんでした</p>'; return; }
      const from = item.from ? (item.from.displayName || item.from.emailAddress) : '';
      analyze({ subject: item.subject || '', body: res.value || '', received: item.dateTimeCreated, from }, item.internetMessageId || item.itemId);
    });
  }

  // ---------- 起動 ----------
  $('ver').textContent = 'v' + VERSION;
  $('pfx').onchange = (e) => (state.prefix = e.target.checked);
  $('add').onclick = () => {
    if (!state.mail) return toast('先にメールを選んでください');
    const t = P.ymd(new Date());
    state.candidates.push({ raw: '', startDate: t, endDate: t, startTime: '10:00', endTime: '11:00', allDay: false, subject: state.shortTitle, context: '', checked: true, location: state.defaultLocation || '' });
    render();
    window.scrollTo(0, document.body.scrollHeight);
  };

  function startDemo() {
    $('demo').hidden = false;
    $('demoSubject').value = '【11/20締切】年末調整書類の提出について';
    $('demoBody').value = [
      '各位', '', 'お疲れさまです。総務部です。', '年末調整の書類提出をお願いします。', '',
      '■提出期限：11月20日(木)17:00', '■説明会：11/12(水) 13:00〜14:00', '■場所：3F 会議室A', '',
      '不明点は 11/14 までにご連絡ください。', '', '-----Original Message-----', 'From: 総務部', '送信日時: 2026年10月1日 10:00',
    ].join('\n');
    $('demoRun').onclick = () => analyze({ subject: $('demoSubject').value, body: $('demoBody').value, received: new Date(), from: '（試し）' }, 'demo:' + $('demoSubject').value);
    $('demoRun').click();
  }

  let started = false;
  if (window.Office && Office.onReady) {
    // Office.js が初期化できない環境（Outlook の外）でも 3 秒で試し画面にする
    setTimeout(() => { if (!started) { started = true; startDemo(); } }, 3000);
    Office.onReady((info) => {
      if (started) return;
      started = true;
      inOutlook = !!(info && info.host === Office.HostType.Outlook && Office.context && Office.context.mailbox);
      if (!inOutlook) return startDemo();
      loadCurrentItem();
      // ピン留め中は別のメールを選ぶと読み直す
      try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadCurrentItem); } catch (e) { /* 古い Outlook では無視 */ }
    });
  } else startDemo();
})();
