// メール本文から日付の候補を抜き出す（社内記事用ブックマークレット v0.5.0 の読み取りルールをメール向けに移植）
// ブラウザでは window.MailDateParser、Node では module.exports として使う。外部には何も送らない。
(function (global) {
  'use strict';

  const clean = (s) => (s || '').replace(/\r/g, '').replace(/[ \t　]+/g, ' ').replace(/\s*\n\s*/g, '\n').trim();
  const z = (s) => String(s).replace(/[０-９：～]/g, (c) => ({ '：': ':', '～': '~' }[c] || String.fromCharCode(c.charCodeAt(0) - 0xfee0)));
  const pad = (n) => String(n).padStart(2, '0');
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const addHour = (t) => { const [h, m] = t.split(':').map(Number); return `${pad(Math.min(h + 1, 23))}:${pad(m)}`; };
  const nextDay = (s) => { const [y, m, d] = s.split('-').map(Number); return ymd(new Date(y, m - 1, d + 1)); };

  const WD = '(?:\\s*[（(][月火水木金土日祝・\\s]+[)）])?';
  const DATE = '(?:(\\d{4})\\s*[年/.]\\s*)?(\\d{1,2})\\s*[月/]\\s*(\\d{1,2})\\s*日?' + WD;
  const TIME = '(?:\\s*(\\d{1,2})\\s*[:時]\\s*(\\d{2})?\\s*分?|\\s*(午前|午後|正午|AM|PM|ＡＭ|ＰＭ)(?![a-z]))?';
  const RANGE = '(?:\\s*[~〜\\-－ー–]\\s*(?:' + DATE + '|(\\d{1,2})\\s*日?' + WD + '(?![\\d/月:時]))?' + TIME + ')?';

  // 返信・転送の引用部分の始まり（ここから下は古いやり取りなので読まない）
  const QUOTE_START = /^(?:-{2,}\s*(?:Original Message|元のメッセージ|Forwarded message|転送されたメッセージ)|_{5,}|(?:From|差出人|送信者)\s*[:：]|On .+wrote:|.*\d{4}年\d{1,2}月\d{1,2}日.*(?:書きました|wrote))/i;
  const DEADLINE_AFTER = /^(?:まで|迄|必着|締切|締め切り|〆切)/;
  const DEADLINE_WORD = /締切|締め切り|〆切|期限|必着|提出日/;

  const stripNum = (s) => s.replace(/^[\s　]*(?:[\d０-９]+\s*[.．、:：)）]|[一二三四五六七八九十]+\s*[.．、)）]|[①-⑳]|[(（][\d０-９]+[)）])\s*/, '')
    .replace(/^[■□●◆◇▼▶【\[・※\s]+|[】\]]+$/g, '').trim();
  const labelOf = (before) => {
    const b = before.trim();
    if (!b || /[、。がをにでもと（(「]$/.test(b) || /[、。]/.test(b)) return '';
    if (!(/[:：※]$/.test(b) || /は$/.test(b) || /^[■□●◆◇▼▶・※（(【\[]/.test(b))) return '';
    const l = stripNum(b).replace(/^[・※（(【\[]+|[)）】\]]+/g, '').replace(/[\s:：※]+$|は$/g, '').replace(/[)）】\]]$/, '').trim();
    return l.length >= 2 && l.length <= 20 ? l : '';
  };

  // 件名から RE:/FW: と【11/20締切】のような日付入りの括弧を外す
  const shortSubject = (s) => {
    const t = clean(s).replace(/^(?:\s*(?:re|fw|fwd|返信|転送)\s*[:：]\s*)+/i, '');
    return t.replace(/\s*[【\[(（][^】\])）]*(?:締切|〆切|期限|まで|\d+\s*[月/]\s*\d+)[^】\])）]*[】\])）]\s*/g, ' ').trim() || t;
  };

  // 本文の「自分宛ての部分」だけを返す（引用・署名の手前まで）
  function ownPart(text) {
    const out = [];
    for (const line of clean(text).split('\n')) {
      if (QUOTE_START.test(line.trim())) break;
      if (/^>/.test(line)) continue;
      out.push(line);
    }
    return out;
  }

  /**
   * @param {object} mail { subject, body(text), received(Date), today(Date) }
   * @returns {{ candidates: object[], shortTitle: string, location: string }}
   */
  function extract(mail) {
    const subject = clean(mail.subject || '');
    const shortTitle = shortSubject(subject) || '（件名なし）';
    const baseDate = mail.received ? new Date(mail.received) : new Date();
    const today = ymd(mail.today || new Date());
    const RE = new RegExp(DATE + TIME + RANGE, 'gi');

    const inferYear = (m, d) => {
      let y = baseDate.getFullYear();
      if (new Date(y, m - 1, d) - baseDate < -60 * 864e5) y += 1; // 受信日より2か月以上前なら翌年扱い
      return y;
    };
    const toTime = (h, mi, ap) => {
      if (h != null && h !== '') return +h > 23 ? '' : `${pad(+h)}:${pad(mi ? Math.min(+mi, 59) : 0)}`;
      if (ap) return /正午/.test(ap) ? '12:00' : /午前|AM/i.test(z(ap)) ? '09:00' : '13:00';
      return '';
    };
    function parseAll(text) {
      const out = [];
      const t = z(text);
      RE.lastIndex = 0;
      let m;
      while ((m = RE.exec(t))) {
        const [, y1, mo1, d1, h1, mi1, ap1, y2, mo2, d2, dOnly, h2, mi2, ap2] = m;
        const M1 = +mo1, D1 = +d1;
        if (M1 < 1 || M1 > 12 || D1 < 1 || D1 > 31) continue;
        const Y1 = y1 ? +y1 : inferYear(M1, D1);
        const start = new Date(Y1, M1 - 1, D1);
        let end = start;
        if (mo2 && d2) {
          const M2 = +mo2, D2 = +d2;
          end = new Date(y2 ? +y2 : (M2 < M1 ? Y1 + 1 : Y1), M2 - 1, D2);
          if (end < start) end = start;
        } else if (dOnly && +dOnly > D1 && +dOnly <= 31) {
          end = new Date(Y1, M1 - 1, +dOnly);
        }
        const st = toTime(h1, mi1, ap1);
        let et = toTime(h2, mi2, ap2);
        if (st && !et) et = addHour(st);
        out.push({ raw: m[0].trim(), index: m.index, startDate: ymd(start), endDate: ymd(end), startTime: st, endTime: et, allDay: !st });
      }
      return out;
    }

    const candidates = [];
    const dup = (p) => candidates.some((c) => c.startDate === p.startDate && c.endDate === p.endDate && c.startTime === p.startTime);
    const tooOld = (d) => { const [y, m, dd] = d.split('-').map(Number); return new Date(y, m - 1, dd) - baseDate < -180 * 864e5; };
    const add = (p, o) => {
      const warn = tooOld(p.startDate) ? `${p.startDate.slice(0, 4)}年になっています。受信日より半年以上前なので、年の誤記かもしれません` : '';
      const past = !warn && (p.endDate || p.startDate) < today;
      const inside = !past && p.startDate === p.endDate && !DEADLINE_WORD.test(o.subject) &&
        candidates.some((c) => c.startDate !== c.endDate && c.startDate <= p.startDate && p.startDate <= c.endDate);
      const note = past ? '過去の日付です' : inside ? '上の期間の中の日付です（補足の可能性）' : '';
      candidates.push({ ...p, ...o, warn, note, checked: !warn && !past && !inside });
    };

    // 1) 本文（引用の手前まで）
    let prevLine = '';
    for (const line of ownPart(mail.body)) {
      const zl = z(line);
      if (/例[:：]|（例|\(例/.test(zl)) { if (line.trim()) prevLine = line; continue; }
      for (const p of parseAll(line)) {
        const before = zl.slice(0, p.index);
        const after = zl.slice(p.index + p.raw.length).trim();
        if (/^追記/.test(after) || dup(p)) continue;
        const rest = zl.replace(p.raw, '').replace(/[（(]\s*[)）]/g, '').replace(/^[\s・※:：、。●■◆]+|[\s:：、。]+$/g, '');
        const isDeadline = DEADLINE_AFTER.test(after);
        const label = labelOf(before);
        let subj;
        if (label) {
          subj = isDeadline && !DEADLINE_WORD.test(label) ? `${label}（締切）` : label;
        } else if (/^(?:より|から)/.test(after)) {
          const what = stripNum(after.replace(/^(?:より|から)\s*/, '')).split(/[、。]/)[0]
            .replace(/(?:が|を)?(?:始まります|開始します|開始|実施します|受け付け)$|です$|します$/, '').replace(/[がをはの]$/, '').slice(0, 20);
          subj = what ? `${what} 開始` : '開始';
        } else if (isDeadline || DEADLINE_WORD.test(zl)) {
          subj = '締切';
        } else if (rest.replace(/[\s〜~～\-－()（）]/g, '').length < 2) {
          subj = stripNum(prevLine) || shortTitle;
        } else {
          const head = stripNum(rest).replace(/^[はがをにの]/, '').split(/[、。,]/)[0].replace(/(?:に|は|より|から|まで|の|で)$/, '');
          subj = (head.length >= 2 ? head : stripNum(rest)).slice(0, 30);
        }
        add(p, { subject: subj || shortTitle, context: (prevLine && subj === stripNum(prevLine) ? prevLine + '\n' : '') + line, deadline: isDeadline || DEADLINE_WORD.test(subj) });
      }
      if (line.trim()) prevLine = line;
    }

    // 2) 件名：【11/20締切】など。本文に同じ日付がなければ足す
    for (const p of parseAll(subject)) {
      const zt = z(subject);
      // 本文に同じ日が出ていれば（時刻つきでも）本文の方が詳しいので足さない
      if (/^\s*追記/.test(zt.slice(p.index + p.raw.length)) || candidates.some((c) => c.startDate === p.startDate)) continue;
      const br = (zt.slice(0, p.index).match(/[【(（[]([^【(（[\]】)）]*)$/) || [])[1] || '';
      let subj = br.replace(/[\s:：]+$/, '').trim();
      if (!subj || /^\d/.test(subj)) subj = DEADLINE_WORD.test(subject) ? '締切' : '期間';
      add(p, { subject: subj, context: `件名：${subject}`, deadline: DEADLINE_WORD.test(subj) });
    }

    // 場所：行頭が「場所」「会場」「開催場所」の行
    const lines = ownPart(mail.body);
    let location = '';
    for (let i = 0; i < lines.length && !location; i++) {
      const m = stripNum(lines[i]).match(/^(?:開催場所|場所|会場)\s*[:：]?\s*(.*)$/);
      if (m) location = (m[1] || lines[i + 1] || '').trim().slice(0, 40);
    }
    candidates.forEach((c) => (c.location = c.deadline ? "" : location)); // 締切に会場は付けない
    return { candidates, shortTitle, location };
  }

  const api = { extract, shortSubject, ownPart, nextDay, addHour, ymd };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.MailDateParser = api;
})(typeof window !== 'undefined' ? window : globalThis);
