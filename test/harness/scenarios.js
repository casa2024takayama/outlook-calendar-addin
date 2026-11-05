// 「メールから予定」の画面テスト。1シナリオ＝1回のページ読み込み。
// メールの日付は今日からの相対で作る（過去・未来の判定を毎日同じ結果にするため）。
(function () {
  'use strict';
  const pad = (n) => String(n).padStart(2, '0');
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const now = new Date();
  const day = (off, h = 0, m = 0) => new Date(now.getFullYear(), now.getMonth(), now.getDate() + off, h, m);
  const md = (d) => `${d.getMonth() + 1}/${d.getDate()}`;
  const mdJ = (d) => `${d.getMonth() + 1}月${d.getDate()}日`;
  const D7 = day(7), D10 = day(10), D14 = day(14), P3 = day(-3);
  const I = window.MockItems;
  const open = (item) => { Office.context.mailbox.item = item; };
  const SUBJ = `【${md(D14)}締切】年末調整書類の提出について`;
  const BODY = [
    '各位', '', '総務部です。年末調整の書類提出をお願いします。', '',
    `■提出期限：${mdJ(D14)}(木)17:00`,
    `■説明会：${md(D7)}(水) 13:00〜14:00`,
    '■場所：3F 会議室A', '',
    `不明点は ${md(D10)} までにご連絡ください。`, '',
    '-----Original Message-----', 'From: 総務部', `送信日時: ${md(day(-40))} 10:00`, `前回の締切は${md(day(-30))}でした`,
  ].join('\n');
  const MAIL = { subject: SUBJ, body: BODY, internetMessageId: '<nenmatsu@test>' };
  const cards = (t) => t.$$('#list .c');
  const card = (t, subj) => cards(t).find((c) => c.querySelector('.sj').value === subj);
  const set = (el, v) => { el.value = v; el.dispatchEvent(new Event('input')); el.dispatchEvent(new Event('change')); };

  window.SCENARIOS = {
    'read-candidates': {
      title: '本文の日付を候補にする（引用部分は読まない・件名の日付は重複させない）',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => cards(t).length > 0, '候補の表示');
        const s = cards(t).map((c) => [c.querySelector('.sj').value, c.querySelector('.sd').value, c.querySelector('.st').value, c.querySelector('.ad').checked]);
        t.eq(s, [
          ['提出期限', ymd(D14), '17:00', false],
          ['説明会', ymd(D7), '13:00', false],
          ['不明点（締切）', ymd(D10), '', true],
        ], '候補（件名・日付・時刻・終日）');
        t.has(t.$('#mail').textContent, '日付候補 3 件', '件数');
        t.ok(!t.visible(t.$('#demo')), 'Outlook の中なのに試し入力欄が出ている');
        t.eq(card(t, '提出期限').querySelector('.lc').value, '', '締切に会場が入っている');
        t.eq(card(t, '説明会').querySelector('.lc').value, '3F 会議室A', '説明会の会場');
      },
    },
    'create-timed': {
      title: '「予定を作成」で入力済みの予定画面を開き、登録済みにする',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => card(t, '説明会'), '候補');
        t.click(card(t, '説明会').querySelector('.go'));
        await t.until(() => t.calls.appointmentForms.length === 1, '予定画面を開く');
        const p = t.calls.appointmentForms[0];
        t.eq(p.subject, '説明会｜年末調整書類の提出について', '件名（RE/日付括弧を外したメール件名付き）');
        t.eq([p.start.getTime(), p.end.getTime()], [day(7, 13).getTime(), day(7, 14).getTime()], '開始・終了');
        t.eq(p.location, '3F 会議室A', '場所');
        t.has(p.body, '■ メールの記載', '本文に読み取り元');
        t.has(p.body, '差出人：総務部', '本文に差出人');
        await t.until(() => card(t, '説明会').querySelector('.badge'), '登録済み表示');
        t.has(t.$('#mail').textContent, '登録済み 1 件', '件数の更新');
        t.ok(JSON.parse(localStorage.getItem('olCalAddin.v1'))['<nenmatsu@test>'], '記録がメールの ID で残っていない');
      },
    },
    'create-allday': {
      title: '終日は 0:00〜翌0:00 で開き、「終日をオン」と案内する',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => card(t, '不明点（締切）'), '候補');
        const c = card(t, '不明点（締切）');
        t.ok(t.visible(c.querySelector('.adhint')), '終日の案内が出ていない');
        t.click(c.querySelector('.go'));
        await t.until(() => t.calls.appointmentForms.length === 1, '予定画面を開く');
        const p = t.calls.appointmentForms[0];
        t.eq([p.start.getTime(), p.end.getTime()], [day(10).getTime(), day(11).getTime()], '0:00〜翌0:00');
        t.has(t.toast(), '終日', '保存前の案内');
      },
    },
    'web-allday': {
      title: '「Web版で終日として開く」は終日の deeplink（終わりは翌日）',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => card(t, '不明点（締切）'), '候補');
        t.click(card(t, '不明点（締切）').querySelector('.web'));
        await t.until(() => t.calls.browser.length === 1, 'ブラウザで開く');
        const u = new URL(t.calls.browser[0]);
        t.eq(u.origin + u.pathname, 'https://outlook.cloud.microsoft/calendar/deeplink/compose', 'deeplink の宛先');
        t.eq([u.searchParams.get('allday'), u.searchParams.get('startdt'), u.searchParams.get('enddt')], ['true', ymd(D10), ymd(day(11))], '終日の日付');
      },
    },
    'edit-then-create': {
      title: '件名・時刻を直してから作成すると、直した値で開く',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => card(t, '説明会'), '候補');
        const c = card(t, '説明会');
        set(c.querySelector('.sj'), '年末調整 説明会');
        set(c.querySelector('.st'), '15:30');
        t.eq(c.querySelector('.et').value, '16:30', '開始を終了より後にしたら終了を1時間後に');
        t.click('#pfx');
        t.click(c.querySelector('.go'));
        await t.until(() => t.calls.appointmentForms.length === 1, '予定画面');
        const p = t.calls.appointmentForms[0];
        t.eq(p.subject, '年末調整 説明会', 'メール件名を付けない設定');
        t.eq([p.start.getHours(), p.start.getMinutes(), p.end.getHours(), p.end.getMinutes()], [15, 30, 16, 30], '時刻');
      },
    },
    'past-date': {
      title: '過去の日付は注記つきで下に並べ、最初は薄く表示',
      setup() { open(I.mail({ subject: 'お知らせ', body: `${md(P3)} に実施しました。\n次回は ${md(D7)} 10:00 からです。`, internetMessageId: '<past@test>' })); },
      async run(t) {
        await t.until(() => cards(t).length === 2, '候補2件');
        const last = cards(t)[1];
        t.eq(last.querySelector('.sd').value, ymd(P3), '過去が後ろ');
        t.has(last.textContent, '過去の日付です', '注記');
        t.ok(last.classList.contains('off'), '薄く表示されていない');
      },
    },
    'subject-only-date': {
      title: '本文に日付がなく件名だけにある締切も拾う',
      setup() { open(I.mail({ subject: `RE: 【締切：${mdJ(D14)}】アンケート回答のお願い`, body: 'ご協力をお願いします。', internetMessageId: '<subj@test>' })); },
      async run(t) {
        await t.until(() => cards(t).length === 1, '候補1件');
        t.eq([cards(t)[0].querySelector('.sj').value, cards(t)[0].querySelector('.sd').value], ['締切', ymd(D14)], '件名から');
      },
    },
    'no-dates': {
      title: '日付がないメールは案内し、手動で追加できる',
      setup() { open(I.mail({ subject: '資料送付', body: '添付をご確認ください。', internetMessageId: '<none@test>' })); },
      async run(t) {
        await t.until(() => t.$('#list').textContent.includes('見つかりませんでした'), '案内');
        t.click('#add');
        await t.until(() => cards(t).length === 1, '手動の行');
        t.eq(cards(t)[0].querySelector('.sj').value, '資料送付', '件名の初期値');
        t.click(cards(t)[0].querySelector('.go'));
        await t.until(() => t.calls.appointmentForms.length === 1, '予定画面');
        t.ok(!cards(t)[0].querySelector('.badge'), '手動の行は記録しない');
      },
    },
    'quoted-only': {
      title: '引用部分にしか日付がないメールは候補なし',
      setup() { open(I.mail({ subject: 'Re: 打合せ', body: `承知しました。\n\nOn ${md(D7)} 総務部 wrote:\n> ${md(D7)} 10:00 からです`, internetMessageId: '<q@test>' })); },
      async run(t) { await t.until(() => t.$('#list').textContent.includes('見つかりませんでした'), '候補なしの案内'); },
    },
    'registered-persist': {
      title: '登録済みは次に開いても残り、「未登録に戻す」で戻せる',
      setup() {
        localStorage.setItem('olCalAddin.v1', JSON.stringify({ '<nenmatsu@test>': { [`${ymd(D7)}|13:00|説明会`]: { at: new Date().toISOString(), subject: '説明会', start: ymd(D7) } } }));
        open(I.mail(MAIL));
      },
      async run(t) {
        await t.until(() => card(t, '説明会'), '候補');
        const c = card(t, '説明会');
        t.ok(c.querySelector('.badge'), '登録済みが出ない');
        t.eq(c.querySelector('.go').textContent, 'もう一度開く', 'ボタンの文言');
        t.click(c.querySelector('.undo'));
        await t.until(() => !card(t, '説明会').querySelector('.badge'), '未登録に戻る');
        t.eq(JSON.parse(localStorage.getItem('olCalAddin.v1') || '{}')['<nenmatsu@test>'], undefined, '記録が消えていない');
      },
    },
    'item-changed': {
      title: 'ピン留め中に別のメールを選ぶと読み直す',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => cards(t).length === 3, '1通目');
        Office.context.mailbox.item = I.mail({ subject: '面談の件', body: `${md(D10)} 15:00〜15:30 で面談です。`, internetMessageId: '<m2@test>' });
        t.calls.handlers['olkItemSelectedChanged']();
        await t.until(() => t.$('#mail').textContent.startsWith('面談の件'), '2通目');
        t.eq(cards(t).length, 1, '2通目の候補数');
      },
    },
    'body-fail': {
      title: '本文を読めないときは案内する',
      setup() { open(I.mail({ subject: 'x', body: '', bodyFail: true })); },
      async run(t) { await t.until(() => t.$('#list').textContent.includes('読み取れませんでした'), '案内'); },
    },
    'need-date': {
      title: '日付を消して作成しようとすると止める',
      setup() { open(I.mail(MAIL)); },
      async run(t) {
        await t.until(() => card(t, '説明会'), '候補');
        const c = card(t, '説明会');
        set(c.querySelector('.sd'), '');
        t.click(c.querySelector('.go'));
        await t.sleep(30);
        t.eq(t.calls.appointmentForms.length, 0, '日付なしで開いた');
        t.has(t.toast(), '日付', '案内');
      },
    },
    'no-item': {
      title: 'メールを選んでいないときは案内だけ',
      setup() { open(undefined); },
      async run(t) { await t.until(() => t.$('#list').textContent.includes('メールを選ぶと'), '案内'); },
    },
    'outside-outlook': {
      title: 'Outlook の外では試し画面になり、サンプルを読み取る',
      setup() { window.__MOCK = { host: null }; },
      async run(t) {
        await t.until(() => t.visible(t.$('#demo')) && cards(t).length > 0, '試し画面と候補');
        t.click(cards(t)[0].querySelector('.go'));
        await t.sleep(30);
        t.eq(t.calls.appointmentForms.length, 0, '試し画面なのに Outlook を呼んだ');
        t.has(t.toast(), '（試し）', '試しの案内');
      },
    },
  };

  const name = decodeURIComponent(location.hash.slice(1));
  try { localStorage.clear(); } catch (e) { /* なし */ }
  window.__MOCK = {};
  if (window.SCENARIOS[name]) window.SCENARIOS[name].setup();
})();
