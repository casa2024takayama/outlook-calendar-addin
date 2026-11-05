import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const P = createRequire(import.meta.url)('../src/parser.js');
const today = new Date(2026, 10, 5);
const r = P.extract({
  subject: 'RE: 【11/20締切】年末調整書類の提出について', received: new Date(2026, 10, 4, 9), today,
  body: ['各位','■提出期限：11月20日(木)17:00','■説明会：11/12(水) 13:00〜14:00','■場所：3F 会議室A','不明点は 11/14 までにご連絡ください。','受付は12月1日から開始します。','実施期間','12/15-17','',
         '-----Original Message-----','From: 総務部','送信日時: 2026年10月1日 10:00','前回の締切は10/10でした'].join('\n'),
});
const s = r.candidates.map(c => `${c.subject}|${c.startDate}|${c.startTime}|${c.endDate}|${c.endTime}|${c.allDay}|${c.checked}`);
console.log(s.join('\n'), '\nshort:', r.shortTitle, '\nloc:', r.location);
assert.equal(r.shortTitle, '年末調整書類の提出について');
assert.equal(r.location, '3F 会議室A');
assert.ok(s.includes('提出期限|2026-11-20|17:00|2026-11-20|18:00|false|true'));
assert.ok(s.includes('説明会|2026-11-12|13:00|2026-11-12|14:00|false|true'));
assert.ok(s.some(x => x.startsWith('不明点（締切）|2026-11-14')));
assert.ok(s.some(x => x.startsWith('受付|2026-12-01')));
assert.ok(s.includes('実施期間|2026-12-15||2026-12-17||true|true'));
assert.ok(!s.some(x => x.includes('10-10') || x.includes('10-01')), '引用部分は読まない');
assert.equal(r.candidates.length, 5, '件名の11/20は本文と重複するので足さない');
// 年またぎ・過去日
const r2 = P.extract({ subject: '打ち合わせ', received: new Date(2026, 11, 20), today: new Date(2026, 11, 20), body: '1/8(木) 10時から定例です。\n12/1は終わりました' });
console.log(r2.candidates.map(c => `${c.subject}|${c.startDate}|${c.startTime}|${c.note}`));
assert.equal(r2.candidates[0].startDate, '2027-01-08');
assert.equal(r2.candidates[0].startTime, '10:00');
assert.equal(r2.candidates[1].note, '過去の日付です');
console.log('all ok');
