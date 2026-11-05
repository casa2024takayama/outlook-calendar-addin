// 自動テスト一式：単体テスト（node）＋画面テスト（ヘッドレス Chrome）＋ manifest 検証（--manifest のとき）
// 使い方: node test/run.mjs [--manifest] [シナリオ名…]
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { readFileSync, writeFileSync, mkdirSync, rmSync, cpSync, readdirSync, existsSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';

const run = promisify(execFile);
const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const OUT = join(ROOT, 'test', '.out');
const CHROME = process.env.CHROME || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const args = process.argv.slice(2);
const only = args.filter((a) => !a.startsWith('--'));
const env = { ...process.env, TZ: 'Asia/Tokyo' }; // 予定の時刻は日本時間（+09:00）で確かめる
let failed = 0;

// ---------- 1) 単体テスト ----------
const unit = readdirSync(join(ROOT, 'test')).filter((f) => f.endsWith('.test.mjs'));
for (const f of unit) {
  try {
    await run(process.execPath, [join(ROOT, 'test', f)], { env });
    console.log(`✓ 単体 ${f}`);
  } catch (e) {
    failed++;
    console.log(`✗ 単体 ${f}\n${(e.stdout || '') + (e.stderr || '')}`);
  }
}

// ---------- 2) 画面テスト ----------
// src/taskpane.html から Office.js（本物）を外し、偽物とシナリオを差し込んだページを作る
rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });
cpSync(join(ROOT, 'src'), OUT, { recursive: true });
cpSync(join(ROOT, 'test', 'harness'), join(OUT, 'harness'), { recursive: true });
let html = readFileSync(join(ROOT, 'src', 'taskpane.html'), 'utf8');
const officeTag = /<script src="https:\/\/appsforoffice\.microsoft\.com[^"]*"><\/script>\s*/;
if (!officeTag.test(html)) throw new Error('taskpane.html に Office.js の読み込みが見つからない');
html = html.replace(officeTag, '');
const firstLocal = html.search(/<script src="(?!https?:)/);
html = html.slice(0, firstLocal) + '<script src="harness/office-mock.js"></script>\n  <script src="harness/scenarios.js"></script>\n  ' + html.slice(firstLocal);
html = html.replace('</body>', '  <script src="harness/runner.js"></script>\n</body>');
writeFileSync(join(OUT, 'harness.html'), html);

async function page(hash) {
  const prof = join(tmpdir(), `addin-test-${process.pid}-${Math.random().toString(36).slice(2)}`);
  try {
    const { stdout } = await run(CHROME, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', `--user-data-dir=${prof}`,
      '--allow-file-access-from-files', '--virtual-time-budget=15000', '--dump-dom', `file://${join(OUT, 'harness.html')}#${encodeURIComponent(hash)}`], { env, timeout: 60000, maxBuffer: 1 << 24 });
    const m = stdout.match(/<pre id="__result">([\s\S]*?)<\/pre>/);
    if (!m) return { name: hash, pass: false, errors: ['結果が出力されなかった（スクリプトのエラー？）'] };
    return JSON.parse(m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&'));
  } finally {
    rmSync(prof, { recursive: true, force: true });
  }
}

if (!existsSync(CHROME)) {
  failed++;
  console.log(`✗ 画面テスト：Chrome が見つからない（${CHROME}）`);
} else {
  const { list } = await page('__list');
  const names = only.length ? list.filter((n) => only.includes(n)) : list;
  const results = [];
  for (let i = 0; i < names.length; i += 4) results.push(...(await Promise.all(names.slice(i, i + 4).map(page))));
  for (const r of results) {
    if (r.pass) console.log(`✓ 画面 ${r.name}：${r.title}`);
    else { failed++; console.log(`✗ 画面 ${r.name}：${r.title || ''}\n  - ${r.errors.join('\n  - ')}`); }
  }
}

// ---------- 3) manifest 検証（Microsoft の検証サービスに送る。manifest は公開済みの内容） ----------
if (args.includes('--manifest')) {
  try {
    const { stdout } = await run('npx', ['--yes', 'office-addin-manifest', 'validate', join(ROOT, 'docs', 'manifest.xml')], { env, timeout: 180000 });
    if (/The manifest is valid/.test(stdout)) console.log('✓ manifest 検証');
    else { failed++; console.log('✗ manifest 検証\n' + stdout.split('\n').filter((l) => /Error|not valid/.test(l)).join('\n')); }
  } catch (e) {
    failed++;
    console.log('✗ manifest 検証\n' + ((e.stdout || '') + (e.stderr || '')).split('\n').filter((l) => /Error|not valid|valid/.test(l)).join('\n'));
  }
}

console.log(failed ? `\n${failed} 件失敗` : '\nすべて成功');
process.exit(failed ? 1 : 0);
