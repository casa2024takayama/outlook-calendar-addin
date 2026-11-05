// src/ と assets/ を docs/ にまとめ、配置先 URL を入れた manifest.xml を作る
// 使い方: node build.mjs https://<配置先>   （末尾の / は不要）
import { readFileSync, writeFileSync, mkdirSync, cpSync, rmSync } from 'node:fs';

const base = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\//.test(base)) {
  console.error('配置先の https URL を指定してください。例: node build.mjs https://localhost:3000');
  process.exit(1);
}
const VERSION = '0.1.0.0';

rmSync('docs', { recursive: true, force: true });
mkdirSync('docs/assets', { recursive: true });
cpSync('src', 'docs', { recursive: true });
cpSync('assets', 'docs/assets', { recursive: true });
const manifest = readFileSync('manifest.template.xml', 'utf8').replaceAll('{{BASE}}', base).replaceAll('{{VERSION}}', VERSION);
writeFileSync('docs/manifest.xml', manifest);
console.log(`docs/ を作成しました（配置先 ${base}）。Outlook には docs/manifest.xml を「カスタムアドインをファイルから追加」で入れてください。`);
