import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const scriptsDir = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(scriptsDir, '..');

const src = path.join(projectRoot, 'src', 'outlook', 'scripts');
const dest = path.join(projectRoot, 'dist', 'outlook', 'scripts');

if (!existsSync(src)) {
  throw new Error(`PowerShell 스크립트 디렉터리를 찾을 수 없습니다: ${src}`);
}

// fs.cpSync 는 일부 OneDrive 동기화 경로에서 네이티브 크래시를 일으키는 것이
// 관찰되어, 파일 단위로 직접 읽고 쓰는 방식을 사용합니다.
mkdirSync(dest, { recursive: true });

let copied = 0;
for (const entry of readdirSync(src, { withFileTypes: true })) {
  if (!entry.isFile()) continue;
  const content = readFileSync(path.join(src, entry.name));
  writeFileSync(path.join(dest, entry.name), content);
  copied += 1;
}

console.log(`PowerShell 스크립트 복사 완료: ${src} -> ${dest} (${copied}개 파일)`);
