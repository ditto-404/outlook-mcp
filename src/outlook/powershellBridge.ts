import { spawn } from 'node:child_process';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { logger } from '../utils/logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SCRIPTS_DIR = path.join(__dirname, 'scripts');

const POWERSHELL_EXE = process.env.OUTLOOK_MCP_POWERSHELL_EXE ?? 'powershell.exe';
const DEFAULT_TIMEOUT_MS = 60_000;

export class OutlookScriptError extends Error {
  constructor(
    message: string,
    readonly scriptName: string,
  ) {
    super(message);
    this.name = 'OutlookScriptError';
  }
}

interface ScriptResponse<T> {
  ok: boolean;
  data?: T;
  error?: string;
}

/**
 * Classic Outlook COM 은 Node.js 프로세스에서 직접 호출할 수 없으므로,
 * PowerShell 프로세스를 자식으로 실행해 요청/응답을 임시 JSON 파일로 주고받습니다.
 * (PowerShell 5.1 의 stdout 인코딩은 로케일에 따라 신뢰하기 어려워 파일 교환 방식을 사용합니다.)
 */
export async function runOutlookScript<T = unknown>(
  scriptName: string,
  params: Record<string, unknown> = {},
  timeoutMs = DEFAULT_TIMEOUT_MS,
): Promise<T> {
  const scriptPath = path.join(SCRIPTS_DIR, scriptName);
  const workDir = await mkdtemp(path.join(tmpdir(), 'outlook-mcp-'));
  const requestPath = path.join(workDir, 'request.json');
  const responsePath = path.join(workDir, 'response.json');

  try {
    await writeFile(requestPath, JSON.stringify(params), 'utf8');

    await new Promise<void>((resolve, reject) => {
      const child = spawn(
        POWERSHELL_EXE,
        [
          '-NoProfile',
          '-NonInteractive',
          '-NoLogo',
          '-ExecutionPolicy',
          'Bypass',
          '-File',
          scriptPath,
          '-RequestPath',
          requestPath,
          '-ResponsePath',
          responsePath,
        ],
        { windowsHide: true },
      );

      let stderr = '';
      const timer = setTimeout(() => {
        child.kill();
        reject(new OutlookScriptError(`${scriptName} 실행이 ${timeoutMs}ms 내에 끝나지 않았습니다.`, scriptName));
      }, timeoutMs);

      child.stderr.on('data', (chunk: Buffer) => {
        stderr += chunk.toString('utf8');
      });
      child.on('error', (err) => {
        clearTimeout(timer);
        reject(new OutlookScriptError(`PowerShell(${POWERSHELL_EXE}) 실행 실패: ${err.message}`, scriptName));
      });
      child.on('close', (code) => {
        clearTimeout(timer);
        if (code !== 0) {
          logger.warn(`${scriptName} 종료 코드 ${code}${stderr.trim() ? `: ${stderr.trim()}` : ''}`);
        }
        resolve();
      });
    });

    const raw = await readFile(responsePath, 'utf8').catch(() => '');
    const withoutBom = raw.charCodeAt(0) === 0xfeff ? raw.slice(1) : raw;
    const cleaned = withoutBom.trim();
    if (!cleaned) {
      throw new OutlookScriptError(`${scriptName} 이(가) 결과를 반환하지 않았습니다. Outlook 이 실행 중인지 확인하세요.`, scriptName);
    }

    let parsed: ScriptResponse<T>;
    try {
      parsed = JSON.parse(cleaned) as ScriptResponse<T>;
    } catch {
      throw new OutlookScriptError(`${scriptName} 의 응답을 JSON 으로 해석할 수 없습니다: ${cleaned.slice(0, 500)}`, scriptName);
    }

    if (!parsed.ok) {
      throw new OutlookScriptError(parsed.error ?? '알 수 없는 Outlook 오류', scriptName);
    }
    return parsed.data as T;
  } finally {
    await rm(workDir, { recursive: true, force: true }).catch(() => {});
  }
}
