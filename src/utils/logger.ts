// MCP stdio 전송은 stdout 을 JSON-RPC 메시지 채널로 사용하므로,
// 로그는 반드시 stderr 로만 출력해야 합니다.

type Level = 'debug' | 'info' | 'warn' | 'error';

const LEVEL_WEIGHT: Record<Level, number> = { debug: 10, info: 20, warn: 30, error: 40 };

function currentLevelWeight(): number {
  const raw = (process.env.LOG_LEVEL ?? 'info').toLowerCase();
  const level = (raw in LEVEL_WEIGHT ? raw : 'info') as Level;
  return LEVEL_WEIGHT[level];
}

function write(level: Level, message: string): void {
  if (LEVEL_WEIGHT[level] < currentLevelWeight()) return;
  const timestamp = new Date().toISOString();
  process.stderr.write(`[${timestamp}] [${level.toUpperCase()}] ${message}\n`);
}

export const logger = {
  debug: (message: string): void => write('debug', message),
  info: (message: string): void => write('info', message),
  warn: (message: string): void => write('warn', message),
  error: (message: string): void => write('error', message),
};
