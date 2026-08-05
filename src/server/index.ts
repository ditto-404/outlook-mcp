#!/usr/bin/env node
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig } from '../config/loader.js';
import { registerAllTools } from '../tools/index.js';
import { logger } from '../utils/logger.js';

const server = new McpServer({ name: 'outlook-mcp', version: '0.1.0' });

registerAllTools(server);

async function main(): Promise<void> {
  try {
    await loadConfig();
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    logger.error(`설정 로드 실패: ${message}`);
    logger.error('config/customers.yml, config/categories.yml 파일을 확인하거나 OUTLOOK_MCP_CONFIG_DIR 환경변수를 설정하세요.');
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  logger.info('outlook-mcp 서버가 시작되었습니다 (stdio).');
}

main().catch((err: unknown) => {
  const message = err instanceof Error ? (err.stack ?? err.message) : String(err);
  logger.error(`서버 시작 실패: ${message}`);
  process.exit(1);
});
