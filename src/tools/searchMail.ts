import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { loadConfig } from '../config/loader.js';
import { outlookClient } from '../outlook/client.js';

export function registerSearchMailTool(server: McpServer): void {
  server.registerTool(
    'search_mail',
    {
      title: '메일 검색',
      description:
        '제목/본문/발신자/날짜/고객사/키워드 조건으로 메일을 검색합니다. ' +
        '"회의록 메일 찾아줘", "지난주 수달 메일", "여우 관련 메일" 같은 요청에 사용하세요. ' +
        'customer 를 지정하면 config/customers.yml 에 등록된 별칭들로 자동 검색합니다.',
      inputSchema: {
        folderPath: z
          .string()
          .optional()
          .describe('검색할 받은편지함 하위 폴더 경로 ("/" 구분, 예: "고객사/수달"). 생략하면 받은편지함 전체(하위 폴더 포함)를 검색합니다.'),
        subject: z.string().optional().describe('제목에 포함되어야 하는 문자열'),
        body: z.string().optional().describe('본문에 포함되어야 하는 문자열'),
        sender: z.string().optional().describe('발신자 이름 또는 이메일에 포함되어야 하는 문자열'),
        customer: z.string().optional().describe('customers.yml 에 등록된 고객사 이름. 등록된 모든 별칭으로 검색합니다.'),
        keyword: z.string().optional().describe('제목/본문/발신자/참조 어디에든 포함되면 일치하는 자유 키워드'),
        dateFrom: z.string().optional().describe('검색 시작일 (ISO 8601, 예: 2026-07-28)'),
        dateTo: z.string().optional().describe('검색 종료일 (ISO 8601, 예: 2026-08-04)'),
        maxCount: z.number().int().positive().max(200).default(25).describe('최대 결과 수'),
      },
    },
    async ({ folderPath, subject, body, sender, customer, keyword, dateFrom, dateTo, maxCount }) => {
      const keywords: string[] = [];
      if (keyword) keywords.push(keyword);

      if (customer) {
        const { customers } = await loadConfig();
        const found = customers.customers.find((c) => c.name === customer);
        keywords.push(...(found ? found.aliases : [customer]));
      }

      const result = await outlookClient.searchMail({
        folderPath,
        subject,
        body,
        sender,
        keywords: keywords.length > 0 ? keywords : undefined,
        dateFrom,
        dateTo,
        maxCount,
      });

      const lines = [
        `검색 결과: ${result.items.length}건 (스캔 ${result.scanned}건${result.truncated ? ', 검색 범위 초과로 일부만 스캔됨' : ''})`,
        '',
        ...result.items.map(
          (m) => `- [${m.folderPath || '받은 편지함'}] "${m.subject}" - ${m.senderName} (${m.receivedTime})  entryId=${m.entryId}`,
        ),
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
