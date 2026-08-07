import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { outlookClient } from '../outlook/client.js';

export function registerListDraftsTool(server: McpServer): void {
  server.registerTool(
    'list_drafts',
    {
      title: '임시보관함 조회',
      description:
        'Outlook 임시보관함(Drafts)에 저장된, 아직 발송하지 않은 초안 목록을 조회합니다. ' +
        '"내가 쓰던 임시보관함 메일", "저장해둔 초안 보여줘" 같은 요청에 사용하세요. ' +
        '최근에 수정한 초안이 먼저 나오며, 각 항목의 entryId 로 read_mail(본문 확인) 또는 ' +
        'save_draft(mode="update", 이어서 완성 후 같은 초안에 덮어쓰기)를 호출할 수 있습니다.',
      inputSchema: {
        subject: z.string().optional().describe('제목에 포함되어야 하는 문자열'),
        keyword: z.string().optional().describe('제목/본문/받는사람 어디에든 포함되면 일치하는 자유 키워드'),
        maxCount: z.number().int().positive().max(200).default(25).describe('최대 결과 수'),
      },
    },
    async ({ subject, keyword, maxCount }) => {
      const result = await outlookClient.listDrafts({ subject, keyword, maxCount });

      const lines = [
        `임시보관함 초안: ${result.items.length}건 (전체 ${result.totalCount}건 중)`,
        '',
        ...result.items.map(
          (d) =>
            `- "${d.subject || '(제목 없음)'}" → ${d.to || '(받는사람 없음)'} ` +
            `(수정: ${d.lastModifiedTime})  entryId=${d.entryId}\n  ${d.bodyPreview.replace(/\s+/g, ' ').slice(0, 120)}`,
        ),
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
