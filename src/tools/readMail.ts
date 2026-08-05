import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { outlookClient } from '../outlook/client.js';
import { cleanupPlainText } from '../utils/text.js';

export function registerReadMailTool(server: McpServer): void {
  server.registerTool(
    'read_mail',
    {
      title: '메일 읽기',
      description:
        'entryId(및 storeId)로 지정한 메일 한 건의 본문/제목/발신자/수신자/참조/날짜/첨부파일 목록을 읽어옵니다. ' +
        'search_mail 이나 organize_mail 결과에 포함된 entryId 를 사용하세요. HTML 메일은 읽기 쉬운 텍스트로 변환되어 반환됩니다.',
      inputSchema: {
        entryId: z.string().min(1).describe('Outlook 메일 항목의 EntryID'),
        storeId: z.string().optional().describe('메일이 속한 Outlook 저장소(Store)의 ID (생략 가능)'),
      },
    },
    async ({ entryId, storeId }) => {
      const mail = await outlookClient.getMail(entryId, storeId);
      const body = cleanupPlainText(mail.body);

      const lines = [
        `제목: ${mail.subject}`,
        `발신자: ${mail.senderName} <${mail.senderEmail}>`,
        `수신자: ${mail.to.map((r) => `${r.name} <${r.email}>`).join(', ') || '(없음)'}`,
        `참조: ${mail.cc.map((r) => `${r.name} <${r.email}>`).join(', ') || '(없음)'}`,
        `수신일: ${mail.receivedTime}`,
        `폴더: ${mail.folderPath || '받은 편지함'}`,
        `첨부파일: ${
          mail.attachments.length > 0 ? mail.attachments.map((a) => `${a.fileName} (${a.size} bytes)`).join(', ') : '없음'
        }`,
        `entryId: ${mail.entryId}`,
        `storeId: ${mail.storeId}`,
        '',
        '--- 본문 ---',
        body,
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
