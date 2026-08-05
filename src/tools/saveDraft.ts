import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { outlookClient } from '../outlook/client.js';

export function registerSaveDraftTool(server: McpServer): void {
  server.registerTool(
    'save_draft',
    {
      title: '초안 저장',
      description:
        '작성한 답장 또는 새 메일을 Outlook 임시보관함(Drafts)에 저장합니다. 메일을 발송하지는 않습니다. ' +
        'mode="reply"/"replyAll" 인 경우 sourceEntryId 로 지정한 원본 메일에 대한 답장 스레드로 저장되며 ' +
        '(수신자/제목/원본 인용문이 Outlook 에 의해 자동으로 채워집니다), mode="new" 인 경우 새 메일로 저장됩니다. ' +
        '저장된 초안은 사용자가 Outlook(Classic 또는 New)에서 검토 후 직접 발송해야 합니다.',
      inputSchema: {
        mode: z.enum(['reply', 'replyAll', 'new']).describe('저장할 초안의 종류'),
        sourceEntryId: z.string().optional().describe('mode 가 reply/replyAll 일 때 원본 메일의 EntryID (필수)'),
        sourceStoreId: z.string().optional().describe('원본 메일이 속한 저장소 ID (생략 가능)'),
        to: z
          .array(z.string())
          .optional()
          .describe('받는 사람 이메일 목록. reply/replyAll 에서 생략하면 원본 발신자(및 참조인)를 그대로 사용합니다. new 모드에서는 필수.'),
        cc: z.array(z.string()).optional().describe('참조 이메일 목록'),
        subject: z.string().optional().describe('제목. 생략하면 reply/replyAll 은 원본 제목에 RE: 를 붙여 자동 생성합니다.'),
        bodyHtml: z.string().min(1).describe('저장할 본문 (HTML). 줄바꿈은 <br> 또는 <p> 태그를 사용하세요.'),
      },
    },
    async ({ mode, sourceEntryId, sourceStoreId, to, cc, subject, bodyHtml }) => {
      if ((mode === 'reply' || mode === 'replyAll') && !sourceEntryId) {
        throw new Error('mode 가 reply 또는 replyAll 일 때는 sourceEntryId 가 필요합니다.');
      }
      if (mode === 'new' && (!to || to.length === 0)) {
        throw new Error('mode 가 new 일 때는 to 수신자가 최소 1명 필요합니다.');
      }

      const result = await outlookClient.saveDraft({ mode, sourceEntryId, sourceStoreId, to, cc, subject, bodyHtml });

      const text = [
        '임시보관함(Drafts)에 초안을 저장했습니다.',
        `제목: ${result.subject}`,
        `받는 사람: ${result.to}`,
        `entryId: ${result.draftEntryId}`,
        '',
        'Outlook 임시보관함에서 내용을 검토한 후 직접 발송하세요.',
      ].join('\n');

      return {
        content: [{ type: 'text', text }],
      };
    },
  );
}
