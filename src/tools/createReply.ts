import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { loadConfig } from '../config/loader.js';
import { outlookClient } from '../outlook/client.js';
import { cleanupPlainText } from '../utils/text.js';

export function registerCreateReplyTool(server: McpServer): void {
  server.registerTool(
    'create_reply',
    {
      title: '답장 초안 준비',
      description:
        '지정한 메일의 원문과 config/reply_style.md 에 정의된 답장 스타일 가이드를 함께 가져옵니다. ' +
        'AI 는 이 정보를 바탕으로 답장 본문을 스타일 가이드에 맞게 작성한 뒤, save_draft 도구를 호출해 ' +
        'Outlook 임시보관함에 저장해야 합니다. 이 도구 자체는 메일을 발송하거나 저장하지 않습니다.',
      inputSchema: {
        entryId: z.string().min(1).describe('답장할 원본 메일의 EntryID'),
        storeId: z.string().optional().describe('원본 메일이 속한 저장소 ID (생략 가능)'),
        replyAll: z.boolean().default(false).describe('true 이면 전체 답장(참조 포함) 대상으로 준비합니다.'),
      },
    },
    async ({ entryId, storeId, replyAll }) => {
      const [mail, config] = await Promise.all([outlookClient.getMail(entryId, storeId), loadConfig()]);

      const mode = replyAll ? 'replyAll' : 'reply';
      const suggestedSubject = /^\s*(RE|FW|답장|회신)\s*:/i.test(mail.subject) ? mail.subject : `RE: ${mail.subject}`;

      const lines = [
        '## 원본 메일',
        `제목: ${mail.subject}`,
        `발신자: ${mail.senderName} <${mail.senderEmail}>`,
        `수신일: ${mail.receivedTime}`,
        `참조: ${mail.cc.map((r) => `${r.name} <${r.email}>`).join(', ') || '(없음)'}`,
        '',
        '--- 원본 본문 ---',
        cleanupPlainText(mail.body),
        '',
        '## 답장 스타일 가이드 (config/reply_style.md)',
        config.replyStyle || '(reply_style.md 가 비어 있습니다. 정중한 존댓말로 간결하게 작성하세요.)',
        '',
        '## 다음 단계',
        `위 스타일 가이드에 맞게 답장 본문을 HTML로 작성한 뒤, save_draft 도구를 다음 값으로 호출해 임시보관함에 저장하세요:`,
        `  mode="${mode}", sourceEntryId="${entryId}"${storeId ? `, sourceStoreId="${storeId}"` : ''}, subject="${suggestedSubject}"`,
      ];

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
