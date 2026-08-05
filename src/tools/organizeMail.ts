import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { classifyMail } from '../classification/classifier.js';
import { loadConfig } from '../config/loader.js';
import { outlookClient } from '../outlook/client.js';
import { logger } from '../utils/logger.js';

export function registerOrganizeMailTool(server: McpServer): void {
  server.registerTool(
    'organize_mail',
    {
      title: '메일 정리',
      description:
        '받은 편지함의 메일을 검사하여 제목/본문/참조 내용을 기준으로 고객사 또는 카테고리 폴더로 자동 이동합니다. ' +
        '"메일 정리해", "받은편지함 정리", "오늘 메일 분류" 같은 요청에 사용하세요. ' +
        '분류 규칙은 config/customers.yml, config/categories.yml 을 따릅니다.',
      inputSchema: {
        scope: z
          .enum(['all', 'unread', 'today'])
          .default('all')
          .describe('정리 대상 범위. all=받은편지함 전체, unread=읽지 않은 메일만, today=오늘 도착한 메일만'),
        dryRun: z
          .boolean()
          .default(false)
          .describe('true 이면 실제로 이동하지 않고 이동 계획만 미리 보여줍니다.'),
        maxCount: z.number().int().positive().max(5000).default(200).describe('검사할 최대 메일 수'),
      },
    },
    async ({ scope, dryRun, maxCount }) => {
      const { customers, categories } = await loadConfig();
      const mails = await outlookClient.listInboxMails(scope, maxCount);

      const classified = mails.map((mail) => ({
        mail,
        classification: classifyMail(
          {
            subject: mail.subject,
            bodyPreview: mail.bodyPreview,
            senderName: mail.senderName,
            senderEmail: mail.senderEmail,
            ccNames: mail.ccNames,
            ccEmails: mail.ccEmails,
            itemType: mail.itemType,
          },
          customers,
          categories,
        ),
      }));

      const moved: Array<{ subject: string; from: string; folder: string; reason: string }> = [];
      const errors: Array<{ subject: string; error: string }> = [];

      if (dryRun) {
        for (const { mail, classification } of classified) {
          moved.push({
            subject: mail.subject,
            from: mail.senderName,
            folder: classification.folderPath,
            reason: classification.reason,
          });
        }
      } else {
        // 건마다 PowerShell 프로세스를 새로 띄우면 대량 이동 시 매우 느려지므로,
        // 한 번의 Outlook 연결로 일괄 이동한다.
        const result = await outlookClient.moveMailBatch(
          classified.map(({ mail, classification }) => ({
            entryId: mail.entryId,
            storeId: mail.storeId,
            targetPath: classification.folderPath,
          })),
        );

        const bySubject = new Map(classified.map(({ mail, classification }) => [mail.entryId, { mail, classification }]));
        for (const m of result.moved) {
          const info = bySubject.get(m.entryId);
          moved.push({
            subject: info?.mail.subject ?? m.entryId,
            from: info?.mail.senderName ?? '',
            folder: m.targetPath,
            reason: info?.classification.reason ?? '',
          });
        }
        for (const e of result.errors) {
          const info = bySubject.get(e.entryId);
          const subject = info?.mail.subject ?? e.entryId;
          logger.error(`메일 이동 실패 ("${subject}"): ${e.error}`);
          errors.push({ subject, error: e.error });
        }
      }

      const summaryLine = `${dryRun ? '[미리보기] ' : ''}검사한 메일: ${mails.length}건, 이동 대상: ${moved.length}건, 오류: ${errors.length}건`;
      const lines = [summaryLine, '', ...moved.map((m) => `- "${m.subject}" (${m.from}) → ${m.folder}  [${m.reason}]`)];
      if (errors.length > 0) {
        lines.push('', '오류:', ...errors.map((e) => `- "${e.subject}": ${e.error}`));
      }

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
