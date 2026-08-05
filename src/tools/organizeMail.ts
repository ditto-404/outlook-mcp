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
        '받은 편지함(또는 보낸 편지함)의 메일을 검사하여 제목/본문/참조 내용을 기준으로 고객사 또는 카테고리 폴더로 ' +
        '자동 이동합니다. "메일 정리해", "받은편지함 정리", "오늘 메일 분류" 같은 요청에 사용하세요. ' +
        '분류 규칙은 config/customers.yml, config/categories.yml 을 따릅니다.\n' +
        '규칙: (1) 읽지 않은 메일은 절대 이동하지 않습니다 - 사용자가 아직 확인하지 않은 메일이기 때문입니다. ' +
        '(2) 고객사/카테고리 어디에도 확실히 일치하지 않는 애매한 메일은 임의로 미분류 폴더로 옮기지 않고, ' +
        '"확인이 필요한 메일" 목록으로만 보고합니다 - 이 목록은 반드시 사용자에게 보여주고 어떻게 분류할지 물어보세요. ' +
        '(3) rootFolder 로 보낸편지함을 검사하더라도, 고객사/카테고리 폴더는 항상 받은편지함 쪽 폴더 트리 하나로 ' +
        '모입니다 - 같은 고객사에게 받은 메일과 보낸 메일이 한 폴더에 함께 쌓입니다.',
      inputSchema: {
        rootFolder: z
          .enum(['inbox', 'sent'])
          .default('inbox')
          .describe(
            '어디를 검사할지. inbox=받은편지함, sent=보낸편지함 (동일한 제목/본문 기준 분류 규칙을 그대로 적용). ' +
              '검사 대상만 다를 뿐, 이동 대상 폴더는 항상 받은편지함 쪽 트리로 통일됩니다.',
          ),
        scope: z
          .enum(['all', 'unread', 'today'])
          .default('all')
          .describe(
            '검사 범위. all=전체, unread=읽지 않은 메일만 훑어보기(단, 이동은 하지 않음), today=오늘 항목만',
          ),
        dryRun: z
          .boolean()
          .default(false)
          .describe('true 이면 실제로 이동하지 않고 이동 계획만 미리 보여줍니다.'),
        maxCount: z.number().int().positive().max(5000).default(200).describe('검사할 최대 메일 수'),
      },
    },
    async ({ rootFolder, scope, dryRun, maxCount }) => {
      const { customers, categories } = await loadConfig();
      const mails = await outlookClient.listInboxMails(scope, maxCount, rootFolder);

      // 규칙 1: 읽지 않은 메일은 사용자가 아직 확인하지 않았다는 뜻이므로 절대 옮기지 않는다.
      const unreadSkipped = mails.filter((m) => m.unread);
      const candidates = mails.filter((m) => !m.unread);

      const classified = candidates.map((mail) => ({
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

      // 규칙 2: 고객사/카테고리/캘린더 규칙에 확실히 일치하지 않는(=default 로 떨어지는) 메일은
      // 임의로 미분류 폴더에 넣지 않고, 사용자에게 물어볼 "확인 필요" 목록으로만 남긴다.
      const confident = classified.filter((c) => c.classification.matchedType !== 'default');
      const needsReview = classified.filter((c) => c.classification.matchedType === 'default');

      const moved: Array<{ subject: string; from: string; folder: string; reason: string }> = [];
      const errors: Array<{ subject: string; error: string }> = [];

      if (dryRun) {
        for (const { mail, classification } of confident) {
          moved.push({
            subject: mail.subject,
            from: mail.senderName,
            folder: classification.folderPath,
            reason: classification.reason,
          });
        }
      } else if (confident.length > 0) {
        // 건마다 PowerShell 프로세스를 새로 띄우면 대량 이동 시 매우 느려지므로,
        // 한 번의 Outlook 연결로 일괄 이동한다.
        // 주의: rootFolder 는 "어디를 스캔할지"만 결정한다. 고객사/카테고리 폴더는
        // 보낸편지함을 정리할 때도 항상 받은편지함 쪽 트리 하나로 모은다 (받은/보낸 메일이
        // 같은 고객사 폴더에 함께 쌓이도록) - 그래서 target 은 항상 'inbox' 로 고정한다.
        const result = await outlookClient.moveMailBatch(
          confident.map(({ mail, classification }) => ({
            entryId: mail.entryId,
            storeId: mail.storeId,
            targetPath: classification.folderPath,
          })),
          { rootFolder: 'inbox' },
        );

        const byEntryId = new Map(confident.map(({ mail, classification }) => [mail.entryId, { mail, classification }]));
        for (const m of result.moved) {
          const info = byEntryId.get(m.entryId);
          moved.push({
            subject: info?.mail.subject ?? m.entryId,
            from: info?.mail.senderName ?? '',
            folder: m.targetPath,
            reason: info?.classification.reason ?? '',
          });
        }
        for (const e of result.errors) {
          const info = byEntryId.get(e.entryId);
          const subject = info?.mail.subject ?? e.entryId;
          logger.error(`메일 이동 실패 ("${subject}"): ${e.error}`);
          errors.push({ subject, error: e.error });
        }
      }

      const lines = [
        `${dryRun ? '[미리보기] ' : ''}검사한 메일: ${mails.length}건, 이동 대상: ${moved.length}건, ` +
          `확인 필요: ${needsReview.length}건, 읽지 않아 건너뜀: ${unreadSkipped.length}건, 오류: ${errors.length}건`,
        '',
        ...moved.map((m) => `- "${m.subject}" (${m.from}) → ${m.folder}  [${m.reason}]`),
      ];

      if (needsReview.length > 0) {
        lines.push(
          '',
          '⚠ 확인이 필요한 메일 (고객사/카테고리가 애매해서 임의로 옮기지 않았습니다. 사용자에게 어느 폴더로 보낼지 물어보세요):',
          ...needsReview.map(
            (c) => `- "${c.mail.subject}" (${c.mail.senderName})  entryId=${c.mail.entryId}${c.mail.storeId ? ` storeId=${c.mail.storeId}` : ''}`,
          ),
        );
      }

      if (unreadSkipped.length > 0) {
        lines.push('', `읽지 않아 건너뛴 메일: ${unreadSkipped.length}건 (예: "${unreadSkipped[0]?.subject}")`);
      }

      if (errors.length > 0) {
        lines.push('', '오류:', ...errors.map((e) => `- "${e.subject}": ${e.error}`));
      }

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
