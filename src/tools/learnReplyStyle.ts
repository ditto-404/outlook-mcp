import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { outlookClient } from '../outlook/client.js';

const MIN_BODY_LENGTH = 150;

export function registerLearnReplyStyleTool(server: McpServer): void {
  server.registerTool(
    'learn_reply_style',
    {
      title: '답장 문체 학습용 자료 수집',
      description:
        '보낸 편지함에서 실제로 작성한 메일들을 모아옵니다. 이 도구 자체는 파일을 쓰지 않습니다 - ' +
        'AI가 반환된 메일 본문들을 읽고 인사말/맺음말/서명/톤/자주 쓰는 표현의 패턴을 분석해서 ' +
        'config/reply_style.md 초안을 작성하는 데 사용하세요. ' +
        '초기 설정 시 "내 메일 문체를 학습해서 reply_style.md 만들어줘" 같은 요청에 사용하세요.',
      inputSchema: {
        maxCount: z.number().int().positive().max(200).default(40).describe('가져올 최대 보낸 메일 수 (최근 순)'),
      },
    },
    async ({ maxCount }) => {
      const mails = await outlookClient.listInboxMails('all', maxCount * 3, 'sent');

      const samples = mails
        .filter((m) => m.itemType === 'mail')
        .filter((m) => m.bodyPreview.trim().length >= MIN_BODY_LENGTH)
        .slice(0, maxCount);

      if (samples.length === 0) {
        return {
          content: [
            {
              type: 'text',
              text: '분석할 만한 보낸 메일을 찾지 못했습니다 (본문이 너무 짧은 메일만 있음). maxCount 를 늘려서 다시 시도해보세요.',
            },
          ],
        };
      }

      const lines = [
        `보낸 편지함에서 실제 작성한 메일 ${samples.length}건을 가져왔습니다.`,
        '아래 본문들을 읽고 다음을 분석해서 config/reply_style.md 초안을 작성하세요:',
        '- 인사말 패턴 (예: "OOO님, 안녕하세요" 를 쓰는지, 소속/직함을 붙이는지)',
        '- 맺음말 패턴 (예: "감사합니다", "확인 부탁드립니다" 등 실제로 반복되는 문구)',
        '- 서명 형식 (이름/직급/부서/연락처를 어떤 순서·형식으로 쓰는지)',
        '- 전반적인 톤 (문장 길이, 존댓말 격식 수준, 이모지·기호 사용 여부)',
        '- 자주 반복되는 답장 유형이 있다면(예: 기술지원 접수 확인, 방문 일정 조율 등) 짧은 템플릿으로 정리',
        '',
        '주의: 아래 본문에는 실제 발신자/수신자 이름과 사내 정보가 포함되어 있습니다.',
        'reply_style.md 는 로컬 파일(.gitignore 처리됨)이므로 그대로 반영해도 되지만, 공개 저장소로',
        '내보내거나 다른 사람에게 공유하지 마세요.',
        '',
        '=== 메일 샘플 ===',
      ];

      for (const m of samples) {
        lines.push('', `--- "${m.subject}" ---`, m.bodyPreview.trim());
      }

      return {
        content: [{ type: 'text', text: lines.join('\n') }],
      };
    },
  );
}
