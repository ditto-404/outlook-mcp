import { z } from 'zod';
import { runOutlookScript } from './powershellBridge.js';
import {
  DraftResultSchema,
  FolderNodeSchema,
  MailDetailSchema,
  MailSummarySchema,
  MoveBatchResultSchema,
  MoveResultSchema,
  SearchResultSchema,
  type DraftResult,
  type FolderNode,
  type MailDetail,
  type MailSummary,
  type MoveBatchResult,
  type MoveResult,
  type SearchResult,
} from './types.js';

export type OrganizeScope = 'all' | 'unread' | 'today';
export type DraftMode = 'reply' | 'replyAll' | 'new';
export type RootFolder = 'inbox' | 'sent';

// 발신자 주소를 Exchange 디렉터리에서 조회하는 COM 호출(GetExchangeUser)이 건마다
// 누적되면 수천 건 스캔 시 2분을 훌쩍 넘길 수 있어 넉넉히 잡는다.
const LONG_TIMEOUT_MS = 480_000;
const BATCH_TIMEOUT_MS = 580_000;

export interface MoveBatchItem {
  entryId: string;
  storeId?: string;
  targetPath: string;
}

export interface MoveBatchOptions {
  rootFolder?: RootFolder;
}

export interface SearchMailParams {
  folderPath?: string;
  subject?: string;
  body?: string;
  sender?: string;
  keywords?: string[];
  dateFrom?: string;
  dateTo?: string;
  maxCount?: number;
}

export interface SaveDraftParams {
  mode: DraftMode;
  sourceEntryId?: string;
  sourceStoreId?: string;
  to?: string[];
  cc?: string[];
  subject?: string;
  bodyHtml: string;
}

/** Classic Outlook COM 과의 상호작용을 감싸는 상위 레벨 클라이언트. */
export class OutlookClient {
  async listInboxMails(scope: OrganizeScope, maxCount = 200, rootFolder: RootFolder = 'inbox'): Promise<MailSummary[]> {
    const data = await runOutlookScript<{ items: unknown[]; count: number }>(
      'list-inbox.ps1',
      { scope, maxCount, rootFolder },
      LONG_TIMEOUT_MS,
    );
    return z.array(MailSummarySchema).parse(data.items);
  }

  async getMail(entryId: string, storeId?: string): Promise<MailDetail> {
    const data = await runOutlookScript('get-mail.ps1', { entryId, storeId });
    return MailDetailSchema.parse(data);
  }

  async moveMail(entryId: string, storeId: string | undefined, targetPath: string): Promise<MoveResult> {
    const data = await runOutlookScript('move-mail.ps1', {
      entryId,
      storeId,
      targetPath,
      createIfMissing: true,
    });
    return MoveResultSchema.parse(data);
  }

  /** move-mail.ps1 을 건마다 반복 호출하는 대신, 한 번의 Outlook 연결로 대량 이동한다. */
  async moveMailBatch(moves: MoveBatchItem[], options: MoveBatchOptions = {}): Promise<MoveBatchResult> {
    const data = await runOutlookScript(
      'move-mail-batch.ps1',
      { moves, rootFolder: options.rootFolder ?? 'inbox' },
      BATCH_TIMEOUT_MS,
    );
    return MoveBatchResultSchema.parse(data);
  }

  async searchMail(params: SearchMailParams): Promise<SearchResult> {
    const data = await runOutlookScript('search-mail.ps1', { ...params }, LONG_TIMEOUT_MS);
    return SearchResultSchema.parse(data);
  }

  async saveDraft(params: SaveDraftParams): Promise<DraftResult> {
    const data = await runOutlookScript('save-draft.ps1', { ...params });
    return DraftResultSchema.parse(data);
  }

  async listFolders(): Promise<FolderNode> {
    const data = await runOutlookScript('list-folders.ps1', {}, LONG_TIMEOUT_MS);
    return FolderNodeSchema.parse(data);
  }
}

export const outlookClient = new OutlookClient();
