import { z } from 'zod';
import { runOutlookScript } from './powershellBridge.js';
import {
  DraftResultSchema,
  FolderNodeSchema,
  MailDetailSchema,
  MailSummarySchema,
  MoveResultSchema,
  SearchResultSchema,
  type DraftResult,
  type FolderNode,
  type MailDetail,
  type MailSummary,
  type MoveResult,
  type SearchResult,
} from './types.js';

export type OrganizeScope = 'all' | 'unread' | 'today';
export type DraftMode = 'reply' | 'replyAll' | 'new';

const LONG_TIMEOUT_MS = 120_000;

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
  async listInboxMails(scope: OrganizeScope, maxCount = 200): Promise<MailSummary[]> {
    const data = await runOutlookScript<{ items: unknown[]; count: number }>(
      'list-inbox.ps1',
      { scope, maxCount },
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
