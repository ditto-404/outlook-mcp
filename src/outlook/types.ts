import { z } from 'zod';

export const RecipientSchema = z.object({
  name: z.string().default(''),
  email: z.string().default(''),
});
export type Recipient = z.infer<typeof RecipientSchema>;

export const MailSummarySchema = z.object({
  entryId: z.string(),
  storeId: z.string().default(''),
  subject: z.string().default(''),
  senderName: z.string().default(''),
  senderEmail: z.string().default(''),
  receivedTime: z.string(),
  unread: z.boolean().default(false),
  hasAttachments: z.boolean().default(false),
  attachmentNames: z.array(z.string()).default([]),
  bodyPreview: z.string().default(''),
  ccNames: z.array(z.string()).default([]),
  ccEmails: z.array(z.string()).default([]),
  folderPath: z.string().default(''),
  // 'mail' = 일반 메일(olMail), 'calendar' = 회의 요청/응답, 연차 등 캘린더성 항목
  itemType: z.enum(['mail', 'calendar']).default('mail'),
});
export type MailSummary = z.infer<typeof MailSummarySchema>;

export const AttachmentSchema = z.object({
  fileName: z.string(),
  size: z.number(),
});

export const MailDetailSchema = MailSummarySchema.extend({
  sentOn: z.string().nullable().default(null),
  to: z.array(RecipientSchema).default([]),
  cc: z.array(RecipientSchema).default([]),
  body: z.string().default(''),
  attachments: z.array(AttachmentSchema).default([]),
});
export type MailDetail = z.infer<typeof MailDetailSchema>;

export const MoveResultSchema = z.object({
  newEntryId: z.string(),
  folderPath: z.string(),
});

export const MoveBatchResultSchema = z.object({
  moved: z.array(
    z.object({
      entryId: z.string(),
      targetPath: z.string(),
      newEntryId: z.string(),
    }),
  ),
  errors: z.array(
    z.object({
      entryId: z.string(),
      targetPath: z.string(),
      error: z.string(),
    }),
  ),
});
export type MoveBatchResult = z.infer<typeof MoveBatchResultSchema>;
export type MoveResult = z.infer<typeof MoveResultSchema>;

export const SearchResultSchema = z.object({
  items: z.array(MailSummarySchema),
  scanned: z.number(),
  truncated: z.boolean(),
});
export type SearchResult = z.infer<typeof SearchResultSchema>;

export const DraftResultSchema = z.object({
  draftEntryId: z.string(),
  draftStoreId: z.string().default(''),
  subject: z.string().default(''),
  to: z.string().default(''),
});
export type DraftResult = z.infer<typeof DraftResultSchema>;

export const DraftSummarySchema = z.object({
  entryId: z.string(),
  storeId: z.string().default(''),
  subject: z.string().default(''),
  to: z.string().default(''),
  cc: z.string().default(''),
  lastModifiedTime: z.string().default(''),
  createdTime: z.string().default(''),
  hasAttachments: z.boolean().default(false),
  bodyPreview: z.string().default(''),
});
export type DraftSummary = z.infer<typeof DraftSummarySchema>;

export const DraftListResultSchema = z.object({
  items: z.array(DraftSummarySchema),
  totalCount: z.number().default(0),
});
export type DraftListResult = z.infer<typeof DraftListResultSchema>;

export interface FolderNode {
  name: string;
  path: string;
  itemCount: number;
  unreadCount: number;
  children: FolderNode[];
}

export const FolderNodeSchema: z.ZodType<FolderNode> = z.lazy(() =>
  z.object({
    name: z.string(),
    path: z.string(),
    itemCount: z.number(),
    unreadCount: z.number(),
    children: z.array(FolderNodeSchema),
  }),
);
