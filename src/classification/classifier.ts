import type { CategoriesConfig, CategoryRule, CustomersConfig } from '../config/schema.js';

export interface ClassifiableMail {
  subject: string;
  bodyPreview: string;
  senderName: string;
  senderEmail: string;
  ccNames: string[];
  ccEmails: string[];
  attachmentNames?: string[];
  /** 'calendar' 이면 회의 요청/응답, 연차 등 - 내용 판별 없이 항상 calendar_folder 로 이동 */
  itemType?: 'mail' | 'calendar';
}

export interface ClassificationResult {
  /** 받은 편지함 기준 상대 폴더 경로 ("/" 구분) */
  folderPath: string;
  matchedType: 'customer' | 'category' | 'calendar' | 'default';
  matchedName: string;
  reason: string;
}

function domainOf(email: string): string {
  const at = email.lastIndexOf('@');
  return at === -1 ? '' : email.slice(at + 1).toLowerCase();
}

function normalizeDomain(domain: string): string {
  return domain.toLowerCase().replace(/^@/, '');
}

function matchCategory(
  category: CategoryRule,
  subjectBodyHaystack: string,
  senderDomain: string,
): ClassificationResult | null {
  for (const keyword of category.keywords) {
    if (subjectBodyHaystack.includes(keyword.toLowerCase())) {
      return {
        folderPath: category.folder,
        matchedType: 'category',
        matchedName: category.name,
        reason: `키워드 "${keyword}" 이(가) 제목/본문/첨부파일명에서 발견되었습니다.`,
      };
    }
  }
  for (const domain of category.sender_domains) {
    if (senderDomain && senderDomain === normalizeDomain(domain)) {
      return {
        folderPath: category.folder,
        matchedType: 'category',
        matchedName: category.name,
        reason: `발신자 도메인 "${domain}" 이(가) 일치했습니다.`,
      };
    }
  }
  return null;
}

/**
 * 제목/본문/발신자/참조/첨부파일명을 종합하여 메일이 어느 고객사 또는 카테고리
 * 폴더로 이동해야 하는지 판단합니다. 발신자의 이메일 도메인만으로 고객사를
 * 판단하지 않고, 메일 내용 전반에서 고객사 별칭을 찾습니다 (customers.yml 참고).
 *
 * 판별 순서: 캘린더성 항목 -> priority 카테고리(사내 공지 등, 고객사 언급이 있어도
 * 항상 우선) -> 고객사 -> 나머지 카테고리 -> 기본 폴더.
 */
export function classifyMail(
  mail: ClassifiableMail,
  customersConfig: CustomersConfig,
  categoriesConfig: CategoriesConfig,
): ClassificationResult {
  if (mail.itemType === 'calendar') {
    return {
      folderPath: categoriesConfig.calendar_folder,
      matchedType: 'calendar',
      matchedName: categoriesConfig.calendar_folder,
      reason: '회의 요청/응답, 연차 등 캘린더성 항목입니다.',
    };
  }

  const attachmentNames = mail.attachmentNames ?? [];
  const subjectBodyHaystack = `${mail.subject}\n${mail.bodyPreview}\n${attachmentNames.join(' ')}`.toLowerCase();
  const senderDomain = domainOf(mail.senderEmail);

  const priorityCategories = categoriesConfig.categories.filter((c) => c.priority);
  const normalCategories = categoriesConfig.categories.filter((c) => !c.priority);

  for (const category of priorityCategories) {
    const result = matchCategory(category, subjectBodyHaystack, senderDomain);
    if (result) return result;
  }

  const contentHaystack = [
    mail.subject,
    mail.bodyPreview,
    mail.senderName,
    mail.ccNames.join(' '),
    mail.ccEmails.join(' '),
    attachmentNames.join(' '),
  ]
    .join('\n')
    .toLowerCase();

  for (const customer of customersConfig.customers) {
    for (const alias of customer.aliases) {
      if (contentHaystack.includes(alias.toLowerCase())) {
        return {
          folderPath: `${categoriesConfig.customer_root_folder}/${customer.name}`,
          matchedType: 'customer',
          matchedName: customer.name,
          reason: `별칭 "${alias}" 이(가) 제목/본문/참조/첨부파일명에서 발견되었습니다.`,
        };
      }
    }
  }

  for (const category of normalCategories) {
    const result = matchCategory(category, subjectBodyHaystack, senderDomain);
    if (result) return result;
  }

  return {
    folderPath: categoriesConfig.default_folder,
    matchedType: 'default',
    matchedName: categoriesConfig.default_folder,
    reason: '일치하는 고객사 또는 카테고리 규칙이 없습니다.',
  };
}
