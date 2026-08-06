import { z } from 'zod';

export const CustomerSchema = z.object({
  name: z.string().min(1, 'customers.yml: name 은 비어 있을 수 없습니다.'),
  aliases: z.array(z.string().min(1)).min(1, 'customers.yml: aliases 는 최소 1개 이상이어야 합니다.'),
});
export type Customer = z.infer<typeof CustomerSchema>;

export const CustomersConfigSchema = z.object({
  customers: z.array(CustomerSchema).default([]),
});
export type CustomersConfig = z.infer<typeof CustomersConfigSchema>;

export const CategoryRuleSchema = z.object({
  name: z.string().min(1, 'categories.yml: name 은 비어 있을 수 없습니다.'),
  folder: z.string().min(1, 'categories.yml: folder 는 비어 있을 수 없습니다.'),
  keywords: z.array(z.string()).default([]),
  sender_domains: z.array(z.string()).default([]),
  // true 이면 고객사 판별보다 먼저 검사한다. 사내 공지/행정 메일처럼, 본문에 고객사
  // 이름이 우연히 언급되더라도 항상 이 카테고리로 보내야 하는 경우에 사용한다.
  priority: z.boolean().default(false),
});
export type CategoryRule = z.infer<typeof CategoryRuleSchema>;

export const CategoriesConfigSchema = z.object({
  categories: z.array(CategoryRuleSchema).default([]),
  default_folder: z.string().min(1).default('미분류'),
  customer_root_folder: z.string().min(1).default('고객사'),
  // 회의 요청/응답, 연차 등 캘린더성 항목(올Mail 이 아닌 항목)이 이동할 폴더.
  // 이 폴더로 가는 항목은 고객사/카테고리 키워드 판별을 거치지 않고 항상 여기로 간다.
  calendar_folder: z.string().min(1).default('일정-캘린더알림'),
});
export type CategoriesConfig = z.infer<typeof CategoriesConfigSchema>;
