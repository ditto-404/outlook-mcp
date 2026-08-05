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
});
export type CategoryRule = z.infer<typeof CategoryRuleSchema>;

export const CategoriesConfigSchema = z.object({
  categories: z.array(CategoryRuleSchema).default([]),
  default_folder: z.string().min(1).default('미분류'),
  customer_root_folder: z.string().min(1).default('고객사'),
});
export type CategoriesConfig = z.infer<typeof CategoriesConfigSchema>;
