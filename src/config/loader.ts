import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { parse as parseYaml } from 'yaml';
import { logger } from '../utils/logger.js';
import { CategoriesConfigSchema, CustomersConfigSchema, type CategoriesConfig, type CustomersConfig } from './schema.js';

export interface AppConfig {
  customers: CustomersConfig;
  categories: CategoriesConfig;
  replyStyle: string;
}

/**
 * config 디렉터리 위치. OUTLOOK_MCP_CONFIG_DIR 환경변수로 재정의하면
 * 조직마다 다른 config 세트를 코드 변경 없이 사용할 수 있습니다.
 */
export function resolveConfigDir(): string {
  const override = process.env.OUTLOOK_MCP_CONFIG_DIR;
  return override ? path.resolve(override) : path.resolve(process.cwd(), 'config');
}

let cached: AppConfig | null = null;

export async function loadConfig(forceReload = false): Promise<AppConfig> {
  if (cached && !forceReload) return cached;

  const dir = resolveConfigDir();
  const customersPath = path.join(dir, 'customers.yml');
  const categoriesPath = path.join(dir, 'categories.yml');
  const replyStylePath = path.join(dir, 'reply_style.md');

  const [customersRaw, categoriesRaw, replyStyle] = await Promise.all([
    readFile(customersPath, 'utf8').catch((err: NodeJS.ErrnoException) => {
      throw new Error(`고객사 설정 파일을 읽을 수 없습니다 (${customersPath}): ${err.message}`);
    }),
    readFile(categoriesPath, 'utf8').catch((err: NodeJS.ErrnoException) => {
      throw new Error(`카테고리 설정 파일을 읽을 수 없습니다 (${categoriesPath}): ${err.message}`);
    }),
    readFile(replyStylePath, 'utf8').catch(() => {
      logger.warn(`reply_style.md 를 찾을 수 없습니다 (${replyStylePath}). 스타일 가이드 없이 답장을 준비합니다.`);
      return '';
    }),
  ]);

  const customers = CustomersConfigSchema.parse(parseYaml(customersRaw) ?? {});
  const categories = CategoriesConfigSchema.parse(parseYaml(categoriesRaw) ?? {});

  const seen = new Set<string>();
  for (const customer of customers.customers) {
    if (seen.has(customer.name)) {
      throw new Error(`customers.yml 에 중복된 고객사 이름이 있습니다: ${customer.name}`);
    }
    seen.add(customer.name);
  }

  cached = { customers, categories, replyStyle };
  logger.info(
    `설정 로드 완료: 고객사 ${customers.customers.length}건, 카테고리 ${categories.categories.length}건 (config: ${dir})`,
  );
  return cached;
}
