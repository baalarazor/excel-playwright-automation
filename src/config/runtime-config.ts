import path from 'node:path';

import dotenv from 'dotenv';

export interface RuntimeConfig {
  readonly actionTimeoutMs: number;
  readonly locale: string;
  readonly navigationTimeoutMs: number;
  readonly targetCell: string;
  readonly timeZone: string;
  readonly workbookTimeoutMs: number;
  readonly workbookUrl: string;
}

const ALLOWED_WORKBOOK_HOSTS = [
  '1drv.ms',
  'live.com',
  'microsoft365.com',
  'office.com',
  'sharepoint.com',
] as const;

export function loadEnvironmentFile(
  environment: NodeJS.ProcessEnv = process.env,
): void {
  const configuredPath = environment.TEST_CONFIG_FILE;
  const filePath = configuredPath
    ? path.resolve(configuredPath)
    : path.resolve(process.cwd(), 'config/test.env');

  dotenv.config({ path: filePath, override: false, quiet: true });
}

export function readRuntimeConfig(
  environment: NodeJS.ProcessEnv = process.env,
): RuntimeConfig {
  const workbookUrl = required(environment, 'EXCEL_WORKBOOK_URL');
  validateWorkbookUrl(workbookUrl);

  const locale = environment.EXCEL_LOCALE?.trim() || 'en-US';
  const timeZone = environment.EXCEL_TIME_ZONE?.trim() || 'UTC';
  validateIntlSettings(locale, timeZone);

  const targetCell = (environment.EXCEL_TARGET_CELL?.trim() || 'A2').toUpperCase();
  if (!isExcelCellReference(targetCell)) {
    throw new ConfigurationError(
      'EXCEL_TARGET_CELL must be within Excel bounds A1:XFD1048576.',
    );
  }

  return {
    actionTimeoutMs: boundedInteger(environment, 'ACTION_TIMEOUT_MS', 15_000, 1_000, 120_000),
    locale,
    navigationTimeoutMs: boundedInteger(
      environment,
      'NAVIGATION_TIMEOUT_MS',
      60_000,
      5_000,
      180_000,
    ),
    targetCell,
    timeZone,
    workbookTimeoutMs: boundedInteger(
      environment,
      'WORKBOOK_TIMEOUT_MS',
      90_000,
      10_000,
      300_000,
    ),
    workbookUrl,
  };
}

export class ConfigurationError extends Error {
  public constructor(message: string) {
    super(message);
    this.name = 'ConfigurationError';
  }
}

function required(environment: NodeJS.ProcessEnv, name: string): string {
  const value = environment[name]?.trim();
  if (!value) {
    throw new ConfigurationError(`${name} is required. See config/test.env.example.`);
  }
  return value;
}

function validateWorkbookUrl(value: string): void {
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new ConfigurationError('EXCEL_WORKBOOK_URL must be a valid URL.');
  }

  if (url.protocol !== 'https:') {
    throw new ConfigurationError('EXCEL_WORKBOOK_URL must use HTTPS.');
  }

  const hostname = url.hostname.toLowerCase();
  const isAllowed = ALLOWED_WORKBOOK_HOSTS.some(
    (allowedHost) => hostname === allowedHost || hostname.endsWith(`.${allowedHost}`),
  );
  if (!isAllowed) {
    throw new ConfigurationError(
      `EXCEL_WORKBOOK_URL host "${hostname}" is not a trusted Microsoft workbook host.`,
    );
  }
}

function validateIntlSettings(locale: string, timeZone: string): void {
  try {
    new Intl.DateTimeFormat(locale, { timeZone }).format();
  } catch {
    throw new ConfigurationError('EXCEL_LOCALE or EXCEL_TIME_ZONE is invalid.');
  }
}

function boundedInteger(
  environment: NodeJS.ProcessEnv,
  name: string,
  fallback: number,
  minimum: number,
  maximum: number,
): number {
  const rawValue = environment[name];
  if (rawValue === undefined || rawValue.trim() === '') return fallback;

  const value = Number(rawValue);
  if (!Number.isSafeInteger(value) || value < minimum || value > maximum) {
    throw new ConfigurationError(`${name} must be an integer from ${minimum} to ${maximum}.`);
  }
  return value;
}

function isExcelCellReference(value: string): boolean {
  const match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(value);
  if (!match) return false;

  const [, columnLetters = '', rowText = ''] = match;
  const column = [...columnLetters].reduce(
    (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
    0,
  );
  const row = Number(rowText);
  return column >= 1 && column <= 16_384 && row >= 1 && row <= 1_048_576;
}
