import { expect, test } from '@playwright/test';

import {
  ConfigurationError,
  readRuntimeConfig,
  requireCredentials,
} from '../../src/config/runtime-config';

const validEnvironment = (): NodeJS.ProcessEnv => ({
  EXCEL_WORKBOOK_URL: 'https://tenant.sharepoint.com/sites/qa/workbook.xlsx',
});

test.describe('runtime configuration', () => {
  test('applies safe deterministic defaults', () => {
    const config = readRuntimeConfig(validEnvironment());

    expect(config.targetCell).toBe('A2');
    expect(config.locale).toBe('en-US');
    expect(config.timeZone).toBe('UTC');
    expect(config.actionTimeoutMs).toBe(15_000);
  });

  test('rejects an unsafe workbook URL', () => {
    expect(() => readRuntimeConfig({ EXCEL_WORKBOOK_URL: 'http://attacker.example/book.xlsx' })).toThrow(
      ConfigurationError,
    );
  });

  test('does not accept placeholder or missing credentials', () => {
    const config = readRuntimeConfig({
      ...validEnvironment(),
      MS_PASSWORD: 'replace-me',
      MS_USERNAME: '',
    });

    expect(() => requireCredentials(config)).toThrow(/required/);
  });

  test('returns credentials without logging or transforming them', () => {
    const config = readRuntimeConfig({
      ...validEnvironment(),
      MS_PASSWORD: ' secret ',
      MS_USERNAME: ' qa@example.com ',
    });

    expect(requireCredentials(config)).toEqual({
      password: 'secret',
      username: 'qa@example.com',
    });
  });
});
