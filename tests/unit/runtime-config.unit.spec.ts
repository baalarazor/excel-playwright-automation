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

  test('accepts trusted Microsoft workbook hosts and normalizes the cell', () => {
    const config = readRuntimeConfig({
      ...validEnvironment(),
      EXCEL_TARGET_CELL: ' xfd1048576 ',
    });

    expect(config.targetCell).toBe('XFD1048576');
  });

  for (const unsafeUrl of [
    'http://tenant.sharepoint.com/book.xlsx',
    'https://sharepoint.com.attacker.example/book.xlsx',
    'not-a-url',
  ]) {
    test(`rejects unsafe or malformed workbook URL: ${unsafeUrl}`, () => {
      expect(() =>
        readRuntimeConfig({ EXCEL_WORKBOOK_URL: unsafeUrl }),
      ).toThrow(ConfigurationError);
    });
  }

  for (const malformedCell of [
    'A0',
    '1A',
    'A-2',
    '=A2',
    'AAAA1',
    'XFE1',
    'A1048577',
  ]) {
    test(`rejects malformed target cell: ${malformedCell}`, () => {
      expect(() =>
        readRuntimeConfig({ ...validEnvironment(), EXCEL_TARGET_CELL: malformedCell }),
      ).toThrow(/Excel bounds/);
    });
  }

  test('rejects timeout values outside operational boundaries', () => {
    expect(() =>
      readRuntimeConfig({ ...validEnvironment(), ACTION_TIMEOUT_MS: '999' }),
    ).toThrow(/1000 to 120000/);
    expect(() =>
      readRuntimeConfig({ ...validEnvironment(), ACTION_TIMEOUT_MS: 'NaN' }),
    ).toThrow(/integer/);
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
