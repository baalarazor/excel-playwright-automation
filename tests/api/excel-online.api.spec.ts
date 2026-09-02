import { expect, test } from '@playwright/test';

import { loadEnvironmentFile, readRuntimeConfig } from '../../src/config/runtime-config';

loadEnvironmentFile();

test('Excel Online workbook endpoint is available over HTTPS', async ({ request }) => {
  const config = readRuntimeConfig();
  const startedAt = performance.now();
  const response = await request.get(config.workbookUrl, {
    failOnStatusCode: false,
    timeout: config.navigationTimeoutMs,
  });
  const elapsedMs = performance.now() - startedAt;

  expect(response.status(), 'Endpoint returned a server error').toBeLessThan(500);
  expect(response.url()).toMatch(/^https:\/\//);
  expect(response.headers()['content-type']).toContain('text/html');
  expect(
    elapsedMs,
    `Workbook endpoint exceeded ${config.apiResponseBudgetMs}ms response budget`,
  ).toBeLessThan(config.apiResponseBudgetMs);
});
