import { defineConfig, devices } from '@playwright/test';

import { loadEnvironmentFile } from './src/config/runtime-config';

loadEnvironmentFile();

const demoMode = process.env.DEMO === 'true';
export default defineConfig({
  expect: { timeout: 10_000 },
  forbidOnly: Boolean(process.env.CI),
  fullyParallel: true,
  outputDir: 'test-results',
  reporter: [['list'], ['html', { open: 'never', outputFolder: 'playwright-report' }]],
  retries: process.env.CI ? 1 : 0,
  testDir: './tests',
  timeout: 120_000,
  use: {
    actionTimeout: Number(process.env.ACTION_TIMEOUT_MS ?? 15_000),
    navigationTimeout: Number(process.env.NAVIGATION_TIMEOUT_MS ?? 60_000),
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
  },
  projects: [
    {
      name: 'unit',
      testMatch: /.*\.unit\.spec\.ts/,
    },
    {
      name: 'chrome',
      testMatch: /.*\.e2e\.spec\.ts/,
      use: {
        ...devices['Desktop Chrome'],
        channel: 'chrome',
        headless: !demoMode,
        locale: process.env.EXCEL_LOCALE ?? 'en-US',
        permissions: ['clipboard-read', 'clipboard-write'],
        timezoneId: process.env.EXCEL_TIME_ZONE ?? 'UTC',
        video: demoMode ? 'on' : 'retain-on-failure',
      },
    },
  ],
  // A2 is shared server-side state: parallel UI workers would race each other.
  ...(process.env.CI ? { workers: 1 } : {}),
});
