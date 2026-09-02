import fs from 'node:fs';
import path from 'node:path';

import { expect, test as setup } from '@playwright/test';

import {
  loadEnvironmentFile,
  readRuntimeConfig,
  requireCredentials,
} from '../../src/config/runtime-config';
import { MicrosoftLoginPage } from '../../src/pages/microsoft-login.page';
import { ExcelWorkbookPage } from '../../src/pages/excel-workbook.page';

const authFile = path.resolve('playwright/.auth/user.json');

setup('authenticate Microsoft test account', async ({ browser }) => {
  loadEnvironmentFile();
  const config = readRuntimeConfig();
  const context = await browser.newContext({
    locale: config.locale,
    timezoneId: config.timeZone,
    ...(fs.existsSync(authFile) ? { storageState: authFile } : {}),
  });
  const page = await context.newPage();

  await page.goto(config.workbookUrl, { waitUntil: 'domcontentloaded' });
  const loginHosts = ['login.live.com', 'login.microsoft.com', 'login.microsoftonline.com'];
  const isLoginHost = loginHosts.includes(new URL(page.url()).hostname.toLowerCase());
  if (isLoginHost) {
    await new MicrosoftLoginPage(page).signIn(requireCredentials(config), config.staySignedIn);
  }

  await expect(page, 'Authentication remained on the Microsoft login page').not.toHaveURL(
    /\/\/login\.(?:live|microsoft|microsoftonline)\.com\//i,
    { timeout: config.navigationTimeoutMs },
  );
  await new ExcelWorkbookPage(page, config.workbookTimeoutMs).open(config.workbookUrl);

  fs.mkdirSync(path.dirname(authFile), { recursive: true });
  await context.storageState({ path: authFile });
  await context.close();
});
