import { test } from '@playwright/test';

import { readRuntimeConfig } from '../../src/config/runtime-config';
import { ExcelWorkbookPage } from '../../src/pages/excel-workbook.page';
import { TODAY_DATE_FORMATS, TodayWorkflow } from '../support/today-workflow';

test.describe('Excel Online TODAY()', () => {
  test.describe.configure({ mode: 'serial' });
  let today: TodayWorkflow;

  test.beforeEach(async ({ page }) => {
    const config = readRuntimeConfig();
    const workbook = new ExcelWorkbookPage(page, config.workbookTimeoutMs);
    today = new TodayWorkflow(workbook, config);
    await today.prepare();
  });

  test.afterEach(async () => {
    await today.cleanup();
  });

  test('returns the current date in A2', async () => {
    const startedAt = await today.enterTodayFormula();
    await today.verifyTodayCalculation(startedAt);
  });

  test('displays TODAY() using the yyyy-mm-dd date format', async () => {
    const startedAt = await today.enterTodayFormula();
    await today.verifyDateFormat(startedAt, TODAY_DATE_FORMATS.iso);
  });

  test('displays TODAY() using the d mmm yyyy date format', async () => {
    test.skip(today.config.locale !== 'en-US', 'This display-pattern assertion is specific to en-US.');
    const startedAt = await today.enterTodayFormula();
    await today.verifyDateFormat(startedAt, TODAY_DATE_FORMATS.textMonth);
  });
});
