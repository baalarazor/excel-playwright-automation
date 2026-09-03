import { expect, test } from '@playwright/test';

import { loadEnvironmentFile, readRuntimeConfig } from '../../src/config/runtime-config';
import {
  describeCalendarDate,
  matchesDisplayedDate,
  observedCalendarDates,
} from '../../src/domain/calendar-date';
import { ExcelWorkbookPage } from '../../src/pages/excel-workbook.page';

loadEnvironmentFile();

test.describe('Excel Online TODAY()', () => {
  test.describe.configure({ mode: 'serial' });

  test('returns the current date in A2', async ({ page }, testInfo) => {
    const config = readRuntimeConfig();
    const workbook = new ExcelWorkbookPage(page, config.workbookTimeoutMs);
    let cellWasChanged = false;

    await test.step('Open the dedicated Excel Online workbook', async () => {
      await workbook.open(config.workbookUrl);
    });

    await test.step(`Clear ${config.targetCell} before the test`, async () => {
      await workbook.clearCell(config.targetCell);
    });

    try {
      const startedAt = new Date();
      await test.step(`Execute =TODAY() in ${config.targetCell}`, async () => {
        // Cleanup owns the cell before mutation begins, including partial UI failures.
        cellWasChanged = true;
        await workbook.enterFormula(config.targetCell, '=TODAY()');
      });

      const formula = await test.step('Verify the entered formula', async () =>
        workbook.readSelectedCellFormula(),
      );
      expect(formula.replace(/\s/g, '').toUpperCase()).toBe('=TODAY()');

      const displayedValue = await test.step('Read the calculated cell value', async () =>
        workbook.readSelectedCellValue(config.targetCell),
      );
      const finishedAt = new Date();
      const expectedDates = observedCalendarDates(
        startedAt,
        finishedAt,
        config.locale,
        config.timeZone,
      );

      await testInfo.attach('date-verification.txt', {
        body: Buffer.from(
          [
            `Cell: ${config.targetCell}`,
            `Formula: ${formula}`,
            `Displayed: ${displayedValue}`,
            `Expected: ${expectedDates.map(describeCalendarDate).join(' or ')}`,
            `Time zone: ${config.timeZone}`,
            `Locale: ${config.locale}`,
          ].join('\n'),
        ),
        contentType: 'text/plain',
      });

      expect(
        matchesDisplayedDate(displayedValue, expectedDates, config.locale),
        `Expected ${config.targetCell} value "${displayedValue}" to represent ${expectedDates
          .map(describeCalendarDate)
          .join(' or ')}`,
      ).toBe(true);
    } finally {
      if (cellWasChanged) {
        await workbook.clearCell(config.targetCell).catch(() => undefined);
      }
    }
  });

  test('displays TODAY() using Excel\'s Short Date number format', async ({ page }) => {
    const config = readRuntimeConfig();
    test.skip(config.locale !== 'en-US', 'This display-pattern assertion is specific to en-US.');
    const workbook = new ExcelWorkbookPage(page, config.workbookTimeoutMs);
    let cellWasChanged = false;

    await workbook.open(config.workbookUrl);
    await workbook.clearCell(config.targetCell);
    await workbook.setNumberFormat(config.targetCell, 'General');

    try {
      const startedAt = new Date();
      cellWasChanged = true;
      await workbook.enterFormula(config.targetCell, '=TODAY()');
      await workbook.setNumberFormat(config.targetCell, 'Short Date');

      const displayedValue = await workbook.readSelectedCellValue(config.targetCell);
      expect(await workbook.readNumberFormat(), 'Excel should apply the Date format category').toBe(
        'Date',
      );
      const expectedDates = observedCalendarDates(
        startedAt,
        new Date(),
        config.locale,
        config.timeZone,
      );

      expect(
        matchesDisplayedDate(displayedValue, expectedDates, config.locale),
        `Expected ${config.targetCell} value "${displayedValue}" to represent today's date`,
      ).toBe(true);
      expect(displayedValue, 'Excel Short Date should use the M/D/YYYY pattern').toMatch(
        /^\d{1,2}\/\d{1,2}\/\d{4}$/,
      );
    } finally {
      if (cellWasChanged) {
        await workbook.clearCell(config.targetCell).catch(() => undefined);
        await workbook.setNumberFormat(config.targetCell, 'General').catch(() => undefined);
      }
    }
  });
});
