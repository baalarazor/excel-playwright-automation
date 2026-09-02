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
});
