import { expect, test } from '@playwright/test';

import { readRuntimeConfig } from '../../src/config/runtime-config';
import {
  describeCalendarDate,
  matchesDisplayedDate,
  observedCalendarDates,
} from '../../src/domain/calendar-date';
import { ExcelWorkbookPage } from '../../src/pages/excel-workbook.page';

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

  test('displays TODAY() using the yyyy-mm-dd date format', async ({ page }) => {
    const config = readRuntimeConfig();
    const workbook = new ExcelWorkbookPage(page, config.workbookTimeoutMs);
    let cellWasChanged = false;

    await workbook.open(config.workbookUrl);
    await workbook.clearCell(config.targetCell);
    await workbook.setNumberFormat(config.targetCell, 'General');

    try {
      const startedAt = new Date();
      cellWasChanged = true;
      await workbook.enterFormula(config.targetCell, '=TODAY()');
      await workbook.setIsoDateFormat(config.targetCell);

      const expectedDates = observedCalendarDates(
        startedAt,
        new Date(),
        config.locale,
        config.timeZone,
      );

      const expectedIsoDates = expectedDates.map(
        (date) =>
          `${date.year}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`,
      );
      await workbook.waitForSelectedCellValue(
        config.targetCell,
        (value) => expectedIsoDates.includes(value),
        `Expected ${config.targetCell} to display today's date in yyyy-mm-dd format`,
      );
    } finally {
      if (cellWasChanged) {
        await workbook.clearCell(config.targetCell).catch(() => undefined);
        await workbook.setNumberFormat(config.targetCell, 'General').catch(() => undefined);
      }
    }
  });

  test('displays TODAY() using the d mmm yyyy date format', async ({ page }) => {
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
      await workbook.setCustomNumberFormat(config.targetCell, 'd mmm yyyy');

      const expectedDates = observedCalendarDates(
        startedAt,
        new Date(),
        config.locale,
        config.timeZone,
      );
      const expectedTextDates = expectedDates.map((date) => {
        const month = new Intl.DateTimeFormat('en-US', {
          month: 'short',
          timeZone: 'UTC',
        }).format(new Date(Date.UTC(date.year, date.month - 1, date.day)));
        return `${date.day} ${month} ${date.year}`;
      });

      await workbook.waitForSelectedCellValue(
        config.targetCell,
        (value) => expectedTextDates.includes(value),
        `Expected ${config.targetCell} to display today's date in d mmm yyyy format`,
      );
    } finally {
      if (cellWasChanged) {
        await workbook.clearCell(config.targetCell).catch(() => undefined);
        await workbook.setNumberFormat(config.targetCell, 'General').catch(() => undefined);
      }
    }
  });
});
