import { expect } from '@playwright/test';

import type { RuntimeConfig } from '../../src/config/runtime-config';
import {
  type CalendarDate,
  matchesDisplayedDate,
  observedCalendarDates,
} from '../../src/domain/calendar-date';
import type { ExcelWorkbookPage } from '../../src/pages/excel-workbook.page';

export interface TodayDateFormat {
  readonly code: string;
  readonly description: string;
  readonly expectedValue: (date: CalendarDate) => string;
}

export const TODAY_DATE_FORMATS = {
  iso: {
    code: 'yyyy-mm-dd',
    description: 'yyyy-mm-dd',
    expectedValue: (date) =>
      `${date.year}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`,
  },
  textMonth: {
    code: 'd mmm yyyy',
    description: 'd mmm yyyy',
    expectedValue: (date) => {
      const month = new Intl.DateTimeFormat('en-US', {
        month: 'short',
        timeZone: 'UTC',
      }).format(new Date(Date.UTC(date.year, date.month - 1, date.day)));
      return `${date.day} ${month} ${date.year}`;
    },
  },
} satisfies Record<string, TodayDateFormat>;

export class TodayWorkflow {
  private cellWasChanged = false;

  public constructor(
    private readonly workbook: ExcelWorkbookPage,
    public readonly config: RuntimeConfig,
  ) {}

  public async prepare(): Promise<void> {
    await this.workbook.open(this.config.workbookUrl);
    await this.workbook.clearCell(this.config.targetCell);
    await this.workbook.setNumberFormat(this.config.targetCell, 'General');
  }

  public async cleanup(): Promise<void> {
    if (!this.cellWasChanged) return;

    await this.workbook.clearCell(this.config.targetCell).catch(() => undefined);
    await this.workbook.setNumberFormat(this.config.targetCell, 'General').catch(() => undefined);
  }

  public async enterTodayFormula(): Promise<Date> {
    const startedAt = new Date();
    this.cellWasChanged = true;
    await this.workbook.enterFormula(this.config.targetCell, '=TODAY()');
    return startedAt;
  }

  public async verifyTodayCalculation(startedAt: Date): Promise<void> {
    const formula = await this.workbook.readSelectedCellFormula();
    expect(formula.replace(/\s/g, '').toUpperCase()).toBe('=TODAY()');

    const displayedValue = await this.workbook.readSelectedCellValue(this.config.targetCell);
    const expectedDates = this.expectedTodayDates(startedAt);
    expect(
      matchesDisplayedDate(displayedValue, expectedDates, this.config.locale),
      `Expected ${this.config.targetCell} value "${displayedValue}" to represent today's date`,
    ).toBe(true);
  }

  public async verifyDateFormat(startedAt: Date, format: TodayDateFormat): Promise<void> {
    await this.workbook.setCustomNumberFormat(this.config.targetCell, format.code);
    const expectedDates = this.expectedTodayDates(startedAt);
    const expectedValues = expectedDates.map(format.expectedValue);

    await this.workbook.waitForSelectedCellValue(
      this.config.targetCell,
      (value) => expectedValues.includes(value),
      `Expected ${this.config.targetCell} to display today's date in ${format.description} format`,
    );
  }

  private expectedTodayDates(startedAt: Date): readonly CalendarDate[] {
    return observedCalendarDates(startedAt, new Date(), this.config.locale, this.config.timeZone);
  }
}
