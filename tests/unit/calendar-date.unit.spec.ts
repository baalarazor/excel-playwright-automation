import { expect, test } from '@playwright/test';

import {
  calendarDateAt,
  describeCalendarDate,
  matchesDisplayedDate,
  observedCalendarDates,
} from '../../src/domain/calendar-date';

test.describe('Excel displayed date matching', () => {
  const septemberSecond = { day: 2, month: 9, year: 2026 };

  test('matches common US Excel date representations', () => {
    expect(matchesDisplayedDate('9/2/2026', [septemberSecond], 'en-US')).toBe(true);
    expect(matchesDisplayedDate('Sep 2, 2026', [septemberSecond], 'en-US')).toBe(true);
  });

  test('matches localized and clipboard-delimited representations', () => {
    expect(matchesDisplayedDate('02.09.2026\tignored', [septemberSecond], 'de-DE')).toBe(
      true,
    );
    expect(matchesDisplayedDate('02/09/2026\r\n', [septemberSecond], 'en-GB')).toBe(true);
  });

  test('rejects wrong, malformed, empty, and formula-like values', () => {
    for (const value of ['', 'not a date', '=TODAY()', '9/3/2026', '2026-99-99']) {
      expect(matchesDisplayedDate(value, [septemberSecond], 'en-US')).toBe(false);
    }
  });

  test('uses the configured time zone at a UTC date boundary', () => {
    const instant = new Date('2026-01-01T00:30:00.000Z');
    expect(calendarDateAt(instant, 'en-US', 'America/Los_Angeles')).toEqual({
      day: 31,
      month: 12,
      year: 2025,
    });
  });

  test('accepts either calendar date when execution crosses midnight', () => {
    const dates = observedCalendarDates(
      new Date('2026-09-02T23:59:59.900Z'),
      new Date('2026-09-03T00:00:00.100Z'),
      'en-GB',
      'UTC',
    );

    expect(dates.map(describeCalendarDate)).toEqual(['2026-09-02', '2026-09-03']);
    expect(matchesDisplayedDate('03/09/2026', dates, 'en-GB')).toBe(true);
  });

  test('rejects an invalid execution window', () => {
    expect(() =>
      observedCalendarDates(
        new Date('2026-09-03T00:00:00Z'),
        new Date('2026-09-02T00:00:00Z'),
        'en-US',
        'UTC',
      ),
    ).toThrow(RangeError);
  });
});
