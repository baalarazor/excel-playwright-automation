export interface CalendarDate {
  readonly day: number;
  readonly month: number;
  readonly year: number;
}

const DATE_STYLES = ['short', 'medium', 'long', 'full'] as const;

export function calendarDateAt(
  instant: Date,
  locale: string,
  timeZone: string,
): CalendarDate {
  const parts = new Intl.DateTimeFormat(locale, {
    day: 'numeric',
    month: 'numeric',
    timeZone,
    year: 'numeric',
  }).formatToParts(instant);

  return {
    day: numericPart(parts, 'day'),
    month: numericPart(parts, 'month'),
    year: numericPart(parts, 'year'),
  };
}

export function observedCalendarDates(
  startedAt: Date,
  finishedAt: Date,
  locale: string,
  timeZone: string,
): readonly CalendarDate[] {
  if (finishedAt.getTime() < startedAt.getTime()) {
    throw new RangeError('finishedAt must be at or after startedAt.');
  }

  const first = calendarDateAt(startedAt, locale, timeZone);
  const last = calendarDateAt(finishedAt, locale, timeZone);
  return sameCalendarDate(first, last) ? [first] : [first, last];
}

export function matchesDisplayedDate(
  displayedValue: string,
  expectedDates: readonly CalendarDate[],
  locale: string,
): boolean {
  if (expectedDates.length === 0) return false;
  const actual = normalize(displayedValue.split(/\r?\n|\t/, 1)[0] ?? '');
  if (!actual) return false;

  return expectedDates.some((date) =>
    dateRepresentations(date, locale).has(actual),
  );
}

export function describeCalendarDate(date: CalendarDate): string {
  return `${date.year.toString().padStart(4, '0')}-${date.month
    .toString()
    .padStart(2, '0')}-${date.day.toString().padStart(2, '0')}`;
}

function dateRepresentations(
  date: CalendarDate,
  locale: string,
): ReadonlySet<string> {
  // UTC keeps the supplied civil date stable; locale controls only its representation.
  const instant = new Date(Date.UTC(date.year, date.month - 1, date.day, 12));
  const values = new Set<string>();

  for (const dateStyle of DATE_STYLES) {
    values.add(normalize(new Intl.DateTimeFormat(locale, { dateStyle, timeZone: 'UTC' }).format(instant)));
  }
  values.add(
    normalize(
      new Intl.DateTimeFormat(locale, {
        day: 'numeric',
        month: 'numeric',
        timeZone: 'UTC',
        year: 'numeric',
      }).format(instant),
    ),
  );
  values.add(
    normalize(
      new Intl.DateTimeFormat(locale, {
        day: '2-digit',
        month: '2-digit',
        timeZone: 'UTC',
        year: 'numeric',
      }).format(instant),
    ),
  );

  return values;
}

function normalize(value: string): string {
  return value
    .normalize('NFKC')
    .replace(/[\u200e\u200f\u202a-\u202e]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleLowerCase();
}

function numericPart(
  parts: readonly Intl.DateTimeFormatPart[],
  type: Intl.DateTimeFormatPartTypes,
): number {
  const value = parts.find((part) => part.type === type)?.value;
  if (!value) throw new RangeError(`Intl did not return a ${type} date part.`);
  return Number(value);
}

function sameCalendarDate(left: CalendarDate, right: CalendarDate): boolean {
  return left.year === right.year && left.month === right.month && left.day === right.day;
}
