import { expect, test } from '@playwright/test';

import { extractCellValueFromAccessibleLabel } from '../../src/pages/excel-workbook.page';

test.describe('Excel accessibility value extraction', () => {
  test('extracts the displayed value before formula metadata', () => {
    expect(
      extractCellValueFromAccessibleLabel(
        '9/2/2026 . A2 . Contains Formula . Selected by Guest Contributor . ',
        'A2',
      ),
    ).toBe('9/2/2026');
  });

  test('returns empty for a blank selected cell', () => {
    expect(
      extractCellValueFromAccessibleLabel('A2 . Selected by Guest Contributor .', 'A2'),
    ).toBe('');
  });

  test('does not return a value for another cell', () => {
    expect(extractCellValueFromAccessibleLabel('value . A20 . Selected .', 'A2')).toBe('');
  });
});
