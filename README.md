# Excel Online `TODAY()` test

This TypeScript + Playwright project verifies that Excel Online workbook stores `=TODAY()` in A2 and displays today's date in a configured timezone.

## Setup

Requirements: Node.js 20+, Google Chrome, and an editable Excel Online workbook.

```bash
npm ci
cp config/test.env.example config/test.env
```

Set `EXCEL_WORKBOOK_URL` in `config/test.env`. The workbook must be publicly editable; this project deliberately has no Microsoft sign-in flow.

## Run

```bash
npm run check       # typecheck, lint, and fast unit tests
npm run test:e2e    # Chrome UI test
npm run test:demo   # visible Chrome run with video
npm run report      # open the HTML report
```

## What the E2E test does

1. Opens the configured workbook in Chrome.
2. Clears A2 and confirms the cell is empty.
3. Enters `=TODAY()` and verifies the formula bar.
4. Reads A2's displayed value from Excel's accessibility label.
5. Compares it with today's date in the configured locale and timezone, allowing a midnight boundary.
6. In two separate format tests, opens **More Number Formats → Custom** and verifies `yyyy-mm-dd` and `d mmm yyyy` (`3 Sep 2026`).
7. Clears A2 and resets it to the General format in the `afterEach` hook.

Excel's canvas grid is not a stable DOM source. The test therefore uses the selected cell's accessibility label; if that label is unavailable, the test fails with an actionable error rather than relying on clipboard or coordinates.

## Design notes

- The Name Box provides deterministic A2 navigation without grid coordinates.
- The Excel page object keeps the UI selectors in one place.
- Locale and timezone are explicit, so the expected date does not depend on the machine location.
- A2 is shared state; do not run this test concurrently against the same workbook.

Failures retain a screenshot, trace, and video under `test-results/`.

## Demo video

https://www.loom.com/share/ba9026f1f2d2435facd010f9b5a04959
