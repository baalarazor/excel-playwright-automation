# Demo and verbal walkthrough

## Record the demo

1. Complete `config/test.env` using the example file.
2. Close confidential applications and notification popups.
3. Run `npm run test:demo`.
4. The headed Chrome journey is visible while Playwright also records video.
5. Find `video.webm` under `test-results/`; open the HTML report with `npm run report`.

Do not publish a recording that contains a Microsoft password, session cookie, workbook sharing token, or personal account data. Authentication recording is disabled; the test video starts from the stored session.

## Suggested 60-second narration

“This project uses Playwright Test with TypeScript and drives the stable Google Chrome channel. Configuration and credentials come from an ignored environment file or CI secrets. Authentication is isolated in a setup project and its storage state is reused, so the business test remains focused on behavior.

The test opens a dedicated Excel Online workbook, navigates to A2 through the Name Box, enters the exact `=TODAY()` formula, and checks the formula bar. It then copies the calculated cell value and validates it against localized date formats for the configured time zone. The assertion allows both dates only if execution crosses midnight. A finally block clears A2.

Fast unit tests cover malformed configuration, phishing-style URLs, timeout boundaries, locales, bad date values, and midnight behavior. An API smoke test separates endpoint availability from UI failures. Reports, screenshots, traces, and optional video support diagnosis.

The main limitations are Microsoft MFA or Conditional Access, upstream UI changes, and concurrent use of one shared workbook. The workarounds are a dedicated automation account and workbook per execution lane, cached auth state, serialized UI execution, and resilient accessible-name locators. Graph API or Office Scripts would be faster for formula validation, but neither proves the required Chrome user journey.”
