# Excel Online `TODAY()` E2E test

A TypeScript + Playwright test that opens an editable workbook in stable Google Chrome, signs in to Microsoft, executes `=TODAY()` in `A2`, and verifies the result is the current date.

## Prerequisites

- Node.js 20+
- Google Chrome
- A dedicated Microsoft automation account without an interactive MFA requirement
- A dedicated editable workbook URL in OneDrive or SharePoint

## Setup

```bash
npm ci
cp config/test.env.example config/test.env
```

Edit `config/test.env`. Real credentials, the cached authenticated session, videos, traces, and reports are ignored by Git.

## Run

```bash
npm run check       # typecheck, lint, and fast unit/negative tests
npm run test:api    # configured endpoint availability/performance contract
npm run test:e2e    # authenticated Chrome E2E test
npm run test:demo   # headed run with video recording
npm run report      # inspect the HTML report
```

The first E2E run logs in and writes `playwright/.auth/user.json`. Delete that file to force a new login. If MFA is mandatory, create the authenticated state interactively with Playwright UI mode or use a policy-approved automation account; never commit the state file.

## Evidence and cleanup

Failures retain a trace, screenshot, and video under `test-results/`. Demo mode records successful execution too. The test attaches the actual formula, displayed value, expected ISO date, locale, and time zone to the report. It clears `A2` in a `finally` block.

See [requirements and test strategy](docs/REQUIREMENTS.md) and [demo/narration guide](docs/DEMO.md) for architecture decisions, risks, limitations, workarounds, and alternative solutions.

The implementation follows Playwright's official guidance for [reusable authentication state](https://playwright.dev/docs/auth), [stable Chrome channels](https://playwright.dev/docs/browsers#google-chrome--microsoft-edge), and [failure evidence](https://playwright.dev/docs/test-use-options#recording-options). The expected function behavior comes from Microsoft's [`TODAY()` reference](https://support.microsoft.com/en-us/office/today-function-5eb3078d-a82c-4736-8930-2f51a028fdd9).

## GitHub

Initialize and publish after substituting your repository URL:

```bash
git init
git add .
git commit -m "Add Excel Online TODAY Playwright test"
git branch -M main
git remote add origin https://github.com/<owner>/<repository>.git
git push -u origin main
```

Add `MS_USERNAME`, `MS_PASSWORD`, and `EXCEL_WORKBOOK_URL` as GitHub Actions secrets. Quality checks run on pushes and pull requests; the external E2E job is intentionally manual to protect the shared account and workbook.
