# Requirements analysis and test strategy

## Acceptance criteria

1. Launch the installed stable Google Chrome channel.
2. Authenticate to Microsoft with configuration held outside source control.
3. Open an editable Excel Online workbook.
4. enter the exact formula `=TODAY()` in cell `A2`.
5. Assert both the formula and its displayed result.
6. Compare against the calendar date at execution time in an explicit time zone.
7. Produce reproducible evidence: HTML report, trace/screenshot on failure, and optional video.

## Architecture

- `config`: parses and validates environment input, including trusted-host and timeout boundaries.
- `domain`: pure, localized calendar-date matching with midnight protection.
- `pages`: UI responsibilities for Microsoft sign-in and Excel workbook interaction.
- `tests/auth`: reusable authentication-state setup.
- `tests/e2e`: business-readable acceptance scenario.
- `tests/api`: low-cost availability contract for the configured workbook endpoint.
- `tests/unit`: fast negative and boundary coverage without network or credentials.

This is intentionally a small Page Object + fixture architecture. Adding a dependency-injection container, generic base page, or repository abstraction would add indirection without another implementation to substitute.

## Risk-based coverage

| Risk | Coverage / control |
| --- | --- |
| Wrong/stale value looks like a date | Assert formula bar contains `=TODAY()` and separately assert displayed value. |
| Midnight during execution | Accept start/end calendar dates when they differ. |
| Locale ambiguity | Browser locale is explicit; compare only locale-produced representations. |
| Hostile workbook URL steals credentials | HTTPS plus Microsoft-owned host allowlist before login. |
| Credentials leak into Git/artifacts | Ignored config/auth files; recording disabled during authentication. |
| Slow Microsoft UI/network | Bounded navigation/action/workbook waits; diagnostics retained on failure. |
| Shared workbook races | One CI worker and workflow concurrency; dedicated workbook required. |
| MFA/CAPTCHA/account challenge | Fail fast with an actionable test-account/session-state message. |
| Malformed configuration | Parameterized negative unit tests. |
| Endpoint outage/performance regression | API availability check with configurable response budget. |

## Deliberate limitations

- Cross-repository or developer-to-CI concurrency cannot be locked through Playwright; use a dedicated workbook/account per execution lane.
- Microsoft can change its internal Excel DOM. Semantic ARIA selectors are preferred, with narrowly scoped fallbacks for the canvas-based editor.
- Automated username/password login is incompatible with many MFA and Conditional Access policies. A purpose-built test account or pre-created storage state is required.
- This validates Excel Online end to end. Microsoft Graph or Office Scripts would be faster alternatives, but they would not satisfy the browser/UI requirement.
