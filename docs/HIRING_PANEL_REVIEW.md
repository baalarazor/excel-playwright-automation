# Hiring-panel review

## Verdict

Ready for review after one environment-dependent proof: a successful run against the reviewer's dedicated Microsoft account and workbook. The implementation is complete, but no honest E2E pass or demo video can be claimed without those external inputs.

## What is strong

- The acceptance scenario reads as business behavior and contains no login or selector detail.
- Configuration, date-domain logic, login workflow, and workbook UI are independently owned components.
- Stable Chrome is explicit (`channel: chrome`), not implied by Playwright's bundled Chromium.
- Assertions prove both inputs and outputs: exact formula plus calculated current date, read through Excel's accessibility surface with a clipboard fallback.
- Locale, time zone, midnight crossover, malformed values, Excel grid limits, URL validation, response time, cleanup, and diagnostic evidence are addressed.
- Credentials and authenticated browser state are excluded from Git and auth is excluded from recordings/traces.
- CI serialization acknowledges that a workbook cell is shared mutable state.
- Dependencies are current and the audit reports no known vulnerabilities at delivery time.

## What I would challenge in interview

1. **Why UI instead of Graph API?** Because the requirement asks for Excel Online in Chrome. The API smoke test helps triage outages; Graph or Office Scripts would be alternatives, not an E2E substitute.
2. **Why not add retries everywhere?** Playwright locators already auto-wait. Bounded waits and one CI retry cover transient infrastructure without hiding deterministic defects.
3. **Why Page Objects?** Microsoft login and Excel are separate change boundaries. The test remains readable while selectors remain localized. A generic base page was avoided as premature abstraction.
4. **How is midnight handled?** The test records instants immediately before formula execution and after value retrieval, then accepts both civil dates only if the configured time zone crossed midnight.
5. **Can it run concurrently?** Not safely against the same A2. Local UI scope contains one test; CI uses one worker and a workflow concurrency group. Scale-out requires one workbook/account per lane.

## Residual risks and next investments

- Run locator discovery against the actual tenant because Microsoft rolls out Excel UI variants gradually.
- If the suite grows, provision test workbooks per worker through Microsoft Graph, then delete them in teardown.
- For organizations enforcing MFA, bootstrap short-lived storage state through an approved interactive flow or workload identity; do not weaken tenant controls.
- Add a second regional configuration only if localized Excel behavior is in product scope; one configurable locale is sufficient for this assignment.
