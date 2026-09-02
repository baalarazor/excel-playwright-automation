import { expect, type Frame, type Locator, type Page } from '@playwright/test';

type WorkbookScope = Page | Frame;

export class ExcelWorkbookPage {
  private scope: WorkbookScope;

  public constructor(
    private readonly page: Page,
    private readonly workbookTimeoutMs: number,
  ) {
    this.scope = page;
  }

  public async open(url: string): Promise<void> {
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });
    this.scope = await this.findWorkbookScope();
  }

  public async enterFormula(cellReference: string, formula: string): Promise<void> {
    await this.selectCell(cellReference);
    const grid = await this.attachedLocator(this.gridCandidates());
    await grid.focus();
    await this.page.keyboard.type(formula, { delay: 50 });
    await this.page.keyboard.press('Enter');
    await this.selectCell(cellReference);
  }

  public async readSelectedCellValue(cellReference: string): Promise<string> {
    const accessibleValue = await this.readAccessibleCellValue(cellReference);
    if (accessibleValue) return accessibleValue;

    await this.page.keyboard.press('ControlOrMeta+C');

    const value = await expect
      .poll(
        async () =>
          this.scope.evaluate(async () => {
            try {
              return (await navigator.clipboard.readText()).trim();
            } catch {
              return '';
            }
          }),
        {
          message: 'Excel did not expose the selected cell value through the clipboard',
          timeout: 10_000,
        },
      )
      .not.toBe('')
      .then(async () =>
        this.scope.evaluate(async () => (await navigator.clipboard.readText()).trim()),
      );

    return value;
  }

  public async readSelectedCellFormula(): Promise<string> {
    const formulaBar = await this.visibleLocator(this.formulaBarCandidates());
    await expect(formulaBar, 'Excel formula bar was not available').toBeVisible();

    return formulaBar.evaluate((element) => {
      if (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) {
        return element.value.trim();
      }
      return (element.textContent ?? '').trim();
    });
  }

  public async clearCell(cellReference: string): Promise<void> {
    await this.selectCell(cellReference);
    const grid = await this.attachedLocator(this.gridCandidates());
    await grid.focus();
    await grid.press('Delete');
  }

  private async selectCell(cellReference: string): Promise<void> {
    await this.dismissBlockingDialog();
    const nameBox = await this.visibleLocator(this.nameBoxCandidates());
    await nameBox.click();
    await nameBox.press('ControlOrMeta+A');
    await nameBox.pressSequentially(cellReference, { delay: 50 });
    await nameBox.press('Enter');
    await expect
      .poll(
        async () => {
          const labels = await this.scope
            .locator('[role="textbox"][aria-label]')
            .evaluateAll((elements) => elements.map((element) => element.getAttribute('aria-label')));
          return labels.some(
            (label) =>
              label !== null &&
              labelContainsCellReference(label, cellReference) &&
              /\bselected\b/i.test(label),
          );
        },
        { message: `Excel did not select ${cellReference}` },
      )
      .toBe(true);
  }

  private async dismissBlockingDialog(): Promise<void> {
    const confirmation = this.scope.getByRole('button', { name: /^ok$/i }).first();
    if (await confirmation.isVisible().catch(() => false)) await confirmation.click();
  }

  private async readAccessibleCellValue(cellReference: string): Promise<string> {
    let value = '';
    await expect
      .poll(
        async () => {
          const cellTextboxes = this.scope.getByRole('textbox', {
            name: new RegExp(`(?:^|\\s\\.\\s)${escapeRegex(cellReference)}(?:\\s\\.\\s|$)`, 'i'),
          });
          const count = await cellTextboxes.count();
          for (let index = 0; index < count; index += 1) {
            const label = await cellTextboxes.nth(index).getAttribute('aria-label');
            const parsed = label ? extractCellValueFromAccessibleLabel(label, cellReference) : '';
            if (parsed) {
              value = parsed;
              return value;
            }
          }
          return '';
        },
        { timeout: 3_000 },
      )
      .not.toBe('')
      .catch(() => undefined);

    return value;
  }

  private nameBoxCandidates(): readonly Locator[] {
    return [
      this.scope.getByRole('combobox', { name: /name box/i }).first(),
      this.scope.locator('input[aria-label*="Name Box" i]').first(),
      this.scope.locator('[data-unique-id="formula-bar-name-box"] input').first(),
      this.scope.locator('input[id*="nameBox" i]').first(),
    ];
  }

  private formulaBarCandidates(): readonly Locator[] {
    return [
      this.scope.getByRole('textbox', { name: /formula bar/i }).first(),
      this.scope.locator('[contenteditable="true"][aria-label*="formula bar" i]').first(),
      this.scope.locator('textarea[aria-label*="formula bar" i]').first(),
      this.scope.locator('[id*="formulaBarTextDivId"] [contenteditable="true"]').first(),
      this.scope.locator('[id*="formulaBarTextDivId"]').first(),
    ];
  }

  private gridCandidates(): readonly Locator[] {
    return [
      this.scope.getByRole('textbox', { name: /^grid$/i }).first(),
      this.scope.locator('#gridKeyboardContentEditable_textElement').first(),
      this.scope.locator('[contenteditable="true"][aria-label="grid" i]').first(),
    ];
  }

  private async findWorkbookScope(): Promise<WorkbookScope> {
    const deadline = Date.now() + this.workbookTimeoutMs;

    while (Date.now() < deadline) {
      const scopes: WorkbookScope[] = [this.page, ...this.page.frames()];
      for (const scope of scopes) {
        this.scope = scope;
        const nameBox = this.nameBoxCandidates()[0]!;
        if (await nameBox.isVisible().catch(() => false)) return scope;
      }
      await this.page.waitForTimeout(250);
    }

    throw new Error(
      `Excel workbook did not become editable. Confirm the URL and edit permission. URL: ${this.page.url()}`,
    );
  }

  private async visibleLocator(
    candidates: readonly Locator[],
    timeout = 10_000,
    errorMessage = 'Expected Excel control was not visible.',
  ): Promise<Locator> {
    try {
      return await Promise.any(
        candidates.map(async (candidate) => {
          await candidate.waitFor({ state: 'visible', timeout });
          return candidate;
        }),
      );
    } catch {
      throw new Error(`${errorMessage} URL: ${this.page.url()}`);
    }
  }

  private async attachedLocator(candidates: readonly Locator[]): Promise<Locator> {
    try {
      return await Promise.any(
        candidates.map(async (candidate) => {
          await candidate.waitFor({ state: 'attached', timeout: 10_000 });
          return candidate;
        }),
      );
    } catch {
      throw new Error(`Excel keyboard editor was not attached. URL: ${this.page.url()}`);
    }
  }
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function extractCellValueFromAccessibleLabel(
  label: string,
  cellReference: string,
): string {
  const segments = label.split(/\s+\.\s+/).map((segment) => segment.trim());
  const cellIndex = segments.findIndex(
    (segment) => segment.toUpperCase() === cellReference.toUpperCase(),
  );
  if (cellIndex < 0) return '';

  // Excel's readout is normally "value . A2 . metadata"; blank cells start at "A2".
  return cellIndex > 0 ? (segments[cellIndex - 1] ?? '') : '';
}

function labelContainsCellReference(label: string, cellReference: string): boolean {
  return label
    .split(/\s+\.\s+/)
    .some((segment) => segment.trim().toUpperCase() === cellReference.toUpperCase());
}
