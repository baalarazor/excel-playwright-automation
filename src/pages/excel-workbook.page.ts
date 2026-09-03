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
    return this.waitForSelectedCellValue(
      cellReference,
      () => true,
      `Excel did not expose ${cellReference}'s displayed value through accessibility.`,
    );
  }

  public async waitForSelectedCellValue(
    cellReference: string,
    matches: (value: string) => boolean,
    message: string,
  ): Promise<string> {
    let matchedValue = '';
    await expect
      .poll(
        async () => {
          matchedValue = (await this.accessibleCellValues(cellReference)).find(matches) ?? '';
          return matchedValue;
        },
        { message, timeout: 10_000 },
      )
      .not.toBe('');

    return matchedValue;
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
    await this.selectCell(cellReference);
    await expect.poll(() => this.readSelectedCellFormula(), {
      message: `Excel did not clear ${cellReference}`,
    }).toBe('');
  }

  public async setNumberFormat(cellReference: string, formatName: string): Promise<void> {
    await this.selectCell(cellReference);
    await this.openNumberFormatMenu();

    const format = this.scope.getByRole('option', { name: new RegExp(`^${escapeRegex(formatName)}$`, 'i') }).first();
    await expect(format, `Excel did not offer the ${formatName} number format`).toBeVisible();
    await format.click();
    await this.selectCell(cellReference);
  }

  public async setCustomNumberFormat(cellReference: string, formatCode: string): Promise<void> {
    await this.selectCell(cellReference);
    await this.openMoreNumberFormats();

    const category = this.scope.getByRole('combobox', { name: /choose category/i }).first();
    await category.click();
    const custom = this.scope.getByRole('option', { name: /^custom$/i }).first();
    await expect(custom, 'Excel did not offer the Custom number-format category').toBeVisible();
    await custom.click();

    const formatInput = await this.visibleLocator([
      this.scope.locator('#customFormatInput').first(),
    ]);
    await formatInput.fill(formatCode);
    await this.scope.getByRole('button', { name: /^confirm$/i }).click();
    await this.selectCell(cellReference);
  }

  private async selectCell(cellReference: string): Promise<void> {
    await this.dismissBlockingDialog();
    await this.page.keyboard.press('Escape');
    const nameBox = await this.visibleLocator(this.nameBoxCandidates());
    await nameBox.click();
    await nameBox.fill(cellReference);
    await nameBox.press('Enter');
    await expect(nameBox, `Excel did not select ${cellReference}`).toHaveValue(
      new RegExp(`^${escapeRegex(cellReference)}$`, 'i'),
    );
  }

  private async openNumberFormatMenu(): Promise<void> {
    const numberFormat = await this.visibleLocator([
      this.scope.getByRole('button', { name: /number format/i }).first(),
      this.scope.locator('button[aria-label*="number format" i]').first(),
    ]);
    await numberFormat.click();
  }

  private async openMoreNumberFormats(): Promise<void> {
    await this.openNumberFormatMenu();
    const moreFormats = this.scope
      .getByRole('option', { name: /^more number formats\.\.\.$/i })
      .first();
    await expect(moreFormats, 'Excel did not offer More Number Formats').toBeVisible();
    await moreFormats.click();
  }

  private async dismissBlockingDialog(): Promise<void> {
    const confirmation = this.scope.getByRole('button', { name: /^ok$/i }).first();
    if (await confirmation.isVisible().catch(() => false)) await confirmation.click();
  }

  private async accessibleCellValues(cellReference: string): Promise<string[]> {
    const cellTextboxes = this.scope.getByRole('textbox', {
      name: new RegExp(`(?:^|\\s\\.\\s)${escapeRegex(cellReference)}(?:\\s\\.\\s|$)`, 'i'),
    });
    const count = await cellTextboxes.count();
    const values: string[] = [];
    for (let index = 0; index < count; index += 1) {
      const label = await cellTextboxes.nth(index).getAttribute('aria-label');
      const parsed = label ? extractCellValueFromAccessibleLabel(label, cellReference) : '';
      if (parsed) values.push(parsed);
    }
    return values;
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
    return this.findLocator(candidates, 'visible', timeout, `${errorMessage} URL: ${this.page.url()}`);
  }

  private async attachedLocator(candidates: readonly Locator[]): Promise<Locator> {
    return this.findLocator(
      candidates,
      'attached',
      10_000,
      `Excel keyboard editor was not attached. URL: ${this.page.url()}`,
    );
  }

  private async findLocator(
    candidates: readonly Locator[],
    state: 'attached' | 'visible',
    timeout: number,
    errorMessage: string,
  ): Promise<Locator> {
    try {
      return await Promise.any(
        candidates.map(async (candidate) => {
          await candidate.waitFor({ state, timeout });
          return candidate;
        }),
      );
    } catch {
      throw new Error(errorMessage);
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
