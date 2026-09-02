import { expect, type Locator, type Page } from '@playwright/test';

export class ExcelWorkbookPage {
  public constructor(
    private readonly page: Page,
    private readonly workbookTimeoutMs: number,
  ) {}

  public async open(url: string): Promise<void> {
    await this.page.goto(url, { waitUntil: 'domcontentloaded' });
    await this.visibleLocator(
      this.nameBoxCandidates(),
      this.workbookTimeoutMs,
      'Excel workbook did not become editable. Confirm the URL and edit permission.',
    );
  }

  public async enterFormula(cellReference: string, formula: string): Promise<void> {
    await this.selectCell(cellReference);
    await this.page.keyboard.type(formula);
    await this.page.keyboard.press('Enter');
    await this.selectCell(cellReference);
  }

  public async readSelectedCellValue(): Promise<string> {
    await this.page.keyboard.press('ControlOrMeta+C');

    const value = await expect
      .poll(
        async () =>
          this.page.evaluate(async () => {
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
        this.page.evaluate(async () => (await navigator.clipboard.readText()).trim()),
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
    await this.page.keyboard.press('Delete');
  }

  private async selectCell(cellReference: string): Promise<void> {
    const nameBox = await this.visibleLocator(this.nameBoxCandidates());
    await nameBox.click();
    await nameBox.fill(cellReference);
    await nameBox.press('Enter');
    await expect(nameBox).toHaveValue(new RegExp(`^${escapeRegex(cellReference)}$`, 'i'));
  }

  private nameBoxCandidates(): readonly Locator[] {
    return [
      this.page.getByRole('combobox', { name: /name box/i }).first(),
      this.page.locator('input[aria-label*="Name Box" i]').first(),
      this.page.locator('[data-unique-id="formula-bar-name-box"] input').first(),
      this.page.locator('input[id*="nameBox" i]').first(),
    ];
  }

  private formulaBarCandidates(): readonly Locator[] {
    return [
      this.page.getByRole('textbox', { name: /formula bar/i }).first(),
      this.page.locator('[contenteditable="true"][aria-label*="formula bar" i]').first(),
      this.page.locator('textarea[aria-label*="formula bar" i]').first(),
      this.page.locator('[id*="formulaBarTextDivId"] [contenteditable="true"]').first(),
      this.page.locator('[id*="formulaBarTextDivId"]').first(),
    ];
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
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
