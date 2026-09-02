import { expect, type Locator, type Page } from '@playwright/test';

export interface MicrosoftCredentials {
  readonly password: string;
  readonly username: string;
}

export class MicrosoftLoginPage {
  public constructor(private readonly page: Page) {}

  public async signIn(
    credentials: MicrosoftCredentials,
    staySignedIn: boolean,
  ): Promise<void> {
    const hostname = new URL(this.page.url()).hostname.toLowerCase();
    if (!['login.live.com', 'login.microsoft.com', 'login.microsoftonline.com'].includes(hostname)) {
      throw new Error(`Refusing to enter Microsoft credentials on untrusted host "${hostname}".`);
    }

    const email = this.page.locator('input[type="email"], input[name="loginfmt"]').first();
    if (!(await email.isVisible({ timeout: 2_000 }).catch(() => false))) {
      const matchingAccount = this.page.getByText(credentials.username, { exact: true }).first();
      const otherAccount = this.page
        .locator('#otherTileText')
        .or(this.page.getByText(/use another account/i))
        .first();
      if (await matchingAccount.isVisible({ timeout: 1_000 }).catch(() => false)) {
        await matchingAccount.click();
      } else if (await otherAccount.isVisible({ timeout: 1_000 }).catch(() => false)) {
        await otherAccount.click();
      }
    }

    if (await email.isVisible({ timeout: 3_000 }).catch(() => false)) {
      await email.fill(credentials.username);
      await this.submit();
    }

    const password = this.page.locator('input[type="password"], input[name="passwd"]').first();
    await expect(password, 'Microsoft password field was not displayed').toBeVisible();
    await password.fill(credentials.password);
    await this.submit();

    await this.failFastOnCredentialError();
    await this.failFastOnInteractiveChallenge();
    await this.answerStaySignedInPrompt(staySignedIn);
  }

  private async submit(): Promise<void> {
    const submit = this.page.locator('#idSIButton9, button[type="submit"], input[type="submit"]').first();
    await expect(submit, 'Microsoft sign-in submit control was not displayed').toBeVisible();
    await submit.click();
  }

  private async answerStaySignedInPrompt(staySignedIn: boolean): Promise<void> {
    const prompt = this.page.getByText(/stay signed in/i).first();
    if (!(await prompt.isVisible({ timeout: 3_000 }).catch(() => false))) return;

    const selector = staySignedIn ? '#idSIButton9' : '#idBtn_Back';
    const button = this.page.locator(selector).first();
    await expect(button).toBeVisible();
    await button.click();
  }

  private async failFastOnInteractiveChallenge(): Promise<void> {
    const challenge: Locator = this.page.getByText(
      /approve sign.?in request|enter code|verify your identity|more information required/i,
    );
    if (await challenge.first().isVisible({ timeout: 2_000 }).catch(() => false)) {
      throw new Error(
        'Microsoft requested MFA or an interactive account challenge. Use a test account exempted from interactive MFA, or create playwright/.auth/user.json manually with npm run test:ui.',
      );
    }
  }

  private async failFastOnCredentialError(): Promise<void> {
    const error = this.page
      .locator('#passwordError, #usernameError')
      .or(this.page.getByText(/incorrect password|account.*doesn.t exist|couldn.t sign you in/i))
      .first();
    if (await error.isVisible({ timeout: 2_000 }).catch(() => false)) {
      throw new Error(`Microsoft sign-in failed: ${(await error.textContent())?.trim() ?? 'unknown error'}`);
    }
  }
}
