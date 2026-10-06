import { expect, type Locator, type Page } from '@playwright/test';

/**
 * Browser-contract helpers for the demo host and the `<xlsx-editor>` element. Selectors use the
 * element's shadow parts (`[part="grid"]`, `[part="formula-bar"]`, ...), which Playwright's CSS
 * engine reaches through the open shadow root.
 */
export const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;
export const SAMPLE_SHEETS = ['Sales', 'Budget', 'Summary'] as const;

export const editor = (page: Page) => page.locator('xlsx-editor');
export const part = (page: Page, name: string): Locator => editor(page).locator(`[part="${name}"]`);
export const grid = (page: Page) => part(page, 'grid');
export const nameBox = (page: Page) => part(page, 'name-box');
export const formulaBar = (page: Page) => part(page, 'formula-bar');
export const sheetTabs = (page: Page) => part(page, 'sheet-tabs');
export const ribbon = (page: Page) => part(page, 'ribbon');
export const statusBar = (page: Page) => part(page, 'status-bar');

const landingVisible = (page: Page) => page.locator('#landing').isVisible();
/** Waits until the demo script has wired the landing page's buttons. */
export const demoReady = (page: Page) =>
	page.locator('html[data-demo-ready="true"]').waitFor({ state: 'attached' });

/** Collects uncaught page errors; assert it stays empty at the end of a test. */
export function pageErrors(page: Page): string[] {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	return errors;
}

/** Opens the demo landing page in `framework`, optionally with extra query parameters. */
export async function openLanding(page: Page, framework = 'vanilla', query = '') {
	page.on('dialog', (dialog) => void dialog.accept());
	await page.goto(`/?framework=${framework}${query ? `&${query}` : ''}`);
	await demoReady(page);
}

/** Opens the demo in `framework` and loads the sample workbook from the landing page. */
export async function openSample(page: Page, framework = 'vanilla') {
	await openLanding(page, framework);
	await page.locator('#sample').click();
	await expect(grid(page)).toBeVisible();
	await expect(sheetTabs(page)).toContainText(SAMPLE_SHEETS[0]);
}

/** Starts a blank workbook from the landing page. */
export async function newWorkbook(page: Page, framework = 'vanilla') {
	if (!(await landingVisible(page))) await openLanding(page, framework);
	await page.locator('#blank').click();
	await expect(grid(page)).toBeVisible();
}

/** The landing page's file input (accepts .xlsx, .xlsm, .xls and .csv). */
export async function landingFileInput(page: Page) {
	await demoReady(page);
	return page.locator('#landing-file');
}

/** Selects a cell through the name box, as a user would type `B3` and press Enter. */
export async function goToCell(page: Page, ref: string) {
	await nameBox(page).click();
	await page.keyboard.press('Control+A');
	await page.keyboard.type(ref);
	await page.keyboard.press('Enter');
	await expect.poll(() => nameBoxValue(page)).toBe(ref.toUpperCase());
}

/** Types into the active cell of the focused grid and commits with Enter. */
export async function typeInActiveCell(page: Page, text: string) {
	await grid(page).focus();
	await page.keyboard.type(text);
	await page.keyboard.press('Enter');
}

/** The value the name box reports (it may be an input or plain text). */
export async function nameBoxValue(page: Page): Promise<string> {
	return nameBox(page).evaluate((node) => {
		const input = node instanceof HTMLInputElement ? node : node.querySelector('input');
		return (input?.value ?? node.textContent ?? '').trim();
	});
}

/** The formula bar's current text (its formula textarea; the name box sits in the same bar). */
export async function formulaBarValue(page: Page): Promise<string> {
	return formulaBar(page).evaluate((node) => {
		const field = node.querySelector<HTMLTextAreaElement>('textarea');
		return (field?.value ?? '').trim();
	});
}

/** Reads one property of the mounted `<xlsx-editor>` (fileName, dirty, locale, ...). */
export async function editorProperty<T>(page: Page, name: string): Promise<T> {
	const value = await editor(page).evaluate(
		(node, key) => (node as unknown as Record<string, unknown>)[key],
		name,
	);
	return value as T;
}
