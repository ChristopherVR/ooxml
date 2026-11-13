import { expect, test, type Page } from '@playwright/test';
import { createWorkbook, saveXlsx, type SheetProtection } from 'ooxml-core/xlsx';
import { FRAMEWORKS, demoReady, editor } from './helpers';
import type { NativeHarness } from '../../demos/xlsx/demo-vanilla/native-building-blocks';
import type { HostUiState } from '../../demos/xlsx/demo-vanilla/host-workflows';

declare global {
	interface Window {
		xlsxNativeBlocks: NativeHarness;
		xlsxHostUi: HostUiState;
	}
}

/** apple / apple pie / pear in A1:A3, 5 in B1; a second sheet that is protected. */
async function seedBytes(): Promise<number[]> {
	const workbook = createWorkbook({ sheets: ['Data', 'Locked'] });
	const data = workbook.sheets[0]!;
	data.rows.set(
		0,
		new Map([
			[0, { value: 'apple' }],
			[1, { value: 5 }],
		]),
	);
	data.rows.set(1, new Map([[0, { value: 'apple pie' }]]));
	data.rows.set(2, new Map([[0, { value: 'pear' }]]));
	// B1 only accepts whole numbers above 10, so replacing its 5 with 7 breaks the rule.
	data.dataValidations.push({
		ranges: [{ start: { row: 0, col: 1 }, end: { row: 0, col: 1 } }],
		type: 'whole',
		operator: 'greaterThan',
		formula1: '10',
		showErrorMessage: true,
		error: 'Too small',
	});
	const locked = workbook.sheets[1]!;
	locked.rows.set(0, new Map([[0, { value: 'secret apple' }]]));
	locked.protection = { sheet: true } as SheetProtection;
	return [...(await saveXlsx(workbook))];
}

async function open(page: Page, framework: string): Promise<void> {
	await page.goto(`/?framework=${framework}&blocks=1&hostui=1`);
	await demoReady(page);
	await page.getByRole('button', { name: 'or create a New Workbook', exact: true }).click();
	await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
	const bytes = await seedBytes();
	await page.evaluate((data) => window.xlsxNativeBlocks.handle.load(new Uint8Array(data)), bytes);
	await expect(page.getByRole('tab', { name: 'Locked' })).toBeVisible();
	// The loaded sheet is painted before the first click lands on it.
	await expect(editor(page).getByRole('grid', { name: 'Spreadsheet grid' })).toContainText(
		'apple pie',
	);
}

const log = (page: Page) => page.evaluate(() => window.xlsxHostUi.log);
const gridBox = async (page: Page) => {
	const grid = editor(page).getByRole('grid', { name: 'Spreadsheet grid' });
	// Some demos stack a tall toolbar above the canvas: bring the grid into the viewport first.
	await grid.scrollIntoViewIfNeeded();
	// Let the scroll event of that call land now: one delivered after a click would be mistaken
	// for the user scrolling, which dismisses an open menu by design.
	await grid
		.page()
		.evaluate(
			() => new Promise((done) => requestAnimationFrame(() => requestAnimationFrame(done))),
		);
	return (await grid.boundingBox())!;
};
/** Clicks at an offset inside the grid, measured when the click happens (the page may scroll). */
const clickGrid = async (page: Page, at: { x: number; y: number }, button: 'left' | 'right') => {
	const box = await gridBox(page);
	await page.mouse.click(box.x + at.x, box.y + at.y, { button });
};
const openNative = async (page: Page) => {
	await clickGrid(page, { x: 160, y: 80 }, 'right');
	await expect(editor(page).locator('office-ui-context-menu')).toHaveCount(1);
};
const cellValue = (page: Page, ref: string) =>
	page.evaluate((r) => {
		const workbook = window.xlsxNativeBlocks.handle.element.workbook!;
		const match = /^([A-Z]+)(\d+)$/.exec(r)!;
		const col = match[1]!.charCodeAt(0) - 65;
		return workbook.sheets[0]!.rows.get(Number(match[2]) - 1)?.get(col)?.value ?? null;
	}, ref);
const dirtyUndo = (page: Page) =>
	page.evaluate(() => window.xlsxNativeBlocks.handle.element.undo());

for (const framework of FRAMEWORKS) {
	test.describe(framework, () => {
		test('the host menu replaces the native menu on cell, headers and tabs', async ({ page }) => {
			const errors: string[] = [];
			page.on('pageerror', (e) => errors.push(e.message));
			await open(page, framework);
			const hostMenu = page.locator('#host-menu');
			const native = editor(page).locator('office-ui-context-menu');
			const spots = {
				cell: { x: 160, y: 80 },
				row: { x: 10, y: 80 },
				column: { x: 160, y: 10 },
			} as const;
			for (const [kind, at] of Object.entries(spots)) {
				await clickGrid(page, at, 'right');
				await expect(hostMenu).toHaveAttribute('data-kind', kind);
				await expect(native).toHaveCount(0);
				await page.keyboard.press('Escape');
				await clickGrid(page, { x: 400, y: 300 }, 'left');
			}
			await page.getByRole('tab', { name: 'Locked' }).click({ button: 'right' });
			await expect(hostMenu).toHaveAttribute('data-kind', 'tab');
			await expect(native).toHaveCount(0);
			await page.keyboard.press('Escape');

			// Executing an item works: Format Cells reaches the host callback.
			await clickGrid(page, { x: 160, y: 80 }, 'right');
			await hostMenu.locator('[data-id="format-cells"]').click();
			expect(await log(page)).toEqual(
				expect.arrayContaining(['open-format-cells', 'cell:format-cells:true']),
			);
			// A disabled item is refused by execute().
			await clickGrid(page, { x: 160, y: 80 }, 'right');
			await hostMenu.locator('[aria-disabled="true"]').first().click({ force: true });
			expect((await log(page)).at(-1)).toMatch(/:false$/);
			expect(errors).toEqual([]);
		});

		test('the native menu can be filtered and extended by the host', async ({ page }) => {
			await open(page, framework);
			await page.evaluate(() => (window.xlsxHostUi.mode = 'native'));
			const native = editor(page).locator('office-ui-context-menu');
			await openNative(page);
			await expect(page.locator('#host-menu')).toHaveCount(0);
			const labels = native.getByRole('menuitem');
			await expect(labels.filter({ hasText: /^Host stamp$/ })).toHaveCount(1);
			await expect(labels.filter({ hasText: /^Copy$/ })).toHaveCount(0);
			await labels.filter({ hasText: /^Host stamp$/ }).click();
			expect(await log(page)).toContain('stamp:cell');
			// The edit-only host item is refused once the workbook is read-only.
			await page.evaluate(() => window.xlsxNativeBlocks.update({ canEdit: false }));
			await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeDisabled();
			await openNative(page);
			await expect(native.getByRole('menuitem', { name: 'Host stamp (edit)' })).toBeDisabled();
			await native.getByRole('menuitem', { name: 'Host stamp (edit)' }).click({ force: true });
			expect(await log(page)).not.toContain('stamp-edit:cell');
		});

		test('Format Cells shows mixed values and applies as one undo step', async ({ page }) => {
			await open(page, framework);
			await page.evaluate(() => window.xlsxNativeBlocks.handle.select('A1'));
			await page.evaluate(() => window.xlsxNativeBlocks.snapshot().toolbarProps!.onBold());
			await page.evaluate(() => window.xlsxNativeBlocks.handle.select('A1:A3'));
			await page.getByRole('button', { name: 'Host load format' }).click();
			const status = page.getByLabel('Host format status');
			await expect(status).toContainText('mixed=bold');
			await expect(page.getByLabel('Host bold')).toHaveJSProperty('indeterminate', true);
			await page.getByLabel('Host bold').check();
			await page.getByLabel('Host number format').fill('0.00');
			await page.getByRole('button', { name: 'Host apply format' }).click();
			await expect(status).toHaveText('applied=true');
			await page.getByRole('button', { name: 'Host load format' }).click();
			await expect(status).not.toContainText('mixed=bold');
			await dirtyUndo(page);
			// One undo step restores both the weight and the number format.
			await page.getByRole('button', { name: 'Host load format' }).click();
			await expect(status).toContainText('mixed=bold');
			await expect(status).toContainText('numFmt=General');
		});

		test('a protected sheet refuses formatting', async ({ page }) => {
			await open(page, framework);
			await page.getByRole('tab', { name: 'Locked' }).click();
			await page.evaluate(() => window.xlsxNativeBlocks.handle.select('A1'));
			await page.getByRole('button', { name: 'Host apply format' }).click();
			await expect(page.getByLabel('Host format status')).toHaveText('applied=false');
		});

		test('find works read-only, replace is refused, then replace and undo', async ({ page }) => {
			await open(page, framework);
			await page.getByLabel('Host find what').fill('apple');
			await page.getByLabel('Host replace with').fill('plum');
			await page.evaluate(() => window.xlsxNativeBlocks.update({ canEdit: false }));
			const status = page.getByLabel('Host find status');
			await page.getByRole('button', { name: 'Host find all' }).click();
			await expect(status).toContainText('3 found');
			await expect(status).toContainText('canReplace=false');
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(status).toContainText('replaced 0');
			expect(await cellValue(page, 'A1')).toBe('apple');
			await page.evaluate(() => window.xlsxNativeBlocks.update({ canEdit: true }));
			await page.getByRole('button', { name: 'Host replace all' }).click();
			// The protected sheet's locked cell is skipped, like Excel.
			await expect(status).toContainText('replaced 2');
			expect(await cellValue(page, 'A1')).toBe('plum');
			expect(await cellValue(page, 'A2')).toBe('plum pie');
			await page.evaluate(() => window.xlsxNativeBlocks.handle.element.undo());
			expect(await cellValue(page, 'A1')).toBe('apple');
			expect(await cellValue(page, 'A2')).toBe('apple pie');
			await page.evaluate(() => window.xlsxNativeBlocks.handle.select('A1'));
			await page.getByLabel('Host find what').fill('apple pie');
			await page.getByLabel('Host replace with').fill('fig');
			await page.getByRole('button', { name: 'Host find next' }).click();
			await page.getByRole('button', { name: 'Host replace', exact: true }).click();
			await expect(status).toContainText('replaced true');
			expect(await cellValue(page, 'A2')).toBe('fig');
			await page.evaluate(() => window.xlsxNativeBlocks.handle.element.undo());
			expect(await cellValue(page, 'A2')).toBe('apple pie');
		});

		test('replace ignores data validation by default, like Excel', async ({ page }) => {
			await open(page, framework);
			await page.getByLabel('Host find what').fill('5');
			await page.getByLabel('Host replace with').fill('7');
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(page.getByLabel('Host find status')).toContainText('replaced 1; refused=;');
			expect(await cellValue(page, 'B1')).toBe(7);
		});

		test('a call can ask replace to validate and reports the refused cells', async ({ page }) => {
			await open(page, framework);
			await page.getByLabel('Host find what').fill('5');
			await page.getByLabel('Host replace with').fill('7');
			await page.getByLabel('Host validate replace').check();
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(page.getByLabel('Host find status')).toContainText('replaced 0; refused=B1;');
			expect(await cellValue(page, 'B1')).toBe(5);
			// A value the rule accepts goes through.
			await page.getByLabel('Host replace with').fill('70');
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(page.getByLabel('Host find status')).toContainText('replaced 1; refused=;');
			expect(await cellValue(page, 'B1')).toBe(70);
		});

		test('the validateReplace option applies to every call and a call can opt out', async ({
			page,
		}) => {
			await open(page, framework);
			await page.evaluate(() =>
				window.xlsxNativeBlocks.update({ canEdit: true, validateReplace: true }),
			);
			await expect
				.poll(() => page.evaluate(() => window.xlsxNativeBlocks.handle.element.validateReplace))
				.toBe(true);
			await page.getByLabel('Host find what').fill('5');
			await page.getByLabel('Host replace with').fill('7');
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(page.getByLabel('Host find status')).toContainText('replaced 0; refused=B1;');
			expect(await cellValue(page, 'B1')).toBe(5);
			await page.getByLabel('Host skip validation').check();
			await page.getByRole('button', { name: 'Host replace all' }).click();
			await expect(page.getByLabel('Host find status')).toContainText('replaced 1; refused=;');
			expect(await cellValue(page, 'B1')).toBe(7);
		});

		test('Insert Function inserts at the caret of an active formula draft', async ({ page }) => {
			await open(page, framework);
			await page.evaluate(() => window.xlsxNativeBlocks.handle.select('C1'));
			const formula = page.getByRole('textbox', { name: 'Formula Bar', exact: true });
			await formula.fill('=1+');
			await page.getByLabel('Host function search').fill('SUM');
			const insert = page.locator('[data-function="SUM"]');
			await expect(insert).toBeVisible();
			await insert.click();
			await expect(page.getByLabel('Host function status')).toHaveText('=1+SUM(|7');
			await expect(formula).toHaveValue('=1+SUM(');
			await formula.press('Escape');
		});
	});
}
