import { expect, test } from '@playwright/test';
import type { Workbook } from 'ooxml-core/xlsx';
import {
	clickRibbonControl,
	editor,
	FRAMEWORKS,
	goToCell,
	grid,
	newWorkbook,
	openLanding,
	pageErrors,
	ribbon,
	typeInActiveCell,
} from './helpers';

interface FormulaEditor extends HTMLElement {
	workbook: Workbook;
	saveBytes(): Promise<Uint8Array>;
	load(bytes: Uint8Array, name: string): Promise<void>;
}

test('the first FILTERXML formatting rule repaints after XPath loads', async ({ page }) => {
	const errors = pageErrors(page);
	await newWorkbook(page);
	await goToCell(page, 'A1');
	await typeInActiveCell(page, '<r><a>7</a></r>');
	await goToCell(page, 'B1');
	await typeInActiveCell(page, '9');
	await goToCell(page, 'B1');
	const cell = grid(page).locator('.xg-c').filter({ hasText: /^9$/ });
	await expect(cell).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');

	await clickRibbonControl(
		page,
		ribbon(page).getByRole('button', {
			name: 'Conditional Formatting',
			exact: true,
			includeHidden: true,
		}),
	);
	await editor(page).getByRole('menuitem', { name: 'New Rule...', exact: true }).click();
	const rule = editor(page).locator('[data-dialog="cf-rule"]');
	await rule
		.getByRole('option', { name: 'Use a formula to determine which cells to format' })
		.click();
	await rule
		.getByLabel('Format values where this formula is true:')
		.fill('=FILTERXML($A$1,"//a")=7');
	await rule.getByRole('button', { name: 'Format...', exact: true }).click();
	const format = editor(page).locator('[data-dialog="format-cells"]');
	await format.getByRole('tab', { name: 'Fill', exact: true }).click();
	await format
		.locator('office-ui-color-grid[aria-label="Background Color:"] [aria-label="Yellow"]')
		.click();
	await format.getByRole('button', { name: 'OK', exact: true }).click();

	const xpathAsset = /\/assets\/xpath-[^/]+\.js$/;
	let releaseXPath: () => void = () => {};
	const blocked = new Promise<void>((resolve) => {
		releaseXPath = resolve;
	});
	await page.route(xpathAsset, async (route) => {
		await blocked;
		await route.continue();
	});
	try {
		await Promise.all([
			page.waitForRequest(xpathAsset),
			rule.getByRole('button', { name: 'OK', exact: true }).click(),
		]);
		await expect(rule).toBeHidden();
		await expect(cell).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
	} finally {
		releaseXPath();
	}
	await expect(cell).toHaveCSS('background-color', 'rgb(255, 255, 0)');
	await page.keyboard.press('Control+Z');
	await expect(cell).toHaveCSS('background-color', 'rgba(0, 0, 0, 0)');
	expect(errors).toEqual([]);
});

for (const framework of FRAMEWORKS) {
	test(`${framework}: FILTERXML spills, recalculates and survives save`, async ({ page }) => {
		const errors = pageErrors(page);
		await openLanding(page, framework);
		await newWorkbook(page, framework);
		await goToCell(page, 'A1');
		await typeInActiveCell(page, '<r><a>001</a><a>TRUE</a><a>tail</a></r>');
		await goToCell(page, 'C1');
		await typeInActiveCell(page, '=FILTERXML(A1,"//a")');
		const values = () =>
			editor(page).evaluate((node) => {
				const sheet = (node as FormulaEditor).workbook.sheets[0]!;
				return [0, 1, 2].map((row) => sheet.rows.get(row)?.get(2)?.value ?? null);
			});
		await expect.poll(values).toEqual([1, true, 'tail']);
		await goToCell(page, 'A1');
		await typeInActiveCell(page, '<r><a>changed</a></r>');
		await expect.poll(values).toEqual(['changed', null, null]);
		await page.keyboard.press('Control+z');
		await expect.poll(values).toEqual([1, true, 'tail']);
		await editor(page).evaluate(async (node) => {
			const el = node as FormulaEditor;
			await el.load(await el.saveBytes(), 'filterxml.xlsx');
		});
		await expect.poll(values).toEqual([1, true, 'tail']);
		expect(errors).toEqual([]);
	});
}
