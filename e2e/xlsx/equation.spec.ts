import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { loadXlsx } from 'ooxml-core/xlsx';
import { editor, FRAMEWORKS, grid, newWorkbook, pageErrors } from './helpers';

for (const framework of FRAMEWORKS)
	test(`${framework}: insert, edit and export an Office Math text box`, async ({ page }) => {
		await page.setViewportSize({ width: 1700, height: 900 });
		const errors = pageErrors(page);
		await newWorkbook(page, framework);
		const host = editor(page);
		await host.getByRole('tab', { name: 'Insert', exact: true }).click();
		async function clickEquation(name: string) {
			const symbols = host.getByRole('button', { name: 'Symbols', exact: true });
			const button = host.getByRole('button', { name, exact: true });
			if (!(await button.isVisible()) && (await symbols.isVisible())) await symbols.click();
			await button.click();
		}
		await clickEquation('Equation');
		const dialog = host.locator('[data-dialog="equation"]');
		const source = dialog.getByRole('textbox', { name: 'LaTeX equation', exact: true });
		await source.fill('\\frac{a');
		await expect(dialog.getByRole('button', { name: 'OK' })).toBeDisabled();
		await source.fill('\\frac{a}{b}');
		await expect(dialog.locator('mfrac')).toBeVisible();
		await source.press('Control+Enter');
		await expect(grid(page).locator('.xg-shape mfrac')).toBeVisible();
		await clickEquation('Edit Equation');
		await expect(source).toHaveValue(/frac/);
		await source.fill('\\sqrt{2}');
		await source.press('Control+Enter');
		await expect(grid(page).locator('.xg-shape msqrt')).toBeVisible();
		const save = page.waitForEvent('download');
		await host
			.locator('office-ui-title-bar')
			.getByRole('button', { name: 'Save', exact: true })
			.click();
		const path = await (await save).path();
		const book = await loadXlsx(await readFile(path!));
		const drawing = book.sheets[0]?.drawings[0];
		expect(
			drawing?.kind === 'unsupported' && drawing.shape?.equationParagraphs?.[0]?.[0],
		).toMatchObject({
			kind: 'equation',
			fallback: '\\sqrt{2}',
			omml: expect.stringContaining('<m:rad>'),
		});
		expect(errors).toEqual([]);
	});
