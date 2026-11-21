import { expect, test } from '@playwright/test';
import { mkdir, writeFile } from 'node:fs/promises';
import { loadXlsx } from 'ooxml-core/xlsx';
import { clickRibbonControl, editor, newWorkbook, openLanding } from './helpers';

test('inserts styled WordArt, reopens the workbook, and edits the drawing', async ({
	page,
}, testInfo) => {
	await page.setViewportSize({ width: 1600, height: 850 });
	await openLanding(page);
	await newWorkbook(page);
	const host = editor(page);
	await host.getByRole('tab', { name: 'Insert', exact: true }).click();
	await clickRibbonControl(page, host.getByRole('button', { name: 'WordArt', exact: true }));
	const picker = host.locator('office-ui-dialog[open]');
	await picker.getByRole('textbox', { name: 'WordArt text' }).fill('Sales');
	await picker.getByRole('button', { name: 'WordArt style 30' }).click();
	await picker.getByRole('combobox', { name: 'Transform' }).selectOption('textArchUp');
	await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(
		host.locator('.xg-obj [data-word-art-transform="textArchUp"] textPath').first(),
	).toHaveText('Sales');
	const save = () =>
		host.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
	const bytes = await save();
	await mkdir(testInfo.outputDir, { recursive: true });
	await writeFile(testInfo.outputPath('word-art.xlsx'), Uint8Array.from(bytes));
	await openLanding(page);
	await page.locator('#landing-file').setInputFiles({
		name: 'word-art.xlsx',
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(bytes),
	});
	await host.locator('.xg-obj[aria-label="WordArt"]', { hasText: 'Sales' }).first().dblclick();
	const edit = host.locator('office-ui-dialog[open]');
	await expect(edit.getByRole('combobox', { name: 'Transform' })).toHaveValue('textArchUp');
	await edit.getByRole('combobox', { name: 'Transform' }).selectOption('textWave1');
	await edit.getByRole('textbox', { name: 'WordArt text' }).fill('Total sales');
	await edit.getByRole('button', { name: 'Apply', exact: true }).click();
	const saved = await loadXlsx(Uint8Array.from(await save()));
	expect(saved.sheets[0]?.drawings[0]).toMatchObject({
		shape: {
			text: { text: 'Total sales' },
			wordArt: { outline: '45460D', shadow: '45460D', warp: { preset: 'textWave1' } },
		},
	});
});

test('Symbol replaces the selected characters in the active cell edit', async ({ page }) => {
	await page.setViewportSize({ width: 1600, height: 850 });
	await openLanding(page);
	await newWorkbook(page);
	const host = editor(page);
	const field = host.getByRole('textbox', { name: 'Formula Bar', exact: true });
	await field.fill('abcdef');
	await field.evaluate((node: HTMLTextAreaElement) => node.setSelectionRange(2, 4));
	await host.getByRole('tab', { name: 'Insert', exact: true }).click();
	await clickRibbonControl(page, host.getByRole('button', { name: 'Symbol', exact: true }));
	const picker = host.locator('office-ui-dialog[data-dialog="symbol"]');
	await picker.getByRole('grid').getByRole('button', { name: 'U+00A9', exact: true }).click();
	await picker.getByRole('button', { name: 'Insert', exact: true }).click();
	await expect(field).toHaveValue('ab©ef');
});
