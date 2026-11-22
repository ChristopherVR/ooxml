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

test.describe('phone Insert ribbon', () => {
	test.use({ viewport: { width: 375, height: 812 }, hasTouch: true, isMobile: true });

	test('inserts a symbol at the Formula Bar selection through the folded ribbon', async ({
		page,
	}) => {
		await newWorkbook(page);
		const host = editor(page);
		const field = host.getByRole('textbox', { name: 'Formula Bar', exact: true });
		await field.fill('abcdef');
		await field.evaluate((node: HTMLTextAreaElement) => node.setSelectionRange(2, 4));
		await host.getByRole('tab', { name: 'Insert', exact: true }).tap();
		await host.getByRole('button', { name: 'Symbols', exact: true }).tap();
		await host.getByRole('button', { name: 'Symbol', exact: true }).tap();
		const picker = host.locator('office-ui-dialog[data-dialog="symbol"]');
		await picker.getByRole('textbox', { name: 'Character code', exact: true }).fill('U+00A9');
		await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
		await expect(field).toHaveValue('ab©ef');
	});

	test('inserts WordArt and an icon through folded groups', async ({ page }) => {
		await newWorkbook(page);
		const host = editor(page);
		await host.getByRole('tab', { name: 'Insert', exact: true }).tap();
		await host.getByRole('button', { name: 'Symbols', exact: true }).tap();
		await host.getByRole('button', { name: 'WordArt', exact: true }).tap();
		const picker = host.locator('office-ui-dialog[open]');
		await picker.getByRole('textbox', { name: 'WordArt text' }).fill('Phone WordArt');
		await picker.getByRole('button', { name: 'WordArt style 31', exact: true }).tap();
		await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
		const wordArt = host.locator('.xg-obj[aria-label="WordArt"]');
		await expect(wordArt).toContainText('Phone WordArt');
		await expect(wordArt.locator('linearGradient')).toHaveCount(1);

		await host.getByRole('button', { name: 'Illustrations', exact: true }).tap();
		await host.getByRole('button', { name: 'Icons', exact: true }).tap();
		await picker.getByRole('searchbox', { name: 'Search icons' }).fill('Light bulb');
		await picker.getByRole('button', { name: 'Light bulb', exact: true }).tap();
		await picker.getByRole('searchbox', { name: 'Search icons' }).fill('Check mark');
		await picker.getByRole('button', { name: 'Check mark', exact: true }).tap();
		await picker.getByRole('button', { name: 'Insert', exact: true }).tap();
		await expect(host.locator('.xg-obj img')).toHaveCount(2);
	});
});
