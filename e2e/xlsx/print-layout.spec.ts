import { expect, test } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('manual page breaks, physical preview, reset, undo and saved reopening', async ({
	page,
}, info) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'First');
	await goToCell(page, 'D10');
	await typeInActiveCell(page, 'Last');
	await ribbon(page).getByRole('tab', { name: 'Page Layout', exact: true }).click();
	const breaks = () => ribbon(page).getByRole('button', { name: 'Breaks', exact: true }).click();
	await goToCell(page, 'A6');
	await breaks();
	await editor(page).getByRole('menuitem', { name: 'Insert Row Page Break', exact: true }).click();
	await goToCell(page, 'C1');
	await breaks();
	await editor(page)
		.getByRole('menuitem', { name: 'Insert Column Page Break', exact: true })
		.click();
	const preview = editor(page).locator('[data-dialog="print-preview"]');
	const open = () => ribbon(page).locator('[data-command="page.preview"]').click();
	await open();
	await expect(preview).toContainText('Pages: 4 | Scale: 100%');
	await expect(preview.frameLocator('iframe').locator('.xve-print-page')).toHaveCount(4);
	await expect(preview.frameLocator('iframe').locator('.xve-print-page').first()).toContainText(
		'First',
	);
	await expect(preview.frameLocator('iframe').locator('.xve-print-page').last()).toContainText(
		'Last',
	);
	await preview.getByRole('button', { name: 'Next Page', exact: true }).click();
	await expect(preview).toContainText('Page 2 of 4');
	await expect(
		preview.frameLocator('iframe').locator('.xve-print-page[data-page="2"]'),
	).toBeVisible();
	await preview.getByRole('button', { name: 'Previous Page', exact: true }).click();
	await preview.getByRole('button', { name: 'Actual Size', exact: true }).click();
	await expect(preview.getByRole('button', { name: 'Fit Page', exact: true })).toBeVisible();
	await preview.getByRole('button', { name: 'Fit Page', exact: true }).click();
	await page.screenshot({ path: info.outputPath('print-preview.png') });
	const html = await preview.locator('iframe').getAttribute('srcdoc');
	const printPage = await page.context().newPage();
	await printPage.setContent(html!);
	const pdf = await printPage.pdf({
		path: info.outputPath('browser-print.pdf'),
		preferCSSPageSize: true,
	});
	expect(
		Buffer.from(pdf)
			.toString('latin1')
			.match(/\/Type\s*\/Page\b/g),
	).toHaveLength(4);
	await printPage.close();
	await preview.getByRole('button', { name: 'Close', exact: true }).first().click();
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'page-breaks.xlsx',
			),
		bytes,
	);
	await ribbon(page).getByRole('tab', { name: 'Page Layout', exact: true }).click();
	await open();
	await expect(preview).toContainText('Pages: 4');
	await preview.getByRole('button', { name: 'Close', exact: true }).first().click();
	await breaks();
	await editor(page).getByRole('menuitem', { name: 'Reset All Page Breaks', exact: true }).click();
	await open();
	await expect(preview).toContainText('Pages: 1');
	await preview.getByRole('button', { name: 'Close', exact: true }).first().click();
	await page.keyboard.press('Control+Z');
	await open();
	await expect(preview).toContainText('Pages: 4');
});
