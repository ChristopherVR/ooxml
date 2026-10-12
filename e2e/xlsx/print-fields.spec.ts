import { expect, test } from '@playwright/test';
import { editor, goToCell, newWorkbook, ribbon, typeInActiveCell } from './helpers';

test('authors repeated titles and native header fields with numbering and saved reopening', async ({
	page,
}, info) => {
	await newWorkbook(page);
	await typeInActiveCell(page, 'Repeated title');
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
	await ribbon(page).locator('[data-command="page.setup"]').first().click();
	const setup = editor(page).locator('[data-dialog="page-setup"]');
	await setup.getByLabel('Use first page number', { exact: true }).check();
	await setup.getByLabel('First page number', { exact: true }).fill('5');
	await setup.getByRole('tab', { name: 'Sheet', exact: true }).click();
	await setup.getByLabel('Rows to repeat at top', { exact: true }).fill('$1:$1');
	await setup.getByLabel('Columns to repeat at left', { exact: true }).fill('$A:$A');
	await setup.getByRole('tab', { name: 'Header/Footer', exact: true }).click();
	await setup.getByLabel('Header:', { exact: true }).fill('&RPage ');
	await setup.getByLabel('Insert Field', { exact: true }).selectOption('&P');
	await expect(setup.getByLabel('Header:', { exact: true })).toHaveValue('&RPage &P');
	await setup.getByLabel('Header:', { exact: true }).fill('&L&A&RPage &P of &N');
	await setup.getByLabel('Footer:', { exact: true }).fill('&L&F&R&D &T');
	await setup.getByLabel('Different first page', { exact: true }).check();
	await setup.getByLabel('First Page Header:', { exact: true }).fill('&CFirst &P of &N');
	await setup.getByLabel('Different odd and even pages', { exact: true }).check();
	await setup.getByLabel('Even Page Header:', { exact: true }).fill('&CEven &P of &N');
	await setup.getByLabel('Even Page Footer:', { exact: true }).fill('&L&F&R&D &T');
	await setup.getByRole('button', { name: 'OK', exact: true }).click();
	const bytes = await editor(page).evaluate(async (node) => [
		...new Uint8Array(
			await (await (node as unknown as { save(): Promise<Blob> }).save()).arrayBuffer(),
		),
	]);
	await editor(page).evaluate(
		async (node, data) =>
			(node as unknown as { load(bytes: Uint8Array, name: string): Promise<void> }).load(
				new Uint8Array(data),
				'print-fields.xlsx',
			),
		bytes,
	);
	await ribbon(page).getByRole('tab', { name: 'Page Layout', exact: true }).click();
	await ribbon(page).locator('[data-command="page.preview"]').click();
	const preview = editor(page).locator('[data-dialog="print-preview"]');
	const pages = preview.frameLocator('iframe').locator('.xve-print-page');
	await expect(pages).toHaveCount(4);
	await expect(pages.first().locator('.xve-print-header')).toContainText('First 5 of 4');
	await preview.getByRole('button', { name: 'Next Page', exact: true }).click();
	await expect(pages.nth(1).locator('.xve-print-header')).toContainText('Even 6 of 4');
	await expect(pages.nth(1).locator('.xve-print-footer')).toContainText('print-fields.xlsx');
	await expect(pages.nth(1).locator('.xve-print-footer')).toContainText(
		/\d{1,2}\/\d{1,2}\/\d{2,4}/,
	);
	await expect(pages.nth(1).locator('.xve-print-footer')).toContainText(/\d{1,2}:\d{2}/);
	for (let i = 0; i < 4; i++) await expect(pages.nth(i)).toContainText('Repeated title');
	await page.screenshot({ path: info.outputPath('print-fields.png') });
	const html = await preview.locator('iframe').getAttribute('srcdoc');
	const printPage = await page.context().newPage();
	await printPage.setContent(html!);
	const pdf = await printPage.pdf({
		path: info.outputPath('browser-fields.pdf'),
		preferCSSPageSize: true,
	});
	expect(
		Buffer.from(pdf)
			.toString('latin1')
			.match(/\/Type\s*\/Page\b/g),
	).toHaveLength(4);
	await printPage.close();
});
