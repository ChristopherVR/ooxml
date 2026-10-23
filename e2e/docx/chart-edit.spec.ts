import { expect, test, type Page } from '@playwright/test';
import JSZip from 'jszip';
import { dialogByHeading, fileInput, newDocument } from './helpers';

const FRAMEWORKS = ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'] as const;
const WORKBOOK = 'word/embeddings/Microsoft_Excel_Worksheet.xlsx';

const editorOf = (page: Page) => page.locator('docx-editor');

async function saveBytes(page: Page): Promise<Uint8Array> {
	const list = await editorOf(page).evaluate(async (element) =>
		Array.from(await (element as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
	);
	return Uint8Array.from(list);
}

/** `ref -> text` of the first worksheet of the chart's embedded workbook. */
async function workbookCells(zip: JSZip): Promise<string> {
	const workbook = await JSZip.loadAsync(await zip.file(WORKBOOK)!.async('uint8array'));
	return workbook.file('xl/worksheets/sheet1.xml')!.async('string');
}

for (const framework of FRAMEWORKS) {
	test(`${framework}: Insert Chart, Edit Data and Change Chart Type save and reload`, async ({
		page,
	}) => {
		test.setTimeout(120_000);
		await page.setViewportSize({ width: 1400, height: 900 });
		await page.goto(`/?framework=${framework}`);
		await newDocument(page);
		const editor = editorOf(page);
		await editor.locator('.ProseMirror').click();
		await page.keyboard.type('Quarterly results:');

		// Insert > Chart: the shared picker with a live preview; OK puts a column chart inline.
		await editor.getByRole('tab', { name: 'Insert', exact: true }).click();
		await editor.getByRole('button', { name: 'Insert chart', exact: true }).click();
		const insert = dialogByHeading(editor, 'Insert Chart');
		const gallery = insert.locator('office-ui-gallery[mode="panel"]');
		await expect(gallery.locator('.tile svg')).toHaveCount(8);
		await insert.getByLabel('Chart title', { exact: true }).fill('Results');
		await insert.getByRole('button', { name: 'OK', exact: true }).click();
		const chart = editor.locator('.ProseMirror [data-docx-chart="1"]');
		await expect(chart).toHaveCount(1);
		await expect(chart.locator('svg g[data-chart-series="2"]').first()).toBeAttached();
		await expect(chart.locator('svg text', { hasText: 'Results' })).toHaveCount(1);
		await expect(editor.getByRole('tab', { name: 'Chart Design', exact: true })).toBeVisible();

		// Edit Data: the shared data sheet; a typed value repaints the chart at once.
		await editor.getByRole('button', { name: 'Edit Data', exact: true }).click();
		const data = dialogByHeading(editor, 'Edit Data');
		const sheet = data.locator('office-ui-chart-data-sheet');
		const cell = sheet.locator('input[data-row="1"][data-col="1"]');
		await expect(cell).toHaveValue('4.3');
		await cell.fill('12');
		await cell.press('Enter');
		await expect(sheet.locator('input[data-row="1"][data-col="1"]')).toHaveValue('12');
		await data.getByRole('button', { name: 'OK', exact: true }).click();

		// Change Chart Type: the picker in change-only mode switches the chart to lines.
		await editor.locator('.ProseMirror [data-docx-chart="1"]').click();
		await editor.getByRole('tab', { name: 'Chart Design', exact: true }).click();
		await editor.getByRole('button', { name: 'Change Chart Type', exact: true }).click();
		const change = dialogByHeading(editor, 'Change Chart Type');
		await change.getByRole('button', { name: 'Line', exact: true }).click();
		await change.getByRole('button', { name: 'OK', exact: true }).click();
		await expect(editor.locator('.ProseMirror [data-docx-chart="1"]')).toHaveCount(1);

		// Saved: a chart part with the line group and the edited value, and a workbook that agrees.
		const bytes = await saveBytes(page);
		const zip = await JSZip.loadAsync(bytes);
		const part = await zip.file('word/charts/chart1.xml')!.async('string');
		expect(part).toContain('<c:lineChart>');
		expect(part).toContain('<c:v>12</c:v>');
		expect(part).toContain('<c:f>Sheet1!$B$2:$B$5</c:f>');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'/word/charts/chart1.xml',
		);
		const sheetXml = await workbookCells(zip);
		expect(sheetXml).toContain('<c r="B2"><v>12</v></c>');
		expect(sheetXml).toContain('Series 1');

		// Reloaded: the chart comes back drawn, with its title.
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'chart-edited.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		const reloaded = editor.locator('.ProseMirror [data-docx-chart="1"]');
		await expect(reloaded).toHaveCount(1);
		await expect(reloaded.locator('svg text', { hasText: 'Results' })).toHaveCount(1);

		// Delete removes the selected chart.
		await reloaded.click();
		await expect(editor.getByRole('tab', { name: 'Chart Design', exact: true })).toBeVisible();
		await page.keyboard.press('Delete');
		await expect(editor.locator('.ProseMirror [data-docx-chart]')).toHaveCount(0);
		await expect(editor.getByRole('tab', { name: 'Chart Design', exact: true })).toBeHidden();
	});
}
