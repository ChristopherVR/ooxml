import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import { styleRefFixture } from './support/styleref-fixture';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: header STYLEREF fields follow the body page by page and are saved`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'styleref.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await styleRefFixture()),
		});
		const editor = page.locator('docx-editor');
		await expect(editor.locator('.ProseMirror')).toContainText('How the work is done.');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		await editor.getByRole('combobox', { name: 'Layout view', exact: true }).selectOption('print');
		const sheets = editor.locator('.dve-canvas > .dve-print-pages .dve-print-page');
		await expect(sheets).toHaveCount(2);
		const headers = () =>
			sheets.evaluateAll((elements) =>
				elements.map((sheet) => sheet.querySelector('.dve-print-header')?.textContent?.trim()),
			);
		// Each page's header shows the document number and that page's heading.
		await expect.poll(headers).toEqual(['No. DOC-0001 | Scope', 'No. DOC-0001 | Method']);

		const number = editor.locator('.dve-print-column .dve-print-line', { hasText: 'DOC-0001' });
		await number.locator('span').first().click();
		await page.keyboard.press('End');
		for (let index = 0; index < 4; index++) await page.keyboard.press('Backspace');
		await page.keyboard.type('0042');
		await expect.poll(headers).toEqual(['No. DOC-0042 | Scope', 'No. DOC-0042 | Method']);

		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const header = await (
			await JSZip.loadAsync(Buffer.from(bytes))
		)
			.file('word/header1.xml')!
			.async('string');
		// The number is the same on every page, so its saved result follows the body; the heading
		// depends on the page and keeps the result it was saved with.
		expect(header).toContain('<w:t>DOC-0042</w:t>');
		expect(header).toContain('<w:t>Scope</w:t>');
		expect(errors).toEqual([]);
	});
