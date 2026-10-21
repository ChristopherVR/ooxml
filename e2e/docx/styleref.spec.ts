import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { fileInput } from './helpers';
import { styleRefFixture, styleRefLineBreakFixture } from './support/styleref-fixture';

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

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: Update Fields keeps a line break in a body STYLEREF result, written once`, async ({
		page,
	}) => {
		const errors: string[] = [];
		page.on('pageerror', (error) => errors.push(error.message));
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'styleref-break.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(await styleRefLineBreakFixture()),
		});
		const editor = page.locator('docx-editor');
		const paragraphs = editor.locator('.ProseMirror > p');
		await expect(paragraphs.nth(2)).toContainText('Then old.');
		// Each paragraph's text, a line break shown as `|`.
		const shown = () =>
			paragraphs.evaluateAll((elements) =>
				elements.map((element) => {
					const copy = element.cloneNode(true) as HTMLElement;
					for (const item of copy.querySelectorAll('br')) item.replaceWith('|');
					return copy.textContent;
				}),
			);
		await editor.getByRole('tab', { name: 'References', exact: true }).click();
		const update = editor.getByRole('button', { name: 'Update fields', exact: true });
		await update.click();
		const expected = [
			'Chapter 1|Introduction',
			'See Chapter 1|Introduction.',
			'Then Chapter 1|Introduction.',
		];
		await expect.poll(shown).toEqual(expected);
		// A second update has nothing to change.
		await update.click();
		await expect.poll(shown).toEqual(expected);

		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		const document = await (
			await JSZip.loadAsync(Buffer.from(bytes))
		)
			.file('word/document.xml')!
			.async('string');
		expect(document.match(/<w:br\/>/g)).toHaveLength(3);
		expect(document).not.toContain('old');
		expect(errors).toEqual([]);
	});
