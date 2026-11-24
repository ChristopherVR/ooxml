import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, test } from '@playwright/test';
import { fileNameLabel, newDocument } from './helpers';

const openFile = async (page: import('@playwright/test').Page) => {
	const editor = page.locator('docx-editor');
	await editor.locator('.dve-ribbon .file').click();
	return editor;
};
const navigate = (editor: import('@playwright/test').Locator, name: string) =>
	editor.locator('.dve-backstage .item', { hasText: name }).click();
/** File > Export > Create a PDF's Options step (its fields are the shared dialog's slotted children). */
const pdfOptions = (editor: import('@playwright/test').Locator) =>
	editor.locator('.dve-pdf-options-dialog');
/** A downloaded PDF's raw bytes and its page text (the `Tj` strings of its content streams). */
async function readPdf(path: string): Promise<{ raw: string; words: string }> {
	const bytes = await readFile(path);
	const raw = bytes.toString('latin1');
	const { inflateSync } = await import('node:zlib');
	const text = [...raw.matchAll(/<< \/Filter \/FlateDecode \/Length (\d+) >>\nstream\n/g)]
		.map((match) => {
			const start = match.index + match[0].length;
			return inflateSync(bytes.subarray(start, start + Number(match[1]))).toString('latin1');
		})
		.join('\n');
	const words = [...text.matchAll(/\(((?:\\.|[^\\)])*)\) Tj/g)].map((match) => match[1]).join('');
	return { raw, words };
}

test.describe('File tab', () => {
	test.beforeEach(async ({ page }) => {
		await page.goto('/?framework=vanilla');
		await newDocument(page);
		await page.locator('docx-editor .ProseMirror').first().click();
		await page.keyboard.type('Hello file tab');
	});

	test('Info reports the document and its properties', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Info');
		const content = editor.locator('.dve-backstage-content:not([hidden])');
		await expect(content).toContainText('Document1.docx');
		await expect(content).toContainText('Unsaved changes');
		await expect(editor.locator('.dve-backstage-properties')).toContainText('3');
	});

	test('Home and Options work on a narrow screen and contain keyboard focus', async ({ page }) => {
		await page.setViewportSize({ width: 390, height: 844 });
		const editor = await openFile(page);
		await expect(editor.locator('.dve-backstage-content:not([hidden]) h2')).toHaveText('Home');
		await navigate(editor, 'Options');
		const dialog = editor.locator('office-ui-options-dialog');
		await expect(dialog.getByRole('dialog')).toBeVisible();
		await dialog.getByRole('tab', { name: 'Display' }).click();
		await dialog.getByRole('checkbox', { name: 'Ruler', exact: true }).check();
		await dialog.getByRole('checkbox', { name: 'Show paragraph marks', exact: true }).check();
		const overflow = await dialog
			.getByRole('tabpanel')
			.evaluate((el) => el.scrollWidth > el.clientWidth);
		expect(overflow).toBe(false);
		await dialog.locator('[data-action="ok"]').click();
		await expect(editor.locator('.dve-ruler')).toBeAttached();
		await expect(editor.locator('.dve-paper')).toHaveAttribute('data-show-marks', '');
		await expect(editor.locator('.dve-ribbon .file')).toBeFocused();
	});

	test('Options closes with Escape without applying the edits', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Options');
		const dialog = editor.locator('office-ui-options-dialog');
		await dialog.getByRole('tab', { name: 'Display' }).click();
		await dialog.getByRole('checkbox', { name: 'Ruler', exact: true }).check();
		await page.keyboard.press('Escape');
		await expect(dialog.getByRole('dialog')).toBeHidden();
		await expect(editor.locator('.dve-ruler')).toHaveCount(0);
	});

	test('Save As renames the document and downloads it', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Save As');
		await editor.getByLabel('File name', { exact: true }).fill('Quarterly report');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Save' })
			.click();
		const download = await pending;
		expect(download.suggestedFilename()).toBe('Quarterly report.docx');
		await expect(fileNameLabel(page)).toHaveText('Quarterly report.docx');
		const zip = await JSZip.loadAsync(await readFile((await download.path())!));
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Hello file tab');
	});

	test('Export writes plain text', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Export as plain text' })
			.click();
		const download = await pending;
		expect(download.suggestedFilename()).toBe('Document1.txt');
		expect(await readFile((await download.path())!, 'utf8')).toBe('Hello file tab');
	});

	test('Create a PDF downloads a PDF with real text, without a print dialog', async ({ page }) => {
		let printed = false;
		await page.exposeFunction('notePrint', () => (printed = true));
		await page.evaluate(() => {
			window.print = () => void (window as unknown as { notePrint(): void }).notePrint();
		});
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Create PDF' })
			.click();
		// The Options step, then Export with its defaults.
		const options = pdfOptions(editor);
		await expect(options).toBeVisible();
		await options.getByRole('button', { name: 'Export' }).click();
		const download = await pending;
		expect(download.suggestedFilename()).toBe('Document1.pdf');
		const pdf = await readPdf((await download.path())!);
		expect(pdf.raw.startsWith('%PDF-')).toBe(true);
		expect(pdf.raw).toContain('/Type /Page ');
		expect(pdf.raw).not.toContain('/Subtype /Image');
		// The text is real PDF text, written word by word with its spaces.
		expect(pdf.words).toBe('Hello file tab');
		expect(pdf.raw).toContain('/StructTreeRoot');
		expect(printed).toBe(false);
	});

	test('the PDF Options step is keyboard accessible and Escape cancels', async ({ page }) => {
		let downloads = 0;
		page.on('download', () => downloads++);
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Create PDF' })
			.click();
		const options = pdfOptions(editor);
		await expect(options).toBeVisible();
		await expect(options.getByRole('radio', { name: 'All', exact: true })).toBeFocused();
		// Draft view has no current page and nothing is selected.
		await expect(options.getByRole('radio', { name: 'Current page' })).toBeDisabled();
		await expect(options.getByRole('radio', { name: 'Selection' })).toBeDisabled();
		// Tab keeps focus inside the dialog.
		for (let step = 0; step < 20; step++) {
			await page.keyboard.press('Tab');
			const inside = await options.evaluate((dialog) => {
				let active: Element | null = document.activeElement;
				while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
				return dialog.contains(active) || Boolean(dialog.shadowRoot?.contains(active));
			});
			expect(inside).toBe(true);
		}
		await page.keyboard.press('Escape');
		await expect(options).toBeHidden();
		await expect(editor.locator('.dve-ribbon .file')).toBeFocused();
		expect(downloads).toBe(0);
	});

	test('the PDF Options step exports the selection as PDF/A without bookmarks', async ({
		page,
	}) => {
		await page.keyboard.press('Enter');
		await page.keyboard.type('Second line');
		await page.keyboard.press('Shift+Home');
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Create PDF' })
			.click();
		const options = pdfOptions(editor);
		await options.getByRole('radio', { name: 'Selection' }).check();
		await options.getByRole('radio', { name: 'Document showing markup' }).check();
		await options.getByRole('checkbox', { name: 'Create bookmarks using' }).uncheck();
		await expect(options.getByRole('radio', { name: 'Headings' })).toBeDisabled();
		await options
			.getByRole('checkbox', { name: 'Document structure tags for accessibility' })
			.uncheck();
		await options.getByRole('checkbox', { name: 'PDF/A compliant' }).check();
		await page.keyboard.press('Enter');
		const pdf = await readPdf((await (await pending).path())!);
		expect(pdf.words).toBe('Second line');
		expect(pdf.raw).toContain('<pdfaid:part>2</pdfaid:part>');
		expect(pdf.raw).not.toContain('/StructTreeRoot');
		expect(pdf.raw).not.toContain('/Type /Outlines');
	});

	test('the PDF Options step exports a page range of Print Layout', async ({ page }) => {
		await page.keyboard.press('Control+Enter');
		await page.keyboard.type('Page two');
		await page
			.locator('docx-editor')
			.getByRole('button', { name: 'Print Layout', exact: true })
			.click();
		const editor = await openFile(page);
		await navigate(editor, 'Export');
		const pending = page.waitForEvent('download');
		await editor
			.locator('.dve-backstage-content:not([hidden])')
			.getByRole('button', { name: 'Create PDF' })
			.click();
		const options = pdfOptions(editor);
		await expect(options.getByRole('radio', { name: 'Current page' })).toBeEnabled();
		const from = options.getByRole('spinbutton', { name: 'From' });
		await from.fill('3');
		await options.getByRole('button', { name: 'Export' }).click();
		// Page 3 does not exist: the field is marked and the dialog stays open.
		await expect(from).toHaveAttribute('aria-invalid', 'true');
		await from.fill('2');
		await options.getByRole('spinbutton', { name: 'To' }).fill('2');
		await options.getByRole('button', { name: 'Export' }).click();
		const pdf = await readPdf((await (await pending).path())!);
		expect(pdf.words).toBe('Page two');
		expect(pdf.raw.match(/\/Type \/Page /g)).toHaveLength(1);
	});

	test('Options switches the display language and the review author', async ({ page }) => {
		const editor = await openFile(page);
		await navigate(editor, 'Options');
		const dialog = editor.locator('office-ui-options-dialog');
		await dialog.getByLabel('User name', { exact: true }).fill('Dana');
		await dialog.getByRole('tab', { name: 'Language' }).click();
		await dialog.getByLabel('Display language', { exact: true }).selectOption('fr');
		await dialog.locator('[data-action="ok"]').click();
		await expect(editor).toHaveAttribute('review-author', 'Dana');
		await expect(editor).toHaveAttribute('locale', 'fr');
		await editor.locator('.dve-ribbon .file').click();
		await expect(editor.locator('.dve-backstage .item', { hasText: 'Exporter' })).toBeVisible();
	});

	test('Escape closes the backstage and returns to the document', async ({ page }) => {
		const editor = await openFile(page);
		await expect(editor.locator('.dve-backstage')).toBeVisible();
		await page.keyboard.press('Escape');
		await expect(editor.locator('.dve-backstage')).toBeHidden();
	});
});
