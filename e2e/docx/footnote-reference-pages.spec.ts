import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import { clickInEditor, fileInput } from './helpers';

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(): Promise<Buffer> {
	const zip = new JSZip();
	const opening = Array.from(
		{ length: 9 },
		(_, index) => `<w:t>Opening ${index}</w:t><w:br/>`,
	).join('');
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p><w:pPr><w:widowControl w:val="0"/></w:pPr><w:r>${opening}<w:t>Claim</w:t></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p><w:sectPr><w:pgSz w:w="4500" w:h="2250"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="300"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/footnotes.xml',
		`<w:footnotes xmlns:w="${w}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Evidence sources</w:t></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>Reference source</w:t></w:r></w:p></w:footnote></w:footnotes>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: footnotes follow their reference line and custom separators survive editing, undo and save/reopen`, async ({
		page,
	}) => {
		page.on('dialog', (dialog) => void dialog.accept());
		await page.goto(`/?framework=${framework}`);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'late-footnote.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: await fixture(),
		});
		const editor = page.locator('docx-editor');
		await editor.getByRole('tab', { name: 'View', exact: true }).click();
		const layoutView = editor.getByRole('combobox', { name: 'Layout view', exact: true });
		await layoutView.selectOption('print');
		const sheets = editor.locator('.dve-print-page');
		const verify = async () => {
			await expect(sheets.first()).toBeVisible();
			await expect(sheets.first().locator('.dve-print-footnotes')).toHaveCount(0);
			const reference = sheets.filter({
				has: page.locator('.dve-print-column .dve-print-line', { hasText: 'Claim' }),
			});
			await expect(reference).toHaveCount(1);
			await expect(reference.locator('.dve-print-custom-footnote-separator')).toContainText(
				'Evidence sources',
			);
			await expect(reference.locator('.dve-print-footnote-separator')).toHaveCount(0);
			await expect(reference.locator('.dve-print-footnote[data-note-id="2"]')).toContainText(
				'Reference source',
			);
			await expect(editor.locator('.dve-print-footnote[data-note-id="2"]')).toHaveCount(1);
		};
		await verify();
		await layoutView.selectOption('draft');
		const body = editor.locator('.ProseMirror').first();
		await clickInEditor(body);
		await page.keyboard.press('Control+Home');
		await page.keyboard.type('Prefix ');
		await expect(body).toContainText('Prefix Opening 0');
		await page.keyboard.press('Control+z');
		await expect(body).not.toContainText('Prefix');
		await layoutView.selectOption('print');
		await verify();
		const bytes = await editor.evaluate(async (node) =>
			Array.from(await (node as unknown as { saveBytes(): Promise<Uint8Array> }).saveBytes()),
		);
		await (
			await fileInput(page)
		).setInputFiles({
			name: 'reopened.docx',
			mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			buffer: Buffer.from(bytes),
		});
		await layoutView.selectOption('print');
		await verify();
	});
}
