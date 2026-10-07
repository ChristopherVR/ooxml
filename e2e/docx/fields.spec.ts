import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { fileInput, fileNameLabel, newDocument, saveButton } from './helpers';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import type { EditorView } from 'prosemirror-view';
import type { TextSelection } from 'prosemirror-state';

import { fieldDocx, WORD_NS as w } from './field-fixtures';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const simple of [true, false])
		test(`${framework}: pasted ${simple ? 'simple' : 'complex'} field result HTML remains literal text`, async ({
			page,
		}) => {
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'field.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: await fieldDocx(simple),
			});
			const editor = page.locator('docx-editor');
			const result = editor.locator('[data-field="AUTHOR"]').first();
			await expect(result).toContainText('Ann');
			const html = await result.evaluate((element) => element.outerHTML);
			const body = editor.locator('.ProseMirror');
			await editor.evaluate((element, html) => {
				// Synthetic paste events have no native clipboard selection sequencing.
				const view = (element as unknown as { view: EditorView }).view;
				const Selection = view.state.selection.constructor as typeof TextSelection;
				view.dispatch(
					view.state.tr.setSelection(
						Selection.create(view.state.doc, view.state.doc.content.size - 1),
					),
				);
				view.focus();
				const data = new DataTransfer();
				data.setData('text/html', html);
				view.dom.dispatchEvent(
					new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }),
				);
			}, html);
			await expect(body).toContainText('wrote this.Ann');
			const model = await editor.evaluate(
				(element) => (element as DocxEditorElement).documentModel!,
			);
			expect(
				model.blocks.flatMap((block) =>
					block.type === 'paragraph' ? block.runs.filter((run) => run.field) : [],
				),
			).toHaveLength(1);
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml.match(simple ? /<w:fldSimple\b/g : /w:fldCharType="begin"/g)).toHaveLength(1);
			await editor.getByRole('button', { name: 'Undo', exact: true }).click();
			await expect(body).not.toContainText('wrote this.Ann');
		});

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const kind of ['complex', 'simple', 'adjacent'] as const)
		test(`${framework}: commenting on part of a ${kind} field result anchors the complete field`, async ({
			page,
		}) => {
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'field.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: await fieldDocx(kind !== 'complex', kind === 'adjacent'),
			});
			const editor = page.locator('docx-editor');
			const result = editor.locator('[data-field="AUTHOR"]').first();
			await expect(result).toContainText('Ann');
			await editor.locator('.ProseMirror').focus();
			await result.evaluate((element) => {
				const text = document.createTreeWalker(element, NodeFilter.SHOW_TEXT).nextNode()!;
				const selection = window.getSelection();
				selection!.setBaseAndExtent(text, 1, text, 2);
			});
			await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe('n');
			await editor.getByRole('button', { name: 'Show comments', exact: true }).click();
			const pane = editor.locator('.dve-comments-panel');
			await pane.getByRole('textbox', { name: 'New comment', exact: true }).fill('Whole field');
			await pane.getByRole('button', { name: 'Add comment', exact: true }).click();
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const zip = await JSZip.loadAsync(new Uint8Array(bytes));
			const xml = await zip.file('word/document.xml')!.async('string');
			expect(xml.indexOf('<w:commentRangeStart')).toBeLessThan(
				xml.indexOf(kind !== 'complex' ? '<w:fldSimple' : 'w:fldCharType="begin"'),
			);
			expect(xml.indexOf('<w:commentRangeEnd')).toBeGreaterThan(
				xml.indexOf(kind !== 'complex' ? '</w:fldSimple>' : 'w:fldCharType="end"'),
			);
			expect(xml.match(/<w:commentRangeStart\b/g)).toHaveLength(1);
			if (kind === 'adjacent') {
				expect(xml.match(/<w:fldSimple\b/g)).toHaveLength(2);
				expect(xml.indexOf('<w:commentRangeEnd')).toBeLessThan(xml.lastIndexOf('<w:fldSimple'));
			}
			await editor.getByRole('button', { name: 'Undo', exact: true }).click();
			await expect(pane).not.toContainText('Whole field');
			const undone = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const undoneZip = await JSZip.loadAsync(new Uint8Array(undone));
			expect(await undoneZip.file('word/document.xml')!.async('string')).not.toContain(
				'<w:commentRangeStart',
			);
			expect(undoneZip.file('word/comments.xml')).toBeNull();
			await editor.getByRole('button', { name: 'Redo', exact: true }).click();
			await expect(pane).toContainText('Whole field');
		});

test('edits text around and inside a complex field and saves the field intact', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'field.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await fieldDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('field.docx');
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await expect(body).toContainText('Author: Ann wrote this.');
	await expect(body).toHaveAttribute('contenteditable', 'true');

	await editor.locator('[data-field="AUTHOR"]').dblclick();
	// Windows word selection includes the trailing space, crossing the hidden field end marker.
	await page.keyboard.type('Bea');
	await page.keyboard.press('End');
	await page.keyboard.type(' Twice.');
	await expect(body).toContainText(/Author: Bea ?wrote this\. Twice\./);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(
		/<w:fldChar w:fldCharType="begin"\/><\/w:r><w:r><w:instrText xml:space="preserve"> AUTHOR <\/w:instrText><\/w:r><w:r><w:fldChar w:fldCharType="separate"\/><\/w:r><w:r><w:t>Bea<\/w:t><\/w:r><w:r><w:fldChar w:fldCharType="end"\/>/,
	);
	expect(xml).toMatch(/ ?wrote this\. Twice\./);
	expect(errors).toEqual([]);
});

async function headingsDocx(): Promise<Buffer> {
	const zip = await JSZip.loadAsync(await fieldDocx());
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${w}"><w:body><w:p/><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Overview</w:t></w:r></w:p><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading2"/><w:pageBreakBefore/></w:pPr><w:r><w:t>Details</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'nodebuffer' });
}

test('inserts a table of contents from the References tab and saves it as a TOC field', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await (
		await fileInput(page)
	).setInputFiles({
		name: 'headings.docx',
		mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		buffer: await headingsDocx(),
	});
	await expect(fileNameLabel(page)).toHaveText('headings.docx');
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await body.locator('p').first().click();
	await editor.getByRole('tab', { name: 'References', exact: true }).click();
	await editor.getByRole('combobox', { name: 'Insert table of contents' }).selectOption('3');
	await expect(body.locator('p').first()).toHaveText(/^Overview\s+1$/);
	await expect(body.locator('p').nth(1)).toHaveText(/^Details\s+2$/);
	// The editing surface sizes the tab to the right-aligned stop: the page number ends at the
	// right margin, and the tab shows a dot leader.
	await expect(body.locator('.dve-tab-leader-dot')).toHaveCount(2);
	await expect
		.poll(() =>
			body.evaluate((root) => {
				const paragraph = root.querySelectorAll('p')[1];
				if (!paragraph) throw new Error('The second paragraph is missing.');
				const range = document.createRange();
				range.selectNodeContents(paragraph);
				return Math.round(root.getBoundingClientRect().right - range.getBoundingClientRect().right);
			}),
		)
		.toBeLessThanOrEqual(2);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	expect(xml).toMatch(/<w:instrText xml:space="preserve"> TOC [^<]*<\/w:instrText>/);
	expect(xml).toContain('<w:tab w:val="right" w:leader="dot" w:pos="9360"/>');
	// Entries are hyperlinks to _Toc bookmarks on the headings, with PAGEREF page numbers.
	expect(xml).toMatch(
		/<w:hyperlink w:anchor="(_Toc\d+)"[^>]*><w:r><w:t>Details<\/w:t><w:tab\/><\/w:r>.*?PAGEREF \1 .*?<w:t>2<\/w:t>/,
	);
	expect(xml).toMatch(/<w:bookmarkStart w:id="\d+" w:name="_Toc\d+"\/><w:r><w:t>Details<\/w:t>/);
	expect(errors).toEqual([]);
});

test('records cut and paste under Track Changes as a tracked move and saves it', async ({
	page,
	context,
}) => {
	await context.grantPermissions(['clipboard-read', 'clipboard-write']);
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.goto('/');
	await newDocument(page);
	const editor = page.locator('docx-editor');
	const body = editor.locator('.ProseMirror').first();
	await body.locator('p').first().click();
	await page.keyboard.type('Alpha Beta');
	await page.keyboard.press('Enter');
	await page.keyboard.type('Gamma');
	await editor.getByRole('tab', { name: 'Review', exact: true }).click();
	await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
	// Select "Beta" and cut it, then paste it at the end of "Gamma".
	await body
		.locator('p')
		.first()
		.dblclick({ position: { x: 60, y: 8 } });
	await page.keyboard.press('Control+x');
	await page.keyboard.press('ArrowDown');
	await page.keyboard.press('End');
	// Let the editor read the caret move (a selectionchange task) before pasting.
	await page.evaluate(
		() => new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 50))),
	);
	await page.keyboard.press('Control+v');
	await expect(editor.locator('.dve-revision-move')).toHaveCount(2);

	const pending = page.waitForEvent('download');
	await saveButton(page).click();
	const zip = await JSZip.loadAsync(await readFile((await (await pending).path())!));
	const xml = await zip.file('word/document.xml')!.async('string');
	const name = /<w:moveFromRangeStart w:id="\d+"[^>]* w:name="(move\d+)"\/>/.exec(xml)?.[1];
	expect(name).toBeTruthy();
	expect(xml).toContain(`<w:moveToRangeStart`);
	// Native Word represents moved source text with w:t, rather than a deletion run.
	expect(xml).toMatch(/<w:moveFrom w:id="\d+"[^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t>Beta<\/w:t>/);
	expect(xml).toMatch(/<w:moveTo w:id="\d+"[^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:t>Beta<\/w:t>/);
	// Pasted CSS values are converted to what Word accepts, never written verbatim.
	expect(xml).not.toMatch(/rgb\(|sans-serif/);
	expect(xml).not.toContain('dve-rev');
	expect(errors).toEqual([]);
});
