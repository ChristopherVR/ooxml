import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { TextSelection } from 'prosemirror-state';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';
import { fieldDocx } from './field-fixtures';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const side of ['prefix', 'suffix'] as const)
		test(`${framework}: ${side} clipboard slice removes incomplete field structure`, async ({
			page,
		}) => {
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'field.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: await fieldDocx(),
			});
			const editor = page.locator('docx-editor');
			await expect(editor.locator('.ProseMirror')).toContainText('Author: Ann wrote this.');
			const copied = await editor.evaluate((element, boundary) => {
				const view = (element as unknown as { view: EditorView }).view;
				const Selection = view.state.selection.constructor as typeof TextSelection;
				let result = -1;
				view.state.doc.descendants((node, pos) => {
					if (result < 0 && node.isText && node.marks.some((mark) => mark.type.name === 'field'))
						result = pos;
				});
				if (result < 0) throw new Error('Missing field result');
				const from = boundary === 'prefix' ? result - 4 : result + 2;
				const to = boundary === 'prefix' ? result + 1 : result + 5;
				view.dispatch(view.state.tr.setSelection(Selection.create(view.state.doc, from, to)));
				const original = view.state.doc;
				const slice = view.serializeForClipboard(view.state.selection.content());
				if (!view.state.doc.eq(original)) throw new Error('Copy changed source');
				view.dispatch(
					view.state.tr.setSelection(
						Selection.create(view.state.doc, view.state.doc.content.size - 1),
					),
				);
				view.focus();
				if (!view.pasteHTML(slice.dom.innerHTML)) throw new Error('Paste failed');
				return { text: slice.text, html: slice.dom.innerHTML };
			}, side);
			expect(copied.text).toBe(side === 'prefix' ? ' A' : 'n ');
			expect(copied.html).not.toContain('data-field');
			await expect(editor.locator('.ProseMirror')).toContainText(
				`Author: Ann wrote this.${copied.text}`,
			);
			const bytes = await editor.evaluate(async (element) =>
				Array.from(await (element as DocxEditorElement).saveBytes()),
			);
			const xml = await (
				await JSZip.loadAsync(new Uint8Array(bytes))
			)
				.file('word/document.xml')!
				.async('string');
			expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1);
			expect(xml.match(/w:fldCharType="end"/g)).toHaveLength(1);
			expect(xml).toContain('AUTHOR');
			await editor.getByRole('button', { name: 'Undo', exact: true }).click();
			await expect(editor.locator('.ProseMirror')).toHaveText('Author: Ann wrote this.');
		});
