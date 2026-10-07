import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { TextSelection } from 'prosemirror-state';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';
import { fieldDocx } from './field-fixtures';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const simple of [true, false])
		for (const method of ['keyboard', 'fallback'])
			test(`${framework}: ${method} cut retains an empty ${simple ? 'simple' : 'complex'} field`, async ({
				page,
				context,
			}) => {
				await context.grantPermissions(['clipboard-read', 'clipboard-write']);
				await page.goto(`/?framework=${framework}`);
				await (
					await fileInput(page)
				).setInputFiles({
					name: 'field.docx',
					mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
					buffer: await fieldDocx(simple),
				});
				const editor = page.locator('docx-editor');
				const body = editor.locator('.ProseMirror');
				await expect(body).toContainText('Ann');
				await editor.evaluate((element) => {
					const view = (element as unknown as { view: EditorView }).view;
					let start = -1;
					view.state.doc.descendants((node, pos) => {
						if (start < 0 && node.isText && node.marks.some((mark) => mark.type.name === 'field'))
							start = pos;
					});
					if (start < 0) throw new Error('Missing field result');
					const Selection = view.state.selection.constructor as typeof TextSelection;
					view.dispatch(
						view.state.tr.setSelection(Selection.create(view.state.doc, start, start + 3)),
					);
					view.focus();
				});
				if (method === 'keyboard') await page.keyboard.press('Control+x');
				else {
					await page.evaluate(() => (document.execCommand = () => false));
					await editor.getByRole('button', { name: 'Cut', exact: true }).click();
				}
				await expect(body).toContainText('Author:  wrote this.');
				expect(await page.evaluate(() => navigator.clipboard.readText())).toBe('Ann');
				const bytes = await editor.evaluate(async (element) =>
					Array.from(await (element as DocxEditorElement).saveBytes()),
				);
				const xml = await (
					await JSZip.loadAsync(new Uint8Array(bytes))
				)
					.file('word/document.xml')!
					.async('string');
				expect(xml.match(/w:fldCharType="begin"/g)).toHaveLength(1);
				expect(xml).toContain('AUTHOR');
				await editor.getByRole('button', { name: 'Undo', exact: true }).click();
				await expect(body).toContainText('Author: Ann wrote this.');
			});
