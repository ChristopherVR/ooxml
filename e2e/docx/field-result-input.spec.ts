import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { TextSelection } from 'prosemirror-state';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';
import { fieldDocx } from './field-fixtures';

const replacements = [
	{ name: 'partial', from: 1, to: 2, result: 'AXn' },
	{ name: 'whole', from: 0, to: 3, result: 'X' },
	{ name: 'start', from: 0, to: 1, result: 'Xnn' },
	{ name: 'end', from: 2, to: 3, result: 'AnX' },
];
for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const simple of [true, false])
		for (const method of ['typing', 'paste'])
			for (const replacement of replacements)
				test(`${framework}: ${method} replaces ${replacement.name} ${simple ? 'simple' : 'complex'} field result`, async ({
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
					await expect(editor.locator('[data-field="AUTHOR"]').first()).toContainText('Ann');
					await editor.evaluate(
						(element, { from, to, method }) => {
							const view = (element as unknown as { view: EditorView }).view;
							let start = -1;
							view.state.doc.descendants((node, pos) => {
								if (
									start < 0 &&
									node.isText &&
									node.marks.some((mark) => mark.type.name === 'field')
								)
									start = pos;
							});
							if (start < 0) throw new Error('Missing cached field result');
							const Selection = view.state.selection.constructor as typeof TextSelection;
							view.dispatch(
								view.state.tr.setSelection(
									Selection.create(view.state.doc, start + from, start + to),
								),
							);
							view.focus();
							if (method === 'paste') {
								const data = new DataTransfer();
								data.setData('text/html', '<b>X</b>');
								view.dom.dispatchEvent(
									new ClipboardEvent('paste', {
										clipboardData: data,
										bubbles: true,
										cancelable: true,
									}),
								);
							}
						},
						{ ...replacement, method },
					);
					if (method === 'typing') await page.keyboard.type('X');
					await expect(editor.locator('.ProseMirror')).toContainText(
						`Author: ${replacement.result} wrote this.`,
					);
					const model = await editor.evaluate(
						(element) => (element as DocxEditorElement).documentModel!,
					);
					const runs = model.blocks.flatMap((block) =>
						block.type === 'paragraph' ? block.runs.filter((run) => run.field) : [],
					);
					expect(runs.map((run) => run.text).join('')).toBe(replacement.result);
					if (simple) expect(new Set(runs.map((run) => run.fieldInstanceId)).size).toBe(1);
					const bytes = await editor.evaluate(async (element) =>
						Array.from(await (element as DocxEditorElement).saveBytes()),
					);
					const xml = await (
						await JSZip.loadAsync(new Uint8Array(bytes))
					)
						.file('word/document.xml')!
						.async('string');
					expect(xml.match(simple ? /<w:fldSimple\b/g : /w:fldCharType="begin"/g)).toHaveLength(1);
					expect(xml).toContain(' AUTHOR ');
					await editor.getByRole('button', { name: 'Undo', exact: true }).click();
					await expect(editor.locator('.ProseMirror')).toContainText('Author: Ann wrote this.');
				});
