import { expect, test } from '@playwright/test';
import JSZip from 'jszip';
import type { EditorView } from 'prosemirror-view';
import type { TextSelection } from 'prosemirror-state';
import type { DocxEditorElement } from '../../viewers/docx/packages/web-component/src';
import { fileInput } from './helpers';
import { fieldDocx } from './field-fixtures';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
	for (const action of ['Accept all', 'Reject all'])
		test(`${framework}: ${action} resolves a tracked simple-field result deletion with one instruction`, async ({
			page,
		}) => {
			await page.goto(`/?framework=${framework}`);
			await (
				await fileInput(page)
			).setInputFiles({
				name: 'field.docx',
				mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
				buffer: await fieldDocx(true),
			});
			const editor = page.locator('docx-editor');
			await expect(editor.locator('.ProseMirror')).toContainText('Ann');
			await editor.getByRole('tab', { name: 'Review', exact: true }).click();
			await editor.getByRole('button', { name: 'Track changes', exact: true }).click();
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
			await page.keyboard.press('Delete');
			const savedXml = async () => {
				const bytes = await editor.evaluate(async (element) =>
					Array.from(await (element as DocxEditorElement).saveBytes()),
				);
				return (await JSZip.loadAsync(new Uint8Array(bytes)))
					.file('word/document.xml')!
					.async('string');
			};
			const tracked = await savedXml();
			expect(tracked.match(/w:fldCharType="begin"/g)).toHaveLength(1);
			expect(tracked).not.toContain('<w:fldSimple');
			expect(tracked.match(/<w:del\b/g)).toHaveLength(1);
			expect(tracked).toContain('<w:delText>Ann</w:delText>');
			await editor.getByRole('button', { name: action, exact: true }).click();
			const resolved = await savedXml();
			expect(resolved.match(/w:fldCharType="begin"/g)).toHaveLength(1);
			expect(resolved).not.toContain('<w:del');
			expect(resolved.includes('<w:t>Ann</w:t>')).toBe(action === 'Reject all');
			await editor.getByRole('button', { name: 'Undo', exact: true }).click();
			expect(await savedXml()).toContain('<w:delText>Ann</w:delText>');
		});
