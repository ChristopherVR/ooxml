// @vitest-environment jsdom
import { expect, it } from 'vitest';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { loadDocx, listRevisions, type DocumentModel } from 'ooxml-core/docx';
import { parseXml, buildXml } from 'ooxml-core/xml';
import { undo } from 'prosemirror-history';
import type { EditorView } from 'prosemirror-view';
import { DocxEditorElement } from './index';
import './index';

function content(model: DocumentModel) {
	return model.blocks.map((block) =>
		block.type === 'paragraph'
			? block.runs.map(
					({ sourceRunPropertiesXml: _source, restoredRunPropertiesXml: _restored, ...run }) => run,
				)
			: '[table]',
	);
}
function drawing(xml: string): string {
	return buildXml(
		parseXml(xml).getElementsByTagNameNS(
			'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
			'drawing',
		)[0]!,
	);
}
for (const name of ['picture', 'note', 'break', 'field'])
	for (const mode of ['accept', 'reject'] as const)
		it(`${mode}s native ${name} run formatting and preserves its content on export and undo`, async () => {
			const directory = resolve('../core/docx/__fixtures__/review-object-formatting');
			const sourceBytes = new Uint8Array(
				await readFile(resolve(directory, `${name}-tracked.docx`)),
			);
			const editor = document.createElement('docx-editor') as DocxEditorElement;
			document.body.append(editor);
			try {
				await editor.load(sourceBytes);
				const source = structuredClone(editor.documentModel!);
				expect(listRevisions(source)).toHaveLength(1);
				editor
					.shadowRoot!.querySelector<HTMLButtonElement>(
						`[aria-label="${mode === 'accept' ? 'Accept' : 'Reject'} all"]`,
					)!
					.click();
				expect(listRevisions(editor.documentModel!)).toHaveLength(0);
				const bytes = await editor.saveBytes();
				const exported = (await loadDocx(bytes)).model;
				const native = (
					await loadDocx(
						await readFile(
							resolve(directory, `${name}-${mode === 'accept' ? 'accepted' : 'rejected'}.docx`),
						),
					)
				).model;
				expect(listRevisions(exported)).toHaveLength(0);
				expect(content(exported)).toEqual(content(native));
				if (name === 'picture') {
					const before = await JSZip.loadAsync(sourceBytes);
					const after = await JSZip.loadAsync(bytes);
					expect(await after.file('word/media/image1.png')!.async('uint8array')).toEqual(
						await before.file('word/media/image1.png')!.async('uint8array'),
					);
					const xml = await after.file('word/document.xml')!.async('string');
					expect(xml).toContain('noProof');
					expect(drawing(xml)).toBe(
						drawing(await before.file('word/document.xml')!.async('string')),
					);
				}
				const view = (editor as unknown as { view: EditorView }).view;
				expect(undo(view.state, view.dispatch)).toBe(true);
				expect(editor.documentModel).toEqual(source);
			} finally {
				editor.remove();
			}
		});
