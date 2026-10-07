import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { createDocument, loadDocx, rejectAllRevisions, type Paragraph } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { modelToDoc, docToModel } from './model-adapter';

for (const name of ['alignment', 'spacing', 'indent', 'multiple'])
	it(`retains native ${name} paragraph history through a text transaction`, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					resolve('../core/docx/__fixtures__/review-paragraph-formatting', `${name}-tracked.docx`),
				),
			),
		);
		let state = EditorState.create({ doc: modelToDoc(loaded.model) });
		expect(docToModel(state.doc, loaded.model).blocks).toEqual(loaded.model.blocks);
		state = state.apply(state.tr.insertText('!', 5));
		const model = docToModel(state.doc, loaded.model);
		expect((model.blocks[0] as Paragraph).formatRevision).toEqual(
			(loaded.model.blocks[0] as Paragraph).formatRevision,
		);
		const reopened = await loadDocx(await loaded.save(model));
		expect((reopened.model.blocks[0] as Paragraph).formatRevision).toEqual(
			(loaded.model.blocks[0] as Paragraph).formatRevision,
		);
	});

describe('paragraph mark revision preservation', () => {
	for (const kind of ['insert', 'delete'] as const)
		it(`keeps the ${kind} paragraph mark when its text changes`, () => {
			const model = createDocument();
			model.blocks = [
				{
					type: 'paragraph',
					id: 'p',
					runs: [{ text: 'Text' }],
					markRevision: { kind, id: 'mark', author: 'Ada' },
				},
			];
			let state = EditorState.create({ doc: modelToDoc(model) });
			state = state.apply(state.tr.insertText('!', 3));
			expect((docToModel(state.doc, model).blocks[0] as Paragraph).markRevision).toEqual(
				(model.blocks[0] as Paragraph).markRevision,
			);
		});
});

for (const name of ['alignment', 'spacing', 'indent', 'multiple'])
	it(`retains restored native ${name} paragraph properties through editor text editing`, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				await readFile(
					resolve('../core/docx/__fixtures__/review-paragraph-formatting', `${name}-tracked.docx`),
				),
			),
		);
		const model = rejectAllRevisions(loaded.model);
		let state = EditorState.create({ doc: modelToDoc(model) });
		expect(docToModel(state.doc, model).blocks).toEqual(model.blocks);
		state = state.apply(state.tr.insertText('!', 5));
		const actual = docToModel(state.doc, model);
		expect((actual.blocks[0] as Paragraph).restoredParagraphPropertiesXml).toBe(
			(model.blocks[0] as Paragraph).restoredParagraphPropertiesXml,
		);
		const reopened = await loadDocx(await loaded.save(actual));
		const expected = {
			...(model.blocks[0] as Paragraph),
			runs: (reopened.model.blocks[0] as Paragraph).runs,
		};
		delete expected.restoredParagraphPropertiesXml;
		expect(reopened.model.blocks[0]).toEqual(expected);
	});
