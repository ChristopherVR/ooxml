import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { modelToDoc, docToModel } from './model-adapter';

for (const name of ['bold', 'multiple'])
	describe(`editor ${name} formatting history`, () => {
		it('keeps the formatting marker and its prior snapshot through a text transaction', async () => {
			const bytes = await readFile(
				resolve('../core/docx/__fixtures__/review-formatting', `${name}-tracked.docx`),
			);
			const loaded = await loadDocx(new Uint8Array(bytes));
			let state = EditorState.create({ doc: modelToDoc(loaded.model) });
			const original = loaded.model.blocks[0]!;
			if (original.type !== 'paragraph') throw new Error('Expected paragraph');
			const revision = original.runs.find((run) => run.revision)?.revision;
			expect(docToModel(state.doc, loaded.model).blocks).toEqual(loaded.model.blocks);
			state = state.apply(state.tr.insertText('!', 5));
			const model = docToModel(state.doc, loaded.model);
			const actual = model.blocks[0]!;
			if (actual.type !== 'paragraph') throw new Error('Expected paragraph');
			expect(actual.runs.find((run) => run.revision)?.revision).toEqual(revision);
			const reopened = await loadDocx(await loaded.save(model));
			const paragraph = reopened.model.blocks[0]!;
			if (paragraph.type !== 'paragraph') throw new Error('Expected paragraph');
			expect(paragraph.runs.find((run) => run.revision)?.revision).toEqual(revision);
		});
	});
