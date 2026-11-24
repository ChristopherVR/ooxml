import { describe, expect, it } from 'vitest';
import { createDocument } from 'ooxml-core/docx';
import { EditorState, TextSelection } from 'prosemirror-state';
import { modelToDoc } from '../model-adapter';
import { selectionModel } from './selection-model';

function state() {
	const model = createDocument();
	model.blocks = [
		{ type: 'paragraph', id: 'a', runs: [{ text: 'First paragraph' }] },
		{ type: 'paragraph', id: 'b', style: 'Heading1', runs: [{ text: 'Second', bold: true }] },
		{ type: 'paragraph', id: 'c', runs: [{ text: 'Third paragraph' }] },
	];
	return { model, state: EditorState.create({ doc: modelToDoc(model) }) };
}
const text = (blocks: readonly { type: string; runs?: readonly { text: string }[] }[]) =>
	blocks.map((block) => block.runs?.map((run) => run.text).join('') ?? '');

describe('selectionModel', () => {
	it('keeps only the selected blocks, cut at the selection ends, with their formatting', () => {
		const { model, state: initial } = state();
		const doc = initial.doc;
		// From "paragraph" in the first block to "Thi" in the third.
		const from = 1 + 'First '.length;
		const to = doc.child(0).nodeSize + doc.child(1).nodeSize + 1 + 'Thi'.length;
		const selected = selectionModel(
			initial.apply(initial.tr.setSelection(TextSelection.create(doc, from, to))),
			model,
		)!;
		expect(text(selected.blocks)).toEqual(['paragraph', 'Second', 'Thi']);
		expect(selected.blocks[1]).toMatchObject({ style: 'Heading1', runs: [{ bold: true }] });
		expect(selected.page).toEqual(model.page);
		expect(text(model.blocks)).toEqual(['First paragraph', 'Second', 'Third paragraph']);
	});

	it('is undefined for an empty selection', () => {
		const { model, state: initial } = state();
		expect(selectionModel(initial, model)).toBeUndefined();
	});
});
