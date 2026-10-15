// @vitest-environment jsdom
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { afterEach, describe, expect, it } from 'vitest';
import { modelToDoc } from './model-adapter';
import { noteParagraphs, referencePosition } from './print-note-editing';

function model(): DocumentModel {
	const base = createDocument();
	return {
		...base,
		blocks: [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{ text: 'Claim' },
					{ text: '', noteReference: { kind: 'footnote', id: '2' } },
					{ text: ' tail' },
				],
			},
		],
		footnotes: [
			{
				id: '2',
				blocks: [
					{
						type: 'paragraph',
						id: 'fn-2-p0',
						runs: [{ text: '', noteMark: 'footnote' }, { text: 'Source' }],
					},
				],
			},
		],
		endnotes: [
			{ id: '3', blocks: [{ type: 'paragraph', id: 'en-3-p0', runs: [{ text: 'End' }] }] },
		],
	};
}

let view: EditorView | undefined;
afterEach(() => {
	view?.destroy();
	view = undefined;
});

describe('page note editing helpers', () => {
	it('maps every note paragraph to its note', () => {
		const map = noteParagraphs(model());
		expect(map.get('fn-2-p0')).toEqual({ kind: 'footnote', id: '2', firstId: 'fn-2-p0' });
		expect(map.get('en-3-p0')).toEqual({ kind: 'endnote', id: '3', firstId: 'en-3-p0' });
		expect(map.get('p1')).toBeUndefined();
	});

	it('finds the caret position just after a body reference', () => {
		const doc = modelToDoc(model());
		view = new EditorView(document.createElement('div'), { state: EditorState.create({ doc }) });
		const reference = referencePosition(view, { kind: 'footnote', id: '2' });
		expect(reference).toMatchObject({ blockId: 'p1', offset: 6 });
		expect(view.state.doc.textBetween(reference!.pos, reference!.pos + 5)).toBe(' tail');
		expect(referencePosition(view, { kind: 'endnote', id: '3' })).toBeUndefined();
	});
});
