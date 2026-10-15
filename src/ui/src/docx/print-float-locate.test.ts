// @vitest-environment jsdom
import { createDocument, type DocumentModel, type InlineImage } from 'ooxml-core/docx';
import { EditorState, NodeSelection } from 'prosemirror-state';
import { describe, expect, it } from 'vitest';
import { modelToDoc } from './model-adapter';
import { locateFloat, selectedFloat } from './print-float-locate';

const picture = (partName: string, anchored: boolean): InlineImage => ({
	relId: 'rId1',
	partName,
	contentType: 'image/png',
	widthPx: 40,
	heightPx: 20,
	...(anchored ? { anchored: true, placement: { wrap: 'square' as const, offsetXPx: 4 } } : {}),
});

function model(): DocumentModel {
	return {
		...createDocument(),
		blocks: [
			{
				type: 'paragraph',
				id: 'p1',
				runs: [
					{ text: 'a' },
					{ text: '', image: picture('word/media/inline.png', false) },
					{ text: '', image: picture('word/media/first.png', true) },
					{ text: 'b' },
					{ text: '', image: picture('word/media/second.png', true) },
				],
			},
		],
	};
}

describe('floating picture lookup', () => {
	it('counts only anchored pictures, in run order, like the layout does', () => {
		const doc = modelToDoc(model());
		expect(locateFloat(doc, 'p1', 0)?.node.attrs.partName).toBe('word/media/first.png');
		expect(locateFloat(doc, 'p1', 1)?.node.attrs.partName).toBe('word/media/second.png');
		expect(locateFloat(doc, 'p1', 2)).toBeUndefined();
		expect(locateFloat(doc, 'missing', 0)).toBeUndefined();
	});

	it('reads the picture a node selection holds back as paragraph and index', () => {
		const doc = modelToDoc(model());
		const second = locateFloat(doc, 'p1', 1)!;
		const state = EditorState.create({ doc, selection: NodeSelection.create(doc, second.pos) });
		expect(selectedFloat(state)).toMatchObject({ blockId: 'p1', index: 1, pos: second.pos });
		const inline = NodeSelection.create(doc, locateFloat(doc, 'p1', 0)!.pos - 2);
		expect(selectedFloat(EditorState.create({ doc, selection: inline }))).toBeUndefined();
	});
});
