import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { expect, it } from 'vitest';
import { createDocument, loadDocx, saveDocx, type DocumentModel } from '../index';
import { paragraphAttrs, paragraphFromAttrs } from './paragraph-attributes';
import {
	changeOutlineLevel,
	indexOutline,
	moveOutline,
	outlinePeer,
	visibleOutline,
} from './outline';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: {
			content: 'text*',
			attrs: { id: { default: null }, style: { default: '' }, outlineLevel: { default: null } },
		},
		text: {},
	},
});

const model: DocumentModel = {
	...createDocument(),
	blocks: [
		{ type: 'paragraph', id: 'first', style: 'Heading1', runs: [{ text: 'First' }] },
		{ type: 'paragraph', id: 'child', outlineLevel: 2, runs: [{ text: 'Nested' }] },
		{ type: 'paragraph', id: 'body', runs: [{ text: 'Body' }] },
		{ type: 'paragraph', id: 'second', outlineLevel: 1, runs: [{ text: 'Second' }] },
	],
};

function stateFrom(model: DocumentModel) {
	const blocks = model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? [
					schema.node(
						'paragraph',
						paragraphAttrs(block),
						schema.text(block.runs[0]?.text ?? block.id),
					),
				]
			: [],
	);
	return EditorState.create({ doc: schema.node('doc', null, blocks), plugins: [history()] });
}

it('indexes style-based and direct headings and hides descendants without editing the document', () => {
	const { doc } = stateFrom(model);
	const entries = indexOutline(doc, model);
	expect(entries.map(({ level, end }) => [level, end])).toEqual([
		[1, 3],
		[2, 3],
		[10, 3],
		[1, 4],
	]);
	expect(visibleOutline(entries, 1, new Set())).toEqual([true, false, false, true]);
	expect(visibleOutline(entries, 10, new Set(['first']))).toEqual([true, false, false, true]);
	expect(doc.textContent).toBe('FirstNestedBodySecond');
});

it('moves a branch as a unit, retaining its anchors, and undoes the move', () => {
	const state = stateFrom(model);
	const entries = indexOutline(state.doc, model);
	const first = entries[0]!;
	const change = moveOutline(state, entries, first, 1)!;
	const next = state.apply(change);
	expect(next.doc.textContent).toBe('SecondFirstNestedBody');
	expect(next.selection.$from.parent.attrs.id).toBe('first');
	expect(indexOutline(next.doc, model).map((entry) => entry.id)).toEqual([
		'second',
		'first',
		'child',
		'body',
	]);
	expect(undo(next, (tr) => expect(next.apply(tr).doc.eq(state.doc)).toBe(true))).toBe(true);

	const sectioned = [...entries.slice(0, 3), { ...entries[3]!, section: 1 }];
	expect(outlinePeer(sectioned, first, 1)).toBeUndefined();
});

it('persists promoted levels and the explicit body-text override through save and reopen', async () => {
	const initial = stateFrom(model);
	const handle = await loadDocx(await saveDocx(model));
	const base = handle.model;
	const entries = indexOutline(initial.doc, model);
	const promoted = initial.apply(changeOutlineLevel(initial, [entries[1]!], -1)!);
	const selected = promoted.apply(promoted.tr.setSelection(TextSelection.create(promoted.doc, 1)));
	expect(selected.doc.child(1).attrs.outlineLevel).toBe(1);
	const blocks = base.blocks.map((block, index) => {
		if (block.type !== 'paragraph') return block;
		const attrs = selected.doc.child(index).attrs;
		return {
			...block,
			...paragraphFromAttrs(attrs, block.id, block.runs),
			...(index === 0 && { outlineLevel: 0 }),
		};
	});
	const saved = await loadDocx(await handle.save({ ...base, blocks }));
	expect(saved.model.blocks[0]).toMatchObject({ outlineLevel: 0 });
	expect(saved.model.blocks[1]).toMatchObject({ outlineLevel: 1 });
});
