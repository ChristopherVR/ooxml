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
		doc: { content: 'paragraph+', attrs: { styleCatalogs: { default: null } } },
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
	const promoted = initial.apply(changeOutlineLevel(initial, model, [entries[1]!], -1)!);
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

it('switches built-in styles in both directions and creates missing body and heading styles', async () => {
	let state = stateFrom(model);
	const applyLevel = (index: number, direction: -1 | 1) => {
		const entries = indexOutline(state.doc, model);
		state = state.apply(
			changeOutlineLevel(
				state,
				{
					...model,
					...state.doc.attrs.styleCatalogs,
				},
				[entries[index]!],
				direction,
			)!,
		);
	};
	applyLevel(0, 1);
	expect(state.doc.child(0).attrs).toMatchObject({ style: 'Heading2', outlineLevel: null });
	applyLevel(0, -1);
	expect(state.doc.child(0).attrs).toMatchObject({ style: 'Heading1', outlineLevel: null });
	applyLevel(2, -1);
	expect(state.doc.child(2).attrs).toMatchObject({ style: 'Heading9', outlineLevel: null });
	expect(indexOutline(state.doc, model)[2]?.level).toBe(9);
	applyLevel(2, 1);
	expect(state.doc.child(2).attrs).toMatchObject({ style: 'BodyText', outlineLevel: null });
	const blocks = model.blocks.map((block, index) =>
		block.type === 'paragraph'
			? paragraphFromAttrs(state.doc.child(index).attrs, block.id, block.runs)
			: block,
	);
	const saved = await loadDocx(
		await saveDocx({ ...model, ...state.doc.attrs.styleCatalogs, blocks }),
	);
	expect(saved.model.blocks[2]).toMatchObject({ style: 'BodyText' });
	expect(saved.model.paragraphStyles?.styles.BodyText?.formatting.outlineLevel).toBe(0);
	expect(saved.model.paragraphStyles?.styles.Heading9?.formatting.outlineLevel).toBe(9);
});
