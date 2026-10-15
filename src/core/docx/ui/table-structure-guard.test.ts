import { Schema } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';
import { history, undo } from 'prosemirror-history';
import { describe, expect, it } from 'vitest';
import { tableStructureGuardPlugin } from './table-structure-guard';
const schema = new Schema({
	nodes: {
		doc: { content: 'block+' },
		text: { group: 'inline' },
		paragraph: { group: 'block', content: 'inline*', attrs: { id: { default: null } } },
		nestedTablePreview: { group: 'block', atom: true, attrs: { rowsJson: { default: '[]' } } },
		table: {
			group: 'block',
			content: 'tableRow+',
			attrs: { id: { default: 't0' }, structureEditable: { default: false } },
		},
		tableRow: { content: 'tableCell+' },
		tableCell: {
			content: 'block+',
			attrs: {
				sourceCellKey: { default: 'p0' },
				colspan: { default: 1 },
				rowspan: { default: 1 },
				widthTwips: { default: 3200 },
			},
		},
	},
});
const paragraph = (id: string, text: string) => schema.node('paragraph', { id }, schema.text(text));
const create = () =>
	EditorState.create({
		doc: schema.node('doc', null, [
			schema.node('table', null, [
				schema.node('tableRow', null, [
					schema.node('tableCell', null, [
						paragraph('p0', 'Outer'),
						schema.node('nestedTablePreview'),
					]),
				]),
			]),
		]),
		plugins: [history(), tableStructureGuardPlugin()],
	});
describe('complex table transaction boundary', () => {
	it('rejects Enter, multiline paste and removal of the read-only preview before history changes', () => {
		const state = create();
		const transactions = [
			state.tr.split(6),
			state.tr.replaceWith(3, 10, [paragraph('p0', 'One'), paragraph('new', 'Two')]),
			state.tr.delete(10, 11),
			state.tr.setNodeMarkup(0, undefined, { id: 't0', structureEditable: true }),
		];
		for (const tr of transactions) {
			const result = state.applyTransaction(tr);
			expect(result.transactions).toEqual([]);
			expect(result.state.doc).toBe(state.doc);
		}
	});
	it('allows text edits and undo while retaining the source preview', () => {
		const original = create();
		let state = original.applyTransaction(original.tr.insertText('Edit', 4)).state;
		expect(state.doc.textContent).toBe('EditOuter');
		expect(
			undo(state, (tr) => {
				state = state.applyTransaction(tr).state;
			}),
		).toBe(true);
		expect(state.doc.eq(original.doc)).toBe(true);
	});
	it('allows deleting a complete complex table and ordinary table paragraph structure edits', () => {
		const state = create();
		expect(
			state.applyTransaction(state.tr.delete(0, state.doc.firstChild!.nodeSize)).transactions,
		).toHaveLength(1);
		// Recreate a simple table directly: changing a source table's boundary is not a user workflow.
		const simple = EditorState.create({
			doc: schema.node('doc', null, [
				state.doc.firstChild!.type.create(
					{ id: 'simple', structureEditable: true },
					state.doc.firstChild!.content,
				),
			]),
			plugins: [tableStructureGuardPlugin()],
		});
		expect(simple.applyTransaction(simple.tr.split(6)).transactions).toHaveLength(1);
	});
});

describe('nested table transaction boundary', () => {
	const nested = () =>
		schema.node('table', { id: 'n0' }, [
			schema.node('tableRow', null, [
				schema.node('tableCell', { sourceCellKey: 'n1' }, [paragraph('n1', 'Inner')]),
				schema.node('tableCell', { sourceCellKey: 'n2' }, [paragraph('n2', 'Other')]),
			]),
		]);
	const start = () =>
		EditorState.create({
			doc: schema.node('doc', null, [
				schema.node('table', null, [
					schema.node('tableRow', null, [
						schema.node('tableCell', null, [paragraph('p0', 'Outer'), nested()]),
					]),
				]),
			]),
			plugins: [history(), tableStructureGuardPlugin()],
		});
	it('allows text edits inside a nested cell and rejects splitting its paragraphs', () => {
		const state = start();
		let pos = -1;
		state.doc.descendants((node, at) => {
			if (node.isText && node.text === 'Inner') pos = at + 5;
		});
		const edited = state.applyTransaction(state.tr.insertText('!', pos));
		expect(edited.transactions).toHaveLength(1);
		expect(edited.state.doc.textContent).toContain('Inner!');
		const split = state.applyTransaction(state.tr.split(pos));
		expect(split.transactions).toEqual([]);
	});
});
