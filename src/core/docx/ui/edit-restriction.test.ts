import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection } from 'prosemirror-state';
import type { DocumentProtection, PermissionRange } from '../model';
import {
	editRestrictionOf,
	editRestrictionPlugin,
	editableRanges,
	forcesTrackChanges,
	locksStories,
	setEditableRegion,
} from './edit-restriction';

const schema = new Schema({
	nodes: {
		doc: {
			content: 'paragraph+',
			attrs: { trackChanges: { default: false }, commentThreads: { default: '[]' } },
		},
		paragraph: {
			content: 'text*',
			attrs: { id: { default: '' }, permissions: { default: null } },
			toDOM: () => ['p', 0],
		},
		text: {},
	},
	marks: { comment: { attrs: { ids: { default: [] } } }, bold: {} },
});

const everyone = (from: number, to: number, id = '1'): PermissionRange => ({
	id,
	group: 'everyone',
	from,
	to,
});

/** "hello world" with "world" editable by everyone, and a second paragraph that is not editable. */
function start(protection: DocumentProtection | undefined, user = 'Ann') {
	const doc = schema.node('doc', null, [
		schema.node(
			'paragraph',
			{ id: 'a', permissions: [everyone(6, 11)] },
			schema.text('hello world'),
		),
		schema.node('paragraph', { id: 'b' }, schema.text('locked')),
	]);
	return EditorState.create({
		doc,
		schema,
		plugins: [
			editRestrictionPlugin(
				() => protection,
				() => user,
			),
		],
	});
}
const protectedBy = (edit: DocumentProtection['edit']): DocumentProtection => ({
	edit,
	enforced: true,
});
/** Position of the character at offset `n` of the first paragraph. */
const at = (n: number) => 1 + n;
const textOf = (state: EditorState) => state.doc.textContent;

describe('edit restriction guard', () => {
	it('knows which protections it enforces', () => {
		expect(editRestrictionOf(protectedBy('readOnly'))).toBe('readOnly');
		expect(editRestrictionOf(protectedBy('forms'))).toBeUndefined();
		expect(editRestrictionOf({ edit: 'readOnly', enforced: false })).toBeUndefined();
		expect(forcesTrackChanges(protectedBy('trackedChanges'))).toBe(true);
		expect(locksStories(protectedBy('comments'))).toBe(true);
		expect(locksStories(protectedBy('trackedChanges'))).toBe(false);
	});

	it('read only: blocks edits outside editable ranges and allows them inside', () => {
		const state = start(protectedBy('readOnly'));
		expect(textOf(state.applyTransaction(state.tr.insertText('X', at(1))).state)).toBe(
			textOf(state),
		);
		expect(textOf(state.applyTransaction(state.tr.delete(14, 16)).state)).toBe(textOf(state));
		const typed = state.applyTransaction(state.tr.insertText('big ', at(6))).state;
		expect(textOf(typed)).toBe('hello big worldlocked');
		const deleted = state.applyTransaction(state.tr.delete(at(6), at(11))).state;
		expect(textOf(deleted)).toBe('hello locked');
	});

	it('read only: the editable range follows the text, and a range cannot be left by an edit', () => {
		let state = start(protectedBy('readOnly'));
		state = state.applyTransaction(state.tr.insertText('ab', at(8))).state;
		expect(editableRanges(state, 'Ann')).toEqual([
			expect.objectContaining({ from: at(6), to: at(13) }),
		]);
		expect(state.doc.firstChild!.attrs.permissions).toEqual([everyone(6, 13)]);
		const across = state.tr.delete(at(3), at(8));
		expect(textOf(state.applyTransaction(across).state)).toBe(textOf(state));
	});

	it('read only: only the exceptions of the current user count', () => {
		const doc = schema.node('doc', null, [
			schema.node(
				'paragraph',
				{ permissions: [{ id: '1', user: 'bob@example.com', from: 0, to: 5 }] },
				schema.text('hello'),
			),
		]);
		const make = (user: string) =>
			EditorState.create({
				doc,
				plugins: [
					editRestrictionPlugin(
						() => protectedBy('readOnly'),
						() => user,
					),
				],
			});
		const edit = (state: EditorState) =>
			textOf(state.applyTransaction(state.tr.insertText('X', 1)).state);
		expect(edit(make('Bob@example.com'))).toBe('Xhello');
		expect(edit(make('Ann'))).toBe('hello');
	});

	it('comments only: allows comment marks and threads but not text edits', () => {
		const state = start(protectedBy('comments'));
		const mark = schema.marks.comment!.create({ ids: ['c1'] });
		const commented = state.applyTransaction(state.tr.addMark(at(0), at(5), mark));
		expect(commented.state.doc.firstChild!.firstChild!.marks).toHaveLength(1);
		const thread = state.applyTransaction(
			state.tr.setDocAttribute('commentThreads', '[{"id":"c1"}]'),
		);
		expect(thread.state.doc.attrs.commentThreads).toBe('[{"id":"c1"}]');
		const bold = state.applyTransaction(
			state.tr.addMark(at(0), at(5), schema.marks.bold!.create()),
		);
		expect(bold.state.doc.firstChild!.firstChild!.marks).toHaveLength(0);
		expect(textOf(state.applyTransaction(state.tr.insertText('X', at(1))).state)).toBe(
			textOf(state),
		);
		expect(textOf(state.applyTransaction(state.tr.insertText('X', at(7))).state)).toContain(
			'wXorld',
		);
	});

	it('tracked changes only: keeps Track Changes from being turned off', () => {
		let state = start(protectedBy('trackedChanges'));
		state = state.applyTransaction(state.tr.setDocAttribute('trackChanges', true)).state;
		expect(state.doc.attrs.trackChanges).toBe(true);
		expect(
			state.applyTransaction(state.tr.setDocAttribute('trackChanges', false)).state.doc.attrs
				.trackChanges,
		).toBe(true);
		expect(textOf(state.applyTransaction(state.tr.insertText('X', at(1))).state)).toContain(
			'hXello',
		);
	});

	it('does nothing while no restriction is enforced or for selection-only changes', () => {
		const state = start(undefined);
		expect(textOf(state.applyTransaction(state.tr.insertText('X', at(1))).state)).toContain(
			'hXello',
		);
		const guarded = start(protectedBy('readOnly'));
		const moved = guarded.applyTransaction(
			guarded.tr.setSelection(TextSelection.create(guarded.doc, 2)),
		).state;
		expect(moved.selection.from).toBe(2);
	});
});

describe('exceptions', () => {
	it('adds and removes an editable region for the selection', () => {
		const state = start(undefined);
		const added = setEditableRegion(state, true, at(0), at(5));
		if (!('tr' in added)) throw new Error('expected a transaction');
		const after = state.applyTransaction(added.tr).state;
		expect(after.doc.firstChild!.attrs.permissions).toEqual([
			{ id: '2', group: 'everyone', from: 0, to: 5 },
			everyone(6, 11),
		]);
		const removed = setEditableRegion(after, false, at(8), at(10));
		if (!('tr' in removed)) throw new Error('expected a transaction');
		const split = after.applyTransaction(removed.tr).state;
		expect(split.doc.firstChild!.attrs.permissions).toEqual([
			{ id: '2', group: 'everyone', from: 0, to: 5 },
			{ id: '1', group: 'everyone', from: 6, to: 8 },
			{ id: '3', group: 'everyone', from: 10, to: 11 },
		]);
	});

	it('leaves a range that goes on into another paragraph alone', () => {
		const doc = schema.node('doc', null, [
			schema.node(
				'paragraph',
				{ permissions: [{ ...everyone(0, 5), toNext: true }] },
				schema.text('hello'),
			),
			schema.node('paragraph', null, schema.text('more')),
		]);
		const state = EditorState.create({
			doc,
			plugins: [
				editRestrictionPlugin(
					() => undefined,
					() => 'Ann',
				),
			],
		});
		expect(setEditableRegion(state, false, 2, 4)).toEqual({ error: 'spansParagraphs' });
		expect(setEditableRegion(state, true, 3, 3)).toEqual({ error: 'empty' });
	});
});
