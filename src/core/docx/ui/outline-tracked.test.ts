import { describe, expect, it } from 'vitest';
import { Schema } from 'prosemirror-model';
import { EditorState, TextSelection, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { history } from 'prosemirror-history';
import { createDocument, type DocumentModel, type Paragraph } from '../index';
import { changeOutlineLevel, indexOutline, outlineParagraphs } from './outline';
import { trackedMoveOutline, trackedOutlineMoveRefusal } from './outline-tracked';
import { paragraphAttrs } from './paragraph-attributes';
import { acceptRevisionRange, collectRevisionRanges, rejectRevisionRange } from './review-commands';
import { paragraphFormattingRevision } from './review-paragraph-formatting';
import { reviewMarks } from './review-schema';
import { markSpecs } from './schema-marks';
import { trackChangesPlugin } from './track-changes-mode';

const empty: Paragraph = { type: 'paragraph', id: 'empty', runs: [] };
const schema = new Schema({
	nodes: {
		doc: {
			content: 'paragraph+',
			attrs: {
				styleCatalogs: { default: null },
				trackMoves: { default: true },
				trackFormatting: { default: true },
			},
		},
		paragraph: {
			content: 'text*',
			attrs: Object.fromEntries(
				Object.entries(paragraphAttrs(empty)).map(([key, value]) => [key, { default: value }]),
			),
		},
		text: {},
	},
	marks: { ...markSpecs, insertion: reviewMarks.insertion, deletion: reviewMarks.deletion },
});

const paragraphs: Paragraph[] = [
	{ type: 'paragraph', id: 'a', outlineLevel: 1, runs: [{ text: 'Alpha' }] },
	{ type: 'paragraph', id: 'a-body', runs: [{ text: 'Alpha body' }] },
	{ type: 'paragraph', id: 'b', outlineLevel: 1, runs: [{ text: 'Beta' }] },
	// A heading of its own: body text would belong to Beta's branch and end the document there.
	{ type: 'paragraph', id: 'tail', outlineLevel: 1, runs: [{ text: 'Tail' }] },
];
const model: DocumentModel = { ...createDocument(), blocks: paragraphs };

function editor(blocks = paragraphs, trackMoves = true): EditorView {
	const doc = schema.node(
		'doc',
		{ trackMoves },
		blocks.map((paragraph) =>
			schema.node(
				'paragraph',
				paragraphAttrs(paragraph),
				paragraph.runs.map((run) => schema.text(run.text)),
			),
		),
	);
	const host = {
		state: EditorState.create({
			doc,
			plugins: [
				history(),
				trackChangesPlugin(
					() => 'Ada',
					() => true,
				),
			],
		}),
		editable: true,
		dispatch(tr: Transaction) {
			this.state = this.state.apply(tr);
		},
		focus() {},
	};
	return host as unknown as EditorView;
}
const texts = (view: EditorView) => {
	const result: string[] = [];
	view.state.doc.forEach((node) => result.push(node.textContent));
	return result;
};
function moveFirstDown(view: EditorView) {
	view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 2)));
	const entries = indexOutline(view.state.doc, model);
	const tr = trackedMoveOutline(view.state, entries, entries[0]!, 1, { author: 'Ada' });
	expect(tr).toBeDefined();
	view.dispatch(tr!);
}

describe('tracked outline moves', () => {
	it('records a branch move as linked moveFrom and moveTo text and paragraph marks', () => {
		const view = editor();
		moveFirstDown(view);
		expect(texts(view)).toEqual(['Alpha', 'Alpha body', 'Beta', 'Alpha', 'Alpha body', 'Tail']);
		const marks: unknown[] = [];
		view.state.doc.forEach((node) => marks.push(node.attrs.markRevision?.kind ?? null));
		expect(marks).toEqual(['moveFrom', 'moveFrom', null, 'moveTo', 'moveTo', null]);
		const ranges = collectRevisionRanges(view.state.doc);
		const names = new Set(ranges.map((range) => range.move));
		expect(names.size).toBe(1);
		expect([...names][0]).toMatch(/^move/);
		const ids = new Set<string>();
		view.state.doc.forEach((node) => ids.add(String(node.attrs.id)));
		expect(ids.size).toBe(6);
		// The caret follows the branch to its new place.
		expect(view.state.selection.$from.index(0)).toBe(3);
	});

	it('accepting the move keeps only the moved copy, rejecting restores the original order', () => {
		const accepted = editor();
		moveFirstDown(accepted);
		acceptRevisionRange(accepted, collectRevisionRanges(accepted.state.doc)[0]!);
		expect(texts(accepted)).toEqual(['Beta', 'Alpha', 'Alpha body', 'Tail']);
		expect(collectRevisionRanges(accepted.state.doc)).toEqual([]);

		const rejected = editor();
		moveFirstDown(rejected);
		rejectRevisionRange(rejected, collectRevisionRanges(rejected.state.doc)[0]!);
		expect(texts(rejected)).toEqual(['Alpha', 'Alpha body', 'Beta', 'Tail']);
		expect(collectRevisionRanges(rejected.state.doc)).toEqual([]);
	});

	it('records a deletion and an insertion when move tracking is off', () => {
		const view = editor(paragraphs, false);
		moveFirstDown(view);
		const ranges = collectRevisionRanges(view.state.doc);
		expect(ranges.every((range) => range.move === undefined)).toBe(true);
		const marks: unknown[] = [];
		view.state.doc.forEach((node) => marks.push(node.attrs.markRevision?.kind ?? null));
		expect(marks).toEqual(['delete', 'delete', null, 'insert', 'insert', null]);
	});

	it('refuses branches it cannot track honestly', () => {
		const last = editor(paragraphs.slice(0, 3));
		const entries = indexOutline(last.state.doc, model);
		// Moving Alpha down would leave its copy as the last paragraph of the document.
		expect(trackedOutlineMoveRefusal(last.state, entries, entries[0]!, 1)).toBe('boundary');
		expect(trackedOutlineMoveRefusal(last.state, entries, entries[2]!, -1)).toBe('boundary');

		const moved = editor();
		moveFirstDown(moved);
		const after = indexOutline(moved.state.doc, model);
		expect(trackedOutlineMoveRefusal(moved.state, after, after[0]!, 1)).toBe('revisions');

		const free = editor();
		const fresh = indexOutline(free.state.doc, model);
		expect(trackedOutlineMoveRefusal(free.state, fresh, fresh[0]!, 1)).toBeUndefined();
	});
});

describe('tracked outline levels', () => {
	it('records promote and demote as paragraph property changes, defining a missing style', () => {
		const view = editor([
			{ type: 'paragraph', id: 'h', style: 'Heading1', runs: [{ text: 'Heading' }] },
			{ type: 'paragraph', id: 'p', runs: [{ text: 'Body' }] },
		]);
		const headings: DocumentModel = {
			...createDocument(),
			blocks: [],
		};
		view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 12)));
		const entries = indexOutline(view.state.doc, headings);
		const tr = changeOutlineLevel(view.state, headings, outlineParagraphs(view.state, entries), -1);
		expect(tr).toBeDefined();
		view.dispatch(tr!);
		const body = view.state.doc.child(1);
		expect(String(body.attrs.style)).toMatch(/^Heading/);
		expect(paragraphFormattingRevision(body)?.author).toBe('Ada');
	});
});
