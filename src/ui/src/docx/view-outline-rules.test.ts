// @vitest-environment jsdom
import { createDocument, type DocumentModel } from 'ooxml-core/docx';
import { indexOutline, outlineParagraphs } from 'ooxml-core/docx/ui';
import { EditorState, TextSelection } from 'prosemirror-state';
import { describe, expect, it } from 'vitest';
import { schema } from './schema';
import {
	outlineAvailability,
	outlineLevelEdit,
	outlineMoveEdit,
	type OutlineEditContext,
} from './view-outline-rules';

function state(texts: [string, number | null][]) {
	const doc = schema.node(
		'doc',
		null,
		texts.map(([text, outlineLevel], index) =>
			schema.node('paragraph', { id: `p${index}`, outlineLevel }, schema.text(text)),
		),
	);
	const created = EditorState.create({ doc });
	return created.apply(created.tr.setSelection(TextSelection.create(created.doc, 2)));
}
const outline: [string, number | null][] = [
	['Alpha', 1],
	['Alpha body', null],
	['Beta', 1],
	// A heading of its own: body text would belong to Beta's branch and end the document there,
	// where a tracked move cannot land (see outline-tracked.test.ts in the core).
	['Tail', 1],
];
const context = (model: Partial<DocumentModel>, collaborating = false): OutlineEditContext => ({
	editable: true,
	model: { ...createDocument(), ...model },
	collaborating,
	author: 'Ada',
});
const keys = (value: ReturnType<typeof outlineAvailability>) => [...value.enabled].sort();

describe('Outline structural edits', () => {
	it('stay available while changes are tracked and record the move', () => {
		const current = state(outline);
		const tracked = context({ trackChanges: true });
		const entries = indexOutline(current.doc, tracked.model);
		const available = outlineAvailability(current, tracked, entries, entries[0], new Set());
		expect(keys(available)).toEqual(
			['close', 'collapse', 'demote', 'down', 'firstLine', 'format', 'paragraphLevel'].sort(),
		);
		const tr = outlineMoveEdit(current, tracked, entries, entries[0]!, 1);
		const kinds: unknown[] = [];
		tr!.doc.forEach((node) => kinds.push(node.attrs.markRevision?.kind ?? null));
		expect(kinds).toEqual(['moveFrom', 'moveFrom', null, 'moveTo', 'moveTo', null]);
	});

	it('explains a move that cannot be tracked', () => {
		const current = state(outline.slice(0, 3));
		const tracked = context({ trackChanges: true });
		const entries = indexOutline(current.doc, tracked.model);
		const available = outlineAvailability(current, tracked, entries, entries[0], new Set());
		expect(available.enabled.has('down')).toBe(false);
		expect(available.reasons.get('down')).toBe('outline.reason.trackedMoveBoundary');
		// Without tracking, the same move is an ordinary reorder.
		const plain = context({});
		expect(
			outlineAvailability(current, plain, entries, entries[0], new Set()).enabled.has('down'),
		).toBe(true);
	});

	it('works in collaboration unless a level change must define a new heading style', () => {
		const current = state(outline);
		const shared = context({}, true);
		const entries = indexOutline(current.doc, shared.model);
		const paragraphs = outlineParagraphs(current, entries);
		// Direct outline levels change in place, which is an ordinary shared step.
		expect(outlineLevelEdit(current, shared, paragraphs, { direction: 1 })).toBeDefined();
		expect(outlineMoveEdit(current, shared, entries, entries[0]!, 1)).toBeDefined();

		const body = current.apply(current.tr.setSelection(TextSelection.create(current.doc, 9)));
		const bodyParagraphs = outlineParagraphs(body, entries);
		expect(outlineLevelEdit(body, shared, bodyParagraphs, { direction: -1 })).toBeUndefined();
		const available = outlineAvailability(body, shared, entries, entries[1], new Set());
		expect(available.enabled.has('promote')).toBe(false);
		expect(available.reasons.get('promote')).toBe('outline.reason.sharedStyle');
		expect(outlineLevelEdit(body, context({}), bodyParagraphs, { direction: -1 })).toBeDefined();
	});

	it('stays off in a document protected for anything but tracked changes', () => {
		const current = state(outline);
		const forms = context({ protection: { enforced: true, edit: 'forms' } });
		const entries = indexOutline(current.doc, forms.model);
		expect(keys(outlineAvailability(current, forms, entries, entries[0], new Set()))).toEqual(
			['close', 'collapse', 'firstLine', 'format'].sort(),
		);
		const tracked = context({ protection: { enforced: true, edit: 'trackedChanges' } });
		expect(
			outlineAvailability(current, tracked, entries, entries[0], new Set()).enabled.has('down'),
		).toBe(true);
	});
});
