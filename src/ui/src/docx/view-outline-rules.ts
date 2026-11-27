import type { DocumentModel } from 'ooxml-core/docx';
import {
	canChangeOutlineLevel,
	changeOutlineLevel,
	forcesTrackChanges,
	moveOutline,
	OUTLINE_BODY,
	outlineParagraphs,
	outlinePeer,
	setOutlineLevel,
	trackedMoveOutline,
	trackedOutlineMoveRefusal,
	type OutlineEntry,
	type TrackedOutlineMoveRefusal,
} from 'ooxml-core/docx/ui';
import type { EditorState, Transaction } from 'prosemirror-state';
import { DocAttrStep } from 'prosemirror-transform';
import type { LocalizationKey } from './localization';
import type { OutlineKey } from './view-outline-ribbon';

/** How Outline's structural edits are recorded: tracked, shared with a room, or both. */
export interface OutlineEditContext {
	editable: boolean;
	model: DocumentModel;
	collaborating: boolean;
	author: string;
	revisionIds?: ((kind: string) => string) | undefined;
}

export interface OutlineAvailability {
	enabled: Set<OutlineKey>;
	reasons: Map<OutlineKey, LocalizationKey>;
	/** Levels the Outline Level list may apply to the selected paragraphs. */
	levels: Set<number>;
	paragraphLevel: number | undefined;
}

const MOVE_REASONS: Record<TrackedOutlineMoveRefusal, LocalizationKey> = {
	content: 'outline.reason.trackedMoveContent',
	revisions: 'outline.reason.trackedMoveRevisions',
	boundary: 'outline.reason.trackedMoveBoundary',
};

export const tracksOutlineEdits = (model: DocumentModel): boolean =>
	Boolean(model.trackChanges) || forcesTrackChanges(model.protection);

/** A protected document allows Outline edits only when the protection is Track Changes. */
function canEdit(context: OutlineEditContext): boolean {
	const { protection } = context.model;
	return context.editable && (!protection?.enforced || forcesTrackChanges(protection));
}

/**
 * Style catalogs stay on each device in a collaboration session, so a level change that has to
 * define a new heading style is refused there; applying an existing style is an ordinary step.
 */
function sharedLevel(context: OutlineEditContext, tr: Transaction | undefined): boolean {
	return Boolean(
		tr && !(context.collaborating && tr.steps.some((step) => step instanceof DocAttrStep)),
	);
}

export function outlineLevelEdit(
	state: EditorState,
	context: OutlineEditContext,
	paragraphs: readonly OutlineEntry[],
	change: { direction: -1 | 1 } | { level: number },
): Transaction | undefined {
	if (!canEdit(context)) return undefined;
	const tr =
		'level' in change
			? setOutlineLevel(state, context.model, paragraphs, change.level)
			: changeOutlineLevel(state, context.model, paragraphs, change.direction);
	return sharedLevel(context, tr) ? tr : undefined;
}

/** Moves a branch, as a tracked move when Track Changes is on. */
export function outlineMoveEdit(
	state: EditorState,
	context: OutlineEditContext,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry,
	direction: -1 | 1,
): Transaction | undefined {
	if (!canEdit(context)) return undefined;
	if (!tracksOutlineEdits(context.model)) return moveOutline(state, entries, entry, direction);
	return trackedMoveOutline(state, entries, entry, direction, {
		author: context.author || 'Author',
		...(context.revisionIds ? { nextId: context.revisionIds } : {}),
	});
}

export function outlineAvailability(
	state: EditorState,
	context: OutlineEditContext,
	entries: readonly OutlineEntry[],
	entry: OutlineEntry | undefined,
	collapsed: ReadonlySet<string>,
): OutlineAvailability {
	const enabled = new Set<OutlineKey>(['firstLine', 'format', 'close']);
	const reasons = new Map<OutlineKey, LocalizationKey>();
	const levels = new Set<number>();
	const result: OutlineAvailability = { enabled, reasons, levels, paragraphLevel: undefined };
	if (!entry) return result;
	if (entry.end > entry.index + 1) enabled.add(collapsed.has(entry.id) ? 'expand' : 'collapse');
	if (!canEdit(context)) return result;
	const paragraphs = outlineParagraphs(state, entries);
	if (paragraphs.length) {
		const first = paragraphs[0]?.level;
		if (paragraphs.every((paragraph) => paragraph.level === first)) result.paragraphLevel = first;
		for (let level = 1; level <= OUTLINE_BODY; level++)
			if (
				paragraphs.every((paragraph) => paragraph.level === level) ||
				outlineLevelEdit(state, context, paragraphs, { level })
			)
				levels.add(level);
		if (levels.size) enabled.add('paragraphLevel');
	}
	for (const [key, direction] of [
		['promote', -1],
		['demote', 1],
	] as const) {
		if (!paragraphs.some((paragraph) => canChangeOutlineLevel(paragraph, direction))) continue;
		if (outlineLevelEdit(state, context, paragraphs, { direction })) enabled.add(key);
		else reasons.set(key, 'outline.reason.sharedStyle');
	}
	const end = entries[entry.end]?.pos ?? state.doc.content.size;
	if (state.selection.to >= end) return result;
	for (const [key, direction] of [
		['up', -1],
		['down', 1],
	] as const) {
		if (!outlinePeer(entries, entry, direction)) continue;
		const refusal = tracksOutlineEdits(context.model)
			? trackedOutlineMoveRefusal(state, entries, entry, direction)
			: undefined;
		if (refusal) reasons.set(key, MOVE_REASONS[refusal]);
		else enabled.add(key);
	}
	return result;
}
