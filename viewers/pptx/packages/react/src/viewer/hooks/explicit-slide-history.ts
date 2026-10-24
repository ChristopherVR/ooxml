import type { EditorHistoryController } from 'ooxml-ui/pptx';

import type { EditorHistorySnapshot } from '../types';
import { cloneHistorySnapshot } from '../utils/clone';

function historyElementContent(
	element: EditorHistorySnapshot['slides'][number]['elements'][number],
): EditorHistorySnapshot['slides'][number]['elements'][number] {
	return {
		...element,
		// Saving refreshes these archive-only fields in place. They are needed by
		// serialization, but neither one represents a user edit.
		rawXml: undefined,
		shapeId: undefined,
		...(element.type === 'group' ? { children: element.children.map(historyElementContent) } : {}),
	};
}

/**
 * The part of a history snapshot that IS the document.
 *
 * `activeSlideIndex` is deliberately excluded. It rides along in the STORED
 * snapshot so undo/redo return the user to the slide the edit happened on, but
 * it must never take part in deciding whether the deck changed: clicking a
 * thumbnail reassigns nothing but the index, and comparing the whole snapshot
 * made that read as a document mutation. The consequences were both visible to
 * the user - the deck was marked dirty, so autosave wrote a crash-recovery
 * snapshot and the next visit offered to "recover unsaved changes" for a deck
 * that had only been read, and every slide click pushed an undo entry, so
 * Ctrl+Z walked back through navigation instead of edits. Angular and Vanilla
 * raise dirty from explicit commit choke points and never had either symptom.
 *
 * Note this is only the CHANGE GATE: an edit still announces itself through
 * `markDirty()` the moment it commits, so narrowing the comparison cannot
 * swallow an edit made immediately after a navigation.
 */
export function serializeHistoryDocument(snapshot: EditorHistorySnapshot): string {
	return JSON.stringify({
		width: snapshot.width,
		height: snapshot.height,
		slides: snapshot.slides.map((slide) => ({
			...slide,
			rawXml: undefined,
			elements: slide.elements.map(historyElementContent),
		})),
		templateElementsBySlideId: Object.fromEntries(
			Object.entries(snapshot.templateElementsBySlideId).map(([slideId, elements]) => [
				slideId,
				elements.map(historyElementContent),
			]),
		),
		// A theme pick that changes nothing on the slides is still an edit, and an undo step.
		theme: snapshot.theme,
	});
}

/** `snapshot` recording the theme parts of the masters in `paths` only. */
function withMasterPathsIn(
	snapshot: EditorHistorySnapshot,
	paths: ReadonlySet<string>,
): EditorHistorySnapshot {
	const { theme, ...rest } = snapshot;
	if (!theme?.masterThemePaths) {
		return snapshot;
	}
	const { masterThemePaths, ...parts } = theme;
	const kept = Object.entries(masterThemePaths).filter(([path]) => paths.has(path));
	const next = kept.length > 0 ? { ...parts, masterThemePaths: Object.fromEntries(kept) } : parts;
	return Object.keys(next).length > 0 ? { ...rest, theme: next } : rest;
}

/**
 * The two snapshots differ only in which slide masters exist (Master view adds, duplicates or
 * deletes one, or a paste brings its own): every master both record keeps its theme part, and the
 * rest of the document is the same. That is no undo step; the tracking effect takes `next` as its
 * baseline instead.
 */
export function differsOnlyInMasterSet(
	previous: EditorHistorySnapshot,
	next: EditorHistorySnapshot,
): boolean {
	const before = Object.keys(previous.theme?.masterThemePaths ?? {});
	const after = new Set(Object.keys(next.theme?.masterThemePaths ?? {}));
	if (before.length === after.size && before.every((path) => after.has(path))) {
		return false;
	}
	const shared = new Set(before.filter((path) => after.has(path)));
	return (
		serializeHistoryDocument(withMasterPathsIn(previous, shared)) ===
		serializeHistoryDocument(withMasterPathsIn(next, shared))
	);
}

/** Record pending ordinary edits, then the explicit batch, as separate entries. */
export function recordExplicitSlideUpdate(
	history: EditorHistoryController<EditorHistorySnapshot>,
	previous: EditorHistorySnapshot | null,
	before: EditorHistorySnapshot,
	label: string | undefined,
): void {
	if (previous && serializeHistoryDocument(previous) !== serializeHistoryDocument(before)) {
		history.record(cloneHistorySnapshot(previous), previous.actionLabel);
	}
	history.record({ ...cloneHistorySnapshot(before), actionLabel: label }, label);
}
