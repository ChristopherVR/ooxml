import { createEditorHistoryController, crossedHistoryTheme, historyTheme } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHeaderFooter, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { useRef, useState, useCallback, useEffect } from 'react';

import type { CanvasSize, EditorHistorySnapshot, EditorHistoryTheme } from '../types';
import { cloneHistorySnapshot, cloneSlide, cloneTemplateElementsBySlideId } from '../utils/clone';
import {
	recordExplicitSlideUpdate,
	serializeHistoryDocument as serializeDocument,
} from './explicit-slide-history';
import { useController } from './useController';
import { usePointerReleaseWakeup } from './usePointerReleaseWakeup';

// ---------------------------------------------------------------------------
// Input / output interfaces
// ---------------------------------------------------------------------------

export interface EditorHistoryInput {
	slides: PptxSlide[];
	canvasSize: CanvasSize;
	activeSlideIndex: number;
	templateElementsBySlideId: Record<string, PptxElement[]>;
	selectedElementId: string | null;
	selectedElementIds: string[];
	editTemplateMode: boolean;
	headerFooter: PptxHeaderFooter;
	loading: boolean;
	error: string | null;
	/**
	 * Undo-stack depth (File > Options > Advanced > "Maximum number of undos",
	 * resolved via the shared `resolveHistoryDepth`). Defaults to the shared
	 * `DEFAULT_EDITOR_HISTORY_DEPTH`.
	 */
	maxHistoryEntries?: number;
	hasActivePointerInteraction: () => boolean;
	pointerCommitNonce: number;
	/**
	 * Raised whenever a local edit is committed, i.e. exactly when this hook
	 * learns the deck has diverged from what was last loaded or saved.
	 *
	 * This is what drives `state.isDirty`, and through it the status bar, the
	 * host's `onDirtyChange` (the demos hang their "* filename" title marker off
	 * it) and - critically - `useAutosave`, which short-circuits on a clean
	 * document. Before this existed, `isDirty` was raised only by a few
	 * master-view and document-property paths, so an element nudge, Home > New
	 * Slide or a notes edit left the flag false and crash recovery never ran.
	 */
	onDirty?: () => void;
	/**
	 * The deck's theme. Its colour scheme, font scheme and name are recorded with every snapshot,
	 * so a theme pick is an undo step like any other edit.
	 */
	theme?: PptxTheme | undefined;
	/**
	 * Put a recorded theme back (an undo or redo crossing a theme change). The slides come back
	 * with the snapshot; the theme has to be written to the handler's theme part too, which this
	 * hook cannot reach, so the theme handlers do it.
	 */
	onRestoreTheme?: (theme: EditorHistoryTheme) => void;
	// Setters for applying snapshots
	setSlides: (slides: PptxSlide[]) => void;
	setCanvasSize: (size: CanvasSize) => void;
	setActiveSlideIndex: (index: number) => void;
	setTemplateElementsBySlideId: (map: Record<string, PptxElement[]>) => void;
	setSelectedElementId: (id: string | null) => void;
	setSelectedElementIds: (ids: string[]) => void;
	setEditTemplateMode: (mode: boolean) => void;
	setHeaderFooter: (hf: PptxHeaderFooter) => void;
}

export interface EditorHistoryResult {
	commitSlides: (slides: PptxSlide[], label?: string, beforeSlides?: PptxSlide[]) => void;
	canUndo: boolean;
	canRedo: boolean;
	undoLabel: string | undefined;
	redoLabel: string | undefined;
	handleUndo: () => void;
	handleRedo: () => void;
	resetHistory: () => void;
	/**
	 * Pause (`true`) and resume (`false`) change tracking. A gallery hover preview swaps the
	 * slides in for a moment; while tracking is held that is neither an undo step nor an edit, and
	 * when the original is back the tracker sees no change.
	 */
	holdTracking: (held: boolean) => void;
	markDirty: () => void;
	buildHistorySnapshot: (actionLabel?: string) => EditorHistorySnapshot;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useEditorHistory(input: EditorHistoryInput): EditorHistoryResult {
	const {
		slides,
		canvasSize,
		activeSlideIndex,
		templateElementsBySlideId,
		loading,
		error,
		maxHistoryEntries,
		hasActivePointerInteraction,
		pointerCommitNonce,
		onDirty,
		theme,
		onRestoreTheme,
		setSlides,
		setCanvasSize,
		setActiveSlideIndex,
		setTemplateElementsBySlideId,
		setSelectedElementId,
		setSelectedElementIds,
	} = input;

	// -- Refs ---------------------------------------------------------------

	/**
	 * The shared undo/redo controller: one stack and one default depth for every
	 * binding. This hook keeps only React's change detection (below), which
	 * decides WHEN a snapshot is recorded; the stack itself is shared.
	 */
	const [history] = useState(() => createEditorHistoryController<EditorHistorySnapshot>());
	const historyState = useController(history);
	const lastHistorySnapshotRef = useRef<EditorHistorySnapshot | null>(null);
	const lastHistorySerializedRef = useRef<string>('');
	/**
	 * Cheap structural hash so we can short-circuit the expensive
	 * JSON.stringify when no slide-shape change has occurred. Only when the
	 * cheap hash differs do we fall back to a full deep stringify comparison.
	 */
	const lastCheapHashRef = useRef<string>('');
	const isApplyingHistoryRef = useRef(false);
	const trackingHeldRef = useRef(false);
	const unlockHistoryTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

	// -- State --------------------------------------------------------------

	/**
	 * Monotonic counter bumped by every `markDirty()` call, i.e. by every local
	 * edit-commit choke point in the editor.
	 *
	 * It exists because the cheap hash below is deliberately blind to an
	 * element's CONTENT: it only sees slide / element counts. An edit that
	 * rewrites a property in place (any inspector field, a ribbon format
	 * command, an inline-text commit) changes no count, so without a nonce the
	 * hash is byte-identical before and after and the effect returns before it
	 * ever reaches the snapshot push. The edit lands on screen and is silently
	 * absent from the undo stack.
	 *
	 * `markDirty` previously flipped a boolean, which changes state exactly
	 * once in a session and so could not re-open the gate. Counting instead
	 * makes every commit distinguishable.
	 */
	const [editCommitNonce, setEditCommitNonce] = useState(0);
	/**
	 * Re-runs the tracking effect once a pointer interaction it deferred behind
	 * has ended. `pointerCommitNonce` only covers a pointer-up that MOVED
	 * something; see `usePointerReleaseWakeup` for the click-away case.
	 */
	const { pointerReleaseNonce, deferUntilPointerReleased } = usePointerReleaseWakeup(
		hasActivePointerInteraction,
	);

	// -- Helpers ------------------------------------------------------------

	useEffect(() => {
		if (maxHistoryEntries !== undefined) {
			history.setMaxDepth(maxHistoryEntries);
		}
	}, [history, maxHistoryEntries]);

	const buildHistorySnapshot = useCallback(
		(actionLabel?: string): EditorHistorySnapshot => {
			const recordedTheme = historyTheme(theme);
			return {
				width: canvasSize.width,
				height: canvasSize.height,
				activeSlideIndex,
				slides: slides.map(cloneSlide),
				templateElementsBySlideId: cloneTemplateElementsBySlideId(templateElementsBySlideId),
				...(recordedTheme ? { theme: recordedTheme } : {}),
				...(actionLabel ? { actionLabel } : {}),
			};
		},
		[activeSlideIndex, canvasSize, slides, templateElementsBySlideId, theme],
	);

	const applyHistorySnapshot = useCallback(
		(snapshot: EditorHistorySnapshot) => {
			const maxSlideIndex = Math.max(snapshot.slides.length - 1, 0);
			setCanvasSize({
				width: snapshot.width,
				height: snapshot.height,
			});
			setSlides(snapshot.slides.map(cloneSlide));
			setTemplateElementsBySlideId(
				cloneTemplateElementsBySlideId(snapshot.templateElementsBySlideId),
			);
			setActiveSlideIndex(Math.min(snapshot.activeSlideIndex, maxSlideIndex));
			setSelectedElementIds([]);
			setSelectedElementId(null);
		},
		[
			setActiveSlideIndex,
			setCanvasSize,
			setSelectedElementId,
			setSelectedElementIds,
			setSlides,
			setTemplateElementsBySlideId,
		],
	);

	const unlockHistoryTracking = useCallback(() => {
		if (unlockHistoryTimerRef.current) {
			clearTimeout(unlockHistoryTimerRef.current);
		}
		unlockHistoryTimerRef.current = setTimeout(() => {
			isApplyingHistoryRef.current = false;
		}, 0);
	}, []);

	/**
	 * Held in a ref so `markDirty` can stay identity-stable: it is passed down
	 * into dozens of memoised handlers, and a changing identity would invalidate
	 * all of them on every render.
	 */
	const onDirtyRef = useRef(onDirty);
	useEffect(() => {
		onDirtyRef.current = onDirty;
	}, [onDirty]);

	const markDirty = useCallback(() => {
		setEditCommitNonce((previous) => previous + 1);
		onDirtyRef.current?.();
	}, []);

	const onRestoreThemeRef = useRef(onRestoreTheme);
	useEffect(() => {
		onRestoreThemeRef.current = onRestoreTheme;
	}, [onRestoreTheme]);
	/** Moving from `current` to `target` crosses a theme change: put the target's theme back. */
	const restoreTheme = useCallback(
		(current: EditorHistorySnapshot, target: EditorHistorySnapshot) => {
			const crossed = crossedHistoryTheme(current.theme, target.theme);
			if (crossed) {
				onRestoreThemeRef.current?.(crossed);
			}
		},
		[],
	);

	// -- Stack navigation ---------------------------------------------------

	const handleUndo = useCallback(() => {
		if (!history.getSnapshot().canUndo) {
			return;
		}
		const currentSnapshot = buildHistorySnapshot();
		const result = history.undo(currentSnapshot);
		if (!result) {
			return;
		}

		isApplyingHistoryRef.current = true;
		const nextSnapshot = cloneHistorySnapshot(result.snapshot);
		lastHistorySnapshotRef.current = cloneHistorySnapshot(nextSnapshot);
		lastHistorySerializedRef.current = serializeDocument(nextSnapshot);
		applyHistorySnapshot(nextSnapshot);
		restoreTheme(currentSnapshot, nextSnapshot);
		unlockHistoryTracking();
		markDirty();
	}, [
		applyHistorySnapshot,
		buildHistorySnapshot,
		history,
		markDirty,
		restoreTheme,
		unlockHistoryTracking,
	]);

	const handleRedo = useCallback(() => {
		if (!history.getSnapshot().canRedo) {
			return;
		}
		const currentSnapshot = buildHistorySnapshot();
		const result = history.redo(currentSnapshot);
		if (!result) {
			return;
		}

		isApplyingHistoryRef.current = true;
		const targetSnapshot = cloneHistorySnapshot(result.snapshot);
		lastHistorySnapshotRef.current = cloneHistorySnapshot(targetSnapshot);
		lastHistorySerializedRef.current = serializeDocument(targetSnapshot);
		applyHistorySnapshot(targetSnapshot);
		restoreTheme(currentSnapshot, targetSnapshot);
		unlockHistoryTracking();
		markDirty();
	}, [
		applyHistorySnapshot,
		buildHistorySnapshot,
		history,
		markDirty,
		restoreTheme,
		unlockHistoryTracking,
	]);

	// -- Reset --------------------------------------------------------------

	const resetHistory = useCallback(
		(initialSnapshot?: EditorHistorySnapshot | null) => {
			history.clear();
			if (initialSnapshot) {
				const clonedInitial = cloneHistorySnapshot(initialSnapshot);
				lastHistorySnapshotRef.current = clonedInitial;
				lastHistorySerializedRef.current = serializeDocument(clonedInitial);
			} else {
				lastHistorySnapshotRef.current = null;
				lastHistorySerializedRef.current = '';
			}
		},
		[history],
	);

	// -- Cleanup timer on unmount -------------------------------------------

	useEffect(() => {
		return () => {
			if (unlockHistoryTimerRef.current) {
				clearTimeout(unlockHistoryTimerRef.current);
				unlockHistoryTimerRef.current = null;
			}
		};
	}, []);

	// -- History tracking effect --------------------------------------------

	useEffect(() => {
		if (loading || error) {
			return;
		}
		if (isApplyingHistoryRef.current || trackingHeldRef.current) {
			return;
		}
		if (hasActivePointerInteraction()) {
			// Defer, but make sure this run happens once the gesture ends: a
			// plain click never bumps `pointerCommitNonce`, so without this the
			// pending change waited for the NEXT edit and undid with it.
			deferUntilPointerReleased();
			return;
		}

		// ── Cheap hash gate ────────────────────────────────────────────
		// Skip the deep stringify when slide / element counts and ids
		// match the previous snapshot. This catches the very common case
		// where the effect re-runs but no real state shape changed.
		//
		// Both nonces are part of the hash because the counts alone cannot see
		// a content-only edit, and a content-only edit is still an undo step:
		//   - `pointerCommitNonce` covers a committed pointer interaction
		//     (move / resize / adjust / on-canvas chart edit).
		//   - `editCommitNonce` covers every other local commit, via the
		//     `markDirty()` that each edit choke point already calls: inspector
		//     fields, ribbon formatting, inline-text commits, table and theme
		//     edits. Without it none of those armed Undo.
		//   - `pointerReleaseNonce` covers a run this effect deferred behind an
		//     active pointer interaction (a click-away commit; see
		//     `usePointerReleaseWakeup`), which no other nonce re-opens.
		// Opening the gate only costs a stringify; the serialized comparison
		// below still rejects a commit that changed nothing, so a handler that
		// calls `markDirty()` without touching the deck pushes no entry.
		const cheapHash = `${pointerCommitNonce}|${editCommitNonce}|${pointerReleaseNonce}|${slides.length}|${activeSlideIndex}|${canvasSize.width}x${canvasSize.height}|${slides
			.map((s) => `${s.id}:${s.elements.length}`)
			.join('/')}`;
		if (cheapHash === lastCheapHashRef.current) {
			return;
		}

		const snapshot = buildHistorySnapshot();
		// Document only: see `serializeDocument`. A slide selection reaches here
		// (the cheap hash above deliberately still lets it through, so the gate
		// cannot swallow a same-tick edit) and stops on this comparison instead of
		// being reported as a mutation.
		const serialized = serializeDocument(snapshot);
		if (serialized === lastHistorySerializedRef.current) {
			// Nothing about the deck moved, only the view. Keep the stored
			// snapshot's index current so a later undo returns to the slide the
			// user was actually editing, then stop: no dirty flag, no undo entry.
			const previous = lastHistorySnapshotRef.current;
			if (previous && previous.activeSlideIndex !== snapshot.activeSlideIndex) {
				lastHistorySnapshotRef.current = {
					...previous,
					activeSlideIndex: snapshot.activeSlideIndex,
				};
			}
			lastCheapHashRef.current = cheapHash;
			return;
		}

		const previousSnapshot = lastHistorySnapshotRef.current;
		// Both snapshots are already isolated from live state; transfer ownership.
		if (!previousSnapshot) {
			lastHistorySnapshotRef.current = snapshot;
			lastHistorySerializedRef.current = serialized;
			lastCheapHashRef.current = cheapHash;
			return;
		}

		// Reaching here means the deck genuinely differs from the last snapshot,
		// which is the strongest "this document is dirty" signal there is: it is
		// true even for an edit path that forgot to call `markDirty()`. The very
		// first snapshot after a load takes the `!previousSnapshot` branch above
		// and so never reports dirty.
		onDirtyRef.current?.();
		history.record(previousSnapshot, previousSnapshot.actionLabel);
		lastHistorySnapshotRef.current = snapshot;
		lastHistorySerializedRef.current = serialized;
		lastCheapHashRef.current = cheapHash;
	}, [
		activeSlideIndex,
		buildHistorySnapshot,
		canvasSize.height,
		canvasSize.width,
		deferUntilPointerReleased,
		editCommitNonce,
		error,
		hasActivePointerInteraction,
		history,
		loading,
		pointerCommitNonce,
		pointerReleaseNonce,
		slides,
	]);

	// -- Public API ---------------------------------------------------------

	const commitSlides = (nextSlides: PptxSlide[], label?: string, beforeSlides = slides): void => {
		const before = { ...buildHistorySnapshot(), slides: beforeSlides.map(cloneSlide) };
		const after = { ...before, slides: nextSlides.map(cloneSlide) };
		recordExplicitSlideUpdate(history, lastHistorySnapshotRef.current, before, label);
		lastHistorySnapshotRef.current = cloneHistorySnapshot(after);
		lastHistorySerializedRef.current = serializeDocument(after);
		isApplyingHistoryRef.current = false;
		setSlides(nextSlides);
		onDirtyRef.current?.();
	};

	const holdTracking = useCallback((held: boolean) => {
		trackingHeldRef.current = held;
	}, []);

	return {
		commitSlides,
		holdTracking,
		canUndo: historyState.canUndo,
		canRedo: historyState.canRedo,
		undoLabel: historyState.undoLabel,
		redoLabel: historyState.redoLabel,
		handleUndo,
		handleRedo,
		resetHistory,
		markDirty,
		buildHistorySnapshot,
	};
}
