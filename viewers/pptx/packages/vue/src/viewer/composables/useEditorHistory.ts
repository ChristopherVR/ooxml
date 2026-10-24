import { createEditorHistoryController, historyTheme, restoreCrossedTheme } from 'ooxml-ui/pptx';
import type { HistoryTheme } from 'ooxml-ui/pptx';
import { cloneElement, cloneSlide } from 'pptx-viewer-core';
import type { PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import { computed, nextTick } from 'vue';
import type { ComputedRef, Ref } from 'vue';

import type { TemplateElementMap } from './template-editing';
import { useController } from './useController';

/**
 * useEditorHistory: framework-idiomatic undo/redo stack over `PptxSlide[]`.
 *
 * This is the Vue port of the React `useEditorHistory` hook, reduced to its
 * load-bearing core: an undo / redo stack of deep-cloned slide snapshots.
 * Unlike the React version it does **not** track canvas size or
 * template-element layers; the Vue editor foundation operates purely on the
 * `PptxSlide[]` model, so a snapshot is simply a cloned slide array.
 *
 * The stack itself is the shared history controller
 * (`createEditorHistoryController` in `ooxml-ui/pptx`), read through the
 * generic `useController` adapter; this composable adds only the `cloneSlide`
 * deep-snapshotting the shared stack deliberately leaves to the caller.
 *
 * Usage pattern (push-before-mutate):
 *
 * ```ts
 * const slides = shallowRef<PptxSlide[]>([]);
 * const history = useEditorHistory(slides);
 *
 * // before any mutation that should be undoable:
 * history.pushHistory();
 * slides.value = nextSlides;
 * ```
 *
 * `pushHistory` snapshots the *current* slide state onto the past stack and
 * clears the redo stack; callers invoke it immediately before committing a new
 * `slides.value`. `undo`/`redo` swap snapshots in and out of the live ref.
 */

// ---------------------------------------------------------------------------
// Result interface
// ---------------------------------------------------------------------------

export interface EditorHistoryResult {
	/** True when there is at least one snapshot to undo to. */
	canUndo: ComputedRef<boolean>;
	/** True when there is at least one snapshot to redo to. */
	canRedo: ComputedRef<boolean>;
	/** Description of the action Undo would revert, when one was recorded. */
	undoLabel: ComputedRef<string | undefined>;
	/** Description of the action Redo would re-apply, when one was recorded. */
	redoLabel: ComputedRef<string | undefined>;
	/**
	 * Snapshot the current `slides.value` onto the undo stack and clear the redo
	 * stack. Call this immediately **before** committing a mutating change.
	 */
	pushHistory: (label?: string) => void;
	/** Revert to the previous snapshot, pushing the current state onto redo. */
	undo: () => void;
	/** Re-apply the next snapshot, pushing the current state back onto undo. */
	redo: () => void;
	/** Drop all undo/redo history (e.g. when new content is loaded). */
	clearHistory: () => void;
	/**
	 * Apply a new File > Options > Advanced > "Maximum number of undos" value
	 * at runtime (`resolveHistoryDepth`). Trims the past stack immediately if
	 * the new limit is smaller.
	 */
	setMaxDepth: (depth: number) => void;
}

// ---------------------------------------------------------------------------
// Composable
// ---------------------------------------------------------------------------

/**
 * One undoable snapshot: the slide array plus (optionally) the separate template
 * (master/layout) element store, so edits in `editTemplateMode` undo/redo too.
 */
interface HistorySnapshot {
	slides: PptxSlide[];
	template: TemplateElementMap;
	/** The theme's colour scheme, font scheme and name, and the colour map built from them. */
	theme?: HistoryTheme;
	themeColorMap?: Record<string, string>;
}

/**
 * The deck theme an undo step carries along with the slides (`render/theme-history`). An undo or
 * redo that crosses a theme change puts the recorded theme and colour map back, and writes the
 * parts that differ into the handler's theme parts, so a save after it writes the theme it shows.
 */
export interface EditorHistoryThemeSource {
	theme: Ref<PptxTheme | undefined>;
	themeColorMap: Ref<Record<string, string> | undefined>;
	handler: Ref<PptxHandler | null | undefined>;
	/** The primary theme part's loaded name; the parsed theme carries none. */
	loadedThemeName?: () => string | undefined;
}

/** Optional tuning knobs for the history stack. */
export interface UseEditorHistoryOptions {
	/**
	 * Maximum undo depth. Defaults to the shared `DEFAULT_EDITOR_HISTORY_DEPTH`;
	 * File > Options > Advanced > "Maximum number of undos" threads
	 * `resolveHistoryDepth` here.
	 */
	maxDepth?: number;
	/**
	 * The slide on screen. An undo or redo can shorten the deck under it (undoing
	 * New Slide on the last slide removes the very slide being shown), so the
	 * index is pulled back into the restored deck, as React's history does.
	 * Without it the stage pointed past the end of the deck and drew nothing.
	 */
	activeSlideIndex?: Ref<number>;
	/**
	 * Called once per tick after an edit is recorded, undone or redone: the host's `change`
	 * event. Deferred a tick so a listener reads the committed deck, not the one before it.
	 */
	onChange?: () => void;
	/** The deck theme, recorded with every snapshot so a theme pick undoes like any other edit. */
	theme?: EditorHistoryThemeSource;
}

/**
 * @param slides Reactive reference to the live slide array the editor mutates.
 *               A `shallowRef` is recommended for large decks.
 * @param templateElementsBySlideId Optional reactive store of the per-slide
 *               master/layout (template) elements, snapshotted alongside slides
 *               so edits in `editTemplateMode` are undoable.
 * @param options Optional tuning knobs (undo depth).
 */
export function useEditorHistory(
	slides: Ref<PptxSlide[]>,
	templateElementsBySlideId?: Ref<TemplateElementMap>,
	options?: UseEditorHistoryOptions,
): EditorHistoryResult {
	// The shared history controller (one stack and one default depth for every
	// binding). It performs no cloning, so we snapshot with `cloneSlide` /
	// `cloneElement` before every record/undo/redo round-trip.
	const stack = createEditorHistoryController<HistorySnapshot>(
		options?.maxDepth === undefined ? {} : { maxDepth: options.maxDepth },
	);
	const state = useController(stack);

	const cloneTemplateMap = (source: TemplateElementMap): TemplateElementMap => {
		const next: TemplateElementMap = {};
		for (const [slideId, elements] of Object.entries(source)) {
			next[slideId] = elements.map(cloneElement);
		}
		return next;
	};

	const themeSource = options?.theme;
	const snapshot = (): HistorySnapshot => {
		const theme = historyTheme(themeSource?.theme.value);
		const themeColorMap = themeSource?.themeColorMap.value;
		return {
			slides: slides.value.map(cloneSlide),
			template: cloneTemplateMap(templateElementsBySlideId?.value ?? {}),
			...(theme ? { theme } : {}),
			...(themeColorMap ? { themeColorMap: { ...themeColorMap } } : {}),
		};
	};

	/** Moving from `current` to `snap` crosses a theme change: put the recorded theme back. */
	const restoreTheme = (current: HistorySnapshot, snap: HistorySnapshot): void => {
		if (!themeSource) {
			return;
		}
		const theme = restoreCrossedTheme(
			themeSource.theme.value,
			current.theme,
			snap.theme,
			themeSource.handler.value,
			themeSource.loadedThemeName?.(),
		);
		if (theme) {
			themeSource.theme.value = theme;
			themeSource.themeColorMap.value = snap.themeColorMap ? { ...snap.themeColorMap } : undefined;
		}
	};

	const restore = (snap: HistorySnapshot): void => {
		slides.value = snap.slides.map(cloneSlide);
		if (templateElementsBySlideId) {
			templateElementsBySlideId.value = cloneTemplateMap(snap.template);
		}
		const active = options?.activeSlideIndex;
		if (active) {
			const last = Math.max(snap.slides.length - 1, 0);
			if (active.value > last) {
				active.value = last;
			}
		}
	};

	const canUndo = computed(() => state.value.canUndo);
	const canRedo = computed(() => state.value.canRedo);
	const undoLabel = computed(() => state.value.undoLabel);
	const redoLabel = computed(() => state.value.redoLabel);

	let changePending = false;
	const notifyChange = (): void => {
		if (!options?.onChange || changePending) {
			return;
		}
		changePending = true;
		void nextTick(() => {
			changePending = false;
			options.onChange?.();
		});
	};

	const pushHistory = (label = ''): void => {
		stack.record(snapshot(), label);
		notifyChange();
	};

	const undo = (): void => {
		const current = snapshot();
		const result = stack.undo(current);
		if (!result) {
			return;
		}
		restore(result.snapshot);
		restoreTheme(current, result.snapshot);
		notifyChange();
	};

	const redo = (): void => {
		const current = snapshot();
		const result = stack.redo(current);
		if (!result) {
			return;
		}
		restore(result.snapshot);
		restoreTheme(current, result.snapshot);
		notifyChange();
	};

	const clearHistory = (): void => {
		stack.clear();
	};

	const setMaxDepth = (depth: number): void => {
		stack.setMaxDepth(depth);
	};

	return {
		canUndo,
		canRedo,
		undoLabel,
		redoLabel,
		pushHistory,
		undo,
		redo,
		clearHistory,
		setMaxDepth,
	};
}
