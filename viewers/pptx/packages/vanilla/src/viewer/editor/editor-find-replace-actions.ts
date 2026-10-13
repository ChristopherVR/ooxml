import {
	findInSlides,
	findReplacementAvailability,
	replaceInSlides,
	replaceMatch,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import type { EditorOps } from './editor-operations';

/**
 * Find & Replace actions for the ribbon's Home > Editing group, backed by the
 * shared `find-replace.ts` helpers. A lean, docked-panel implementation:
 * `search` reports a match count for the query, `replaceCurrent` replaces the
 * first match, and `replaceAll` replaces every match; there is no in-canvas
 * match highlighting or "next/previous match" cursor (that needs per-match
 * selection/scroll wiring into the renderer, out of scope for this wave).
 * Every replace is history-integrated like every other editor action.
 */
export interface FindReplaceActions {
	replacementAvailability?(
		query: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): ReturnType<typeof findReplacementAvailability>;
	/** Count occurrences of `query` across all slides (does not mutate). */
	search(
		query: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
	/** Replace the first match of `query`; returns the number replaced (0 or 1). */
	replaceCurrent(
		query: string,
		replacement: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
	/** Replace every match of `query`; returns the number replaced. */
	replaceAll(
		query: string,
		replacement: string,
		matchCase: boolean,
		includeNotes?: boolean,
		includeChartTitles?: boolean,
		includeChartData?: boolean,
	): number;
}

export interface FindReplaceActionsDeps {
	store: Store<ViewerState>;
	ops: EditorOps;
}

export function createFindReplaceActions(deps: FindReplaceActionsDeps): FindReplaceActions {
	const { store, ops } = deps;

	const commitReplace = (slides: readonly PptxSlide[], replacements: number): void => {
		if (replacements === 0) {
			return;
		}
		ops.pushHistory();
		store.set({ slides: [...slides] });
		ops.commitChange();
	};

	return {
		replacementAvailability(query, matchCase, includeNotes, includeChartTitles, includeChartData) {
			return findReplacementAvailability(
				findInSlides(store.get().slides, query, {
					matchCase,
					includeNotes,
					includeChartTitles,
					includeChartData,
				}),
				0,
			);
		},
		search(query, matchCase, includeNotes, includeChartTitles, includeChartData) {
			if (!query) {
				return 0;
			}
			return findInSlides(store.get().slides, query, {
				matchCase,
				includeNotes,
				includeChartTitles,
				includeChartData,
			}).length;
		},

		replaceCurrent(
			query,
			replacement,
			matchCase,
			includeNotes,
			includeChartTitles,
			includeChartData,
		) {
			const state = store.get();
			if (!state.editable || !query) {
				return 0;
			}
			const results = findInSlides(state.slides, query, {
				matchCase,
				includeNotes,
				includeChartTitles,
				includeChartData,
			});
			if (results[0]?.readOnly) {
				return 0;
			}
			const { slides, replacements } = replaceMatch(state.slides, results, 0, replacement);
			commitReplace(slides, replacements);
			return replacements;
		},

		replaceAll(query, replacement, matchCase, includeNotes, includeChartTitles, includeChartData) {
			const state = store.get();
			if (!state.editable || !query) {
				return 0;
			}
			if (
				findInSlides(state.slides, query, {
					matchCase,
					includeNotes,
					includeChartTitles,
					includeChartData,
				}).some((match) => match.readOnly)
			) {
				return 0;
			}
			const { slides, replacements } = replaceInSlides(state.slides, query, replacement, {
				matchCase,
				includeNotes,
				includeChartTitles,
				includeChartData,
			});
			commitReplace(slides, replacements);
			return replacements;
		},
	};
}
