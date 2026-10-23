import { findReplaceSessionActions, INITIAL_FIND_REPLACE_SESSION } from 'ooxml-ui/pptx';
import type { FindReplaceOptions, FindReplaceSession, FindResult } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import type { ComputedRef, Ref, WritableComputedRef } from 'vue';
import { computed, shallowRef, watch } from 'vue';

// The match descriptor, the search / replace implementations and the session around them
// (`find-replace-session`) are shared with the other bindings; this composable keeps the
// session in a ref.
export type { FindResult } from 'ooxml-ui/pptx';

export interface UseFindReplaceInput {
	/** Reactive list of slides being edited. */
	slides: Ref<PptxSlide[]>;
	/**
	 * Show a match: go to its slide and select its element there (an empty list for a
	 * speaker-notes match) in one step, the viewer's `showSlideSelection`, which keeps that
	 * selection across the slide change.
	 */
	showMatch: (slideIndex: number, ids: string[]) => void;
	/** Snapshot the current document onto the undo/redo history stack. */
	pushHistory: () => void;
	/** Whether the bar is open: closing it drops the matches, opening it searches again. */
	open?: Ref<boolean>;
}

export interface UseFindReplaceResult {
	query: WritableComputedRef<string>;
	replacement: WritableComputedRef<string>;
	matchCase: WritableComputedRef<boolean>;
	includeNotes: WritableComputedRef<boolean>;
	includeChartTitles: WritableComputedRef<boolean>;
	includeChartData: WritableComputedRef<boolean>;
	matches: ComputedRef<readonly FindResult[]>;
	matchCount: ComputedRef<number>;
	/** Zero-based current match; -1 when there is none. */
	currentMatch: ComputedRef<number>;
	next: () => void;
	prev: () => void;
	replaceCurrent: () => void;
	replaceAll: () => void;
}

/**
 * Find-and-replace state for the Vue editor: the shared `find-replace-session` in a ref, exposed
 * as refs the bar can `v-model`. Setting the query or a scope searches at once; a replace
 * snapshots history first, reassigns `slides.value` to the core's immutable result and searches
 * again, and an ordinary edit leaves the matches alone.
 */
export function useFindReplace({
	slides,
	showMatch,
	pushHistory,
	open,
}: UseFindReplaceInput): UseFindReplaceResult {
	const session = shallowRef<FindReplaceSession>(INITIAL_FIND_REPLACE_SESSION);
	const actions = findReplaceSessionActions(
		{ get: () => session.value, set: (next) => (session.value = next) },
		{
			slides: () => slides.value,
			commit: (next) => {
				pushHistory();
				slides.value = next as PptxSlide[];
			},
			navigate: ({ slideIndex, elementId }) => showMatch(slideIndex, elementId ? [elementId] : []),
		},
	);

	const option = (key: keyof FindReplaceOptions): WritableComputedRef<boolean> =>
		computed({
			get: () => session.value.options[key],
			set: (value) => actions.search({ options: { [key]: value } }),
		});

	if (open) {
		watch(open, (isOpen) => (isOpen ? actions.search() : actions.close()));
	}

	return {
		query: computed({
			get: () => session.value.query,
			set: (query) => actions.search({ query }),
		}),
		replacement: computed({
			get: () => session.value.replacement,
			set: actions.setReplacement,
		}),
		matchCase: option('matchCase'),
		includeNotes: option('includeNotes'),
		includeChartTitles: option('includeChartTitles'),
		includeChartData: option('includeChartData'),
		matches: computed(() => session.value.matches),
		matchCount: computed(() => session.value.matches.length),
		currentMatch: computed(() => session.value.current),
		next: actions.next,
		prev: actions.previous,
		replaceCurrent: () => {
			actions.replaceCurrent();
		},
		replaceAll: () => {
			actions.replaceAll();
		},
	};
}
