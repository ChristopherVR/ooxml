import { createFindReplaceController } from 'ooxml-ui/pptx';
import type { FindReplaceMode, FindReplaceOptions, FindResult } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import type { ComputedRef, Ref, WritableComputedRef } from 'vue';
import { computed, watch } from 'vue';

import { useController } from './useController';

// The match descriptor, the search / replace implementations, the session and the bar's open
// state are shared with the other bindings (`createFindReplaceController`); this composable
// exposes them as refs the bar can `v-model`.
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
	/**
	 * A host-owned open flag to follow (closing it drops the matches, opening it searches
	 * again). Without it the bar's open state is {@link UseFindReplaceResult.open} alone.
	 */
	open?: Ref<boolean>;
}

export interface UseFindReplaceResult {
	/** Whether the bar is open; writing it opens (searching again) or closes the bar. */
	open: WritableComputedRef<boolean>;
	/** Ctrl+F / Ctrl+H: close the bar open in `mode`, switch its mode, or open it. */
	toggle: (mode?: FindReplaceMode) => void;
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
 * Find-and-replace state for the Vue editor: the shared `createFindReplaceController` read
 * through `useController`. Setting the query or a scope searches at once; a replace snapshots
 * history first, reassigns `slides.value` to the core's immutable result and searches again,
 * and an ordinary edit leaves the matches alone.
 */
export function useFindReplace({
	slides,
	showMatch,
	pushHistory,
	open: hostOpen,
}: UseFindReplaceInput): UseFindReplaceResult {
	const controller = createFindReplaceController({
		slides: () => slides.value,
		commit: (next) => {
			pushHistory();
			slides.value = next as PptxSlide[];
		},
		navigate: ({ slideIndex, elementId }) => showMatch(slideIndex, elementId ? [elementId] : []),
	});
	const state = useController(controller);
	const session = computed(() => state.value.session);

	if (hostOpen) {
		// Two-way: the host's flag drives the bar, and a shortcut toggle writes back to it.
		watch(hostOpen, (isOpen) => controller.setOpen(isOpen), { immediate: true });
		watch(
			() => state.value.open,
			(isOpen) => {
				if (hostOpen.value !== isOpen) {
					hostOpen.value = isOpen;
				}
			},
		);
	}

	const option = (key: keyof FindReplaceOptions): WritableComputedRef<boolean> =>
		computed({
			get: () => session.value.options[key],
			set: (value) => controller.search({ options: { [key]: value } }),
		});

	return {
		open: computed({
			get: () => state.value.open,
			set: (value) => controller.setOpen(value),
		}),
		toggle: controller.toggle,
		query: computed({
			get: () => session.value.query,
			set: (query) => controller.search({ query }),
		}),
		replacement: computed({
			get: () => session.value.replacement,
			set: controller.setReplacement,
		}),
		matchCase: option('matchCase'),
		includeNotes: option('includeNotes'),
		includeChartTitles: option('includeChartTitles'),
		includeChartData: option('includeChartData'),
		matches: computed(() => session.value.matches),
		matchCount: computed(() => session.value.matches.length),
		currentMatch: computed(() => session.value.current),
		next: controller.next,
		prev: controller.previous,
		replaceCurrent: () => {
			controller.replaceCurrent();
		},
		replaceAll: () => {
			controller.replaceAll();
		},
	};
}
