import { buildOutline, createOutlineViewSession } from 'ooxml-ui/pptx';
import type { OutlineEdit, OutlineRow } from 'ooxml-ui/pptx';
/**
 * Outline-view state and keyboard, kept out of the SFC.
 *
 * The outline's rules and its edit session (`createOutlineViewSession`: refuse on a read-only
 * deck, apply, commit, hand the caret on) live in `ooxml-ui/pptx`; this composable derives the
 * rows from the deck and restores the caret after Vue has re-rendered.
 *
 * The composable does NOT own the deck. It takes a getter and a `commit`
 * callback, so the host decides how a new deck reaches history: this binding's
 * contract is `pushHistory()` immediately BEFORE `slides.value` is replaced,
 * and burying that in here would hide it from the component that owns it.
 */
import type { PptxSlide } from 'pptx-viewer-core';
import { computed, nextTick, ref } from 'vue';
import type { ComputedRef, Ref } from 'vue';

import type { CanvasSize } from '../types';

export interface UseOutlineViewInput {
	slides: () => PptxSlide[];
	canvasSize: () => CanvasSize;
	canEdit: () => boolean;
	/** Replace the deck and make `activeSlideIndex` current, with history. */
	commit: (slides: PptxSlide[], activeSlideIndex: number) => void;
}

export interface UseOutlineViewResult {
	rows: ComputedRef<OutlineRow[]>;
	containerRef: Ref<HTMLElement | null>;
	run: (edit: OutlineEdit) => void;
	onRowKeyDown: (event: KeyboardEvent, rowKey: string) => void;
}

export function useOutlineView(input: UseOutlineViewInput): UseOutlineViewResult {
	const rows = computed(() => buildOutline(input.slides()));
	const containerRef = ref<HTMLElement | null>(null);
	const session = createOutlineViewSession({
		slides: input.slides,
		canvasSize: input.canvasSize,
		canEdit: input.canEdit,
		commit: (slides, activeSlideIndex) => {
			input.commit(slides, activeSlideIndex);
			// After nextTick, because a new slide's title row does not exist in the DOM until Vue
			// has re-rendered with the new deck.
			void nextTick(() => session.restoreFocus(containerRef.value));
		},
	});

	return {
		rows,
		containerRef,
		run: (edit) => {
			session.run(edit);
		},
		onRowKeyDown: (event, rowKey) => session.handleKey(event, rowKey),
	};
}
