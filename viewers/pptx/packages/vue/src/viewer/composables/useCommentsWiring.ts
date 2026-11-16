import { createCommentsPanelController } from 'ooxml-ui/pptx';
import type { ReviewCommentsHost } from 'ooxml-ui/pptx';
import type { PptxComment, PptxSlide } from 'pptx-viewer-core';
import { computed, provide, ref, watch } from 'vue';
import type { ComputedRef, Ref } from 'vue';

import { COMMENTS_PANEL_CONTEXT } from './comments-panel-context';
import type { CommentsPanelContext } from './comments-panel-context';

export interface UseCommentsWiringInput {
	activeSlide: ComputedRef<PptxSlide | undefined>;
	activeSlideIndex: Ref<number>;
	slides: Ref<PptxSlide[]>;
	authorName: ComputedRef<string>;
	pushHistory: () => void;
	/** Whether the host allows edits; the comment commands are refused otherwise. */
	canEdit?: () => boolean;
	/** The element a new comment is anchored to (the selection). */
	selectedElementId?: () => string | undefined;
}

export interface UseCommentsWiringResult {
	showComments: Ref<boolean>;
	activeComments: ComputedRef<PptxComment[]>;
	/** The shared panel state, also provided to every `CommentsPanel` below the viewer. */
	panel: CommentsPanelContext;
	onCommentMarkerClick: (id: string) => void;
	/** Review > Comments Previous / Next / Delete over the whole deck. */
	reviewHost: ReviewCommentsHost;
}

/**
 * useCommentsWiring: the comments panel, its numbered on-canvas markers, and
 * the history-aware commit path that writes a slide's comments back. The panel's
 * drafts, edits and reply composers are the shared `createCommentsPanelController`,
 * provided to every `CommentsPanel` (side rail, inspector tab, mobile sheet) so they
 * all show the same per-slide drafts.
 */
export function useCommentsWiring(input: UseCommentsWiringInput): UseCommentsWiringResult {
	const { activeSlide, activeSlideIndex, slides, authorName, pushHistory } = input;

	const showComments = ref(false);
	const activeComments = computed(() => activeSlide.value?.comments ?? []);
	const commit = (next: PptxSlide[]): void => {
		pushHistory();
		slides.value = next;
	};
	const controller = createCommentsPanelController({
		slides: () => slides.value,
		canEdit: () => input.canEdit?.() ?? true,
		authorName: () => authorName.value,
		selectedElementId: () => input.selectedElementId?.(),
		commit,
	});
	// Drafts for slides that no longer exist go with them.
	watch(slides, () => controller.pruneDrafts());
	const panel: CommentsPanelContext = {
		controller,
		slideIndex: () => activeSlideIndex.value,
		slideId: () => activeSlide.value?.id,
		canEdit: () => input.canEdit?.() ?? true,
	};
	provide(COMMENTS_PANEL_CONTEXT, panel);

	/** Open the comments panel and focus the deck on the marker's slide. */
	function onCommentMarkerClick(_id: string): void {
		showComments.value = true;
	}

	const reviewHost: ReviewCommentsHost = {
		slides: () => slides.value,
		currentSlide: () => activeSlideIndex.value,
		goToSlide: (index) => {
			activeSlideIndex.value = index;
		},
		ensureCommentsOpen: () => {
			showComments.value = true;
		},
		applyComments: (changes) => {
			const nextSlides = slides.value.slice();
			for (const { slideIndex, comments } of changes) {
				const slide = nextSlides[slideIndex];
				if (slide) {
					nextSlides[slideIndex] = { ...slide, comments };
				}
			}
			commit(nextSlides);
		},
	};

	return { showComments, activeComments, panel, onCommentMarkerClick, reviewHost };
}
