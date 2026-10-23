import {
	acceptsPresentationInput,
	beginZoomExcursion,
	buildZoomTransitionOverride,
	resolveShowBackward,
	resolveShowForward,
} from 'ooxml-ui/pptx';
import type { ZoomExcursion, ZoomNavigationTarget } from 'ooxml-ui/pptx';
/**
 * usePresentationNavigation: where a slideshow advance actually lands.
 *
 * One composable rather than several because "next" is a three-level decision
 * that only reads correctly in one place: step the current slide's remaining
 * animation builds; if none are left, move to the next slide the SHOW visits
 * (hidden slides skipped, custom show honoured); if there is none, show the
 * black end screen or end the show outright. The slide-transition overlay is
 * driven by the same index change, so it lives here too.
 */
import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import type { ComputedRef, Ref } from 'vue';
import { computed, ref, watch } from 'vue';

import type { UseAnimationPlaybackResult } from './useAnimationPlayback';
import type { PresentationShowOrder } from './usePresentationShowOrder';

/** The active slide's transition, plus the two slides it animates between. */
export interface PresentationTransitionState {
	outgoing: PptxSlide | undefined;
	incoming: PptxSlide | undefined;
	transition: NonNullable<PptxSlide['transition']>;
}

export interface UsePresentationNavigationOptions {
	slides: () => PptxSlide[];
	startIndex: () => number;
	/**
	 * A GETTER, not the value: the playback controller is built FROM
	 * `activeSlide`, which this composable owns, so the two would otherwise be a
	 * construction cycle. Every read here happens at user-input time, long after
	 * both exist.
	 */
	playback: () => UseAnimationPlaybackResult;
	showOrder: PresentationShowOrder;
	/** File > Options > Advanced: show a black slide instead of ending outright. */
	endWithBlackSlide: () => boolean;
	/**
	 * Slide Show Setup > "Loop continuously until 'Esc'" (or a kiosk-type show,
	 * which forces it): advancing past the show's last slide wraps to its first
	 * slide instead of raising the end screen or exiting. Callers should pass
	 * shared's `shouldLoopContinuously(presentationProperties)`.
	 */
	loopContinuously?: () => boolean;
	/** Leave the show (goes through the keep-annotations prompt when needed). */
	requestClose: () => void;
	onSlideChange: (index: number) => void;
	/**
	 * Checked before the loop / black-slide / close fallback, when the running
	 * show has no next slide. Returning `true` means the caller already
	 * navigated (e.g. a `ppaction://customshow?...&return=true` sub-show
	 * restoring the origin show and slide) and `next()` should do nothing
	 * further. Returning `false` (or omitting the option) falls through to the
	 * normal end-of-show handling.
	 */
	onShowEnd?: () => boolean;
}

export interface UsePresentationNavigationResult {
	currentIndex: Ref<number>;
	activeSlide: ComputedRef<PptxSlide | undefined>;
	/** Black "End of slide show" screen shown past the last slide (option-gated). */
	showEndScreen: Ref<boolean>;
	/**
	 * Jump straight to a deck index. Deliberately NOT filtered by the show order:
	 * PowerPoint's typed "slide number + Enter" reaches a hidden slide on purpose,
	 * and slide-zoom tiles link to whatever slide they name.
	 */
	goTo: (index: number) => void;
	next: () => void;
	prev: () => void;
	/**
	 * Click/tap/swipe advance. Like `next()` it first steps the current slide's
	 * remaining animation builds, but once they are exhausted it advances the
	 * slide only when the slide's transition allows click-advance
	 * (`advanceOnClick !== false`), matching PowerPoint's "on mouse click" gate.
	 * Keyboard, the toolbar next button and the end screen keep calling `next()`.
	 */
	advanceFromClick: () => void;
	/**
	 * Handle a Slide Zoom / Section Zoom / Summary Zoom tile click: navigates
	 * to the target slide, playing the zoom's own `zmPr/@transitionDur` when
	 * authored, and (when `returnToParent` is set) arms an excursion so the
	 * next forward `next()` past the target's range returns to the slide the
	 * zoom was clicked from instead of continuing linearly through the deck.
	 */
	navigateToZoomTarget: (target: ZoomNavigationTarget) => void;
	/** Non-null while a slide transition is playing over the frame. */
	transitionState: Ref<PresentationTransitionState | null>;
	onTransitionDone: () => void;
}

export function usePresentationNavigation(
	options: UsePresentationNavigationOptions,
): UsePresentationNavigationResult {
	const { showOrder } = options;
	const playback = (): UseAnimationPlaybackResult => options.playback();

	function clampIndex(index: number): number {
		const last = Math.max(0, options.slides().length - 1);
		if (index < 0) {
			return 0;
		}
		return index > last ? last : index;
	}

	const currentIndex = ref(clampIndex(options.startIndex()));
	const activeSlide = computed<PptxSlide | undefined>(() => options.slides()[currentIndex.value]);
	const showEndScreen = ref(false);
	const transitionState = ref<PresentationTransitionState | null>(null);

	// A pending "return to zoom" excursion armed by `navigateToZoomTarget`,
	// consumed by `next()` once the show reaches the end of the target's
	// range. A Slide Zoom / Section Zoom click's own `zmPr/@transitionDur`
	// (see `pendingTransitionOverride` below) is independent of this: a zoom
	// without `returnToParent` still plays its own transition length, it just
	// never arms an excursion.
	const zoomExcursion = ref<ZoomExcursion | undefined>(undefined);
	// A one-shot override for the transition the very next `currentIndex`
	// change should play, consumed by the `watch` below and cleared
	// immediately after. Not reactive state: nothing reads it directly.
	let pendingTransitionOverride: PptxSlideTransition | undefined;

	function goTo(index: number, transitionOverride?: PptxSlideTransition): void {
		const target = clampIndex(index);
		if (target === currentIndex.value) {
			return;
		}
		pendingTransitionOverride = transitionOverride;
		currentIndex.value = target;
	}

	function navigateToZoomTarget(target: ZoomNavigationTarget): void {
		zoomExcursion.value = beginZoomExcursion(target, currentIndex.value, options.slides());
		goTo(target.targetSlideIndex, buildZoomTransitionOverride(target.transitionDurationMs));
	}

	/**
	 * A forward input, decided by the shared show step: the end screen ends the
	 * show, a remaining build is revealed, a click past the last build is gated
	 * by the slide's transition, a Zoom excursion returns to its zoom slide, and
	 * past the show's end a returning sub-show resumes, a looping show wraps,
	 * the black end screen rises or the show ends.
	 */
	function forward(fromClick: boolean): void {
		const step = resolveShowForward({
			currentIndex: currentIndex.value,
			currentSlide: activeSlide.value,
			showOrder: showOrder.indexes.value,
			endOfShow: showEndScreen.value,
			fromClick,
			loop: options.loopContinuously?.() === true,
			endWithBlackSlide: options.endWithBlackSlide(),
			zoomExcursion: zoomExcursion.value,
			returnFromSubShow: () => options.onShowEnd?.() === true,
			advanceBuild: () => playback().advance(),
		});
		switch (step.kind) {
			case 'exit':
				// A second advance on the end screen exits the show, like PowerPoint.
				options.requestClose();
				return;
			case 'end-screen':
				showEndScreen.value = true;
				return;
			case 'go':
				if (step.zoomReturn) {
					zoomExcursion.value = step.zoomReturn.excursion;
					goTo(step.index, step.zoomReturn.transitionOverride);
					return;
				}
				goTo(step.index);
				break;
			default:
				break;
		}
	}

	function next(): void {
		forward(false);
	}

	function prev(): void {
		const step = resolveShowBackward({
			currentIndex: currentIndex.value,
			showOrder: showOrder.indexes.value,
			endOfShow: showEndScreen.value,
			seededCompleted: playback().seededCompleted.value,
		});
		if (step.kind === 'dismiss-end-screen') {
			showEndScreen.value = false;
		} else if (step.kind === 'replay') {
			// A slide entered backward shows its builds already complete; the next
			// back press replays them rather than leaving the slide (PowerPoint).
			playback().reset();
		} else if (step.kind === 'go') {
			// PowerPoint shows a slide you step BACK onto with its builds played.
			playback().markNextEntryCompleted();
			goTo(step.index);
		}
	}

	function advanceFromClick(): void {
		// An audience display never drives itself: a tap or swipe of its own would
		// move it off the presenter's slide, and the next snapshot would drag it back.
		if (!acceptsPresentationInput()) {
			return;
		}
		forward(true);
	}

	watch(currentIndex, (index, previousIndex) => {
		options.onSlideChange(index);
		// The playback controller rebuilds itself on the active-slide change (it
		// watches `activeSlide`), so no explicit reset is needed here.
		const incoming = options.slides()[index];
		// A Slide Zoom / Section Zoom / Summary Zoom navigation's own
		// `zmPr/@transitionDur` takes priority over either slide's own authored
		// transition; it applies to this ONE jump only.
		const override = pendingTransitionOverride;
		pendingTransitionOverride = undefined;
		// Forward steps play the ENTERING slide's transition; a backward step
		// replays the LEAVING slide's transition in reverse (a morph glides its
		// shapes back to where they came from).
		const transition =
			override ?? (index < previousIndex ? options.slides()[previousIndex] : incoming)?.transition;
		transitionState.value =
			transition && transition.type && transition.type !== 'none'
				? { outgoing: options.slides()[previousIndex], incoming, transition }
				: null;
	});

	return {
		currentIndex,
		activeSlide,
		showEndScreen,
		goTo,
		next,
		prev,
		advanceFromClick,
		navigateToZoomTarget,
		transitionState,
		onTransitionDone: () => {
			transitionState.value = null;
		},
	};
}
