import { resolveAutoAdvanceDelayMs, resolveShowBackward, resolveShowForward } from 'ooxml-ui/pptx';
import type { ZoomExcursion } from 'ooxml-ui/pptx';
import type { PptxAction, PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { useRef, useCallback } from 'react';

import type { ViewerMode } from '../../types';
import { handlePresentationActionImpl } from './presentation-actions';
import { executeSlideTransition } from './slide-transition';
import type { SlideTransitionDeps } from './slide-transition';
import type { PresentationTransitionOverlayState, SlideAdvanceTrigger } from './types';

// ---------------------------------------------------------------------------
// Sub-hook interface
// ---------------------------------------------------------------------------

export interface UseSlideNavigationInput {
	slides: PptxSlide[];
	visibleSlideIndexes: number[];
	presentationSlideIndex: number;
	setPresentationSlideIndex: (index: number) => void;
	setPresentationSlideVisible: (visible: boolean) => void;
	setTransitionOverlay: (state: PresentationTransitionOverlayState | null) => void;
	onSetMode: (mode: ViewerMode) => void;
	onSetActiveSlideIndex: (index: number) => void;
	onPlayActionSound?: (soundPath: string) => void;
	/** Stop the current transition sound (`p:sndAc/p:endSnd`). See `slide-transition.ts`. */
	onStopActionSound?: () => void;
	loopContinuously?: boolean;
	/** Whether to use rehearsed auto-advance timings. When false, slides advance only on click. */
	useTimings?: boolean;
	/**
	 * Fired when the user advances past the last slide (no loop, no rehearsal).
	 * The presentation-mode hook uses it for PowerPoint's "End with black
	 * slide" behavior. When absent, advancing past the end clamps as before.
	 */
	onAdvancePastLastSlide?: () => void;
	playNextAnimationGroup: () => boolean;
	clearPresentationTimers: () => void;
	/** Seed the incoming slide's initial animation states (no playback). */
	seedSlideAnimations: (slideIndex: number, options?: { completed?: boolean }) => void;
	/** Start the seeded slide's playback (auto-play group + entrance timers). */
	startSlideAnimations: (slideIndex: number) => void;
	presentationTimersRef: { current: number[] };
	rehearsing: boolean;
	recordCurrentSlideTime: (slideIndex: number) => void;
	setShowRehearsalSummary: (value: boolean) => void;
	/** `ppaction://hlinkshowjump?jump=lastslideviewed`. */
	onLastViewed?: () => void;
	/** `ppaction://customshow?id=<id>[&return=true]`. */
	onCustomShow?: (customShowId: string, returnAfter: boolean) => void;
	/** `ppaction://hlinkfile`. */
	onOpenFile?: (target: string) => void;
	/** `ppaction://hlinkpres`. */
	onOpenPresentation?: (target: string) => void;
	/** `ppaction://media`. */
	onPlayMedia?: (elementId: string | undefined) => void;
	/** `ppaction://ole?verb=<n>`. */
	onOleVerb?: (verb: number, elementId: string | undefined) => void;
	/** `ppaction://program`. */
	onRunProgram?: (target: string) => void;
	/**
	 * A pending "return to zoom" excursion armed by `useZoomNavigation`'s
	 * `handleZoomClick`. Consulted on every forward advance: once the show
	 * reaches `excursion.endSlideIndex`, the step jumps back to
	 * `excursion.returnSlideIndex` (playing the zoom's own `transitionDur`)
	 * instead of continuing linearly through the deck.
	 */
	zoomExcursionRef: React.RefObject<ZoomExcursion | undefined>;
}

export interface UseSlideNavigationResult {
	movePresentationSlide: (direction: 1 | -1, trigger?: SlideAdvanceTrigger) => void;
	navigateToSlide: (slideIndex: number, transitionOverride?: PptxSlideTransition) => void;
	/** `elementId` is the clicked element, for the verbs that act on it (`playMedia`, `oleVerb`). */
	handlePresentationAction: (action: PptxAction, elementId?: string) => void;
	scheduleAutoAdvanceForSlide: (slideIndex: number) => void;
}

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export function useSlideNavigation(input: UseSlideNavigationInput): UseSlideNavigationResult {
	const {
		slides,
		visibleSlideIndexes,
		presentationSlideIndex,
		setPresentationSlideIndex,
		setPresentationSlideVisible,
		setTransitionOverlay,
		onSetMode,
		onSetActiveSlideIndex,
		onPlayActionSound,
		onStopActionSound,
		loopContinuously,
		useTimings,
		onAdvancePastLastSlide,
		playNextAnimationGroup,
		clearPresentationTimers,
		seedSlideAnimations,
		startSlideAnimations,
		presentationTimersRef,
		rehearsing,
		recordCurrentSlideTime,
		setShowRehearsalSummary,
		onLastViewed,
		onCustomShow,
		onOpenFile,
		onOpenPresentation,
		onPlayMedia,
		onOleVerb,
		onRunProgram,
		zoomExcursionRef,
	} = input;

	const movePresentationSlideRef = useRef<(direction: 1 | -1) => void>(() => {});

	// -----------------------------------------------------------------------
	// Auto-advance scheduling (uses ref to break circular dependency)
	// -----------------------------------------------------------------------

	const scheduleAutoAdvanceForSlide = useCallback(
		(slideIndex: number) => {
			// PowerPoint's "Advance slide: After <n>" (`p:transition/@advTm`),
			// resolved by the shared helper every binding shares so a deck cannot
			// advance in one framework and stall in another. `useTimings === false`
			// is "Advance slides: Manually", which ignores authored timings.
			const delayMs = resolveAutoAdvanceDelayMs(slides[slideIndex], { useTimings });
			if (delayMs === undefined) {
				return;
			}

			const timer = window.setTimeout(() => {
				movePresentationSlideRef.current(1);
			}, delayMs);
			presentationTimersRef.current.push(timer);
		},
		[slides, presentationTimersRef, useTimings],
	);

	// -----------------------------------------------------------------------
	// Slide navigation
	// -----------------------------------------------------------------------

	const movePresentationSlide = useCallback(
		(direction: 1 | -1, trigger: SlideAdvanceTrigger = 'explicit') => {
			// The show order (custom-show membership minus hidden slides) is resolved
			// upstream. The end screen and the replay of a slide entered backward
			// are handled by `usePresentationMode` before this runs.
			const showOrder =
				visibleSlideIndexes.length > 0
					? visibleSlideIndexes
					: slides.map((_slide, slideIndex) => slideIndex);
			const current = presentationSlideIndex;
			const transitionTo = (index: number, options: Partial<SlideTransitionDeps> = {}): void =>
				executeSlideTransition(index, {
					slides,
					currentSlideIndex: current,
					onPlayActionSound,
					onStopActionSound,
					setPresentationSlideVisible,
					clearPresentationTimers,
					setPresentationSlideIndex,
					onSetActiveSlideIndex,
					seedSlideAnimations,
					startSlideAnimations,
					scheduleAutoAdvanceForSlide: rehearsing ? undefined : scheduleAutoAdvanceForSlide,
					presentationTimersRef,
					setTransitionOverlay,
					// PowerPoint plays a slide's transition only when advancing into it,
					// and replays the LEAVING slide's transition in reverse when stepping
					// back (a morph glides its shapes back to where they came from).
					playTransition: true,
					...options,
				});

			if (direction === -1) {
				const step = resolveShowBackward({
					currentIndex: current,
					showOrder,
					endOfShow: false,
					seededCompleted: false,
				});
				if (step.kind === 'go') {
					// Stepping back onto a slide shows it with its builds already played.
					transitionTo(step.index, { reverse: true, seedCompleted: true });
				}
				return;
			}

			// The shared show step: builds first, then the click-advance gate
			// (`advanceOnClick`), a Zoom excursion's return, and the end of the show.
			const step = resolveShowForward({
				currentIndex: current,
				currentSlide: slides[current],
				showOrder,
				endOfShow: false,
				fromClick: trigger === 'click',
				loop: Boolean(loopContinuously) && !rehearsing,
				endWithBlackSlide: true,
				zoomExcursion: zoomExcursionRef.current,
				advanceBuild: playNextAnimationGroup,
			});
			if (step.kind === 'go') {
				if (step.zoomReturn) {
					zoomExcursionRef.current = step.zoomReturn.excursion;
					transitionTo(step.index, {
						scheduleAutoAdvanceForSlide,
						transitionOverride: step.zoomReturn.transitionOverride,
					});
					return;
				}
				// Record timing for the slide we are leaving (rehearsal mode only).
				if (rehearsing) {
					recordCurrentSlideTime(current);
				}
				transitionTo(step.index);
				return;
			}
			if (step.kind !== 'end-screen' && step.kind !== 'exit') {
				return;
			}
			// Past the last slide. Rehearsal ends with its summary.
			if (rehearsing) {
				recordCurrentSlideTime(current);
				try {
					if (document.fullscreenElement) {
						void document.exitFullscreen().catch(() => {
							/* ignore */
						});
					}
				} catch {
					/* ignore */
				}
				onSetMode('edit');
				setShowRehearsalSummary(true);
				return;
			}
			// The presentation hook decides: a returning custom show, the black
			// end-of-show slide, or exit. Without it, the show stays on its last slide.
			onAdvancePastLastSlide?.();
		},
		[
			clearPresentationTimers,
			loopContinuously,
			onAdvancePastLastSlide,
			onPlayActionSound,
			onStopActionSound,
			onSetActiveSlideIndex,
			onSetMode,
			playNextAnimationGroup,
			presentationSlideIndex,
			presentationTimersRef,
			recordCurrentSlideTime,
			rehearsing,
			seedSlideAnimations,
			startSlideAnimations,
			scheduleAutoAdvanceForSlide,
			setShowRehearsalSummary,
			slides,
			visibleSlideIndexes,
			setPresentationSlideVisible,
			setPresentationSlideIndex,
			setTransitionOverlay,
			zoomExcursionRef,
		],
	);

	// Keep the ref in sync so scheduleAutoAdvanceForSlide always calls the
	// latest version of movePresentationSlide.
	movePresentationSlideRef.current = movePresentationSlide;

	// -----------------------------------------------------------------------
	// Direct slide navigation (for action buttons / slide jumps)
	// -----------------------------------------------------------------------

	const navigateToSlide = useCallback(
		(targetIndex: number, transitionOverride?: PptxSlideTransition) => {
			if (targetIndex < 0 || targetIndex >= slides.length) {
				return;
			}
			if (targetIndex === presentationSlideIndex) {
				return;
			}

			executeSlideTransition(targetIndex, {
				slides,
				currentSlideIndex: presentationSlideIndex,
				onPlayActionSound,
				onStopActionSound,
				setPresentationSlideVisible,
				clearPresentationTimers,
				setPresentationSlideIndex,
				onSetActiveSlideIndex,
				seedSlideAnimations,
				startSlideAnimations,
				scheduleAutoAdvanceForSlide,
				presentationTimersRef,
				setTransitionOverlay,
				// A jump ENTERS the target slide, so PowerPoint plays that slide's
				// transition exactly as it does for a forward advance - hyperlinks,
				// action buttons, Home/End and "type a number + Enter" all animate.
				// Suppressing it here is why a deck navigated by clicking its own
				// on-slide links (this is how the issue #131 reporter drives their
				// wheel menu) appeared to have no morph at all, while the same
				// transition played fine on PageDown.
				playTransition: true,
				// A Slide Zoom / Section Zoom / Summary Zoom click carries its own
				// `zmPr/@transitionDur`, independent of the destination's own
				// authored transition (see `useZoomNavigation`).
				transitionOverride,
			});
		},
		[
			clearPresentationTimers,
			onPlayActionSound,
			onStopActionSound,
			onSetActiveSlideIndex,
			presentationSlideIndex,
			presentationTimersRef,
			seedSlideAnimations,
			startSlideAnimations,
			scheduleAutoAdvanceForSlide,
			slides,
			setPresentationSlideIndex,
			setPresentationSlideVisible,
			setTransitionOverlay,
		],
	);

	// -----------------------------------------------------------------------
	// Presentation action handler (action buttons, hyperlinks, slide jumps)
	// -----------------------------------------------------------------------

	const handlePresentationAction = useCallback(
		(action: PptxAction, elementId?: string) => {
			handlePresentationActionImpl(action, {
				elementId,
				movePresentationSlide,
				navigateToSlide,
				onPlayActionSound,
				onSetMode,
				slidesLength: slides.length,
				onLastViewed,
				onCustomShow,
				onOpenFile,
				onOpenPresentation,
				onPlayMedia,
				onOleVerb,
				onRunProgram,
			});
		},
		[
			movePresentationSlide,
			navigateToSlide,
			onPlayActionSound,
			onSetMode,
			slides.length,
			onLastViewed,
			onCustomShow,
			onOpenFile,
			onOpenPresentation,
			onPlayMedia,
			onOleVerb,
			onRunProgram,
		],
	);

	return {
		movePresentationSlide,
		navigateToSlide,
		handlePresentationAction,
		scheduleAutoAdvanceForSlide,
	};
}
