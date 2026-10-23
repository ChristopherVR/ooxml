/**
 * presentation-show-navigator.ts: the slide show's navigation state machine.
 *
 * "Which slide is up" is not a single number in PowerPoint's model: a forward
 * step may reveal the next animation build instead of changing slide, a
 * backward step may replay the current slide's builds instead of leaving it,
 * running off the end raises an end-of-show screen rather than wrapping, and an
 * authored `p:transition/@advTm` arms a timer re-armed (and cancelled) on every
 * change. It lives here as a plain signal-holding class (no Angular injection
 * context needed) so {@link PresentationOverlayComponent} keeps only view
 * wiring, and the rules above stay in one place.
 */
import { signal } from '@angular/core';
import {
	beginZoomExcursion,
	buildZoomTransitionOverride,
	resolveShowBackward,
	resolveShowForward,
	resolveShowSlideIndexes,
} from 'ooxml-ui/pptx';
import type { ShowForwardStep, ZoomExcursion, ZoomNavigationTarget } from 'ooxml-ui/pptx';
import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';

import type { AnimationPlaybackService } from './animation-playback.service';
import type { PresentationAnnotationsService } from './presentation-annotations.service';
import { clampIndex, firstVisibleIndex, lastVisibleIndex } from './presentation-overlay-helpers';
import type { ActiveShow, AuthoredRange } from './presentation-overlay-helpers';

/** Where a navigation request wants to go. */
export type ShowDirection = 'next' | 'prev' | 'first' | 'last';

/**
 * The outgoing slide plus the incoming slide's transition, played over the new
 * slide. Cleared once the transition overlay reports completion.
 */
export interface ActiveSlideTransition {
	outgoing: PptxSlide;
	transition: PptxSlideTransition;
}

/** Everything the navigator needs from its host component. */
export interface ShowNavigatorDeps {
	slides: () => readonly PptxSlide[];
	/**
	 * The running custom show, when one is active. `slides()` is always the whole
	 * deck: the subset is a navigation rule, not a different slide array, so that
	 * a hidden slide inside the show is still skipped and the editor's index
	 * space never has to be translated.
	 */
	activeCustomShow?: () => ActiveShow;
	/**
	 * The `p:showPr/p:sldRg` range, when the deck opens into a range rather than
	 * the whole deck or a custom show. Applied like `activeCustomShow`: a filter
	 * on the navigable order, not a pre-filtered slide array.
	 */
	authoredRange?: () => AuthoredRange;
	/** The slide at the CURRENT index (a computed over `currentIndex`). */
	currentSlide: () => PptxSlide | undefined;
	showWithAnimation: () => boolean | undefined;
	/**
	 * The slide canvas size (px), for a `p:anim` formula needing the animated
	 * shape's real box (e.g. Grow And Turn's `-#ppt_w/2` fly-in). Optional so a
	 * host constructed before this existed still compiles; omitting it just
	 * keeps the pre-existing fallback behaviour.
	 */
	canvasSize?: () => { width: number; height: number };
	/** The deck's resolved theme colour map, for a scheme-colour (`a:schemeClr`) animation stop. */
	themeColorMap?: () => Readonly<Record<string, string>> | undefined;
	/**
	 * `viewerOptions.advanced.pixelateMosaicAnimation`. `true` plays the
	 * blocky mosaic reveal for `p:animEffect/@filter="pixelate"`; omitted or
	 * `false` (the default, matching PowerPoint's own behaviour) snaps the
	 * element to its end state. See `resolveFilterEffect`'s doc in
	 * `pptx-viewer-shared`'s `animation-filter-effects.ts`.
	 */
	pixelateMosaicAnimation?: () => boolean | undefined;
	playback: AnimationPlaybackService;
	annotations: PresentationAnnotationsService;
	/** Publish a committed index change to the host's `indexChange` output. */
	emitIndex: (index: number) => void;
	/** Ask the host to end the show; the host guards against double-closing. */
	requestClose: () => void;
	/**
	 * File > Options > Advanced > "End with black slide". PowerPoint's default is
	 * ON: running past the last slide raises a black end screen and only the NEXT
	 * forward input ends the show; off ends the show immediately instead.
	 */
	endWithBlackSlide?: () => boolean | undefined;
	/**
	 * Set Up Slide Show > "Loop continuously until 'Esc'": wraps to the show's
	 * first slide instead of raising the end screen (or exiting).
	 */
	loopContinuously?: () => boolean | undefined;
}

export class PresentationShowNavigator {
	/** Zero-based index into `slides()`. */
	readonly currentIndex = signal(0);

	/**
	 * True once the show has run past its last slide and the black end screen is
	 * up: while it is up the next input either goes nowhere (backward) or ends
	 * the show (forward), so a stuck deck must surface this rather than swallow
	 * every advance.
	 */
	readonly endOfShow = signal(false);

	/** Active slide-transition animation, or null when none is playing. */
	readonly activeTransition = signal<ActiveSlideTransition | null>(null);

	/**
	 * Set just before a BACKWARD change so the host's slide effect seeds the
	 * incoming slide as fully built. Read-and-cleared via
	 * {@link takePendingCompletedEntry}.
	 */
	private pendingCompletedEntry = false;

	/** Pending `p:transition/@advTm` auto-advance timer for the current slide. */
	private autoAdvanceTimer: ReturnType<typeof setTimeout> | undefined;

	/** Deck index before each committed change; backs {@link goToLastViewed}. */
	private previousIndex: number | null = null;

	/**
	 * A pending "return to zoom" excursion armed by {@link navigateToZoomTarget},
	 * consumed by a forward `navigate('next')` once the show reaches the end of
	 * the target's range. See `pptx-viewer-shared`'s `zoom-return-navigation`.
	 */
	private zoomExcursion: ZoomExcursion | undefined;

	constructor(private readonly deps: ShowNavigatorDeps) {}

	/** Consume the "entered backward" flag: reads and clears it in one call. */
	takePendingCompletedEntry(): boolean {
		const completed = this.pendingCompletedEntry;
		this.pendingCompletedEntry = false;
		return completed;
	}

	/**
	 * Adopt an index pushed in by the host (its `startIndex` input, which an
	 * audience display drives from the presenter's snapshot). Silent: it does not
	 * echo back through `emitIndex`, which would fight the host for control.
	 */
	syncFromHost(requestedIndex: number): void {
		const count = this.deps.slides().length;
		if (count === 0) {
			return;
		}
		const requested = clampIndex(requestedIndex, count);
		if (requested !== this.currentIndex()) {
			this.currentIndex.set(requested);
			this.deps.annotations.setActiveSlide(requested);
			// A host-forced jump (customShow return-after restoring the origin
			// slide) supersedes the black end screen: the host would not push a
			// new slide if it considered the show over.
			this.endOfShow.set(false);
		}
	}

	/**
	 * Jump back to the slide the show was on immediately before the current one
	 * (`ppaction://hlinkshowjump?jump=lastslideviewed`). A no-op before any
	 * navigation has happened yet.
	 */
	goToLastViewed(): void {
		if (this.previousIndex !== null) {
			this.goToSlide(this.previousIndex);
		}
	}

	/**
	 * (Re)arm PowerPoint's "Advance slide: After <n>" timer, always cancelling
	 * whatever was pending first: a manual advance must never leave a stale
	 * timer running that skips the slide the presenter just moved to.
	 */
	armAutoAdvance(delayMs: number | undefined): void {
		this.clearAutoAdvance();
		if (delayMs === undefined) {
			return;
		}
		this.autoAdvanceTimer = setTimeout(() => {
			this.autoAdvanceTimer = undefined;
			this.navigate('next');
		}, delayMs);
	}

	/** Cancel any pending timed auto-advance (also called on teardown). */
	clearAutoAdvance(): void {
		if (this.autoAdvanceTimer !== undefined) {
			clearTimeout(this.autoAdvanceTimer);
			this.autoAdvanceTimer = undefined;
		}
	}

	navigate(direction: ShowDirection): void {
		const slides = this.deps.slides();
		const count = slides.length;
		if (count === 0) {
			return;
		}
		const current = this.currentIndex();
		const activeShow = this.deps.activeCustomShow?.();
		const authoredRange = this.deps.authoredRange?.();
		// Forward and backward are the shared show-step decisions every binding
		// makes (end screen, builds, Zoom return, end of show, show order).
		if (direction === 'next') {
			this.applyForward(
				resolveShowForward({
					currentIndex: current,
					currentSlide: slides[current],
					showOrder: resolveShowSlideIndexes(slides, activeShow, authoredRange),
					endOfShow: this.endOfShow(),
					loop: this.deps.loopContinuously?.() === true,
					endWithBlackSlide: this.deps.endWithBlackSlide?.() !== false,
					zoomExcursion: this.zoomExcursion,
					advanceBuild: () => this.deps.playback.advance(),
				}),
			);
			return;
		}
		if (direction === 'prev') {
			this.applyBackward(current);
			return;
		}
		// Home goes to the START OF THE SHOW, which is not slide 1 when the author
		// hid it. Clamped anyway so an empty order cannot produce -1.
		const next = clampIndex(
			direction === 'first'
				? firstVisibleIndex(slides, activeShow, authoredRange)
				: lastVisibleIndex(slides, activeShow, authoredRange),
			count,
		);
		if (next === current) {
			return;
		}
		const incoming = slides[next];
		const outgoing = slides[current];
		this.commit(
			next,
			direction === 'first' && incoming?.transition && outgoing
				? { outgoing, transition: incoming.transition }
				: null,
		);
	}

	private applyForward(step: ShowForwardStep): void {
		switch (step.kind) {
			case 'exit':
				// A forward input on the end screen ends the show (PowerPoint's "click
				// to exit"), as does running off the end without a black slide.
				this.endOfShow.set(false);
				this.deps.requestClose();
				return;
			case 'end-screen':
				this.endOfShow.set(true);
				return;
			case 'go':
				if (step.zoomReturn) {
					this.zoomExcursion = step.zoomReturn.excursion;
					this.goToSlide(step.index, step.zoomReturn.transitionOverride);
					return;
				}
				this.goToSlide(step.index);
				break;
			default:
				break;
		}
	}

	private applyBackward(current: number): void {
		const slides = this.deps.slides();
		const step = resolveShowBackward({
			currentIndex: current,
			showOrder: resolveShowSlideIndexes(
				slides,
				this.deps.activeCustomShow?.(),
				this.deps.authoredRange?.(),
			),
			endOfShow: this.endOfShow(),
			seededCompleted: this.deps.playback.isSeededCompleted(),
		});
		if (step.kind === 'dismiss-end-screen') {
			this.endOfShow.set(false);
			return;
		}
		if (step.kind === 'replay') {
			const size = this.deps.canvasSize?.();
			this.deps.playback.setSlide(this.deps.currentSlide(), this.deps.showWithAnimation(), {
				slideHeightPx: size?.height,
				slideWidthPx: size?.width,
				themeColorMap: this.deps.themeColorMap?.(),
				pixelateMosaic: this.deps.pixelateMosaicAnimation?.(),
			});
			return;
		}
		if (step.kind !== 'go') {
			return;
		}
		// PowerPoint shows a slide you step BACK onto with its builds played, and
		// replays the LEAVING slide's transition in reverse (a morph glides its
		// shapes back to where they came from).
		this.pendingCompletedEntry = true;
		const outgoing = slides[current];
		this.commit(
			step.index,
			outgoing?.transition && slides[step.index]
				? { outgoing, transition: outgoing.transition }
				: null,
		);
	}

	/**
	 * Jump directly to `index` (clamped). Used by on-slide Action Settings
	 * (`ppaction://hlinksldjump`) and by {@link navigateToZoomTarget}; ENTERS
	 * the target slide, so PowerPoint plays its transition exactly as a forward
	 * step does.
	 *
	 * @param transitionOverride Play this transition INSTEAD of the
	 *   destination's own authored `<p:transition>`. Set by a Slide Zoom /
	 *   Section Zoom / Summary Zoom navigation that authored its own
	 *   `zmPr/@transitionDur` (see `buildZoomTransitionOverride`).
	 */
	goToSlide(index: number, transitionOverride?: PptxSlideTransition): void {
		const slides = this.deps.slides();
		const count = slides.length;
		if (count === 0) {
			return;
		}
		const next = clampIndex(index, count);
		const current = this.currentIndex();
		if (next === current) {
			return;
		}
		const incoming = slides[next];
		const outgoing = slides[current];
		const transition = transitionOverride
			? outgoing
				? { outgoing, transition: transitionOverride }
				: null
			: incoming?.transition && outgoing
				? { outgoing, transition: incoming.transition }
				: null;
		this.commit(next, transition);
	}

	/**
	 * Handle a Slide Zoom / Section Zoom / Summary Zoom tile click: navigates
	 * to the target slide, playing the zoom's own `zmPr/@transitionDur` when
	 * authored, and (when `returnToParent` is set) arms an excursion so the
	 * next forward `navigate('next')` past the target's range returns to the
	 * slide the zoom was clicked from instead of continuing linearly through
	 * the deck.
	 */
	navigateToZoomTarget(target: ZoomNavigationTarget): void {
		this.zoomExcursion = beginZoomExcursion(target, this.currentIndex(), this.deps.slides());
		this.goToSlide(
			target.targetSlideIndex,
			buildZoomTransitionOverride(target.transitionDurationMs),
		);
	}

	/**
	 * The single place a slide change becomes visible: set the transition, move
	 * the index, retarget the annotation layer, and tell the host.
	 */
	private commit(next: number, transition: ActiveSlideTransition | null): void {
		this.previousIndex = this.currentIndex();
		this.activeTransition.set(transition);
		this.currentIndex.set(next);
		this.deps.annotations.setActiveSlide(next);
		this.deps.emitIndex(next);
	}
}
