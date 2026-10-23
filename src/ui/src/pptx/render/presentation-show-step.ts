import type { PptxSlide, PptxSlideTransition } from 'ooxml-core/pptx';

import { isClickAdvanceAllowed } from './presentation-setup';
import {
	firstShowSlideIndex,
	hasShowSlideAfter,
	nextShowSlideIndex,
	previousShowSlideIndex,
} from './presentation-show-order';
import type { ZoomExcursion } from './zoom-return-navigation';
import {
	buildZoomTransitionOverride,
	resolveForwardSlideWithZoomReturn,
} from './zoom-return-navigation';

/**
 * Where a slide show's forward or backward input lands, decided once for all
 * five bindings. Each binding ran the same sequence of shared calls in its own
 * order (Angular's navigator, Svelte's presentation controller, Vue's
 * navigation composable, vanilla's controls and React's slide navigation):
 *
 * forward: the black end screen ends the show; otherwise the next animation
 * build is revealed; a click past the last build does nothing on a slide whose
 * transition forbids click-advance; a pending Zoom excursion returns to its
 * zoom slide; past the show's last slide a returning sub-show resumes, a
 * looping show wraps, the black end screen rises, or the show ends; otherwise
 * the next show slide (hidden slides skipped, custom show and range honoured).
 *
 * backward: the end screen is dismissed; a slide entered backward replays its
 * builds; otherwise the previous show slide, entered with its builds complete.
 *
 * The functions only decide; the binding performs the answer (its own index,
 * transition overlay, timers and end screen), so it keeps its own state shape.
 */

/** What a forward input needs to know about the running show. */
export interface ShowForwardInput {
	/** The deck index on screen. */
	currentIndex: number;
	/** The slide at `currentIndex`, for the click-advance gate. */
	currentSlide: PptxSlide | undefined;
	/** The deck indexes the show visits, in order (`resolveShowSlideIndexes`). */
	showOrder: readonly number[];
	/** The black "End of slide show" screen is up. */
	endOfShow: boolean;
	/** A click, tap or swipe on the slide (keyboard and buttons are never gated). */
	fromClick?: boolean;
	/** Set Up Slide Show > "Loop continuously until 'Esc'" (or a kiosk show). */
	loop: boolean;
	/** File > Options > Advanced > "End with black slide" (PowerPoint's default is on). */
	endWithBlackSlide: boolean;
	/** The pending "return to zoom" excursion, if a Zoom with `returnToParent` armed one. */
	zoomExcursion?: ZoomExcursion | undefined;
	/**
	 * Called only past the show's last slide: resume the show a returning
	 * `ppaction://customshow?...&return=true` sub-show interrupted. True when it did.
	 */
	returnFromSubShow?(): boolean;
	/** Reveal the next animation build; true when one was revealed (stay on the slide). */
	advanceBuild(): boolean;
}

export type ShowForwardStep =
	/** End the show (forward on the end screen, or past the end without a black slide). */
	| { kind: 'exit' }
	/** Nothing to do (a build was revealed, a click was gated, or there is no slide). */
	| { kind: 'stay' }
	/** Raise the black end-of-show screen. */
	| { kind: 'end-screen' }
	/** A returning sub-show ran off its end and `returnFromSubShow` resumed its origin. */
	| { kind: 'resumed' }
	/**
	 * Go to `index`. A Zoom return also carries the excursion to keep (cleared)
	 * and the zoom's own transition, when it authored one.
	 */
	| {
			kind: 'go';
			index: number;
			zoomReturn?: {
				excursion: ZoomExcursion | undefined;
				transitionOverride: PptxSlideTransition | undefined;
			};
	  };

export function resolveShowForward(input: ShowForwardInput): ShowForwardStep {
	if (input.endOfShow) {
		return { kind: 'exit' };
	}
	if (input.advanceBuild()) {
		return { kind: 'stay' };
	}
	if (input.fromClick && !isClickAdvanceAllowed(input.currentSlide)) {
		return { kind: 'stay' };
	}
	const current = input.currentIndex;
	const excursion = input.zoomExcursion;
	const zoomStep = resolveForwardSlideWithZoomReturn(current, undefined, excursion);
	if (zoomStep.returnedToZoom && zoomStep.nextSlideIndex !== undefined) {
		return {
			kind: 'go',
			index: zoomStep.nextSlideIndex,
			zoomReturn: {
				excursion: zoomStep.excursion,
				transitionOverride: buildZoomTransitionOverride(excursion?.transitionDurationMs),
			},
		};
	}
	const order = input.showOrder;
	if (!hasShowSlideAfter(current, order)) {
		if (input.returnFromSubShow?.()) {
			return { kind: 'resumed' };
		}
		if (input.loop) {
			const first = firstShowSlideIndex(order);
			return first === undefined || first === current
				? { kind: 'stay' }
				: { kind: 'go', index: first };
		}
		return input.endWithBlackSlide ? { kind: 'end-screen' } : { kind: 'exit' };
	}
	const next = nextShowSlideIndex(current, order);
	return next === undefined || next === current ? { kind: 'stay' } : { kind: 'go', index: next };
}

/** What a backward input needs to know about the running show. */
export interface ShowBackwardInput {
	currentIndex: number;
	showOrder: readonly number[];
	endOfShow: boolean;
	/** The slide on screen was entered backward and shows its builds complete. */
	seededCompleted: boolean;
}

export type ShowBackwardStep =
	/** Dismiss the black end screen (back onto the last slide). */
	| { kind: 'dismiss-end-screen' }
	/** Replay the current slide's builds from the start. */
	| { kind: 'replay' }
	/** At the start of the show: stay (PowerPoint never wraps backward). */
	| { kind: 'stay' }
	/** Go to `index`, entering it with its builds already played. */
	| { kind: 'go'; index: number };

export function resolveShowBackward(input: ShowBackwardInput): ShowBackwardStep {
	if (input.endOfShow) {
		return { kind: 'dismiss-end-screen' };
	}
	if (input.seededCompleted) {
		return { kind: 'replay' };
	}
	const previous = previousShowSlideIndex(input.currentIndex, input.showOrder);
	return previous === undefined || previous === input.currentIndex
		? { kind: 'stay' }
		: { kind: 'go', index: previous };
}
