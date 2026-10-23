import type { PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { resolveShowBackward, resolveShowForward } from './presentation-show-step';
import type { ShowForwardInput } from './presentation-show-step';

const slide = (advanceOnClick?: boolean): PptxSlide =>
	({
		id: 's',
		slideNumber: 1,
		elements: [],
		...(advanceOnClick === undefined ? {} : { transition: { type: 'fade', advanceOnClick } }),
	}) as unknown as PptxSlide;

/** Deck of five; slide 2 hidden, so the show visits 0, 1, 3, 4. */
const order = [0, 1, 3, 4];

const forward = (over: Partial<ShowForwardInput> = {}) =>
	resolveShowForward({
		currentIndex: 1,
		currentSlide: slide(),
		showOrder: order,
		endOfShow: false,
		loop: false,
		endWithBlackSlide: true,
		advanceBuild: () => false,
		...over,
	});

describe('resolveShowForward', () => {
	it('steps to the next show slide, skipping a hidden one', () => {
		expect(forward()).toStrictEqual({ kind: 'go', index: 3 });
	});

	it('reveals a remaining build before changing slide', () => {
		expect(forward({ advanceBuild: () => true })).toStrictEqual({ kind: 'stay' });
	});

	it('ends the show on a forward input while the end screen is up, without stepping builds', () => {
		let stepped = false;
		const step = forward({
			endOfShow: true,
			advanceBuild: () => {
				stepped = true;
				return true;
			},
		});
		expect(step).toStrictEqual({ kind: 'exit' });
		expect(stepped).toBeFalsy();
	});

	it('gates a click past the last build when the transition forbids click-advance', () => {
		expect(forward({ fromClick: true, currentSlide: slide(false) })).toStrictEqual({
			kind: 'stay',
		});
		expect(forward({ fromClick: false, currentSlide: slide(false) }).kind).toBe('go');
		expect(forward({ fromClick: true, currentSlide: slide(true) }).kind).toBe('go');
	});

	it('returns to the zoom slide at the end of an excursion, with the zoom transition', () => {
		const step = forward({
			currentIndex: 3,
			zoomExcursion: { returnSlideIndex: 0, endSlideIndex: 3, transitionDurationMs: 700 },
		});
		expect(step).toStrictEqual({
			kind: 'go',
			index: 0,
			zoomReturn: { excursion: undefined, transitionOverride: { type: 'zoom', durationMs: 700 } },
		});
	});

	it('past the last slide: sub-show return, then loop, then end screen, then exit', () => {
		expect(forward({ currentIndex: 4, returnFromSubShow: () => true, loop: true })).toStrictEqual({
			kind: 'resumed',
		});
		expect(forward({ currentIndex: 4, returnFromSubShow: () => false })).toStrictEqual({
			kind: 'end-screen',
		});
		expect(forward({ currentIndex: 4, loop: true })).toStrictEqual({ kind: 'go', index: 0 });
		expect(forward({ currentIndex: 4 })).toStrictEqual({ kind: 'end-screen' });
		expect(forward({ currentIndex: 4, endWithBlackSlide: false })).toStrictEqual({
			kind: 'exit',
		});
	});
});

describe('resolveShowBackward', () => {
	const backward = (over: Partial<Parameters<typeof resolveShowBackward>[0]> = {}) =>
		resolveShowBackward({
			currentIndex: 3,
			showOrder: order,
			endOfShow: false,
			seededCompleted: false,
			...over,
		});

	it('steps back over a hidden slide', () => {
		expect(backward()).toStrictEqual({ kind: 'go', index: 1 });
	});

	it('dismisses the end screen first, then replays a slide entered backward', () => {
		expect(backward({ endOfShow: true, seededCompleted: true })).toStrictEqual({
			kind: 'dismiss-end-screen',
		});
		expect(backward({ seededCompleted: true })).toStrictEqual({ kind: 'replay' });
	});

	it('stays at the start of the show (never wraps backward)', () => {
		expect(backward({ currentIndex: 0 })).toStrictEqual({ kind: 'stay' });
	});
});
