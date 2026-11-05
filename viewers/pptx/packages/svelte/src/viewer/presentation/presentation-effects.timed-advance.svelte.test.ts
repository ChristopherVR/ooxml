import type { PptxSlide } from 'pptx-viewer-core';
import { flushSync } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { PresentationController } from './presentation-controller.svelte';
import { usePresentationEffects } from './presentation-effects.svelte';

/** Only what `usePresentationEffects` reads; no stage root, so no DOM wiring. */
function stubController(advance: () => void): PresentationController {
	return {
		start: vi.fn(),
		stop: vi.fn(),
		onSlideChange: vi.fn(),
		advance,
		elementStates: new Map(),
		interactiveTriggerShapeIds: new Set(),
		hoverTriggerShapeIds: new Set(),
		keyframesCss: '',
		endOfShowVisible: false,
	} as unknown as PresentationController;
}

// Issue #141: PowerPoint counts a slide's "After" time from the end of its own
// transition. Arming the timer on entry let a slide set to advance after 0 s
// move on at once, so the next slide's transition replaced its 700 ms fade and
// the slide appeared in a single frame.
describe('usePresentationEffects timed advance', () => {
	let destroy: (() => void) | undefined;

	beforeEach(() => {
		vi.useFakeTimers();
	});

	afterEach(() => {
		destroy?.();
		destroy = undefined;
		vi.useRealTimers();
	});

	it('waits out the entry transition before a 0 s timing advances', () => {
		const advance = vi.fn();
		const slide = {
			id: 's6',
			elements: [],
			transition: { type: 'fade', durationMs: 700, advanceOnClick: false, advanceAfterMs: 0 },
		} as unknown as PptxSlide;
		destroy = $effect.root(() => {
			usePresentationEffects({
				controller: stubController(advance),
				getPresenting: () => true,
				getCurrentIndex: () => 5,
				getActiveSlide: () => slide,
				getStageRoot: () => null,
			});
		});
		flushSync();

		vi.advanceTimersByTime(699);
		expect(advance).not.toHaveBeenCalled();
		vi.advanceTimersByTime(1);
		expect(advance).toHaveBeenCalledOnce();
	});

	it('adds the transition to the authored time', () => {
		const advance = vi.fn();
		const slide = {
			id: 's2',
			elements: [],
			transition: { type: 'fade', durationMs: 500, advanceAfterMs: 2000 },
		} as unknown as PptxSlide;
		destroy = $effect.root(() => {
			usePresentationEffects({
				controller: stubController(advance),
				getPresenting: () => true,
				getCurrentIndex: () => 1,
				getActiveSlide: () => slide,
				getStageRoot: () => null,
			});
		});
		flushSync();

		vi.advanceTimersByTime(2499);
		expect(advance).not.toHaveBeenCalled();
		vi.advanceTimersByTime(1);
		expect(advance).toHaveBeenCalledOnce();
	});
});
