// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { applyAnimationGroupSteps } from './animation-playback-engine';
import type { PlaybackContext } from './animation-playback-engine';
import type { ElementAnimationState, TimelineStep } from './animation-timeline-types';

function colorStep(
	colorTargets: TimelineStep['colorTargets'],
	holdEndState: boolean,
): TimelineStep {
	return {
		elementId: 'shape',
		keyframeName: 'pptx-tl-color-hold',
		cssAnimation: 'pptx-tl-color-hold 100ms linear 0ms 1 both',
		trigger: 'onClick',
		delayMs: 0,
		durationMs: 100,
		fillMode: 'both',
		presetClass: 'emph',
		colorTargets,
		holdEndState,
	};
}

function play(steps: TimelineStep[]) {
	let states = new Map<string, ElementAnimationState>();
	const context: PlaybackContext = {
		setStates: (update) => {
			states = update(states);
		},
		timers: [],
		buildHandle: { current: null },
		playSound: vi.fn(),
		stopSound: vi.fn(),
	};
	applyAnimationGroupSteps({ steps, totalDurationMs: 300 }, context);
	return () => states.get('shape');
}

describe('held animation paint ownership', () => {
	beforeEach(() => vi.useFakeTimers());
	afterEach(() => vi.useRealTimers());

	it.each<{ targets: NonNullable<TimelineStep['colorTargets']> }>([
		{ targets: ['fill'] },
		{ targets: ['stroke'] },
		{ targets: ['fill', 'stroke'] },
	])('keeps $targets paint animated while its final frame is held', ({ targets }) => {
		const step = colorStep([...targets], true);
		const current = play([step]);
		vi.advanceTimersByTime(110);
		expect(current()?.cssAnimation).toBe(step.cssAnimation);
		expect(current()?.animatesFill).toBe(targets.includes('fill') || undefined);
		expect(current()?.animatesStroke).toBe(targets.includes('stroke') || undefined);
	});

	it('retains held color ownership when a disjoint motion component finishes', () => {
		const held = colorStep(['fill'], true);
		const current = play([
			held,
			{
				...colorStep(undefined, false),
				keyframeName: 'pptx-tl-motion-end',
				cssAnimation: 'pptx-tl-motion-end 300ms linear 0ms 1 both',
				durationMs: 300,
			},
		]);
		vi.advanceTimersByTime(310);
		expect(current()?.cssAnimation).toBe(held.cssAnimation);
		expect(current()?.animatesFill).toBe(true);
		expect(current()?.animatesStroke).toBeUndefined();
	});

	it('restores static paint when no color animation remains', () => {
		const current = play([colorStep(['fill', 'stroke'], false)]);
		vi.advanceTimersByTime(110);
		expect(current()?.cssAnimation).toBeUndefined();
		expect(current()?.animatesFill).toBeUndefined();
		expect(current()?.animatesStroke).toBeUndefined();
	});
});
