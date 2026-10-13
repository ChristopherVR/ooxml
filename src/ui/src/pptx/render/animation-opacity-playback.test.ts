// @vitest-environment jsdom
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import type { PlaybackContext } from './animation-playback-engine';
import { applyAnimationGroupSteps } from './animation-playback-engine';
import { clearPlaybackTimers } from './animation-playback-seek';
import { finishOpacitySetSteps } from './animation-opacity-playback';
import type {
	ElementAnimationState,
	TimelineClickGroup,
	TimelineStep,
} from './animation-timeline-types';

beforeEach(() => vi.useFakeTimers());
afterEach(() => {
	vi.clearAllTimers();
	vi.useRealTimers();
});

function setup() {
	let states = new Map<string, ElementAnimationState>();
	const ctx: PlaybackContext = {
		setStates: (update) => {
			states = update(states);
		},
		timers: [],
		buildHandle: { current: null },
		playSound: vi.fn(),
		stopSound: vi.fn(),
	};
	const play = (value: number, durationMs: number, holdEndState: boolean, delayMs = 0) => {
		const step: TimelineStep = {
			elementId: 's',
			cssAnimation: '',
			keyframeName: '',
			trigger: 'onClick',
			delayMs,
			durationMs,
			fillMode: 'both',
			presetClass: 'emph',
			opacitySet: { value, durationMs, delayMs: 0, holdEndState },
		};
		const group: TimelineClickGroup = { steps: [step], totalDurationMs: delayMs + durationMs };
		applyAnimationGroupSteps(group, ctx);
		return group;
	};
	return { ctx, play, state: () => states.get('s') };
}

it('sets the fixed value at its delayed start and restores at removal', () => {
	const test = setup();
	test.play(0.4, 200, false, 100);
	expect(test.state()?.opacity).toBeUndefined();
	vi.advanceTimersByTime(100);
	expect(test.state()?.opacity).toBe(0.4);
	expect(test.state()?.cssAnimation).toBeUndefined();
	vi.advanceTimersByTime(200);
	expect(test.state()?.opacity).toBeUndefined();
});
it('keeps held values through CSS cleanup and restores a surviving held predecessor', () => {
	const test = setup();
	test.play(0.2, 100, true);
	vi.advanceTimersByTime(108);
	expect(test.state()?.opacity).toBe(0.2);
	test.play(0.4, 200, false);
	expect(test.state()?.opacity).toBe(0.4);
	vi.advanceTimersByTime(200);
	expect(test.state()?.opacity).toBe(0.2);
});
it('retains a held zero value when a later temporary set restores its predecessor', () => {
	const test = setup();
	test.play(0, 100, true);
	test.play(0.4, 100, false);
	vi.advanceTimersByTime(100);
	expect(test.state()?.opacity).toBe(0);
});
it('does not restore an already expired overlapping predecessor', () => {
	const test = setup();
	test.play(0.2, 200, false);
	vi.advanceTimersByTime(100);
	test.play(0.4, 200, false);
	vi.advanceTimersByTime(100);
	expect(test.state()?.opacity).toBe(0.4);
	vi.advanceTimersByTime(100);
	expect(test.state()?.opacity).toBeUndefined();
});
it('seek settles delayed held sets, and cancel prevents stale delayed writes', () => {
	const test = setup();
	const group = test.play(0.2, 200, true, 1000);
	finishOpacitySetSteps(group, test.ctx);
	clearPlaybackTimers(test.ctx);
	expect(test.state()?.opacity).toBe(0.2);
	test.play(0.4, 200, true, 1000);
	clearPlaybackTimers(test.ctx);
	vi.advanceTimersByTime(1200);
	expect(test.state()?.opacity).toBe(0.2);
});
