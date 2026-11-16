import { describe, expect, it } from 'vitest';

import { createPresentationRehearsalController } from './presentation-rehearsal-controller';
import type { RehearsalTimers } from './presentation-rehearsal-controller';

function setup() {
	let clock = 1_000;
	let handler: (() => void) | null = null;
	const timers: RehearsalTimers = {
		setInterval: (next) => {
			handler = next;
			return 1;
		},
		clearInterval: () => {
			handler = null;
		},
	};
	const controller = createPresentationRehearsalController({ now: () => clock, timers });
	const advance = (ms: number): void => {
		clock += ms;
		handler?.();
	};
	return { controller, advance, ticking: () => handler !== null };
}

describe('createPresentationRehearsalController', () => {
	it('records the time of each slide and the total', () => {
		const { controller, advance } = setup();
		controller.start(2);
		advance(1500);
		expect(controller.getSnapshot()).toMatchObject({
			active: true,
			slideIndex: 2,
			slideElapsedMs: 1500,
			totalElapsedMs: 1500,
		});
		controller.moveTo(3);
		advance(700);
		controller.finish();
		const state = controller.getSnapshot();
		expect(state.timings).toEqual({ 2: 1500, 3: 700 });
		expect(state.recordedTotalMs).toBe(2200);
		expect(state).toMatchObject({ active: false, summaryOpen: true, totalElapsedMs: 2200 });
	});

	it('stops both clocks while paused, across a slide change', () => {
		const { controller, advance, ticking } = setup();
		controller.start(0);
		advance(1000);
		controller.togglePause();
		expect(ticking()).toBe(false);
		advance(5000);
		controller.moveTo(1);
		advance(4000);
		controller.togglePause();
		advance(500);
		controller.finish();
		const state = controller.getSnapshot();
		expect(state.timings).toEqual({ 0: 1000, 1: 500 });
		expect(state.totalElapsedMs).toBe(1500);
	});

	it('records a very short slide as it was, with no floor', () => {
		const { controller, advance } = setup();
		controller.start(0);
		advance(40);
		controller.finish();
		expect(controller.getSnapshot().timings).toEqual({ 0: 40 });
	});

	it('ignores moves to the same slide and actions while idle', () => {
		const { controller, advance } = setup();
		controller.moveTo(1);
		controller.togglePause();
		controller.finish();
		expect(controller.getSnapshot().summaryOpen).toBe(false);
		controller.start(0);
		advance(100);
		controller.moveTo(0);
		expect(controller.getSnapshot().timings).toEqual({});
	});

	it('save returns the timings and closes; discard drops them', () => {
		const { controller, advance } = setup();
		controller.start(0);
		advance(900);
		controller.finish();
		expect(controller.save()).toEqual({ 0: 900 });
		expect(controller.getSnapshot()).toMatchObject({ summaryOpen: false, timings: {} });
		controller.start(0);
		advance(100);
		controller.finish();
		controller.discard();
		expect(controller.getSnapshot()).toMatchObject({ summaryOpen: false, timings: {} });
	});

	it('restarting discards earlier timings', () => {
		const { controller, advance } = setup();
		controller.start(0);
		advance(100);
		controller.moveTo(1);
		controller.start(4);
		expect(controller.getSnapshot()).toMatchObject({ slideIndex: 4, timings: {}, paused: false });
	});
});
