/**
 * Rehearse Timings (Slide Show > Rehearse Timings) as a headless controller: the per-slide and
 * total clocks, pause and resume, the recorded time of each slide and the summary that offers
 * to keep them.
 *
 * Every binding used to keep its own copy of this state machine, and they had drifted (see
 * `docs/structure-plan.md` section 5): one floored each slide at one second, one at 250 ms; one
 * counted paused time into the total, one subtracted a pause twice; one also switched "On mouse
 * click" on when saving. The one behaviour is now:
 *
 *  - a slide's time is the time it was shown, less the time the clock was paused while it was,
 *    rounded to whole milliseconds (no floor);
 *  - the total clock also stops while paused;
 *  - saving writes `advanceAfterMs` through {@link applyRehearsalTimings} and leaves
 *    "On mouse click" as the slide had it.
 *
 * The controller owns its display tick (every 250 ms while recording and not paused), so a
 * template reads `slideElapsedMs` and `totalElapsedMs` from the snapshot and keeps no timer.
 *
 * @module render/controller/presentation-rehearsal-controller
 */

import { createControllerStore } from './controller';
import type { PptxController } from './controller';

/** The rehearsal's state; replaced (never mutated) on every change. */
export interface PresentationRehearsalState {
	/** A rehearsal is recording. */
	readonly active: boolean;
	/** The clocks are paused. */
	readonly paused: boolean;
	/** The "keep the new timings?" summary is open. */
	readonly summaryOpen: boolean;
	/** The slide being timed (a deck index). */
	readonly slideIndex: number;
	/** Recorded time per slide index, in milliseconds. */
	readonly timings: Readonly<Record<number, number>>;
	/** Time on the current slide, paused time excluded, as of the last tick. */
	readonly slideElapsedMs: number;
	/** Time since the rehearsal started, paused time excluded, as of the last tick. */
	readonly totalElapsedMs: number;
	/** Sum of {@link timings}. */
	readonly recordedTotalMs: number;
}

/** Interval seam; defaults to the global `setInterval` / `clearInterval`. */
export interface RehearsalTimers {
	setInterval(handler: () => void, ms: number): unknown;
	clearInterval(id: unknown): void;
}

export interface PresentationRehearsalControllerOptions {
	/** Clock in milliseconds; defaults to `Date.now`. */
	now?(): number;
	timers?: RehearsalTimers;
	/** Display tick period; defaults to 250 ms. */
	tickMs?: number;
}

export interface PresentationRehearsalController extends PptxController<PresentationRehearsalState> {
	/** Begin recording on `slideIndex`, discarding any earlier timings. */
	start(slideIndex: number): void;
	/** The show moved to `slideIndex`: record the slide being left and time the new one. */
	moveTo(slideIndex: number): void;
	togglePause(): void;
	/** The show ended: record the current slide and open the summary. No-op when idle. */
	finish(): void;
	/** Close the summary and return the timings to save (apply with `applyRehearsalTimings`). */
	save(): Readonly<Record<number, number>>;
	/** Close the summary and drop the timings. */
	discard(): void;
	/** Stop the tick; the controller can still be read. */
	dispose(): void;
}

const IDLE: PresentationRehearsalState = {
	active: false,
	paused: false,
	summaryOpen: false,
	slideIndex: 0,
	timings: {},
	slideElapsedMs: 0,
	totalElapsedMs: 0,
	recordedTotalMs: 0,
};

const defaultTimers: RehearsalTimers = {
	setInterval: (handler, ms) => setInterval(handler, ms),
	clearInterval: (id) => clearInterval(id as ReturnType<typeof setInterval>),
};

function sum(timings: Readonly<Record<number, number>>): number {
	return Object.values(timings).reduce((total, ms) => total + ms, 0);
}

/** Create the Rehearse Timings controller (idle until {@link PresentationRehearsalController.start}). */
export function createPresentationRehearsalController(
	options: PresentationRehearsalControllerOptions = {},
): PresentationRehearsalController {
	const now = options.now ?? Date.now;
	const timers = options.timers ?? defaultTimers;
	const tickMs = options.tickMs ?? 250;
	const store = createControllerStore<PresentationRehearsalState>(IDLE);

	let startedAt = 0;
	let slideStartedAt = 0;
	let pausedAt: number | null = null;
	let slidePausedMs = 0;
	let totalPausedMs = 0;
	let interval: unknown = null;

	const openPause = (at: number): number => (pausedAt === null ? 0 : at - pausedAt);
	const slideElapsed = (at: number): number =>
		Math.max(0, Math.round(at - slideStartedAt - slidePausedMs - openPause(at)));
	const totalElapsed = (at: number): number =>
		Math.max(0, Math.round(at - startedAt - totalPausedMs - openPause(at)));

	const stopTick = (): void => {
		if (interval !== null) {
			timers.clearInterval(interval);
			interval = null;
		}
	};
	const tick = (): void => {
		const state = store.getSnapshot();
		if (!state.active) {
			return;
		}
		const at = now();
		const slideElapsedMs = slideElapsed(at);
		const totalElapsedMs = totalElapsed(at);
		if (slideElapsedMs !== state.slideElapsedMs || totalElapsedMs !== state.totalElapsedMs) {
			store.set({ ...state, slideElapsedMs, totalElapsedMs });
		}
	};
	const startTick = (): void => {
		stopTick();
		interval = timers.setInterval(tick, tickMs);
	};
	const record = (state: PresentationRehearsalState, at: number) => {
		const timings = { ...state.timings, [state.slideIndex]: slideElapsed(at) };
		return { timings, recordedTotalMs: sum(timings) };
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		start(slideIndex) {
			const at = now();
			startedAt = at;
			slideStartedAt = at;
			pausedAt = null;
			slidePausedMs = 0;
			totalPausedMs = 0;
			store.set({ ...IDLE, active: true, slideIndex });
			startTick();
		},
		moveTo(slideIndex) {
			const state = store.getSnapshot();
			if (!state.active || slideIndex === state.slideIndex) {
				return;
			}
			const at = now();
			const recorded = record(state, at);
			// A pause carries over to the next slide: its clock starts paused.
			totalPausedMs += openPause(at);
			pausedAt = pausedAt === null ? null : at;
			slideStartedAt = at;
			slidePausedMs = 0;
			store.set({
				...state,
				...recorded,
				slideIndex,
				slideElapsedMs: 0,
				totalElapsedMs: totalElapsed(at),
			});
		},
		togglePause() {
			const state = store.getSnapshot();
			if (!state.active) {
				return;
			}
			const at = now();
			if (pausedAt === null) {
				pausedAt = at;
				stopTick();
				store.set({
					...state,
					paused: true,
					slideElapsedMs: slideElapsed(at),
					totalElapsedMs: totalElapsed(at),
				});
				return;
			}
			const pause = at - pausedAt;
			slidePausedMs += pause;
			totalPausedMs += pause;
			pausedAt = null;
			store.set({ ...state, paused: false });
			startTick();
		},
		finish() {
			const state = store.getSnapshot();
			if (!state.active) {
				return;
			}
			stopTick();
			const at = now();
			store.set({
				...state,
				...record(state, at),
				active: false,
				paused: false,
				summaryOpen: true,
				totalElapsedMs: totalElapsed(at),
			});
			pausedAt = null;
		},
		save() {
			const { timings } = store.getSnapshot();
			stopTick();
			store.set(IDLE);
			return timings;
		},
		discard() {
			stopTick();
			store.set(IDLE);
		},
		dispose: stopTick,
	};
}
