/**
 * Builds the open workbook's formula dependency graph while the editor is idle. Opening a large
 * workbook paints without it, and without this the first edit would wait for every formula to be
 * parsed and indexed (half a second on a 66,000-formula workbook). After the first paint the work
 * runs in `requestIdleCallback` slices of at most 8 ms (a `setTimeout` fallback elsewhere),
 * pauses while the user types, scrolls or edits a cell, and stops when another workbook loads.
 * An edit made before it finishes is still correct: the core completes what is left first.
 */
import type { EditSession } from 'ooxml-core/xlsx';

/** The longest slice of work, so a frame is never held up noticeably. */
export const SLICE_MS = 8;
/** How long after the last key, wheel or pointer event the work stays paused. */
export const QUIET_MS = 250;

/** What a slice is told about the time it has (the part of `IdleDeadline` used here). */
export interface IdleTime {
	timeRemaining(): number;
}

/** The timers the preparation runs on (the window's, replaced in tests). */
export interface IdleScheduler {
	now(): number;
	idle(run: (deadline: IdleTime) => void): () => void;
	later(run: () => void, ms: number): () => void;
	/** Runs `run` once the next frame has been painted. */
	afterPaint(run: () => void): () => void;
}

/** Events that mean the user is busy with the grid. */
const ACTIVITY = ['keydown', 'wheel', 'pointerdown', 'touchmove'] as const;

export function windowScheduler(view: (Window & typeof globalThis) | null): IdleScheduler {
	const clock = () => (view?.performance ?? performance).now();
	const later = (run: () => void, ms: number) => {
		const id = setTimeout(run, ms);
		return () => clearTimeout(id);
	};
	return {
		now: clock,
		later,
		afterPaint(run) {
			if (!view || typeof view.requestAnimationFrame !== 'function') return later(run, 0);
			let cancelTimer: (() => void) | undefined;
			const id = view.requestAnimationFrame(() => {
				cancelTimer = later(run, 0);
			});
			return () => {
				view.cancelAnimationFrame(id);
				cancelTimer?.();
			};
		},
		idle(run) {
			if (view && typeof view.requestIdleCallback === 'function') {
				const id = view.requestIdleCallback(run);
				return () => view.cancelIdleCallback(id);
			}
			const start = clock();
			return later(() => run({ timeRemaining: () => SLICE_MS - (clock() - start) }), 1);
		},
	};
}

export class CalcPreparation {
	private session: EditSession | undefined;
	private cancel: (() => void) | undefined;
	private lastActivity = Number.NEGATIVE_INFINITY;

	constructor(
		host: HTMLElement,
		/** True while a cell is being edited: no slice runs then. */
		private readonly busy: () => boolean,
		private readonly scheduler: IdleScheduler = windowScheduler(host.ownerDocument.defaultView),
	) {
		const active = () => {
			this.lastActivity = this.scheduler.now();
		};
		for (const type of ACTIVITY) host.addEventListener(type, active, { passive: true });
	}

	/** Prepares `session` (the workbook just opened) after the next paint; stops earlier work. */
	start(session: EditSession | undefined): void {
		this.stop();
		if (!session) return;
		this.session = session;
		this.cancel = this.scheduler.afterPaint(() => this.next());
	}

	stop(): void {
		this.cancel?.();
		this.cancel = undefined;
		this.session = undefined;
	}

	/** Whether work is still waiting (tests and diagnostics). */
	get pending(): boolean {
		return this.session !== undefined;
	}

	private next(): void {
		this.cancel = this.scheduler.idle((deadline) => this.slice(deadline));
	}

	private slice(deadline: IdleTime): void {
		this.cancel = undefined;
		const session = this.session;
		if (!session) return;
		const quietFor = this.scheduler.now() - this.lastActivity;
		if (this.busy() || quietFor < QUIET_MS) {
			const wait = Math.max(QUIET_MS - quietFor, 50);
			this.cancel = this.scheduler.later(() => this.next(), wait);
			return;
		}
		const end = this.scheduler.now() + Math.min(SLICE_MS, Math.max(deadline.timeRemaining(), 1));
		let done = true;
		try {
			done = session.prepareCalculation({ timeRemaining: () => end - this.scheduler.now() });
		} catch {
			// The first edit builds the graph itself and reports any failure there.
		}
		if (done) this.session = undefined;
		else this.next();
	}
}
