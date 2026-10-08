// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import {
	createEditSession,
	createWorkbook,
	getCell,
	putCell,
	type Workbook,
} from 'ooxml-core/xlsx';
import { CalcPreparation, type IdleScheduler, type IdleTime, QUIET_MS } from './calc-preparation';

/** A hand-driven clock and task queue standing in for the window's timers. */
function fakeScheduler() {
	let clock = 0;
	let tasks: { at: number; idle: boolean; run: (deadline: IdleTime) => void }[] = [];
	const add = (task: (typeof tasks)[number]) => {
		tasks.push(task);
		return () => {
			tasks = tasks.filter((t) => t !== task);
		};
	};
	const scheduler: IdleScheduler = {
		// Every reading moves the clock on, so work in a slice uses up its time.
		now: () => ++clock,
		later: (run, ms) => add({ at: clock + ms, idle: false, run: () => run() }),
		afterPaint: (run) => add({ at: clock + 16, idle: false, run: () => run() }),
		idle: (run) => add({ at: clock, idle: true, run }),
	};
	return {
		scheduler,
		advance(ms: number) {
			clock += ms;
		},
		get queued() {
			return tasks.length;
		},
		/** Runs the next due task; idle slices get a spent deadline (one slice of work each). */
		step(): boolean {
			const due = tasks.filter((t) => t.at <= clock).sort((a, b) => a.at - b.at)[0];
			const next = due ?? [...tasks].sort((a, b) => a.at - b.at)[0];
			if (!next) return false;
			if (next.at > clock) clock = next.at;
			tasks = tasks.filter((t) => t !== next);
			next.run({ timeRemaining: () => 0 });
			return true;
		},
	};
}

/** 200 rows of `B = A*10` whose stored results are stale, as a file loads them. */
function opened(): Workbook {
	const wb = createWorkbook();
	const sheet = wb.sheets[0]!;
	for (let r = 0; r < 200; r++) {
		putCell(sheet, r, 0, { value: r });
		putCell(sheet, r, 1, { formula: `A${r + 1}*10`, value: -1, legacyFormula: true });
	}
	return wb;
}

describe('CalcPreparation', () => {
	it('prepares the graph in idle slices until it is done', () => {
		const timers = fakeScheduler();
		const host = document.createElement('div');
		const preparation = new CalcPreparation(host, () => false, timers.scheduler);
		const session = createEditSession(opened());
		preparation.start(session);
		let steps = 0;
		while (timers.step()) steps++;
		expect(steps).toBeGreaterThan(3);
		expect(preparation.pending).toBe(false);
		expect(session.prepareCalculation({ timeRemaining: () => 0 })).toBe(true);
	});

	it('pauses while the user types, scrolls or edits a cell', () => {
		const timers = fakeScheduler();
		const host = document.createElement('div');
		let editing = true;
		const preparation = new CalcPreparation(host, () => editing, timers.scheduler);
		const session = createEditSession(opened());
		let slices = 0;
		preparation.start({
			...session,
			prepareCalculation: (options) => {
				slices++;
				return session.prepareCalculation(options);
			},
		});
		for (let i = 0; i < 20; i++) timers.step();
		expect(slices).toBe(0);
		editing = false;
		host.dispatchEvent(new Event('wheel'));
		timers.step();
		timers.step();
		// Still within the quiet period after the wheel event: nothing prepared yet.
		expect(slices).toBe(0);
		timers.advance(QUIET_MS);
		while (timers.step());
		expect(slices).toBeGreaterThan(1);
		expect(preparation.pending).toBe(false);
	});

	it('stops when another workbook loads, and edits before it finishes stay correct', () => {
		const timers = fakeScheduler();
		const preparation = new CalcPreparation(
			document.createElement('div'),
			() => false,
			timers.scheduler,
		);
		const first = opened();
		const session = createEditSession(first);
		preparation.start(session);
		timers.step();
		timers.step();
		session.setCellValue(0, 150, 0, 2);
		expect(getCell(first.sheets[0]!, 150, 1)?.value).toBe(20);
		expect(getCell(first.sheets[0]!, 149, 1)?.value).toBe(-1);
		preparation.start(undefined);
		expect(preparation.pending).toBe(false);
		expect(timers.queued).toBe(0);
	});
});
