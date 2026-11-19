// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';

import { createDeferredRows } from './deferred-rows';
import type { FrameScheduler } from './virtualized-list';
import {
	computeVirtualRange,
	createVirtualWindow,
	DEFERRED_FILL_FRAME_BUDGET_MS,
	DEFERRED_FILL_SETTLE_MS,
	SLIDE_VIRTUALIZATION_THRESHOLD,
} from './virtualized-list';

describe('computeVirtualRange', () => {
	it('returns an empty window for an empty list', () => {
		expect(computeVirtualRange(0, 100, 0, 600)).toStrictEqual({
			startIndex: 0,
			endIndex: -1,
			totalHeight: 0,
			offsetY: 0,
			visibleRange: { start: 0, end: -1 },
		});
	});

	it('adds overscan without exceeding list bounds', () => {
		const range = computeVirtualRange(100, 100, 5000, 600, 3);
		expect(range).toMatchObject({
			startIndex: 47,
			endIndex: 59,
			totalHeight: 10_000,
			offsetY: 4700,
			visibleRange: { start: 50, end: 56 },
		});
	});

	it('clamps invalid dimensions and scroll positions', () => {
		const range = computeVirtualRange(3, 0, -100, -20, -1);
		expect(range.startIndex).toBe(0);
		expect(range.endIndex).toBe(0);
		expect(range.totalHeight).toBe(3);
	});

	it('uses the same large-deck threshold as React', () => {
		expect(SLIDE_VIRTUALIZATION_THRESHOLD).toBe(50);
	});
});

describe('createVirtualWindow', () => {
	const manualFrames = () => {
		const queue: (() => void)[] = [];
		const frames: FrameScheduler = {
			request: (callback) => queue.push(callback),
			cancel: (handle) => {
				queue[(handle as number) - 1] = () => {};
			},
		};
		return {
			frames,
			flush() {
				for (const callback of queue.splice(0)) {
					callback();
				}
			},
		};
	};
	const scroller = (height = 600) => {
		const el = document.createElement('div');
		Object.defineProperty(el, 'clientHeight', { configurable: true, value: height });
		document.body.appendChild(el);
		return el;
	};
	const scrollTo = (el: HTMLElement, top: number) => {
		el.scrollTop = top;
		el.dispatchEvent(new Event('scroll'));
	};

	it('reads a burst of scroll events once, in the next frame', () => {
		const { frames, flush } = manualFrames();
		const changes: number[] = [];
		const el = scroller();
		const view = createVirtualWindow(
			{ totalItems: 200, itemHeight: 100 },
			(range) => changes.push(range.startIndex),
			frames,
		);
		view.attach(el);
		scrollTo(el, 2000);
		scrollTo(el, 2500);
		scrollTo(el, 3000);
		expect(changes).toStrictEqual([]);
		flush();
		expect(changes).toStrictEqual([25]);
		expect(view.range.startIndex).toBe(25);
		view.destroy();
		el.remove();
	});

	it('does not report a scroll that keeps the same rows', () => {
		const { frames, flush } = manualFrames();
		const onChange = vi.fn();
		const el = scroller();
		const view = createVirtualWindow({ totalItems: 200, itemHeight: 100 }, onChange, frames);
		view.attach(el);
		scrollTo(el, 40);
		flush();
		expect(onChange).not.toHaveBeenCalled();
		view.destroy();
		el.remove();
	});

	it('reveals a row by the least scroll and reports its window at once', () => {
		const { frames } = manualFrames();
		const onChange = vi.fn();
		const el = scroller();
		const view = createVirtualWindow({ totalItems: 200, itemHeight: 100 }, onChange, frames);
		view.attach(el);
		view.reveal(10);
		expect(el.scrollTop).toBe(500);
		expect(onChange).toHaveBeenCalledTimes(1);
		expect(view.range.visibleRange).toStrictEqual({ start: 5, end: 11 });
		view.reveal(8);
		expect(el.scrollTop).toBe(500);
		expect(onChange).toHaveBeenCalledTimes(1);
		view.destroy();
		el.remove();
	});

	it('stops reading the container after destroy', () => {
		const { frames, flush } = manualFrames();
		const onChange = vi.fn();
		const el = scroller();
		const view = createVirtualWindow({ totalItems: 200, itemHeight: 100 }, onChange, frames);
		view.attach(el);
		scrollTo(el, 5000);
		view.destroy();
		flush();
		scrollTo(el, 9000);
		flush();
		expect(onChange).not.toHaveBeenCalled();
		el.remove();
	});
});

describe('deferred row fills', () => {
	const clockFrames = () => {
		let time = 0;
		let frames: (() => void)[] = [];
		let timers: { at: number; run: () => void }[] = [];
		const scheduler: FrameScheduler = {
			request: (callback) => frames.push(callback),
			cancel: () => {},
			delay: (callback, ms) => timers.push({ at: time + ms, run: callback }),
			cancelDelay: () => {},
			now: () => time,
		};
		return {
			scheduler,
			advance(ms: number) {
				time += ms;
				const due = timers.filter((timer) => timer.at <= time);
				timers = timers.filter((timer) => timer.at > time);
				due.forEach((timer) => timer.run());
			},
			frame() {
				const run = frames;
				frames = [];
				run.forEach((callback) => callback());
			},
			tick(ms = 4) {
				time += ms;
			},
		};
	};
	const scroller = () => {
		const el = document.createElement('div');
		document.body.appendChild(el);
		return el;
	};

	it('fills rows only once scrolling has paused', () => {
		const clock = clockFrames();
		const el = scroller();
		const view = createVirtualWindow(
			{ totalItems: 200, itemHeight: 100 },
			() => {},
			clock.scheduler,
		);
		view.attach(el);
		const filled: string[] = [];
		el.dispatchEvent(new Event('scroll'));
		view.defer('a', () => filled.push('a'));
		clock.frame();
		expect(filled).toStrictEqual([]);
		clock.advance(DEFERRED_FILL_SETTLE_MS);
		clock.frame();
		expect(filled).toStrictEqual(['a']);
		view.destroy();
		el.remove();
	});

	it('stops a frame of fills at its budget and continues in the next frame', () => {
		const clock = clockFrames();
		const view = createVirtualWindow(
			{ totalItems: 200, itemHeight: 100 },
			() => {},
			clock.scheduler,
		);
		const filled: string[] = [];
		for (const key of ['a', 'b', 'c']) {
			view.defer(key, () => {
				filled.push(key);
				clock.tick(DEFERRED_FILL_FRAME_BUDGET_MS);
			});
		}
		clock.frame();
		expect(filled).toStrictEqual(['a']);
		clock.frame();
		expect(filled).toStrictEqual(['a', 'b']);
		view.destroy();
	});

	it('never fills a row that left the window first', () => {
		const clock = clockFrames();
		const view = createVirtualWindow(
			{ totalItems: 200, itemHeight: 100 },
			() => {},
			clock.scheduler,
		);
		const ready: string[][] = [];
		const rows = createDeferredRows(view, (keys) => ready.push([...keys]));
		rows.sync(['a', 'b']);
		rows.sync(['b', 'c']);
		// One declarative row per frame: its render lands after the fill.
		clock.frame();
		expect(ready.at(-1)).toStrictEqual(['b']);
		clock.frame();
		expect(ready.at(-1)).toStrictEqual(['b', 'c']);
		rows.sync(['c']);
		expect(ready.at(-1)).toStrictEqual(['c']);
		rows.destroy();
		view.destroy();
	});
});

describe('virtual window scroll container', () => {
	it('turns scroll anchoring off while attached, so a reveal is not undone', () => {
		const el = document.createElement('div');
		const view = createVirtualWindow({ totalItems: 200, itemHeight: 100 }, () => {});
		view.attach(el);
		expect(el.style.overflowAnchor).toBe('none');
		view.attach(null);
		expect(el.style.overflowAnchor).toBe('');
		view.destroy();
	});
});
