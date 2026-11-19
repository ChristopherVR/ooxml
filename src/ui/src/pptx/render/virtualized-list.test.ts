// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';

import type { FrameScheduler } from './virtualized-list';
import {
	computeVirtualRange,
	createVirtualWindow,
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
