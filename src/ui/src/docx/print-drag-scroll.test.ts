// @vitest-environment jsdom
import { expect, it, vi } from 'vitest';
import { createPrintDragScroll } from './print-drag-scroll';
it('extends selection while scrolling both edges and cancels pending scrolling on disposal', () => {
	vi.useFakeTimers();
	const viewport = document.createElement('div');
	Object.defineProperties(viewport, { clientHeight: { value: 100 }, scrollHeight: { value: 500 } });
	vi.spyOn(viewport, 'getBoundingClientRect').mockReturnValue(new DOMRect(20, 100, 200, 100));
	const select = vi.fn(),
		scroll = createPrintDragScroll(viewport, select);
	scroll.update(50, 205);
	vi.advanceTimersByTime(48);
	expect(viewport.scrollTop).toBe(72);
	expect(select).toHaveBeenLastCalledWith(50, 199);
	scroll.update(50, 150);
	vi.advanceTimersByTime(64);
	expect(viewport.scrollTop).toBe(72);
	scroll.update(50, 95);
	vi.advanceTimersByTime(16);
	expect(viewport.scrollTop).toBe(48);
	scroll.destroy();
	const calls = select.mock.calls.length;
	vi.advanceTimersByTime(100);
	expect(viewport.scrollTop).toBe(48);
	expect(select).toHaveBeenCalledTimes(calls);
	scroll.update(50, 205);
	vi.advanceTimersByTime(100);
	expect(viewport.scrollTop).toBe(48);
	vi.restoreAllMocks();
	vi.useRealTimers();
});
