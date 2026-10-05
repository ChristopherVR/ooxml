// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { PRINT_FRAME_TIMEOUT_MS, printHtml } from './print';

const frames = () => document.querySelectorAll('iframe');
const print = vi.fn();
/** Prints and stubs the new frame's `print()` and `focus()` (jsdom implements neither). */
const start = (html: string) => {
	printHtml(document, html);
	const frame = [...frames()].at(-1)!;
	const target = frame.contentWindow!;
	Object.defineProperty(target, 'print', { value: print, configurable: true });
	Object.defineProperty(target, 'focus', { value: () => undefined, configurable: true });
	return frame;
};

beforeEach(() => {
	vi.useFakeTimers();
	print.mockClear();
});
afterEach(() => {
	vi.useRealTimers();
	vi.restoreAllMocks();
	document.body.replaceChildren();
});

describe('printHtml', () => {
	it('keeps the frame after print() returns and removes it on afterprint', () => {
		const frame = start('<p>Sheet</p>');
		vi.advanceTimersByTime(5000);
		expect(print).toHaveBeenCalledTimes(1);
		expect(frame.isConnected).toBe(true);
		frame.contentWindow!.dispatchEvent(new Event('afterprint'));
		expect(frame.isConnected).toBe(false);
	});

	it('removes a frame that never gets afterprint after the safety timeout', () => {
		const frame = start('<p>Sheet</p>');
		vi.advanceTimersByTime(PRINT_FRAME_TIMEOUT_MS - 1);
		expect(frame.isConnected).toBe(true);
		vi.advanceTimersByTime(1);
		expect(frame.isConnected).toBe(false);
	});

	it('removes the previous frame when the next print starts', () => {
		const first = start('<p>One</p>');
		vi.advanceTimersByTime(100);
		start('<p>Two</p>');
		expect(first.isConnected).toBe(false);
		expect(frames()).toHaveLength(1);
	});
});
