// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { printHtmlDocument, printNeedsWindow } from './print-frame';
import { printWhenLoaded } from './print-window';

const IPAD = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15';
const DESKTOP = 'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 Chrome/140.0 Safari/537.36';

/** A print window stand-in recording listeners and calls. */
function fakeWindow(readyState: DocumentReadyState) {
	const listeners = new Map<string, () => void>();
	const win = {
		document: { readyState, open: vi.fn(), write: vi.fn(), close: vi.fn() },
		addEventListener: vi.fn((type: string, listener: () => void) => listeners.set(type, listener)),
		removeEventListener: vi.fn(),
		focus: vi.fn(),
		print: vi.fn(),
		close: vi.fn(),
	};
	return { win, fire: (type: string) => listeners.get(type)?.() };
}

afterEach(() => {
	vi.useRealTimers();
	vi.restoreAllMocks();
	document.body.innerHTML = '';
});

describe('printWhenLoaded', () => {
	it('prints on load, once, and closes the window after printing', () => {
		vi.useFakeTimers();
		const { win, fire } = fakeWindow('loading');
		printWhenLoaded(win as unknown as Window);
		expect(win.print).not.toHaveBeenCalled();
		fire('load');
		vi.runAllTimers();
		expect(win.print).toHaveBeenCalledOnce();
		fire('afterprint');
		expect(win.close).toHaveBeenCalledOnce();
	});
	it('prints a document that already finished loading, and falls back when load never fires', () => {
		vi.useFakeTimers();
		const complete = fakeWindow('complete');
		printWhenLoaded(complete.win as unknown as Window);
		vi.runAllTimers();
		expect(complete.win.print).toHaveBeenCalledOnce();
		const stuck = fakeWindow('interactive');
		printWhenLoaded(stuck.win as unknown as Window);
		vi.runAllTimers();
		expect(stuck.win.print).toHaveBeenCalledOnce();
	});
});

describe('printHtmlDocument', () => {
	it('opens a window on iOS and iPadOS, where an iframe prints the parent page', () => {
		expect(printNeedsWindow({ userAgent: IPAD, maxTouchPoints: 5 })).toBe(true);
		expect(printNeedsWindow({ userAgent: DESKTOP })).toBe(false);
		const { win } = fakeWindow('complete');
		const open = vi.spyOn(window, 'open').mockReturnValue(win as unknown as Window);
		expect(printHtmlDocument('<p>notes</p>', { env: { userAgent: IPAD, maxTouchPoints: 5 } })).toBe(
			true,
		);
		expect(open).toHaveBeenCalledOnce();
		expect(win.document.write).toHaveBeenCalledWith('<p>notes</p>');
		expect(document.querySelector('iframe')).toBeNull();
	});
	it('reports a blocked popup on iOS', () => {
		vi.spyOn(window, 'open').mockReturnValue(null);
		expect(printHtmlDocument('<p>x</p>', { env: { userAgent: IPAD, maxTouchPoints: 5 } })).toBe(
			false,
		);
	});
	it('prints a hidden, script-free iframe elsewhere', () => {
		const open = vi.spyOn(window, 'open');
		expect(printHtmlDocument('<p>notes</p>', { env: { userAgent: DESKTOP } })).toBe(true);
		const frame = document.querySelector('iframe')!;
		expect(frame.getAttribute('sandbox')).toBe('allow-same-origin allow-modals');
		expect(frame.srcdoc).toBe('<p>notes</p>');
		expect(open).not.toHaveBeenCalled();
	});
});
