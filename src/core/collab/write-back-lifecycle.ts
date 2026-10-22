// The hook point that flushes a write-back when a page is going away. DOM-free: the targets are
// structural, so a UI layer registers the real `window` and `document` (or passes none and the
// globals are used), and tests pass fakes.
import type { WriteBackScheduler } from './write-back';

/** The part of `window` used: `pagehide` is listened on it. */
export interface PageHideTarget {
	addEventListener(type: 'pagehide', listener: () => void): void;
	removeEventListener(type: 'pagehide', listener: () => void): void;
}

/** The part of `document` used: `visibilitychange` is listened on it. */
export interface VisibilityTarget {
	readonly visibilityState?: string;
	addEventListener(type: 'visibilitychange', listener: () => void): void;
	removeEventListener(type: 'visibilitychange', listener: () => void): void;
}

export interface FlushOnPageCloseOptions {
	/** Default `globalThis` when it can listen (a browser window), otherwise no `pagehide` hook. */
	window?: PageHideTarget | null;
	/** Default `globalThis.document` when present. */
	document?: VisibilityTarget | null;
	/** Receives a flush that failed (it never throws). Default ignores it. */
	onError?: (error: unknown) => void;
}

/**
 * Flushes `scheduler` on `pagehide` and when the page turns hidden, and returns the function that
 * removes the listeners. Only unsaved work is written.
 *
 * Browsers do not wait for asynchronous work when a page unloads: the flush starts the write, but
 * a write that needs a network round trip can still be lost on close. A host that must survive a
 * close should send the bytes with `navigator.sendBeacon(url, bytes)` or
 * `fetch(url, { method: 'PUT', body: bytes, keepalive: true })` (the browser caps keepalive bodies
 * at about 64 KiB, so a large document may still not make it), which the browser completes after
 * the page is gone.
 */
export function flushWriteBackOnPageClose(
	scheduler: Pick<WriteBackScheduler, 'flush' | 'dirty'>,
	options: FlushOnPageCloseOptions = {},
): () => void {
	const globals = globalThis as unknown as {
		addEventListener?: unknown;
		document?: VisibilityTarget;
	};
	const win =
		options.window === undefined
			? typeof globals.addEventListener === 'function'
				? (globalThis as unknown as PageHideTarget)
				: null
			: options.window;
	const doc = options.document === undefined ? (globals.document ?? null) : options.document;
	const flush = (): void => {
		if (!scheduler.dirty) return;
		scheduler.flush().catch((error: unknown) => options.onError?.(error));
	};
	const hidden = (): void => {
		if (doc?.visibilityState === 'hidden') flush();
	};
	win?.addEventListener('pagehide', flush);
	doc?.addEventListener('visibilitychange', hidden);
	return () => {
		win?.removeEventListener('pagehide', flush);
		doc?.removeEventListener('visibilitychange', hidden);
	};
}
