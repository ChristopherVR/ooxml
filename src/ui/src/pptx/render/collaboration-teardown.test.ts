import { describe, expect, it, vi } from 'vitest';

import type {
	TeardownEventLike,
	TeardownListener,
	TeardownWindowLike,
} from './collaboration-teardown';
import {
	COLLAB_LEAVE_MESSAGE,
	clearLocalAwareness,
	registerCollaborationTeardown,
} from './collaboration-teardown';

interface FakeWindow extends TeardownWindowLike {
	emit: (type: string, event?: TeardownEventLike) => void;
	count: (type: string) => number;
}

/** The page embedding the viewer frame, and its origin. */
const embedder = {};
const ORIGIN = 'https://viewer.example';

/** A leave request as the embedding page sends it. */
const fromEmbedder = (data: unknown): TeardownEventLike => ({
	data,
	origin: ORIGIN,
	source: embedder,
});

function makeWindow(): FakeWindow {
	const listeners = new Map<string, Set<TeardownListener>>();
	return {
		parent: embedder,
		location: { origin: ORIGIN },
		addEventListener: (type, listener) => {
			const set = listeners.get(type) ?? new Set<TeardownListener>();
			set.add(listener);
			listeners.set(type, set);
		},
		removeEventListener: (type, listener) => {
			listeners.get(type)?.delete(listener);
		},
		emit: (type, event = {}) => {
			for (const listener of listeners.get(type) ?? []) {
				listener(event);
			}
		},
		count: (type) => listeners.get(type)?.size ?? 0,
	};
}

describe('registerCollaborationTeardown', () => {
	it('evaluates dynamic ownership on each attempted navigation', () => {
		const target = makeWindow();
		const leave = vi.fn();
		let borrowed = true;
		registerCollaborationTeardown({ leave, target, leaveOnBeforeUnload: () => !borrowed });
		target.emit('beforeunload');
		expect(leave).not.toHaveBeenCalled();
		borrowed = false;
		target.emit('beforeunload');
		expect(leave).toHaveBeenCalledOnce();
	});

	it('keeps a borrowed session attached when navigation is canceled', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target, leaveOnBeforeUnload: false });
		target.emit('beforeunload');
		expect(leave).not.toHaveBeenCalled();
		target.emit('pagehide');
		expect(leave).toHaveBeenCalledOnce();
	});

	it('leaves the room on pagehide (the iframe-removal / tab-close path)', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('pagehide', { persisted: false });

		expect(leave).toHaveBeenCalledOnce();
	});

	it('leaves only once across beforeunload + pagehide of the same unload', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('beforeunload');
		target.emit('pagehide', { persisted: false });

		expect(leave).toHaveBeenCalledOnce();
	});

	it('keeps the session on a bfcache pagehide when no rejoin is available', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('pagehide', { persisted: true });

		expect(leave).not.toHaveBeenCalled();
	});

	it('leaves on a bfcache pagehide and rejoins on restore when rejoin is given', () => {
		const target = makeWindow();
		const leave = vi.fn();
		const rejoin = vi.fn();
		registerCollaborationTeardown({ leave, rejoin, target });

		target.emit('pagehide', { persisted: true });
		expect(leave).toHaveBeenCalledOnce();

		target.emit('pageshow', { persisted: true });
		expect(rejoin).toHaveBeenCalledOnce();

		// A second departure after the restore leaves again.
		target.emit('pagehide', { persisted: false });
		expect(leave).toHaveBeenCalledTimes(2);
	});

	it('ignores a pageshow that is not a bfcache restore', () => {
		const target = makeWindow();
		const rejoin = vi.fn();
		registerCollaborationTeardown({ leave: vi.fn(), rejoin, target });

		target.emit('pageshow', { persisted: false });

		expect(rejoin).not.toHaveBeenCalled();
	});

	it('leaves when an embedding page posts the leave message', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('message', fromEmbedder({ type: COLLAB_LEAVE_MESSAGE }));

		expect(leave).toHaveBeenCalledOnce();
	});

	it('accepts a bare string leave message and ignores unrelated messages', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('message', fromEmbedder({ type: 'something-else' }));
		target.emit('message', fromEmbedder('hello'));
		target.emit('message', fromEmbedder(null));
		expect(leave).not.toHaveBeenCalled();

		target.emit('message', fromEmbedder(COLLAB_LEAVE_MESSAGE));
		expect(leave).toHaveBeenCalledOnce();
	});

	it('ignores a leave message from another origin or another window', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target });

		target.emit('message', {
			data: COLLAB_LEAVE_MESSAGE,
			origin: 'https://evil.example',
			source: embedder,
		});
		target.emit('message', { data: COLLAB_LEAVE_MESSAGE, origin: ORIGIN, source: {} });
		target.emit('message', { data: COLLAB_LEAVE_MESSAGE });
		expect(leave).not.toHaveBeenCalled();
	});

	it('obeys an embedder on another origin the host names', () => {
		const target = makeWindow();
		const leave = vi.fn();
		registerCollaborationTeardown({ leave, target, leaveMessageOrigin: 'https://host.example' });

		target.emit('message', fromEmbedder(COLLAB_LEAVE_MESSAGE));
		expect(leave).not.toHaveBeenCalled();
		target.emit('message', {
			data: COLLAB_LEAVE_MESSAGE,
			origin: 'https://host.example',
			source: embedder,
		});
		expect(leave).toHaveBeenCalledOnce();
	});

	it('removes every listener when disposed', () => {
		const target = makeWindow();
		const leave = vi.fn();
		const dispose = registerCollaborationTeardown({ leave, target });

		expect(target.count('pagehide')).toBe(1);
		dispose();

		expect(target.count('pagehide')).toBe(0);
		expect(target.count('beforeunload')).toBe(0);
		expect(target.count('pageshow')).toBe(0);
		expect(target.count('message')).toBe(0);
		target.emit('pagehide', { persisted: false });
		expect(leave).not.toHaveBeenCalled();
	});

	it('is a no-op disposer outside a browser', () => {
		const leave = vi.fn();
		expect(() => registerCollaborationTeardown({ leave })()).not.toThrow();
		expect(leave).not.toHaveBeenCalled();
	});
});

describe('clearLocalAwareness', () => {
	it('withdraws the local presence', () => {
		const setLocalState = vi.fn();
		clearLocalAwareness({ setLocalState });
		expect(setLocalState).toHaveBeenCalledWith(null);
	});

	it('tolerates a missing awareness or method', () => {
		expect(() => clearLocalAwareness(null)).not.toThrow();
		expect(() => clearLocalAwareness(undefined)).not.toThrow();
		expect(() => clearLocalAwareness({})).not.toThrow();
	});
});
