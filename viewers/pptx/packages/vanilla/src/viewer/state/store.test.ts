import { describe, expect, it, vi } from 'vitest';

import { createStore } from './store';
import { clampSlideIndex, createInitialViewerState } from './viewer-state';

describe('createStore', () => {
	it('returns the initial state', () => {
		const store = createStore({ a: 1, b: 'x' });
		expect(store.get()).toStrictEqual({ a: 1, b: 'x' });
	});

	it('merges patches and notifies subscribers with next + previous state', () => {
		const store = createStore({ a: 1, b: 'x' });
		const listener = vi.fn();
		store.subscribe(listener);

		store.set({ a: 2 });

		expect(store.get()).toStrictEqual({ a: 2, b: 'x' });
		expect(listener).toHaveBeenCalledExactlyOnceWith({ a: 2, b: 'x' }, { a: 1, b: 'x' });
	});

	it('does not notify when the patch changes nothing', () => {
		const store = createStore({ a: 1 });
		const listener = vi.fn();
		store.subscribe(listener);

		store.set({ a: 1 });
		store.set({});

		expect(listener).not.toHaveBeenCalled();
	});

	it('stops notifying after unsubscribe', () => {
		const store = createStore({ a: 1 });
		const listener = vi.fn();
		const unsubscribe = store.subscribe(listener);

		unsubscribe();
		store.set({ a: 2 });

		expect(listener).not.toHaveBeenCalled();
	});

	it('delivers a set made by a listener once, in order, to every listener', () => {
		const store = createStore({ slide: 0, zoom: 2 });
		const first: string[] = [];
		const later: string[] = [];
		const change = (state: { slide: number; zoom: number }, previous: typeof state) =>
			`slide ${previous.slide}->${state.slide}, zoom ${previous.zoom}->${state.zoom}`;
		store.subscribe((state, previous) => {
			first.push(change(state, previous));
			if (state.slide !== previous.slide) {
				store.set({ zoom: 1 });
			}
		});
		store.subscribe((state, previous) => later.push(change(state, previous)));

		store.set({ slide: 1 });

		// The listener after the one that set sees both changes as one; the one
		// that set sees its own change after it returns. Nobody sees one twice.
		expect(later).toStrictEqual(['slide 0->1, zoom 2->1']);
		expect(first).toStrictEqual(['slide 0->1, zoom 2->2', 'slide 1->1, zoom 2->1']);
		expect(store.get()).toStrictEqual({ slide: 1, zoom: 1 });
	});

	it('does not call a listener that an earlier listener unsubscribed during the same change', () => {
		const store = createStore({ n: 0 });
		const seen: string[] = [];
		let offB = (): void => {};
		store.subscribe((state) => {
			seen.push(`A:${state.n}`);
			offB();
		});
		offB = store.subscribe((state) => seen.push(`B:${state.n}`));

		store.set({ n: 1 });

		// The old store still called B, from the copy of the list it took before A ran.
		expect(seen).toStrictEqual(['A:1']);
	});

	it('lets a listener that writes read only its own write until it returns', () => {
		const store = createStore({ slide: 0, rendered: -1 });
		const seen: number[] = [];
		store.subscribe((state, previous) => {
			if (state.slide === 1 && previous.slide === 0) {
				store.set({ slide: 2 });
				seen.push(store.get().rendered);
			}
		});
		// A later listener that derives `rendered` from `slide`.
		store.subscribe((state) => {
			if (state.rendered !== state.slide) {
				store.set({ rendered: state.slide });
			}
		});

		store.set({ slide: 1 });

		// The old store ran the deriving listener inside the write, so the writer read
		// 2. That listener now runs after the writer returns.
		expect(seen).toStrictEqual([-1]);
		expect(store.get()).toStrictEqual({ slide: 2, rendered: 2 });
	});

	it('gives a listener the state it last saw as previous after another listener threw', () => {
		const store = createStore({ n: 0 });
		const seen: string[] = [];
		store.subscribe((state) => {
			if (state.n === 1) {
				throw new Error('boom');
			}
		});
		store.subscribe((state, previous) => seen.push(`${previous.n}->${state.n}`));

		expect(() => store.set({ n: 1 })).toThrow('boom');
		store.set({ n: 2 });

		// The throw kept the second listener from hearing about 1, so it hears 0 -> 2.
		// The old store passed 1 as previous, a state this listener never saw.
		expect(seen).toStrictEqual(['0->2']);
	});

	it('throws, rather than spinning, when a listener never stops changing the state', () => {
		const store = createStore({ n: 0 });
		store.subscribe((state) => store.set({ n: state.n + 1 }));
		expect(() => store.set({ n: 1 })).toThrow(/never settles/u);
	});

	it('supports multiple subscribers', () => {
		const store = createStore({ a: 1 });
		const first = vi.fn();
		const second = vi.fn();
		store.subscribe(first);
		store.subscribe(second);

		store.set({ a: 3 });

		expect(first).toHaveBeenCalledOnce();
		expect(second).toHaveBeenCalledOnce();
	});
});

describe('viewer state helpers', () => {
	it('creates a sane initial state', () => {
		const state = createInitialViewerState();
		expect(state.slides).toStrictEqual([]);
		expect(state.currentSlide).toBe(0);
		expect(state.zoom).toBe('fit');
		expect(state.loading).toBeFalsy();
		expect(state.error).toBeNull();
		expect(state.presenting).toBeFalsy();
		expect(state.canvasSize.width).toBeGreaterThan(0);
	});

	it('clamps slide indices into range', () => {
		expect(clampSlideIndex(-5, 10)).toBe(0);
		expect(clampSlideIndex(3, 10)).toBe(3);
		expect(clampSlideIndex(99, 10)).toBe(9);
		expect(clampSlideIndex(2.7, 10)).toBe(2);
		expect(clampSlideIndex(4, 0)).toBe(0);
	});
});
