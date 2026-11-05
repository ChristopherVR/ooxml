import type { PptxSlide, PptxSlideTransition } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { EditorController } from '../editor/editor-controller.svelte';
import type { EditorState } from '../editor/editor-state.svelte';
import { PresentationAnnotations } from '../presentation/presentation-annotations.svelte';
import type { TransitionState } from '../presentation/presentation-controller.svelte';
import SlideOverlays from './SlideOverlays.svelte';
import type { SlideOverlaysProps } from './viewer-body-props';

const slide = (id: string): PptxSlide =>
	({ id, elements: [], backgroundColor: '#ffffff' }) as unknown as PptxSlide;

const transitionOf = (
	transition: PptxSlideTransition,
	from: string,
	to: string,
): TransitionState => ({
	outgoing: slide(from),
	incoming: slide(to),
	transition,
});

// Issue #141, slides 6 -> 7: slide 6 fades in and advances after 0 s, so its
// timer fires as its fade ends, while that overlay is still mounted. The
// overlay arms its completion timer on mount, so a reused instance let the
// fade's leftover timer tear down slide 7's transition 50 ms in, and slide 7
// appeared in one frame.
describe('slide transition overlay mounting', () => {
	let cleanup: (() => void) | undefined;

	beforeEach(() => {
		vi.useFakeTimers();
	});

	afterEach(() => {
		cleanup?.();
		cleanup = undefined;
		vi.useRealTimers();
	});

	it("plays a timed slide's next transition in full", () => {
		const done = vi.fn();
		const props = $state({
			editor: {} as EditorState,
			controller: {} as EditorController,
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map<string, string>(),
			current: 1,
			activeSlide: slide('s2'),
			scale: 1,
			presenting: true,
			presentationTransition: transitionOf({ type: 'fade', durationMs: 200 }, 's1', 's2'),
			onTransitionDone: done,
			editingActive: false,
			annotations: new PresentationAnnotations(),
		} as SlideOverlaysProps);
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(SlideOverlays, { target, props });
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
		};

		// Slide 2's fade (200 ms) ends and its 0 s timing starts slide 3's
		// transition while the fade's overlay is still up.
		vi.advanceTimersByTime(200);
		props.current = 2;
		props.activeSlide = slide('s3');
		props.presentationTransition = transitionOf({ type: 'fade', durationMs: 400 }, 's2', 's3');
		flushSync();

		// The fade's own completion timer (250 ms) passes: slide 3's transition
		// must still be on screen.
		vi.advanceTimersByTime(140);
		expect(done).not.toHaveBeenCalled();
		expect(target.querySelector('[data-pptx-transition-overlay]')).not.toBeNull();

		// And it completes once, after its own 400 ms (+ 50 ms buffer).
		vi.advanceTimersByTime(320);
		expect(done).toHaveBeenCalledOnce();
	});
});
