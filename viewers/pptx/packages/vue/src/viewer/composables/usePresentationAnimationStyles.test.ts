import { describe, expect, it, vi } from 'vitest';
import { effectScope, nextTick, ref, shallowRef } from 'vue';

import type { UseAnimationPlaybackResult } from './useAnimationPlayback';
import { usePresentationAnimationStyles } from './usePresentationAnimationStyles';

describe('presentation animation trigger hit testing', () => {
	it('applies fixed whole opacity without CSS and restores authored alpha', async () => {
		const root = document.createElement('div');
		root.innerHTML = '<div data-element-id="s" style="opacity:0.6"></div>';
		const states = shallowRef(
			new Map([['s', { visible: true, cssAnimation: undefined, opacity: 0.4 }]]),
		);
		const scope = effectScope();
		try {
			scope.run(() =>
				usePresentationAnimationStyles({
					frameRef: ref(root),
					activeSlide: () => 'slide',
					playback: {
						presentationElementStates: states,
						interactiveTriggerShapeIds: shallowRef(new Set()),
						hoverTriggerShapeIds: shallowRef(new Set()),
					} as unknown as UseAnimationPlaybackResult,
				}),
			);
			await nextTick();
			await nextTick();
			const shape = root.firstElementChild as HTMLElement;
			expect(shape.style.opacity).toBe('0.4');
			expect(shape.style.animation).toBe('');
			states.value = new Map();
			await nextTick();
			await nextTick();
			expect(shape.style.opacity).toBe('0.6');
		} finally {
			scope.stop();
		}
	});
	it('marks trigger shapes and removes stale markers', async () => {
		const root = document.createElement('div');
		root.innerHTML = '<div data-element-id="trigger"></div>';
		const ids = shallowRef<ReadonlySet<string>>(new Set(['trigger']));
		const scope = effectScope();
		scope.run(() =>
			usePresentationAnimationStyles({
				frameRef: ref(root),
				activeSlide: () => 'slide',
				playback: {
					presentationElementStates: shallowRef(new Map()),
					interactiveTriggerShapeIds: ids,
					hoverTriggerShapeIds: shallowRef(new Set()),
					handleInteractiveShapeClick: vi.fn(),
				} as unknown as UseAnimationPlaybackResult,
			}),
		);
		await nextTick();
		await nextTick();
		expect(root.firstElementChild?.hasAttribute('data-pptx-anim-trigger')).toBeTruthy();
		ids.value = new Set();
		await nextTick();
		await nextTick();
		expect(root.firstElementChild?.hasAttribute('data-pptx-anim-trigger')).toBeFalsy();
		scope.stop();
	});
});
