/**
 * Touch long-press and double-tap in the Slide Sorter and the slide rail
 * (issue #290). iOS Safari fires no `contextmenu` on a long-press and no
 * reliable `dblclick`; the tiles opt into the shared `installTouchPress` with
 * `data-pptx-touch-press` and must reach the handlers a mouse does.
 */
import { mount } from '@vue/test-utils';
import { installTouchPress, TOUCH_LONG_PRESS_MS } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import SlideSorter from './SlideSorter.vue';

let wrapper: ReturnType<typeof mount> | undefined;
let uninstall: (() => void) | undefined;

beforeEach(() => {
	vi.useFakeTimers();
});

afterEach(() => {
	uninstall?.();
	wrapper?.unmount();
	wrapper = undefined;
	document.body.replaceChildren();
	vi.useRealTimers();
});

function touch(target: Element, type: string): void {
	target.dispatchEvent(
		new PointerEvent(type, {
			bubbles: true,
			cancelable: true,
			pointerId: 4,
			pointerType: 'touch',
			isPrimary: true,
			clientX: 20,
			clientY: 30,
		}),
	);
}

function mountSorter() {
	const slides = Array.from({ length: 3 }, (_, i) => ({ id: `s${i}`, elements: [] }));
	wrapper = mount(SlideSorter, {
		attachTo: document.body,
		props: {
			slides: slides as unknown as PptxSlide[],
			canvasSize: { width: 960, height: 540 },
			mediaDataUrls: new Map<string, string>(),
			activeIndex: 0,
			canEdit: true,
		},
	});
	uninstall = installTouchPress(wrapper.element as HTMLElement);
	return wrapper;
}

describe('slide sorter touch gestures', () => {
	it('opens a slide on a double-tap', () => {
		const sorter = mountSorter();
		const tile = sorter.findAll('[data-pptx-chrome="sorter-tile"]')[1]!.element;
		for (let i = 0; i < 2; i++) {
			touch(tile, 'pointerdown');
			touch(tile, 'pointerup');
		}
		expect(sorter.emitted('select')).toStrictEqual([[1]]);
	});

	it('opens the slide menu on a long-press', async () => {
		const sorter = mountSorter();
		const tile = sorter.findAll('[data-pptx-chrome="sorter-tile"]')[2]!.element;
		touch(tile, 'pointerdown');
		vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
		touch(tile, 'pointerup');
		await sorter.vm.$nextTick();
		expect(document.querySelector('pptx-ui-context-menu')).not.toBeNull();
		expect(sorter.emitted('select')).toBeUndefined();
	});
});
