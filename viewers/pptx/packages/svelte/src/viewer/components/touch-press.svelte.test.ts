/**
 * Touch long-press and double-tap on the slide rail, its section headers and
 * the Slide Sorter (issue #290). iOS Safari fires no `contextmenu` on a
 * long-press and no reliable `dblclick`; these controls opt into the shared
 * `installTouchPress` with `data-pptx-touch-press` and must reach the same
 * handlers a mouse does.
 */
import { installTouchPress, TOUCH_LONG_PRESS_MS } from 'ooxml-ui/pptx';
import type { PptxSection, PptxSlide } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import SlideSorterOverlay from './SlideSorterOverlay.svelte';
import ThumbnailRail from './ThumbnailRail.svelte';

let cleanup: (() => void) | undefined;
beforeEach(() => {
	vi.useFakeTimers();
});
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
	document.body.replaceChildren();
	vi.useRealTimers();
});

const CANVAS = { width: 960, height: 540 };
const slides = [1, 2, 3].map(
	(n) =>
		({
			id: `s${n}`,
			rId: `rId${n}`,
			slideNumber: n,
			sectionId: 'sec1',
			elements: [],
		}) as PptxSlide,
);
const sections = [{ id: 'sec1', name: 'Intro', slideIds: ['s1', 's2', 's3'] }] as PptxSection[];

function render(component: never, props: Record<string, unknown>): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(component, { target, props });
	const uninstall = installTouchPress(target);
	cleanup = () => {
		uninstall();
		void unmount(instance);
		target.remove();
	};
	flushSync();
	return target;
}

function touch(target: Element, type: string): void {
	target.dispatchEvent(
		new PointerEvent(type, {
			bubbles: true,
			cancelable: true,
			pointerId: 5,
			pointerType: 'touch',
			isPrimary: true,
			clientX: 10,
			clientY: 10,
		}),
	);
}

function longPress(target: Element): void {
	touch(target, 'pointerdown');
	vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
	touch(target, 'pointerup');
	flushSync();
}

function doubleTap(target: Element): void {
	for (let i = 0; i < 2; i++) {
		touch(target, 'pointerdown');
		touch(target, 'pointerup');
	}
	flushSync();
}

const menuItems = (): Element[] =>
	Array.from(document.querySelectorAll('[role="menu"] [role="menuitem"]'));

describe('touch long-press and double-tap', () => {
	it('opens the slide rail and section menus on a long-press', () => {
		const target = render(ThumbnailRail as never, {
			slides,
			canvasSize: CANVAS,
			mediaDataUrls: new Map<string, string>(),
			current: 0,
			sections,
			editable: true,
			onselect: () => undefined,
			onaddslide: () => undefined,
		});
		longPress(target.querySelector('[data-pptx-chrome="section-header"]')!);
		expect(menuItems().length).toBeGreaterThan(0);
	});

	it('renames a section on a double-tap of its header', () => {
		const target = render(ThumbnailRail as never, {
			slides,
			canvasSize: CANVAS,
			mediaDataUrls: new Map<string, string>(),
			current: 0,
			sections,
			editable: true,
			onselect: () => undefined,
			onaddslide: () => undefined,
		});
		doubleTap(target.querySelector('.pptx-svelte-section-toggle')!);
		expect(target.querySelector('[data-pptx-chrome="section-header"] input')).not.toBeNull();
	});

	it('opens a sorter slide on a double-tap and its menu on a long-press', () => {
		const onselect = vi.fn();
		const target = render(SlideSorterOverlay as never, {
			slides,
			canvasSize: CANVAS,
			mediaDataUrls: new Map<string, string>(),
			current: 0,
			canEdit: true,
			onselect,
			onmove: () => undefined,
			ondelete: () => undefined,
			onduplicate: () => undefined,
			onpaste: () => undefined,
			ontogglehidden: () => undefined,
			onclose: () => undefined,
		});
		const tile = target.querySelectorAll('[data-pptx-chrome="sorter-tile"]')[1]!;
		longPress(tile);
		expect(menuItems().length).toBeGreaterThan(0);
		doubleTap(tile.querySelector('button.preview')!);
		expect(onselect).toHaveBeenCalledWith(1);
	});
});
