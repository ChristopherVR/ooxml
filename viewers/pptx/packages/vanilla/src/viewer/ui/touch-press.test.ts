/**
 * Touch long-press and double-tap on the section headers and the Slide Sorter
 * (issue #290). iOS Safari fires no `contextmenu` on a long-press and no
 * reliable `dblclick`; these controls opt into the shared `installTouchPress`
 * with `data-pptx-touch-press` and must reach the handlers a mouse does.
 */
import { installTouchPress, TOUCH_LONG_PRESS_MS } from 'ooxml-ui/pptx';
import type { PptxSection, PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { openSlideSorterOverlay } from './slide-sorter-overlay';
import { renderThumbnailSections } from './thumbnail-sections';

let host: HTMLElement;
let uninstall: () => void;

beforeEach(() => {
	vi.useFakeTimers();
	host = document.createElement('div');
	document.body.appendChild(host);
	uninstall = installTouchPress(host);
});

afterEach(() => {
	uninstall();
	document.body.replaceChildren();
	vi.useRealTimers();
});

const slides = [1, 2].map((n) => ({
	id: `s${n}`,
	rId: `rId${n}`,
	slideNumber: n,
	sectionId: 'sec1',
	elements: [],
})) as PptxSlide[];

function touch(target: Element, type: string): void {
	target.dispatchEvent(
		new PointerEvent(type, {
			bubbles: true,
			cancelable: true,
			pointerId: 6,
			pointerType: 'touch',
			isPrimary: true,
			clientX: 8,
			clientY: 9,
		}),
	);
}

function longPress(target: Element): void {
	touch(target, 'pointerdown');
	vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS);
	touch(target, 'pointerup');
}

function doubleTap(target: Element): void {
	for (let i = 0; i < 2; i++) {
		touch(target, 'pointerdown');
		touch(target, 'pointerup');
	}
}

describe('touch long-press and double-tap', () => {
	it('renames a section on a double-tap and opens its menu on a long-press', () => {
		const sections = [{ id: 'sec1', name: 'Intro', slideIds: ['s1', 's2'] }] as PptxSection[];
		host.append(
			...renderThumbnailSections({
				doc: document,
				t: createTranslator(),
				sections,
				slides,
				actions: {
					toggle: vi.fn(),
					rename: vi.fn(),
					delete: vi.fn(),
					move: vi.fn(),
					addAfter: vi.fn(),
				} as never,
				buildSlide: () => document.createElement('button'),
			}),
		);
		const header = host.querySelector('[data-pptx-chrome="section-header"]')!;
		const menus = vi.fn();
		header.addEventListener('contextmenu', menus);
		longPress(header);
		expect(menus).toHaveBeenCalledTimes(1);
		doubleTap(header.querySelector('.pptxv-thumb-section-toggle')!);
		expect(header.querySelector('input')).not.toBeNull();
	});

	it('opens a sorter slide on a double-tap and its menu on a long-press', () => {
		const onSelect = vi.fn();
		openSlideSorterOverlay(document, host, createTranslator(), {
			slides,
			onSelect,
			onReorder: vi.fn(),
			onDelete: vi.fn(),
			onDuplicate: vi.fn(),
			onToggleHidden: vi.fn(),
			canEdit: true,
		} as never);
		const card = host.querySelectorAll('[data-pptx-chrome="sorter-tile"]')[1]!;
		const menus = vi.fn();
		card.addEventListener('contextmenu', menus);
		longPress(card);
		expect(menus).toHaveBeenCalledTimes(1);
		doubleTap(card.querySelector(':scope > button')!);
		expect(onSelect).toHaveBeenCalledWith(1);
	});
});
