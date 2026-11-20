// @vitest-environment happy-dom
/**
 * Touch long-press and double-tap on the slide rail, the section headers and
 * the Slide Sorter (issue #290). iOS Safari fires no `contextmenu` on a
 * long-press and no reliable `dblclick`, so these controls opt into the shared
 * `installTouchPress` with `data-pptx-touch-press` and must reach the same
 * handlers a mouse does.
 */
import { createInstance } from 'i18next';
import { installTouchPress, TOUCH_LONG_PRESS_MS } from 'ooxml-ui/pptx';
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import type { PptxSlide } from 'pptx-viewer-core';
import React, { act, createRef } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { I18nextProvider, initReactI18next } from 'react-i18next';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { SlideCard } from './slide-sorter/SlideCard';
import { SectionHeader } from './slides-pane/SectionHeader';
import { SlideItem } from './slides-pane/SlideItem';

const i18n = createInstance();
let container: HTMLDivElement;
let root: Root;
let uninstall: () => void;

beforeEach(async () => {
	vi.useFakeTimers();
	await i18n.use(initReactI18next).init({
		lng: 'en',
		resources: { en: { translation: translationsEn } },
		interpolation: { escapeValue: false },
	});
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	uninstall = installTouchPress(container);
});

afterEach(() => {
	uninstall();
	act(() => root.unmount());
	container.remove();
	vi.useRealTimers();
});

const slide = { id: 's1', rId: 'rId1', slideNumber: 1, elements: [] } as unknown as PptxSlide;

function touch(target: Element, type: string, x = 12, y = 14): void {
	target.dispatchEvent(
		new PointerEvent(type, {
			bubbles: true,
			cancelable: true,
			pointerId: 3,
			pointerType: 'touch',
			isPrimary: true,
			clientX: x,
			clientY: y,
		}),
	);
}

function longPress(target: Element): void {
	touch(target, 'pointerdown');
	act(() => vi.advanceTimersByTime(TOUCH_LONG_PRESS_MS));
	touch(target, 'pointerup');
}

function doubleTap(target: Element): void {
	for (let i = 0; i < 2; i++) {
		touch(target, 'pointerdown');
		touch(target, 'pointerup');
	}
}

function render(node: React.ReactElement): void {
	act(() => root.render(<I18nextProvider i18n={i18n}>{node}</I18nextProvider>));
}

describe('touch long-press and double-tap', () => {
	it('opens the slide rail context menu on a long-press', () => {
		const onOpenSlideCtxMenu = vi.fn();
		render(
			<SlideItem
				slide={slide}
				templateElements={[]}
				slideIndex={0}
				isActive={false}
				canvasSize={{ width: 960, height: 540 }}
				canEdit
				onSelectSlide={vi.fn()}
				onSlideContextMenu={vi.fn()}
				onOpenSlideCtxMenu={onOpenSlideCtxMenu}
				onDragStart={vi.fn()}
				onDragOver={vi.fn()}
				onDrop={vi.fn()}
				slideRef={vi.fn()}
			/>,
		);
		longPress(container.querySelector('[data-pptx-chrome="slide-row"]')!);
		// The touch point, not the element's corner, positions the menu.
		expect(onOpenSlideCtxMenu).toHaveBeenCalledWith(12, 14, 0);
	});

	it('opens the section menu on a long-press and renames on a double-tap', () => {
		const onContextMenu = vi.fn();
		const onStartRename = vi.fn();
		render(
			<SectionHeader
				sectionId='sec'
				label='Intro'
				slideCount={1}
				isCollapsed={false}
				isRenaming={false}
				renameValue=''
				canEdit
				sectionIndex={0}
				totalSections={1}
				renameInputRef={createRef<HTMLInputElement>()}
				onToggle={vi.fn()}
				onContextMenu={onContextMenu}
				onStartRename={onStartRename}
				onRenameValueChange={vi.fn()}
				onCommitRename={vi.fn()}
				onCancelRename={vi.fn()}
			/>,
		);
		const header = container.querySelector('[data-pptx-chrome="section-header"]')!;
		longPress(header);
		expect(onContextMenu).toHaveBeenCalledTimes(1);
		doubleTap(header);
		expect(onStartRename).toHaveBeenCalledWith('sec', 'Intro');
	});

	it('opens a sorter slide on a double-tap and its menu on a long-press', () => {
		const onDoubleClick = vi.fn();
		const onContextMenu = vi.fn();
		render(
			<SlideCard
				slide={slide}
				index={0}
				isActive={false}
				isDragTarget={false}
				isSelected={false}
				selectedCount={0}
				selectionOrder={0}
				canvasSize={{ width: 960, height: 540 }}
				canEdit
				onSlideClick={vi.fn()}
				onDoubleClick={onDoubleClick}
				onContextMenu={onContextMenu}
				onDragStart={vi.fn()}
				onDragOver={vi.fn()}
				onDragLeave={vi.fn()}
				onDrop={vi.fn()}
			/>,
		);
		const card = container.querySelector('[data-pptx-chrome="sorter-tile"]')!;
		doubleTap(card);
		expect(onDoubleClick).toHaveBeenCalledWith(0);
		longPress(card);
		expect(onContextMenu).toHaveBeenCalledTimes(1);
	});
});
