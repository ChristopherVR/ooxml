// @vitest-environment happy-dom
/**
 * Design > Variants > Background Styles in the React ribbon: a hovered tile is shown over the
 * slides without an undo step, leaving restores the original, a pick commits through the real
 * `useDeckStyleHost` + `useRibbonGalleryCommands`, and the footer command reaches the Design tab.
 */
import { FIXED_TAB_GALLERIES } from 'ooxml-ui/pptx';
import type { RibbonGalleryPlacement } from 'ooxml-ui/pptx';
import type { PptxElement, PptxHandler, PptxSlide, PptxTheme } from 'pptx-viewer-core';
import React, { act, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock(import('react-i18next'), () => ({
	useTranslation: () => ({ t: (key: string) => key }),
}));

const { RibbonGallery } = await import('./RibbonGallery');
const { RibbonGalleryCommandsContext } = await import('../ribbon-gallery-context');
const { useRibbonGalleryCommands } = await import('../../hooks/useRibbonGalleryCommands');
const { useDeckStyleHost } = await import('../../hooks/useDeckStyleHost');

const BACKGROUNDS = FIXED_TAB_GALLERIES.find(
	(placement) => placement.gallery === 'themeBackgrounds',
) as RibbonGalleryPlacement;

const handler = {
	resolveStyleMatrixReferences: () => ({ shapeStyle: {} }),
	getThemeVariants: () => [],
	applyBackgroundStyleToSlides: (slides: PptxSlide[], templates: unknown, index: number) => ({
		slides: slides.map((slide) => ({ ...slide, backgroundColor: `#00000${index}` })),
		templateElementsBySlideId: templates,
	}),
	commitMasterBackgroundStyle: vi.fn(async () => true),
} as unknown as PptxHandler;

const slideOf = (color: string): PptxSlide =>
	({ id: 's1', rId: 'r1', slideNumber: 1, elements: [], backgroundColor: color }) as PptxSlide;

let container: HTMLDivElement;
let root: Root;
let state: {
	slides: () => PptxSlide[];
	held: boolean[];
	bumps: number;
	dirty: number;
	commands: string[];
};

beforeEach(() => {
	vi.useFakeTimers();
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
	state = { slides: () => [], held: [], bumps: 0, dirty: 0, commands: [] };
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	vi.useRealTimers();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

function Harness() {
	const [slides, setSlides] = useState<PptxSlide[]>(() => [slideOf('#FFFFFF')]);
	const [templates, setTemplates] = useState<Record<string, PptxElement[]>>({});
	const [theme, setTheme] = useState<PptxTheme | undefined>({ name: 'Office' });
	const handlerRef = useRef<PptxHandler | null>(handler);
	const slidesRef = useRef(slides);
	slidesRef.current = slides;
	state.slides = () => slidesRef.current;
	const deckStyle = useDeckStyleHost({
		slides,
		templateElementsBySlideId: templates,
		theme,
		themeColorMap: undefined,
		handlerRef,
		selectedSlideIds: ['s1'],
		setSlides,
		setTemplateElementsBySlideId: setTemplates,
		setTheme,
		holdTracking: (held) => state.held.push(held),
		bumpHistory: () => {
			state.bumps += 1;
		},
		markDirty: () => {
			state.dirty += 1;
		},
	});
	const commands = useRibbonGalleryCommands({
		selectedElement: null,
		theme,
		themeColorMap: undefined,
		handlerRef,
		editable: true,
		activeSlide: slides[0] ?? null,
		deckStyle,
		updateElementById: vi.fn(),
		updateThemeColorScheme: vi.fn<() => Promise<void>>(),
		updateThemeFontScheme: vi.fn<() => Promise<void>>(),
	});
	return (
		<RibbonGalleryCommandsContext.Provider value={commands}>
			<RibbonGallery placement={BACKGROUNDS} onCommand={(c) => state.commands.push(c)} />
		</RibbonGalleryCommandsContext.Provider>
	);
}

function open(): HTMLElement {
	act(() => root.render(<Harness />));
	container.querySelector<HTMLButtonElement>('[data-ribbon-gallery]')!.click();
	return container;
}

describe('React Background Styles gallery', () => {
	it('shows a hovered style over the slides, holds the undo tracker, and restores on leave', () => {
		open();
		const original = state.slides();
		const tile = container.querySelector('[data-gallery-item="3"]')!;
		act(() => {
			tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
			vi.advanceTimersByTime(150);
		});
		expect(state.slides()[0]!.backgroundColor).toBe('#000003');
		expect(state.held).toEqual([true]);
		act(() => {
			tile.dispatchEvent(new PointerEvent('pointerleave', { pointerType: 'mouse' }));
			vi.advanceTimersByTime(100);
		});
		expect(state.slides()).toBe(original);
		expect(state.held).toEqual([true, false]);
		expect(state.bumps).toBe(0);
		expect(state.dirty).toBe(0);
	});

	it('commits a pick as one undo step after settling the preview', async () => {
		open();
		const tile = container.querySelector<HTMLButtonElement>('[data-gallery-item="4"]')!;
		act(() => {
			tile.dispatchEvent(new PointerEvent('pointerenter', { pointerType: 'mouse' }));
			vi.advanceTimersByTime(150);
		});
		await act(async () => {
			tile.click();
			await vi.advanceTimersByTimeAsync(0);
		});
		expect(state.slides()[0]!.backgroundColor).toBe('#000004');
		expect(state.bumps).toBe(1);
		expect(state.dirty).toBe(1);
		expect(state.held).toEqual([true, false]);
	});

	it('runs the Format Background command from the footer', () => {
		open();
		act(() => {
			container
				.querySelector<HTMLButtonElement>('[data-gallery-footer="formatBackground"]')!
				.click();
		});
		expect(state.commands).toEqual(['formatBackground']);
	});
});
