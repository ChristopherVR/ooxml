// @vitest-environment happy-dom
/**
 * Applying a layout has to refresh the artwork the slide inherits from it.
 *
 * The viewer keeps layout / master shapes outside `slide.elements`, so replacing
 * the slide alone leaves the canvas painting the previous layout's decoration
 * until the deck is reopened.
 */
import type { PptxElement, PptxHandler, PptxSlide } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';

import type { EditorHistoryResult } from './useEditorHistory';
import type { ElementOperations } from './useElementOperations';
import { useLayoutSwitching } from './useLayoutSwitching';

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
});

function makeSlide(id: string): PptxSlide {
	return { id, rId: '', slideNumber: 1, elements: [] } as unknown as PptxSlide;
}

const TEMPLATE_ELEMENT = { id: 'layout-9', type: 'shape' } as unknown as PptxElement;

describe('useLayoutSwitching.applyLayout', () => {
	it('reports the slide, and its refreshed template artwork, after a switch', async () => {
		const slide = makeSlide('ppt/slides/slide1.xml');
		const switched = { ...slide, layoutPath: 'ppt/slideLayouts/slideLayout3.xml' };
		const handler = {
			applyLayoutToSlide: vi.fn().mockResolvedValue(switched),
			getTemplateElementsForSlide: vi.fn().mockResolvedValue([TEMPLATE_ELEMENT]),
		} as unknown as PptxHandler;
		const updateSlides = vi.fn();
		const markDirty = vi.fn();
		const onTemplateElementsChanged = vi.fn();

		let applyLayout: ((path: string) => Promise<void>) | undefined;
		function Harness(): null {
			applyLayout = useLayoutSwitching({
				handler,
				slides: [slide],
				activeSlideIndex: 0,
				ops: { updateSlides } as unknown as ElementOperations,
				history: { markDirty } as unknown as EditorHistoryResult,
				onTemplateElementsChanged,
			}).applyLayout;
			return null;
		}
		act(() => root.render(<Harness />));

		await act(async () => {
			await applyLayout!('ppt/slideLayouts/slideLayout3.xml');
		});

		expect(handler.applyLayoutToSlide).toHaveBeenCalledWith(
			0,
			'ppt/slideLayouts/slideLayout3.xml',
			[slide],
		);
		expect(updateSlides).toHaveBeenCalledOnce();
		expect(handler.getTemplateElementsForSlide).toHaveBeenCalledWith(switched.id);
		expect(onTemplateElementsChanged).toHaveBeenCalledWith(switched.id, [TEMPLATE_ELEMENT]);
		expect(markDirty).toHaveBeenCalledOnce();
	});

	it('leaves the slides state it was given as it is (the core replaces the slide in its array)', async () => {
		const slide = makeSlide('ppt/slides/slide1.xml');
		const given = [slide];
		const handler = {
			applyLayoutToSlide: vi.fn(async (index: number, layoutPath: string, slides: PptxSlide[]) => {
				const updated = { ...slides[index]!, layoutPath };
				slides[index] = updated;
				return updated;
			}),
			getTemplateElementsForSlide: vi.fn().mockResolvedValue([]),
		} as unknown as PptxHandler;
		const updateSlides = vi.fn();

		let applyLayout: ((path: string) => Promise<void>) | undefined;
		function Harness(): null {
			applyLayout = useLayoutSwitching({
				handler,
				slides: given,
				activeSlideIndex: 0,
				ops: { updateSlides } as unknown as ElementOperations,
				history: { markDirty: vi.fn() } as unknown as EditorHistoryResult,
			}).applyLayout;
			return null;
		}
		act(() => root.render(<Harness />));

		await act(async () => {
			await applyLayout!('ppt/slideLayouts/slideLayout3.xml');
		});

		expect(given).toStrictEqual([slide]);
		expect(given[0]).toBe(slide);
		const updater = updateSlides.mock.calls[0]?.[0] as (prev: PptxSlide[]) => PptxSlide[];
		expect(updater(given)[0]?.layoutPath).toBe('ppt/slideLayouts/slideLayout3.xml');
	});

	it('does nothing without a loaded handler', async () => {
		const markDirty = vi.fn();
		let applyLayout: ((path: string) => Promise<void>) | undefined;
		function Harness(): null {
			applyLayout = useLayoutSwitching({
				handler: null,
				slides: [makeSlide('ppt/slides/slide1.xml')],
				activeSlideIndex: 0,
				ops: { updateSlides: vi.fn() } as unknown as ElementOperations,
				history: { markDirty } as unknown as EditorHistoryResult,
			}).applyLayout;
			return null;
		}
		act(() => root.render(<Harness />));

		await act(async () => {
			await applyLayout!('ppt/slideLayouts/slideLayout3.xml');
		});

		expect(markDirty).not.toHaveBeenCalled();
	});
});
