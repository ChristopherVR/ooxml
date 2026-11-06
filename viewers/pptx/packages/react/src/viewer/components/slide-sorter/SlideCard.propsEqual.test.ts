import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { SlideCardProps } from './SlideCard';
import { slideCardPropsEqual } from './SlideCard';

const noop = (): void => {};

function makeProps(overrides: Partial<SlideCardProps> = {}): SlideCardProps {
	return {
		slide: {
			id: 'ppt/slides/slide1.xml',
			rId: 'rId2',
			slideNumber: 1,
			elements: [],
		} as unknown as PptxSlide,
		index: 0,
		isActive: false,
		isDragTarget: false,
		isSelected: false,
		selectedCount: 0,
		selectionOrder: 0,
		canvasSize: { width: 960, height: 540 },
		canEdit: true,
		onSlideClick: noop,
		onDoubleClick: noop,
		onContextMenu: noop,
		onDragStart: noop,
		onDragOver: noop,
		onDragLeave: noop,
		onDrop: noop,
		...overrides,
	};
}

describe('slideCardPropsEqual', () => {
	it('skips a parent re-render that passes the same slide object', () => {
		const props = makeProps();
		expect(slideCardPropsEqual(props, { ...props })).toBeTruthy();
	});

	it('repaints the sorter card when only slide metadata changes', () => {
		// The card's thumbnail shows the transition badge and the background, which
		// live on the slide object, not on its elements.
		const base = makeProps();
		const edits: Partial<PptxSlide>[] = [
			{ transition: { type: 'fade' } } as Partial<PptxSlide>,
			{ backgroundColor: '#336699' },
		];
		for (const edit of edits) {
			const next = { ...base, slide: { ...base.slide, ...edit } };
			expect(next.slide.elements).toBe(base.slide.elements);
			expect(slideCardPropsEqual(base, next)).toBeFalsy();
		}
	});
});
