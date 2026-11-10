import type { PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import type { SlideThumbnailProps } from './SlideThumbnail';
import { slideThumbnailPropsEqual } from './SlideThumbnail';

function makeSlide(overrides: Partial<PptxSlide> = {}): PptxSlide {
	return {
		id: 'ppt/slides/slide1.xml',
		slideNumber: 1,
		elements: [],
		showMasterShapes: true,
		...overrides,
	} as unknown as PptxSlide;
}

function makeProps(overrides: Partial<SlideThumbnailProps> = {}): SlideThumbnailProps {
	return {
		slide: makeSlide(),
		templateElements: [],
		canvasSize: { width: 960, height: 540 },
		...overrides,
	};
}

describe('slideThumbnailPropsEqual', () => {
	it('treats identical props as equal', () => {
		const props = makeProps();
		expect(slideThumbnailPropsEqual(props, props)).toBeTruthy();
	});

	it('is unequal when showMasterShapes flips, even with the same elements/templateElements references', () => {
		// buildPreviewElements -> visibleTemplateElements drops inherited
		// layout/master artwork entirely when showMasterShapes is false, so the
		// memo must not skip a re-render just because `elements` and
		// `templateElements` themselves are unchanged.
		const templateElements = [{ id: 'layout-1' }] as SlideThumbnailProps['templateElements'];
		const prev = makeProps({
			slide: makeSlide({ showMasterShapes: true }),
			templateElements,
		});
		const next = makeProps({
			slide: makeSlide({ showMasterShapes: false }),
			templateElements,
		});
		expect(slideThumbnailPropsEqual(prev, next)).toBeFalsy();
	});

	it('is unequal when the slide id, dirty flag or hidden flag changes', () => {
		const base = makeProps();
		expect(
			slideThumbnailPropsEqual(base, makeProps({ slide: makeSlide({ id: 'other' }) })),
		).toBeFalsy();
		expect(
			slideThumbnailPropsEqual(base, makeProps({ slide: makeSlide({ isDirty: true }) })),
		).toBeFalsy();
		expect(
			slideThumbnailPropsEqual(base, makeProps({ slide: makeSlide({ hidden: true }) })),
		).toBeFalsy();
	});

	it('is unequal when only slide metadata changes, with the same element references', () => {
		// The thumbnail draws the transition badge, the slide number and title
		// fields and the background from the slide object itself, so an edit that
		// keeps `elements` and `templateElements` must still repaint it.
		const base = makeProps();
		const edits: Partial<PptxSlide>[] = [
			{ transition: { type: 'fade' } } as Partial<PptxSlide>,
			{ slideNumber: 2 },
			{ backgroundPattern: { preset: 'pct5' } } as Partial<PptxSlide>,
		];
		for (const edit of edits) {
			const next = { ...base, slide: { ...base.slide, ...edit } };
			expect(next.slide.elements).toBe(base.slide.elements);
			expect(slideThumbnailPropsEqual(base, next)).toBeFalsy();
		}
	});

	it('skips a parent re-render that passes the same slide object', () => {
		const base = makeProps();
		expect(
			slideThumbnailPropsEqual(base, { ...base, canvasSize: { width: 960, height: 540 } }),
		).toBeTruthy();
	});

	it('is unequal when canvasSize changes', () => {
		const base = makeProps();
		expect(
			slideThumbnailPropsEqual(base, makeProps({ canvasSize: { width: 1280, height: 720 } })),
		).toBeFalsy();
	});

	it('repaints when resizing the rail changes the preview width', () => {
		const props = makeProps({ previewWidth: 132 });
		expect(slideThumbnailPropsEqual(props, { ...props, previewWidth: 180 })).toBeFalsy();
	});
});
