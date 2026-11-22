import type { PptxSlide, ZoomPptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildSummaryZoomView,
	resolveSummaryZoomNavigation,
	zoomPreviewSource,
	zoomTargetInfo,
} from './summary-zoom';

const summary: ZoomPptxElement = {
	id: 'summary',
	type: 'zoom',
	zoomType: 'summary',
	targetSlideIndex: 0,
	x: 100,
	y: 50,
	width: 400,
	height: 200,
	summaryLayout: 'fixed',
	summaryTargets: [
		{
			sectionId: 'intro',
			targetSlideIndex: 0,
			x: 100,
			y: 50,
			width: 180,
			height: 200,
			title: 'Welcome',
		},
		{
			sectionId: 'details',
			targetSlideIndex: 3,
			x: 320,
			y: 50,
			width: 180,
			height: 200,
		},
	],
};

describe('summary zoom view model', () => {
	it('preserves tile order, fixed geometry, labels, and targets', () => {
		const view = buildSummaryZoomView(summary, (index) => ({
			slideNumber: index + 10,
			sectionName: index === 3 ? 'Deep dive' : undefined,
			backgroundColor: '#123456',
		}));
		expect(view).toMatchObject({ layout: 'fixed', ariaLabel: 'Summary Zoom with 2 sections' });
		expect(view?.tiles).toMatchObject([
			{ label: 'Welcome', targetSlideIndex: 0, slideLabel: 'Slide 10', style: { left: '0%' } },
			{
				label: 'Deep dive',
				targetSlideIndex: 3,
				slideLabel: 'Slide 13',
				style: { left: '55%' },
			},
		]);
		expect(resolveSummaryZoomNavigation(view, 1)).toBe(3);
	});

	it('renders current slide artwork in a preview and keeps custom covers', () => {
		const slide: PptxSlide = {
			id: 'target',
			rId: 'rId2',
			slideNumber: 2,
			elements: [
				{ type: 'text', id: 'text', text: 'Original', x: 20, y: 20, width: 140, height: 60 },
			],
		};
		const size = { width: 960, height: 540 };
		const thumbnail = { zoomImageType: 'preview' as const, imageData: 'embedded' };
		const before = zoomPreviewSource(thumbnail, zoomTargetInfo(slide, size));
		const updated: PptxSlide = {
			...slide,
			elements: [
				{ type: 'text', id: 'text', text: 'Revised', x: 20, y: 20, width: 140, height: 60 },
			],
		};
		expect(decodeURIComponent(before ?? '')).toContain('Original');
		expect(
			decodeURIComponent(zoomPreviewSource(thumbnail, zoomTargetInfo(updated, size)) ?? ''),
		).toContain('Revised');
		expect(
			zoomPreviewSource({ ...thumbnail, zoomImageType: 'cover' }, zoomTargetInfo(updated, size)),
		).toBe('embedded');
	});

	it('builds a grid and rejects invalid navigation indexes', () => {
		const view = buildSummaryZoomView({ ...summary, summaryLayout: 'grid' });
		expect(view?.containerStyle.gridTemplateColumns).toBe('repeat(2, minmax(0, 1fr))');
		expect(resolveSummaryZoomNavigation(view, -1)).toBeUndefined();
		expect(resolveSummaryZoomNavigation(view, 3)).toBeUndefined();
	});
});
