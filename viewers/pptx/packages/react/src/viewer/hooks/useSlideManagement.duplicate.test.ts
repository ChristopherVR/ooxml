// @vitest-environment happy-dom
import { createChartElement } from 'pptx-viewer-core';
import type { ChartPptxElement, PptxSlide } from 'pptx-viewer-core';
import React, { act, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

import { useSlideManagement } from './useSlideManagement';

describe('slide duplication', () => {
	it('retains package provenance and isolates chart data before saving', async () => {
		const source: PptxSlide = {
			id: 'ppt/slides/slide1.xml',
			rId: 'rId2',
			slideNumber: 1,
			elements: [
				createChartElement('line', { categories: ['A'], series: [{ name: 'One', values: [1] }] }),
			],
		};
		const container = document.createElement('div');
		document.body.append(container);
		const root = createRoot(container);
		let current: PptxSlide[] = [];
		let management: ReturnType<typeof useSlideManagement>;
		function Harness() {
			const [slides, setSlides] = useState([source]);
			current = slides;
			management = useSlideManagement({
				slides,
				activeSlide: slides[0],
				activeSlideIndex: 0,
				setActiveSlideIndex: vi.fn(),
				ops: { updateSlides: setSlides } as never,
				history: { markDirty: vi.fn() } as never,
			});
			return null;
		}
		try {
			await act(async () => root.render(React.createElement(Harness)));
			await act(async () => management.handleDuplicateSlides([0]));
			expect(current[1].sourceSlideId).toBe(source.id);
			(current[1].elements[0] as ChartPptxElement).chartData!.series[0].values[0] = 2;
			expect((source.elements[0] as ChartPptxElement).chartData!.series[0].values[0]).toBe(1);
			expect(current[1].elements[0].id).not.toBe(source.elements[0].id);
		} finally {
			await act(async () => root.unmount());
			container.remove();
		}
	});
});
