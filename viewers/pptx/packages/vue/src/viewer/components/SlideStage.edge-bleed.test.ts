/**
 * Issue #141: two abutting filled rectangles must not show a light seam where
 * their shared edge falls inside a device pixel. The lower box grows under the
 * opaque upper one (shared `resolveStageEdgeBleeds`); a selected box keeps its
 * exact geometry, and a Morph layer never grows.
 */
import { mount } from '@vue/test-utils';
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import SlideStage from './SlideStage.vue';

function rect(id: string, x: number, fillColor: string): PptxElement {
	return {
		type: 'shape',
		id,
		x,
		y: 0,
		width: 100,
		height: 100,
		shapeType: 'rect',
		text: '',
		shapeStyle: { fillMode: 'solid', fillColor, strokeWidth: 0, strokeFillMode: 'none' },
	} as PptxElement;
}

const slide = {
	id: 's1',
	elements: [rect('lower', 100, '#112233'), rect('upper', 200, '#445566')],
} as unknown as PptxSlide;

function boxStyle(extra: Record<string, unknown>, id: string): string {
	const wrapper = mount(SlideStage, {
		props: { slide, canvasSize: { width: 960, height: 540 }, mediaDataUrls: new Map(), ...extra },
	});
	return wrapper.get(`[data-element-id="${id}"]`).attributes('style') ?? '';
}

describe('slideStage abutting-edge bleed', () => {
	it('grows the lower box under the abutting opaque box above it', () => {
		const lower = boxStyle({ interactive: true }, 'lower');
		expect(lower).toContain('left: 100px');
		expect(lower).toContain('width: 102px');
		const upper = boxStyle({ interactive: true }, 'upper');
		expect(upper).toContain('left: 200px');
		expect(upper).toContain('width: 100px');
	});

	it('keeps the exact box while either box is selected, and on a Morph layer', () => {
		expect(boxStyle({ interactive: true, selectedElementIds: ['upper'] }, 'lower')).toContain(
			'width: 100px',
		);
		expect(boxStyle({ preserveElementIds: true }, 'lower')).toContain('width: 100px');
	});
});
