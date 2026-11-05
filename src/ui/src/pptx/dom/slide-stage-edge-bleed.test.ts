// @vitest-environment happy-dom
import type { PptxElement, PptxSlide } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../i18n/translator';
import { createDefaultRegistry } from './elements';
import { renderSlideStage } from './slide-stage';

/**
 * Issue #141 on the vanilla stage: two abutting filled rectangles must not
 * show a light seam where their shared edge falls inside a device pixel. The
 * lower box grows under the opaque upper one (`resolveStageEdgeBleeds`); its
 * gradient keeps its authored box, the selection keeps its exact geometry.
 */

const fading = {
	fillMode: 'gradient',
	fillGradientType: 'linear',
	fillGradientAngle: 0,
	fillGradientStops: [
		{ color: '#00405E', position: 0 },
		{ color: '#33667E', position: 95, opacity: 0 },
	],
	fillGradient: 'linear-gradient(90deg, #00405E 0%, rgba(51, 102, 126, 0) 95%)',
	strokeWidth: 0,
	strokeFillMode: 'none',
};

function rect(id: string, x: number, width: number, shapeStyle: object): PptxElement {
	return {
		type: 'shape',
		id,
		x,
		y: 0,
		width,
		height: 720,
		shapeType: 'rect',
		text: '',
		shapeStyle,
	} as PptxElement;
}

/** Slide 13 of the issue deck: "Object 8" (lower, fading) then "Object 2" (upper, opaque). */
const slide = {
	id: 'slide-13',
	elements: [
		rect('object-8', 581, 350, fading),
		rect('object-2', 0, 581, {
			fillMode: 'solid',
			fillColor: '#00405E',
			strokeWidth: 0,
			strokeFillMode: 'none',
		}),
	],
} as unknown as PptxSlide;

function box(id: string, selectedElementIds?: ReadonlySet<string>): HTMLElement {
	const stage = renderSlideStage({
		document,
		slide,
		canvasSize: { width: 1280, height: 720 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator(),
		interactive: true,
		...(selectedElementIds ? { selectedElementIds } : {}),
	});
	return stage.querySelector<HTMLElement>(`[data-element-id="${id}"]`)!;
}

describe('vanilla stage abutting-edge bleed', () => {
	it('grows the lower panel under the upper one and keeps its gradient at the authored box', () => {
		const lower = box('object-8');
		expect(lower.style.left).toBe('579px');
		expect(lower.style.width).toBe('352px');
		expect(lower.style.backgroundSize).toBe('350px 720px, 3px 720px');
		expect(lower.style.backgroundPosition).toBe('2px 0px, 0px 0px');
		const upper = box('object-2');
		expect(upper.style.left).toBe('0px');
		expect(upper.style.width).toBe('581px');
	});

	it('keeps the exact box while either panel is selected', () => {
		expect(box('object-8', new Set(['object-2'])).style.width).toBe('350px');
	});
});
