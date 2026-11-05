// @vitest-environment jsdom
import type { PptxElement, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTranslator } from './i18n';
import { createDefaultRegistry, renderSlideStage } from './render';

/**
 * PowerPoint's dash lengths (4:3 for `dash`, 8:3 for `lgDash`, ...) in units of
 * the pen width, with a round cap folded into the dash. This binding paints
 * connectors through the shared DOM renderer: pin the pattern it ends up with.
 */
function dashOf(shapeStyle: Record<string, unknown>): Element | null {
	const line = {
		type: 'connector',
		id: 'c',
		x: 10,
		y: 10,
		width: 400,
		height: 0,
		shapeType: 'straightConnector1',
		shapeStyle: { strokeColor: '#000000', strokeWidth: 6, ...shapeStyle },
	} as unknown as PptxElement;
	const stage = renderSlideStage({
		document,
		slide: { id: 's1', rId: 'r1', slideNumber: 1, elements: [line] } as PptxSlide,
		canvasSize: { width: 960, height: 540 },
		mediaDataUrls: new Map(),
		registry: createDefaultRegistry(),
		t: createTranslator('en'),
		scale: 1,
		interactive: true,
	});
	return stage.querySelector('[data-element-id="c"] [stroke-dasharray]');
}

describe('vanilla: connector dash patterns follow PowerPoint', () => {
	it('uses the measured dash and gap lengths', () => {
		expect(dashOf({ strokeDash: 'dash' })?.getAttribute('stroke-dasharray')).toBe('24 18');
		expect(dashOf({ strokeDash: 'sysDot' })?.getAttribute('stroke-dasharray')).toBe('6 6');
	});

	it('folds a round cap into the dash and keeps square dashes flat', () => {
		const round = dashOf({ strokeDash: 'dash', lineCap: 'rnd' });
		expect(round?.getAttribute('stroke-dasharray')).toBe('18 24');
		expect(round?.getAttribute('stroke-linecap')).toBe('round');
		expect(dashOf({ strokeDash: 'dash', lineCap: 'sq' })?.getAttribute('stroke-linecap')).toBe(
			'butt',
		);
	});
});
