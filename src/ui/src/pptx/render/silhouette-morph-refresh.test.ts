// @vitest-environment jsdom
import type { PptxElement } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import { buildMorphEffectCarry } from './morph-effect-carry';
import { getSilhouetteEffects } from './silhouette-effects';
import { refreshSilhouetteMorphFilters } from './silhouette-morph-refresh';

function star(id: string, width: number, distance: number): PptxElement {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width,
		height: width * 0.6,
		shapeType: 'star5',
		shapeStyle: {
			fillColor: '#336699',
			fillMode: 'solid',
			shadowColor: '#000000',
			shadowOpacity: 0.5,
			shadowBlur: 10,
			shadowAngle: 45,
			shadowDistance: distance,
			outerShadowXml: {},
		},
	} as unknown as PptxElement;
}

describe('refreshSilhouetteMorphFilters', () => {
	afterEach(() => {
		document.body.innerHTML = '';
	});

	it('adds the dissolve layers to a filter painted before the plan existed, wherever it sits', () => {
		const from = star('refresh-a', 200, 8);
		const to = star('refresh-b', 400, 20);
		// The stage was painted first: static markup, the filter beside its container.
		const staged = getSilhouetteEffects(to)?.filterMarkup ?? '';
		expect(staged).not.toContain('--pptx-morph-p');
		document.body.innerHTML = `<svg data-pptx-morph-fx="refresh-b">${staged}</svg>`;
		buildMorphEffectCarry(from, to, 'to');
		refreshSilhouetteMorphFilters(document.body);
		const filter = document.querySelector('filter');
		expect(filter?.id).toBe('shape-fx-refresh-b');
		expect(filter?.innerHTML).toContain('--pptx-morph-p');
		// A second pass leaves it alone.
		const before = document.body.innerHTML;
		refreshSilhouetteMorphFilters(document.body);
		expect(document.body.innerHTML).toBe(before);
	});
});
