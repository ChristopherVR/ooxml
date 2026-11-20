// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { labelInName } from '../../test-support/label-in-name';
import { createRibbonControlIcon, RIBBON_CONTROL_ICONS } from './ribbon-icons';

describe('canonical Home ribbon icons', () => {
	it.each(Object.keys(RIBBON_CONTROL_ICONS))(
		'renders visible artwork with the common 24-unit geometry: %s',
		(control) => {
			const svg = createRibbonControlIcon(document, control);
			expect(svg.getAttribute('viewBox')).toBe('0 0 24 24');
			expect(svg.childElementCount).toBeGreaterThan(0);
			for (const rectangle of svg.querySelectorAll('rect')) {
				expect(Number(rectangle.getAttribute('width'))).toBeGreaterThan(0);
				expect(Number(rectangle.getAttribute('height'))).toBeGreaterThan(0);
			}
		},
	);

	it('keeps strokes and text-shadow layers from the reference artwork', () => {
		expect(
			createRibbonControlIcon(document, 'home.font.characterSpacing').getAttribute('stroke-width'),
		).toBe('1.5');
		const layers = createRibbonControlIcon(document, 'home.font.shadow').querySelectorAll('path');
		expect(layers).toHaveLength(2);
		// A filled bold S, then its outline offset as the shadow.
		expect(layers[0]!.getAttribute('fill')).toBe('currentColor');
		expect(layers[1]!.getAttribute('d')).toBe(layers[0]!.getAttribute('d'));
		expect(layers[1]!.getAttribute('transform')).toBe('translate(1.5 1.5)');
		expect(layers[1]!.getAttribute('opacity')).toBe('0.4');
	});

	it('draws no text that a button name would have to contain (WCAG 2.5.3)', () => {
		const shown = Object.keys(RIBBON_CONTROL_ICONS)
			.map((id) => [id, createRibbonControlIcon(document, id).textContent ?? ''] as const)
			.filter(([, text]) => !labelInName(text, ''));
		expect(shown).toEqual([]);
	});
});
