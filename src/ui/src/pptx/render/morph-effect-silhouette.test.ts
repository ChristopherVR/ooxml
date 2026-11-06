import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildMorphEffectCarry } from './morph-effect-carry';
import { getSilhouetteEffects } from './silhouette-effects';

function shape(id: string, shapeType: string, width: number, style: Record<string, unknown> = {}) {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width,
		height: width * 0.6,
		shapeType,
		shapeStyle: {
			fillColor: '#336699',
			fillMode: 'solid',
			shadowColor: '#000000',
			shadowOpacity: 0.5,
			shadowBlur: 10,
			shadowAngle: 45,
			shadowDistance: 8,
			outerShadowXml: {},
			...style,
		},
	} as unknown as PptxElement;
}

describe('silhouette morph dissolve', () => {
	it('keeps one filter layer set per end, the outgoing one fading on 1 - p and the incoming on p', () => {
		const from = shape('morph-a', 'triangle', 200);
		const to = shape('morph-b', 'star5', 400, { shadowDistance: 20 });
		const carry = buildMorphEffectCarry(from, to, 'to');
		expect(carry?.prelude).toContain('@property --pptx-morph-p');
		expect(carry?.keyframes).toContain('--pptx-morph-p: 0;');
		expect(carry?.keyframes).toContain('--pptx-morph-p: 1;');
		const markup = getSilhouetteEffects(to)?.filterMarkup ?? '';
		expect(markup).toContain('calc(0.5 * (1 - var(--pptx-morph-p, 1)))');
		expect(markup).toContain('calc(0.5 * var(--pptx-morph-p, 1))');
		// The two ends sum, so a shadow they share never dips mid-morph.
		expect(markup).toContain('operator="arithmetic" k1="0" k2="1" k3="1" k4="0"');
		// The outgoing offset rides the box: 8 pt at 45 degrees, in the 2x frame.
		expect(markup).toContain('dx="11.31"');
	});

	it('leaves a shape alone that no morph brings in', () => {
		const lone = shape('morph-c', 'star5', 400);
		expect(getSilhouetteEffects(lone)?.filterMarkup).not.toContain('--pptx-morph-p');
	});

	it('does not dissolve on the outgoing half, and under a ghost fades only the incoming end in', () => {
		const from = shape('morph-d', 'triangle', 200);
		const to = shape('morph-e', 'star5', 400, { shadowDistance: 20 });
		expect(buildMorphEffectCarry(from, to, 'from')?.prelude).toBeUndefined();
		expect(buildMorphEffectCarry(from, to, 'to', 'underGhost')?.prelude).toContain('@property');
		const markup = getSilhouetteEffects(to)?.filterMarkup ?? '';
		expect(markup).toContain('calc(0.5 * var(--pptx-morph-p, 1))');
		expect(markup).not.toContain('(1 - var(--pptx-morph-p, 1))');
	});
});
