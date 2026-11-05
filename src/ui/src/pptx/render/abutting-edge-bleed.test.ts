import type { PptxElement, PptxSlide, ShapeStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	resolveAbuttingEdgeBleeds,
	resolveStageEdgeBleeds,
	slideAnimatedElementIds,
} from './abutting-edge-bleed';

const solid = (fillColor: string, extra: Partial<ShapeStyle> = {}): ShapeStyle => ({
	fillMode: 'solid',
	fillColor,
	strokeColor: 'transparent',
	strokeWidth: 0,
	strokeFillMode: 'none',
	...extra,
});

/** Issue #141 slide 13's "Object 8": accent4 fading to transparent, left to right. */
const fadingGradient: ShapeStyle = {
	fillMode: 'gradient',
	fillGradientType: 'linear',
	fillGradientAngle: 0,
	fillGradientStops: [
		{ color: '#00405E', position: 0 },
		{ color: '#33667E', position: 95, opacity: 0 },
	],
	strokeColor: 'transparent',
	strokeWidth: 0,
	strokeFillMode: 'none',
};

function rect(
	id: string,
	box: [number, number, number, number],
	shapeStyle: ShapeStyle,
	extra: Record<string, unknown> = {},
): PptxElement {
	const [x, y, width, height] = box;
	return {
		id,
		type: 'shape',
		x,
		y,
		width,
		height,
		shapeType: 'rect',
		text: '',
		shapeStyle,
		...extra,
	} as PptxElement;
}

describe('resolveAbuttingEdgeBleeds', () => {
	it('grows issue #141 slide 13 lower gradient panel under the opaque panel to its left', () => {
		// Painted in this order: Object 8 (lower, x 581..931), then Object 2 (upper, x 0..581).
		const object8 = rect('o8', [581, 0, 350, 720], fadingGradient);
		const object2 = rect('o2', [0, 0, 581, 720], {
			...fadingGradient,
			fillGradientStops: [
				{ color: '#00405E', position: 0 },
				{ color: '#00405E', position: 100 },
			],
		});
		const bleeds = resolveAbuttingEdgeBleeds([object8, object2]);
		expect([...bleeds]).toEqual([['o8', { left: { color: '#00405E' } }]]);
	});

	it('never grows the upper box, and grows a solid lower box toward each opaque neighbour', () => {
		const lower = rect('lower', [100, 100, 100, 100], solid('#112233'));
		const right = rect('right', [200, 50, 80, 200], solid('#445566'));
		const above = rect('above', [80, 20, 140, 80], solid('#778899'));
		const bleeds = resolveAbuttingEdgeBleeds([lower, right, above]);
		expect(bleeds.get('lower')).toEqual({ right: { color: '#112233' }, top: { color: '#112233' } });
		expect(bleeds.has('right')).toBe(false);
		expect(bleeds.has('above')).toBe(false);
	});

	it('does nothing when the boxes only touch at a gap or overlap', () => {
		const lower = rect('lower', [0, 0, 100, 100], solid('#112233'));
		expect(
			resolveAbuttingEdgeBleeds([lower, rect('gap', [101, 0, 50, 100], solid('#000'))]).size,
		).toBe(0);
		expect(
			resolveAbuttingEdgeBleeds([lower, rect('over', [99, 0, 50, 100], solid('#000'))]).size,
		).toBe(0);
	});

	it('needs an opaque, unrotated, plain rectangle above that covers the whole strip', () => {
		const lower = rect('lower', [0, 0, 100, 100], solid('#112233'));
		const cases: PptxElement[] = [
			rect('translucent', [100, 0, 50, 100], solid('#000', { fillOpacity: 0.5 })),
			rect('transparent', [100, 0, 50, 100], solid('transparent')),
			rect('faded', [100, 0, 50, 100], solid('#000'), { opacity: 0.8 }),
			rect('rotated', [100, 0, 50, 100], solid('#000'), { rotation: 90 }),
			rect('ellipse', [100, 0, 50, 100], solid('#000'), { shapeType: 'ellipse' }),
			rect('soft', [100, 0, 50, 100], solid('#000', { softEdgeRadius: 4 })),
			rect('short', [100, 10, 50, 80], solid('#000')),
			rect('alpha-stop', [100, 0, 50, 100], fadingGradient),
			rect('hidden', [100, 0, 50, 100], solid('#000'), { hidden: true }),
			{ ...rect('pic', [100, 0, 50, 100], solid('#000')), type: 'picture' } as PptxElement,
		];
		for (const upper of cases) {
			expect(resolveAbuttingEdgeBleeds([lower, upper]).size, upper.id).toBe(0);
		}
	});

	it('leaves a lower box that carries text, a line, effects or an unsupported fill alone', () => {
		const upper = rect('upper', [100, 0, 50, 100], solid('#000'));
		const lowers: PptxElement[] = [
			rect('text', [0, 0, 100, 100], solid('#123'), { text: 'Hello' }),
			rect('placeholder', [0, 0, 100, 100], solid('#123'), { placeholderType: 'body' }),
			rect(
				'line',
				[0, 0, 100, 100],
				solid('#123', { strokeWidth: 2, strokeColor: '#000', strokeFillMode: 'solid' }),
			),
			rect('shadow', [0, 0, 100, 100], solid('#123', { shadowColor: '#000' })),
			rect('translucent-edge', [0, 0, 100, 100], solid('#123', { fillOpacity: 0.4 })),
			rect('pattern', [0, 0, 100, 100], { fillMode: 'pattern', fillColor: '#123' }),
			rect('oblique', [0, 0, 100, 100], { ...fadingGradient, fillGradientAngle: 45 }),
			rect('flipped', [0, 0, 100, 100], fadingGradient, { flipHorizontal: true }),
		];
		for (const lower of lowers) {
			expect(resolveAbuttingEdgeBleeds([lower, upper]).size, lower.id).toBe(0);
		}
	});

	it('reads a gradient edge colour from the end that meets the shared edge', () => {
		const upperLeft = rect('ul', [0, 0, 100, 100], solid('#000'));
		const reversed = rect('rev', [100, 0, 100, 100], {
			...fadingGradient,
			fillGradientAngle: 180,
			fillGradientStops: [
				{ color: '#AA0000', position: 0 },
				{ color: '#00AA00', position: 100 },
			],
		});
		expect(resolveAbuttingEdgeBleeds([reversed, upperLeft]).get('rev')).toEqual({
			left: { color: '#00AA00' },
		});
		// A vertical gradient runs along a vertical edge: the fill itself extends.
		const vertical = rect('vert', [100, 0, 100, 100], { ...fadingGradient, fillGradientAngle: 90 });
		expect(resolveAbuttingEdgeBleeds([vertical, upperLeft]).get('vert')).toEqual({ left: {} });
		// Across the edge, a transparent end cannot be extended.
		const transparentEdge = rect('t', [100, 0, 100, 100], {
			...fadingGradient,
			fillGradientAngle: 180,
		});
		expect(resolveAbuttingEdgeBleeds([transparentEdge, upperLeft]).size).toBe(0);
	});

	it('drops a side whose grown corner would stick out from under its cover', () => {
		const lower = rect('lower', [100, 100, 100, 100], solid('#112233'));
		const left = rect('left', [0, 100, 100, 100], solid('#000')); // exactly the lower's height
		const top = rect('top', [100, 0, 100, 100], solid('#000')); // exactly the lower's width
		const bleed = resolveAbuttingEdgeBleeds([lower, left, top]).get('lower');
		// Growing both would put the top-left corner under neither cover.
		expect(Object.keys(bleed ?? {})).toHaveLength(1);
	});

	it('skips excluded and animated elements on either side of the pair', () => {
		const lower = rect('lower', [0, 0, 100, 100], solid('#112233'), { shapeId: '7' });
		const upper = rect('upper', [100, 0, 50, 100], solid('#000'));
		expect(resolveAbuttingEdgeBleeds([lower, upper], { excludeIds: new Set(['upper']) }).size).toBe(
			0,
		);
		expect(resolveAbuttingEdgeBleeds([lower, upper], { excludeIds: new Set(['lower']) }).size).toBe(
			0,
		);
		expect(resolveAbuttingEdgeBleeds([lower, upper], { animatedIds: new Set(['7']) }).size).toBe(0);
		expect(
			resolveAbuttingEdgeBleeds([lower, upper], { animatedIds: new Set(['upper']) }).size,
		).toBe(0);
		expect(resolveAbuttingEdgeBleeds([lower, upper]).size).toBe(1);
	});
});

describe('resolveStageEdgeBleeds', () => {
	const lower = rect('lower', [0, 0, 100, 100], solid('#112233'), { shapeId: '3' });
	const upper = rect('upper', [100, 0, 50, 100], solid('#000'));

	it('stacks template elements under the slide and honours the selection', () => {
		const slide = { id: 's', elements: [upper] } as unknown as PptxSlide;
		expect([...resolveStageEdgeBleeds([lower], slide).keys()]).toEqual(['lower']);
		expect(resolveStageEdgeBleeds([lower], slide, new Set(['upper'])).size).toBe(0);
		// Reversed stacking: the template box is now the lower one and grows instead.
		const reversed = { id: 's', elements: [lower] } as unknown as PptxSlide;
		expect([...resolveStageEdgeBleeds([upper], reversed)]).toEqual([
			['upper', { left: { color: '#000' } }],
		]);
	});

	it('keeps boxes exact for animated elements and on a slide entered by Morph', () => {
		const animated = {
			id: 's',
			elements: [lower, upper],
			nativeAnimations: [{ targetId: '3' }],
		} as unknown as PptxSlide;
		expect(resolveStageEdgeBleeds(undefined, animated).size).toBe(0);
		const morph = {
			id: 's',
			elements: [lower, upper],
			transition: { type: 'morph' },
		} as unknown as PptxSlide;
		expect(resolveStageEdgeBleeds(undefined, morph).size).toBe(0);
		expect(resolveStageEdgeBleeds(undefined, undefined).size).toBe(0);
	});
});

describe('slideAnimatedElementIds', () => {
	it('collects editor and native animation targets', () => {
		const slide = {
			animations: [{ elementId: 'a' }],
			nativeAnimations: [{ targetId: '12' }, {}],
		} as unknown as PptxSlide;
		expect([...slideAnimatedElementIds(slide)].sort()).toEqual(['12', 'a']);
		expect(slideAnimatedElementIds(undefined).size).toBe(0);
	});
});
