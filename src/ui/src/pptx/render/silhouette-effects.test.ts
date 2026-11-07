import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { buildSubpathFillOverlay, suppressesCssFill } from './subpath-fill-overlay';
import { resolveShapeGeometry } from './shape-geometry-cascade';
import {
	buildSilhouetteFilterMarkup,
	getSilhouetteEffects,
	usesSilhouetteEffects,
} from './silhouette-effects';
import { getSilhouetteFillLayer } from './silhouette-fill-layer';
import { getComputedEffectStyle, getShapeEffectFilterDefs } from './visual-effects';

function shape(shapeType: string, style: Record<string, unknown> = {}): PptxElement {
	return {
		type: 'shape',
		id: 'sp1',
		x: 0,
		y: 0,
		width: 200,
		height: 120,
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

describe('silhouette effects', () => {
	it('a triangle with a shadow is not clipped and paints its shadow as a filter', () => {
		const el = shape('triangle');
		expect(usesSilhouetteEffects(el)).toBe(true);
		expect(resolveShapeGeometry(el).kind).toBe('none');
		const fx = getComputedEffectStyle(el);
		expect(fx.boxShadow).toBeUndefined();
		expect(fx.filter).toContain('url(#shape-fx-sp1)');
		const defs = getShapeEffectFilterDefs(el);
		expect(defs?.filterMarkup).toContain('<filter id="shape-fx-sp1"');
		expect(defs?.filterMarkup).toContain('feGaussianBlur');
	});

	it('the fill moves to the SVG overlay and the css fill is dropped', () => {
		const el = shape('triangle');
		expect(suppressesCssFill(el)).toBe(true);
		const overlay = buildSubpathFillOverlay(el);
		expect(overlay?.silhouette).toBe(true);
		expect(overlay?.paints[0]?.fill).not.toBe('none');
	});

	it('a glow is a ring of offsets blurred by a sixth of its radius', () => {
		const el = shape('rightArrow', {
			shadowColor: undefined,
			glowColor: '#00C0FF',
			glowRadius: 18,
		});
		const markup = getSilhouetteEffects(el)?.filterMarkup ?? '';
		expect(markup).toContain('<feOffset');
		expect(markup).toContain('stdDeviation="3"');
		expect(getComputedEffectStyle(el).boxShadow).toBeUndefined();
	});

	it('leaves a filled rectangle, ellipse and round rectangle on box-shadow', () => {
		for (const type of ['rect', 'ellipse', 'roundRect']) {
			const el = shape(type);
			expect(usesSilhouetteEffects(el)).toBe(false);
			expect(getComputedEffectStyle(el).boxShadow).toContain('rgba');
		}
	});

	it('shows the shadow through a translucent rectangle', () => {
		const el = shape('rect', { fillOpacity: 0.5 });
		expect(usesSilhouetteEffects(el)).toBe(true);
		expect(resolveShapeGeometry(el).kind).toBe('none');
		expect(getComputedEffectStyle(el).boxShadow).toBeUndefined();
	});

	it('shadows only the outline of a shape with no fill', () => {
		const el = shape('rect', { fillMode: 'none', fillColor: undefined });
		expect(usesSilhouetteEffects(el)).toBe(true);
		expect(getComputedEffectStyle(el).boxShadow).toBeUndefined();
		expect(buildSubpathFillOverlay(el)).toBeUndefined();
	});

	it('keeps the old paint when other effects share the box-shadow', () => {
		expect(usesSilhouetteEffects(shape('triangle', { lineShadowColor: '#000000' }))).toBe(false);
		expect(usesSilhouetteEffects(shape('triangle', { fillMode: 'gradient' }))).toBe(true);
	});

	it('paints the inner shadow of a clipped shape in the filter, not the box-shadow', () => {
		const inner = {
			innerShadowColor: '#000000',
			innerShadowOffsetX: 8,
			innerShadowOffsetY: 0,
			innerShadowBlur: 10,
		};
		const triangle = shape('triangle', inner);
		const fx = getSilhouetteEffects(triangle);
		expect(fx?.paintsInnerShadow).toBe(true);
		expect(fx?.filterMarkup).toContain('innerLayer');
		// The offset points at the edge it darkens: the inverse alpha moves the other way.
		expect(fx?.filterMarkup).toContain('dx="-8"');
		expect(getComputedEffectStyle(triangle).boxShadow).toBeUndefined();
		// A rectangle keeps its inset box-shadow.
		const rect = shape('rect', inner);
		expect(usesSilhouetteEffects(rect)).toBe(false);
		expect(getComputedEffectStyle(rect).boxShadow).toContain('inset -8px 0px');
	});

	it('shades a clipped shape by its inner shadow alone', () => {
		const only = shape('triangle', { shadowColor: undefined, innerShadowColor: '#000000' });
		expect(getSilhouetteEffects(only)?.paintsInnerShadow).toBe(true);
	});
});

describe('gradient, pattern and picture fills', () => {
	const fx = {
		shadowColor: '#000000',
		shadowOpacity: 0.5,
		shadowBlur: 10,
		shadowAngle: 45,
		shadowDistance: 8,
		outerShadowXml: {},
	};
	const mk = (shapeType: string, shapeStyle: Record<string, unknown>) =>
		({
			id: 'g1',
			type: 'shape',
			x: 0,
			y: 0,
			width: 160,
			height: 130,
			shapeType,
			shapeStyle: { ...fx, ...shapeStyle },
		}) as unknown as PptxElement;
	const gradient = {
		fillColor: '#2070C0',
		fillMode: 'gradient',
		fillGradientType: 'linear',
		fillGradientAngle: 90,
		fillGradientStops: [
			{ color: '#FF0000', position: 0 },
			{ color: '#0000FF', position: 100 },
		],
	};

	it('paints a non-rectangular gradient shape as a silhouette with a fill layer', () => {
		const triangle = mk('triangle', gradient);
		expect(usesSilhouetteEffects(triangle)).toBe(true);
		const layer = getSilhouetteFillLayer(triangle);
		expect(layer?.style['backgroundImage']).toContain('linear-gradient');
		expect(layer?.style['clipPath']).toBeTruthy();
		expect(resolveShapeGeometry(triangle).kind).toBe('none');
	});

	it('leaves a gradient rectangle on its box-shadow, untouched', () => {
		const rect = mk('rect', gradient);
		expect(usesSilhouetteEffects(rect)).toBe(false);
		expect(getSilhouetteFillLayer(rect)).toBeUndefined();
	});

	it('scales a layer by an attribute, never a CSS property, so a hidden copy of the filter agrees', () => {
		// Copies of one shape's filter share an id and `url(#id)` takes the first; one
		// in a display:none subtree (the editor canvas under a show) resolves no CSS,
		// which painted every shadow at opacity 1 in the Vue and vanilla shows.
		const markup = getSilhouetteEffects(shape('triangle'))?.filterMarkup ?? '';
		expect(markup).not.toContain('flood-opacity');
		expect(markup).toContain('<feFuncA type="linear" slope="0.5"/>');
	});

	it('an inner shadow blurs by 0.327 of its blurRad, like the outer one', () => {
		const el = shape('triangle', {
			shadowColor: undefined,
			innerShadowColor: '#000000',
			innerShadowBlur: 16,
			innerShadowOpacity: 0.7,
		});
		const markup = getSilhouetteEffects(el)?.filterMarkup ?? '';
		// 16 px of blurRad: CSS blur 10.46, deviation 5.23.
		expect(markup).toContain('stdDeviation="5.23"');
	});
});

describe('buildSilhouetteFilterMarkup', () => {
	it('escapes the filter id it writes into the markup', () => {
		const markup = buildSilhouetteFilterMarkup('shape-fx-a"><img src=x onerror=alert(1)>', 10, 10, {
			shadows: [],
			glow: { spread: 1, blur: 2, color: '#ff0000', opacity: 1 },
		});
		expect(markup).toContain('id="shape-fx-a&quot;&gt;&lt;img src=x onerror=alert(1)&gt;"');
		expect(markup).not.toContain('<img');
	});
});
