import { describe, it, expect } from 'vitest';

import type { XmlObject } from '../../types';
import { PptxEffectDagExtractor } from './PptxEffectDagExtractor';
import type { PptxEffectDagExtractorContext } from './PptxEffectDagExtractor';

/**
 * Minimal context whose `parseColor`/`extractColorOpacity` understand a colour
 * *container* node (e.g. `a:solidFill` or a gradient stop) the same way the
 * real codec does: they look inside for an `a:srgbClr` child.
 */
function makeContext(
	overrides: Partial<PptxEffectDagExtractorContext> = {},
): PptxEffectDagExtractorContext {
	return {
		emuPerPx: 9525,
		parseColor: (node: XmlObject | undefined): string | undefined => {
			const clr = node?.['a:srgbClr'] as XmlObject | undefined;
			const val = clr?.['@_val'] ?? node?.['@_val'];
			return val ? `#${String(val)}` : undefined;
		},
		extractColorOpacity: (node: XmlObject | undefined): number | undefined => {
			const clr = node?.['a:srgbClr'] as XmlObject | undefined;
			const alpha = (clr?.['a:alpha'] as XmlObject | undefined)?.['@_val'];
			return alpha === undefined ? undefined : parseInt(String(alpha), 10) / 100000;
		},
		ensureArray: (value: unknown): XmlObject[] => {
			if (Array.isArray(value)) {
				return value as XmlObject[];
			}
			return value === undefined || value === null ? [] : [value as XmlObject];
		},
		...overrides,
	};
}

function dagWith(fillOverlay: XmlObject): XmlObject {
	return { 'a:effectDag': { 'a:fillOverlay': fillOverlay } };
}

describe('extractEffectDagStyle fill-overlay colour', () => {
	it('parses a solidFill overlay colour, opacity and blend mode', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle(
			dagWith({
				'@_blend': 'mult',
				'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000', 'a:alpha': { '@_val': '50000' } } },
			}),
		);
		expect(style.dagFillOverlayColor).toBe('#FF0000');
		expect(style.dagFillOverlayOpacity).toBe(0.5);
		expect(style.dagFillOverlayBlend).toBe('mult');
	});

	it('parses the first gradient stop colour for a gradFill overlay', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle(
			dagWith({
				'@_blend': 'screen',
				'a:gradFill': {
					'a:gsLst': {
						'a:gs': [
							{ 'a:srgbClr': { '@_val': '00FF00' } },
							{ 'a:srgbClr': { '@_val': '0000FF' } },
						],
					},
				},
			}),
		);
		expect(style.dagFillOverlayColor).toBe('#00FF00');
		expect(style.dagFillOverlayBlend).toBe('screen');
	});

	it('leaves the colour unset when the overlay carries no fill child', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle(dagWith({ '@_blend': 'over' }));
		expect(style.dagFillOverlayColor).toBeUndefined();
		expect(style.dagFillOverlayOpacity).toBeUndefined();
		expect(style.dagFillOverlayBlend).toBe('over');
	});

	it('does not set overlay colour when there is no fillOverlay at all', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle({ 'a:effectDag': { 'a:grayscl': {} } });
		expect(style.dagFillOverlayColor).toBeUndefined();
		expect(style.dagGrayscale).toBeTruthy();
	});
});

describe('extractEffectDagStyle through the shared DAG interpretation', () => {
	it('finds list effects and colour operations inside nested containers', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle({
			'a:effectDag': {
				'@_type': 'tree',
				'a:cont': {
					'@_type': 'tree',
					'a:outerShdw': {
						'@_blurRad': '38100',
						'@_dist': '19050',
						'@_dir': '5400000',
						'a:srgbClr': { '@_val': '000000' },
					},
					'a:grayscl': '',
				},
				'a:alphaModFix': { '@_amt': '40000' },
				'a:hsl': { '@_hue': '10800000', '@_sat': '25000' },
			},
		});
		expect(style.shadowColor).toBe('#000000');
		expect(style.shadowBlur).toBe(4);
		expect(style.shadowOffsetX).toBe(0);
		expect(style.shadowOffsetY).toBe(2);
		expect(style.dagGrayscale).toBe(true);
		expect(style.dagAlphaModFix).toBeCloseTo(40);
		expect(style.dagHslHue).toBe(180);
		expect(style.dagHslSaturation).toBeCloseTo(125);
		expect(style.effectDagApproximations).toBeUndefined();
		expect(style.effectDagXml).toBeDefined();
	});

	it('maps each colour operation with the schema defaults, the first of a kind winning', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle({
			'a:effectDag': {
				'@_type': 'tree',
				'a:biLevel': { '@_thresh': '75000' },
				'a:lum': { '@_bright': '-20000' },
				'a:tint': [
					{ '@_hue': '3600000', '@_amt': '50000' },
					{ '@_hue': '0', '@_amt': '0' },
				],
				'a:duotone': { 'a:srgbClr': [{ '@_val': '112233' }, { '@_val': 'AABBCC' }] },
				'a:alphaModFix': [{ '@_amt': '50000' }, { '@_amt': '50000' }],
			},
		});
		expect(style.dagBiLevel).toBe(75);
		expect(style.dagLumBrightness).toBeCloseTo(-20);
		expect(style.dagLumContrast).toBe(0);
		expect(style.dagTintHue).toBe(60);
		expect(style.dagTintAmount).toBe(50);
		expect(style.dagDuotone).toEqual({ color1: '#112233', color2: '#AABBCC' });
		expect(style.dagAlphaModFix).toBeCloseTo(25);
		expect(style.dagHslSaturation).toBeUndefined();
	});

	it('keeps an hsl without @sat at unchanged saturation', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle({
			'a:effectDag': { 'a:hsl': { '@_hue': '600000' } },
		});
		expect(style.dagHslHue).toBe(10);
		expect(style.dagHslSaturation).toBe(100);
		expect(style.dagHslLuminance).toBe(0);
	});

	it('reports what the interpretation approximates', () => {
		const extractor = new PptxEffectDagExtractor(makeContext());
		const style = extractor.extractEffectDagStyle({
			'a:effectDag': {
				'a:blend': { '@_blend': 'screen', 'a:cont': { 'a:glow': { '@_rad': '9525' } } },
				'a:effect': { '@_ref': 'fillLine' },
			},
		});
		expect(style.glowRadius).toBe(1);
		expect(style.effectDagApproximations).toEqual([
			'sibling outputs composited as one paint',
			'blend screen painted as over',
			'effect references not applied',
		]);
	});
});
