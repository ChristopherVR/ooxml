import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { parseTrueType } from '../../text/index';
import { pairKerning } from '../../text/truetype/kerning';
import { createEstimatedMeasurer, createFontFileMeasurer } from './font-measurer';

const font = (name: string) =>
	new Uint8Array(readFileSync(new URL(`./__fixtures__/fonts/${name}.subset.ttf`, import.meta.url)));
const px = (pt: number) => (pt * 96) / 72;
/** Latin subsets of the free fonts (`__fixtures__/fonts`, see its README). */
const referenceFonts = () => [
	{ family: 'Carlito', data: font('FixtureMetricSans-Regular') },
	{ family: 'Liberation Sans', data: font('LiberationSans-Regular') },
	{ family: 'Liberation Sans', bold: true, data: font('LiberationSans-Bold') },
	{ family: 'Liberation Serif', data: font('LiberationSerif-Regular') },
	{ family: 'Noto Serif', data: font('NotoSerif-Regular') },
];

describe('createFontFileMeasurer', () => {
	const measurer = createFontFileMeasurer(referenceFonts());

	it('gives the single line heights Word 16 lays these fonts out on', () => {
		// Measured in Word's PDFs: Carlito 11 pt and Noto Serif 11 pt at w:line="259" advance
		// 14.49 and 16.14 pt; Liberation Sans 10 pt single 11.5 pt (Windows metrics plus the hhea
		// gap they leave out; Noto Serif sets USE_TYPO_METRICS).
		const em = (family: string) => measurer.lineHeightOf({ family, sizePx: 100 }) / 100;
		expect(em('Carlito')).toBeCloseTo(2500 / 2048, 6);
		expect(em('Liberation Sans')).toBeCloseTo(2355 / 2048, 6);
		expect(em('Liberation Serif')).toBeCloseTo(2355 / 2048, 6);
		expect(em('Noto Serif')).toBeCloseTo(1.362, 6);
	});

	it('sums the advance widths, and kerns only where Word does', () => {
		const spec = { family: 'Liberation Sans', sizePx: px(20) };
		const plain = measurer.widthOf('AVATAR To', spec);
		const parsed = parseTrueType(font('LiberationSans-Regular'))!;
		const advances = [...'AVATAR To'].reduce(
			(sum, char) => sum + parsed.advance(parsed.glyphFor(char.codePointAt(0)!)),
			0,
		);
		expect(plain).toBeCloseTo((advances / parsed.unitsPerEm) * spec.sizePx, 6);
		expect(measurer.widthOf('AVATAR To', { ...spec, kerning: 'none' })).toBe(plain);
		expect(measurer.widthOf('AVATAR To', { ...spec, kerning: 'normal' })).toBeLessThan(plain - 1);
	});

	it('resolves a family through the substitution table and records it', () => {
		const fresh = createFontFileMeasurer(referenceFonts());
		const calibri = fresh.widthOf('Hello world', { family: 'Calibri', sizePx: 20 });
		expect(calibri).toBe(fresh.widthOf('Hello world', { family: 'Carlito', sizePx: 20 }));
		expect(fresh.families.get('Calibri')).toEqual({ family: 'Carlito', source: 'substitute' });
		fresh.widthOf('x', { family: 'Nowhere', sizePx: 20 });
		expect(fresh.families.get('Nowhere')).toBeUndefined();
		expect(fresh.families.has('Nowhere')).toBe(true);
	});

	it('estimates a family with no file from the Office advance tables', () => {
		const estimate = createEstimatedMeasurer();
		const fresh = createFontFileMeasurer([]);
		expect(fresh.widthOf('Hello', { family: 'Arial', sizePx: 20 })).toBe(
			estimate.widthOf('Hello', { family: 'Arial', sizePx: 20 }),
		);
		expect(estimate.widthOf('Hello', { family: 'Arial', sizePx: 20 })).toBeCloseTo(
			((722 + 556 + 222 + 222 + 556) / 1000) * 20,
			1,
		);
	});
});

describe('pairKerning', () => {
	it('reads GPOS pair adjustments', () => {
		const parsed = parseTrueType(font('LiberationSans-Regular'))!;
		const kern = pairKerning(parsed)!;
		const glyph = (char: string) => parsed.glyphFor(char.codePointAt(0)!);
		expect(kern(glyph('A'), glyph('V'))).toBeLessThan(0);
		expect(kern(glyph('T'), glyph('o'))).toBeLessThan(0);
		expect(kern(glyph('o'), glyph('o'))).toBe(0);
	});
});
