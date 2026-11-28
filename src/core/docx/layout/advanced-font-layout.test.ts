import { describe, expect, it } from 'vitest';
import type { LayoutParagraph } from './input';
import type { TextMeasurer } from './measure';
import { fontOf, runTextWidth } from './paragraph-tokens';
import { layoutParagraph } from './paragraph-layout';

const measurer: TextMeasurer = {
	widthOf: (text) => [...text].length * 10,
	lineHeightOf: () => 20,
	ascentOf: () => 15,
};
const paragraph = (runs: LayoutParagraph['runs']): LayoutParagraph => ({
	kind: 'paragraph',
	id: 'advanced',
	runs,
});

describe('advanced font line layout', () => {
	it('adds spacing after scale, matching Word COM measurements', () => {
		expect(
			runTextWidth(
				'MMMM',
				{ text: 'MMMM', textScalePercent: 200, characterSpacingPx: 2 },
				measurer,
			),
		).toBe(88);
		expect(
			runTextWidth(
				'MMMM',
				{ text: 'MMMM', textScalePercent: 50, characterSpacingPx: -1 },
				measurer,
			),
		).toBe(16);
	});

	it('uses natural advances for a zero character scale', () => {
		expect(runTextWidth('MMMM', { text: 'MMMM', textScalePercent: 0 }, measurer)).toBe(40);
	});

	it('counts spacing by graphemes rather than splitting combining marks or emoji', () => {
		const constant: TextMeasurer = { ...measurer, widthOf: () => 10 };
		expect(runTextWidth('e\u0301👩‍💻', { text: '', characterSpacingPx: 2 }, constant)).toBe(14);
	});

	it('uses scaled/expanded widths for line wrapping and positions', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aaaa bbbb', textScalePercent: 200, characterSpacingPx: 2 }]),
			100,
			measurer,
			() => {},
		);
		expect(
			result.lines.map((line) => line.fragments.map((fragment) => fragment.text).join('')),
		).toEqual(['aaaa', 'bbbb']);
		expect(result.lines[0]!.fragments[0]).toMatchObject({
			widthPx: 88,
			textScalePercent: 200,
			characterSpacingPx: 2,
		});
	});

	it('uses scaled decimal prefixes when placing decimal tabs', () => {
		const p = paragraph([{ text: '\t12.34', textScalePercent: 200 }]);
		p.tabStops = [{ posPx: 100, align: 'decimal' }];
		const line = layoutParagraph(p, 300, measurer, () => {}).lines[0]!;
		expect(line.fragments[0]!.widthPx).toBe(60);
		expect(line.fragments[1]!.xPx).toBe(60);
	});

	it('uses the kerning threshold and explicit off without changing font size', () => {
		expect(fontOf({ text: 'AV', fontSizePt: 12, kerningThresholdPt: 12 })).toMatchObject({
			kerning: 'normal',
			sizePx: 16,
		});
		expect(fontOf({ text: 'AV', fontSizePt: 11, kerningThresholdPt: 12 }).kerning).toBe('none');
		expect(fontOf({ text: 'AV', fontSizePt: 24, kerningThresholdPt: 0 }).kerning).toBe('none');
		expect(fontOf({ text: 'AV', fontSizePt: 12 }).kerning).toBeUndefined();
	});

	it('raises and lowers full-size fragments around a shared baseline', () => {
		const line = layoutParagraph(
			paragraph([{ text: 'a', positionPx: 4 }, { text: 'b' }, { text: 'c', positionPx: -4 }]),
			200,
			measurer,
			() => {},
		).lines[0]!;
		expect(line.heightPx).toBe(28);
		expect(line.fragments.map((fragment) => fragment.topPx)).toEqual([0, 4, 8]);
		expect(line.fragments.map((fragment) => fragment.boxHeightPx)).toEqual([20, 20, 20]);
	});
});

it.each([
	{ markerPt: 12, heightPx: 18.464, baselinePx: 15.008 },
	{ markerPt: 18, heightPx: 25.968, baselinePx: 22.512 },
])('measures a $markerPt pt Arial list marker with 12 pt Times New Roman text', (expected) => {
	// Word keeps the line at about 13.8 pt with a 12 pt marker (13.85 pt from the two fonts'
	// ascents with their leading and the larger descent, #395). A larger marker adds room
	// above the baseline while keeping the body text's descent.
	const mixedFonts: TextMeasurer = {
		...measurer,
		lineHeightOf: (font) => font.sizePx * 1.15,
		ascentOf: (font) => font.sizePx * (font.family === 'Arial' ? 0.938 : 0.934),
	};
	const numbered = paragraph([
		{
			text: '-\t',
			fontFamily: 'Arial',
			fontSizePt: expected.markerPt,
			synthetic: true,
			marker: { length: 1, alignment: 'left' },
		},
		{ text: 'Numbered', fontFamily: 'Times New Roman', fontSizePt: 12 },
	]);
	const { lines } = layoutParagraph(numbered, 300, mixedFonts, () => {});
	expect(lines[0]?.heightPx).toBeCloseTo(expected.heightPx);
	expect(lines[0]?.baselinePx).toBeCloseTo(expected.baselinePx);
});

it('takes the larger descent of a same-size list label and its text (#395)', () => {
	// Symbol 11 pt (ascent 11.06 pt, descent 2.42 pt) over Calibri 11 pt (10.47 pt, 2.95 pt): Word's
	// first line is Symbol's ascent plus Calibri's descent, 14.01 pt, not 13.48 pt.
	const fonts: TextMeasurer = {
		...measurer,
		lineHeightOf: (font) => font.sizePx * (font.family === 'Symbol' ? 1.2251 : 1.2207),
		ascentOf: (font) => font.sizePx * (font.family === 'Symbol' ? 1.0054 : 0.9521),
	};
	const bullet = paragraph([
		{
			text: '\t',
			fontFamily: 'Symbol',
			fontSizePt: 11,
			synthetic: true,
			marker: { length: 1, alignment: 'left' },
		},
		{ text: 'Item', fontFamily: 'Calibri', fontSizePt: 11 },
	]);
	const { lines } = layoutParagraph(bullet, 300, fonts, () => {});
	expect(lines[0]?.heightPx).toBeCloseTo(((11.06 + 2.95) * 4) / 3, 1);
	expect(lines[0]?.baselinePx).toBeCloseTo((11.06 * 4) / 3, 1);
});
