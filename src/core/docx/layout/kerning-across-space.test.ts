import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutFontSpec, TextMeasurer } from './measure';
import type { LayoutParagraph } from './input';

// Issue #384: Word kerns a pair that has the space in it (Y and the space, the space and T). The
// layout counts such a pair in the width of the token on its left, and reports it on the fragment
// (`kernAfterPx`) so the export, which draws that token at its own width, can tell it apart.

/** 10 px per character, 4 px less for the pair "Y " and 3 px less for " T", when kerning is on. */
const measurer: TextMeasurer = {
	widthOf(text: string, font: LayoutFontSpec) {
		let width = text.length * 10;
		if (font.kerning === 'normal') {
			if (text.includes('Y ')) width -= 4 * text.split('Y ').length - 4;
			if (text.includes(' T')) width -= 3 * text.split(' T').length - 3;
		}
		return width;
	},
	lineHeightOf: () => 12,
};

function widths(kerningThresholdPt: number | undefined) {
	const paragraph = {
		runs: [
			{
				text: 'Y T',
				fontFamily: 'X',
				fontSizePt: 12,
				...(kerningThresholdPt === undefined ? {} : { kerningThresholdPt }),
			},
		],
	} as unknown as LayoutParagraph;
	const result = layoutParagraph(paragraph, 400, measurer, () => {});
	return result.lines[0]!.fragments.map((fragment) => ({
		text: fragment.text,
		widthPx: fragment.widthPx,
		kernAfterPx: fragment.kernAfterPx ?? 0,
	}));
}

describe('kerning pairs with the space', () => {
	it('adds the pair to the width of the glyph on its left', () => {
		expect(widths(1)).toEqual([
			{ text: 'Y', widthPx: 6, kernAfterPx: -4 },
			{ text: ' ', widthPx: 7, kernAfterPx: -3 },
			{ text: 'T', widthPx: 10, kernAfterPx: 0 },
		]);
	});

	it('is the control without w:kern', () => {
		expect(widths(undefined).map((fragment) => fragment.widthPx)).toEqual([10, 10, 10]);
		expect(widths(0).map((fragment) => fragment.widthPx)).toEqual([10, 10, 10]);
	});
});
