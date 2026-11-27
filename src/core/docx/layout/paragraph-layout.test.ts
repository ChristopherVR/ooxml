import { signedTwips, twips } from '../index';
import { describe, expect, it } from 'vitest';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutParagraph, LayoutRun } from './input';
import type { TextMeasurer } from './measure';
import { at } from './test-helpers';
import { tokenPosition } from './text-breaks';

// A measurer with clean, font-agnostic numbers: 10px per character, 20px lines.
const measurer: TextMeasurer = {
	widthOf: (text) => text.length * 10,
	lineHeightOf: () => 20,
};
const noop = () => {};
const lineText = (fragments: { text: string }[]) => fragments.map((f) => f.text).join('');
function paragraph(runs: LayoutRun[], overrides: Partial<LayoutParagraph> = {}): LayoutParagraph {
	return { kind: 'paragraph', id: 'p1', runs, ...overrides };
}

describe('layoutParagraph line breaking', () => {
	it('breaks oversized words at complete graphemes and resumes from inside a word', () => {
		const grapheme = 'e\u0301';
		const text = grapheme.repeat(150);
		const input = paragraph([{ text }]);
		const result = layoutParagraph(input, 55, measurer, noop);
		expect(result.lines).toHaveLength(75);
		for (const [index, line] of result.lines.entries()) {
			expect(lineText(line.fragments)).toBe(grapheme.repeat(2));
			expect(line).toMatchObject({ sourceStart: index * 4, sourceEnd: index * 4 + 4 });
			expect(at(line.fragments, 0).widthPx).toBe(40);
		}
		const continued = layoutParagraph(
			input,
			65,
			measurer,
			noop,
			undefined,
			tokenPosition(at(result.lines, 0).nextToken!, at(result.lines, 0).nextTokenOffset),
		);
		expect(at(continued.lines, 0)).toMatchObject({ sourceStart: 4, sourceEnd: 10 });
		expect(continued.lines.map((line) => lineText(line.fragments)).join('')).toBe(text.slice(4));
	});

	it('wraps at word boundaries to fit the available width', () => {
		const result = layoutParagraph(paragraph([{ text: 'aaaa bbbb cccc' }]), 90, measurer, noop);
		expect(result.lines).toHaveLength(2);
		expect(lineText(at(result.lines, 0).fragments)).toBe('aaaa bbbb');
		expect(lineText(at(result.lines, 1).fragments)).toBe('cccc');
	});

	it('honors manual line breaks without wrapping the rest of the width', () => {
		const result = layoutParagraph(paragraph([{ text: 'one\ntwo' }]), 200, measurer, noop);
		expect(result.lines).toHaveLength(2);
		expect(lineText(at(result.lines, 0).fragments)).toBe('one');
		expect(lineText(at(result.lines, 1).fragments)).toBe('two');
	});

	it('reports an explicit page break as a line boundary', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'first', breakAfter: 'page' }, { text: 'second' }]),
			200,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(2);
		expect(result.pageBreakAfterLine.has(0)).toBe(true);
		expect(result.columnBreakAfterLine.size).toBe(0);
		expect(lineText(at(result.lines, 0).fragments)).toBe('first');
		expect(lineText(at(result.lines, 1).fragments)).toBe('second');
	});

	it('advances tabs to the next default half-inch stop', () => {
		const result = layoutParagraph(paragraph([{ text: 'A\tB' }]), 200, measurer, noop);
		const [a, tab, b] = at(result.lines, 0).fragments;
		expect(a).toMatchObject({ text: 'A', xPx: 0, widthPx: 10 });
		expect(tab).toMatchObject({ text: '', xPx: 10, widthPx: 38 }); // 48px stop - 10px used
		expect(b).toMatchObject({ text: 'B', xPx: 48, widthPx: 10 });
	});
});

describe('layoutParagraph alignment', () => {
	it('keeps natural spaces at soft breaks, including consecutive and terminal breaks', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aa bb\n\ncc dd ee ff\n' }], {
				align: 'justify',
				doNotExpandShiftReturn: true,
			}),
			60,
			measurer,
			noop,
		);
		expect(result.lines.map((line) => lineText(line.fragments))).toEqual([
			'aa bb',
			'',
			'cc dd',
			'ee ff',
		]);
		expect(
			result.lines.flatMap((line) =>
				line.fragments.filter((f) => f.text === ' ').map((f) => f.widthPx),
			),
		).toEqual([10, 20, 10]);
	});

	it.each([
		['page', true],
		['page', false],
		['column', true],
		['column', false],
	] as const)(
		'keeps a line ending in a %s break at its natural width (doNotExpandShiftReturn: %s)',
		(breakAfter, doNotExpandShiftReturn) => {
			const result = layoutParagraph(
				paragraph([{ text: 'aa bb', breakAfter }, { text: 'cc dd ee ff' }], {
					align: 'justify',
					...(doNotExpandShiftReturn ? { doNotExpandShiftReturn } : {}),
				}),
				60,
				measurer,
				noop,
			);
			expect(at(at(result.lines, 0).fragments, 1)).toMatchObject({ text: ' ', widthPx: 10 });
			// A wrapped line after the break is still justified.
			expect(at(at(result.lines, 1).fragments, 1)).toMatchObject({ text: ' ', widthPx: 20 });
		},
	);

	it('can still tighten a line that slightly overflows before a soft break', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aaaa bbbb\ncc' }], { align: 'justify', doNotExpandShiftReturn: true }),
			89,
			measurer,
			noop,
		);
		expect(result.lines.map((line) => lineText(line.fragments))).toEqual(['aaaa bbbb', 'cc']);
		expect(at(at(result.lines, 0).fragments, 1)).toMatchObject({ text: ' ', widthPx: 9 });
	});

	it('distributes extra space on justified lines but not the paragraph’s last line', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'aa bb cc dd' }], { align: 'justify' }),
			60,
			measurer,
			noop,
		);
		expect(result.lines).toHaveLength(2);
		const spaceFragments = at(result.lines, 0).fragments.filter((f) => f.text === ' ');
		expect(spaceFragments.every((f) => f.widthPx > 10)).toBe(true);
		const lastLineSpaces = at(result.lines, 1).fragments.filter((f) => f.text === ' ');
		expect(lastLineSpaces.every((f) => f.widthPx === 10)).toBe(true);
	});

	it('applies firstLineTwips only to the paragraph’s first line', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], { align: 'right', firstLineTwips: twips(150) }),
			100,
			measurer,
			noop,
		);
		// The first line's box starts 10px in and still ends at the right edge: 10 + (90 - 20).
		expect(at(at(result.lines, 0).fragments, 0).xPx).toBe(80);
		expect(at(at(result.lines, 1).fragments, 0).xPx).toBe(80); // 100 - 20
	});

	it('applies hangingTwips as a first-line outdent', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], { align: 'right', hangingTwips: twips(150) }),
			100,
			measurer,
			noop,
		);
		// The first line's box starts 10px out (-10) and still ends at the right edge: -10 + (110 - 20).
		expect(at(at(result.lines, 0).fragments, 0).xPx).toBe(80);
		expect(at(at(result.lines, 1).fragments, 0).xPx).toBe(80); // 100 - 20
	});

	it('offsets left-aligned lines by the left indent and first-line indent', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'AB\nCD' }], {
				indentLeftTwips: signedTwips(300),
				firstLineTwips: twips(150),
			}),
			100,
			measurer,
			noop,
		);
		expect(at(at(result.lines, 0).fragments, 0).xPx).toBe(30); // 20 + 10
		expect(at(at(result.lines, 1).fragments, 0).xPx).toBe(20);
	});
});

describe('layoutParagraph line spacing rules', () => {
	it('scales natural height by the auto multiple', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: signedTwips(360), lineSpacingRule: 'auto' }),
			200,
			measurer,
			noop,
		);
		expect(at(result.lines, 0).heightPx).toBe(30); // 20 * 1.5
	});

	it('uses the exact height regardless of natural line height', () => {
		const result = layoutParagraph(
			paragraph([{ text: 'x' }], { lineSpacingTwips: signedTwips(300), lineSpacingRule: 'exact' }),
			200,
			measurer,
			noop,
		);
		expect(at(result.lines, 0).heightPx).toBe(20); // 300 twips = 20px, smaller than natural but forced
	});

	it('uses at-least as a floor over the natural height', () => {
		const tall = layoutParagraph(
			paragraph([{ text: 'x' }], {
				lineSpacingTwips: signedTwips(450),
				lineSpacingRule: 'atLeast',
			}),
			200,
			measurer,
			noop,
		);
		expect(at(tall.lines, 0).heightPx).toBe(30);
		const short = layoutParagraph(
			paragraph([{ text: 'x' }], {
				lineSpacingTwips: signedTwips(150),
				lineSpacingRule: 'atLeast',
			}),
			200,
			measurer,
			noop,
		);
		expect(at(short.lines, 0).heightPx).toBe(20);
	});
});

describe('layoutParagraph baselines', () => {
	// Ascent is 80% of a line height proportional to font size.
	const sized: TextMeasurer = {
		widthOf: (text) => text.length * 10,
		lineHeightOf: (font) => font.sizePx * 1.25,
	};

	it('aligns fragments of different sizes on one baseline, with extra spacing above', () => {
		const result = layoutParagraph(
			paragraph(
				[
					{ text: 'Big ', fontSizePt: 24 },
					{ text: 'small', fontSizePt: 12 },
				],
				{
					lineSpacingTwips: signedTwips(480),
					lineSpacingRule: 'auto',
				},
			),
			400,
			sized,
			noop,
		);
		const line = at(result.lines, 0);
		// 24pt = 32px → 40px natural (32 above, 8 below); double spacing doubles it to 80px.
		expect(line.heightPx).toBe(80);
		expect(line.baselinePx).toBe(72);
		const big = at(line.fragments, 0);
		const small = at(line.fragments, 2);
		expect(big.topPx! + big.boxHeightPx! * 0.8).toBe(72);
		expect(small.topPx! + small.boxHeightPx! * 0.8).toBe(72);
	});

	it('raises superscripts and sits inline pictures on the baseline', () => {
		const result = layoutParagraph(
			paragraph([
				{ text: 'x', fontSizePt: 12 },
				{ text: '2', fontSizePt: 12, script: 'super' },
				{
					text: '',
					object: { partName: 'p', contentType: 'image/png', widthPx: 10, heightPx: 30 },
				},
			]),
			400,
			sized,
			noop,
		);
		const line = at(result.lines, 0);
		const x = at(line.fragments, 0);
		const two = at(line.fragments, 1);
		const picture = at(line.fragments, 2);
		const baseline = line.baselinePx!;
		expect(x.topPx! + x.boxHeightPx! * 0.8).toBeCloseTo(baseline);
		expect(two.topPx! + two.boxHeightPx! * 0.8).toBeCloseTo(baseline - 0.33 * 16);
		expect(picture.topPx).toBeCloseTo(baseline - 30);
	});
});
