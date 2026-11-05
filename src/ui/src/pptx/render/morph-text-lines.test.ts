import { describe, expect, it } from 'vitest';

import type { FrameParagraph, MorphTextMeasure, TextFrame } from './morph-text-frame-read';
import type { MorphTextInkMeasure } from './morph-text-ink';
import { layoutMorphTextWords } from './morph-text-lines';

/** A monospace stand-in for the canvas: every character is half the font size wide. */
const measure: MorphTextMeasure = (text, font) => {
	const [, size = '0'] = /([\d.]+)px/u.exec(font) ?? [];
	return [...text].length * Number(size) * 0.5;
};

/** Ink of that face: 0.7 em above the baseline and 0.2 em below, 0.8 em ascent and 0.2 em descent. */
const ink: MorphTextInkMeasure = (text, font) => {
	const size = Number(/([\d.]+)px/u.exec(font)?.[1] ?? 0);
	return {
		left: 0,
		right: [...text].length * size * 0.5,
		ascent: size * 0.7,
		descent: size * 0.2,
		fontAscent: size * 0.8,
		fontDescent: size * 0.2,
	};
};

function paragraph(
	text: string,
	size: number,
	extra: Partial<FrameParagraph> = {},
	runs: { text: string; size: number; family?: string }[] = [{ text, size }],
): FrameParagraph {
	return {
		text,
		fonts: runs.map((run) => `normal ${run.size}px ${run.family ?? 'Calibri'}`),
		size: Math.max(...runs.map((run) => run.size)),
		runs: runs.map((run) => ({ text: run.text, style: { fontSize: `${run.size}px` } })),
		strut: Math.max(size, ...runs.map((run) => run.size)),
		lineHeight: 1.2,
		spaceAfterPx: 0,
		marginLeftPx: 0,
		textIndentPx: 0,
		ownStrut: size,
		textAlign: 'left',
		textAlignLast: undefined,
		...extra,
	} as FrameParagraph;
}

function frame(paragraphs: FrameParagraph[], box: Partial<TextFrame> = {}): TextFrame {
	return {
		anchorX: 0,
		anchorY: 0,
		align: 'left',
		vAlign: 'top',
		available: 200,
		wraps: true,
		insetLeft: 0,
		insetTop: 0,
		insetBottom: 0,
		boxHeight: 200,
		anchorCenter: false,
		paragraphs,
		...box,
	};
}

// 20 px monospace: 10 px a character, 24 px a line, 20 characters a 200 px line.
const TEXT = 'aaaa bbbb cccc dddd eeee ffff';

describe('laying out a morphing text box word by word', () => {
	it('narrows and shifts the first line by its indent', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph(TEXT, 20, { textIndentPx: 30 })]),
			measure,
		)!;
		// 170 px of room on the first line holds three words, then four.
		expect(layout.words.map((word) => word.line)).toEqual([0, 0, 0, 1, 1, 1]);
		expect(layout.words[0]).toMatchObject({ x0: 30, x1: 70 });
		expect(layout.words[3]).toMatchObject({ x0: 0, x1: 40 });
	});

	it('hangs the first line out of an indented paragraph', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph(TEXT, 20, { marginLeftPx: 30, textIndentPx: -30 })]),
			measure,
		)!;
		expect(layout.words.map((word) => word.line)).toEqual([0, 0, 0, 0, 1, 1]);
		expect(layout.words[0]!.x0).toBe(0);
		expect(layout.words[4]!.x0).toBe(30);
	});

	it('spreads every line but the last of a justified paragraph over its spaces', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph(`${TEXT} gg`, 20, { textAlign: 'justify' })], { available: 210 }),
			measure,
		)!;
		// "aaaa bbbb cccc dddd" is 190 px wide: 20 px over three spaces.
		expect(layout.words.slice(0, 4).map((word) => word.x0)).toEqual(
			[0, 50, 100, 150].map((x, k) => expect.closeTo(x + (k * 20) / 3, 6)),
		);
		expect(layout.lines[0]).toMatchObject({ justified: true, x1: expect.closeTo(210, 6) });
		expect(layout.lines[1]).toMatchObject({ justified: false });
		expect(layout.words[4]!.x0).toBe(0);
		expect(layout.words[5]!.x0).toBe(50);
	});

	it('centres and right-aligns each line in its own room', () => {
		const centred = layoutMorphTextWords(
			frame([paragraph(TEXT, 20)], { align: 'center' }),
			measure,
		)!;
		expect(centred.lines.map((line) => line.x0)).toEqual([5, 55]);
		const right = layoutMorphTextWords(frame([paragraph(TEXT, 20)], { align: 'right' }), measure)!;
		expect(right.lines.map((line) => line.x0)).toEqual([10, 110]);
	});

	it('makes a line as tall as its largest run, and only that line', () => {
		const runs = [
			{ text: 'aaaa ', size: 20 },
			{ text: 'BB', size: 40 },
			{ text: ' cccc dddd eeee ffff gggg', size: 20 },
		];
		const layout = layoutMorphTextWords(
			frame([paragraph(runs.map((run) => run.text).join(''), 20, {}, runs)]),
			measure,
		)!;
		expect(layout.lines.map((line) => line.height)).toEqual(
			[48, 24].map((v) => expect.closeTo(v, 6)),
		);
		expect(layout.lines[1]!.top).toBeCloseTo(48, 6);
	});

	it('stacks paragraphs with the space after each', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph('aaaa', 20, { spaceAfterPx: 10 }), paragraph('bbbb', 20)]),
			measure,
		)!;
		expect(layout.lines.map((line) => line.top)).toEqual([0, 34]);
	});

	it('puts a middle or bottom anchored block where the flex column does', () => {
		const middle = layoutMorphTextWords(
			frame([paragraph(TEXT, 20)], {
				vAlign: 'middle',
				boxHeight: 100,
				insetTop: 4,
				insetBottom: 6,
			}),
			measure,
		)!;
		// 90 px of room for two 24 px lines.
		expect(middle.lines[0]!.top).toBeCloseTo(4 + 21, 6);
		const bottom = layoutMorphTextWords(
			frame([paragraph(TEXT, 20)], { vAlign: 'bottom', boxHeight: 100 }),
			measure,
		)!;
		expect(bottom.lines[1]!.top).toBeCloseTo(76, 6);
	});

	// A piece is a copy of its box and loses what spills past the box's edge:
	// on PowerPoint 16 morph renders carrying an overflowing anchored box piece
	// by piece measured worse than the box (middle 4.27 -> 5.35 mean abs diff).
	it('refuses an anchored block that overflows its box', () => {
		for (const vAlign of ['bottom', 'middle']) {
			expect(
				layoutMorphTextWords(frame([paragraph(TEXT, 20)], { vAlign, boxHeight: 40 }), measure),
			).toBeUndefined();
		}
	});

	// PowerPoint 16 exports: `anchorCtr` centres the paragraphs as one block as
	// wide as the longest, each aligned inside it ("Short line" starts where
	// the longer line below it starts).
	it('centres a shrink-wrapped block as wide as its widest paragraph', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph('aaaa', 20), paragraph('bbbb cccc', 20)], { anchorCenter: true }),
			measure,
		)!;
		// 90 px wide in a 200 px column: 55 px either side.
		expect(layout.words.map((word) => word.x0)).toEqual([55, 55, 105]);
		const right = layoutMorphTextWords(
			frame([paragraph('aaaa', 20), paragraph('bbbb cccc', 20)], {
				anchorCenter: true,
				align: 'right',
			}),
			measure,
		)!;
		expect(right.words.map((word) => word.x0)).toEqual([105, 55, 105]);
	});

	// PowerPoint 16 exports: a wrapped `anchorCtr` block shrinks to its longest
	// line and that box is centred, whatever the alignment.
	it('shrinks a wrapped anchorCtr block to its longest line and centres it', () => {
		const left = layoutMorphTextWords(
			frame([paragraph(TEXT, 20)], { available: 150, anchorCenter: true }),
			measure,
		)!;
		// "aaaa bbbb cccc" is 140 px of a 150 px column: 5 px either side.
		expect(left.words.map((word) => word.line)).toEqual([0, 0, 0, 1, 1, 1]);
		expect(left.words.map((word) => word.x0)).toEqual([5, 55, 105, 5, 55, 105]);
		const right = layoutMorphTextWords(
			frame([paragraph('aaaa bbbb cccc dddd eeee ffff gg', 20)], {
				available: 150,
				align: 'right',
				anchorCenter: true,
			}),
			measure,
		)!;
		// Lines of 140, 140 and 20 px: the block is 140 px, every line ends at 145.
		expect(right.lines.map((line) => line.x1)).toEqual([145, 145, 145]);
	});

	it('spreads a justified line over every space, a doubled one twice', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph('aaaa  bbbb ccc ddddddd', 20, { textAlign: 'justify' })]),
			measure,
		)!;
		// Line 1 "aaaa  bbbb ccc" is 140 px with 3 spaces: 60 px slack, 20 each.
		expect(layout.words[1]!.x0).toBeCloseTo(60 + 2 * 20, 6);
		expect(layout.words[2]!.x0).toBeCloseTo(110 + 3 * 20, 6);
		expect(layout.lines[0]!.x1).toBeCloseTo(200, 6);
	});

	it('places each word ink on its line baseline', () => {
		const layout = layoutMorphTextWords(frame([paragraph(TEXT, 20)]), measure, ink)!;
		// Half-leading 2 px above a 16 px ascent: the baseline is 18 px down.
		expect(layout.words[0]!.ink).toEqual({
			left: 0,
			right: 40,
			top: expect.closeTo(18 - 14, 6),
			bottom: expect.closeTo(18 + 4, 6),
		});
		expect(layout.words[4]!.ink!.top).toBeCloseTo(24 + 4, 6);
	});

	it('moves the ink, not the lines, by the paragraph baseline shift', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph(TEXT, 20, { baselineShiftPx: 3 })]),
			measure,
			ink,
		)!;
		expect(layout.lines[0]!.top).toBe(0);
		expect(layout.words[0]!.ink!.top).toBeCloseTo(18 - 14 + 3, 6);
	});

	it('starts below the first paragraph own space before', () => {
		const layout = layoutMorphTextWords(
			frame([paragraph('aaaa', 20, { spaceBeforePx: 10, spaceAfterPx: 6 }), paragraph('bbbb', 20)]),
			measure,
		)!;
		expect(layout.lines.map((line) => line.top)).toEqual([10, 40]);
	});

	it('refuses what it cannot reproduce', () => {
		const mixed = [
			{ text: 'aaaa ', size: 20 },
			{ text: 'bbbb', size: 30, family: 'Cambria' },
		];
		const refused = [
			// Mixed sizes in different faces, and under exact spacing.
			paragraph('aaaa bbbb', 20, {}, mixed),
			paragraph('aaaa bbbb', 20, { lineHeight: '20pt' }, [mixed[0]!, { text: 'bbbb', size: 30 }]),
			// Distributed alignment, and a non-breaking space a justified line spreads.
			paragraph(TEXT, 20, { textAlign: 'justify', textAlignLast: 'justify' }),
			paragraph('aaaa\u00a0bbbb', 20, { textAlign: 'justify' }),
		];
		for (const p of refused) {
			expect(layoutMorphTextWords(frame([p]), measure)).toBeUndefined();
		}
		// A shrink-wrapped block with a word as wide as the column, or a line break.
		expect(
			layoutMorphTextWords(frame([paragraph('a'.repeat(20), 20)], { anchorCenter: true }), measure),
		).toBe(undefined);
		expect(
			layoutMorphTextWords(frame([paragraph('aaaa\nbbbb', 20)], { anchorCenter: true }), measure),
		).toBe(undefined);
		// Several paragraphs spread over a distributed anchor.
		expect(
			layoutMorphTextWords(
				frame([paragraph('aaaa', 20), paragraph('bbbb', 20)], {
					vAlign: 'middle',
					anchorSpread: true,
				}),
				measure,
			),
		).toBe(undefined);
	});
});
