import { describe, expect, it } from 'vitest';

import type { MorphFontMetrics } from './morph-text-baseline';
import { morphTextFirstBaseline } from './morph-text-baseline';
import type { FrameParagraph, TextFrame } from './morph-text-frame-read';

/** A face whose ascent is 0.9 and descent 0.2 of its size. */
const metrics: MorphFontMetrics = (font) => {
	const size = Number.parseFloat(/([\d.]+)px/u.exec(font)?.[1] ?? '');
	return { ascent: 0.9 * size, descent: 0.2 * size };
};

function paragraph(extra: Partial<FrameParagraph> = {}): FrameParagraph {
	return {
		text: 'Lorem',
		fonts: ['normal 40px Calibri'],
		size: 40,
		runs: [{ text: 'Lorem', style: {} }] as FrameParagraph['runs'],
		strut: 40,
		strutFont: 'normal 40px Calibri',
		lineHeight: 1.5,
		spaceAfterPx: 0,
		marginLeftPx: 0,
		textIndentPx: 0,
		ownStrut: 40,
		textAlign: 'left',
		textAlignLast: undefined,
		indented: false,
		...extra,
	};
}

function frame(vAlign: string, paragraphs: FrameParagraph[]): TextFrame {
	return {
		anchorX: 0,
		anchorY: 0,
		align: 'left',
		vAlign,
		available: 400,
		wraps: false,
		insetLeft: 0,
		insetTop: 0,
		insetBottom: 0,
		boxHeight: 200,
		anchorCenter: false,
		paragraphs,
	};
}

// A 40 px line at 150%: CSS centres the 60 px line box on the face (baseline
// (60 + 36 - 8) / 2 = 44 px down); non-single spacing moves the glyphs onto
// PowerPoint's baseline with a relative `top` (`paragraph-baseline`), which
// the carried text has to follow or its halves land off each other by it.
describe('morphTextFirstBaseline', () => {
	it('follows the first paragraph baseline shift at every anchor', () => {
		const plain = [paragraph(), paragraph()];
		const shifted = [paragraph({ baselineShiftPx: 6 }), paragraph({ baselineShiftPx: 6 })];
		for (const anchor of ['top', 'middle', 'bottom']) {
			const before = morphTextFirstBaseline(frame(anchor, plain), metrics)!;
			const after = morphTextFirstBaseline(frame(anchor, shifted), metrics)!;
			expect(after - before).toBeCloseTo(6, 6);
		}
		expect(morphTextFirstBaseline(frame('top', plain), metrics)).toBeCloseTo(44, 6);
	});
});
