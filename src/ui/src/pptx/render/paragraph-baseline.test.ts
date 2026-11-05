import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	applyParagraphBaselineShift,
	cssBaseline,
	type FontBoxMeasure,
	powerPointBaseline,
	resolveParagraphBaselineShift,
} from './paragraph-baseline';
import type { ParagraphRun, RenderParagraph } from './paragraph-types';

// Calibri's Windows metrics (the ones Chromium lays lines out with): ascent
// 0.952 em, descent 0.269 em.
const calibri: FontBoxMeasure = (font) => {
	const size = Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1]);
	return { ascent: 0.952 * size, descent: 0.269 * size };
};
const run = (text: string, style: ParagraphRun['style'] = {}): ParagraphRun => ({ text, style });
const shiftOf = (lineHeight: number | string, runs = [run('Hxlb')]) =>
	resolveParagraphBaselineShift({
		runs,
		lineHeight,
		strutSizePx: 40,
		strutFamily: 'Calibri',
		measure: calibri,
	});

describe('powerPointBaseline', () => {
	// PowerPoint 16 exports, Calibri 40pt (baselines from the line top, pt).
	it('puts a line taller than single spacing at 75% of its pitch', () => {
		expect(powerPointBaseline(96, 40, 10.75)).toBe(72); // 200%: 72.5 measured with overshoot
		expect(powerPointBaseline(57.6, 40, 10.75)).toBeCloseTo(43.2, 10); // 120%: 43.75
		expect(powerPointBaseline(49, 40, 10.75)).toBe(36.75); // 102%: 37.25, above 100%'s 38
	});

	it('keeps the descent at the bottom of a single or tighter line until 75% is lower', () => {
		expect(powerPointBaseline(48, 40, 10.75)).toBe(37.25); // 100%: 38.0 measured
		expect(powerPointBaseline(43.2, 40, 10.75)).toBeCloseTo(32.45, 10); // 90%: 33.25
		expect(powerPointBaseline(36, 40, 10.75)).toBe(27); // 75%: 27.5
		expect(powerPointBaseline(30, 40, 10.75)).toBe(22.5); // exact 30pt: 23.0
	});
});

describe('cssBaseline', () => {
	it('centres each line height on its font and takes the lowest baseline', () => {
		const font = { ascent: 38, descent: 10 };
		expect(cssBaseline([{ lineHeight: 48, font }])).toBe(38);
		expect(
			cssBaseline([
				{ lineHeight: 96, font },
				{ lineHeight: 0, font },
			]),
		).toBe(62);
	});
});

describe('resolveParagraphBaselineShift', () => {
	// From single spacing (48 px pitch), where CSS puts the baseline at
	// (48 + 27.32) / 2 = 37.66 and PowerPoint at 48 - 10.76 = 37.24.
	it('moves a 200% paragraph down to 75% of its pitch (measured 61.5pt -> 72.3pt, PowerPoint 72.5pt)', () => {
		// CSS moves (96 - 48) / 2 = 24; PowerPoint 0.75 x 96 - 37.24 = 34.76.
		expect(shiftOf(2.4)).toBeCloseTo(10.76, 2);
	});

	it('moves a 75% paragraph up (measured 31.8pt -> 27.6pt, PowerPoint 27.5pt)', () => {
		// CSS moves (36 - 48) / 2 = -6; PowerPoint max(27, 36 - 10.76) - 37.24 = -10.24.
		expect(shiftOf(0.9)).toBeCloseTo(-4.24, 2);
	});

	it('leaves single spacing where the renderer already puts it', () => {
		expect(shiftOf(1.2)).toBeUndefined();
	});

	it('places an exactly spaced line on its strut alone', () => {
		const runs = [run('Hxlb', { lineHeight: 0 })];
		// 40 px pitch: CSS (40 + 27.32) / 2 = 33.66, -4 from single; PowerPoint
		// max(30, 29.24) = 30, -7.24.
		expect(shiftOf('30pt', runs)).toBeCloseTo(-3.24, 2);
	});

	it('leaves alone what it cannot model', () => {
		expect(shiftOf(2.4, [run('a', { fontSize: '40px' }), run('b', { fontSize: '20px' })])).toBe(
			undefined,
		);
		expect(shiftOf(2.4, [run('a', { verticalAlign: 'super' })])).toBeUndefined();
		expect(shiftOf(2.4, [run('')])).toBeUndefined();
		expect(
			resolveParagraphBaselineShift({
				runs: [run('a')],
				lineHeight: 2.4,
				strutSizePx: 40,
				strutFamily: 'Calibri',
				measure: () => undefined,
			}),
		).toBeUndefined();
	});
});

describe('applyParagraphBaselineShift', () => {
	const paragraph = (): RenderParagraph => ({
		runs: [run('Hxlb')],
		bulletStyle: {},
		segmentIndices: [0],
		lineHeight: 2.4,
	});
	const block = { fontFamily: 'Calibri', fontSizePx: 40 };

	it('moves the paragraph with a relative top every binding spreads', () => {
		const para = paragraph();
		applyParagraphBaselineShift(para, {} as TextStyle, block, calibri);
		expect(para.baselineShiftPx).toBeCloseTo(10.76, 2);
		expect(para.paragraphStyle).toStrictEqual({ position: 'relative', top: '10.76px' });
	});

	it('leaves vertical text, compatLnSpc and shrink-on-overflow bodies alone', () => {
		for (const body of [
			{ textDirection: 'vert' },
			{ compatibleLineSpacing: true },
			{ autoFitMode: 'shrink' },
		] as TextStyle[]) {
			const para = paragraph();
			applyParagraphBaselineShift(para, body, block, calibri);
			expect(para.paragraphStyle).toBeUndefined();
		}
	});
});
