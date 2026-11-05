import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	applyParagraphBaselineShift,
	cssBaseline,
	type FontBoxMeasure,
	powerPointBaseline,
	powerPointDescent,
	resolveParagraphBaselineShift,
} from './paragraph-baseline';
import type { ParagraphRun, RenderParagraph } from './paragraph-types';

// Windows ascent and descent in em (the ones Chromium lays lines out with).
const metrics = (ascent: number, descent: number): FontBoxMeasure => {
	return (font) => {
		const size = Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1]);
		return { ascent: ascent * size, descent: descent * size };
	};
};
const calibri = metrics(0.952, 0.269);
const arial = metrics(0.905, 0.212);
const segoe = metrics(1.079, 0.251);
const run = (text: string, style: ParagraphRun['style'] = {}): ParagraphRun => ({ text, style });
const shiftOf = (
	lineHeight: number | string,
	runs = [run('Hxlb')],
	measure: FontBoxMeasure = calibri,
) =>
	resolveParagraphBaselineShift({
		runs,
		lineHeight,
		strutSizePx: 40,
		strutFamily: 'Calibri',
		measure,
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

describe('powerPointDescent', () => {
	// PowerPoint 16 single spacing at 40pt: baselines 0.940 em (Calibri), 0.975
	// (Arial), 0.975 (Segoe UI) of the 1.2 em pitch scale the font's own box.
	it('scales the font ascent and descent to fill the 1.2 pitch', () => {
		expect(powerPointDescent(40, { ascent: 0.952, descent: 0.269 })).toBeCloseTo(10.57, 2);
		expect(powerPointDescent(40, { ascent: 0.905, descent: 0.212 })).toBeCloseTo(9.11, 2);
		expect(powerPointDescent(40, { ascent: 1.079, descent: 0.251 })).toBeCloseTo(9.06, 2);
	});

	it('falls back to the font descent for an empty box', () => {
		expect(powerPointDescent(40, { ascent: 0, descent: 0 })).toBe(0);
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

	it('rounds the font box to whole px and floors the leading above it', () => {
		// Calibri at 53.33px in a 64px line: 51 + floor((64 - 51 - 14) / 2) = 50.
		expect(cssBaseline([{ lineHeight: 64, font: { ascent: 50.77, descent: 14.35 } }])).toBe(50);
	});
});

describe('resolveParagraphBaselineShift', () => {
	// Calibri at 40 px: CSS rounds the box to 38 + 11 and floors the leading.
	it('moves a 200% paragraph down to 75% of its pitch (measured 61.5pt -> 72.3pt, PowerPoint 72.5pt)', () => {
		// CSS 38 + floor((96 - 49) / 2) = 61; PowerPoint 0.75 x 96 = 72.
		expect(shiftOf(2.4)).toBeCloseTo(11, 2);
	});

	it('moves a 90% paragraph up (measured 31.8pt -> 27.6pt, PowerPoint 27.5pt)', () => {
		// CSS 38 + floor((36 - 49) / 2) = 31; PowerPoint max(27, 36 - 10.57) = 27.
		expect(shiftOf(0.9)).toBeCloseTo(-4, 2);
	});

	it('places single spacing on the font own proportions, whatever CSS gives it', () => {
		// PowerPoint 48 - 10.57 = 37.43 in every case; CSS differs per font.
		expect(shiftOf(1.2)).toBeCloseTo(0.43, 2); // Calibri: CSS 38 + floor(-0.5) = 37
		expect(shiftOf(1.2, [run('Hxlb')], arial)).toBeCloseTo(0.89, 2); // Arial sat high: CSS 38
		expect(shiftOf(1.2, [run('Hxlb')], segoe)).toBeCloseTo(-1.06, 2); // Segoe UI sat low: CSS 40
	});

	it('places an exactly spaced line on its strut alone', () => {
		const runs = [run('Hxlb', { lineHeight: 0 })];
		// 40 px pitch: the strut puts CSS at 38 + floor((40 - 49) / 2) = 33;
		// PowerPoint max(30, 40 - 10.57) = 30.
		expect(shiftOf('30pt', runs)).toBeCloseTo(-3, 2);
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
		expect(para.baselineShiftPx).toBeCloseTo(11, 2);
		expect(para.paragraphStyle).toStrictEqual({ position: 'relative', top: '11px' });
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
