import type { TextStyle } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	applyParagraphBaselineShift,
	cssBaseline,
	fontSizePx,
	type FontBoxMeasure,
	powerPointBaseline,
	powerPointDescent,
	resolveParagraphBaselineShift,
} from './paragraph-baseline';
import { typoLineGapEm } from './font-line-gap';
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

describe('powerPointDescent with a line gap', () => {
	// Gabriola (typo ascent 0.684, descent 0.316, gap 0.70 em), PowerPoint 16 at 72pt:
	// single-spaced baseline 0.977 em of a 1.2 em pitch (95px of 96px measured).
	it('counts the typo line gap above the ascent when scaling to the pitch', () => {
		const box = { ascent: 684, descent: 316, lineGap: 700 };
		const size = 96;
		const baseline = powerPointBaseline(1.2 * size, size, powerPointDescent(size, box));
		expect(baseline / size).toBeCloseTo(0.977, 2);
		// Without the gap the same box would put it at 0.82 em, 15px high.
		const bare = powerPointBaseline(
			1.2 * size,
			size,
			powerPointDescent(size, { ...box, lineGap: 0 }),
		);
		expect(bare / size).toBeCloseTo(0.82, 2);
	});
});

describe('typoLineGapEm', () => {
	it('knows the flagged fonts that have a gap and nothing else', () => {
		expect(typoLineGapEm('Gabriola')).toBe(0.7);
		expect(typoLineGapEm(' gabriola ')).toBe(0.7);
		expect(typoLineGapEm('Calibri')).toBe(0);
		// From the Windows and Office table: measured 4px low at 72pt without it.
		expect(typoLineGapEm('Liberation Sans Narrow')).toBeCloseTo(0.131, 3);
		expect(typoLineGapEm(undefined)).toBe(0);
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

	it('places mixed sizes at an exact pitch above their single spacing on 75% of it', () => {
		// 30pt is 40 px; the runs sit at `line-height: 0`, so the 40 px strut
		// alone puts CSS at 38 + floor((40 - 49) / 2) = 33, PowerPoint at 30.
		const runs = [
			run('•', { fontSize: '30px', lineHeight: 0 }),
			run('Hxlb', { fontSize: '20px', lineHeight: 0 }),
		];
		expect(shiftOf('30pt', runs)).toBeCloseTo(-3, 2);
		// A run whose single spacing reaches the pitch keeps its own descent: not modelled.
		expect(shiftOf('30pt', [run('a', { fontSize: '40px' }), run('b', { fontSize: '20px' })])).toBe(
			undefined,
		);
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

describe('a missing legacy East Asian Office font', () => {
	// The browser paints a stand-in (Times New Roman here); PowerPoint lays the
	// line out on the named font's own 0.859 / 0.141 em box.
	const times = (provisional: boolean, loading = false): FontBoxMeasure => {
		return (font) => {
			const px = Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1]);
			return {
				ascent: 0.891 * px,
				descent: 0.216 * px,
				...(provisional ? { provisional: true } : {}),
				...(loading ? { loading: true } : {}),
			};
		};
	};
	const size = 106.67;
	const shiftFor = (family: string, measure: FontBoxMeasure) =>
		resolveParagraphBaselineShift({
			runs: [run('Hxlb', { fontFamily: family, fontSize: size })],
			lineHeight: 0.75,
			strutSizePx: size,
			strutFamily: 'Times New Roman',
			measure,
		});
	const stand = { ascent: 0.891 * size, descent: 0.216 * size };
	const named = { ascent: 0.859375 * 1000, descent: 0.140625 * 1000 };
	const expected =
		powerPointBaseline(0.75 * size, size, powerPointDescent(size, named)) -
		cssBaseline([
			{ lineHeight: 0.75 * size, font: stand },
			{ lineHeight: 0.75 * size, font: stand },
		]);

	it('places the line on the named font box, not the stand-in', () => {
		const shift = shiftFor('MS PMincho', times(true));
		expect(shift).toBeCloseTo(expected, 1);
		expect(shift).not.toBeCloseTo(shiftFor('Times New Roman', times(false)) ?? 0, 1);
	});

	it('places a missing Office font on its own metrics, line gap included', () => {
		// OCR A Extended (Windows ascent 0.858, descent 0.177, no gap) is an Office
		// cloud font the browser may not have: PowerPoint puts its single-spaced
		// baseline at 0.995 em of the 1.2 em line, 20px under the stand-in's.
		const shift = shiftFor('OCR A Extended', times(true));
		const office = { ascent: 0.8579 * 1000, descent: 0.1768 * 1000 };
		const want =
			powerPointBaseline(0.75 * size, size, powerPointDescent(size, office)) -
			cssBaseline([
				{ lineHeight: 0.75 * size, font: stand },
				{ lineHeight: 0.75 * size, font: stand },
			]);
		expect(shift).toBeCloseTo(want, 1);
		// At single spacing the descent decides, and it differs from the stand-in's.
		const single = (family: string) =>
			resolveParagraphBaselineShift({
				runs: [run('Hxlb', { fontFamily: family, fontSize: size })],
				lineHeight: 1.2,
				strutSizePx: size,
				strutFamily: 'Times New Roman',
				measure: times(true),
			});
		expect(single('OCR A Extended')).not.toBe(single('Foo Sans'));
		// Nirmala Text Semilight keeps its typographic gap (0.862 em) above the ascent.
		const withGap = { ...office, lineGap: 0.8623 * 1000 };
		expect(powerPointDescent(size, withGap)).toBeLessThan(powerPointDescent(size, office));
	});

	it('keeps a loading web font and other missing fonts on their own path', () => {
		expect(shiftFor('MS PMincho', times(true, true))).toBe(
			shiftFor('Times New Roman', times(true, true)),
		);
		expect(shiftFor('Foo Sans', times(true))).toBe(shiftFor('Times New Roman', times(true)));
	});
});

describe('the pixel snap of the text body top', () => {
	const withTop = (bodyTop?: { painted: number; exact: number }) =>
		resolveParagraphBaselineShift({
			runs: [run('Hxlb')],
			lineHeight: 2.4,
			strutSizePx: 40,
			strutFamily: 'Calibri',
			measure: calibri,
			...(bodyTop ? { bodyTop } : {}),
		});

	it('adds the move onto the pixel PowerPoint snaps the body top to', () => {
		// 40px text: the body top is moved by round(exact + 0.3) - painted.
		expect(withTop()).toBeCloseTo(11, 2);
		expect(withTop({ painted: 13, exact: 13.67 })).toBeCloseTo(12, 2);
		expect(withTop({ painted: 14, exact: 13.33 })).toBeCloseTo(11, 2);
		expect(withTop({ painted: 13, exact: 13.1 })).toBeCloseTo(11, 2);
	});

	it('moves a paragraph that is otherwise already on its baseline', () => {
		const shift = resolveParagraphBaselineShift({
			runs: [run('Hxlb')],
			lineHeight: 1.2,
			strutSizePx: 20,
			strutFamily: 'Calibri',
			measure: calibri,
			bodyTop: { painted: 13, exact: 13.67 },
		});
		expect(shift).toBeCloseTo(1, 0);
	});
});

describe('fontSizePx', () => {
	it('reads the first px size of a font shorthand', () => {
		expect(fontSizePx('italic 700 16px Arial')).toBe(16);
		expect(fontSizePx('bold 13.5px/20px "Segoe UI"')).toBe(13.5);
		expect(fontSizePx('1.5.2px x')).toBe(5.2);
		expect(fontSizePx('px 12px')).toBe(12);
		expect(fontSizePx('12pt Arial')).toBeNaN();
	});

	it('stays linear on long digit runs without px (code scanning: polynomial patterns)', () => {
		const started = performance.now();
		expect(fontSizePx('0'.repeat(50000))).toBeNaN();
		expect(fontSizePx(`${'0'.repeat(50000)}px`)).toBe(0);
		expect(performance.now() - started).toBeLessThan(200);
	});
});
