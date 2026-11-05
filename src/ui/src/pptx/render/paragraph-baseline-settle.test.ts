// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';

import { isFontSettled, resetInstalledFontCache } from './font-settled';
import {
	applyParagraphBaselineShift,
	BASELINE_PENDING_VAR,
	baselineShiftOf,
	decodePendingBaseline,
	encodePendingBaseline,
	type FontBox,
	type FontBoxMeasure,
	powerPointBaseline,
} from './paragraph-baseline';
import { settleBaselines } from './paragraph-baseline-settle';
import type { ParagraphRun, RenderParagraph } from './paragraph-types';

const run = (text: string, style: ParagraphRun['style'] = {}): ParagraphRun => ({ text, style });
const sized = (ascent: number, descent: number, loading = false): FontBoxMeasure => {
	return (font) => {
		const size = Number(/(\d+(?:\.\d+)?)px/u.exec(font)?.[1]);
		const box: FontBox = { ascent: ascent * size, descent: descent * size };
		return loading ? { ...box, provisional: true, loading } : box;
	};
};
// Oswald (typo metrics, 1.193 / 0.289 em) and the sans-serif stand-in browsers paint it with.
const oswald = sized(1.193, 0.289);
const standIn = sized(0.905, 0.212, true);
const block = { fontFamily: 'Oswald, sans-serif', fontSizePx: 96 };
const paragraph = (): RenderParagraph => ({
	runs: [run('WEST')],
	bulletStyle: {},
	segmentIndices: [0],
	lineHeight: 1.2,
});

describe('single spacing of a font with a deep descent', () => {
	it('keeps the font own descent below the baseline instead of lifting it to 75%', () => {
		// Segoe Script: PowerPoint 16 measured 0.854 em; 75% of the pitch (0.9) was too high.
		expect(powerPointBaseline(48, 40, 15)).toBe(33);
		// Tighter spacings still hold the 75% floor.
		expect(powerPointBaseline(43.2, 40, 15)).toBeCloseTo(32.4, 10);
	});
});

describe('web fonts that have not loaded', () => {
	it('moves a single line by the font own proportions once the font is measured', () => {
		// Oswald at 96px: PowerPoint 0.966 em = 92.7px, CSS puts it at 101px.
		const para = paragraph();
		applyParagraphBaselineShift(para, {}, block, oswald);
		expect(para.baselineShiftPx).toBeCloseTo(-8.3, 1);
		expect(para.baselineProvisional).toBeUndefined();
		expect(para.paragraphStyle).not.toHaveProperty(BASELINE_PENDING_VAR);
	});

	it('leaves single spacing to CSS on a stand-in face and marks the paragraph to be placed again', () => {
		const para = paragraph();
		applyParagraphBaselineShift(para, {}, block, standIn);
		expect(para.baselineShiftPx).toBeUndefined();
		expect(para.baselineProvisional).toBe(true);
		expect(para.paragraphStyle).toHaveProperty(BASELINE_PENDING_VAR);
		expect(para.paragraphStyle).not.toHaveProperty('top');
	});

	it('places a font the machine lacks on its fallback and still marks it for a later face', () => {
		const missing: FontBoxMeasure = (font) => {
			const box = oswald(font);
			return box && { ...box, provisional: true };
		};
		const para = paragraph();
		applyParagraphBaselineShift(para, {}, block, missing);
		expect(para.baselineShiftPx).toBeCloseTo(-8.3, 1);
		expect(para.paragraphStyle).toHaveProperty(BASELINE_PENDING_VAR);
	});

	it('does not take the stand-in proportions for the font (the old behaviour measured them)', () => {
		const geometry = (reference: FontBox, loading: boolean) => {
			const font = { ascent: 115, descent: 28 };
			return baselineShiftOf({
				size: 96,
				pitch: 115.2,
				boxes: [
					{ lineHeight: 115.2, font },
					{ lineHeight: 115.2, font },
				],
				reference,
				strutSizePx: 96,
				loading,
			});
		};
		// Stand-in boxes: single spacing is whatever CSS draws.
		expect(geometry({ ascent: 922, descent: 200 }, true)).toBeUndefined();
		// The font's own box (Oswald): lifted onto PowerPoint's baseline.
		expect(geometry({ ascent: 1193, descent: 289 }, false)).toBeLessThan(-7);
	});

	it('round-trips what is needed to place the paragraph again', () => {
		const input = {
			runs: [run('abc', { fontFamily: '"Oswald", sans-serif', fontWeight: 700, fontSize: '96px' })],
			lineHeight: 0.9,
			strutSizePx: 80,
			strutFamily: 'Oswald',
		};
		const decoded = decodePendingBaseline(encodePendingBaseline(input));
		expect(decoded?.lineHeight).toBe(0.9);
		expect(decoded?.strutSizePx).toBe(80);
		expect(decoded?.runs[0]?.style).toStrictEqual(input.runs[0]?.style);
		expect(decodePendingBaseline('%')).toBeUndefined();
	});
});

describe('isFontSettled', () => {
	afterEach(() => {
		vi.unstubAllGlobals();
		resetInstalledFontCache();
	});

	const stubFonts = (faces: Array<{ family: string }>, loaded: boolean) => {
		const load = vi.fn(() => Promise.resolve([]));
		vi.stubGlobal('document', {
			fonts: {
				forEach: (fn: (face: { family: string }) => void) => faces.forEach(fn),
				check: () => loaded,
				load,
			},
		});
		return load;
	};

	it('is true for generic families without asking the font set', () => {
		stubFonts([], false);
		expect(isFontSettled('normal 40px sans-serif')).toBe(true);
	});

	it('is false, and starts the load, while a face of the family is not loaded', () => {
		const load = stubFonts([{ family: '"Oswald"' }], false);
		expect(isFontSettled('normal 1000px "Oswald", sans-serif')).toBe(false);
		expect(load).toHaveBeenCalledOnce();
	});

	it('is true once the faces its text needs have loaded', () => {
		stubFonts([{ family: 'Oswald' }], true);
		expect(isFontSettled('normal 1000px "Oswald", sans-serif')).toBe(true);
	});
});

describe('settleBaselines', () => {
	it('drops a marker it cannot read and keeps one whose font is still a stand-in', () => {
		const root = document.createElement('div');
		const broken = document.createElement('p');
		broken.style.setProperty(BASELINE_PENDING_VAR, '%');
		root.append(broken);
		expect(settleBaselines(root)).toBe(0);
		expect(broken.style.getPropertyValue(BASELINE_PENDING_VAR)).toBe('');
	});
});
