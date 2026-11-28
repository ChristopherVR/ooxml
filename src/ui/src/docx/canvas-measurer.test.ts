import { afterEach, describe, expect, it, vi } from 'vitest';
import { createCanvasMeasurer } from './canvas-measurer';

// Per family, in 2048ths of the em: the ascender and descender canvas reports (the Windows metrics,
// or the typographic ones for a font with USE_TYPO_METRICS) and the CSS normal line height, which
// adds the font's line gap. The browser reports each in whole pixels at the size it is asked for.
const FACES: Record<string, { ascent: number; descent: number; line: number }> = {
	// hhea 2511, OS/2 win 1989 + 451: 71 units of leading.
	'Century Gothic': { ascent: 1989, descent: 451, line: 2511 },
	// hhea 1825 + 443 + 87 line gap, win 1825 + 443: 87 units of leading.
	'Times New Roman': { ascent: 1825, descent: 443, line: 2355 },
	// hhea 1854 + 434 + 67 line gap, win 1854 + 434: 67 units of leading.
	Arial: { ascent: 1854, descent: 434, line: 2355 },
	// Win 1950 + 550 cover the hhea line gap: no leading.
	Calibri: { ascent: 1950, descent: 550, line: 2500 },
	// USE_TYPO_METRICS: typo 1400 + 648 with a 1434 unit (0.7 em) typographic gap.
	Gabriola: { ascent: 1400, descent: 648, line: 3482 },
};

const sizeOf = (font: string) => Number(/([\d.]+)px/u.exec(font)?.[1] ?? NaN);
const faceOf = (font: string) => {
	const face = Object.entries(FACES).find(([name]) => font.includes(name))?.[1];
	if (!face) throw new Error(`no face for ${font}`);
	return face;
};

afterEach(() => vi.restoreAllMocks());

function mockBrowser() {
	vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function () {
		const context = {
			font: '',
			fontKerning: 'auto',
			measureText: () => {
				const size = sizeOf(context.font);
				const face = faceOf(context.font);
				return {
					width: 0,
					fontBoundingBoxAscent: Math.round((size * face.ascent) / 2048),
					fontBoundingBoxDescent: Math.round((size * face.descent) / 2048),
				};
			},
		};
		return context as unknown as CanvasRenderingContext2D;
	} as never);
	vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (
		this: HTMLElement,
	) {
		const size = sizeOf(this.style.font);
		return { height: Math.round((size * faceOf(this.style.font).line) / 2048) } as DOMRect;
	});
}

it('measures line heights without whole-pixel rounding (#383)', () => {
	mockBrowser();
	const font = { family: 'Century Gothic', sizePx: 40 / 3 };
	const measurer = createCanvasMeasurer();
	// Word's 10 pt Century Gothic line is 12.26 pt.
	expect((measurer.lineHeightOf(font) * 3) / 4).toBeCloseTo(12.26, 2);
	expect(measurer.ascentOf?.(font)).toBeCloseTo((40 / 3) * ((1989 + 71) / 2048), 3);
});

describe('where the line gap goes (#382)', () => {
	const sizePx = 96;
	it.each([
		['Times New Roman', 87],
		['Arial', 67],
	])('puts all of %s leading above the ascent, as Word does', (family, leading) => {
		mockBrowser();
		const measurer = createCanvasMeasurer();
		const face = FACES[family]!;
		const font = { family, sizePx };
		expect(measurer.ascentOf?.(font)).toBeCloseTo((sizePx * (face.ascent + leading)) / 2048, 3);
		// The line height is the browser's, leading included; only the baseline moves.
		expect(measurer.lineHeightOf(font)).toBeCloseTo((sizePx * face.line) / 2048, 3);
		// CSS puts half the leading above the text, so Print Layout moves it down the other half.
		expect(measurer.baselineShiftOf(font)).toBeCloseTo((sizePx * leading) / 2 / 2048, 3);
	});

	it('leaves a font without leading where it was', () => {
		mockBrowser();
		const measurer = createCanvasMeasurer();
		const font = { family: 'Calibri', sizePx };
		expect(measurer.ascentOf?.(font)).toBeCloseTo((sizePx * 1950) / 2048, 3);
		expect(measurer.lineHeightOf(font)).toBeCloseTo((sizePx * 2500) / 2048, 3);
		expect(measurer.baselineShiftOf(font)).toBe(0);
	});

	it('splits the typographic gap of a USE_TYPO_METRICS font above and below the text', () => {
		mockBrowser();
		const measurer = createCanvasMeasurer();
		const font = { family: 'Gabriola', sizePx };
		expect(measurer.ascentOf?.(font)).toBeCloseTo((sizePx * (1400 + 1434 / 2)) / 2048, 3);
		expect(measurer.baselineShiftOf(font)).toBe(0);
	});
});
