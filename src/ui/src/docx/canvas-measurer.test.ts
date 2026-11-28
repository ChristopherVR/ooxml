import { afterEach, expect, it, vi } from 'vitest';
import { createCanvasMeasurer } from './canvas-measurer';

// Century Gothic's hhea line is 2511 of 2048 units and its OS/2 win ascender 1989, descender 451.
// The browser reports each in whole pixels at the size it is asked for.
const sizeOf = (font: string) => Number(/([\d.]+)px/u.exec(font)?.[1] ?? NaN);

afterEach(() => vi.restoreAllMocks());

it('measures line heights without whole-pixel rounding (#383)', () => {
	vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function () {
		const context = {
			font: '',
			fontKerning: 'auto',
			measureText: () => {
				const size = sizeOf(context.font);
				return {
					width: 0,
					fontBoundingBoxAscent: Math.round((size * 1989) / 2048),
					fontBoundingBoxDescent: Math.round((size * 451) / 2048),
				};
			},
		};
		return context as unknown as CanvasRenderingContext2D;
	} as never);
	vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (
		this: HTMLElement,
	) {
		return { height: Math.round((sizeOf(this.style.font) * 2511) / 2048) } as DOMRect;
	});
	const font = { family: 'Century Gothic', sizePx: 40 / 3 };
	const measurer = createCanvasMeasurer();
	// Word's 10 pt Century Gothic line is 12.26 pt.
	expect((measurer.lineHeightOf(font) * 3) / 4).toBeCloseTo(12.26, 2);
	expect(measurer.ascentOf?.(font)).toBeCloseTo((40 / 3) * ((1989 + (2511 - 2440) / 2) / 2048), 3);
});
