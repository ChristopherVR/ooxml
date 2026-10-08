import { expect, it } from 'vitest';
import { demoDocument } from './demo-document';
import { visioSvgStrokeStyle } from './svg-stroke';
import { visioLineDashLengths } from './line-dash';

it('uses the native SVG zero-weight width without changing saved paint', () => {
	const style = Object.freeze({
		...demoDocument.pages[0]!.shapes[0]!.style,
		lineWidth: 0,
		linePattern: 3,
		lineDash: [0, 5],
		lineDashDotLength: 0.01 / 72,
	});
	const output = visioSvgStrokeStyle(style);
	expect(output).toEqual({ ...style, lineWidth: 0.75 / 72 });
	expect(style.lineWidth).toBe(0);
	const dash = visioLineDashLengths(output)!;
	expect(dash[0]).toBeCloseTo(0.01 / 72, 14);
	expect(dash[1]).toBeCloseTo((5 * 0.75) / 72, 14);
});

it.each([0.001 / 72, 0.75 / 72, 3 / 72])('preserves a nonzero %s-inch weight', (lineWidth) => {
	const style = { ...demoDocument.pages[0]!.shapes[0]!.style, lineWidth };
	expect(visioSvgStrokeStyle(style)).toBe(style);
});
