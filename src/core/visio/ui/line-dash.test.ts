import { expect, it, vi } from 'vitest';
import { linePattern } from '../line-pattern';
import { visioLineDashLengths } from './line-dash';

it.each([1, 3])('keeps native square dots at 0.01 points for %s-point strokes', (points) => {
	const pattern = linePattern(new Map([['LinePattern', { value: '5' }]]), vi.fn(), 'square');
	const actual = visioLineDashLengths({ ...pattern, lineWidth: points / 72 })!;
	const expected = points === 1 ? [7, 5, 0.01, 5, 0.01, 5] : [21, 15, 0.01, 15, 0.01, 15];
	for (let index = 0; index < expected.length; index++)
		expect(actual[index]! * 72).toBeCloseTo(expected[index]!, 12);
});
it('retains native butt-cap dash endpoints and round dots', () => {
	const cells = new Map([['LinePattern', { value: '21' }]]);
	expect(linePattern(cells, vi.fn(), 'butt').lineDash).toEqual([40, 8, 16, 8]);
	const dots = new Map([['LinePattern', { value: '3' }]]);
	expect(linePattern(dots, vi.fn(), 'round')).toEqual({ linePattern: 3, lineDash: [0, 5] });
	expect(linePattern(dots, vi.fn(), 'butt')).toEqual({ linePattern: 3, lineDash: [1, 4] });
});
it.each([0, 1e-20])('keeps fixed dots finite at width %s without division', (lineWidth) => {
	const pattern = linePattern(new Map([['LinePattern', { value: '3' }]]), vi.fn(), 'square');
	expect(visioLineDashLengths({ ...pattern, lineWidth })).toEqual([0.01 / 72, 5 * lineWidth]);
	expect(visioLineDashLengths({ lineWidth })).toBeUndefined();
});
