import { describe, expect, it } from 'vitest';
import { layoutVisioFilledArrowLine, visioFilledArrow } from './filled-arrow';
import evidence from './__fixtures__/short-arrows-native.json';

describe('native short one-ended arrow stems', () => {
	it('rejects non-finite and inconsistent public marker descriptors', () => {
		for (const [setback, beginSetback] of [
			[NaN, 0],
			[0.1, Infinity],
			[-0.1, 0],
			[0.1, 0.2],
		])
			expect(
				layoutVisioFilledArrowLine('M 0 0 L 2 0', undefined, {
					path: '',
					setback: setback!,
					beginSetback: beginSetback!,
				}),
			).toBeUndefined();
	});
	it.each(evidence.cases)('matches code $code size $size weight $lineWidth ratio $ratio', (row) => {
		const layout = layoutVisioFilledArrowLine(
			`M 0 0 L ${row.length} 0`,
			undefined,
			visioFilledArrow(row.code, row.size, row.lineWidth),
		)!;
		expect(layout.endSetback).toBeCloseTo(row.setback, 12);
		expect(layout.path.match(/[A-Z]/g)).toEqual(row.path.match(/[A-Z]/g));
		const numbers = (path: string) => path.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
		const actual = numbers(layout.path),
			native = numbers(row.path);
		for (let i = 0; i < actual.length; i += 2) expect(actual[i]).toBeCloseTo(native[i]! / 72, 3);
	});
	it('retains direction for a reversed short line without mutating the source', () => {
		const path = 'M 1 1 L 0.99 1';
		expect(layoutVisioFilledArrowLine(path, undefined, visioFilledArrow(6, 2, 0.01))).toEqual({
			path,
			startSetback: 0,
			endSetback: 0,
		});
	});
});
