import { describe, expect, it } from 'vitest';
import { visioFilledArrow, trimVisioArrowLine } from './filled-arrow';
import evidence from './__fixtures__/filled-arrows-native.json';
import bothEvidence from './__fixtures__/filled-arrows-both-native.json';

describe('native filled arrow geometry and setback at unit scale', () => {
	it.each(bothEvidence.cases)(
		'matches both endpoints code $code size $size weight $lineWidth',
		(row) => {
			const arrow = visioFilledArrow(row.code, row.size, row.lineWidth)!;
			expect(arrow.beginSetback).toBeCloseTo(row.beginSetback, 12);
			const path = trimVisioArrowLine(
				'M 0 0 L 2 0',
				arrow.beginSetback,
				arrow.setback,
				arrow.setback - arrow.beginSetback,
			)!;
			const actual = path.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
			const native = row.linePath.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
			expect(path.match(/[A-Z]/g)).toEqual(row.linePath.match(/[A-Z]/g));
			for (let i = 0; i < actual.length; i += 2) expect(actual[i]).toBeCloseTo(native[i]! / 72, 3);
		},
	);
	it.each(evidence.cases)('matches code $code size $size weight $lineWidth', (row) => {
		const arrow = visioFilledArrow(row.code, row.size, row.lineWidth)!;
		expect(arrow.setback).toBeCloseTo(row.setback, 12);
		expect(arrow.path.match(/[A-Z]/g)).toEqual(row.nativeGlyph.match(/[A-Z]/g));
		const numbers = (path: string) => path.match(/-?\d+(?:\.\d+)?/g)!.map(Number);
		const actual = numbers(arrow.path);
		numbers(row.nativeGlyph).forEach((n, i) =>
			expect(actual[i]).toBeCloseTo(n * row.extent * (i % 2 ? 1 : -1), 12),
		);
		const line = numbers(trimVisioArrowLine('M 0 0 L 2 0', 0, arrow.setback)!);
		const nativeLine = numbers(row.linePath);
		expect(line[2]).toBeCloseTo(nativeLine[2]! / 72, 3);
	});
	it.each([
		['M 0 0 L 2 0', 'M 0.1 0 L 1.8 0'],
		['M 0 0 L 0 2', 'M 0 0.1 L 0 1.8'],
		['M 0 0 L -2 0', 'M -0.1 0 L -1.8 0'],
	])('trims both endpoints along their direction: %s', (path, expected) => {
		expect(trimVisioArrowLine(path, 0.1, 0.2)).toBe(expected);
	});
	it.each([
		'M 0 0 Q 1 1 2 0',
		'M 0 0 L 1 0 L 2 0',
		'M 0 0 L 0 0',
		'M 0 0 L 0.1 0',
		'M NaN 0 L 2 0',
		'M 0 0 L 2 0 Z',
	])('does not guess unsupported or overlapping path %s', (path) =>
		expect(trimVisioArrowLine(path, 0.1, 0.2)).toBeUndefined(),
	);
});
