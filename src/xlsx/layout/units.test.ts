import { describe, expect, it } from 'vitest';
import {
	charactersToColumnWidth,
	charactersToPixels,
	columnWidthToPixels,
	defaultColumnPixels,
	pixelsToCharacters,
	pixelsToColumnWidth,
	pixelsToPoints,
	pointsToPixels,
} from './units.js';

describe('column width conversions', () => {
	it.each([
		[8.43, 64],
		[10, 75],
		[20, 145],
		[1, 12],
		[2, 19],
		[15, 110],
		[50, 355],
		[255, 1790],
	])('shows %s characters as %s px at mdw 7', (chars, px) => {
		expect(charactersToPixels(chars)).toBe(px);
	});

	it('converts the default 8.43 to the file width Excel writes', () => {
		expect(charactersToColumnWidth(8.43)).toBe(9.140625);
	});

	it.each([
		[9.140625, 64],
		[10, 70],
		[13, 91],
		[20.7109375, 145],
		[0, 0],
	])('renders file width %s as %s px', (width, px) => {
		expect(columnWidthToPixels(width)).toBe(px);
	});

	it('round-trips every pixel width through the file width', () => {
		for (let px = 1; px <= 2000; px++)
			expect(columnWidthToPixels(pixelsToColumnWidth(px))).toBe(px);
	});

	it('round-trips pixels at another digit width', () => {
		for (let px = 1; px <= 500; px++)
			expect(columnWidthToPixels(pixelsToColumnWidth(px, 8), 8)).toBe(px);
	});

	it('gives Excel UI widths for pixels', () => {
		expect(pixelsToCharacters(64)).toBe(8.43);
		expect(pixelsToCharacters(75)).toBe(10);
		expect(pixelsToCharacters(5)).toBe(0);
	});

	it('derives the default column width from baseColWidth (64 px at mdw 7)', () => {
		expect(defaultColumnPixels(undefined)).toBe(64);
		expect(defaultColumnPixels(undefined, 8)).toBe(72);
		expect(defaultColumnPixels(10)).toBe(70);
	});
});

describe('row heights', () => {
	it('converts points at 96 dpi', () => {
		expect(pointsToPixels(15)).toBe(20);
		expect(pointsToPixels(12.75)).toBe(17);
		expect(pointsToPixels(409.5)).toBe(546);
		expect(pointsToPixels(0)).toBe(0);
		expect(pixelsToPoints(20)).toBe(15);
	});
});
