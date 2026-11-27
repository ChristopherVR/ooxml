import { describe, expect, it } from 'vitest';
import { visioFilledArrow } from './filled-arrow';
import { visioOpenArrowExtent, visioOpenArrowPath } from './open-arrow';
import evidence from './__fixtures__/arrowheads-native.json';

const numbers = (path: string) => path.match(/-?\d+(?:\.\d+(?:e-?\d+)?)?/g)!.map(Number);
const extent = visioOpenArrowExtent(evidence.size, evidence.lineWidthIn);

// Every built-in line end Visio's export draws, including style 13 from the sample drawings.
describe('native arrowheads 1 to 45 at size 2 and 1 pt', () => {
	it.each(evidence.cases)('code $code matches the export glyph and setback', (row) => {
		const arrow = visioFilledArrow(row.code, evidence.size, evidence.lineWidthIn);
		const open = visioOpenArrowPath(row.code, evidence.size, evidence.lineWidthIn);
		const path = arrow?.path ?? open;
		expect(path).toBeDefined();
		// Glyph units are 3.52 pt; the export trims the line by the setback in points.
		expect((arrow?.setback ?? 0) * 72).toBeCloseTo(row.setbackPt, 1);
		if (arrow && row.setbackPt > 0) expect(arrow.filled).toBe(row.filled);
		expect(path!.match(/[A-Z]/g)).toEqual(row.glyph.match(/[A-Z]/g));
		const actual = numbers(path!);
		numbers(row.glyph).forEach((n, i) =>
			expect(actual[i]).toBeCloseTo(n * extent * (i % 2 ? 1 : -1), 9),
		);
	});
});
