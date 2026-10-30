import { describe, expect, it } from 'vitest';
import { parsePath, pathEnds, pathOperators } from './pdf-path';

describe('parsePath', () => {
	it('makes relative, shorthand and arc commands absolute cubic or line segments', () => {
		const segments = parsePath('m 1 1 h 2 v 2 q 1 1 2 0 a 1 1 0 0 1 2 0 z')!;
		expect(segments[0]).toEqual({ t: 'M', x: 1, y: 1 });
		expect(segments[1]).toEqual({ t: 'L', x: 3, y: 1 });
		expect(segments[2]).toEqual({ t: 'L', x: 3, y: 3 });
		expect(segments.at(-1)).toEqual({ t: 'Z' });
		expect(segments.filter((s) => s.t === 'C').length).toBeGreaterThanOrEqual(3);
		expect(pathOperators(segments).at(-1)).toBe('h');
	});

	it('rejects malformed or oversized data and reports the ends', () => {
		expect(parsePath('L 1')).toBeUndefined();
		expect(parsePath('M 0 0 ' + 'L 1 1 '.repeat(400_000))).toBeUndefined();
		const ends = pathEnds(parsePath('M 0 0 L 1 0 L 1 1')!);
		expect(ends.start!.angle).toBeCloseTo(0);
		expect(ends.end!.angle).toBeCloseTo(Math.PI / 2);
	});
});
