import { expect, it } from 'vitest';
import { ooxmlArcToSvg } from './ooxml-arc';

it('represents a complete DrawingML turn with two nondegenerate SVG arcs', () => {
	for (const sweep of [21600000, -21600000]) {
		const result = ooxmlArcToSvg(100, 50, 0, sweep, 100, 0)!;
		expect(result.svg.match(/A /g)).toHaveLength(2);
		expect(result.endX).toBeCloseTo(100);
		expect(result.endY).toBeCloseTo(0);
		expect(result.svg).toContain('-100 0');
	}
});

it('declines invalid geometry and handles excessive turns without unbounded recursion', () => {
	expect(ooxmlArcToSvg(Infinity, 20, 0, 180, 0, 0)).toBeNull();
	expect(ooxmlArcToSvg(20, 20, 0, 21600000 * 1000000, 20, 0)?.svg.match(/A /g)).toHaveLength(2);
});
