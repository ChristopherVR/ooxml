import { describe, expect, it } from 'vitest';
import { evaluatePresetShape } from './preset-shape-evaluator';

const points = (path: string): Array<[number, number]> =>
	[...path.matchAll(/[ML] ([-\d.]+) ([-\d.]+)/g)].map((m) => [Number(m[1]), Number(m[2])]);

describe('wedge callout presets', () => {
	it('draws the pointer to the tip below the box and keeps the box corners', () => {
		const tip = [160 / 2 + (160 * -20833) / 100000, 80 / 2 + (80 * 62500) / 100000];
		for (const name of ['wedgeRectCallout', 'wedgeRoundRectCallout']) {
			const result = evaluatePresetShape(name, 160, 80)!;
			const pts = points(result.svgPath);
			expect(
				pts.some(([x, y]) => Math.abs(x - tip[0]!) < 0.01 && Math.abs(y - tip[1]!) < 0.01),
				name,
			).toBe(true);
			expect(
				pts.every(([x, y]) => x >= -0.01 && x <= 160.01 && y >= -0.01 && y <= 90.01),
				name,
			).toBe(true);
		}
	});

	it('points the wedge at the side the tip is nearest to', () => {
		const left = evaluatePresetShape('wedgeRectCallout', 160, 80, { adj1: -70000, adj2: 0 })!;
		expect(points(left.svgPath).some(([x]) => x < -30)).toBe(true);
	});

	it('closes the oval callout from the pointer tip', () => {
		const result = evaluatePresetShape('wedgeEllipseCallout', 160, 80)!;
		expect(result.svgPath).toContain('A');
		expect(result.svgPath).toContain(`L ${(80 - 33.332).toFixed(3).slice(0, 5)}`);
	});
});
