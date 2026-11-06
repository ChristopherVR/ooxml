/**
 * irregularSeal2, flowChartMagneticDisk and flowChartMultidocument, measured on
 * PowerPoint 16 (150 x 90 pt, defaults; the raster overlays match exactly).
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

function pointsOf(name: string): Array<[number, number]> {
	const svg = evaluatePresetShape(name, 150, 90)!.svgPath;
	const out: Array<[number, number]> = [];
	for (const m of svg.matchAll(/([MLAQC])((?: -?[\d.]+)+)/g)) {
		const n = m[2]!.trim().split(/\s+/).map(Number);
		out.push([n[n.length - 2]!, n[n.length - 1]!]);
	}
	return out;
}

function has(points: Array<[number, number]>, x: number, y: number, tol = 0.2): boolean {
	return points.some(([px, py]) => Math.abs(px - x) < tol && Math.abs(py - y) < tol);
}

describe('irregularSeal2', () => {
	it('reaches the right edge at 6645 / 21600 and the top at x = 14790 / 21600', () => {
		const pts = pointsOf('irregularSeal2');
		expect(pts).toHaveLength(28);
		expect(has(pts, 150, (6645 / 21600) * 90)).toBe(true);
		expect(has(pts, (14790 / 21600) * 150, 0)).toBe(true);
		expect(has(pts, (4917 / 21600) * 150, 90)).toBe(true);
	});
});

describe('flowChartMagneticDisk', () => {
	it('makes both ellipses h / 3 tall and draws the front arc of the top one', () => {
		const result = evaluatePresetShape('flowChartMagneticDisk', 150, 90)!;
		// the side lines run between the ellipse centres, h / 6 and 5h / 6
		expect(result.paths[0]!.d).toContain('L 150 75');
		expect(result.paths[0]!.d.startsWith('M 0 15')).toBe(true);
		expect(result.paths[1]!.fill).toBe('none');
	});
});

describe('flowChartMultidocument', () => {
	it('ends the front document at 18595 / 21600 and the middle one at 20000 / 21600', () => {
		const pts = pointsOf('flowChartMultidocument');
		expect(has(pts, (18595 / 21600) * 150, (3675 / 21600) * 90)).toBe(true);
		expect(has(pts, (20000 / 21600) * 150, (1815 / 21600) * 90)).toBe(true);
	});

	it('strokes only the visible edges of the documents behind', () => {
		const paths = evaluatePresetShape('flowChartMultidocument', 150, 90)!.paths;
		// two fill-only documents, the front one, and one open stroke path
		expect(paths.filter((p) => p.stroke === false)).toHaveLength(2);
		expect(paths.filter((p) => p.fill === 'none')).toHaveLength(1);
		expect(paths.some((p) => p.fill === 'darken')).toBe(false);
	});
});
