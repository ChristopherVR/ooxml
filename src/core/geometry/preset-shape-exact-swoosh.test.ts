/**
 * swooshArrow, measured on PowerPoint 16 (an 840 x 420 pt shape exported at
 * 1920 x 1080 and compared with the evaluated outline at five adjustments and
 * four aspects: mean absolute difference below 0.0005). The head barbs sit a
 * fixed ssd8 from the box and shear by tan(90deg / 14) per unit of drop, so
 * both adjustments move them; the old hand-written definition ignored that.
 */

import { describe, expect, it } from 'vitest';

import { evaluatePresetShape } from './preset-shape-evaluator';

/** Each segment as `[kind, ...numbers]`: M, L or Q (control then end). */
function outline(w: number, h: number, adj1: number, adj2: number): number[][] {
	const d = evaluatePresetShape('swooshArrow', w, h, { adj1, adj2 })!.paths[0]!.d;
	return [...d.matchAll(/([MLQ])((?: -?[\d.]+)+)/g)].map((m) => [
		m[1] === 'M' ? 0 : m[1] === 'L' ? 1 : 2,
		...m[2]!.trim().split(/\s+/).map(Number),
	]);
}

function near(actual: number | undefined, expected: number, tol = 0.2) {
	expect(Math.abs(actual! - expected)).toBeLessThan(tol);
}

describe('swooshArrow', () => {
	it('places the head barbs at the measured positions for the defaults', () => {
		const o = outline(840, 420, 25000, 16667);
		// Upper curve ends at (xB, yB); then xC/t, r/yD, xE/yE, xF/yF.
		near(o[1]![3], 770, 0.1);
		near(o[1]![4], 52.5, 0.1);
		near(o[2]![1], 764.08);
		near(o[2]![2], 0);
		near(o[3]![1], 840);
		near(o[3]![2], 84);
		near(o[4]![1], 787.74);
		near(o[4]![2], 210);
		near(o[5]![1], 781.83);
		near(o[5]![2], 157.5);
	});

	it('moves the barbs with both adjustments', () => {
		const o = outline(840, 420, 75000, 33000);
		near(o[1]![3], 701.5, 0.3);
		near(o[2]![1], 701.5 - 52.5 * Math.tan(Math.PI / 28), 0.3);
		near(o[3]![2], 189);
		near(o[4]![2], 420);
		near(o[5]![2], 367.5);
		near(o[5]![1], 701.5 + 315 * Math.tan(Math.PI / 28), 0.3);
	});

	it('curves through the measured bezier controls', () => {
		const o = outline(300, 500, 25000, 16667);
		// Upper control (w/6, h/3); lower control (w/4, yF + h/12) with yF = 162.5.
		near(o[1]![1], 50);
		near(o[1]![2], 500 / 3);
		near(o[6]![1], 75);
		near(o[6]![2], 162.5 + 500 / 12);
	});
});
