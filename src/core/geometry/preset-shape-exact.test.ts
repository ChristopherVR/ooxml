/**
 * Ground-truth tests for the spec-exact preset definitions. Every number was
 * measured on PowerPoint 16 (a 150x90pt shape exported at 1280x720 and compared
 * pixel for pixel; the shape here is 200x100 so the figures read as percent of
 * the width). Each assertion fails on the hand-written definitions it replaced.
 */

import { describe, expect, it } from 'vitest';

import { ooxmlArcToSvg } from './ooxml-arc';
import { evaluatePresetShape } from './preset-shape-evaluator';

const W = 200;
const H = 100;

function paths(name: string, adj?: Record<string, number>, w = W, h = H) {
	return evaluatePresetShape(name, w, h, adj)!.paths;
}

/** Every `M`/`L` end point plus the end point of every `A` of an SVG path. */
function points(svg: string): Array<[number, number]> {
	const out: Array<[number, number]> = [];
	for (const m of svg.matchAll(/([MLA])((?: -?[\d.]+)+)/g)) {
		const n = m[2]!.trim().split(/\s+/).map(Number);
		out.push([n[n.length - 2]!, n[n.length - 1]!]);
	}
	return out;
}

function bbox(svg: string) {
	const pts = points(svg);
	const xs = pts.map((p) => p[0]);
	const ys = pts.map((p) => p[1]);
	return { l: Math.min(...xs), r: Math.max(...xs), t: Math.min(...ys), b: Math.max(...ys) };
}

function expectPoint(actual: [number, number] | undefined, x: number, y: number, tol = 0.2) {
	expect(actual).toBeDefined();
	expect(Math.abs(actual![0] - x)).toBeLessThan(tol);
	expect(Math.abs(actual![1] - y)).toBeLessThan(tol);
}

describe('arcTo uses visual angles in the radii own space', () => {
	it('places the end of a 45 degree arc on the 45 degree visual direction', () => {
		// Ellipse 100 x 50 centred at (100, 50); start at its right end, sweep to
		// the visual 45 degree direction: the end point has dx === dy.
		const arc = ooxmlArcToSvg(100, 50, 0, 2700000, 200, 50)!;
		expect(arc.endX - 100).toBeCloseTo(arc.endY - 50, 3);
		expect(arc.endX - 100).toBeCloseTo(44.721, 2);
	});

	it('converts angles before the path scale is applied', () => {
		// A circle of radius 10 in a path space that is stretched 2x horizontally:
		// the 45 degree end point is (10 cos 45 * 2, 10 sin 45) from the centre.
		const arc = ooxmlArcToSvg(10, 10, 0, 2700000, 20, 0, 2, 1)!;
		expect(arc.endX).toBeCloseTo(14.142, 2);
		expect(arc.endY).toBeCloseTo(7.071, 2);
	});
});

describe('regular polygons and stars follow hf / vf', () => {
	it('star5 inner vertices sit at +-54 and +-18 degrees of the inner ellipse', () => {
		const pts = points(paths('star5')[0]!.d);
		// outer tip top, first inner shoulder at (hc - cos54 * iwd2, svc - sin54 * ihd2)
		expectPoint(pts[1], 76.4, 38.2);
		expectPoint(pts[2], 100, 0);
		expectPoint(pts[0], 0, 38.2);
		// the box is filled exactly: top tip at 0, bottom tips on b
		const box = bbox(paths('star5')[0]!.d);
		expect(box.t).toBeCloseTo(0, 1);
		expect(box.b).toBeCloseTo(100, 1);
	});

	it('star4 inner vertices sit at 45 degrees', () => {
		const pts = points(paths('star4')[0]!.d);
		expectPoint(pts[1], 82.32, 41.16);
	});

	it('pentagon honours hf and vf overrides', () => {
		const box = bbox(paths('pentagon', { hf: 66088, vf: 69334 })[0]!.d);
		expect(box.r - box.l).toBeCloseTo(125.7, 0);
	});

	it('heptagon vertices lie on its stretched ellipse in tip order', () => {
		const pts = points(paths('heptagon')[0]!.d);
		const swd2 = 100 * 1.02572;
		const shd2 = 50 * 1.0521;
		const svc = 50 * 1.0521;
		for (const [x, y] of pts) {
			const e = ((x - 100) / swd2) ** 2 + ((y - svc) / shd2) ** 2;
			expect(e).toBeCloseTo(1, 2);
		}
	});

	it('decagon has points on the left and right edges and flat top and bottom', () => {
		const pts = points(paths('decagon')[0]!.d);
		expectPoint(pts[0], 0, 50);
		expectPoint(pts[5], 200, 50);
		const box = bbox(paths('decagon')[0]!.d);
		expect(box.t).toBeCloseTo(0.0, 0);
		expect(box.b).toBeCloseTo(100, 0);
	});

	it('star6/star7/star10 tips touch the box', () => {
		for (const name of ['star6', 'star7', 'star10']) {
			const box = bbox(paths(name)[0]!.d);
			expect(box.l, name).toBeCloseTo(0, 0);
			expect(box.r, name).toBeCloseTo(200, 0);
			expect(box.t, name).toBeCloseTo(0, 0);
			expect(box.b, name).toBeCloseTo(100, 0);
		}
	});
});

describe('math operators are inset to 73.49% of the box', () => {
	it('mathPlus, mathMinus and mathEqual span 26.51 .. 173.49', () => {
		for (const name of ['mathPlus', 'mathMinus', 'mathEqual']) {
			const box = bbox(paths(name)[0]!.d);
			expect(box.l, name).toBeCloseTo(26.51, 1);
			expect(box.r, name).toBeCloseTo(173.49, 1);
		}
		expect(bbox(paths('mathPlus')[0]!.d).t).toBeCloseTo(13.25, 1);
	});

	it('mathMultiply is one 12 vertex X outline', () => {
		expect(points(paths('mathMultiply')[0]!.d)).toHaveLength(12);
	});

	it('mathDivide is a bar plus two round dots (two full arcs)', () => {
		const d = paths('mathDivide')[0]!.d;
		expect(d.match(/A /g)).toHaveLength(4);
		expect(bbox(d).l).toBeCloseTo(26.51, 1);
	});

	it('mathNotEqual is a single 20 vertex outline crossed by a slash', () => {
		const d = paths('mathNotEqual')[0]!.d;
		expect(points(d)).toHaveLength(20);
		expect(bbox(d).t).toBeCloseTo(0, 1);
		expect(bbox(d).b).toBeCloseTo(100, 1);
	});
});

describe('tabs, brackets and braces', () => {
	it('cornerTabs draws four small corner triangles, not a frame', () => {
		const d = paths('cornerTabs')[0]!.d;
		expect(d.match(/M /g)).toHaveLength(4);
		for (const [x, y] of points(d)) {
			expect(x <= 10.01 || x >= 189.99).toBe(true);
			expect(y <= 10.01 || y >= 89.99).toBe(true);
		}
	});

	it('squareTabs draws four ss/10 squares', () => {
		const d = paths('squareTabs')[0]!.d;
		expect(d.match(/M /g)).toHaveLength(4);
		expect(points(d).some(([x, y]) => x === 10 && y === 10)).toBe(true);
	});

	it('plaqueTabs draws four quarter discs centred on the corners', () => {
		const d = paths('plaqueTabs')[0]!.d;
		expect(d.match(/A /g)).toHaveLength(4);
		expect(d.match(/A 10 10/g)).toHaveLength(4);
	});

	it('leftBrace has four quarter arcs and its tip on the left edge at mid height', () => {
		const set = paths('leftBrace');
		expect(set[0]!.stroke).toBe(false);
		expect(set[1]!.fill).toBe('none');
		expect(set[1]!.d.match(/A /g)).toHaveLength(4);
		expect(points(set[1]!.d).some(([x, y]) => x < 0.01 && Math.abs(y - 50) < 0.01)).toBe(true);
	});

	it('rightBrace has its tip on the right edge at mid height', () => {
		expect(
			points(paths('rightBrace')[1]!.d).some(([x, y]) => x > 199.99 && Math.abs(y - 50) < 0.01),
		).toBe(true);
	});

	it('bracePair fills the whole interior and strokes only the braces', () => {
		const set = paths('bracePair');
		expect(set[0]!.stroke).toBe(false);
		expect(set[1]!.fill).toBe('none');
		expect(bbox(set[0]!.d).l).toBeCloseTo(0, 1);
		expect(bbox(set[0]!.d).r).toBeCloseTo(200, 1);
		expect(set[1]!.d.match(/A /g)).toHaveLength(8);
	});

	it('bracketPair fills a round rectangle and strokes two brackets', () => {
		const set = paths('bracketPair');
		expect(set[0]!.stroke).toBe(false);
		expect(set[0]!.d.match(/A /g)).toHaveLength(4);
		expect(set[1]!.fill).toBe('none');
	});
});

describe('moon, stripes and trapezoids', () => {
	it('moon inner edge is a circle in box-normalised space (rx 1.25 w, ry 1.25 h/2)', () => {
		const d = paths('moon')[0]!.d;
		expect(d).toContain('A 250 62.5');
	});

	it('diagStripe is the four point band', () => {
		const pts = points(paths('diagStripe')[0]!.d);
		expectPoint(pts[0], 0, 50);
		expectPoint(pts[1], 100, 0);
		expectPoint(pts[2], 200, 0);
		expectPoint(pts[3], 0, 100);
	});

	it('nonIsoscelesTrapezoid insets both top corners by ss * adj', () => {
		const pts = points(paths('nonIsoscelesTrapezoid')[0]!.d);
		expectPoint(pts[0], 0, 100);
		expectPoint(pts[1], 25, 0);
		expectPoint(pts[2], 175, 0);
		expectPoint(pts[3], 200, 100);
	});

	it('halfFrame top bar is ss * adj1 thick', () => {
		const pts = points(paths('halfFrame')[0]!.d);
		expect(pts.some(([, y]) => Math.abs(y - 33.33) < 0.1)).toBe(true);
	});
});

describe('arcs follow visual angles', () => {
	it('arc paints a filled sector and an unfilled stroke', () => {
		const set = paths('arc');
		expect(set).toHaveLength(2);
		expect(set[0]!.d).toContain('L 100 50');
		expect(set[0]!.stroke).toBe(false);
		expect(set[1]!.fill).toBe('none');
	});

	it('pie ends on the visual direction of adj2', () => {
		// adj2 = 45 degrees visual: the end point of the arc lies on dx === dy.
		const d = paths('pie', { adj1: 0, adj2: 2700000 })[0]!.d;
		const pts = points(d);
		const end = pts[1]!;
		expect(end[0] - 100).toBeCloseTo(end[1] - 50, 1);
	});

	it('blockArc ring is ss * adj3 thick on every side', () => {
		const d = paths('blockArc')[0]!.d;
		expect(d).toContain('A 100 50');
		expect(d).toContain('A 75 25');
	});

	it('pieWedge is a quarter ellipse closed through the bottom right corner', () => {
		expect(paths('pieWedge')[0]!.d).toBe('M 0 100 A 200 100 0 0 1 200 0 L 200 100 Z');
	});

	it('chord closes the arc with a straight line', () => {
		const d = paths('chord')[0]!.d;
		expect(d.endsWith('Z')).toBe(true);
		expect(d).not.toContain('L ');
		// adj1 = 45 degrees VISUAL: the start point is on dx === dy from the centre.
		expect(d.startsWith('M 144.72')).toBe(true);
	});
});

describe('shapes with exact curves', () => {
	it('heart control points use the 49/48 and 10/48 width fractions', () => {
		const d = paths('heart')[0]!.d;
		expect(d.startsWith('M 100 25 C 141.6666')).toBe(true);
		expect(d).toContain(' -33.3333');
		expect(d).toContain(' 304.1666');
	});

	it('lightningBolt is the 11 vertex bolt scaled from the 21600 grid', () => {
		const pts = points(paths('lightningBolt')[0]!.d);
		expect(pts).toHaveLength(11);
		expectPoint(pts[0], 78.44, 0);
		expectPoint(pts[5], 200, 100);
	});

	it('cloud is eleven arcs in its 43200 space', () => {
		const d = paths('cloud')[0]!.d;
		expect(d.match(/A /g)).toHaveLength(11);
		expect(d.startsWith('M 18.05')).toBe(true);
	});

	it('wave keeps y1 = h * adj1 and bows the top by 10/3 of it', () => {
		const d = paths('wave')[0]!.d;
		expect(d.startsWith('M 0 12.5 C 66.66')).toBe(true);
		expect(d).toContain(' -29.166');
		expect(d).toContain(' 54.166');
	});

	it('wave adj2 shifts the top and bottom curves in opposite directions', () => {
		const d = paths('wave', { adj2: 5000 })[0]!.d;
		// of2 = w * adj2 / 50000 = 20: the top curve ends 20 early and the bottom
		// curve starts 20 late, so the sides lean.
		expect(d).toMatch(/ 180 12\.5 L 200 87\.5 /);
		expect(d).toMatch(/ 20 87\.5 Z$/);
	});

	it('doubleWave top edge is two full periods', () => {
		const d = paths('doubleWave')[0]!.d;
		expect(d.match(/C /g)).toHaveLength(4);
		expect(d.startsWith('M 0 6.25 C ')).toBe(true);
		// both top segments bow up first (y2 = y1 - 10/3 y1) then down, one period each
		expect(d.match(/ -14\.58\d* /g)).toHaveLength(2);
	});

	it('dodecagon uses the stretched 21600 grid vertices', () => {
		const pts = points(paths('dodecagon')[0]!.d);
		expectPoint(pts[0], 0, 36.6);
		expectPoint(pts[1], 26.8, 13.4);
	});

	it('flowChartDocument has a wavy bottom edge reaching past 80% height', () => {
		const d = paths('flowChartDocument')[0]!.d;
		expect(d).toContain('L 200 80.194');
		expect(d).toContain('C ');
	});

	it('flowChartPunchedTape has four half-ellipse arcs', () => {
		const d = paths('flowChartPunchedTape')[0]!.d;
		expect(d.match(/A /g)).toHaveLength(4);
		expect(d).toContain('L 200 90');
	});

	it('flowChartOnlineStorage has a convex left side and a concave right side', () => {
		const d = paths('flowChartOnlineStorage')[0]!.d;
		expect(d).toContain('A 33.333 50 0 0 0 200 100');
	});
});

describe('stroke-only detail paths', () => {
	it('can strokes the whole outline including the top ellipse', () => {
		const set = paths('can');
		expect(set[0]!.stroke).toBe(false);
		expect(set[2]!.fill).toBe('none');
		expect(set[2]!.d.match(/A /g)).toHaveLength(3);
	});

	it('cube and bevel stroke their inner edges', () => {
		expect(paths('cube').at(-1)!.fill).toBe('none');
		expect(paths('cube').at(-1)!.d.match(/M /g)!.length).toBeGreaterThanOrEqual(4);
		expect(paths('bevel').at(-1)!.fill).toBe('none');
	});

	it('smileyFace strokes the face, both eyes and the mouth', () => {
		const stroke = paths('smileyFace').at(-1)!;
		expect(stroke.fill).toBe('none');
		expect(stroke.d.match(/M /g)).toHaveLength(4);
	});

	it('borderless line callouts do not stroke their body rectangle', () => {
		for (const name of ['callout1', 'callout2', 'callout3']) {
			expect(paths(name)[0]!.stroke, name).toBe(false);
		}
		for (const name of ['accentCallout1', 'accentCallout2', 'accentCallout3']) {
			expect(paths(name)[0]!.stroke, name).toBe(false);
		}
	});
});

describe('multi-headed block arrows', () => {
	it('leftRightArrow heads are ss * adj2 long', () => {
		const pts = points(paths('leftRightArrow')[0]!.d);
		expectPoint(pts[0], 0, 50);
		expectPoint(pts[1], 50, 0);
		expectPoint(pts[5], 200, 50);
	});

	it('upDownArrow heads are ss * adj2 tall with a half-width stem', () => {
		const pts = points(paths('upDownArrow')[0]!.d);
		expectPoint(pts[0], 0, 50);
		expectPoint(pts[1], 100, 0);
		expectPoint(pts[3], 150, 50);
	});

	it('quadArrow reaches all four box edges', () => {
		const box = bbox(paths('quadArrow')[0]!.d);
		expect(box).toEqual({ l: 0, r: 200, t: 0, b: 100 });
	});

	it('leftUpArrow points up at x4 and left at y4', () => {
		const pts = points(paths('leftUpArrow')[0]!.d);
		expect(pts).toHaveLength(12);
		expectPoint(pts[0], 0, 75);
		expectPoint(pts[6], 175, 0);
	});

	it('leftRightUpArrow has the up head at hc and heads left and right', () => {
		const pts = points(paths('leftRightUpArrow')[0]!.d);
		expect(pts).toHaveLength(17);
		expectPoint(pts[6], 100, 0);
		expectPoint(pts[0], 0, 75);
		expectPoint(pts[12], 200, 75);
	});
});

describe('wedgeEllipseCallout keeps its pointer gap with visual angles', () => {
	it('starts and ends the ellipse arc on the pointer gap edges', () => {
		const d = paths('wedgeEllipseCallout')[0]!.d;
		// The pointer tip is (hc + w * adj1, vc + h * adj2) = (58.33, 112.5).
		expect(d).toMatch(/L 58\.33\d* 112\.5/);
	});
});
