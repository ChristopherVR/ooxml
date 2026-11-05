import type { PptxElement } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildConnectorGeometry,
	buildConnectorPathD,
	buildDashArray,
	buildWrapperStyle,
	connectorHitStrokeWidth,
	connectorSnapTransform,
	connectorWrapperTransform,
	markerPath,
} from './connector-path';
import { connectorStrandScreenWidth } from './connector-style';

describe('buildDashArray', () => {
	it('returns undefined for solid / missing dashes', () => {
		expect(buildDashArray('solid', 2)).toBeUndefined();
		expect(buildDashArray(undefined, 2)).toBeUndefined();
	});

	it('produces a distinct pattern per preset (not a single 3w/w approximation)', () => {
		// Aligned with the shape/border `getSvgStrokeDasharray` output.
		// Lengths measured from PowerPoint's own exports (see dash-pattern.ts).
		expect(buildDashArray('dot', 2)).toBe('2 6');
		expect(buildDashArray('sysDot', 2)).toBe('2 2');
		expect(buildDashArray('dash', 2)).toBe('8 6');
		expect(buildDashArray('sysDash', 2)).toBe('6 2');
		expect(buildDashArray('lgDash', 4)).toBe('32 12');
	});

	it('emits multi-part patterns for dash-dot families', () => {
		expect(buildDashArray('dashDot', 2)).toBe('8 6 2 6');
		expect(buildDashArray('lgDashDotDot', 2)).toBe('16 6 2 6 2 6');
	});

	it('honours custDash segments (percent-of-width, 1000ths of a percent)', () => {
		const segments = [
			{ dash: 400000, space: 300000 },
			{ dash: 100000, space: 300000 },
		];
		// custDash implies the custom family even without a prstDash token.
		expect(buildDashArray(undefined, 2, segments)).toBe('8 6 2 6');
	});

	it('clamps sub-1px stroke widths to 1 before scaling', () => {
		expect(buildDashArray('dot', 0)).toBe('1 3');
	});
});

describe('markerPath', () => {
	it('defaults to a med-sized (3x3) marker box measured in PowerPoint', () => {
		const m = markerPath('triangle');
		expect(m.shape).toBe('path');
		expect(m.d).toBe('M0 0 L10 5 L0 10 Z');
		expect(m.markerWidth).toBe(3);
		expect(m.markerHeight).toBe(3);
	});

	it('scales markerWidth by @len and markerHeight by @w', () => {
		const lg = markerPath('triangle', 'lg', 'lg');
		expect(lg.markerWidth).toBe(5);
		expect(lg.markerHeight).toBe(5);

		// arrowWidth=sm (height), arrowLength=med (width)
		const mixed = markerPath('oval', 'sm', 'med');
		expect(mixed.shape).toBe('circle');
		expect(mixed.markerWidth).toBe(3);
		expect(mixed.markerHeight).toBe(2);
	});

	it('keeps the historical shape/d for each arrow type', () => {
		expect(markerPath('diamond').d).toBe('M5 0 L10 5 L5 10 L0 5 Z');
		expect(markerPath('stealth').d).toBe('M0 0 L10 5 L0 10 L3 5 Z');
		expect(markerPath('oval').shape).toBe('circle');
	});
});

describe('connector pointer hit target', () => {
	it('never offers a target narrower than a finger', () => {
		expect(connectorHitStrokeWidth(0)).toBe(14);
		expect(connectorHitStrokeWidth(2)).toBe(14);
	});

	it('scales with a thick line so the target stays proportional', () => {
		expect(connectorHitStrokeWidth(6)).toBe(18);
		expect(connectorHitStrokeWidth(10)).toBe(30);
	});

	it('follows the endpoints of a straight connector, which has no path', () => {
		const geo = buildConnectorGeometry(
			{ id: 'c1', type: 'connector', x: 0, y: 0, width: 100, height: 50 } as PptxElement,
			3,
		);
		expect(geo.pathD).toBeUndefined();
		expect(geo.hitPathD).toBe('M0,0 L100,50');
		expect(geo.hitStrokeWidth).toBe(14);
	});

	it('follows the routed path when the connector bends', () => {
		const geo = buildConnectorGeometry(
			{
				id: 'c2',
				type: 'connector',
				shapeType: 'bentConnector3',
				x: 0,
				y: 0,
				width: 100,
				height: 50,
				shapeStyle: { strokeWidth: 8 },
			} as PptxElement,
			3,
		);
		expect(geo.hitPathD).toBe(geo.pathD);
		expect(geo.hitStrokeWidth).toBe(24);
	});
});

// ---------------------------------------------------------------------------
// buildConnectorPathD: OOXML preset bent/curved elbow routing
//
// PowerPoint's elbow connectors do not avoid obstacles (out of scope; see
// `connector-router.ts` for that, applied separately). What this covers: the
// bend axis remains horizontal-first as authored in the preset path list,
// regardless of the connector box aspect ratio. `bentConnector4`/`bentConnector5` (and their curved counterparts) no
// longer collapse into the exact same 3-segment shape as `bentConnector3`.
// See `connector-elbow-geometry.test.ts` for the underlying formula tests
// with hand-computed bend points; these exercise the same behaviour through
// the public `buildConnectorPathD` / `buildConnectorGeometry` entry points.
// ---------------------------------------------------------------------------

describe('buildConnectorPathD: bent/curved routing', () => {
	it('keeps producing the historical Z-shape for a wide (side-by-side) bentConnector3', () => {
		// Backward-compatibility check: a 100x50 box is horizontal-dominant, so
		// this must still match the pre-fix output exactly.
		expect(buildConnectorPathD('bentConnector3', 0, 0, 100, 50, 0.5)).toBe(
			'M0,0 L50,0 L50,50 L100,50',
		);
	});

	it('retains the horizontal-first preset path for a tall bentConnector3', () => {
		expect(buildConnectorPathD('bentConnector3', 0, 0, 50, 200, 0.5)).toBe(
			'M0,0 L25,0 L25,200 L50,200',
		);
	});

	it("gives bentConnector4 and bentConnector5 their own segment counts instead of bentConnector3's shape", () => {
		const three = buildConnectorPathD('bentConnector3', 0, 0, 200, 100, 0.5);
		const four = buildConnectorPathD('bentConnector4', 0, 0, 200, 100, 0.5, 0.5);
		const five = buildConnectorPathD('bentConnector5', 0, 0, 200, 100, 0.5, 0.5, 0.5);

		expect(three).toBe('M0,0 L100,0 L100,100 L200,100');
		expect(four).toBe('M0,0 L100,0 L100,50 L200,50 L200,100');
		// adj1 === adj3 (both 0.5) here, so the two primary-axis bend lines land
		// on the same x and the waypoint list has a (harmless) repeated point;
		// see the "honours explicit adj1/adj2/adj3" case below for adj1 != adj3.
		expect(five).toBe('M0,0 L100,0 L100,50 L100,50 L100,100 L200,100');
		expect(three).not.toBe(four);
		expect(four).not.toBe(five);
	});

	it('keeps the historical single-cubic curve for a wide curvedConnector3', () => {
		expect(buildConnectorPathD('curvedConnector3', 0, 0, 100, 50, 0.5)).toBe(
			'M0,0 C50,0 50,50 100,50',
		);
	});

	it('retains the horizontal-first curve for a tall curvedConnector3', () => {
		expect(buildConnectorPathD('curvedConnector3', 0, 0, 50, 200, 0.5)).toBe(
			'M0,0 C25,0 25,200 50,200',
		);
	});

	it('differentiates curvedConnector4/5 from curvedConnector3 (more control points, not the same S-curve)', () => {
		const three = buildConnectorPathD('curvedConnector3', 0, 0, 200, 100, 0.5);
		const four = buildConnectorPathD('curvedConnector4', 0, 0, 200, 100, 0.5, 0.5);
		expect(three).not.toBe(four);
		expect(four?.split('C').length).toBeGreaterThan(three?.split('C').length ?? 0);
	});

	it('honours explicit adj1/adj2/adj3 through buildConnectorGeometry end to end', () => {
		const geo = buildConnectorGeometry(
			{
				id: 'c3',
				type: 'connector',
				shapeType: 'bentConnector5',
				x: 0,
				y: 0,
				width: 200,
				height: 100,
				shapeAdjustments: { adj1: 25000, adj2: 50000, adj3: 75000 },
			} as PptxElement,
			1,
		);
		// Matches the hand-computed bentConnector5 case in
		// connector-elbow-geometry.test.ts for the same adjustments.
		expect(geo.pathD).toBe('M 0 0 L 50 0 L 50 50 L 150 50 L 150 100 L 200 100');
	});

	it('does not crash and still produces a monotonic route for diagonally offset shapes', () => {
		const geo = buildConnectorGeometry(
			{
				id: 'c4',
				type: 'connector',
				shapeType: 'bentConnector3',
				x: 10,
				y: 20,
				width: 137,
				height: 89,
			} as PptxElement,
			1,
		);
		expect(geo.pathD).toBe('M 0 0 L 68.5 0 L 68.5 89 L 137 89');
	});
});

// G0: a connector's flip is baked into its endpoints (x1/y1/x2/y2 swap
// above), so the wrapper transform must carry rotation only. Re-applying the
// flip as scaleX(-1)/scaleY(-1) on the wrapper would cancel the endpoint
// swap back out (this is exactly what React's ConnectorElementRenderer did
// by using the generic getElementTransform, which includes flip, before this
// fix routed it through connectorWrapperTransform instead).
describe('connectorWrapperTransform', () => {
	it('omits scale for a flipped bentConnector2, carrying rotation only', () => {
		const el = {
			id: 'c5',
			type: 'connector',
			shapeType: 'bentConnector2',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			flipHorizontal: true,
			rotation: 30,
		} as PptxElement;
		const transform = connectorWrapperTransform(el);
		expect(transform).toBe('rotate(30deg)');
		expect(transform).not.toContain('scale');
	});

	it('is undefined for an unrotated connector regardless of flip', () => {
		const el = {
			id: 'c6',
			type: 'connector',
			shapeType: 'straightConnector1',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			flipHorizontal: true,
			flipVertical: true,
		} as PptxElement;
		expect(connectorWrapperTransform(el)).toBeUndefined();
	});
});

describe('buildWrapperStyle', () => {
	it('never emits a scale transform for a flipped connector', () => {
		const el = {
			id: 'c7',
			type: 'connector',
			shapeType: 'bentConnector2',
			x: 0,
			y: 0,
			width: 100,
			height: 50,
			flipHorizontal: true,
			flipVertical: true,
			rotation: 45,
		} as PptxElement;
		const style = buildWrapperStyle(el, 2);
		expect(style).toContain('transform:rotate(45deg)');
		expect(style).not.toContain('scale');
	});
});

describe('connectorSnapTransform', () => {
	const line = (over: Partial<PptxElement> = {}): PptxElement =>
		({
			type: 'connector',
			id: 'c',
			x: 27,
			y: 27,
			width: 373,
			height: 0,
			shapeType: 'straightConnector1',
			shapeStyle: { strokeColor: '#000', strokeWidth: 1.3333 },
			...over,
		}) as PptxElement;

	it('snaps a horizontal line in y and a vertical line in x', () => {
		expect(buildConnectorGeometry(line(), 0).snapTransform).toMatch(/^translate\(0px, calc\(/u);
		const vertical = buildConnectorGeometry(line({ width: 0, height: 200 }), 0);
		expect(vertical.snapTransform).toMatch(/^translate\(calc\(.*\), 0px\)$/u);
	});

	it('ends a horizontal line at its real y, not at the 1px-floored SVG box', () => {
		const geo = buildConnectorGeometry(line(), 0);
		expect([geo.y1, geo.y2]).toEqual([0, 0]);
	});

	it('leaves diagonal, rotated, compound and bent lines unsnapped', () => {
		expect(buildConnectorGeometry(line({ height: 80 }), 0).snapTransform).toBeUndefined();
		expect(buildConnectorGeometry(line({ rotation: 30 }), 0).snapTransform).toBeUndefined();
		expect(
			buildConnectorGeometry(
				line({ shapeStyle: { strokeColor: '#000', strokeWidth: 6, compoundLine: 'dbl' } }),
				0,
			).snapTransform,
		).toBeUndefined();
		expect(connectorSnapTransform(line(), false, 1, 1.3333)).toBeUndefined();
	});

	it('snaps a single strand but leaves compound strand widths unrounded', () => {
		expect(connectorStrandScreenWidth([1.3333], 0.3333, 0)).toContain('round(0.3333px,');
		expect(connectorStrandScreenWidth([1.3333], 0.3333, 0)).toContain('clamp(');
		const compound = connectorStrandScreenWidth([2, 2, 2], 6, 1);
		expect(compound).toBe('max(2px, var(--pptx-device-px, 0px))');
	});
});
