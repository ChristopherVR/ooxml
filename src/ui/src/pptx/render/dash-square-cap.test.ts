import { describe, expect, it } from 'vitest';

import { connectorSquareCapPath } from './connector-dash';
import { buildConnectorGeometry } from './connector-path';
import { squareCapDashedPathData } from './dash-square-cap';
import { snappedStrokeTransform } from './device-pixel-stroke';

// PowerPoint (1280x720 export, 3 px pen, `dash` = 12 on / 9 off): the first dash reaches back
// 1.5 px, the last reaches on 1.5 px when the line ends inside a dash.
describe('dashed square-cap line ends', () => {
	it('extends the start back by half the pen', () => {
		const d = squareCapDashedPathData('M0,0 L100,0', 3, 'dash', 'sq');
		expect(d).toBe('M-1.5,0 L0,0 M0,0 L100,0');
	});

	it('extends the end when it falls in a dash', () => {
		// period 21: 100 mod 21 = 16 -> in the gap (12..21)
		expect(squareCapDashedPathData('M0,0 L100,0', 3, 'dash', 'sq')).not.toContain('L101.5');
		// 105 mod 21 = 0 -> exactly at a dash start: nothing; 110 mod 21 = 5 -> inside the dash
		expect(squareCapDashedPathData('M0,0 L110,0', 3, 'dash', 'sq')).toBe(
			'M-1.5,0 L0,0 M0,0 L110,0 M110,0 L111.5,0',
		);
	});

	it('follows the line direction', () => {
		expect(squareCapDashedPathData('M0,0 L0,110', 3, 'dash', 'sq')).toBe(
			'M0,-1.5 L0,0 M0,0 L0,110 M0,110 L0,111.5',
		);
	});

	it('leaves flat caps, solid lines and curves alone', () => {
		expect(squareCapDashedPathData('M0,0 L110,0', 3, 'dash', 'flat')).toBe('M0,0 L110,0');
		expect(squareCapDashedPathData('M0,0 L110,0', 3, 'solid', 'sq')).toBe('M0,0 L110,0');
		expect(squareCapDashedPathData('M0,0 C1,1 2,2 3,3', 3, 'dash', 'sq')).toBe('M0,0 C1,1 2,2 3,3');
	});

	it('is used by connectors without arrowheads only', () => {
		const style = { strokeDash: 'dash', lineCap: 'sq' as const };
		expect(connectorSquareCapPath('M0,0 L110,0', style, 3, false)).toContain('M-1.5,0');
		expect(connectorSquareCapPath('M0,0 L110,0', style, 3, true)).toBeUndefined();
	});

	it('puts the overshoot into a straight connector geometry', () => {
		const element = {
			id: 'c1',
			type: 'connector',
			x: 10,
			y: 10,
			width: 110,
			height: 0,
			shapeType: 'line',
			shapeStyle: { strokeWidth: 3, strokeDash: 'dash', lineCap: 'sq' },
		} as never;
		const geo = buildConnectorGeometry(element, 0);
		expect(geo.pathD).toContain('M-1.5,0');
		expect(geo.hitPathD).not.toContain('-1.5');
	});
});

describe('dashed lines start on the snapped device pixel along their axis', () => {
	it('shifts a dashed horizontal line in x as well as y', () => {
		expect(snappedStrokeTransform('h', 40.3, 100, 1)).toContain('translate(0px');
		expect(snappedStrokeTransform('h', 40.3, 100, 1, undefined, undefined, true)).not.toContain(
			'translate(0px',
		);
	});
});
