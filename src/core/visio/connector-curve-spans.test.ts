import { expect, it } from 'vitest';
import { connectorCurveMiddle, connectorCurveSpans, cornerCurve } from './connector-curve-spans';
import { visioConnectorPreviewPath } from './ui/connector-preview';

// Native Visio 16.0.20430.20146 text controls, recorded from the six corner cases
// in scripts/record-visio-curved-route.ps1. Coordinates are local to the connector.
it.each([
	[
		[0, 0],
		[0, -2.5],
		[3.5, -2.5],
		[0.961632325171988, -2.196006568777631],
	],
	[
		[0, 0],
		[0, 1.5],
		[3.5, 1.5],
		[1.321790634219665, 1.387955034379223],
	],
	[
		[0, 0],
		[3.5, 0],
		[3.5, 0.5],
		[1.85134695928658, 0.02334405842632577],
	],
	[
		[0, 0],
		[0, -2.5],
		[0.5, -2.5],
		[0.02548504906688359, -1.363934955605657],
	],
	[
		[0, 0],
		[0, -2.5],
		[-0.5, -2.5],
		[-0.02548504906688359, -1.363934955605657],
	],
	[
		[0, 0],
		[0, 1.5],
		[-1.5, 1.5],
		[-0.28125, 1.21875],
	],
] as const)(
	'places corner-curve text halfway along native geometry %j %j %j',
	(begin, turn, end, text) => {
		const controls = cornerCurve([begin, turn, end].map(([x, y]) => ({ x, y })));
		const middle = connectorCurveMiddle(controls);
		expect(Math.abs(middle.x - text[0])).toBeLessThan(0.0002);
		expect(Math.abs(middle.y - text[1])).toBeLessThan(0.0002);
		const spans = connectorCurveSpans(controls);
		expect(spans).toHaveLength(2);
		expect(spans[0]![3]).toEqual(spans[1]![0]);
		const path = visioConnectorPreviewPath({ connectorId: 'c', route: 'curved', points: controls });
		expect(path.match(/ C /g)).toHaveLength(2);
		expect(path).not.toContain(' L ');
	},
);

it('keeps single cubic and straight preview paths intact', () => {
	const points = [
		{ x: 0, y: 0 },
		{ x: 0, y: 1 },
		{ x: 1, y: 1 },
		{ x: 1, y: 0 },
	];
	expect(connectorCurveMiddle(points)).toEqual({ x: 0.5, y: 0.75 });
	expect(visioConnectorPreviewPath({ connectorId: 'c', route: 'curved', points })).toBe(
		'M 0 0 C 0 1 1 1 1 0',
	);
	expect(
		visioConnectorPreviewPath({
			connectorId: 'c',
			route: 'straight',
			points: [points[0]!, points[3]!],
		}),
	).toBe('M 0 0 L 1 0');
});
